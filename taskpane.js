/* Os Meus Templates — gestor pessoal de templates de email para Outlook (Office.js) */

// Chaves da versão anterior (localStorage) — só usadas para migrar os dados
const LS_TEMPLATES_ANTIGO = "meusTemplates.v1";
const LS_CATEGORIAS_ANTIGO = "meusTemplates.categorias.v1";

const DB_NOME = "meusTemplatesDB";
const DB_STORE = "kv";
const MAX_RECENTES = 8;
const CATEGORIAS_PADRAO = ["Geral", "Comercial", "Suporte", "Propostas"];

let templates = [];
let categoriasCustom = [];
let recentes = {};            // { nomeVariavelMinusculas: [valores recentes] }
let prefs = { ordenacao: "usados", modoInsercao: "cursor" };

let templateEmEdicao = null;  // id do template a editar, ou null se novo
let anexoTemp = null;         // { nome, tipoConteudo, base64 } carregado no editor
let anexoACarregar = false;
let templateAUsar = null;
let varsAUsar = [];           // [{ nome, padrao, opcoes, auto }]
let anexosExtraUso = [];
let anexosExtraACarregar = false;
let backupImportarTemp = null;
let ultimoRangeEditor = null;
let rangeLink = null;

/* =====================================================================
   Armazenamento (IndexedDB, com recurso ao localStorage se indisponível)
   ===================================================================== */

let dbPromise = null;
let usarLocalStorage = false;

function abrirDB() {
  if (!dbPromise) {
    dbPromise = new Promise((resolve, reject) => {
      if (!window.indexedDB) { reject(new Error("IndexedDB indisponível")); return; }
      const req = indexedDB.open(DB_NOME, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(DB_STORE);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
      req.onblocked = () => reject(new Error("IndexedDB bloqueado"));
    });
  }
  return dbPromise;
}

async function kvGet(chave) {
  if (!usarLocalStorage) {
    try {
      const db = await abrirDB();
      return await new Promise((resolve, reject) => {
        const req = db.transaction(DB_STORE, "readonly").objectStore(DB_STORE).get(chave);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
      });
    } catch (e) {
      console.warn("IndexedDB indisponível, a usar localStorage:", e);
      usarLocalStorage = true;
    }
  }
  const raw = localStorage.getItem("meusTemplates.kv." + chave);
  return raw ? JSON.parse(raw) : undefined;
}

async function kvSet(chave, valor) {
  if (!usarLocalStorage) {
    const db = await abrirDB();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, "readwrite");
      tx.objectStore(DB_STORE).put(valor, chave);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error || new Error("Gravação abortada"));
    });
    return;
  }
  localStorage.setItem("meusTemplates.kv." + chave, JSON.stringify(valor));
}

function lerJsonLocalStorage(chave) {
  try {
    const raw = localStorage.getItem(chave);
    return raw ? JSON.parse(raw) : undefined;
  } catch (e) {
    return undefined;
  }
}

async function carregarTudo() {
  let tpl = await kvGet("templates");
  let cats = await kvGet("categorias");

  if (tpl === undefined) {
    // Migração automática da versão anterior (dados no localStorage)
    const antigos = lerJsonLocalStorage(LS_TEMPLATES_ANTIGO);
    const catsAntigas = lerJsonLocalStorage(LS_CATEGORIAS_ANTIGO);
    tpl = Array.isArray(antigos) ? antigos.map(normalizarTemplate) : [];
    if (cats === undefined && Array.isArray(catsAntigas)) cats = catsAntigas;

    if (antigos || catsAntigas) {
      try {
        await kvSet("templates", tpl);
        if (cats) await kvSet("categorias", cats);
        localStorage.removeItem(LS_TEMPLATES_ANTIGO);
        localStorage.removeItem(LS_CATEGORIAS_ANTIGO);
        console.info(`Migrados ${tpl.length} template(s) para o novo armazenamento.`);
      } catch (e) {
        console.error("Falha na migração (os dados antigos foram mantidos):", e);
      }
    }
  }

  templates = (Array.isArray(tpl) ? tpl : []).map(normalizarTemplate);
  categoriasCustom = Array.isArray(cats) ? cats : [...CATEGORIAS_PADRAO];
  recentes = (await kvGet("recentes")) || {};
  prefs = Object.assign({}, prefs, (await kvGet("prefs")) || {});

  // Pede ao browser para não apagar estes dados quando houver pouco espaço
  if (navigator.storage && navigator.storage.persist) {
    navigator.storage.persist().catch(() => { });
  }
}

// Devolve true se gravou. Se falhar, repõe o último estado gravado e avisa.
async function guardarTemplates() {
  try {
    await kvSet("templates", templates);
    return true;
  } catch (e) {
    console.error("Falha ao gravar templates:", e);
    const gravados = await kvGet("templates").catch(() => undefined);
    templates = (Array.isArray(gravados) ? gravados : templates).map(normalizarTemplate);
    mostrarEstado("Não foi possível guardar (sem espaço de armazenamento?). As alterações não foram gravadas.", true);
    return false;
  }
}

async function guardarCategorias() {
  try {
    await kvSet("categorias", categoriasCustom);
    return true;
  } catch (e) {
    console.error("Falha ao gravar categorias:", e);
    mostrarEstado("Não foi possível guardar as categorias.", true);
    return false;
  }
}

function guardarRecentes() {
  return kvSet("recentes", recentes).catch(e => console.warn("Falha ao gravar valores recentes:", e));
}

function guardarPrefs() {
  return kvSet("prefs", prefs).catch(e => console.warn("Falha ao gravar preferências:", e));
}

function novoId() {
  return "t_" + Date.now() + "_" + Math.random().toString(36).slice(2, 7);
}

function normalizarTemplate(t) {
  const n = Object.assign({}, t);
  n.id = n.id || novoId();
  n.nome = n.nome || "Template Sem Nome";
  n.categoria = String(n.categoria || "Geral").trim() || "Geral";
  ["para", "cc", "bcc", "assunto"].forEach(k => { n[k] = n[k] || ""; });
  if (n.formato !== "html") {
    // Templates antigos tinham o corpo em texto simples
    n.corpo = textoParaHtml(n.corpo || "");
    n.formato = "html";
  } else {
    n.corpo = sanitizarHtml(n.corpo || "");
  }
  n.anexo = n.anexo || null;
  n.favorito = Boolean(n.favorito);
  n.usos = Number(n.usos) || 0;
  n.ultimoUso = Number(n.ultimoUso) || 0;
  return n;
}

/* =====================================================================
   HTML: conversão e limpeza
   ===================================================================== */

const TAGS_PERMITIDAS = new Set(["B", "STRONG", "I", "EM", "U", "BR", "P", "DIV", "SPAN", "UL", "OL", "LI", "A"]);
const TAGS_REMOVER = new Set(["SCRIPT", "STYLE", "IFRAME", "OBJECT", "EMBED", "IMG", "SVG", "MATH", "FORM",
  "INPUT", "BUTTON", "SELECT", "TEXTAREA", "META", "LINK", "TITLE", "HEAD", "XML", "NOSCRIPT", "TEMPLATE",
  "VIDEO", "AUDIO", "CANVAS"]);

function sanitizarHtml(html) {
  const doc = new DOMParser().parseFromString(String(html || ""), "text/html");
  limparNo(doc.body);
  return doc.body.innerHTML;
}

function limparNo(no) {
  Array.from(no.childNodes).forEach(filho => {
    if (filho.nodeType === Node.TEXT_NODE) return;
    if (filho.nodeType !== Node.ELEMENT_NODE || TAGS_REMOVER.has(filho.tagName)) {
      filho.remove();
      return;
    }
    limparNo(filho);
    if (!TAGS_PERMITIDAS.has(filho.tagName)) {
      filho.replaceWith(...Array.from(filho.childNodes));
      return;
    }
    Array.from(filho.attributes).forEach(a => {
      const manter = filho.tagName === "A" && a.name === "href" && /^(https?:|mailto:)/i.test(a.value.trim());
      if (!manter) filho.removeAttribute(a.name);
    });
  });
}

function textoParaHtml(texto) {
  return escapeHtml(texto).replace(/\r?\n/g, "<br>");
}

function htmlParaTexto(html) {
  const comQuebras = String(html || "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|li)>/gi, "\n");
  const doc = new DOMParser().parseFromString(comQuebras, "text/html");
  return doc.body.textContent || "";
}

/* =====================================================================
   Variáveis: {{nome}}, {{nome|omissão}}, {{nome:op1,op2}}, {{data+30}}
   ===================================================================== */

// Nome: letras (incl. acentuadas), números e _; opcionalmente +N / -N (datas)
const RE_VAR_SRC = "{{\\s*([A-Za-z0-9_\\u00C0-\\u024F]+(?:[+-]\\d+)?)\\s*(?:([|:])([^{}]*?))?\\s*}}";
function novaRegexVars() { return new RegExp(RE_VAR_SRC, "g"); }

function extrairVariaveis(texto) {
  const mapa = new Map();
  const re = novaRegexVars();
  let m;
  while ((m = re.exec(texto)) !== null) {
    const nome = m[1];
    if (!mapa.has(nome)) mapa.set(nome, { nome, padrao: "", opcoes: [] });
    const v = mapa.get(nome);
    if (m[2] === "|" && !v.padrao) v.padrao = m[3].trim();
    if (m[2] === ":" && !v.opcoes.length) v.opcoes = m[3].split(",").map(s => s.trim()).filter(Boolean);
  }
  return Array.from(mapa.values());
}

function substituirVariaveis(texto, valores, escapar) {
  return String(texto || "").replace(novaRegexVars(), (match, nome) => {
    if (valores[nome] === undefined) return match;
    return escapar ? escapeHtml(valores[nome]) : valores[nome];
  });
}

function valorData(nomeMinusculas) {
  const m = /^(data|data_atual|data_extenso)([+-]\d+)?$/.exec(nomeMinusculas);
  if (!m) return undefined;
  const d = new Date();
  if (m[2]) d.setDate(d.getDate() + parseInt(m[2], 10));
  return m[1] === "data_extenso"
    ? d.toLocaleDateString("pt-PT", { day: "numeric", month: "long", year: "numeric" })
    : d.toLocaleDateString("pt-PT");
}

function obterSaudacaoTempo() {
  const agora = new Date();
  const minTotais = agora.getHours() * 60 + agora.getMinutes();

  // 05:00 até às 13:49 (inclusive) -> Bom dia
  if (minTotais >= 5 * 60 && minTotais <= 13 * 60 + 49) return "Bom dia, ";
  // 13:50 até às 19:59 (inclusive) -> Boa tarde
  if (minTotais >= 13 * 60 + 50 && minTotais < 20 * 60) return "Boa tarde, ";
  // 20:00 às 04:59 -> Boa noite
  return "Boa noite, ";
}

function obterItem() {
  return (window.Office && Office.context && Office.context.mailbox && Office.context.mailbox.item) || null;
}

// Em modo de leitura o item tem displayReplyForm; ao compor não tem.
function modoLeitura() {
  const item = obterItem();
  return Boolean(item && typeof item.displayReplyForm === "function");
}

function obterValoresPreDefinidosAsync() {
  return new Promise((resolve) => {
    const agora = new Date();
    const horaFormatada = agora.toLocaleTimeString("pt-PT", { hour: "2-digit", minute: "2-digit" });
    const cumprimento = obterSaudacaoTempo();

    const valores = {
      cumprimento, cumprimentos: cumprimento, saudacao: cumprimento,
      saudacao_tempo: cumprimento, bom_dia_boa_tarde: cumprimento,
      hora_atual: horaFormatada, hora: horaFormatada,
      numdoc: "", num_doc: "", numero_documento: "",
      remetente_nome: "", remetente_email: "",
      destinatario_nome: "", destinatario_email: ""
    };

    try {
      const perfil = Office.context && Office.context.mailbox && Office.context.mailbox.userProfile;
      if (perfil) {
        valores.remetente_nome = perfil.displayName || "";
        valores.remetente_email = perfil.emailAddress || "";
      }
    } catch (e) {
      console.warn("Não foi possível obter o perfil do utilizador:", e);
    }

    try {
      const item = obterItem();
      if (item && modoLeitura()) {
        // Ao responder, o "destinatário" é quem nos enviou o email
        if (item.from) {
          valores.destinatario_nome = item.from.displayName || item.from.emailAddress || "";
          valores.destinatario_email = item.from.emailAddress || "";
        }
      } else if (item && item.to && typeof item.to.getAsync === "function") {
        item.to.getAsync((res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded && res.value && res.value.length > 0) {
            const primeiro = res.value[0];
            valores.destinatario_nome = primeiro.displayName || primeiro.emailAddress || "";
            valores.destinatario_email = primeiro.emailAddress || "";
          }
          resolve(valores);
        });
        return;
      }
    } catch (e) {
      console.warn("Não foi possível obter destinatários:", e);
    }

    resolve(valores);
  });
}

function valorAutomatico(nomeMinusculas, valoresPreDefinidos) {
  if (valoresPreDefinidos[nomeMinusculas] !== undefined) return valoresPreDefinidos[nomeMinusculas];
  return valorData(nomeMinusculas) || "";
}

/* =====================================================================
   Categorias
   ===================================================================== */

function obterCategoriasUnicas() {
  // As categorias padrão entram em categoriasCustom na primeira utilização,
  // por isso podem ser eliminadas como qualquer outra (exceto "Geral").
  const setCat = new Set(["Geral", ...categoriasCustom]);
  templates.forEach(t => {
    if (t.categoria && t.categoria.trim()) setCat.add(t.categoria.trim());
  });
  return Array.from(setCat).sort((a, b) => a.localeCompare(b, "pt"));
}

function atualizarDropdownCategorias() {
  const categorias = obterCategoriasUnicas();

  const selectFiltro = document.getElementById("filtro-categoria");
  const valorAtual = selectFiltro.value;
  selectFiltro.innerHTML = `<option value="">Todas as categorias</option>`;
  categorias.forEach(cat => {
    const opt = document.createElement("option");
    opt.value = cat;
    opt.textContent = cat;
    selectFiltro.appendChild(opt);
  });
  selectFiltro.value = categorias.includes(valorAtual) ? valorAtual : "";

  const datalist = document.getElementById("lista-categorias-sugeridas");
  datalist.innerHTML = "";
  categorias.forEach(cat => {
    const opt = document.createElement("option");
    opt.value = cat;
    datalist.appendChild(opt);
  });
}

function abrirModalCategorias() {
  renderListaGestaoCategorias();
  document.getElementById("novo-nome-categoria").value = "";
  document.getElementById("modal-categorias").classList.remove("oculto");
}

function renderListaGestaoCategorias() {
  const container = document.getElementById("lista-gestao-categorias");
  container.innerHTML = "";

  obterCategoriasUnicas().forEach(cat => {
    const count = templates.filter(t => t.categoria === cat).length;
    const isGeral = cat.toLowerCase() === "geral";

    const item = document.createElement("div");
    item.className = "item-gestao-categoria";
    item.innerHTML = `
      <div class="item-gestao-categoria-info">
        <strong>${escapeHtml(cat)}</strong>
        <span class="item-gestao-categoria-count">(${count} template${count === 1 ? "" : "s"})</span>
      </div>
      ${!isGeral
        ? `<button class="btn-eliminar-cat" data-cat="${escapeHtml(cat)}" title="Eliminar categoria">🗑️</button>`
        : `<span class="item-gestao-categoria-count">(padrão)</span>`}
    `;
    container.appendChild(item);
  });
}

async function adicionarNovaCategoria() {
  const input = document.getElementById("novo-nome-categoria");
  const nome = input.value.trim();

  if (!nome) {
    mostrarEstado("Escreve o nome da categoria.", true);
    return;
  }
  if (obterCategoriasUnicas().some(c => c.toLowerCase() === nome.toLowerCase())) {
    mostrarEstado("Essa categoria já existe.", true);
    return;
  }

  categoriasCustom.push(nome);
  await guardarCategorias();
  atualizarDropdownCategorias();
  renderListaGestaoCategorias();
  input.value = "";
  mostrarEstado(`Categoria "${nome}" criada.`);
}

async function eliminarCategoria(nomeCat) {
  if (nomeCat.toLowerCase() === "geral") {
    mostrarEstado("A categoria Geral não pode ser eliminada.", true);
    return;
  }

  const afetados = templates.filter(t => t.categoria === nomeCat).length;
  const ok = await confirmar(afetados
    ? `Eliminar a categoria "${nomeCat}"? Os ${afetados} template(s) desta categoria passam para "Geral".`
    : `Eliminar a categoria "${nomeCat}"?`);
  if (!ok) return;

  templates.forEach(t => { if (t.categoria === nomeCat) t.categoria = "Geral"; });
  categoriasCustom = categoriasCustom.filter(c => c !== nomeCat);

  if (afetados) await guardarTemplates();
  await guardarCategorias();
  atualizarDropdownCategorias();
  renderListaGestaoCategorias();
  renderLista();

  mostrarEstado(afetados
    ? `Categoria eliminada. ${afetados} template(s) movido(s) para "Geral".`
    : `Categoria "${nomeCat}" eliminada.`);
}

/* =====================================================================
   Navegação e confirmação
   ===================================================================== */

function mostrarVista(vista) {
  document.getElementById("lista-vista").classList.toggle("oculto", vista !== "lista");
  document.getElementById("editor-vista").classList.toggle("oculto", vista !== "editor");
  document.getElementById("usar-vista").classList.toggle("oculto", vista !== "usar");
  window.scrollTo(0, 0);
}

// Substitui window.confirm(), que não funciona dentro dos add-ins do Office
function confirmar(mensagem, textoSim = "Eliminar") {
  return new Promise(resolve => {
    const modal = document.getElementById("modal-confirmar");
    const sim = document.getElementById("btn-confirmar-sim");
    const nao = document.getElementById("btn-confirmar-nao");
    document.getElementById("confirmar-msg").textContent = mensagem;
    sim.textContent = textoSim;
    modal.classList.remove("oculto");

    const fechar = (resultado) => {
      modal.classList.add("oculto");
      sim.onclick = null;
      nao.onclick = null;
      resolve(resultado);
    };
    sim.onclick = () => fechar(true);
    nao.onclick = () => fechar(false);
    nao.focus();
  });
}

/* =====================================================================
   Lista de templates
   ===================================================================== */

function ordenarTemplates(lista) {
  const modo = prefs.ordenacao;
  return lista.slice().sort((a, b) => {
    if (a.favorito !== b.favorito) return a.favorito ? -1 : 1;
    if (modo === "usados" && b.usos !== a.usos) return b.usos - a.usos;
    if (modo === "recentes" && b.ultimoUso !== a.ultimoUso) return b.ultimoUso - a.ultimoUso;
    return a.nome.localeCompare(b.nome, "pt");
  });
}

function renderLista() {
  const container = document.getElementById("lista-templates");
  container.innerHTML = "";

  const termo = document.getElementById("pesquisa").value.trim().toLowerCase();
  const categoriaFiltro = document.getElementById("filtro-categoria").value;

  document.getElementById("btn-limpar-pesquisa").classList.toggle("oculto", !termo);
  document.getElementById("btn-exportar").textContent = categoriaFiltro
    ? `Exportar "${categoriaFiltro}"`
    : "Exportar Backup";

  const filtrados = templates.filter(t => {
    const correspondeCategoria = !categoriaFiltro || t.categoria === categoriaFiltro;
    if (!correspondeCategoria) return false;
    if (!termo) return true;
    return [t.nome, t.assunto, htmlParaTexto(t.corpo), t.categoria, t.para, t.cc, t.bcc]
      .some(campo => (campo || "").toLowerCase().includes(termo));
  });

  if (filtrados.length === 0) {
    container.innerHTML = templates.length === 0
      ? `<p class="ajuda">Sem templates. Cria o primeiro com "+ Novo".</p>`
      : `<p class="ajuda">Nenhum template encontrado para os filtros selecionados.</p>`;
    return;
  }

  ordenarTemplates(filtrados).forEach(t => {
    const id = escapeHtml(t.id);
    const card = document.createElement("div");
    card.className = "cartao-template";
    card.innerHTML = `
      <div class="cartao-header">
        <button class="btn-fav${t.favorito ? " ativo" : ""}" data-acao="favorito" data-id="${id}"
          title="${t.favorito ? "Remover dos favoritos" : "Marcar como favorito"}">${t.favorito ? "★" : "☆"}</button>
        <h3>${escapeHtml(t.nome)}</h3>
        <span class="badge-categoria">${escapeHtml(t.categoria || "Geral")}</span>
      </div>
      <p>${escapeHtml(t.assunto || "(Sem assunto)")}</p>
      <div class="cartao-acoes">
        <button class="btn btn-primary" data-acao="usar" data-id="${id}">Usar</button>
        <button class="btn btn-secondary" data-acao="editar" data-id="${id}">Editar</button>
        <button class="btn btn-secondary" data-acao="duplicar" data-id="${id}" title="Criar uma cópia deste template">Duplicar</button>
        ${t.usos ? `<span class="contagem-usos" title="Vezes usado">${t.usos}×</span>` : ""}
      </div>
    `;
    container.appendChild(card);
  });
}

async function alternarFavorito(id) {
  const t = templates.find(x => x.id === id);
  if (!t) return;
  t.favorito = !t.favorito;
  await guardarTemplates();
  renderLista();
}

async function duplicarTemplate(id) {
  const t = templates.find(x => x.id === id);
  if (!t) return;
  const copia = normalizarTemplate(Object.assign({}, t, {
    id: novoId(),
    nome: `${t.nome} (cópia)`,
    favorito: false,
    usos: 0,
    ultimoUso: 0
  }));
  templates.push(copia);
  if (!(await guardarTemplates())) { renderLista(); return; }
  renderLista();
  abrirEditorExistente(copia.id);
  mostrarEstado("Template duplicado — podes editar a cópia agora.");
}

/* =====================================================================
   Editor de template (texto com formatação)
   ===================================================================== */

function editorEl() {
  return document.getElementById("edit-corpo");
}

function guardarRangeEditor() {
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount) return;
  const r = sel.getRangeAt(0);
  if (editorEl().contains(r.commonAncestorContainer)) ultimoRangeEditor = r.cloneRange();
}

function restaurarRange(range) {
  const editor = editorEl();
  const sel = window.getSelection();
  sel.removeAllRanges();
  if (range && editor.contains(range.commonAncestorContainer)) {
    sel.addRange(range);
  } else {
    const r = document.createRange();
    r.selectNodeContents(editor);
    r.collapse(false);
    sel.addRange(r);
  }
}

function selecaoDentroEditor() {
  const sel = window.getSelection();
  return Boolean(sel && sel.rangeCount && editorEl().contains(sel.getRangeAt(0).commonAncestorContainer));
}

function executarComando(cmd, valor = null) {
  const editor = editorEl();
  // Usa a seleção atual se estiver no editor; senão, a última que lá esteve
  const dentro = selecaoDentroEditor();
  editor.focus();
  if (!dentro) restaurarRange(ultimoRangeEditor);
  document.execCommand(cmd, false, valor);
  guardarRangeEditor();
  atualizarVarsDetetadas();
}

function inserirTagVariavel(varNome) {
  executarComando("insertText", `{{${varNome}}}`);
}

function abrirLinhaLink() {
  guardarRangeEditor();
  rangeLink = ultimoRangeEditor ? ultimoRangeEditor.cloneRange() : null;
  document.getElementById("link-row").classList.remove("oculto");
  const input = document.getElementById("link-url");
  input.value = "";
  input.focus();
}

function fecharLinhaLink() {
  document.getElementById("link-row").classList.add("oculto");
  rangeLink = null;
}

function aplicarLink() {
  let url = document.getElementById("link-url").value.trim();
  const range = rangeLink;
  fecharLinhaLink();
  if (!url) return;
  if (!/^(https?:|mailto:)/i.test(url)) {
    url = (url.includes("@") && !url.includes("/") ? "mailto:" : "https://") + url;
  }
  const editor = editorEl();
  editor.focus();
  restaurarRange(range);
  if (range && !range.collapsed) {
    document.execCommand("createLink", false, url);
  } else {
    const texto = url.replace(/^mailto:/i, "");
    document.execCommand("insertHTML", false, `<a href="${escapeHtml(url)}">${escapeHtml(texto)}</a>`);
  }
  guardarRangeEditor();
  atualizarVarsDetetadas();
}

function colarNoEditor(e) {
  const dados = e.clipboardData || window.clipboardData;
  if (!dados) return;
  e.preventDefault();
  const html = dados.getData("text/html");
  if (html) {
    // Mantém negrito, listas e links; remove estilos, imagens e lixo do Word
    document.execCommand("insertHTML", false, sanitizarHtml(html));
  } else {
    document.execCommand("insertText", false, dados.getData("text/plain"));
  }
  atualizarVarsDetetadas();
}

function preencherEditor(t) {
  document.getElementById("edit-nome").value = t ? t.nome : "";
  document.getElementById("edit-categoria").value = t ? (t.categoria || "Geral") : "Geral";
  document.getElementById("edit-para").value = t ? t.para : "";
  document.getElementById("edit-cc").value = t ? t.cc : "";
  document.getElementById("edit-bcc").value = t ? t.bcc : "";
  document.getElementById("edit-assunto").value = t ? t.assunto : "";
  editorEl().innerHTML = t ? t.corpo : "";
  document.getElementById("edit-anexo").value = "";
  ultimoRangeEditor = null;
  anexoACarregar = false;
  fecharLinhaLink();
  atualizarAnexoAtualUI("Anexo atual");
  atualizarDropdownCategorias();
  atualizarVarsDetetadas();
  mostrarVista("editor");
}

function abrirEditorNovo() {
  templateEmEdicao = null;
  anexoTemp = null;
  document.getElementById("editor-titulo").textContent = "Novo template";
  document.getElementById("btn-eliminar").classList.add("oculto");
  preencherEditor(null);
  document.getElementById("edit-nome").focus();
}

function abrirEditorExistente(id) {
  const t = templates.find(x => x.id === id);
  if (!t) return;
  templateEmEdicao = id;
  anexoTemp = t.anexo || null;
  document.getElementById("editor-titulo").textContent = "Editar template";
  document.getElementById("btn-eliminar").classList.remove("oculto");
  preencherEditor(t);
}

function descreverVariavel(v) {
  if (v.opcoes.length) return `${v.nome} (lista: ${v.opcoes.join(", ")})`;
  if (v.padrao) return `${v.nome} (= ${v.padrao})`;
  return v.nome;
}

function atualizarVarsDetetadas() {
  const texto = [
    document.getElementById("edit-assunto").value,
    htmlParaTexto(editorEl().innerHTML),
    document.getElementById("edit-para").value,
    document.getElementById("edit-cc").value,
    document.getElementById("edit-bcc").value
  ].join(" ");
  const vars = extrairVariaveis(texto);
  document.getElementById("vars-detetadas").textContent = vars.length
    ? vars.map(descreverVariavel).join(", ")
    : "nenhuma";
}

async function guardarTemplateAtual() {
  if (anexoACarregar) {
    mostrarEstado("Aguarda que o anexo acabe de carregar.", true);
    return;
  }

  const nome = document.getElementById("edit-nome").value.trim();
  const categoria = document.getElementById("edit-categoria").value.trim() || "Geral";
  const dados = {
    nome,
    categoria,
    para: document.getElementById("edit-para").value.trim(),
    cc: document.getElementById("edit-cc").value.trim(),
    bcc: document.getElementById("edit-bcc").value.trim(),
    assunto: document.getElementById("edit-assunto").value.trim(),
    corpo: sanitizarHtml(editorEl().innerHTML),
    formato: "html",
    anexo: anexoTemp
  };

  if (!nome) {
    mostrarEstado("Dá um nome ao template.", true);
    return;
  }

  if (categoria !== "Geral" && !categoriasCustom.includes(categoria)) {
    categoriasCustom.push(categoria);
    await guardarCategorias();
  }

  if (templateEmEdicao) {
    const t = templates.find(x => x.id === templateEmEdicao);
    if (t) Object.assign(t, dados);
  } else {
    templates.push(normalizarTemplate(Object.assign({ id: novoId() }, dados)));
  }

  if (!(await guardarTemplates())) {
    // Fica no editor para o utilizador poder remover/trocar o anexo e tentar de novo
    renderLista();
    return;
  }
  atualizarDropdownCategorias();
  renderLista();
  mostrarVista("lista");
  mostrarEstado("Template guardado.");
}

async function eliminarTemplateAtual() {
  if (!templateEmEdicao) return;
  const t = templates.find(x => x.id === templateEmEdicao);
  const ok = await confirmar(`Eliminar o template "${t ? t.nome : ""}"? Esta ação não pode ser desfeita.`);
  if (!ok) return;

  templates = templates.filter(x => x.id !== templateEmEdicao);
  await guardarTemplates();
  atualizarDropdownCategorias();
  renderLista();
  mostrarVista("lista");
  mostrarEstado("Template eliminado.");
}

function atualizarAnexoAtualUI(prefixo = "Anexo") {
  const el = document.getElementById("anexo-atual");
  el.innerHTML = "";
  if (!anexoTemp) return;
  const span = document.createElement("span");
  span.textContent = `${prefixo}: ${anexoTemp.nome} `;
  const btn = document.createElement("button");
  btn.type = "button";
  btn.id = "btn-remover-anexo";
  btn.className = "btn btn-link";
  btn.textContent = "Remover";
  el.append(span, btn);
}

function removerAnexoTemplate() {
  anexoTemp = null;
  document.getElementById("edit-anexo").value = "";
  atualizarAnexoAtualUI();
}

function lerFicheiroBase64(ficheiro) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve({
      nome: ficheiro.name,
      tipoConteudo: ficheiro.type || "application/octet-stream",
      base64: String(reader.result).split(",")[1] || ""
    });
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(ficheiro);
  });
}

async function lidarComAnexoSelecionado(ficheiro) {
  if (!ficheiro) return;
  anexoACarregar = true;
  document.getElementById("anexo-atual").textContent = "A carregar…";
  try {
    anexoTemp = await lerFicheiroBase64(ficheiro);
    atualizarAnexoAtualUI("Anexo selecionado");
  } catch (e) {
    mostrarEstado("Não foi possível ler o ficheiro.", true);
    atualizarAnexoAtualUI("Anexo atual");
  } finally {
    anexoACarregar = false;
  }
}

/* =====================================================================
   Usar template
   ===================================================================== */

async function abrirUsarTemplate(id) {
  const t = templates.find(x => x.id === id);
  if (!t) return;
  templateAUsar = t;
  const leitura = modoLeitura();

  document.getElementById("usar-titulo").textContent = `Usar: ${t.nome}`;
  document.getElementById("usar-para").value = t.para || "";
  document.getElementById("usar-cc").value = t.cc || "";
  document.getElementById("usar-bcc").value = t.bcc || "";

  document.getElementById("usar-nota-leitura").classList.toggle("oculto", !leitura);
  document.getElementById("usar-seccao-destinatarios").classList.toggle("oculto", leitura);
  document.getElementById("usar-seccao-anexo-extra").classList.toggle("oculto", leitura);
  document.getElementById("usar-modo-insercao").classList.toggle("oculto", leitura);
  document.getElementById("btn-inserir").classList.toggle("oculto", leitura);
  document.getElementById("btn-responder").classList.toggle("oculto", !leitura);
  document.getElementById("btn-responder-todos").classList.toggle("oculto", !leitura);

  const radio = document.querySelector(`input[name="modo-insercao"][value="${prefs.modoInsercao}"]`);
  if (radio) radio.checked = true;

  // Ao responder, só o corpo é aplicado
  const textoVars = leitura
    ? htmlParaTexto(t.corpo)
    : [t.assunto, htmlParaTexto(t.corpo), t.para, t.cc, t.bcc].join(" ");
  varsAUsar = extrairVariaveis(textoVars);

  const container = document.getElementById("usar-campos");
  container.innerHTML = "<p class='ajuda'>A carregar dados do email…</p>";

  anexosExtraUso = [];
  anexosExtraACarregar = false;
  document.getElementById("usar-anexo-extra").value = "";
  document.getElementById("usar-anexo-extra-lista").textContent = "";

  const anexoFixoEl = document.getElementById("usar-anexo-fixo");
  if (t.anexo) {
    anexoFixoEl.textContent = leitura
      ? `Anexo fixo (não incluído na resposta): ${t.anexo.nome}`
      : `Anexo fixo deste template: ${t.anexo.nome}`;
    anexoFixoEl.classList.remove("oculto");
  } else {
    anexoFixoEl.classList.add("oculto");
  }

  mostrarVista("usar");
  const preDefinidos = await obterValoresPreDefinidosAsync();
  container.innerHTML = "";

  if (varsAUsar.length === 0) {
    container.innerHTML = `<p class="ajuda">Este template não tem variáveis de substituição.</p>`;
  }

  varsAUsar.forEach((v, i) => {
    const k = v.nome.toLowerCase();
    const auto = valorAutomatico(k, preDefinidos);
    const inicial = v.padrao || auto || "";
    v.auto = Boolean(auto) && !v.padrao;

    let label = v.nome;
    let placeholder = `Valor para {{${v.nome}}}`;
    if (["numdoc", "num_doc", "numero_documento"].includes(k)) {
      label = `${v.nome} (Número do Documento)`;
      placeholder = "Ex.: FT 2026/001, Orçamento nº...";
    } else if (k === "cumprimento" || k === "cumprimentos") {
      label = "Cumprimento";
    }

    const idCampo = `var-campo-${i}`;
    let controlo;
    let badge = v.auto ? `<span class="badge-auto">Preenchido auto</span>` : "";

    if (v.opcoes.length) {
      controlo = `<select id="${idCampo}" class="select-categoria" data-var="${escapeHtml(v.nome)}">
        ${v.opcoes.map(o => `<option value="${escapeHtml(o)}"${o === inicial ? " selected" : ""}>${escapeHtml(o)}</option>`).join("")}
      </select>`;
    } else {
      const sugestoes = v.auto ? [] : (recentes[k] || []);
      if (sugestoes.length && !badge) badge = `<span class="badge-recente" title="Clica no campo para ver os últimos valores usados">Recentes</span>`;
      const datalist = sugestoes.length
        ? `<datalist id="${idCampo}-rec">${sugestoes.map(s => `<option value="${escapeHtml(s)}"></option>`).join("")}</datalist>`
        : "";
      controlo = `<input id="${idCampo}" type="text" data-var="${escapeHtml(v.nome)}" value="${escapeHtml(inicial)}"
        placeholder="${escapeHtml(placeholder)}" autocomplete="off"${sugestoes.length ? ` list="${idCampo}-rec"` : ""} />${datalist}`;
    }

    const div = document.createElement("div");
    div.className = "campo-variavel";
    div.innerHTML = `
      <div class="campo-variavel-header">
        <label for="${idCampo}">${escapeHtml(label)}</label>
        ${badge}
      </div>
      ${controlo}
    `;
    container.appendChild(div);
  });

  atualizarPreview();
}

function recolherValores() {
  const valores = {};
  document.querySelectorAll("#usar-campos [data-var]").forEach(el => { valores[el.dataset.var] = el.value; });
  return valores;
}

function construirResultado() {
  const valores = recolherValores();
  return {
    valores,
    assunto: substituirVariaveis(templateAUsar.assunto, valores, false),
    corpoHtml: substituirVariaveis(templateAUsar.corpo, valores, true),
    para: substituirVariaveis(document.getElementById("usar-para").value, valores, false),
    cc: substituirVariaveis(document.getElementById("usar-cc").value, valores, false),
    bcc: substituirVariaveis(document.getElementById("usar-bcc").value, valores, false)
  };
}

function atualizarPreview() {
  if (!templateAUsar) return;
  const r = construirResultado();
  document.getElementById("preview-assunto").textContent = modoLeitura()
    ? "Resposta (o assunto mantém-se)"
    : `Assunto: ${r.assunto || "(sem assunto)"}`;
  document.getElementById("preview-corpo").innerHTML = sanitizarHtml(r.corpoHtml) || "<em>(corpo vazio)</em>";
}

async function lidarComAnexosExtra(ficheiros) {
  const lista = document.getElementById("usar-anexo-extra-lista");
  anexosExtraUso = [];
  anexosExtraACarregar = true;
  lista.textContent = "A carregar…";
  try {
    anexosExtraUso = await Promise.all(Array.from(ficheiros).map(lerFicheiroBase64));
    lista.textContent = anexosExtraUso.length
      ? `A anexar: ${anexosExtraUso.map(a => a.nome).join(", ")}`
      : "";
  } catch (e) {
    lista.textContent = "";
    mostrarEstado("Não foi possível ler um dos ficheiros.", true);
  } finally {
    anexosExtraACarregar = false;
  }
}

function parseEmailsParaDestinatarios(str) {
  if (!str || !str.trim()) return [];
  return str.split(/[;,]/)
    .map(email => email.trim())
    .filter(email => email.length > 0)
    .map(email => ({ emailAddress: email }));
}

function chamarAsync(fn) {
  return new Promise((resolve, reject) => {
    fn(res => (res.status === Office.AsyncResultStatus.Succeeded ? resolve(res.value) : reject(res.error)));
  });
}

async function registarUtilizacao(valores) {
  const t = templates.find(x => x.id === templateAUsar.id);
  if (t) {
    t.usos = (t.usos || 0) + 1;
    t.ultimoUso = Date.now();
  }

  varsAUsar.forEach(v => {
    if (v.auto || v.opcoes.length) return;
    const valor = String(valores[v.nome] || "").trim();
    if (!valor || valor === v.padrao) return;
    const k = v.nome.toLowerCase();
    recentes[k] = [valor, ...(recentes[k] || []).filter(x => x !== valor)].slice(0, MAX_RECENTES);
  });

  await Promise.all([guardarTemplates(), guardarRecentes()]);
  renderLista();
}

async function inserirTemplateNoEmail() {
  if (!templateAUsar) return;
  if (anexosExtraACarregar) {
    mostrarEstado("Aguarda que os anexos acabem de carregar.", true);
    return;
  }

  const item = obterItem();
  if (!item || !item.body || typeof item.body.setSelectedDataAsync !== "function") {
    mostrarEstado("Abre uma mensagem nova (ou uma resposta) para inserir o template.", true);
    return;
  }

  const r = construirResultado();
  const modoEl = document.querySelector('input[name="modo-insercao"]:checked');
  const modo = modoEl ? modoEl.value : "cursor";
  if (prefs.modoInsercao !== modo) {
    prefs.modoInsercao = modo;
    guardarPrefs();
  }

  const opcoesHtml = { coercionType: Office.CoercionType.Html };
  const tarefas = [];

  if (r.assunto) tarefas.push(chamarAsync(cb => item.subject.setAsync(r.assunto, cb)));

  if (modo === "substituir") {
    tarefas.push(chamarAsync(cb => item.body.setAsync(r.corpoHtml, opcoesHtml, cb)));
  } else if (modo === "inicio") {
    tarefas.push(chamarAsync(cb => item.body.prependAsync(r.corpoHtml, opcoesHtml, cb)));
  } else {
    tarefas.push(chamarAsync(cb => item.body.setSelectedDataAsync(r.corpoHtml, opcoesHtml, cb)));
  }

  [["to", r.para], ["cc", r.cc], ["bcc", r.bcc]].forEach(([campo, texto]) => {
    const lista = parseEmailsParaDestinatarios(texto);
    if (lista.length && item[campo] && typeof item[campo].addAsync === "function") {
      tarefas.push(chamarAsync(cb => item[campo].addAsync(lista, cb)));
    }
  });

  try {
    await Promise.all(tarefas);

    const anexos = [];
    if (templateAUsar.anexo) anexos.push(templateAUsar.anexo);
    anexos.push(...anexosExtraUso);
    for (const anexo of anexos) {
      await chamarAsync(cb => item.addFileAttachmentFromBase64Async(anexo.base64, anexo.nome, { isInline: false }, cb));
    }

    await registarUtilizacao(r.valores);
    mostrarVista("lista");
    mostrarEstado("Template inserido no email!");
  } catch (err) {
    console.error(err);
    mostrarEstado(`Erro ao inserir template${err && err.message ? `: ${err.message}` : "."}`, true);
  }
}

async function responderComTemplate(todos) {
  if (!templateAUsar) return;
  const item = obterItem();
  if (!item || typeof item.displayReplyForm !== "function") {
    mostrarEstado("Abre um email recebido para responder com o template.", true);
    return;
  }

  const r = construirResultado();
  try {
    const formulario = { htmlBody: r.corpoHtml };
    if (todos) item.displayReplyAllForm(formulario);
    else item.displayReplyForm(formulario);
    await registarUtilizacao(r.valores);
    mostrarVista("lista");
    mostrarEstado("Resposta aberta com o template.");
  } catch (err) {
    console.error(err);
    mostrarEstado("Não foi possível abrir a resposta.", true);
  }
}

/* =====================================================================
   Exportar / Importar
   ===================================================================== */

function slug(texto) {
  return String(texto).normalize("NFD").replace(/[̀-ͯ]/g, "")
    .replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "").toLowerCase() || "categoria";
}

function exportarTemplates() {
  const categoria = document.getElementById("filtro-categoria").value;
  const lista = categoria ? templates.filter(t => t.categoria === categoria) : templates;

  if (lista.length === 0) {
    mostrarEstado("Não existem templates para exportar.", true);
    return;
  }

  const payload = {
    versao: "2.0",
    dataExportacao: new Date().toISOString(),
    categoria: categoria || null,
    totalTemplates: lista.length,
    categoriasCustom: categoria ? [categoria] : categoriasCustom,
    templates: lista
  };

  const data = new Date().toISOString().slice(0, 10);
  const nomeFicheiro = categoria
    ? `os-meus-templates-${slug(categoria)}-${data}.json`
    : `os-meus-templates-backup-${data}.json`;

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nomeFicheiro;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);

  mostrarEstado(categoria
    ? `Exportados ${lista.length} template(s) da categoria "${categoria}".`
    : "Backup exportado com sucesso.");
}

function processarFicheiroImportar(ficheiro) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const conteudo = JSON.parse(reader.result);
      let listaImportada = [];
      let catsImportadas = [];

      if (Array.isArray(conteudo)) {
        listaImportada = conteudo;
      } else if (conteudo && Array.isArray(conteudo.templates)) {
        listaImportada = conteudo.templates;
        if (Array.isArray(conteudo.categoriasCustom)) catsImportadas = conteudo.categoriasCustom;
      } else {
        throw new Error("Formato inválido");
      }

      if (listaImportada.length === 0) {
        mostrarEstado("O ficheiro de backup não contém nenhum template.", true);
        return;
      }

      backupImportarTemp = {
        templates: listaImportada.map(t => normalizarTemplate(Object.assign({}, t, { id: novoId() }))),
        categoriasCustom: catsImportadas.filter(c => typeof c === "string" && c.trim())
      };

      document.getElementById("modal-importar-info").textContent =
        `Ficheiro validado! Encontrados ${backupImportarTemp.templates.length} template(s).`;
      document.getElementById("modal-importar").classList.remove("oculto");
    } catch (e) {
      mostrarEstado("Ficheiro JSON de backup inválido.", true);
    }
  };
  reader.readAsText(ficheiro);
}

async function executarImportacao(modo) {
  if (!backupImportarTemp) return;
  const importados = backupImportarTemp;
  backupImportarTemp = null;
  document.getElementById("modal-importar").classList.add("oculto");

  const categoriasAntes = [...categoriasCustom];

  if (modo === "unir") {
    templates = templates.concat(importados.templates);
    importados.categoriasCustom.forEach(c => {
      if (!categoriasCustom.includes(c)) categoriasCustom.push(c);
    });
  } else {
    templates = importados.templates;
    if (importados.categoriasCustom.length) categoriasCustom = importados.categoriasCustom;
  }

  if (await guardarTemplates()) {
    await guardarCategorias();
    mostrarEstado(modo === "unir"
      ? `${importados.templates.length} template(s) unidos com sucesso!`
      : `Biblioteca substituída com ${importados.templates.length} template(s).`);
  } else {
    // Importação revertida (guardarTemplates já repôs os templates e avisou)
    categoriasCustom = categoriasAntes;
  }
  atualizarDropdownCategorias();
  renderLista();
}

/* =====================================================================
   Utilitários
   ===================================================================== */

function escapeHtml(str) {
  return String(str == null ? "" : str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

let estadoTimeout = null;
function mostrarEstado(msg, erro = false) {
  const el = document.getElementById("estado-msg");
  el.textContent = msg;
  el.classList.toggle("erro", erro);
  el.classList.remove("oculto");
  clearTimeout(estadoTimeout);
  estadoTimeout = setTimeout(() => el.classList.add("oculto"), erro ? 5000 : 3500);
}

/* =====================================================================
   Eventos
   ===================================================================== */

function ligarEventos() {
  const porId = id => document.getElementById(id);

  // Navegação principal
  porId("btn-novo").addEventListener("click", abrirEditorNovo);
  porId("btn-cancelar").addEventListener("click", () => mostrarVista("lista"));
  porId("btn-guardar").addEventListener("click", guardarTemplateAtual);
  porId("btn-eliminar").addEventListener("click", eliminarTemplateAtual);
  porId("btn-usar-cancelar").addEventListener("click", () => mostrarVista("lista"));
  porId("btn-inserir").addEventListener("click", inserirTemplateNoEmail);
  porId("btn-responder").addEventListener("click", () => responderComTemplate(false));
  porId("btn-responder-todos").addEventListener("click", () => responderComTemplate(true));

  // Editor: campos simples
  ["edit-assunto", "edit-para", "edit-cc", "edit-bcc"].forEach(id =>
    porId(id).addEventListener("input", atualizarVarsDetetadas));
  porId("edit-anexo").addEventListener("change", e => lidarComAnexoSelecionado(e.target.files[0]));
  porId("anexo-atual").addEventListener("click", e => {
    if (e.target.closest("#btn-remover-anexo")) removerAnexoTemplate();
  });

  // Editor: texto com formatação
  const editor = editorEl();
  editor.addEventListener("input", atualizarVarsDetetadas);
  editor.addEventListener("paste", colarNoEditor);
  document.addEventListener("selectionchange", guardarRangeEditor);

  // mousedown com preventDefault mantém a seleção no editor ao clicar nos botões
  const manterSelecao = el => el.addEventListener("mousedown", e => e.preventDefault());

  document.querySelectorAll(".editor-toolbar [data-cmd]").forEach(btn => {
    manterSelecao(btn);
    btn.addEventListener("click", () => executarComando(btn.dataset.cmd));
  });
  manterSelecao(porId("btn-link"));
  porId("btn-link").addEventListener("click", abrirLinhaLink);
  porId("btn-link-ok").addEventListener("click", aplicarLink);
  porId("btn-link-cancelar").addEventListener("click", fecharLinhaLink);
  porId("link-url").addEventListener("keydown", e => {
    if (e.key === "Enter") { e.preventDefault(); aplicarLink(); }
    if (e.key === "Escape") fecharLinhaLink();
  });

  document.querySelectorAll(".btn-var-tag").forEach(btn => {
    manterSelecao(btn);
    btn.addEventListener("click", () => inserirTagVariavel(btn.dataset.var));
  });

  // Vista Usar
  porId("usar-anexo-extra").addEventListener("change", e => lidarComAnexosExtra(e.target.files));
  porId("usar-campos").addEventListener("input", atualizarPreview);
  porId("usar-campos").addEventListener("change", atualizarPreview);

  // Categorias
  porId("btn-gerir-categorias").addEventListener("click", abrirModalCategorias);
  porId("btn-adicionar-categoria").addEventListener("click", adicionarNovaCategoria);
  porId("novo-nome-categoria").addEventListener("keydown", e => {
    if (e.key === "Enter") adicionarNovaCategoria();
  });
  porId("btn-fechar-modal-categorias").addEventListener("click", () => {
    porId("modal-categorias").classList.add("oculto");
  });
  porId("lista-gestao-categorias").addEventListener("click", e => {
    const btn = e.target.closest("button.btn-eliminar-cat");
    if (btn && btn.dataset.cat) eliminarCategoria(btn.dataset.cat);
  });

  // Pesquisa, filtro e ordenação
  porId("pesquisa").addEventListener("input", renderLista);
  porId("filtro-categoria").addEventListener("change", renderLista);
  porId("ordenacao").addEventListener("change", e => {
    prefs.ordenacao = e.target.value;
    guardarPrefs();
    renderLista();
  });
  porId("btn-limpar-pesquisa").addEventListener("click", () => {
    porId("pesquisa").value = "";
    renderLista();
  });

  // Cartões da lista
  porId("lista-templates").addEventListener("click", e => {
    const btn = e.target.closest("button[data-acao]");
    if (!btn) return;
    const id = btn.dataset.id;
    const acoes = {
      usar: abrirUsarTemplate,
      editar: abrirEditorExistente,
      duplicar: duplicarTemplate,
      favorito: alternarFavorito
    };
    if (acoes[btn.dataset.acao]) acoes[btn.dataset.acao](id);
  });

  // Exportar / Importar
  porId("btn-exportar").addEventListener("click", exportarTemplates);
  porId("btn-importar").addEventListener("click", () => {
    const input = porId("input-importar");
    input.value = "";
    input.click();
  });
  porId("input-importar").addEventListener("change", e => {
    if (e.target.files[0]) processarFicheiroImportar(e.target.files[0]);
  });
  porId("btn-importar-unir").addEventListener("click", () => executarImportacao("unir"));
  porId("btn-importar-substituir").addEventListener("click", () => executarImportacao("substituir"));
  porId("btn-importar-cancelar").addEventListener("click", () => {
    backupImportarTemp = null;
    porId("modal-importar").classList.add("oculto");
  });
}

/* =====================================================================
   Arranque — no fim do ficheiro: o Office.onReady pode chamar a função
   de imediato, e todas as variáveis acima já têm de estar declaradas.
   ===================================================================== */

Office.onReady(async () => {
  ligarEventos();
  try {
    await carregarTudo();
  } catch (e) {
    console.error("Falha ao carregar dados:", e);
    mostrarEstado("Erro ao carregar os templates guardados.", true);
  }
  document.getElementById("ordenacao").value = prefs.ordenacao;
  atualizarDropdownCategorias();
  renderLista();
});
