/* Os Meus Templates — gestor pessoal de templates de email para Outlook (Office.js) */

const STORAGE_KEY = "meusTemplates.v1";
const CATEGORIAS_KEY = "meusTemplates.categorias.v1";

let templates = [];
let categoriasCustom = [];
let templateEmEdicao = null; // id do template a editar, ou null se novo
let anexoTemp = null; // { nome, tipoConteudo, base64 } carregado no editor
let templateAUsar = null;
let anexosExtraUso = []; // [{ nome, tipoConteudo, base64 }] escolhidos no momento de usar o template
let backupImportarTemp = null; // Guardar temporariamente os templates a importar para o modal

const CATEGORIAS_PADRAO = ["Geral", "Comercial", "Suporte", "Propostas"];

Office.onReady(() => {
  carregarTemplates();
  carregarCategorias();
  atualizarDropdownCategorias();
  renderLista();
  ligarEventos();
});

/* ---------- Armazenamento local ---------- */

function carregarTemplates() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    templates = raw ? JSON.parse(raw) : [];
    // Normalizar templates
    templates.forEach(t => {
      if (!t.categoria) t.categoria = "Geral";
      if (!t.para) t.para = "";
      if (!t.cc) t.cc = "";
      if (!t.bcc) t.bcc = "";
    });
  } catch (e) {
    templates = [];
  }
}

function guardarTemplates() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(templates));
}

function carregarCategorias() {
  try {
    const raw = localStorage.getItem(CATEGORIAS_KEY);
    if (raw) {
      categoriasCustom = JSON.parse(raw);
    } else {
      categoriasCustom = [...CATEGORIAS_PADRAO];
    }
  } catch (e) {
    categoriasCustom = [...CATEGORIAS_PADRAO];
  }
}

function guardarCategorias() {
  localStorage.setItem(CATEGORIAS_KEY, JSON.stringify(categoriasCustom));
}

function obterCategoriasUnicas() {
  const setCat = new Set(["Geral", ...CATEGORIAS_PADRAO, ...categoriasCustom]);
  templates.forEach(t => {
    if (t.categoria && t.categoria.trim()) {
      setCat.add(t.categoria.trim());
    }
  });
  return Array.from(setCat).sort();
}

function atualizarDropdownCategorias() {
  const categorias = obterCategoriasUnicas();

  // Atualizar filtro da lista
  const selectFiltro = document.getElementById("filtro-categoria");
  const valorAtualFiltro = selectFiltro.value;
  selectFiltro.innerHTML = `<option value="">Todas as categorias</option>`;
  categorias.forEach(cat => {
    const opt = document.createElement("option");
    opt.value = cat;
    opt.textContent = cat;
    selectFiltro.appendChild(opt);
  });
  selectFiltro.value = valorAtualFiltro;

  // Atualizar datalist de sugestões no editor
  const datalist = document.getElementById("lista-categorias-sugeridas");
  datalist.innerHTML = "";
  categorias.forEach(cat => {
    const opt = document.createElement("option");
    opt.value = cat;
    datalist.appendChild(opt);
  });
}

/* ---------- Gestão de Categorias ---------- */

function abrirModalCategorias() {
  renderListaGestaoCategorias();
  document.getElementById("novo-nome-categoria").value = "";
  document.getElementById("modal-categorias").classList.remove("oculto");
}

function renderListaGestaoCategorias() {
  const container = document.getElementById("lista-gestao-categorias");
  container.innerHTML = "";

  const categorias = obterCategoriasUnicas();

  categorias.forEach(cat => {
    const count = templates.filter(t => t.categoria === cat).length;
    const isGeral = cat.toLowerCase() === "geral";

    const item = document.createElement("div");
    item.className = "item-gestao-categoria";
    item.innerHTML = `
      <div class="item-gestao-categoria-info">
        <strong>${escapeHtml(cat)}</strong>
        <span class="item-gestao-categoria-count">(${count} template${count === 1 ? '' : 's'})</span>
      </div>
      ${!isGeral ? `<button class="btn-eliminar-cat" data-cat="${escapeHtml(cat)}" title="Eliminar categoria">🗑️</button>` : `<span class="item-gestao-categoria-count">(padrão)</span>`}
    `;
    container.appendChild(item);
  });
}

function adicionarNovaCategoria() {
  const input = document.getElementById("novo-nome-categoria");
  const nome = input.value.trim();

  if (!nome) {
    mostrarEstado("Escreve o nome da categoria.", true);
    return;
  }

  const existentes = obterCategoriasUnicas().map(c => c.toLowerCase());
  if (existentes.includes(nome.toLowerCase())) {
    mostrarEstado("Essa categoria já existe.", true);
    return;
  }

  categoriasCustom.push(nome);
  guardarCategorias();
  atualizarDropdownCategorias();
  renderListaGestaoCategorias();
  input.value = "";
  mostrarEstado(`Categoria "${nome}" criada.`);
}

function eliminarCategoria(nomeCat) {
  if (nomeCat.toLowerCase() === "geral") {
    mostrarEstado("A categoria Geral não pode ser eliminada.", true);
    return;
  }

  let afetados = 0;
  templates.forEach(t => {
    if (t.categoria === nomeCat) {
      t.categoria = "Geral";
      afetados++;
    }
  });

  categoriasCustom = categoriasCustom.filter(c => c !== nomeCat);

  guardarTemplates();
  guardarCategorias();
  atualizarDropdownCategorias();
  renderListaGestaoCategorias();
  renderLista();

  if (afetados > 0) {
    mostrarEstado(`Categoria eliminada. ${afetados} template(s) movido(s) para "Geral".`);
  } else {
    mostrarEstado(`Categoria "${nomeCat}" eliminada.`);
  }
}

/* ---------- Navegação entre vistas ---------- */

function mostrarVista(vista) {
  document.getElementById("lista-vista").classList.toggle("oculto", vista !== "lista");
  document.getElementById("editor-vista").classList.toggle("oculto", vista !== "editor");
  document.getElementById("usar-vista").classList.toggle("oculto", vista !== "usar");
}

/* ---------- Lista de templates ---------- */

function renderLista() {
  const container = document.getElementById("lista-templates");
  container.innerHTML = "";

  const termoPesquisa = document.getElementById("pesquisa").value.trim().toLowerCase();
  const categoriaFiltro = document.getElementById("filtro-categoria").value;
  const btnLimpar = document.getElementById("btn-limpar-pesquisa");

  if (btnLimpar) {
    btnLimpar.classList.toggle("oculto", !termoPesquisa);
  }

  const filtrados = templates.filter(t => {
    const correspondeTexto = !termoPesquisa ||
      t.nome.toLowerCase().includes(termoPesquisa) ||
      t.assunto.toLowerCase().includes(termoPesquisa) ||
      (t.corpo && t.corpo.toLowerCase().includes(termoPesquisa)) ||
      (t.categoria && t.categoria.toLowerCase().includes(termoPesquisa)) ||
      (t.para && t.para.toLowerCase().includes(termoPesquisa)) ||
      (t.cc && t.cc.toLowerCase().includes(termoPesquisa)) ||
      (t.bcc && t.bcc.toLowerCase().includes(termoPesquisa));

    const correspondeCategoria = !categoriaFiltro || t.categoria === categoriaFiltro;

    return correspondeTexto && correspondeCategoria;
  });

  if (filtrados.length === 0) {
    if (templates.length === 0) {
      container.innerHTML = `<p class="ajuda">Sem templates. Cria o primeiro com "+ Novo".</p>`;
    } else {
      container.innerHTML = `<p class="ajuda">Nenhum template encontrado para os filtros selecionados.</p>`;
    }
    return;
  }

  filtrados.forEach(t => {
    const card = document.createElement("div");
    card.className = "cartao-template";
    card.innerHTML = `
      <div class="cartao-header">
        <h3>${escapeHtml(t.nome)}</h3>
        <span class="badge-categoria">${escapeHtml(t.categoria || "Geral")}</span>
      </div>
      <p>${escapeHtml(t.assunto || "(Sem assunto)")}</p>
      <div class="cartao-acoes">
        <button class="btn btn-primary" data-acao="usar" data-id="${t.id}">Usar</button>
        <button class="btn btn-secondary" data-acao="editar" data-id="${t.id}">Editar</button>
      </div>
    `;
    container.appendChild(card);
  });
}

/* ---------- Editor de template ---------- */

function abrirEditorNovo() {
  templateEmEdicao = null;
  anexoTemp = null;
  document.getElementById("editor-titulo").textContent = "Novo template";
  document.getElementById("edit-nome").value = "";
  document.getElementById("edit-categoria").value = "Geral";
  document.getElementById("edit-para").value = "";
  document.getElementById("edit-cc").value = "";
  document.getElementById("edit-bcc").value = "";
  document.getElementById("edit-assunto").value = "";
  document.getElementById("edit-corpo").value = "";
  document.getElementById("edit-anexo").value = "";
  document.getElementById("anexo-atual").textContent = "";
  document.getElementById("btn-eliminar").classList.add("oculto");
  atualizarDropdownCategorias();
  atualizarVarsDetetadas();
  mostrarVista("editor");
}

function abrirEditorExistente(id) {
  const t = templates.find(x => x.id === id);
  if (!t) return;
  templateEmEdicao = id;
  anexoTemp = t.anexo || null;
  document.getElementById("editor-titulo").textContent = "Editar template";
  document.getElementById("edit-nome").value = t.nome;
  document.getElementById("edit-categoria").value = t.categoria || "Geral";
  document.getElementById("edit-para").value = t.para || "";
  document.getElementById("edit-cc").value = t.cc || "";
  document.getElementById("edit-bcc").value = t.bcc || "";
  document.getElementById("edit-assunto").value = t.assunto;
  document.getElementById("edit-corpo").value = t.corpo;
  document.getElementById("edit-anexo").value = "";
  document.getElementById("anexo-atual").textContent = t.anexo ? `Anexo atual: ${t.anexo.nome}` : "";
  document.getElementById("btn-eliminar").classList.remove("oculto");
  atualizarDropdownCategorias();
  atualizarVarsDetetadas();
  mostrarVista("editor");
}

function extrairVariaveis(texto) {
  const regex = /{{\s*([a-zA-Z0-9_]+)\s*}}/g;
  const encontradas = new Set();
  let m;
  while ((m = regex.exec(texto)) !== null) {
    encontradas.add(m[1]);
  }
  return [...encontradas];
}

function atualizarVarsDetetadas() {
  const assunto = document.getElementById("edit-assunto").value;
  const corpo = document.getElementById("edit-corpo").value;
  const para = document.getElementById("edit-para").value;
  const cc = document.getElementById("edit-cc").value;
  const bcc = document.getElementById("edit-bcc").value;

  const vars = extrairVariaveis(`${assunto} ${corpo} ${para} ${cc} ${bcc}`);
  document.getElementById("vars-detetadas").textContent = vars.length ? vars.join(", ") : "nenhuma";
}

function inserirTagVariavel(varNome) {
  const textarea = document.getElementById("edit-corpo");
  const tag = `{{${varNome}}}`;
  const start = textarea.selectionStart;
  const end = textarea.selectionEnd;
  const val = textarea.value;

  if (start !== undefined && end !== undefined) {
    textarea.value = val.substring(0, start) + tag + val.substring(end);
    textarea.selectionStart = textarea.selectionEnd = start + tag.length;
  } else {
    textarea.value += " " + tag;
  }
  textarea.focus();
  atualizarVarsDetetadas();
}

function guardarTemplateAtual() {
  const nome = document.getElementById("edit-nome").value.trim();
  const categoria = document.getElementById("edit-categoria").value.trim() || "Geral";
  const para = document.getElementById("edit-para").value.trim();
  const cc = document.getElementById("edit-cc").value.trim();
  const bcc = document.getElementById("edit-bcc").value.trim();
  const assunto = document.getElementById("edit-assunto").value.trim();
  const corpo = document.getElementById("edit-corpo").value;

  if (!nome) {
    mostrarEstado("Dá um nome ao template.", true);
    return;
  }

  if (!categoriasCustom.includes(categoria) && categoria !== "Geral") {
    categoriasCustom.push(categoria);
    guardarCategorias();
  }

  if (templateEmEdicao) {
    const t = templates.find(x => x.id === templateEmEdicao);
    t.nome = nome;
    t.categoria = categoria;
    t.para = para;
    t.cc = cc;
    t.bcc = bcc;
    t.assunto = assunto;
    t.corpo = corpo;
    t.anexo = anexoTemp;
  } else {
    templates.push({
      id: "t_" + Date.now(),
      nome, categoria, para, cc, bcc, assunto, corpo,
      anexo: anexoTemp
    });
  }

  guardarTemplates();
  atualizarDropdownCategorias();
  renderLista();
  mostrarVista("lista");
  mostrarEstado("Template guardado.");
}

function eliminarTemplateAtual() {
  if (!templateEmEdicao) return;
  templates = templates.filter(x => x.id !== templateEmEdicao);
  guardarTemplates();
  atualizarDropdownCategorias();
  renderLista();
  mostrarVista("lista");
  mostrarEstado("Template eliminado.");
}

function lidarComAnexoSelecionado(ficheiro) {
  if (!ficheiro) return;
  const reader = new FileReader();
  reader.onload = () => {
    const base64 = reader.result.split(",")[1];
    anexoTemp = { nome: ficheiro.name, tipoConteudo: ficheiro.type || "application/octet-stream", base64 };
    document.getElementById("anexo-atual").textContent = `Anexo selecionado: ${ficheiro.name}`;
  };
  reader.readAsDataURL(ficheiro);
}

/* ---------- Resolução de Variáveis Pré-definidas Dinâmicas ---------- */

function obterSaudacaoTempo() {
  const agora = new Date();
  const hora = agora.getHours();
  const min = agora.getMinutes();
  const minTotais = hora * 60 + min;

  // 05:00 até às 13:49 (inclusive) -> Bom dia
  if (minTotais >= 5 * 60 && minTotais <= 13 * 60 + 49) {
    return "Bom dia, ";
  }
  // 13:50 até às 19:59 (inclusive) -> Boa tarde
  if (minTotais >= 13 * 60 + 50 && minTotais < 20 * 60) {
    return "Boa tarde, ";
  }
  // 20:00 às 04:59 -> Boa noite
  return "Boa noite, ";
}

function obterValoresPreDefinidosAsync() {
  return new Promise((resolve) => {
    const agora = new Date();
    const dataFormatted = agora.toLocaleDateString("pt-PT");
    const horaFormatted = agora.toLocaleTimeString("pt-PT", { hour: "2-digit", minute: "2-digit" });
    const cumprimentoCalculado = obterSaudacaoTempo();

    const valores = {
      cumprimento: cumprimentoCalculado,
      saudacao: cumprimentoCalculado,
      saudacao_tempo: cumprimentoCalculado,
      bom_dia_boa_tarde: cumprimentoCalculado,
      data_atual: dataFormatted,
      data: dataFormatted,
      hora_atual: horaFormatted,
      hora: horaFormatted,
      remetente_nome: "",
      remetente_email: "",
      destinatario_nome: "",
      destinatario_email: ""
    };

    try {
      if (Office.context && Office.context.mailbox && Office.context.mailbox.userProfile) {
        valores.remetente_nome = Office.context.mailbox.userProfile.displayName || "";
        valores.remetente_email = Office.context.mailbox.userProfile.emailAddress || "";
      }
    } catch (e) {
      console.warn("Não foi possível obter o perfil do utilizador:", e);
    }

    try {
      const item = Office.context && Office.context.mailbox && Office.context.mailbox.item;
      if (item && item.to && typeof item.to.getAsync === "function") {
        item.to.getAsync((res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded && res.value && res.value.length > 0) {
            const primeiroRec = res.value[0];
            valores.destinatario_nome = primeiroRec.displayName || primeiroRec.emailAddress || "";
            valores.destinatario_email = primeiroRec.emailAddress || "";
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

/* ---------- Usar template ---------- */

async function abrirUsarTemplate(id) {
  const t = templates.find(x => x.id === id);
  if (!t) return;
  templateAUsar = t;

  document.getElementById("usar-titulo").textContent = `Usar: ${t.nome}`;

  // Preencher destinatários (Para / CC / BCC)
  document.getElementById("usar-para").value = t.para || "";
  document.getElementById("usar-cc").value = t.cc || "";
  document.getElementById("usar-bcc").value = t.bcc || "";

  const todoOTexto = `${t.assunto} ${t.corpo} ${t.para || ""} ${t.cc || ""} ${t.bcc || ""}`;
  const vars = extrairVariaveis(todoOTexto);

  const container = document.getElementById("usar-campos");
  container.innerHTML = "<p class='ajuda'>A carregar dados do email…</p>";

  anexosExtraUso = [];
  document.getElementById("usar-anexo-extra").value = "";
  document.getElementById("usar-anexo-extra-lista").textContent = "";

  const anexoFixoEl = document.getElementById("usar-anexo-fixo");
  if (t.anexo) {
    anexoFixoEl.textContent = `Anexo fixo deste template: ${t.anexo.nome}`;
    anexoFixoEl.classList.remove("oculto");
  } else {
    anexoFixoEl.classList.add("oculto");
  }

  const valoresPreDefinidos = await obterValoresPreDefinidosAsync();

  container.innerHTML = "";

  if (vars.length === 0) {
    container.innerHTML = `<p class="ajuda">Este template não tem variáveis de substituição. Clica em "Inserir no email" para aplicar.</p>`;
  } else {
    vars.forEach(v => {
      const div = document.createElement("div");
      div.className = "campo-variavel";

      const vKeyLower = v.toLowerCase();
      const valDefault = valoresPreDefinidos[vKeyLower] !== undefined ? valoresPreDefinidos[vKeyLower] : "";
      const isAuto = Boolean(valDefault);

      div.innerHTML = `
        <div class="campo-variavel-header">
          <label for="var-field-${escapeHtml(v)}">${escapeHtml(v)}</label>
          ${isAuto ? `<span class="badge-auto">Preenchido auto</span>` : ""}
        </div>
        <input id="var-field-${escapeHtml(v)}" type="text" data-var="${escapeHtml(v)}" value="${escapeHtml(valDefault)}" placeholder="Valor para {{${escapeHtml(v)}}}" />
      `;
      container.appendChild(div);
    });
  }

  mostrarVista("usar");
}

function lidarComAnexosExtra(ficheiros) {
  anexosExtraUso = [];
  const lista = document.getElementById("usar-anexo-extra-lista");
  lista.textContent = "A carregar…";

  const leituras = Array.from(ficheiros).map(ficheiro => new Promise(resolve => {
    const reader = new FileReader();
    reader.onload = () => {
      const base64 = reader.result.split(",")[1];
      resolve({ nome: ficheiro.name, tipoConteudo: ficheiro.type || "application/octet-stream", base64 });
    };
    reader.readAsDataURL(ficheiro);
  }));

  Promise.all(leituras).then(resultados => {
    anexosExtraUso = resultados;
    lista.textContent = resultados.length
      ? `A anexar: ${resultados.map(a => a.nome).join(", ")}`
      : "";
  });
}

function substituirVariaveis(texto, valores) {
  return (texto || "").replace(/{{\s*([a-zA-Z0-9_]+)\s*}}/g, (match, nome) => {
    return valores[nome] !== undefined ? valores[nome] : match;
  });
}

function parseEmailsParaDestinatarios(str) {
  if (!str || !str.trim()) return [];
  return str.split(/[;,]/)
    .map(email => email.trim())
    .filter(email => email.length > 0)
    .map(email => ({ emailAddress: email }));
}

function inserirTemplateNoEmail() {
  if (!templateAUsar) return;

  const inputs = document.querySelectorAll("#usar-campos input[data-var]");
  const valores = {};
  inputs.forEach(inp => { valores[inp.dataset.var] = inp.value; });

  const assuntoFinal = substituirVariaveis(templateAUsar.assunto, valores);
  const corpoFinal = substituirVariaveis(templateAUsar.corpo, valores);
  const corpoHtml = corpoFinal.replace(/\n/g, "<br>");

  const paraFinal = substituirVariaveis(document.getElementById("usar-para").value, valores);
  const ccFinal = substituirVariaveis(document.getElementById("usar-cc").value, valores);
  const bccFinal = substituirVariaveis(document.getElementById("usar-bcc").value, valores);

  const paraRecipients = parseEmailsParaDestinatarios(paraFinal);
  const ccRecipients = parseEmailsParaDestinatarios(ccFinal);
  const bccRecipients = parseEmailsParaDestinatarios(bccFinal);

  const item = Office.context.mailbox.item;

  const tarefas = [];

  // Assunto
  tarefas.push(new Promise((resolve, reject) => {
    if (!assuntoFinal) { resolve(); return; }
    item.subject.setAsync(assuntoFinal, res => {
      res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
    });
  }));

  // Corpo
  tarefas.push(new Promise((resolve, reject) => {
    item.body.setSelectedDataAsync(corpoHtml, { coercionType: Office.CoercionType.Html }, res => {
      res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
    });
  }));

  // Destinatários Para
  if (paraRecipients.length > 0 && item && item.to && typeof item.to.addAsync === "function") {
    tarefas.push(new Promise((resolve, reject) => {
      item.to.addAsync(paraRecipients, res => {
        res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
      });
    }));
  }

  // Destinatários CC
  if (ccRecipients.length > 0 && item && item.cc && typeof item.cc.addAsync === "function") {
    tarefas.push(new Promise((resolve, reject) => {
      item.cc.addAsync(ccRecipients, res => {
        res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
      });
    }));
  }

  // Destinatários BCC
  if (bccRecipients.length > 0 && item && item.bcc && typeof item.bcc.addAsync === "function") {
    tarefas.push(new Promise((resolve, reject) => {
      item.bcc.addAsync(bccRecipients, res => {
        res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
      });
    }));
  }

  const todosAnexos = [];
  if (templateAUsar.anexo) todosAnexos.push(templateAUsar.anexo);
  todosAnexos.push(...anexosExtraUso);

  function anexarUmAUm(lista) {
    if (lista.length === 0) return Promise.resolve();
    const [anexo, ...resto] = lista;
    return new Promise((resolve, reject) => {
      item.addFileAttachmentFromBase64Async(
        anexo.base64,
        anexo.nome,
        { isInline: false },
        res => {
          res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
        }
      );
    }).then(() => anexarUmAUm(resto));
  }

  Promise.all(tarefas)
    .then(() => anexarUmAUm(todosAnexos))
    .then(() => {
      mostrarEstado("Template e destinatários inseridos no email!");
      mostrarVista("lista");
    })
    .catch(err => {
      console.error(err);
      mostrarEstado("Erro ao inserir template. Garante que estás a compor um email.", true);
    });
}

/* ---------- Exportar / Importar Backups ---------- */

function exportarTemplates() {
  if (templates.length === 0) {
    mostrarEstado("Não existem templates para exportar.", true);
    return;
  }

  const payload = {
    versao: "1.2",
    dataExportacao: new Date().toISOString(),
    totalTemplates: templates.length,
    categoriasCustom: categoriasCustom,
    templates: templates
  };

  const jsonStr = JSON.stringify(payload, null, 2);
  const blob = new Blob([jsonStr], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `os-meus-templates-backup-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
  mostrarEstado("Backup exportado com sucesso.");
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
        if (Array.isArray(conteudo.categoriasCustom)) {
          catsImportadas = conteudo.categoriasCustom;
        }
      } else {
        throw new Error("Formato inválido");
      }

      if (listaImportada.length === 0) {
        mostrarEstado("O ficheiro de backup não contém nenhum template validado.", true);
        return;
      }

      backupImportarTemp = {
        templates: listaImportada.map(t => ({
          id: "t_" + Date.now() + "_" + Math.random().toString(36).slice(2, 7),
          nome: t.nome || "Template Sem Nome",
          categoria: t.categoria || "Geral",
          para: t.para || "",
          cc: t.cc || "",
          bcc: t.bcc || "",
          assunto: t.assunto || "",
          corpo: t.corpo || "",
          anexo: t.anexo || null
        })),
        categoriasCustom: catsImportadas
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

function executarImportacao(modo) {
  if (!backupImportarTemp) return;

  if (modo === "unir") {
    templates = templates.concat(backupImportarTemp.templates);
    backupImportarTemp.categoriasCustom.forEach(c => {
      if (!categoriasCustom.includes(c)) categoriasCustom.push(c);
    });
    mostrarEstado(`${backupImportarTemp.templates.length} template(s) unidos com sucesso!`);
  } else if (modo === "substituir") {
    templates = backupImportarTemp.templates;
    if (backupImportarTemp.categoriasCustom.length) {
      categoriasCustom = backupImportarTemp.categoriasCustom;
    }
    mostrarEstado(`Biblioteca substituída com ${backupImportarTemp.templates.length} template(s).`);
  }

  backupImportarTemp = null;
  document.getElementById("modal-importar").classList.add("oculto");

  guardarTemplates();
  guardarCategorias();
  atualizarDropdownCategorias();
  renderLista();
}

/* ---------- Utilitários ---------- */

function escapeHtml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

let estadoTimeout = null;
function mostrarEstado(msg, erro = false) {
  const el = document.getElementById("estado-msg");
  el.textContent = msg;
  el.classList.toggle("erro", erro);
  el.classList.remove("oculto");
  clearTimeout(estadoTimeout);
  estadoTimeout = setTimeout(() => el.classList.add("oculto"), 3500);
}

/* ---------- Eventos ---------- */

function ligarEventos() {
  document.getElementById("btn-novo").addEventListener("click", abrirEditorNovo);
  document.getElementById("btn-cancelar").addEventListener("click", () => mostrarVista("lista"));
  document.getElementById("btn-guardar").addEventListener("click", guardarTemplateAtual);
  document.getElementById("btn-eliminar").addEventListener("click", eliminarTemplateAtual);
  document.getElementById("btn-usar-cancelar").addEventListener("click", () => mostrarVista("lista"));
  document.getElementById("btn-inserir").addEventListener("click", inserirTemplateNoEmail);

  document.getElementById("edit-assunto").addEventListener("input", atualizarVarsDetetadas);
  document.getElementById("edit-corpo").addEventListener("input", atualizarVarsDetetadas);
  document.getElementById("edit-para").addEventListener("input", atualizarVarsDetetadas);
  document.getElementById("edit-cc").addEventListener("input", atualizarVarsDetetadas);
  document.getElementById("edit-bcc").addEventListener("input", atualizarVarsDetetadas);
  document.getElementById("edit-anexo").addEventListener("change", e => lidarComAnexoSelecionado(e.target.files[0]));
  document.getElementById("usar-anexo-extra").addEventListener("change", e => lidarComAnexosExtra(e.target.files));

  // Gestão de Categorias Modal
  document.getElementById("btn-gerir-categorias").addEventListener("click", abrirModalCategorias);
  document.getElementById("btn-adicionar-categoria").addEventListener("click", adicionarNovaCategoria);
  document.getElementById("novo-nome-categoria").addEventListener("keypress", e => {
    if (e.key === "Enter") adicionarNovaCategoria();
  });
  document.getElementById("btn-fechar-modal-categorias").addEventListener("click", () => {
    document.getElementById("modal-categorias").classList.add("oculto");
  });

  document.getElementById("lista-gestao-categorias").addEventListener("click", e => {
    const btn = e.target.closest("button.btn-eliminar-cat");
    if (!btn) return;
    const cat = btn.dataset.cat;
    if (cat) eliminarCategoria(cat);
  });

  // Eventos de atalho para inserir variáveis no corpo
  document.querySelectorAll(".btn-var-tag").forEach(btn => {
    btn.addEventListener("click", e => {
      const varNome = e.target.getAttribute("data-var");
      if (varNome) inserirTagVariavel(varNome);
    });
  });

  // Pesquisa e filtro por categoria
  document.getElementById("pesquisa").addEventListener("input", renderLista);
  document.getElementById("filtro-categoria").addEventListener("change", renderLista);

  const btnLimpar = document.getElementById("btn-limpar-pesquisa");
  if (btnLimpar) {
    btnLimpar.addEventListener("click", () => {
      document.getElementById("pesquisa").value = "";
      renderLista();
    });
  }

  // Cliques na lista (Usar / Editar)
  document.getElementById("lista-templates").addEventListener("click", e => {
    const btn = e.target.closest("button[data-acao]");
    if (!btn) return;
    const id = btn.dataset.id;
    if (btn.dataset.acao === "usar") abrirUsarTemplate(id);
    if (btn.dataset.acao === "editar") abrirEditorExistente(id);
  });

  // Exportar / Importar com Modal
  document.getElementById("btn-exportar").addEventListener("click", exportarTemplates);
  document.getElementById("btn-importar").addEventListener("click", () => {
    const input = document.getElementById("input-importar");
    input.value = "";
    input.click();
  });

  document.getElementById("input-importar").addEventListener("change", e => {
    if (e.target.files[0]) processarFicheiroImportar(e.target.files[0]);
  });

  document.getElementById("btn-importar-unir").addEventListener("click", () => executarImportacao("unir"));
  document.getElementById("btn-importar-substituir").addEventListener("click", () => executarImportacao("substituir"));
  document.getElementById("btn-importar-cancelar").addEventListener("click", () => {
    backupImportarTemp = null;
    document.getElementById("modal-importar").classList.add("oculto");
  });
}
