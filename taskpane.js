/* Meus Templates — gestor pessoal de templates de email para Outlook (Office.js) */

const STORAGE_KEY = "meusTemplates.v1";
let templates = [];
let templateEmEdicao = null; // id do template a editar, ou null se novo
let anexoTemp = null; // { nome, tipoConteudo, base64 } carregado no editor
let templateAUsar = null;
let anexosExtraUso = []; // [{ nome, tipoConteudo, base64 }] escolhidos no momento de usar o template

Office.onReady(() => {
  carregarTemplates();
  renderLista();
  ligarEventos();
});

/* ---------- Armazenamento local ---------- */

function carregarTemplates() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    templates = raw ? JSON.parse(raw) : [];
  } catch (e) {
    templates = [];
  }
}

function guardarTemplates() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(templates));
}

/* ---------- Navegação entre vistas ---------- */

function mostrarVista(vista) {
  document.getElementById("lista-vista").classList.toggle("oculto", vista !== "lista");
  document.getElementById("editor-vista").classList.toggle("oculto", vista !== "editor");
  document.getElementById("usar-vista").classList.toggle("oculto", vista !== "usar");
}

/* ---------- Lista de templates ---------- */

function renderLista(filtro = "") {
  const container = document.getElementById("lista-templates");
  container.innerHTML = "";

  const filtrados = templates.filter(t =>
    t.nome.toLowerCase().includes(filtro.toLowerCase()) ||
    t.assunto.toLowerCase().includes(filtro.toLowerCase())
  );

  if (filtrados.length === 0) {
    container.innerHTML = `<p class="ajuda">Sem templates. Cria o primeiro com "+ Novo".</p>`;
    return;
  }

  filtrados.forEach(t => {
    const card = document.createElement("div");
    card.className = "cartao-template";
    card.innerHTML = `
      <h3>${escapeHtml(t.nome)}</h3>
      <p>${escapeHtml(t.assunto)}</p>
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
  document.getElementById("edit-assunto").value = "";
  document.getElementById("edit-corpo").value = "";
  document.getElementById("edit-anexo").value = "";
  document.getElementById("anexo-atual").textContent = "";
  document.getElementById("btn-eliminar").classList.add("oculto");
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
  document.getElementById("edit-assunto").value = t.assunto;
  document.getElementById("edit-corpo").value = t.corpo;
  document.getElementById("edit-anexo").value = "";
  document.getElementById("anexo-atual").textContent = t.anexo ? `Anexo atual: ${t.anexo.nome}` : "";
  document.getElementById("btn-eliminar").classList.remove("oculto");
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
  const vars = extrairVariaveis(assunto + " " + corpo);
  document.getElementById("vars-detetadas").textContent = vars.length ? vars.join(", ") : "nenhuma";
}

function guardarTemplateAtual() {
  const nome = document.getElementById("edit-nome").value.trim();
  const assunto = document.getElementById("edit-assunto").value.trim();
  const corpo = document.getElementById("edit-corpo").value;

  if (!nome) {
    mostrarEstado("Dá um nome ao template.", true);
    return;
  }

  if (templateEmEdicao) {
    const t = templates.find(x => x.id === templateEmEdicao);
    t.nome = nome;
    t.assunto = assunto;
    t.corpo = corpo;
    t.anexo = anexoTemp;
  } else {
    templates.push({
      id: "t_" + Date.now(),
      nome, assunto, corpo,
      anexo: anexoTemp
    });
  }

  guardarTemplates();
  renderLista(document.getElementById("pesquisa").value);
  mostrarVista("lista");
  mostrarEstado("Template guardado.");
}

function eliminarTemplateAtual() {
  if (!templateEmEdicao) return;
  templates = templates.filter(x => x.id !== templateEmEdicao);
  guardarTemplates();
  renderLista(document.getElementById("pesquisa").value);
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

/* ---------- Usar template ---------- */

function abrirUsarTemplate(id) {
  const t = templates.find(x => x.id === id);
  if (!t) return;
  templateAUsar = t;

  const vars = extrairVariaveis(t.assunto + " " + t.corpo);
  document.getElementById("usar-titulo").textContent = `Usar: ${t.nome}`;
  const container = document.getElementById("usar-campos");
  container.innerHTML = "";

  // reiniciar anexos escolhidos na sessão anterior
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

  if (vars.length === 0) {
    container.innerHTML = `<p class="ajuda">Este template não tem variáveis. Basta inserir.</p>`;
  } else {
    vars.forEach(v => {
      const div = document.createElement("div");
      div.className = "campo-variavel";
      div.innerHTML = `
        <label>${escapeHtml(v)}</label>
        <input type="text" data-var="${escapeHtml(v)}" placeholder="Valor para {{${escapeHtml(v)}}}" />
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
  return texto.replace(/{{\s*([a-zA-Z0-9_]+)\s*}}/g, (match, nome) => {
    return valores[nome] !== undefined ? valores[nome] : match;
  });
}

function inserirTemplateNoEmail() {
  if (!templateAUsar) return;

  const inputs = document.querySelectorAll("#usar-campos input[data-var]");
  const valores = {};
  inputs.forEach(inp => { valores[inp.dataset.var] = inp.value; });

  const assuntoFinal = substituirVariaveis(templateAUsar.assunto, valores);
  const corpoFinal = substituirVariaveis(templateAUsar.corpo, valores);
  const corpoHtml = corpoFinal.replace(/\n/g, "<br>");

  const item = Office.context.mailbox.item;

  const tarefas = [];

  tarefas.push(new Promise((resolve, reject) => {
    if (!assuntoFinal) { resolve(); return; }
    item.subject.setAsync(assuntoFinal, res => {
      res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
    });
  }));

  tarefas.push(new Promise((resolve, reject) => {
    item.body.setSelectedDataAsync(corpoHtml, { coercionType: Office.CoercionType.Html }, res => {
      res.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(res.error);
    });
  }));

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
      mostrarEstado("Template inserido no email.");
      mostrarVista("lista");
    })
    .catch(err => {
      console.error(err);
      mostrarEstado("Não foi possível inserir o template. Verifica se estás numa janela de composição.", true);
    });
}

/* ---------- Importar / Exportar ---------- */

function exportarTemplates() {
  const blob = new Blob([JSON.stringify(templates, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "meus-templates.json";
  a.click();
  URL.revokeObjectURL(url);
}

function importarTemplates(ficheiro) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const importados = JSON.parse(reader.result);
      if (!Array.isArray(importados)) throw new Error("formato inválido");
      importados.forEach(t => {
        if (!t.id) t.id = "t_" + Date.now() + "_" + Math.random().toString(36).slice(2, 7);
      });
      templates = templates.concat(importados);
      guardarTemplates();
      renderLista();
      mostrarEstado(`${importados.length} template(s) importado(s).`);
    } catch (e) {
      mostrarEstado("Ficheiro de importação inválido.", true);
    }
  };
  reader.readAsText(ficheiro);
}

/* ---------- Utilitários ---------- */

function escapeHtml(str) {
  return String(str)
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
  estadoTimeout = setTimeout(() => el.classList.add("oculto"), 3000);
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
  document.getElementById("edit-anexo").addEventListener("change", e => lidarComAnexoSelecionado(e.target.files[0]));
  document.getElementById("usar-anexo-extra").addEventListener("change", e => lidarComAnexosExtra(e.target.files));

  document.getElementById("pesquisa").addEventListener("input", e => renderLista(e.target.value));

  document.getElementById("lista-templates").addEventListener("click", e => {
    const btn = e.target.closest("button[data-acao]");
    if (!btn) return;
    const id = btn.dataset.id;
    if (btn.dataset.acao === "usar") abrirUsarTemplate(id);
    if (btn.dataset.acao === "editar") abrirEditorExistente(id);
  });

  document.getElementById("btn-exportar").addEventListener("click", exportarTemplates);
  document.getElementById("btn-importar").addEventListener("click", () => document.getElementById("input-importar").click());
  document.getElementById("input-importar").addEventListener("change", e => {
    if (e.target.files[0]) importarTemplates(e.target.files[0]);
  });
}
