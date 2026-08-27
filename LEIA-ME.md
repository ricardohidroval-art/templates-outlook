# Meus Templates — add-in do Outlook

Add-in pessoal para gerir templates de email com variáveis `{{campo}}` e um anexo por template, guardados localmente no browser (sem servidor de dados — só o ficheiro precisa de estar acessível por HTTPS).

## Ficheiros

- `manifest.xml` — manifesto do add-in (a instalar no Outlook)
- `taskpane.html` / `taskpane.css` / `taskpane.js` — a interface e lógica
- `assets/` — ícones do add-in

## 1. Alojar os ficheiros (obrigatório)

O Outlook **não aceita** ficheiros locais — o `manifest.xml` tem de apontar para um endereço **HTTPS** público (mesmo que só tu o uses). Duas opções simples:

**Opção A — GitHub Pages (grátis, recomendado)**
1. Cria um repositório novo (pode ser privado) e faz upload desta pasta completa.
2. Em *Settings → Pages*, ativa o GitHub Pages a partir do branch principal.
3. Vais obter um endereço do tipo `https://o-teu-user.github.io/nome-repo/`.

**Opção B — Alojamento próprio (ex.: subpasta no site da Hidroval)**
- Copia a pasta para um diretório acessível por HTTPS no teu servidor (ex.: `https://hidroval.pt/apps/meus-templates/`).

Depois de escolheres o endereço, **substitui todas as ocorrências** de `SUBSTITUIR-PELO-TEU-DOMINIO` no `manifest.xml` pelo domínio real (sem barra final), por exemplo:

```
https://o-teu-user.github.io/nome-repo
```

## 2. Instalar no Outlook

**Outlook na Web / novo Outlook:**
1. Definições (engrenagem) → *Ver todas as definições do Outlook* → *Correio → Personalizar ações* ou *Extensões* → *Adicionar suplemento personalizado* → *Adicionar a partir de um ficheiro*.
2. Seleciona o `manifest.xml` (já com o domínio atualizado).

**Outlook clássico (Windows/Mac):**
1. Base de Início → *Obter Suplementos* → *Meus suplementos* → *Adicionar um suplemento personalizado* → *Adicionar a partir de um ficheiro*.
2. Seleciona o `manifest.xml`.

O botão **"Meus Templates"** aparece no friso ao compor (ou ler) um email.

## 3. Usar

1. Abre uma nova mensagem e clica em **Meus Templates**.
2. **+ Novo** para criar um template: nome, assunto, corpo (usa `{{variavel}}` onde quiseres um campo a preencher depois) e, opcionalmente, um anexo.
3. Na lista, clica **Usar**: preenche os valores das variáveis detetadas e clica **Inserir no email** — o assunto é definido, o corpo é inserido no ponto onde estava o cursor, e o anexo (se existir) é adicionado automaticamente.
4. **Exportar/Importar** permitem fazer backup dos templates ou levá-los para outro computador (os templates ficam guardados no browser local, por isso não sincronizam automaticamente entre dispositivos).

## Notas técnicas

- Requer Mailbox API 1.8 (Outlook 2016+ / Microsoft 365 / Outlook na Web) por causa da inserção de anexos em base64.
- Os templates e anexos ficam em `localStorage` do browser onde o Outlook corre — não saem do teu computador.
