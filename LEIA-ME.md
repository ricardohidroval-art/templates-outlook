# Meus Templates — add-in do Outlook

Add-in pessoal para gerir templates de email com variáveis `{{campo}}`, suporte a variáveis dinâmicas pré-definidas, categorias/pastas e anexos, guardados localmente no browser (sem servidor de dados — só os ficheiros precisam de estar acessíveis por HTTPS).

## Ficheiros

- `manifest.xml` — manifesto do add-in (a instalar no Outlook)
- `taskpane.html` / `taskpane.css` / `taskpane.js` — a interface e lógica
- `assets/` — ícones do add-in

## Novidades e Funcionalidades

1. **Categorias / Pastas:** Organiza os teus templates em categorias (ex.: *Comercial*, *Suporte*, *Propostas* ou categorias personalizadas).
2. **Variáveis Pré-definidas Dinâmicas:** 
   - `{{saudacao_tempo}}` — Preenche automaticamente *"Bom dia"*, *"Boa tarde"* ou *"Boa noite"* conforme a hora local.
   - `{{data_atual}}` ou `{{data}}` — Preenche a data de hoje.
   - `{{hora_atual}}` ou `{{hora}}` — Preenche a hora atual.
   - `{{remetente_nome}}` / `{{remetente_email}}` — Obtém os teus dados de perfil no Outlook.
   - `{{destinatario_nome}}` / `{{destinatario_email}}` — Obtém automaticamente os dados do primeiro destinatário no email ativo.
3. **Pesquisa e Filtros Rápido:** Procura em tempo real por nome, assunto, corpo do email ou filtra por categoria.
4. **Melhorias em Backup e Restauro:** Exportação em formato JSON com metadados e assistente de importação com opções de **Unir com existentes** ou **Substituir todos**.

---

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
2. **+ Novo** para criar um template: escolhe/escreve a categoria, nome, assunto, corpo (usa `{{variavel}}` ou os atalhos de variáveis dinâmicas) e, opcionalmente, um anexo fixo.
3. Na lista, clica **Usar**: os valores das variáveis pré-definidas são calculados e preenchidos automaticamente. Podes ajustá-los e clicar em **Inserir no email**.
4. **Exportar/Importar** no rodapé para salvaguarda de backups.

## Notas técnicas

- Requer Mailbox API 1.8 (Outlook 2016+ / Microsoft 365 / Outlook na Web) por causa da inserção de anexos em base64 e leitura de perfil/destinatários.
- Os templates e anexos ficam em `localStorage` do browser onde o Outlook corre — não saem do teu computador.
