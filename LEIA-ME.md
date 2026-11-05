# Meus Templates — add-in do Outlook

Add-in pessoal para gerir templates de email com variáveis `{{campo}}`, suporte a variáveis dinâmicas pré-definidas, categorias/pastas e anexos, guardados localmente no browser (sem servidor de dados — só os ficheiros precisam de estar acessíveis por HTTPS).

## Ficheiros

- `manifest.xml` — manifesto do add-in (a instalar no Outlook)
- `taskpane.html` / `taskpane.css` / `taskpane.js` — a interface e lógica
- `assets/` — ícones do add-in

## Funcionalidades

1. **Categorias / Pastas:** organiza os templates em categorias (todas podem ser eliminadas, exceto *Geral*).
2. **Editor com formatação:** negrito, itálico, sublinhado, listas e links. Ao colar do Word mantém-se a formatação básica e remove-se o resto.
3. **Variáveis pré-definidas dinâmicas:**
   - `{{Cumprimento}}` / `{{saudacao_tempo}}` — *"Bom dia"*, *"Boa tarde"* ou *"Boa noite"* conforme a hora.
   - `{{NumDoc}}` — Número do Documento (faturas, guias, orçamentos…).
   - `{{data_atual}}` / `{{data}}` — data de hoje (DD/MM/AAAA); `{{data_extenso}}` — *5 de novembro de 2026*.
   - **Datas calculadas:** `{{data+30}}`, `{{data-7}}`, `{{data_extenso+15}}` (dias a somar/subtrair).
   - `{{hora_atual}}` / `{{hora}}` — hora atual.
   - `{{remetente_nome}}` / `{{remetente_email}}` — os teus dados no Outlook.
   - `{{destinatario_nome}}` / `{{destinatario_email}}` — primeiro destinatário (ou quem te enviou o email, ao responder).
4. **Variáveis com opções:**
   - `{{prazo|15 dias}}` — valor por omissão, editável.
   - `{{estado:Pendente,Pago,Em atraso}}` — aparece como lista de escolha.
   - Os nomes aceitam acentos: `{{número}}`, `{{orçamento}}`.
5. **Valores recentes:** cada variável lembra os últimos valores usados e sugere-os ao preencher.
6. **Pré-visualização** do assunto e do corpo antes de inserir.
7. **Onde colocar o texto:** no cursor, no início do email ou substituir tudo (atenção: substituir apaga a assinatura). A escolha fica memorizada.
8. **Responder com template:** ao ler um email, o template abre uma resposta (ou resposta a todos) já preenchida.
9. **Favoritos e ordenação:** marca templates com ★ (ficam sempre no topo) e ordena por *mais usados*, *usados recentemente* ou *nome*.
10. **Duplicar** templates para criar variantes.
11. **Confirmação** antes de eliminar templates ou categorias.
12. **Backup:** exporta todos os templates ou, com uma categoria selecionada no filtro, só essa categoria (útil para partilhar com colegas). Importação com **Unir** ou **Substituir**.

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

- Requer Mailbox API 1.8 (Outlook 2019+ / Microsoft 365 / Outlook na Web) por causa da inserção de anexos em base64 e leitura de perfil/destinatários.
- Os templates e anexos ficam guardados no **IndexedDB** do browser onde o Outlook corre (sem o limite de ~5 MB do `localStorage`) — não saem do teu computador. Na primeira abertura desta versão, os templates antigos são migrados automaticamente.
- Os dados são locais a cada computador/browser: faz **Exportar Backup** regularmente e usa **Importar** para os levar para outro PC.
- Ao responder a partir de um email recebido, o Outlook não permite juntar anexos automaticamente: abre os Meus Templates na janela da resposta para usar o template com anexos.
