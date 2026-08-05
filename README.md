# 📦 Sistema de Estoque Dimensional (Google Apps Script)

Sistema de controle de **estoque** em Google Apps Script com **modelo dimensional**
(`Entrada_dim → Central_fato ← Saida_dim`), pensado para evitar erro humano na operação.

> ⚠️ **Aviso sobre os dados**
> Os dados neste repositório são **fictícios**, gerados apenas para demonstração. Os dados reais da
> operação em que o projeto foi usado são **confidenciais e estão protegidos** — nada real, credencial
> ou informação de terceiros foi incluído aqui.

## 🎯 Recursos
- Menu próprio na planilha (📦 Estoque).
- **Dropdowns validados** (EAN, fornecedor, endereço Rua/Coluna/Volume/Nível).
- **UUID automático** por linha; auto-preenchimento de data.
- **Bloqueio** de linhas já processadas; **coloração por status** (🟡 pendente / 🟢 ok / 🔴 erro).
- Processamento **em lote** com relatório (`processAllEntradas` / `processAllSaidas`).

## ▶️ Como usar
Cole o conteúdo de `Code.gs` no editor de Apps Script de uma planilha Google e recarregue.

---

### 🧰 Competências demonstradas
`Google Apps Script` · `JavaScript` · `Google Sheets` · `Modelagem Dimensional` · `Validação de Dados`

### 👤 Autor
**José Vitor Santos Pinheiro** — Analista de Dados / BI / Ciência de Dados · vytorsantt@gmail.com
