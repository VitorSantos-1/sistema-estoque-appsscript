# Sistema de Estoque Dimensional (Google Apps Script)

Sistema de controle de estoque construído em Google Apps Script sobre um modelo dimensional
(`Entrada_dim` e `Saida_dim` alimentando `Central_fato`), desenhado para reduzir o erro humano na
operação. Combina a familiaridade da planilha com validações, automação e um modelo de dados
estruturado, o que é incomum em soluções feitas apenas com Sheets.

> **Nota de confidencialidade:** os dados presentes neste repositório são fictícios, gerados apenas
> para demonstração. Os dados reais da operação são confidenciais e estão protegidos — nenhum dado
> real, credencial ou informação de terceiros foi incluído aqui.

---

## Visão Geral

O sistema registra entradas e saídas de estoque em abas dimensionais que alimentam uma tabela fato
central, com validações que evitam o erro de digitação típico da operação. Menus próprios, dropdowns
validados, identificadores únicos por linha e coloração por status transformam uma planilha em uma
ferramenta de controle de estoque disciplinada.

## Contexto de Negócio

Estoque é capital: erro de lançamento vira divergência, ruptura ou perda. Em muitas operações o
controle é feito em planilhas livres, onde qualquer digitação errada compromete o saldo. Aplicar um
modelo dimensional e validações fortes sobre o Sheets reduz esse erro na origem e dá rastreabilidade
às movimentações, sem custo de um sistema dedicado.

## O Problema que Resolve

- **Erro humano de lançamento** em planilhas de estoque livres.
- **Falta de estrutura** (entradas e saídas sem um modelo de dados claro).
- **Ausência de validação** de campos críticos (EAN, fornecedor, endereço).

## Público e Decisões Apoiadas

- **Estoque e Logística:** registram movimentações com validação e rastreio.
- **Gestão:** acompanha entradas e saídas a partir de uma base estruturada.

## Impacto e Valor Gerado

- Reduz o erro humano com dropdowns validados e bloqueio de linhas processadas.
- Estrutura o estoque em um modelo dimensional (entrada, saída, fato central).
- Dá rastreabilidade por identificador único e status visual por linha.

---

## Arquitetura e Abordagem Técnica

- **Modelo dimensional:** `Entrada_dim` e `Saida_dim` alimentam `Central_fato`.
- **Menu próprio** na planilha e **dropdowns validados** (EAN, fornecedor, endereço Rua/Coluna/Volume/Nível).
- **Identificador único (UUID)** por linha e auto-preenchimento de data.
- **Bloqueio** de linhas já processadas e **coloração por status** (pendente, ok, erro).
- **Processamento em lote** com relatório (`processAllEntradas` / `processAllSaidas`).

## Stack

Google Apps Script - JavaScript - Google Sheets - Modelagem Dimensional - Validação de Dados.

## Como Usar

Cole o conteúdo de `Code.gs` no editor de Apps Script de uma planilha Google e recarregue.

## Estrutura do Projeto

```text
Code.gs   -> Lógica do controle de estoque (modelo dimensional, validações, processamento em lote)
```

## Autor

José Vitor Santos Pinheiro — Análise de Dados e Inteligência Comercial (Varejo e Supply Chain).
Contato: vytorsantt@gmail.com
