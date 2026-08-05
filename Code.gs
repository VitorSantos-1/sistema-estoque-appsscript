/***************
 * Sistema de Estoque v2.0 — Anti-erro humano
 * Entrada_dim → Central_fato ← Saida_dim
 *
 * MELHORIAS v2.0:
 *  1. Menu "📦 Estoque" no topo da planilha
 *  2. Dropdown de EAN na Entrada (de Produtos_dim)
 *  3. Dropdown de Fornecedor na Entrada (de Fornecedor_dim)
 *  4. Dropdowns de Endereço com valores válidos (Rua/Col/Vol/Nív)
 *  5. Auto-preenche Data de entrada com hoje
 *  6. Auto-UUID ao começar qualquer linha
 *  7. Bloqueio de linhas já processadas na Saída
 *  8. Coloração por status: 🟡 pendente / 🟢 ok / 🔴 erro
 *  9. processAllEntradas / processAllSaidas — lote com relatório
 * 10. Cabeçalho congelado e protegido em todas as abas
 * 11. ensureColumn_ fora do hot path de movimentação
 * 12. Dropdown de EAN na Saída (só produtos com saldo no Centro)
 */

const CFG = {
  SHEETS: {
    PRODUTOS: 'Produtos_dim',
    FORNECEDOR: 'Fornecedor_dim',
    ENTRADA: 'Entrada_dim',
    CENTRO: 'Central_fato',
    SAIDA: 'Saida_dim',
    STATUS: 'Status',
  },

  // Aliases — nomes devem bater com o xlsx importado no Sheets
  HEADERS: {
    COD_BARRAS: ['Código de barras'],
    DESCRICAO: ['Descrição'],
    COD_INTERNO: ['Código interno'],
    MERCADOLOGICO: ['Mercadológico'],

    QTD: ['Quantidade'],
    LOTE: ['Lote'],
    VALIDADE: ['Validade'],

    // 'Código fornecedor' (minúsculo) nas abas; 'Código Fornecedor' como fallback
    COD_FORNECEDOR: ['Código fornecedor', 'Código Fornecedor'],
    // Entrada/Saída usam 'Fornecedor'; Centro usa 'Descrição do fornecedor'
    DESC_FORNECEDOR: ['Fornecedor', 'Descrição do fornecedor'],
    DATA_ENTRADA: ['Data de entrada'],
    DATA_ATUALIZACAO: ['Data de atualização'],

    // Entrada: endereço estoque
    E_RUA: ['Localização estoque rua'],
    E_COL: ['Localização estoque coluna'],
    E_VOL: ['Localização estoque volume'],
    E_NIV: ['Localização estoque nivel'],

    // Centro: endereço
    C_RUA: ['Localização rua'],
    C_COL: ['Localização coluna'],
    C_VOL: ['Localização volume'],
    C_NIV: ['Localização nivel'],

    // Saída: origem/destino
    S_O_RUA: ['Localização origem rua'],
    S_O_COL: ['Localização origem coluna'],
    S_O_VOL: ['Localização origem volume'],
    S_O_NIV: ['Localização origem nivel'],
    S_D_RUA: ['Localização destino rua'],
    S_D_COL: ['Localização destino coluna'],
    S_D_VOL: ['Localização destino volume'],
    S_D_NIV: ['Localização destino nivel'],

    // Colunas de controle (nomes exatos do xlsx)
    PROCESSAR: ['Processar'],
    UUID: ['UUID_sis', 'UUid'],
    PROC_QTD: ['Qtd_processada', 'proc_qtd'],
    PROC_AT: ['Data_processamento', 'proc_at'],
    PROCESSED: ['Processado_sis', 'Processado'],
    PROCESSED_AT: ['Data_processado_sis', 'processed_at'],
    ORIGEM_END: ['Origem endereço (Dropdown)', 'Origem endereço'],
  },
};

/** Paleta de cores para status de linha */
const CORES = {
  PENDENTE: '#FFF9C4',  // 🟡 amarelo — tem dados, aguarda processar
  OK: '#C8E6C9',  // 🟢 verde  — processado com sucesso
  ERRO: '#FFCDD2',  // 🔴 vermelho — erro no processamento
  BRANCO: '#FFFFFF',  // branco — linha vazia
};

/** Máximo de linhas de dados para dropdowns e checkboxes */
const MAX_ROWS = 1000;

/* ============================================================
   UTIL
   ============================================================ */

function normalize_(s) {
  if (s === null || s === undefined) return '';
  return String(s).trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');
}

function getSheetOrThrow_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error(`Aba "${name}" não encontrada.`);
  return sh;
}

function getHeaderMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) throw new Error(`Aba ${sheet.getName()} sem colunas.`);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = new Map();
  headers.forEach((h, i) => map.set(normalize_(h), i + 1));
  return { headers, map, lastCol };
}

function findCol_(sheet, aliases) {
  const { map } = getHeaderMap_(sheet);
  for (const a of aliases) {
    const key = normalize_(a);
    if (map.has(key)) return map.get(key);
  }
  // fallback startsWith
  for (const [k, col] of map.entries()) {
    for (const a of aliases) {
      if (k.startsWith(normalize_(a))) return col;
    }
  }
  return null;
}

function ensureColumn_(sheet, headerName) {
  const col = findCol_(sheet, [headerName]);
  if (col) return col;
  const lastCol = sheet.getLastColumn();
  sheet.getRange(1, lastCol + 1).setValue(headerName);
  return lastCol + 1;
}

function ensureColumns_(sheet, headerNames) {
  const out = {};
  headerNames.forEach(h => { out[h] = ensureColumn_(sheet, h); });
  return out;
}

function toNumber_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'number') return v;
  let s = String(v).trim();
  // Detecta formato BR (vírgula como decimal)
  if (/\d,\d/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

function isEmpty_(v) {
  return v === '' || v === null || v === undefined;
}

function formatQty_(q) {
  return ((typeof q === 'number') ? q : toNumber_(q)).toFixed(2).replace('.', ',');
}

function buildAddrKey_(rua, col, vol, niv) {
  return `R${rua}-C${col}-V${vol}-N${niv}`;
}

function parseOrigemEndereco_(text) {
  const m = String(text || '').trim().match(/R(\d+)-C(\d+)-V(\d+)-N(\d+)/i);
  if (!m) return null;
  return { rua: +m[1], col: +m[2], vol: +m[3], niv: +m[4] };
}

function validateEndereco_(rua, col, vol, niv) {
  rua = +rua; col = +col; vol = +vol; niv = +niv;
  if (!(rua >= 1 && rua <= 8)) return `Rua inválida (${rua}).`;
  const maxCol = (rua === 8) ? 7 : 4;
  if (!(col >= 1 && col <= maxCol)) return `Coluna inválida (${col}) para rua ${rua} (máx ${maxCol}).`;
  if (!(vol >= 1 && vol <= 2)) return `Volume inválido (${vol}).`;
  if (!(niv >= 1 && niv <= 2)) return `Nível inválido (${niv}).`;
  return null;
}

/** Retorna lista de colunas válidas (como strings) para a rua informada */
function colListForRua_(rua) {
  const max = (+rua === 8) ? 7 : 4;
  return Array.from({ length: max }, (_, i) => String(i + 1));
}

function setDataAtualizacao_(sheet, row) {
  const c = findCol_(sheet, CFG.HEADERS.DATA_ATUALIZACAO);
  if (c) sheet.getRange(row, c).setValue(new Date());
}

function toast_(msg) {
  SpreadsheetApp.getActive().toast(msg, '📦 Estoque', 7);
}

function logStatus_(type, message, dataObj) {
  try {
    getSheetOrThrow_(CFG.SHEETS.STATUS)
      .appendRow([new Date(), type, message, dataObj ? JSON.stringify(dataObj) : '']);
  } catch (_) { /* Status sheet pode não existir ainda */ }
}

/** Pinta toda a linha com a cor correspondente ao status */
function colorirLinhaStatus_(sheet, row, status) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  sheet.getRange(row, 1, 1, lastCol).setBackground(CORES[status] || CORES.BRANCO);
}

/* ============================================================
   ÍNDICES (produtos / fornecedor)
   ============================================================ */

function buildProdutosIndex_() {
  const sh = getSheetOrThrow_(CFG.SHEETS.PRODUTOS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return new Map();

  const cCod = findCol_(sh, CFG.HEADERS.COD_BARRAS);
  const cDesc = findCol_(sh, CFG.HEADERS.DESCRICAO);
  const cInt = findCol_(sh, CFG.HEADERS.COD_INTERNO);
  const cMerc = findCol_(sh, CFG.HEADERS.MERCADOLOGICO);
  if (!cCod) return new Map();

  const idx = new Map();
  sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues().forEach(r => {
    // EAN pode vir como float (7.89e+13) — normaliza para inteiro
    const cod = String(r[cCod - 1] ?? '').trim().split('.')[0];
    if (!cod) return;
    idx.set(cod, {
      descricao: cDesc ? r[cDesc - 1] : '',
      codInterno: cInt ? r[cInt - 1] : '',
      mercadologico: cMerc ? r[cMerc - 1] : '',
    });
  });
  return idx;
}

function buildFornecedorIndex_() {
  const sh = getSheetOrThrow_(CFG.SHEETS.FORNECEDOR);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return new Map();

  const cCod = findCol_(sh, CFG.HEADERS.COD_FORNECEDOR);
  const cNome = findCol_(sh, ['Fornecedor', 'Descrição fornecedor', 'Descrição do fornecedor']);
  if (!cCod) return new Map();

  const idx = new Map();
  sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues().forEach(r => {
    const cod = String(r[cCod - 1] ?? '').trim().split('.')[0];
    if (!cod) return;
    idx.set(cod, cNome ? r[cNome - 1] : '');
  });
  return idx;
}

/* ============================================================
   MENU PERSONALIZADO (abre automaticamente com a planilha)
   ============================================================ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Estoque')
    .addItem('⚙️  Setup completo (executar 1× após importar)', 'setupEstoque')
    .addSeparator()
    .addItem('▶️  Processar TODAS as Entradas marcadas', 'processAllEntradas')
    .addItem('▶️  Processar TODAS as Saídas marcadas', 'processAllSaidas')
    .addSeparator()
    .addItem('🔄  Atualizar dropdown de produtos (Saída)', 'setupDropdownsCodSaida_')
    .addSeparator()
    .addItem('🔧  Recriar triggers', 'createTriggers')
    .addToUi();
}

/* ============================================================
   SETUP
   ============================================================ */

function setupEstoque() {
  const lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    const entrada = getSheetOrThrow_(CFG.SHEETS.ENTRADA);
    const centro = getSheetOrThrow_(CFG.SHEETS.CENTRO);
    const saida = getSheetOrThrow_(CFG.SHEETS.SAIDA);
    const status = getSheetOrThrow_(CFG.SHEETS.STATUS);

    if (status.getLastRow() === 0) status.appendRow(['timestamp', 'tipo', 'mensagem', 'json']);

    // 1. Garante colunas de controle (cria se não existirem)
    ensureColumn_(centro, 'Descrição do fornecedor');
    ensureColumns_(entrada, ['Processar', 'UUID_sis', 'Qtd_processada', 'Data_processamento']);
    ensureColumns_(saida, [
      'Código fornecedor',
      'Fornecedor',
      'Origem endereço (Dropdown)',
      'Processar',
      'UUID_sis',
      'Processado_sis',
      'Data_processado_sis',
    ]);

    // 2. Checkboxes na coluna "Processar"
    const cProcE = findCol_(entrada, CFG.HEADERS.PROCESSAR);
    if (cProcE) entrada.getRange(2, cProcE, MAX_ROWS).insertCheckboxes();
    const cProcS = findCol_(saida, CFG.HEADERS.PROCESSAR);
    if (cProcS) saida.getRange(2, cProcS, MAX_ROWS).insertCheckboxes();

    // 3. Dropdown de EAN na Entrada (vem de Produtos_dim)
    setupDropdownsProdutos_();

    // 4. Dropdown de Fornecedor na Entrada (vem de Fornecedor_dim)
    setupDropdownsFornecedor_();

    // 5. Dropdowns de Endereço (Rua / Col / Vol / Nív)
    setupDropdownsEndereco_();

    // 6. Dropdown de EAN na Saída (apenas produtos com saldo no Centro)
    setupDropdownsCodSaida_();

    // 7. Congela cabeçalho nas abas principais
    [entrada, centro, saida].forEach(sh => {
      if (sh.getFrozenRows() < 1) sh.setFrozenRows(1);
    });

    // 8. Protege linha 1 (warning-only — impede edição acidental)
    protegerCabecalhos_();

    logStatus_('INFO', 'Setup v2.0 concluído.', {});
    toast_('✅ Setup concluído! Planilha pronta para uso.');
  } finally {
    lock.releaseLock();
  }
}

/** Dropdown de Código de barras na Entrada (fonte: Produtos_dim col EAN) */
function setupDropdownsProdutos_() {
  const produtos = getSheetOrThrow_(CFG.SHEETS.PRODUTOS);
  const entrada = getSheetOrThrow_(CFG.SHEETS.ENTRADA);
  const lastRowP = produtos.getLastRow();
  if (lastRowP < 2) return;
  const cCodP = findCol_(produtos, CFG.HEADERS.COD_BARRAS);
  const cCodE = findCol_(entrada, CFG.HEADERS.COD_BARRAS);
  if (!cCodP || !cCodE) return;
  // requireValueInRange suporta listas grandes sem limite de caracteres
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(produtos.getRange(2, cCodP, lastRowP - 1, 1), true)
    .setAllowInvalid(false).build();
  entrada.getRange(2, cCodE, MAX_ROWS, 1).setDataValidation(rule);
}

/** Dropdown de EAN na Saída — apenas produtos com saldo > 0 no Centro */
function setupDropdownsCodSaida_() {
  const centro = getSheetOrThrow_(CFG.SHEETS.CENTRO);
  const saida = getSheetOrThrow_(CFG.SHEETS.SAIDA);
  const cCodS = findCol_(saida, CFG.HEADERS.COD_BARRAS);
  if (!cCodS) return;

  const lastRowC = centro.getLastRow();
  if (lastRowC < 2) {
    saida.getRange(2, cCodS, MAX_ROWS, 1).clearDataValidations();
    return;
  }
  const cCodC = findCol_(centro, CFG.HEADERS.COD_BARRAS);
  const cQtdC = findCol_(centro, CFG.HEADERS.QTD);
  if (!cCodC || !cQtdC) return;

  const setCods = new Set();
  centro.getRange(2, 1, lastRowC - 1, centro.getLastColumn()).getValues().forEach(r => {
    if (toNumber_(r[cQtdC - 1]) <= 0) return;
    const cod = String(r[cCodC - 1] ?? '').trim().split('.')[0];
    if (cod) setCods.add(cod);
  });

  const list = Array.from(setCods).sort();
  if (!list.length) {
    saida.getRange(2, cCodS, MAX_ROWS, 1).clearDataValidations();
    return;
  }
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(false).build();
  saida.getRange(2, cCodS, MAX_ROWS, 1).setDataValidation(rule);
}

/** Dropdown de Código Fornecedor na Entrada (fonte: Fornecedor_dim) */
function setupDropdownsFornecedor_() {
  const fornecedor = getSheetOrThrow_(CFG.SHEETS.FORNECEDOR);
  const entrada = getSheetOrThrow_(CFG.SHEETS.ENTRADA);
  const lastRowF = fornecedor.getLastRow();
  if (lastRowF < 2) return;
  const cCodF_dim = findCol_(fornecedor, CFG.HEADERS.COD_FORNECEDOR);
  const cCodF_ent = findCol_(entrada, CFG.HEADERS.COD_FORNECEDOR);
  if (!cCodF_dim || !cCodF_ent) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(fornecedor.getRange(2, cCodF_dim, lastRowF - 1, 1), true)
    .setAllowInvalid(false).build();
  entrada.getRange(2, cCodF_ent, MAX_ROWS, 1).setDataValidation(rule);
}

/** Dropdowns de endereço (Rua 1-8 / Col 1-7 / Vol 1-2 / Nív 1-2) */
function setupDropdownsEndereco_() {
  const mkRule = list => SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true).setAllowInvalid(false).build();

  const rRule = mkRule(['1', '2', '3', '4', '5', '6', '7', '8']);
  const cRule = mkRule(['1', '2', '3', '4', '5', '6', '7']); // máx rua 8; validação fina no process
  const vRule = mkRule(['1', '2']);
  const nRule = mkRule(['1', '2']);

  // Entrada
  const entrada = getSheetOrThrow_(CFG.SHEETS.ENTRADA);
  [
    [CFG.HEADERS.E_RUA, rRule],
    [CFG.HEADERS.E_COL, cRule],
    [CFG.HEADERS.E_VOL, vRule],
    [CFG.HEADERS.E_NIV, nRule],
  ].forEach(([hdr, rule]) => {
    const c = findCol_(entrada, hdr);
    if (c) entrada.getRange(2, c, MAX_ROWS, 1).setDataValidation(rule);
  });

  // Saída: apenas destino (origem vem do dropdown do Centro)
  const saida = getSheetOrThrow_(CFG.SHEETS.SAIDA);
  [
    [CFG.HEADERS.S_D_RUA, rRule],
    [CFG.HEADERS.S_D_COL, cRule],
    [CFG.HEADERS.S_D_VOL, vRule],
    [CFG.HEADERS.S_D_NIV, nRule],
  ].forEach(([hdr, rule]) => {
    const c = findCol_(saida, hdr);
    if (c) saida.getRange(2, c, MAX_ROWS, 1).setDataValidation(rule);
  });
}

/** Protege linha 1 com aviso em todas as abas principais */
function protegerCabecalhos_() {
  const ss = SpreadsheetApp.getActive();
  [CFG.SHEETS.ENTRADA, CFG.SHEETS.SAIDA, CFG.SHEETS.CENTRO,
  CFG.SHEETS.PRODUTOS, CFG.SHEETS.FORNECEDOR].forEach(nome => {
    try {
      const sh = ss.getSheetByName(nome);
      if (!sh || sh.getLastColumn() < 1) return;
      // Remove proteções antigas do cabeçalho para não acumular
      sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)
        .filter(p => p.getRange().getRow() === 1 && p.getRange().getNumRows() === 1)
        .forEach(p => p.remove());
      const prot = sh.getRange(1, 1, 1, sh.getLastColumn()).protect();
      prot.setDescription(`Cabeçalho ${nome} — não editar`);
      prot.setWarningOnly(true);
    } catch (_) { /* sem permissão ainda — ok */ }
  });
}

function createTriggers() {
  const existing = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  const ss = SpreadsheetApp.getActive();
  if (!existing.includes('onEditEstoque')) {
    ScriptApp.newTrigger('onEditEstoque').forSpreadsheet(ss).onEdit().create();
  }
  if (!existing.includes('onOpen')) {
    ScriptApp.newTrigger('onOpen').forSpreadsheet(ss).onOpen().create();
  }
  logStatus_('INFO', 'Triggers criados/confirmados.', {});
  toast_('🔧 Triggers criados/confirmados.');
}

/* ============================================================
   onEdit (trigger instalado)
   ============================================================ */

function onEditEstoque(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();
  const name = sheet.getName();
  const row = range.getRow();
  const col = range.getColumn();

  if (row === 1) return; // nunca mexa no cabeçalho

  try {
    if (name === CFG.SHEETS.ENTRADA) {
      setDataAtualizacao_(sheet, row);
      handleEditEntrada_(e);
    } else if (name === CFG.SHEETS.SAIDA) {
      // Bloqueia linhas já processadas ANTES de qualquer ação
      if (isLinhaProcessadaSaida_(sheet, row, col)) {
        if (e.oldValue !== undefined) range.setValue(e.oldValue);
        toast_('⚠️  Esta linha já foi processada e não pode ser alterada.');
        return;
      }
      setDataAtualizacao_(sheet, row);
      handleEditSaida_(e);
    }
  } catch (err) {
    logStatus_('ERRO', err.message, { sheet: name, row, col });
    toast_(`❌ Erro: ${err.message}`);
    colorirLinhaStatus_(sheet, row, 'ERRO');
  }
}

/**
 * Verifica se a linha da Saída já foi processada E a coluna editada
 * não é uma coluna de sistema (portanto deve ser bloqueada).
 */
function isLinhaProcessadaSaida_(sheet, row, editedCol) {
  const cProc = findCol_(sheet, CFG.HEADERS.PROCESSED);
  if (!cProc) return false;
  if (sheet.getRange(row, cProc).getValue() !== true) return false;

  const colsSistema = [
    findCol_(sheet, CFG.HEADERS.PROCESSED),
    findCol_(sheet, CFG.HEADERS.PROCESSED_AT),
    findCol_(sheet, CFG.HEADERS.UUID),
    findCol_(sheet, CFG.HEADERS.DATA_ATUALIZACAO),
  ].filter(Boolean);

  return !colsSistema.includes(editedCol);
}

function handleEditEntrada_(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  const cProc = findCol_(sheet, CFG.HEADERS.PROCESSAR);
  const cCod = findCol_(sheet, CFG.HEADERS.COD_BARRAS);
  const cCodF = findCol_(sheet, CFG.HEADERS.COD_FORNECEDOR);
  const cUuid = findCol_(sheet, CFG.HEADERS.UUID);
  const cDataEnt = findCol_(sheet, CFG.HEADERS.DATA_ENTRADA);
  const cRua = findCol_(sheet, CFG.HEADERS.E_RUA);
  const cColE = findCol_(sheet, CFG.HEADERS.E_COL);

  // Auto-UUID: gera na primeira edição da linha
  if (cUuid && isEmpty_(sheet.getRange(row, cUuid).getValue())) {
    sheet.getRange(row, cUuid).setValue(Utilities.getUuid());
  }

  // Auto-Data de entrada: preenche com hoje se vazio
  if (cDataEnt && isEmpty_(sheet.getRange(row, cDataEnt).getValue())) {
    sheet.getRange(row, cDataEnt).setValue(new Date());
  }

  // Autopreenchimento ao selecionar código de barras
  if (cCod && col === cCod) autopreencheProduto_(sheet, row);

  // Autopreenchimento ao selecionar Código Fornecedor
  if (cCodF && col === cCodF) autopreencheFornecedor_(sheet, row);

  // Quando Rua muda → recalcula dropdown de Coluna conforme máximo permitido
  if (cRua && cColE && col === cRua) {
    const ruaVal = sheet.getRange(row, cRua).getValue();
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(colListForRua_(ruaVal), true)
      .setAllowInvalid(false).build();
    sheet.getRange(row, cColE).clearContent().setDataValidation(rule);
  }

  // Coloração: linha com código → PENDENTE
  if (cCod && !isEmpty_(sheet.getRange(row, cCod).getValue())) {
    colorirLinhaStatus_(sheet, row, 'PENDENTE');
  }

  // Processar?
  if (cProc && col === cProc && e.value === 'TRUE') {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const res = processEntradaRow_(row);
      sheet.getRange(row, cProc).setValue(false);
      colorirLinhaStatus_(sheet, row, 'OK');
      logStatus_('OK', `Entrada processada (linha ${row}).`, res);
      toast_(`✅ Entrada OK — linha ${row}.`);
    } catch (err) {
      sheet.getRange(row, cProc).setValue(false);
      colorirLinhaStatus_(sheet, row, 'ERRO');
      logStatus_('ERRO', `Entrada L${row}: ${err.message}`, {});
      toast_(`❌ L${row}: ${err.message}`);
    } finally {
      lock.releaseLock();
    }
  }
}

function handleEditSaida_(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  const cProc = findCol_(sheet, CFG.HEADERS.PROCESSAR);
  const cCod = findCol_(sheet, CFG.HEADERS.COD_BARRAS);
  const cVal = findCol_(sheet, CFG.HEADERS.VALIDADE);
  const cLote = findCol_(sheet, CFG.HEADERS.LOTE);
  const cCodF = findCol_(sheet, CFG.HEADERS.COD_FORNECEDOR);
  const cOrigemEnd = findCol_(sheet, CFG.HEADERS.ORIGEM_END);
  const cUuid = findCol_(sheet, CFG.HEADERS.UUID);
  const cDRua = findCol_(sheet, CFG.HEADERS.S_D_RUA);
  const cDCol = findCol_(sheet, CFG.HEADERS.S_D_COL);

  // Auto-UUID: gera na primeira edição da linha
  if (cUuid && isEmpty_(sheet.getRange(row, cUuid).getValue())) {
    sheet.getRange(row, cUuid).setValue(Utilities.getUuid());
  }

  // Autopreenchimento produto/fornecedor
  if (cCod && col === cCod) autopreencheProduto_(sheet, row);
  if (cCodF && col === cCodF) autopreencheFornecedor_(sheet, row);

  // Cascata de limpeza ao mudar campo superior (invalida escolhas dependentes)
  if (cCod && col === cCod) {
    [cVal, cLote, cCodF, cOrigemEnd].filter(Boolean)
      .forEach(c => sheet.getRange(row, c).clearContent().clearDataValidations());
  }
  if (cVal && col === cVal) {
    [cLote, cCodF, cOrigemEnd].filter(Boolean)
      .forEach(c => sheet.getRange(row, c).clearContent().clearDataValidations());
  }
  if (cLote && col === cLote) {
    [cCodF, cOrigemEnd].filter(Boolean)
      .forEach(c => sheet.getRange(row, c).clearContent().clearDataValidations());
  }
  if (cCodF && col === cCodF && cOrigemEnd) {
    sheet.getRange(row, cOrigemEnd).clearContent().clearDataValidations();
  }

  // Quando Rua destino muda → recalcula dropdown de Coluna destino
  if (cDRua && cDCol && col === cDRua) {
    const ruaVal = sheet.getRange(row, cDRua).getValue();
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(colListForRua_(ruaVal), true)
      .setAllowInvalid(false).build();
    sheet.getRange(row, cDCol).clearContent().setDataValidation(rule);
  }

  // Atualiza dropdowns em cascata (validade → lote → fornecedor → origem)
  if ([cCod, cVal, cLote, cCodF].includes(col)) {
    updateSaidaValidations_(row);
  }

  // Coloração: linha com código → PENDENTE
  if (cCod && !isEmpty_(sheet.getRange(row, cCod).getValue())) {
    colorirLinhaStatus_(sheet, row, 'PENDENTE');
  }

  // Processar?
  if (cProc && col === cProc && e.value === 'TRUE') {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const res = processSaidaRow_(row);
      sheet.getRange(row, cProc).setValue(false);
      colorirLinhaStatus_(sheet, row, 'OK');
      logStatus_('OK', `Saída processada (linha ${row}).`, res);
      toast_(`✅ Saída OK — linha ${row}.`);
      // Revalida dropdowns e atualiza lista de EANs com saldo (saldos mudaram)
      updateSaidaValidations_(row);
      setupDropdownsCodSaida_();
    } catch (err) {
      sheet.getRange(row, cProc).setValue(false);
      colorirLinhaStatus_(sheet, row, 'ERRO');
      logStatus_('ERRO', `Saída L${row}: ${err.message}`, {});
      toast_(`❌ L${row}: ${err.message}`);
    } finally {
      lock.releaseLock();
    }
  }
}

/* ============================================================
   AUTOPREENCHIMENTO
   ============================================================ */

function autopreencheProduto_(sheet, row) {
  const produtos = buildProdutosIndex_();
  const cCod = findCol_(sheet, CFG.HEADERS.COD_BARRAS);
  if (!cCod) return;

  const cod = String(sheet.getRange(row, cCod).getValue() ?? '').trim().split('.')[0];
  if (!cod) return;
  const prod = produtos.get(cod);
  if (!prod) return;

  const cDesc = findCol_(sheet, CFG.HEADERS.DESCRICAO);
  const cInt = findCol_(sheet, CFG.HEADERS.COD_INTERNO);
  const cMerc = findCol_(sheet, CFG.HEADERS.MERCADOLOGICO);

  if (cDesc && isEmpty_(sheet.getRange(row, cDesc).getValue())) sheet.getRange(row, cDesc).setValue(prod.descricao);
  if (cInt && isEmpty_(sheet.getRange(row, cInt).getValue())) sheet.getRange(row, cInt).setValue(prod.codInterno);
  if (cMerc && isEmpty_(sheet.getRange(row, cMerc).getValue())) sheet.getRange(row, cMerc).setValue(prod.mercadologico);
}

function autopreencheFornecedor_(sheet, row) {
  const forn = buildFornecedorIndex_();
  const cCodF = findCol_(sheet, CFG.HEADERS.COD_FORNECEDOR);
  if (!cCodF) return;

  const codF = String(sheet.getRange(row, cCodF).getValue() ?? '').trim().split('.')[0];
  if (!codF) return;
  const nome = forn.get(codF);
  if (!nome) return;

  const cDescF = findCol_(sheet, CFG.HEADERS.DESC_FORNECEDOR);
  if (cDescF && isEmpty_(sheet.getRange(row, cDescF).getValue())) sheet.getRange(row, cDescF).setValue(nome);
}

/* ============================================================
   DROPDOWNS SAÍDA — cascata baseada no saldo do Centro
   ============================================================ */

function updateSaidaValidations_(row) {
  const saida = getSheetOrThrow_(CFG.SHEETS.SAIDA);
  const centro = getSheetOrThrow_(CFG.SHEETS.CENTRO);

  const cCodS = findCol_(saida, CFG.HEADERS.COD_BARRAS);
  const cValS = findCol_(saida, CFG.HEADERS.VALIDADE);
  const cLoteS = findCol_(saida, CFG.HEADERS.LOTE);
  const cFornS = findCol_(saida, CFG.HEADERS.COD_FORNECEDOR);
  const cOrigemEnd = findCol_(saida, CFG.HEADERS.ORIGEM_END);
  if (!cCodS || !cValS || !cLoteS || !cFornS || !cOrigemEnd) return;

  const cod = String(saida.getRange(row, cCodS).getValue() ?? '').trim().split('.')[0];
  const validade = String(saida.getRange(row, cValS).getValue() || '').trim();
  const lote = String(saida.getRange(row, cLoteS).getValue() || '').trim();
  const codForn = String(saida.getRange(row, cFornS).getValue() || '').trim();

  const lastRowC = centro.getLastRow();
  if (lastRowC < 2) return;

  // Lê como display (evita treta de Date vs string nas chaves)
  const display = centro.getRange(2, 1, lastRowC - 1, centro.getLastColumn()).getDisplayValues();
  const values = centro.getRange(2, 1, lastRowC - 1, centro.getLastColumn()).getValues();

  const cCodC = findCol_(centro, CFG.HEADERS.COD_BARRAS);
  const cValC = findCol_(centro, CFG.HEADERS.VALIDADE);
  const cLoteC = findCol_(centro, CFG.HEADERS.LOTE);
  const cFornC = findCol_(centro, CFG.HEADERS.COD_FORNECEDOR);
  const cQtdC = findCol_(centro, CFG.HEADERS.QTD);
  const cRuaC = findCol_(centro, CFG.HEADERS.C_RUA);
  const cColC = findCol_(centro, CFG.HEADERS.C_COL);
  const cVolC = findCol_(centro, CFG.HEADERS.C_VOL);
  const cNivC = findCol_(centro, CFG.HEADERS.C_NIV);
  if (!cCodC || !cValC || !cLoteC || !cFornC || !cQtdC || !cRuaC || !cColC || !cVolC || !cNivC) return;

  const mkRule = list => SpreadsheetApp.newDataValidation()
    .requireValueInList(list.length ? list : ['(sem estoque)'], true)
    .setAllowInvalid(false).build();

  // 1) Validades disponíveis para o código
  if (cod) {
    const s = new Set();
    for (let i = 0; i < display.length; i++) {
      if (toNumber_(values[i][cQtdC - 1]) <= 0) continue;
      if (String(display[i][cCodC - 1] || '').trim() !== cod) continue;
      const v = String(display[i][cValC - 1] || '').trim();
      if (v) s.add(v);
    }
    saida.getRange(row, cValS).setDataValidation(mkRule(Array.from(s).sort()));
  }

  // 2) Lotes disponíveis para código+validade
  if (cod && validade) {
    const s = new Set();
    for (let i = 0; i < display.length; i++) {
      if (toNumber_(values[i][cQtdC - 1]) <= 0) continue;
      if (String(display[i][cCodC - 1] || '').trim() !== cod) continue;
      if (String(display[i][cValC - 1] || '').trim() !== validade) continue;
      const l = String(display[i][cLoteC - 1] || '').trim();
      if (l) s.add(l);
    }
    saida.getRange(row, cLoteS).setDataValidation(mkRule(Array.from(s).sort()));
  }

  // 3) Fornecedores disponíveis para código+validade+lote
  if (cod && validade && lote) {
    const s = new Set();
    for (let i = 0; i < display.length; i++) {
      if (toNumber_(values[i][cQtdC - 1]) <= 0) continue;
      if (String(display[i][cCodC - 1] || '').trim() !== cod) continue;
      if (String(display[i][cValC - 1] || '').trim() !== validade) continue;
      if (String(display[i][cLoteC - 1] || '').trim() !== lote) continue;
      const f = String(display[i][cFornC - 1] || '').trim();
      if (f) s.add(f);
    }
    saida.getRange(row, cFornS).setDataValidation(mkRule(Array.from(s).sort()));
  }

  // 4) Origens (endereços com saldo) para código+validade+lote+fornecedor
  if (cod && validade && lote && codForn) {
    const origs = [];
    for (let i = 0; i < display.length; i++) {
      const qtd = toNumber_(values[i][cQtdC - 1]);
      if (qtd <= 0) continue;
      if (String(display[i][cCodC - 1] || '').trim() !== cod) continue;
      if (String(display[i][cValC - 1] || '').trim() !== validade) continue;
      if (String(display[i][cLoteC - 1] || '').trim() !== lote) continue;
      if (String(display[i][cFornC - 1] || '').trim() !== codForn) continue;
      const addr = buildAddrKey_(
        display[i][cRuaC - 1], display[i][cColC - 1],
        display[i][cVolC - 1], display[i][cNivC - 1]);
      origs.push(`${addr} | Qtd ${formatQty_(qtd)}`);
    }
    saida.getRange(row, cOrigemEnd).setDataValidation(mkRule(origs.sort()));
  }
}

/* ============================================================
   PROCESSAR EM LOTE
   ============================================================ */

function processAllEntradas() {
  const entrada = getSheetOrThrow_(CFG.SHEETS.ENTRADA);
  const cProc = findCol_(entrada, CFG.HEADERS.PROCESSAR);
  if (!cProc) { toast_('❌ Coluna "Processar" não encontrada.'); return; }

  const lastRow = entrada.getLastRow();
  if (lastRow < 2) { toast_('Sem dados em Entrada_dim.'); return; }

  const procVals = entrada.getRange(2, cProc, lastRow - 1, 1).getValues();

  const lock = LockService.getScriptLock();
  lock.waitLock(60000);

  let ok = 0, erros = 0;
  const errosMsgs = [];

  try {
    for (let i = 0; i < procVals.length; i++) {
      if (procVals[i][0] !== true) continue;
      const row = i + 2;
      try {
        processEntradaRow_(row);
        entrada.getRange(row, cProc).setValue(false);
        colorirLinhaStatus_(entrada, row, 'OK');
        ok++;
      } catch (err) {
        erros++;
        errosMsgs.push(`L${row}: ${err.message}`);
        colorirLinhaStatus_(entrada, row, 'ERRO');
        logStatus_('ERRO', `Lote Entrada L${row}: ${err.message}`, {});
      }
    }
  } finally {
    lock.releaseLock();
  }

  const msg = `✅ ${ok} entrada(s) processada(s).${erros ? ` ❌ ${erros} erro(s).` : ''}`;
  toast_(msg);
  logStatus_('INFO', `Lote Entrada: ${ok} ok, ${erros} erros.`, { erros: errosMsgs });
  if (errosMsgs.length) {
    SpreadsheetApp.getUi().alert('⚠️ Erros no lote de Entradas:\n\n' + errosMsgs.join('\n'));
  }
}

function processAllSaidas() {
  const saida = getSheetOrThrow_(CFG.SHEETS.SAIDA);
  const cProc = findCol_(saida, CFG.HEADERS.PROCESSAR);
  if (!cProc) { toast_('❌ Coluna "Processar" não encontrada.'); return; }

  const lastRow = saida.getLastRow();
  if (lastRow < 2) { toast_('Sem dados em Saida_dim.'); return; }

  const procVals = saida.getRange(2, cProc, lastRow - 1, 1).getValues();

  const lock = LockService.getScriptLock();
  lock.waitLock(60000);

  let ok = 0, erros = 0;
  const errosMsgs = [];

  try {
    for (let i = 0; i < procVals.length; i++) {
      if (procVals[i][0] !== true) continue;
      const row = i + 2;
      try {
        processSaidaRow_(row);
        saida.getRange(row, cProc).setValue(false);
        colorirLinhaStatus_(saida, row, 'OK');
        ok++;
      } catch (err) {
        erros++;
        errosMsgs.push(`L${row}: ${err.message}`);
        colorirLinhaStatus_(saida, row, 'ERRO');
        logStatus_('ERRO', `Lote Saída L${row}: ${err.message}`, {});
      }
    }
  } finally {
    lock.releaseLock();
  }

  // Atualiza dropdown de EAN da Saída (saldos mudaram)
  try { setupDropdownsCodSaida_(); } catch (_) { }

  const msg = `✅ ${ok} saída(s) processada(s).${erros ? ` ❌ ${erros} erro(s).` : ''}`;
  toast_(msg);
  logStatus_('INFO', `Lote Saída: ${ok} ok, ${erros} erros.`, { erros: errosMsgs });
  if (errosMsgs.length) {
    SpreadsheetApp.getUi().alert('⚠️ Erros no lote de Saídas:\n\n' + errosMsgs.join('\n'));
  }
}

/* ============================================================
   PROCESSAR ENTRADA (delta por linha)
   ============================================================ */

function processEntradaRow_(row) {
  const entrada = getSheetOrThrow_(CFG.SHEETS.ENTRADA);
  const produtosIdx = buildProdutosIndex_();
  const fornIdx = buildFornecedorIndex_();

  const cCod = findCol_(entrada, CFG.HEADERS.COD_BARRAS);
  const cDesc = findCol_(entrada, CFG.HEADERS.DESCRICAO);
  const cInt = findCol_(entrada, CFG.HEADERS.COD_INTERNO);
  const cMerc = findCol_(entrada, CFG.HEADERS.MERCADOLOGICO);
  const cQtd = findCol_(entrada, CFG.HEADERS.QTD);
  const cLote = findCol_(entrada, CFG.HEADERS.LOTE);
  const cVal = findCol_(entrada, CFG.HEADERS.VALIDADE);
  const cCodF = findCol_(entrada, CFG.HEADERS.COD_FORNECEDOR);
  const cDescF = findCol_(entrada, CFG.HEADERS.DESC_FORNECEDOR);
  const cDataEnt = findCol_(entrada, CFG.HEADERS.DATA_ENTRADA);
  const cRua = findCol_(entrada, CFG.HEADERS.E_RUA);
  const cCol = findCol_(entrada, CFG.HEADERS.E_COL);
  const cVol = findCol_(entrada, CFG.HEADERS.E_VOL);
  const cNiv = findCol_(entrada, CFG.HEADERS.E_NIV);
  const cProcQtd = findCol_(entrada, CFG.HEADERS.PROC_QTD);
  const cProcAt = findCol_(entrada, CFG.HEADERS.PROC_AT);
  const cUuid = findCol_(entrada, CFG.HEADERS.UUID);

  if (!cCod || !cQtd || !cLote || !cVal || !cCodF || !cRua || !cCol || !cVol || !cNiv) {
    throw new Error('Entrada: faltam colunas obrigatórias (código, quantidade, lote, validade, fornecedor, endereço).');
  }

  const rowVals = entrada.getRange(row, 1, 1, entrada.getLastColumn()).getValues()[0];

  const cod = String(rowVals[cCod - 1] ?? '').trim().split('.')[0];
  if (!cod) throw new Error(`Entrada L${row}: código de barras vazio.`);

  const prod = produtosIdx.get(cod);
  if (prod) {
    if (cDesc && isEmpty_(rowVals[cDesc - 1])) entrada.getRange(row, cDesc).setValue(prod.descricao);
    if (cInt && isEmpty_(rowVals[cInt - 1])) entrada.getRange(row, cInt).setValue(prod.codInterno);
    if (cMerc && isEmpty_(rowVals[cMerc - 1])) entrada.getRange(row, cMerc).setValue(prod.mercadologico);
  }

  const qtd = toNumber_(rowVals[cQtd - 1]);
  if (qtd <= 0) throw new Error(`Entrada L${row}: quantidade inválida (${rowVals[cQtd - 1]}).`);

  const lote = String(rowVals[cLote - 1] || '').trim();
  const rawVal = rowVals[cVal - 1];
  const val = rawVal instanceof Date
    ? Utilities.formatDate(rawVal, Session.getScriptTimeZone(), 'dd/MM/yyyy')
    : String(rawVal || '').trim();
  if (!lote || !val) throw new Error(`Entrada L${row}: lote e validade são obrigatórios.`);

  const codF = String(rowVals[cCodF - 1] || '').trim();
  if (!codF) throw new Error(`Entrada L${row}: Código Fornecedor obrigatório.`);

  if (cDescF && isEmpty_(rowVals[cDescF - 1]) && fornIdx.has(codF)) {
    entrada.getRange(row, cDescF).setValue(fornIdx.get(codF));
  }

  const rua = +rowVals[cRua - 1];
  const col = +rowVals[cCol - 1];
  const vol = +rowVals[cVol - 1];
  const niv = +rowVals[cNiv - 1];

  const errEnd = validateEndereco_(rua, col, vol, niv);
  if (errEnd) throw new Error(`Entrada L${row}: ${errEnd}`);

  if (cUuid && isEmpty_(rowVals[cUuid - 1])) entrada.getRange(row, cUuid).setValue(Utilities.getUuid());

  const prevProc = cProcQtd ? toNumber_(rowVals[cProcQtd - 1]) : 0;
  const delta = qtd - prevProc;

  if (delta === 0) {
    if (cProcAt) entrada.getRange(row, cProcAt).setValue(new Date());
    return { ok: true, message: 'Sem delta (já consolidado).', delta: 0 };
  }

  // Data de entrada: usa o valor da célula (já preenchido automaticamente)
  const dataEntrada = cDataEnt
    ? (rowVals[cDataEnt - 1] instanceof Date ? rowVals[cDataEnt - 1] : new Date())
    : new Date();

  applyDeltaCentro_({
    codBarras: cod,
    lote,
    validade: val,
    codFornecedor: codF,
    endereco: { rua, col, vol, niv },
    deltaQtd: delta,
    descricao: prod ? prod.descricao : (cDesc ? rowVals[cDesc - 1] : ''),
    codInterno: prod ? prod.codInterno : (cInt ? rowVals[cInt - 1] : ''),
    mercadologico: prod ? prod.mercadologico : (cMerc ? rowVals[cMerc - 1] : ''),
    descFornecedor: fornIdx.get(codF) || (cDescF ? rowVals[cDescF - 1] : ''),
    dataEntrada,
    isEntrada: true,
  });

  if (cProcQtd) entrada.getRange(row, cProcQtd).setValue(qtd);
  if (cProcAt) entrada.getRange(row, cProcAt).setValue(new Date());

  return { ok: true, message: 'Entrada consolidada no centro.', delta };
}

/* ============================================================
   PROCESSAR SAÍDA (movimentação origem → destino)
   ============================================================ */

function processSaidaRow_(row) {
  const saida = getSheetOrThrow_(CFG.SHEETS.SAIDA);
  const produtosIdx = buildProdutosIndex_();
  const fornIdx = buildFornecedorIndex_();

  const cCod = findCol_(saida, CFG.HEADERS.COD_BARRAS);
  const cDesc = findCol_(saida, CFG.HEADERS.DESCRICAO);
  const cInt = findCol_(saida, CFG.HEADERS.COD_INTERNO);
  const cMerc = findCol_(saida, CFG.HEADERS.MERCADOLOGICO);
  const cQtd = findCol_(saida, CFG.HEADERS.QTD);
  const cLote = findCol_(saida, CFG.HEADERS.LOTE);
  const cVal = findCol_(saida, CFG.HEADERS.VALIDADE);
  const cCodF = findCol_(saida, CFG.HEADERS.COD_FORNECEDOR);
  const cDescF = findCol_(saida, CFG.HEADERS.DESC_FORNECEDOR);
  const cOrigemEnd = findCol_(saida, CFG.HEADERS.ORIGEM_END);
  const cORua = findCol_(saida, CFG.HEADERS.S_O_RUA);
  const cOCol = findCol_(saida, CFG.HEADERS.S_O_COL);
  const cOVol = findCol_(saida, CFG.HEADERS.S_O_VOL);
  const cONiv = findCol_(saida, CFG.HEADERS.S_O_NIV);
  const cDRua = findCol_(saida, CFG.HEADERS.S_D_RUA);
  const cDCol = findCol_(saida, CFG.HEADERS.S_D_COL);
  const cDVol = findCol_(saida, CFG.HEADERS.S_D_VOL);
  const cDNiv = findCol_(saida, CFG.HEADERS.S_D_NIV);
  const cProcessed = findCol_(saida, CFG.HEADERS.PROCESSED);
  const cProcessedAt = findCol_(saida, CFG.HEADERS.PROCESSED_AT);
  const cUuid = findCol_(saida, CFG.HEADERS.UUID);

  if (!cCod || !cQtd || !cLote || !cVal || !cCodF || !cOrigemEnd || !cDRua || !cDCol || !cDVol || !cDNiv) {
    throw new Error('Saída: faltam colunas obrigatórias (código, qtd, lote, validade, fornecedor, origem, destino).');
  }

  const rowVals = saida.getRange(row, 1, 1, saida.getLastColumn()).getValues()[0];

  // Evita reprocessar linha já processada
  if (cProcessed && rowVals[cProcessed - 1] === true) {
    return { ok: true, message: 'Linha já processada (ignorada).' };
  }

  const cod = String(rowVals[cCod - 1] ?? '').trim().split('.')[0];
  if (!cod) throw new Error(`Saída L${row}: código de barras vazio.`);

  const prod = produtosIdx.get(cod);
  if (prod) {
    if (cDesc && isEmpty_(rowVals[cDesc - 1])) saida.getRange(row, cDesc).setValue(prod.descricao);
    if (cInt && isEmpty_(rowVals[cInt - 1])) saida.getRange(row, cInt).setValue(prod.codInterno);
    if (cMerc && isEmpty_(rowVals[cMerc - 1])) saida.getRange(row, cMerc).setValue(prod.mercadologico);
  }

  const qtd = toNumber_(rowVals[cQtd - 1]);
  if (qtd <= 0) throw new Error(`Saída L${row}: quantidade inválida (${rowVals[cQtd - 1]}).`);

  const lote = String(rowVals[cLote - 1] || '').trim();
  const rawVal = rowVals[cVal - 1];
  const val = rawVal instanceof Date
    ? Utilities.formatDate(rawVal, Session.getScriptTimeZone(), 'dd/MM/yyyy')
    : String(rawVal || '').trim();
  if (!lote || !val) throw new Error(`Saída L${row}: lote e validade são obrigatórios.`);

  const codF = String(rowVals[cCodF - 1] || '').trim();
  if (!codF) throw new Error(`Saída L${row}: Código Fornecedor obrigatório.`);

  if (cDescF && isEmpty_(rowVals[cDescF - 1]) && fornIdx.has(codF)) {
    saida.getRange(row, cDescF).setValue(fornIdx.get(codF));
  }

  // Origem: parse do dropdown
  const origemTxt = String(rowVals[cOrigemEnd - 1] || '').trim();
  const orig = parseOrigemEndereco_(origemTxt);
  if (!orig) throw new Error(`Saída L${row}: selecione a origem no dropdown "Origem endereço (Dropdown)".`);

  const errOrig = validateEndereco_(orig.rua, orig.col, orig.vol, orig.niv);
  if (errOrig) throw new Error(`Saída L${row}: origem inválida. ${errOrig}`);

  // Grava campos de origem (visíveis)
  if (cORua) saida.getRange(row, cORua).setValue(orig.rua);
  if (cOCol) saida.getRange(row, cOCol).setValue(orig.col);
  if (cOVol) saida.getRange(row, cOVol).setValue(orig.vol);
  if (cONiv) saida.getRange(row, cONiv).setValue(orig.niv);

  // Destino
  const dRua = +rowVals[cDRua - 1];
  const dCol = +rowVals[cDCol - 1];
  const dVol = +rowVals[cDVol - 1];
  const dNiv = +rowVals[cDNiv - 1];

  const errDest = validateEndereco_(dRua, dCol, dVol, dNiv);
  if (errDest) throw new Error(`Saída L${row}: destino inválido. ${errDest}`);

  if (cUuid && isEmpty_(rowVals[cUuid - 1])) saida.getRange(row, cUuid).setValue(Utilities.getUuid());

  const shared = {
    codBarras: cod,
    lote,
    validade: val,
    codFornecedor: codF,
    descricao: prod ? prod.descricao : (cDesc ? rowVals[cDesc - 1] : ''),
    codInterno: prod ? prod.codInterno : (cInt ? rowVals[cInt - 1] : ''),
    mercadologico: prod ? prod.mercadologico : (cMerc ? rowVals[cMerc - 1] : ''),
    descFornecedor: fornIdx.get(codF) || (cDescF ? rowVals[cDescF - 1] : ''),
  };

  // 1) Subtrai da origem
  applyDeltaCentro_({
    ...shared,
    endereco: { rua: orig.rua, col: orig.col, vol: orig.vol, niv: orig.niv },
    deltaQtd: -qtd,
    dataEntrada: '',
    isEntrada: false,
  });

  // 2) Soma no destino (registra data de chegada)
  applyDeltaCentro_({
    ...shared,
    endereco: { rua: dRua, col: dCol, vol: dVol, niv: dNiv },
    deltaQtd: +qtd,
    dataEntrada: new Date(),
    isEntrada: true,
  });

  if (cProcessed) saida.getRange(row, cProcessed).setValue(true);
  if (cProcessedAt) saida.getRange(row, cProcessedAt).setValue(new Date());

  return { ok: true, message: 'Movimentação processada.', qtd };
}

/* ============================================================
   CENTRO — cria / soma / subtrai / deleta linha
   ============================================================ */

function centroKey_(codBarras, validade, lote, codFornecedor, rua, col, vol, niv) {
  return [codBarras, validade, lote, codFornecedor, rua, col, vol, niv]
    .map(v => String(v || '').trim()).join('|');
}

function applyDeltaCentro_(p) {
  const centro = getSheetOrThrow_(CFG.SHEETS.CENTRO);

  const cCod = findCol_(centro, CFG.HEADERS.COD_BARRAS);
  const cDesc = findCol_(centro, CFG.HEADERS.DESCRICAO);
  const cInt = findCol_(centro, CFG.HEADERS.COD_INTERNO);
  const cQtd = findCol_(centro, CFG.HEADERS.QTD);
  const cLote = findCol_(centro, CFG.HEADERS.LOTE);
  const cVal = findCol_(centro, CFG.HEADERS.VALIDADE);
  const cCodF = findCol_(centro, CFG.HEADERS.COD_FORNECEDOR);
  const cDescF = findCol_(centro, ['Descrição do fornecedor', 'Fornecedor']);
  const cDataEnt = findCol_(centro, CFG.HEADERS.DATA_ENTRADA);
  const cRua = findCol_(centro, CFG.HEADERS.C_RUA);
  const cCol = findCol_(centro, CFG.HEADERS.C_COL);
  const cVol = findCol_(centro, CFG.HEADERS.C_VOL);
  const cNiv = findCol_(centro, CFG.HEADERS.C_NIV);
  const cMerc = findCol_(centro, CFG.HEADERS.MERCADOLOGICO);
  const cDataAt = findCol_(centro, CFG.HEADERS.DATA_ATUALIZACAO);

  if (!cCod || !cQtd || !cLote || !cVal || !cCodF || !cRua || !cCol || !cVol || !cNiv) {
    throw new Error('Centro: faltam colunas obrigatórias.');
  }

  const cod = String(p.codBarras ?? '').trim().split('.')[0];
  const lote = String(p.lote || '').trim();
  const val = String(p.validade || '').trim();
  const codF = String(p.codFornecedor || '').trim();
  const rua = +p.endereco?.rua;
  const col = +p.endereco?.col;
  const vol = +p.endereco?.vol;
  const niv = +p.endereco?.niv;
  const delta = toNumber_(p.deltaQtd);

  if (!cod || !lote || !val || !codF) throw new Error('Centro: chave incompleta (código/lote/validade/fornecedor).');
  const errEnd = validateEndereco_(rua, col, vol, niv);
  if (errEnd) throw new Error(`Centro: endereço inválido. ${errEnd}`);
  if (delta === 0) return;

  const key = centroKey_(cod, val, lote, codF, rua, col, vol, niv);
  const lastRow = centro.getLastRow();
  const lastCol = centro.getLastColumn();

  // Monta vetor para nova linha
  const buildNewRow_ = () => {
    const r = new Array(lastCol).fill('');
    r[cCod - 1] = cod;
    if (cDesc) r[cDesc - 1] = p.descricao || '';
    if (cInt) r[cInt - 1] = p.codInterno || '';
    r[cQtd - 1] = delta;
    r[cLote - 1] = lote;
    r[cVal - 1] = val;
    r[cCodF - 1] = codF;
    if (cDescF) r[cDescF - 1] = p.descFornecedor || '';
    if (cDataEnt && p.isEntrada && !isEmpty_(p.dataEntrada)) r[cDataEnt - 1] = p.dataEntrada;
    r[cRua - 1] = rua;
    r[cCol - 1] = col;
    r[cVol - 1] = vol;
    r[cNiv - 1] = niv;
    if (cMerc) r[cMerc - 1] = p.mercadologico || '';
    if (cDataAt) r[cDataAt - 1] = new Date();
    return r;
  };

  // Centro vazio — cria primeira linha
  if (lastRow < 2) {
    if (delta < 0) throw new Error('Centro: não existe saldo para subtrair (centro vazio).');
    centro.appendRow(buildNewRow_());
    return;
  }

  // Busca linha existente pela chave composta
  const disp = centro.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  const vals = centro.getRange(2, 1, lastRow - 1, lastCol).getValues();

  let foundRow = null;
  let foundQty = null;

  for (let i = 0; i < disp.length; i++) {
    const k = centroKey_(
      String(disp[i][cCod - 1] || '').trim(),
      String(disp[i][cVal - 1] || '').trim(),
      String(disp[i][cLote - 1] || '').trim(),
      String(disp[i][cCodF - 1] || '').trim(),
      String(disp[i][cRua - 1] || '').trim(),
      String(disp[i][cCol - 1] || '').trim(),
      String(disp[i][cVol - 1] || '').trim(),
      String(disp[i][cNiv - 1] || '').trim(),
    );
    if (k === key) {
      foundRow = i + 2;
      foundQty = toNumber_(vals[i][cQtd - 1]);
      break;
    }
  }

  // Linha não encontrada → cria nova
  if (foundRow === null) {
    if (delta < 0) throw new Error(`Centro: sem saldo na origem (${buildAddrKey_(rua, col, vol, niv)}) para subtrair.`);
    centro.appendRow(buildNewRow_());
    return;
  }

  const newQty = foundQty + delta;
  if (newQty < -1e-9) {
    throw new Error(`Centro: saldo insuficiente. Saldo ${formatQty_(foundQty)}, tentando mover ${formatQty_(-delta)}.`);
  }

  // Saldo zerado → remove linha (sem lixo no Centro)
  if (Math.abs(newQty) < 1e-9) {
    centro.deleteRow(foundRow);
    return;
  }

  // Atualização em lote — 1 getValues() + 1 setValues() por movimentação
  const rowRange = centro.getRange(foundRow, 1, 1, lastCol);
  const rowData = rowRange.getValues()[0];

  rowData[cQtd - 1] = newQty;
  if (cDesc && p.descricao) rowData[cDesc - 1] = p.descricao;
  if (cInt && p.codInterno) rowData[cInt - 1] = p.codInterno;
  if (cMerc && p.mercadologico) rowData[cMerc - 1] = p.mercadologico;
  rowData[cCodF - 1] = codF;
  if (cDescF && p.descFornecedor) rowData[cDescF - 1] = p.descFornecedor;

  if (cDataEnt && p.isEntrada && !isEmpty_(p.dataEntrada)) {
    const current = rowData[cDataEnt - 1];
    if (current instanceof Date && p.dataEntrada instanceof Date) {
      if (p.dataEntrada.getTime() >= current.getTime()) rowData[cDataEnt - 1] = p.dataEntrada;
    } else if (isEmpty_(current)) {
      rowData[cDataEnt - 1] = p.dataEntrada;
    } else {
      rowData[cDataEnt - 1] = p.dataEntrada;
    }
  }

  if (cDataAt) rowData[cDataAt - 1] = new Date();
  rowRange.setValues([rowData]);
}