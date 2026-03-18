/**
 * Script para gerar abas analíticas (LK_US_BASE, LK_US_API, LK_US_TIMES, LK_SNAPSHOT)
 * a partir da aba "WBS Project" (ou alternativa) da planilha de Workflow / FTR Sinistro.
 *
 * Suporta múltiplos nomes de coluna (ex.: ID_US ou US Original, Data_Fim_Plan ou Data_Fim_Planejada).
 * Se Outline Level não existir, considera todas as linhas com ID_US/WBS.
 * Etapa/Processo derivados do WBS quando colunas vazias. LK_US_API e LK_US_TIMES
 * recebem ao menos uma linha por US (N/A quando sem APIs/Times).
 */

/*************** CONFIGURAÇÃO BÁSICA ***************/

// Aba de origem: deve ter ao menos coluna ID (US Original ou ID_US) e de preferência WBS.
// Se seus dados estiverem em outra aba (ex.: EF Funcionalidades RSI), altere aqui.
const SHEET_WBS_PROJECT = 'WBS Project';

// Abas de destino
const SHEET_LK_US_BASE  = 'LK_US_BASE';
const SHEET_LK_US_API   = 'LK_US_API';
const SHEET_LK_US_TIMES = 'LK_US_TIMES';
const SHEET_LK_SNAPSHOT = 'LK_SNAPSHOT';

// Nomes possíveis das COLUNAS (primeiro que existir na planilha será usado)
const COL_ID_US_ALT     = ['US Original', 'ID_US'];
const COL_OUTLINE_ALT   = ['Outline Level'];
const COL_WBS_ALT       = ['WBS'];
const COL_TASK_ALT      = ['Task Name', 'Funcionalidade'];
const COL_ETAPA_ALT     = ['Etapa'];
const COL_PROCESSO_ALT  = ['Processo'];
const COL_REGRA_ALT     = ['Regra'];
const COL_FUNC_ALT      = ['Funcionalidade', 'Task Name'];
const COL_DURACAO_ALT   = ['Duracao_Dias', 'Duracao'];
const COL_PRODUTO_ALT   = ['Produto'];
const COL_SIST_LEGADOS_ALT = ['Sistemas_Legados'];
const COL_APIS_ALT      = ['APIs', 'API'];
const COL_TIMES_ALT     = ['Times', 'Times_Envolvidos', 'Qtd_Times_Interacao'];
const COL_STATUS_ALT    = ['Status', 'Status_Interacao'];
const COL_PCT_ALT       = ['Pct_Realizado'];
const COL_DT_INI_ALT    = ['Data_Inicio_Planejada', 'Data_Inicio_Plan'];
const COL_DT_FIM_ALT    = ['Data_Fim_Planejada', 'Data_Fim_Plan'];
const COL_DT_FIM_REAL_ALT = ['Data_Fim_Real'];
const COL_RESP_ALT      = ['Responsavel'];


/*************** FUNÇÕES PRINCIPAIS ***************/

/**
 * Função principal para gerar/atualizar:
 *  - LK_US_BASE
 *  - LK_US_API
 *  - LK_US_TIMES
 *
 * Execute manualmente em um primeiro momento.
 * Depois você pode criar um gatilho de tempo (time-driven) no Apps Script
 * para rodar 1x ao dia ou conforme necessário.
 */
function atualizarBasesLK() {
  const ss = SpreadsheetApp.getActive();

  const wbsSheet = ss.getSheetByName(SHEET_WBS_PROJECT);
  if (!wbsSheet) {
    throw new Error('Aba origem "' + SHEET_WBS_PROJECT + '" não encontrada.');
  }

  const data = wbsSheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('Aba "' + SHEET_WBS_PROJECT + '" não tem dados suficientes.');
  }

  const header = data[0].map(function(h) { return String(h || '').trim(); });
  const rows = data.slice(1);

  // Mapeia índice de cada coluna pelo nome (case-sensitive match)
  const colIndex = {};
  header.forEach(function(name, i) {
    if (name) colIndex[name] = i;
  });

  function idxFirst(altNames) {
    for (var a = 0; a < altNames.length; a++) {
      if (colIndex[altNames[a]] !== undefined) return colIndex[altNames[a]];
    }
    return -1;
  }

  const iOutline = idxFirst(COL_OUTLINE_ALT);
  const iTask    = idxFirst(COL_TASK_ALT);
  const iWbs     = idxFirst(COL_WBS_ALT);
  const iUsOrig  = idxFirst(COL_ID_US_ALT);
  const iEtapa   = idxFirst(COL_ETAPA_ALT);
  const iProc    = idxFirst(COL_PROCESSO_ALT);
  const iRegra   = idxFirst(COL_REGRA_ALT);
  const iFunc    = idxFirst(COL_FUNC_ALT);
  const iDur     = idxFirst(COL_DURACAO_ALT);
  const iProd    = idxFirst(COL_PRODUTO_ALT);
  const iSist    = idxFirst(COL_SIST_LEGADOS_ALT);
  const iApis    = idxFirst(COL_APIS_ALT);
  const iTimes   = idxFirst(COL_TIMES_ALT);
  const iStatus  = idxFirst(COL_STATUS_ALT);
  const iPct     = idxFirst(COL_PCT_ALT);
  const iDtIni   = idxFirst(COL_DT_INI_ALT);
  const iDtFimP  = idxFirst(COL_DT_FIM_ALT);
  const iDtFimR  = idxFirst(COL_DT_FIM_REAL_ALT);
  const iResp    = idxFirst(COL_RESP_ALT);

  if (iUsOrig === -1) {
    throw new Error(
      'Na aba "' + SHEET_WBS_PROJECT + '" não foi encontrada nenhuma coluna de ID: ' +
      COL_ID_US_ALT.join(' ou ') + '. Verifique o cabeçalho da linha 1.'
    );
  }

  // Deriva Etapa e Processo a partir do WBS (ex.: "1.1.1.1" -> Etapa "1", Processo "1.1")
  function derivarEtapaProcesso(wbsStr) {
    var w = String(wbsStr || '').trim();
    if (!w) return { etapa: '', processo: '' };
    var partes = w.split(/\./).filter(Boolean);
    if (partes.length >= 1) {
      return {
        etapa: partes[0],
        processo: partes.length >= 2 ? partes[0] + '.' + partes[1] : partes[0]
      };
    }
    return { etapa: '', processo: '' };
  }

  // =================== Monta LK_US_BASE ===================
  const baseHeader = [
    'ID_US',
    'WBS',
    'Etapa',
    'Processo',
    'Regra',
    'Funcionalidade',
    'Duracao_Dias',
    'Produto',
    'Sistemas_Legados',
    'Qtd_APIs',
    'Qtd_Times_Envolvidos',
    'Tem_Interseccao',
    'Status',
    'Pct_Realizado',
    'Data_Inicio_Planejada',
    'Data_Fim_Planejada',
    'Data_Fim_Real',
    'Responsavel'
  ];

  const baseRows = [];

  rows.forEach(function(row) {
    // Se existir coluna Outline Level, filtrar só nível 4; senão incluir todas as linhas com ID
    if (iOutline >= 0) {
      var level = String(row[iOutline] || '').trim();
      if (level !== '4' && level !== 4) return;
    }

    var idUs = iUsOrig >= 0 ? (row[iUsOrig] != null ? String(row[iUsOrig]).trim() : '') : '';
    if (!idUs) return;

    var wbs     = iWbs   >= 0 ? (row[iWbs] != null ? row[iWbs] : '') : '';
    var dp      = derivarEtapaProcesso(wbs);
    var etapa   = (iEtapa >= 0 && row[iEtapa] != null && String(row[iEtapa]).trim() !== '')
      ? row[iEtapa] : dp.etapa;
    var processo = (iProc >= 0 && row[iProc] != null && String(row[iProc]).trim() !== '')
      ? row[iProc] : dp.processo;
    var regra   = iRegra >= 0 ? row[iRegra] : '';
    var func    = (iFunc >= 0 ? row[iFunc] : (iTask >= 0 ? row[iTask] : ''));
    var duracao = iDur   >= 0 ? row[iDur] : '';
    var produto = iProd  >= 0 ? row[iProd] : '';
    var sistLeg = iSist  >= 0 ? row[iSist] : '';
    var apisCell  = iApis  >= 0 ? row[iApis] : '';
    var timesCell = iTimes >= 0 ? row[iTimes] : '';
    var status  = iStatus >= 0 ? row[iStatus] : '';
    var pct     = iPct    >= 0 ? row[iPct] : '';
    var dtIni   = iDtIni  >= 0 ? row[iDtIni] : '';
    var dtFimPlan = iDtFimP >= 0 ? row[iDtFimP] : '';
    var dtFimReal = iDtFimR >= 0 ? row[iDtFimR] : '';
    var resp    = iResp   >= 0 ? row[iResp] : '';

    var apisStrVal = apisCell != null ? String(apisCell).trim() : '';
    var listApisVal = [];
    if (apisStrVal !== '') {
      if (/^\d+$/.test(apisStrVal)) {
        listApisVal = [];
      } else {
        listApisVal = apisStrVal.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
      }
    }
    var qtdApis = listApisVal.length;
    if (apisStrVal !== '' && /^\d+$/.test(apisStrVal)) qtdApis = parseInt(apisStrVal, 10) || 0;

    var timesStrVal = timesCell != null ? String(timesCell).trim() : '';
    var listTimesVal = [];
    if (timesStrVal !== '') {
      if (/^\d+$/.test(timesStrVal)) {
        listTimesVal = [];
      } else {
        listTimesVal = timesStrVal.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
      }
    }
    var qtdTimes = listTimesVal.length;
    if (timesStrVal !== '' && /^\d+$/.test(timesStrVal)) qtdTimes = parseInt(timesStrVal, 10) || 0;

    var temInter = qtdTimes > 1 ? 1 : 0;

    baseRows.push([
      idUs,
      wbs,
      etapa,
      processo,
      regra,
      func,
      duracao,
      produto,
      sistLeg,
      qtdApis,
      qtdTimes,
      temInter,
      status,
      pct,
      dtIni,
      dtFimPlan,
      dtFimReal,
      resp
    ]);
  });

  sobrescreverAba(ss, SHEET_LK_US_BASE, baseHeader, baseRows);

  // =================== Monta LK_US_API ===================
  const apiHeader = [
    'ID_US',
    'Etapa',
    'Processo',
    'Funcionalidade',
    'API',
    'Status',
    'Pct_Realizado'
  ];
  const apiRows = [];

  baseRows.forEach(function(r) {
    var idUs   = r[0];
    var etapa  = r[2];
    var proc   = r[3];
    var func   = r[5];
    var status = r[12];
    var pct    = r[13];

    var origRow = encontrarLinhaPorIdUs(rows, iUsOrig, idUs);
    var apisCell = (origRow && iApis >= 0) ? origRow[iApis] : null;
    var apisStr = apisCell != null ? String(apisCell).trim() : '';
    var apisList = [];
    if (apisStr !== '' && !/^\d+$/.test(apisStr)) {
      apisList = apisStr.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    }
    if (apisList.length === 0) {
      apiRows.push([idUs, etapa, proc, func, 'N/A', status, pct]);
    } else {
      apisList.forEach(function(api) {
        apiRows.push([idUs, etapa, proc, func, api, status, pct]);
      });
    }
  });

  sobrescreverAba(ss, SHEET_LK_US_API, apiHeader, apiRows);

  // =================== Monta LK_US_TIMES ===================
  const timesHeader = [
    'ID_US',
    'Etapa',
    'Processo',
    'Funcionalidade',
    'Time',
    'Status'
  ];
  const timesRows = [];

  baseRows.forEach(function(r) {
    var idUs   = r[0];
    var etapa  = r[2];
    var proc   = r[3];
    var func   = r[5];
    var status = r[12];

    var origRow = encontrarLinhaPorIdUs(rows, iUsOrig, idUs);
    var timesCell = (origRow && iTimes >= 0) ? origRow[iTimes] : null;
    var timesStr = timesCell != null ? String(timesCell).trim() : '';
    var timesList = [];
    if (timesStr !== '' && !/^\d+$/.test(timesStr)) {
      timesList = timesStr.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    }
    if (timesList.length === 0) {
      timesRows.push([idUs, etapa, proc, func, 'N/A', status]);
    } else {
      timesList.forEach(function(t) {
        timesRows.push([idUs, etapa, proc, func, t, status]);
      });
    }
  });

  sobrescreverAba(ss, SHEET_LK_US_TIMES, timesHeader, timesRows);
}

/**
 * Função para tirar snapshot e append em LK_SNAPSHOT.
 * Pode ser chamada manualmente ou via gatilho agendado.
 */
function tirarSnapshot() {
  const ss = SpreadsheetApp.getActive();
  const baseSheet = ss.getSheetByName(SHEET_LK_US_BASE);
  if (!baseSheet) {
    throw new Error('Aba "' + SHEET_LK_US_BASE + '" não existe. Rode atualizarBasesLK() antes.');
  }

  const data = baseSheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('Aba "' + SHEET_LK_US_BASE + '" não tem dados suficientes.');
  }

  const header = data[0];
  const rows = data.slice(1);

  const colIndex = {};
  header.forEach((name, i) => (colIndex[name] = i));

  function idx(name) {
    if (!(name in colIndex)) return -1;
    return colIndex[name];
  }

  const iID   = idx('ID_US');
  const iStat = idx('Status');
  const iPct  = idx('Pct_Realizado');
  const iDtFp = idx('Data_Fim_Planejada');

  const hoje = new Date();
  const snapshotRows = [];

  rows.forEach(r => {
    const id = iID   >= 0 ? r[iID]   : '';
    if (!id) return;
    const st = iStat >= 0 ? r[iStat] : '';
    const pc = iPct  >= 0 ? r[iPct]  : '';
    const dt = iDtFp >= 0 ? r[iDtFp] : '';
    snapshotRows.push([
      hoje,
      id,
      st,
      pc,
      dt
    ]);
  });

  let snapSheet = ss.getSheetByName(SHEET_LK_SNAPSHOT);
  if (!snapSheet) {
    snapSheet = ss.insertSheet(SHEET_LK_SNAPSHOT);
    snapSheet.appendRow(['Data_Snapshot', 'ID_US', 'Status', 'Pct_Realizado', 'Data_Fim_Planejada']);
  }

  if (snapshotRows.length > 0) {
    var startRow = snapSheet.getLastRow() + 1;
    var numCols = snapshotRows[0].length;
    snapSheet.getRange(startRow, 1, startRow + snapshotRows.length - 1, numCols).setValues(snapshotRows);
  }
}


/*************** HELPERS ***************/

/**
 * Cria/limpa uma aba e escreve cabeçalho + linhas.
 */
function sobrescreverAba(ss, sheetName, header, rows) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clearContents();
  }
  if (!rows || rows.length === 0) {
    sheet.appendRow(header);
    return;
  }
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
}

/**
 * Encontra linha original na WBS Project usando ID_US.
 */
function encontrarLinhaPorIdUs(rows, colIdx, idUs) {
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][colIdx] == idUs) return rows[i];
  }
  return null;
}

