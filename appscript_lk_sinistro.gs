/**
 * Script para gerar abas analíticas (LK_US_BASE, LK_US_API, LK_US_TIMES, LK_SNAPSHOT)
 * a partir da aba "WBS Project" da planilha de Workflow / FTR Sinistro.
 *
 * Cole este código em:
 *  Extensões → Apps Script, na própria planilha de especificação.
 *
 * Ajuste apenas a seção de CONFIGURAÇÃO se os nomes de abas/colunas forem diferentes.
 */

/*************** CONFIGURAÇÃO BÁSICA ***************/

// Nome da aba de origem principal (WBS Project)
const SHEET_WBS_PROJECT = 'WBS Project';

// Abas de destino (serão criadas/limpas e recriadas)
const SHEET_LK_US_BASE  = 'LK_US_BASE';
const SHEET_LK_US_API   = 'LK_US_API';
const SHEET_LK_US_TIMES = 'LK_US_TIMES';
const SHEET_LK_SNAPSHOT = 'LK_SNAPSHOT'; // só é append, não limpa tudo

// Nomes exatos das COLUNAS na aba WBS Project (linha 1)
const COL_OUTLINE_LEVEL = 'Outline Level';
const COL_TASK_NAME     = 'Task Name';
const COL_WBS           = 'WBS';
const COL_US_ORIGINAL   = 'US Original';     // vira ID_US
const COL_ETAPA         = 'Etapa';           // se não existir, pode ser calculado por hierarquia
const COL_PROCESSO      = 'Processo';        // idem
const COL_REGRA         = 'Regra';           // opcional
const COL_FUNC          = 'Funcionalidade';  // opcional
const COL_DURACAO       = 'Duracao_Dias';
const COL_PRODUTO       = 'Produto';
const COL_SIST_LEGADOS  = 'Sistemas_Legados';
const COL_APIS          = 'APIs';            // coluna com lista: "API-01, API-13"
const COL_TIMES         = 'Times';           // coluna com lista: "GEFEP, GETIV"
const COL_STATUS        = 'Status';
const COL_PCT_REAL      = 'Pct_Realizado';
const COL_DT_INI_PLAN   = 'Data_Inicio_Planejada';
const COL_DT_FIM_PLAN   = 'Data_Fim_Planejada';
const COL_DT_FIM_REAL   = 'Data_Fim_Real';
const COL_RESP          = 'Responsavel';


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

  const header = data[0];
  const rows = data.slice(1);

  // Mapeia índice de cada coluna pelo nome
  const colIndex = {};
  header.forEach((name, i) => {
    colIndex[name] = i;
  });

  function idx(colName) {
    if (!(colName in colIndex)) return -1;
    return colIndex[colName];
  }

  const iOutline = idx(COL_OUTLINE_LEVEL);
  const iTask    = idx(COL_TASK_NAME);
  const iWbs     = idx(COL_WBS);
  const iUsOrig  = idx(COL_US_ORIGINAL);
  const iEtapa   = idx(COL_ETAPA);
  const iProc    = idx(COL_PROCESSO);
  const iRegra   = idx(COL_REGRA);
  const iFunc    = idx(COL_FUNC);
  const iDur     = idx(COL_DURACAO);
  const iProd    = idx(COL_PRODUTO);
  const iSist    = idx(COL_SIST_LEGADOS);
  const iApis    = idx(COL_APIS);
  const iTimes   = idx(COL_TIMES);
  const iStatus  = idx(COL_STATUS);
  const iPct     = idx(COL_PCT_REAL);
  const iDtIni   = idx(COL_DT_INI_PLAN);
  const iDtFimP  = idx(COL_DT_FIM_PLAN);
  const iDtFimR  = idx(COL_DT_FIM_REAL);
  const iResp    = idx(COL_RESP);

  if (iOutline === -1 || iUsOrig === -1) {
    throw new Error(
      'Certifique-se de ter pelo menos as colunas "' +
      COL_OUTLINE_LEVEL + '" e "' + COL_US_ORIGINAL +
      '" na aba "' + SHEET_WBS_PROJECT + '".'
    );
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

  rows.forEach(row => {
    const outlineLevel = row[iOutline];
    if (String(outlineLevel).trim() !== '4') {
      // Só pega US (nível 4)
      return;
    }

    const idUs = iUsOrig >= 0 ? row[iUsOrig] : '';
    if (!idUs) return;

    const wbs        = iWbs    >= 0 ? row[iWbs]    : '';
    const etapa      = iEtapa  >= 0 ? row[iEtapa]  : '';
    const processo   = iProc   >= 0 ? row[iProc]   : '';
    const regra      = iRegra  >= 0 ? row[iRegra]  : '';
    const func       = iFunc   >= 0 ? row[iFunc]   : (iTask >= 0 ? row[iTask] : '');
    const duracao    = iDur    >= 0 ? row[iDur]    : '';
    const produto    = iProd   >= 0 ? row[iProd]   : '';
    const sistLeg    = iSist   >= 0 ? row[iSist]   : '';
    const apisCell   = iApis   >= 0 ? row[iApis]   : '';
    const timesCell  = iTimes  >= 0 ? row[iTimes]  : '';
    const status     = iStatus >= 0 ? row[iStatus] : '';
    const pct        = iPct    >= 0 ? row[iPct]    : '';
    const dtIni      = iDtIni  >= 0 ? row[iDtIni]  : '';
    const dtFimPlan  = iDtFimP >= 0 ? row[iDtFimP] : '';
    const dtFimReal  = iDtFimR >= 0 ? row[iDtFimR] : '';
    const resp       = iResp   >= 0 ? row[iResp]   : '';

    const qtdApis = apisCell
      ? String(apisCell).split(',').filter(s => s.trim()).length
      : 0;
    const qtdTimes = timesCell
      ? String(timesCell).split(',').filter(s => s.trim()).length
      : 0;
    const temInter = qtdTimes > 1 ? 1 : 0;

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

  baseRows.forEach(r => {
    const idUs   = r[0];
    const etapa  = r[2];
    const proc   = r[3];
    const func   = r[5];
    const status = r[12];
    const pct    = r[13];

    const origRow = encontrarLinhaPorIdUs(rows, iUsOrig, idUs);
    if (!origRow || iApis < 0) return;
    const apisCell = origRow[iApis];
    if (!apisCell) return;

    String(apisCell).split(',').forEach(api => {
      const limp = api.trim();
      if (!limp) return;
      apiRows.push([
        idUs,
        etapa,
        proc,
        func,
        limp,
        status,
        pct
      ]);
    });
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

  baseRows.forEach(r => {
    const idUs   = r[0];
    const etapa  = r[2];
    const proc   = r[3];
    const func   = r[5];
    const status = r[12];

    const origRow = encontrarLinhaPorIdUs(rows, iUsOrig, idUs);
    if (!origRow || iTimes < 0) return;
    const timesCell = origRow[iTimes];
    if (!timesCell) return;

    String(timesCell).split(',').forEach(t => {
      const limp = t.trim();
      if (!limp) return;
      timesRows.push([
        idUs,
        etapa,
        proc,
        func,
        limp,
        status
      ]);
    });
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

  snapSheet
    .getRange(snapSheet.getLastRow() + 1, 1, snapshotRows.length, snapshotRows[0].length)
    .setValues(snapshotRows);
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

