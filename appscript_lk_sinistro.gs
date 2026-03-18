/**
 * Gera a base analítica principal (LK_US_BASE) a partir de:
 *
 *  - "WBS Project"  → ID_US, WBS, duração, sistemas, funcionalidade original
 *  - "Validação Cruzada EF×WBS" → Etapa, Processo, Funcionalidade (Baseline),
 *    Cobertura na WBS, APIs envolvidas, IDs USs mapeadas
 *
 * Resultado: uma linha por US em LK_US_BASE, pronta para visão por
 * Etapa / Processo / Funcionalidade, quantidade, %, status e previsão.
 */

/*************** CONFIGURAÇÃO ***************/

// Abas de origem
const SHEET_WBS_PROJECT = 'WBS Project';
const SHEET_VALIDACAO   = 'Validação Cruzada EF×WBS'; // nome EXATO da aba

// Aba de destino
const SHEET_LK_US_BASE  = 'LK_US_BASE';

// Colunas na aba WBS Project
const WBS_COL_ID_US   = 'US Original';
const WBS_COL_WBS     = 'WBS';
const WBS_COL_TASK    = 'Task Name';
const WBS_COL_DUR     = 'Duracao_Dias';
const WBS_COL_SIST    = 'Sistemas_Legados';

// Colunas na aba Validação Cruzada EF×WBS (use os nomes EXATOS da linha 1)
const VAL_COL_ETAPA     = 'ETAPA (Baseline)';
const VAL_COL_PROCESSO  = 'PROCESSO (Baseline)';
const VAL_COL_FUNC      = 'FUNCIONALIDADE (Baseline)';
const VAL_COL_COBERTURA = 'COBERTURA NA WBS';
const VAL_COL_IDS_US    = 'IDs USs MAPEADAS (WBS)';
const VAL_COL_APIS      = 'APIs ENVOLVIDAS (WBS)';


/*************** FUNÇÃO PRINCIPAL ***************/

function atualizarBasesLK() {
  const ss = SpreadsheetApp.getActive();

  const wbsSheet = ss.getSheetByName(SHEET_WBS_PROJECT);
  const valSheet = ss.getSheetByName(SHEET_VALIDACAO);

  if (!wbsSheet) throw new Error('Aba origem "' + SHEET_WBS_PROJECT + '" não encontrada.');
  if (!valSheet) throw new Error('Aba origem "' + SHEET_VALIDACAO + '" não encontrada.');

  // ===== 1) WBS Project → mapa básico por ID_US =====
  const wbsData = wbsSheet.getDataRange().getValues();
  if (wbsData.length < 2) throw new Error('Aba "' + SHEET_WBS_PROJECT + '" não tem dados suficientes.');

  const wbsHeaderIdx = indexByName_(wbsData[0]);
  const wbsRows      = wbsData.slice(1);

  const iIdUs = wbsHeaderIdx[WBS_COL_ID_US];
  const iWbs  = wbsHeaderIdx[WBS_COL_WBS];
  const iTask = wbsHeaderIdx[WBS_COL_TASK];
  const iDur  = wbsHeaderIdx[WBS_COL_DUR];
  const iSist = wbsHeaderIdx[WBS_COL_SIST];

  if (iIdUs === undefined) {
    throw new Error('Na aba "' + SHEET_WBS_PROJECT + '" não encontrei a coluna "' + WBS_COL_ID_US + '".');
  }

  const mapaWbs = {};
  wbsRows.forEach(row => {
    const idUs = safeTrim_(row[iIdUs]);
    if (!idUs) return;
    mapaWbs[idUs] = {
      WBS:      iWbs  !== undefined ? row[iWbs]  : '',
      Duracao:  iDur  !== undefined ? row[iDur]  : '',
      Sistemas: iSist !== undefined ? row[iSist] : '',
      FuncOrig: iTask !== undefined ? row[iTask] : ''
    };
  });

  // ===== 2) Validação Cruzada → explode IDs USs mapeadas =====
  const valData = valSheet.getDataRange().getValues();
  if (valData.length < 2) throw new Error('Aba "' + SHEET_VALIDACAO + '" não tem dados suficientes.');

  const valHeaderIdx = indexByName_(valData[0]);
  const valRows      = valData.slice(1);

  const iEtapa = valHeaderIdx[VAL_COL_ETAPA];
  const iProc  = valHeaderIdx[VAL_COL_PROCESSO];
  const iFunc  = valHeaderIdx[VAL_COL_FUNC];
  const iCob   = valHeaderIdx[VAL_COL_COBERTURA];
  const iIds   = valHeaderIdx[VAL_COL_IDS_US];
  const iApis  = valHeaderIdx[VAL_COL_APIS];

  if ([iEtapa, iProc, iFunc, iCob, iIds].some(v => v === undefined)) {
    throw new Error(
      'Na aba "' + SHEET_VALIDACAO +
      '" faltam colunas obrigatórias: "' +
      [VAL_COL_ETAPA, VAL_COL_PROCESSO, VAL_COL_FUNC, VAL_COL_COBERTURA, VAL_COL_IDS_US].join('", "') +
      '".'
    );
  }

  const mapaVal = {};
  valRows.forEach(row => {
    const etapa = row[iEtapa];
    const proc  = row[iProc];
    const func  = row[iFunc];
    const cob   = safeTrim_(row[iCob]); // COBERTO / PARCIAL / X% SEM COBERTURA
    const apis  = iApis !== undefined ? safeTrim_(row[iApis]) : '';

    const idsStr = safeTrim_(row[iIds]);
    if (!idsStr) return;

    idsStr.split(',')
      .map(s => safeTrim_(s))
      .filter(Boolean)
      .forEach(idUs => {
        if (!mapaVal[idUs]) {
          mapaVal[idUs] = { Etapa: etapa, Processo: proc, Func: func, Cobertura: cob, Apis: apis };
        }
      });
  });

  // ===== 3) Montar LK_US_BASE =====
  const lkHeader = [
    'ID_US',
    'Etapa',
    'Processo',
    'Funcionalidade',
    'WBS',
    'Duracao_Dias',
    'Produto',
    'Sistemas_Legados',
    'Qtd_APIs',
    'Qtd_Times_Envolvidos',
    'Tem_Interseccao',
    'Status_Projeto',
    'Pct_Realizado',
    'Data_Inicio_Projeto',
    'Data_Fim_Projeto',
    'Data_Fim_Planejada',
    'Jira_Key',
    'Jira_Link',
    'Responsavel'
  ];

  const linhas = [];

  Object.keys(mapaWbs).forEach(idUs => {
    const w = mapaWbs[idUs] || {};
    const v = mapaVal[idUs] || {};

    const etapa = v.Etapa || '';
    const proc  = v.Processo || '';
    const func  = v.Func || w.FuncOrig || '';
    const wbs   = w.WBS || '';
    const dur   = w.Duracao || '';
    const sist  = w.Sistemas || '';

    let statusProj = '';
    let pctReal    = 0;
    let dtFimProj  = '';

    const cob = (v.Cobertura || '').toUpperCase();
    if (cob.includes('COBERTO') && !cob.includes('SEM')) {
      statusProj = 'Concluído';
      pctReal    = 1;
      dtFimProj  = new Date(2026, 2, 31); // 31/03/2026
    } else if (cob.includes('PARCIAL')) {
      statusProj = 'Em andamento';
      pctReal    = 0.5;
    } else {
      statusProj = 'Não iniciado';
      pctReal    = 0;
    }

    let qtdApis = 0;
    if (v.Apis) {
      qtdApis = v.Apis.split(',')
        .map(s => safeTrim_(s))
        .filter(Boolean).length;
    }

    linhas.push([
      idUs,
      etapa,
      proc,
      func,
      wbs,
      dur,
      '',
      sist,
      qtdApis,
      0,
      0,
      statusProj,
      pctReal,
      '',
      dtFimProj,
      '',
      '',
      '',
      ''
    ]);
  });

  escreverAba_(ss, SHEET_LK_US_BASE, lkHeader, linhas);
}


/*************** HELPERS ***************/

function indexByName_(headerRow) {
  const idx = {};
  headerRow.forEach((name, i) => {
    const key = safeTrim_(name);
    if (key) idx[key] = i;
  });
  return idx;
}

function safeTrim_(v) {
  return v == null ? '' : String(v).trim();
}

function escreverAba_(ss, sheetName, header, rows) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  else sheet.clearContents();

  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows && rows.length) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
}