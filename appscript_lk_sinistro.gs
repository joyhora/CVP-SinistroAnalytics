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
// Coluna de detalhamento de regra/US na WBS (ex.: "US-FUNCIONALIDADE")
const WBS_COL_REGRA_DET = 'US-FUNCIONALIDADE';

// Colunas na aba Validação Cruzida EF×WBS
// Não dependemos do nome exato: usamos palavras‑chave mínimas.
// Pensando na sua planilha atual:
//  - "ETAPA (Baseline)"               -> contém "etapa"
//  - "PROCESSO (Baseline)"            -> contém "processo"
//  - "FUNCIONALIDADE (Baseline)"      -> contém "funcionalidade"
//  - "REGRA DE NEGÓCIO (Baseline)"    -> contém "regra" e "negocio"
//  - "COBERTURA\nNA WBS"              -> contém "cobertura"
//  - "IDs USs MAPEADAS (WBS)"         -> contém "ids us"
//  - "APIs ENVOLVIDAS (WBS)"          -> contém "apis"
const VAL_COL_ETAPA_TOKENS     = ['etapa'];
const VAL_COL_PROCESSO_TOKENS  = ['processo'];
const VAL_COL_FUNC_TOKENS      = ['funcionalidade'];
const VAL_COL_REGRA_TOKENS     = ['regra', 'negocio'];
const VAL_COL_COBERTURA_TOKENS = ['cobertura'];
const VAL_COL_IDS_US_TOKENS    = ['ids us'];
const VAL_COL_APIS_TOKENS      = ['apis'];


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
  const iRegDet = wbsHeaderIdx[WBS_COL_REGRA_DET];

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
      FuncOrig: iTask !== undefined ? row[iTask] : '',
      RegraDet: iRegDet !== undefined ? row[iRegDet] : ''
    };
  });

  // ===== 2) Validação Cruzada → explode IDs USs mapeadas =====
  const valData = valSheet.getDataRange().getValues();
  if (valData.length < 2) {
    throw new Error('Aba "' + SHEET_VALIDACAO + '" não tem dados suficientes.');
  }

  // Detecta automaticamente a linha de cabeçalho (onde aparecem Etapa/Processo/Funcionalidade)
  var headerRowIndex = -1;
  for (var r = 0; r < Math.min(valData.length, 20); r++) {
    var rowText = safeTrim_(valData[r].join(' ')).toLowerCase();
    if (rowText.includes('etapa') && rowText.includes('processo') && rowText.includes('funcionalidade')) {
      headerRowIndex = r;
      break;
    }
  }
  if (headerRowIndex === -1) {
    throw new Error('Não encontrei a linha de cabeçalho na aba "' + SHEET_VALIDACAO +
      '". Verifique se alguma linha contém \"ETAPA\", \"PROCESSO\" e \"FUNCIONALIDADE\".');
  }

  const valHeaderIdx = indexByName_(valData[headerRowIndex]);
  const valRows      = valData.slice(headerRowIndex + 1);

  const iEtapa = findColByTokens_(valHeaderIdx, VAL_COL_ETAPA_TOKENS);
  const iProc  = findColByTokens_(valHeaderIdx, VAL_COL_PROCESSO_TOKENS);
  const iFunc  = findColByTokens_(valHeaderIdx, VAL_COL_FUNC_TOKENS);
  const iRegra = findColByTokens_(valHeaderIdx, VAL_COL_REGRA_TOKENS);
  const iCob   = findColByTokens_(valHeaderIdx, VAL_COL_COBERTURA_TOKENS);
  const iIds   = findColByTokens_(valHeaderIdx, VAL_COL_IDS_US_TOKENS);
  const iApis  = findColByTokens_(valHeaderIdx, VAL_COL_APIS_TOKENS);

  if ([iEtapa, iProc, iFunc, iRegra, iCob, iIds].some(v => v === undefined)) {
    throw new Error(
      'Na aba "' + SHEET_VALIDACAO +
      '" não consegui localizar automaticamente as colunas de Etapa/Processo/Funcionalidade/Regra/Cobertura/IDs USs. ' +
      'Verifique se os cabeçalhos dessas colunas contêm, respectivamente, as palavras "Etapa", "Processo", "Funcionalidade", "Regra"/"Negócio", "Cobertura" e "IDs US".'
    );
  }

  const mapaVal = {};
  valRows.forEach(row => {
    const etapa = row[iEtapa];
    const proc  = row[iProc];
    const func  = row[iFunc];
    const regra = iRegra !== undefined ? row[iRegra] : '';
    const cob   = safeTrim_(row[iCob]); // COBERTO / PARCIAL / X% SEM COBERTURA
    const apis  = iApis !== undefined ? safeTrim_(row[iApis]) : '';

    const idsStr = safeTrim_(row[iIds]);
    if (!idsStr) return;

    idsStr.split(',')
      .map(s => safeTrim_(s))
      .filter(Boolean)
      .forEach(idUs => {
        if (!mapaVal[idUs]) {
          mapaVal[idUs] = {
            Etapa: etapa,
            Processo: proc,
            Func: func,
            Regra: regra,
            Cobertura: cob,
            Apis: apis
          };
        }
      });
  });

  // ===== 3) Montar LK_US_BASE =====
  const lkHeader = [
    'ID_US',
    'Etapa',
    'Processo',
    'Regra',
    'Funcionalidade',
    'Regra_Detalhada_WBS',
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
    const regra = v.Regra || '';
    const func  = v.Func || w.FuncOrig || '';
    const regraDet = w.RegraDet || '';
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
      regra,
      func,
      regraDet,
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

/**
 * Encontra o índice de coluna procurando por um conjunto de palavras‑chave
 * no cabeçalho, ignorando acentos e maiúsculas/minúsculas.
 */
function findColByTokens_(headerIdx, tokens) {
  const normalize = s =>
    safeTrim_(s)
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, ''); // remove acentos

  const normTokens = tokens.map(normalize);

  for (const [name, i] of Object.entries(headerIdx)) {
    const normName = normalize(name);
    const matchesAll = normTokens.every(tok => normName.includes(tok));
    if (matchesAll) return i;
  }
  return undefined;
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