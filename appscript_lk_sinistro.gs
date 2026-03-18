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
const SHEET_WBS_PROJECT = 'WBS Project';              // estrutura WBS + US Original
const SHEET_WBS_DET     = 'WBS';                      // detalhamento por US (US-FUNCIONALIDADE)
const SHEET_VALIDACAO   = 'Validação Cruzada EF×WBS'; // visão executiva / baseline

// Aba de destino
const SHEET_LK_US_BASE  = 'LK_US_BASE';

// Colunas na aba WBS Project
const WBS_COL_ID_US   = 'US Original';
const WBS_COL_WBS     = 'WBS';
const WBS_COL_TASK    = 'Task Name';
const WBS_COL_DUR     = 'Duracao_Dias';
const WBS_COL_SIST    = 'Sistemas_Legados';
// Coluna de detalhamento de regra/US na aba WBS (algumas planilhas podem usar "US-FUNCIONALIDADE" ou "US+FUNCIONALIDADE")
const WBS_COL_REGRA_DET_PRI = 'US-FUNCIONALIDADE';
const WBS_COL_REGRA_DET_ALT = 'US+FUNCIONALIDADE';
// Coluna de ID da história na aba WBS (usada para casar com ID_US)
const WBS_DET_COL_ID_HIST   = 'ID HISTÓRIA';

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
  const iTask   = wbsHeaderIdx[WBS_COL_TASK];
  const iDur    = wbsHeaderIdx[WBS_COL_DUR];
  const iSist   = wbsHeaderIdx[WBS_COL_SIST];

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
      RegraDet: '' // preenchido depois a partir da aba WBS_DET (US-FUNCIONALIDADE)
    };
  });

  // ===== 1b) WBS (detalhamento) → complementar RegraDet por ID_US, se existir =====
  const detSheet = ss.getSheetByName(SHEET_WBS_DET);
  if (detSheet) {
    const detData = detSheet.getDataRange().getValues();
    if (detData.length > 1) {
      const detHeaderIdx = indexByName_(detData[0]);
      const iDetIdUs  = detHeaderIdx[WBS_DET_COL_ID_HIST];
      const iDetPri   = detHeaderIdx[WBS_COL_REGRA_DET_PRI];
      const iDetAlt   = detHeaderIdx[WBS_COL_REGRA_DET_ALT];
      const iDetTexto = iDetPri !== undefined ? iDetPri : iDetAlt;
      if (iDetIdUs !== undefined && iDetTexto !== undefined) {
        detData.slice(1).forEach(r => {
          const idUs = safeTrim_(r[iDetIdUs]);
          if (!idUs) return;
          if (!mapaWbs[idUs]) mapaWbs[idUs] = {};
          // não sobrescreve se já tiver
          if (!mapaWbs[idUs].RegraDet) {
            mapaWbs[idUs].RegraDet = r[iDetTexto] || '';
          }
        });
      }
    }
  }

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
    const func  = row[iFunc]; // funcionalidade baseline (executiva)
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
    'Funcionalidade',           // texto completo (WBS)
    'Funcionalidade_Baseline',  // texto baseline (validação)
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

    let etapa = v.Etapa || '';
    let proc  = v.Processo || '';
    let regra = v.Regra || '';
    const funcDet  = w.RegraDet || w.FuncOrig || ''; // texto completo da WBS
    const regraDet = w.RegraDet || '';               // coluna específica de detalhamento da WBS

    // Funcionalidade_Baseline: começa pela coluna de validação,
    // mas se o detalhamento da WBS for maior, usa o texto mais completo.
    let funcBase = v.Func || w.FuncOrig || '';
    if (regraDet && regraDet.length > funcBase.length) {
      funcBase = regraDet;
    }
    const wbs   = w.WBS || '';
    const dur   = w.Duracao || '';
    const sist  = w.Sistemas || '';

    // Marcar US não mapeadas na validação de forma explícita
    if (!v.Etapa || !v.Processo || !v.Regra) {
      if (!etapa) etapa = '[NÃO MAPEADO NA VALIDAÇÃO]';
      if (!proc)  proc  = '[NÃO MAPEADO NA VALIDAÇÃO]';
      if (!regra) regra = '[NÃO MAPEADO NA VALIDAÇÃO]';
    }

    let statusProj = '';
    let pctReal    = 0;
    let dtFimProj  = '';

  const cob = (v.Cobertura || '').toUpperCase();
  if (cob.includes('COBERTO') && !cob.includes('SEM')) {
      statusProj = 'Implementado';
      pctReal    = 100;                       // 100%
      dtFimProj  = new Date(2026, 2, 30);     // 30/03/2026 (data de corte)
    } else if (cob.includes('PARCIAL')) {
      statusProj = 'Em andamento';
      pctReal    = 50;                        // 50% (pode ser ajustado depois)
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
      idUs,        // ID_US
      etapa,       // Etapa
      proc,        // Processo
      regra,       // Regra
      funcDet,     // Funcionalidade (texto completo WBS)
      funcBase,    // Funcionalidade_Baseline
      regraDet,    // Regra_Detalhada_WBS
      wbs,         // WBS
      dur,         // Duracao_Dias
      '',          // Produto
      sist,        // Sistemas_Legados
      qtdApis,     // Qtd_APIs
      0,           // Qtd_Times_Envolvidos
      0,           // Tem_Interseccao
      statusProj,  // Status_Projeto
      pctReal,     // Pct_Realizado
      '',          // Data_Inicio_Projeto
      dtFimProj,   // Data_Fim_Projeto
      '',          // Data_Fim_Planejada
      '',          // Jira_Key
      '',          // Jira_Link
      ''           // Responsavel
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

  // Garantir que textos longos não sejam "cortados" visualmente:
  // aplicamos quebra de linha em todas as células da aba
  // e ajustamos a largura das colunas ao conteúdo.
  const lastRow = sheet.getLastRow() || 1;
  const lastCol = header.length;
  const fullRange = sheet.getRange(1, 1, lastRow, lastCol);
  fullRange.setWrap(true);
  try {
    sheet.autoResizeColumns(1, lastCol);
  } catch (e) {
    // Em alguns ambientes (ou se muitas colunas), autoResize pode falhar;
    // nesse caso, apenas ignoramos o erro para não quebrar a execução.
  }
}