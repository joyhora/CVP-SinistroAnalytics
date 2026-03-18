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
const SHEET_WBS_PROJECT   = 'WBS Project';                // estrutura WBS + US Original
const SHEET_WBS_DET       = 'WBS';                        // detalhamento por US (US-FUNCIONALIDADE)
const SHEET_VALIDACAO     = 'Validação Cruzada EF×WBS';   // visão executiva / baseline
const SHEET_VALIDACAO_API = 'Validação Anexos×WBS×APIs';  // mapeamento US x APIs
const SHEET_CAT_APIS      = 'Catálogo APIs Detalhado';    // dicionário de APIs (opcional)

// Abas de destino
const SHEET_LK_US_BASE        = 'LK_US_BASE';
const SHEET_LK_US_NAO_MAPPEDS = 'LK_US_NAO_MAPEADAS';
const SHEET_LK_API_X_REGRA    = 'LK_API_X_REGRA';

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

// Colunas adicionais na aba WBS para quando não existir validação:
// usamos Etapa / Processo diretamente da WBS.
const WBS_DET_COL_ETAPA   = 'ETAPA';
const WBS_DET_COL_PROCESSO = 'PROCESSO';

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
    const idUsOrig = safeTrim_(row[iIdUs]);
    // Em algumas planilhas há uma linha final "IDHISTORIA" que não é uma US válida.
    if (idUsOrig.toUpperCase() === 'IDHISTORIA') return;
    const idUs     = normalizeId_(idUsOrig);
    if (!idUs) return;
    mapaWbs[idUs] = {
      IdOriginal: idUsOrig,
      WBS:      iWbs  !== undefined ? row[iWbs]  : '',
      Duracao:  iDur  !== undefined ? row[iDur]  : '',
      Sistemas: iSist !== undefined ? row[iSist] : '',
      FuncOrig: iTask !== undefined ? row[iTask] : '',
      RegraDet: '', // preenchido depois a partir da aba WBS_DET (US-FUNCIONALIDADE)
      EtapaWbs: '',
      ProcessoWbs: ''
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
      const iDetEtapa   = detHeaderIdx[WBS_DET_COL_ETAPA];
      const iDetProc    = detHeaderIdx[WBS_DET_COL_PROCESSO];
      if (iDetIdUs !== undefined && iDetTexto !== undefined) {
        detData.slice(1).forEach(r => {
          const rawId = safeTrim_(r[iDetIdUs]);
          if (rawId.toUpperCase() === 'IDHISTORIA') return;
          const idUs = normalizeId_(rawId);
          if (!idUs) return;
          if (!mapaWbs[idUs]) mapaWbs[idUs] = {};
          // não sobrescreve se já tiver
          if (!mapaWbs[idUs].RegraDet) {
            mapaWbs[idUs].RegraDet = r[iDetTexto] || '';
          }
          if (iDetEtapa !== undefined && !mapaWbs[idUs].EtapaWbs) {
            mapaWbs[idUs].EtapaWbs = safeTrim_(r[iDetEtapa]);
          }
          if (iDetProc !== undefined && !mapaWbs[idUs].ProcessoWbs) {
            mapaWbs[idUs].ProcessoWbs = safeTrim_(r[iDetProc]);
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
    const etapa = safeTrim_(row[iEtapa]);
    const proc  = safeTrim_(row[iProc]);
    const func  = safeTrim_(row[iFunc]); // funcionalidade baseline (executiva)
    const regra = iRegra !== undefined ? safeTrim_(row[iRegra]) : '';
    const cob   = safeTrim_(row[iCob]); // COBERTO / PARCIAL / X% SEM COBERTURA
    const apis  = iApis !== undefined ? safeTrim_(row[iApis]) : '';

    const idsStr = safeTrim_(row[iIds]);
    if (!idsStr) return;

    // IDs podem vir separados por vírgula, ponto e vírgula ou quebras de linha.
    idsStr.split(/[,\n;]+/)
      .map(s => safeTrim_(s))
      .filter(Boolean)
      .forEach(idUsRaw => {
        const idUs = normalizeId_(idUsRaw);
        if (!idUs) return;
        const atual = mapaVal[idUs];
        // Sempre garantimos que exista um registro...
        if (!atual) {
          mapaVal[idUs] = {
            Etapa: etapa,
            Processo: proc,
            Func: func,
            Regra: regra,
            Cobertura: cob,
            Apis: apis
          };
        } else {
          // ...mas, se essa nova linha tiver informações mais completas
          // (Etapa/Processo/Regra preenchidos), sobrescrevemos.
          const temInfoNova =
            (!!etapa && !atual.Etapa) ||
            (!!proc  && !atual.Processo) ||
            (!!regra && !atual.Regra) ||
            (!!func  && !atual.Func);
          if (temInfoNova) {
            mapaVal[idUs] = {
              Etapa: etapa || atual.Etapa,
              Processo: proc || atual.Processo,
              Func: func || atual.Func,
              Regra: regra || atual.Regra,
              Cobertura: cob || atual.Cobertura,
              Apis: apis || atual.Apis
            };
          }
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
  const mapaUsFinal = {};
  const idsNaoMapeados = [];

  Object.keys(mapaWbs).forEach(idUs => {
    const w = mapaWbs[idUs] || {};

    // 1ª tentativa: mapeamento direto pelo ID normalizado
    let v = mapaVal[idUs] || {};
    // 2ª tentativa (fallback): procurar por um "ID base" equivalente
    // Ex.: usar configuração de SF1-SIS-US06C também para SF1-SIS-US06.
    if (!v || (!v.Etapa && !v.Processo && !v.Regra)) {
      const vAlt = findFromBaseId_(idUs, mapaVal);
      if (vAlt) v = vAlt;
    }
    // 3ª tentativa (definitiva): busca "bruta" na aba de Validação
    // por qualquer célula de IDs que contenha o ID original.
    if (!v || (!v.Etapa && !v.Processo && !v.Regra)) {
      const rawId = w.IdOriginal || idUs;
      const vBruto = findFromValidationRawId_(rawId, valRows, iIds, iEtapa, iProc, iFunc, iRegra, iCob, iApis);
      if (vBruto) v = vBruto;
    }

    if (!v || !v.Etapa || !v.Processo || !v.Regra) {
      idsNaoMapeados.push([
        w.IdOriginal || idUs,
        w.WBS || '',
        w.FuncOrig || ''
      ]);
    }

    // Etapa/Processo: prioriza validação; se não existir validação,
    // preenche a partir da própria WBS.
    let etapa = v.Etapa || w.EtapaWbs || '';
    let proc  = v.Processo || w.ProcessoWbs || '';
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

    // Se não houver mapeamento na validação, mantemos Etapa/Processo/Regra em branco
    // (sem preencher com o texto "[NÃO MAPEADO NA VALIDAÇÃO]"), para evitar esse rótulo na LK.

    let statusProj = '';
    let pctReal    = 0;
    let dtFimProj  = '';

    const cob = (v.Cobertura || '').toUpperCase();
    if (cob.includes('COBERTO') && !cob.includes('SEM')) {
      statusProj = 'Implementado';
      pctReal    = 1;                         // 100%
      dtFimProj  = new Date(2026, 2, 30);     // 30/03/2026 (data de corte)
    } else if (cob.includes('PARCIAL')) {
      statusProj = 'Em andamento';
      pctReal    = 0.5;                       // 50%
    } else {
      statusProj = 'Não iniciado';
      pctReal    = 0;                         // 0%
    }

    let qtdApis = 0;
    if (v.Apis) {
      qtdApis = v.Apis.split(',')
        .map(s => safeTrim_(s))
        .filter(Boolean).length;
    }

    mapaUsFinal[idUs] = {
      IdOriginal: w.IdOriginal || idUs,
      Etapa: etapa,
      Processo: proc,
      Regra: regra,
      FuncDet: funcDet,
      RegraDet: regraDet,
      Status: statusProj,
      PctReal: pctReal
    };

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

  // Aba auxiliar para controle de qualidade: quais US da WBS não foram
  // mapeadas na aba de Validação (impacta diretamente os indicadores).
  // Isso permite corrigir a planilha de origem e voltar a rodar o script.
  escreverAba_(
    ss,
    SHEET_LK_US_NAO_MAPPEDS,
    ['ID_US', 'WBS', 'Funcionalidade_WBS'],
    idsNaoMapeados
  );

  // ===== 4) Montar visão analítica de APIs x Regras =====
  montarLkApiXRegra_(ss, mapaUsFinal);
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
  if (v == null) return '';
  // Converte para string, troca quebras de linha por espaço e
  // compacta espaços múltiplos em um só.
  return String(v)
    .replace(/[\r\n]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

// Normalização única para IDs de US, garantindo correspondência exata
// entre WBS Project, WBS e Validação Cruzada (ignorando espaços e caixa).
function normalizeId_(id) {
  const raw = safeTrim_(id);
  if (!raw) return '';
  // Remove espaços e alguns sinais de pontuação de cauda que às vezes
  // aparecem junto com o ID (p.ex. "SF1-SIS-US06C;" ou com quebras de linha).
  return raw
    .toUpperCase()
    .replace(/\s+/g, '')
    .replace(/[.;,]+$/g, '');
}

// Busca um registro "equivalente" na validação para um ID de US,
// removendo um possível sufixo de letra no final (C, D, etc.).
function findFromBaseId_(idUs, mapaVal) {
  const base = idUs.replace(/[A-Z]$/g, ''); // SF1-SIS-US06C -> SF1-SIS-US06

  // Se já existir exatamente o ID, devolve direto
  if (mapaVal[idUs]) return mapaVal[idUs];

  // Procura por alguma chave cujo "base" coincida
  for (const [k, v] of Object.entries(mapaVal)) {
    const baseK = k.replace(/[A-Z]$/g, '');
    if (baseK === base) return v;
  }

  return null;
}

// Busca direta na aba de validação por um ID "como texto",
// usando o valor original da WBS (sem normalização agressiva).
function findFromValidationRawId_(rawId, valRows, iIds, iEtapa, iProc, iFunc, iRegra, iCob, iApis) {
  const alvo = safeTrim_(rawId);
  if (!alvo) return null;

  for (let r = 0; r < valRows.length; r++) {
    const row = valRows[r];
    const idsCell = safeTrim_(row[iIds]);
    if (!idsCell) continue;
    // Fazemos a busca textual: se a célula contém exatamente o ID
    // como palavra isolada ou lista, vamos usar essa linha.
    const partes = idsCell.split(/[,\n;]+/).map(s => safeTrim_(s)).filter(Boolean);
    if (partes.includes(alvo)) {
      return {
        Etapa: row[iEtapa],
        Processo: row[iProc],
        Func: row[iFunc],
        Regra: iRegra !== undefined ? row[iRegra] : '',
        Cobertura: safeTrim_(row[iCob]),
        Apis: iApis !== undefined ? safeTrim_(row[iApis]) : ''
      };
    }
  }
  return null;
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
  // Formatar coluna de percentual (Pct_Realizado) como porcentagem, se existir
  const pctColIndex = header.indexOf('Pct_Realizado');
  if (pctColIndex >= 0 && lastRow > 1) {
    const pctRange = sheet.getRange(2, pctColIndex + 1, lastRow - 1, 1);
    pctRange.setNumberFormat('0%');
  }
  try {
    sheet.autoResizeColumns(1, lastCol);
  } catch (e) {
    // Em alguns ambientes (ou se muitas colunas), autoResize pode falhar;
    // nesse caso, apenas ignoramos o erro para não quebrar a execução.
  }
}

/**
 * Cria a aba LK_API_X_REGRA com uma linha por combinação API x US/Regra.
 * Depende de:
 *  - SHEET_VALIDACAO_API: contém IDs USs mapeadas e APIs envolvidas
 *  - mapaUsFinal: resumo das informações de cada US já consolidadas em LK_US_BASE
 */
function montarLkApiXRegra_(ss, mapaUsFinal) {
  const apiSheet = ss.getSheetByName(SHEET_VALIDACAO_API);
  if (!apiSheet) {
    // Se a aba de validação de APIs não existir, não criamos a LK de APIs.
    return;
  }

  const apiData = apiSheet.getDataRange().getValues();
  if (apiData.length < 2) {
    escreverAba_(ss, SHEET_LK_API_X_REGRA, ['API_Codigo','ID_US','Etapa','Processo','Regra','Regra_Detalhada_WBS','Status_Projeto','Pct_Realizado'], []);
    return;
  }

  const headerIdx = indexByName_(apiData[0]);
  const iIds  = findColByTokens_(headerIdx, VAL_COL_IDS_US_TOKENS);
  const iApis = findColByTokens_(headerIdx, VAL_COL_APIS_TOKENS);
  if (iIds === undefined || iApis === undefined) {
    // Cabeçalhos não encontrados, não geramos a LK de APIs.
    escreverAba_(ss, SHEET_LK_API_X_REGRA, ['API_Codigo','ID_US','Etapa','Processo','Regra','Regra_Detalhada_WBS','Status_Projeto','Pct_Realizado'], []);
    return;
  }

  const linhas = [];
  apiData.slice(1).forEach(row => {
    const idsStr  = safeTrim_(row[iIds]);
    const apisStr = safeTrim_(row[iApis]);
    if (!idsStr || !apisStr) return;

    const ids  = idsStr.split(/[,\n;]+/).map(s => safeTrim_(s)).filter(Boolean);
    const apis = apisStr.split(/[,\n;]+/).map(s => safeTrim_(s)).filter(Boolean);

    ids.forEach(idRaw => {
      const idUs = normalizeId_(idRaw);
      if (!idUs) return;
      const infoUs = mapaUsFinal[idUs];
      if (!infoUs) return; // só consideramos US que entraram na LK_US_BASE

      apis.forEach(apiCod => {
        if (!apiCod) return;
        linhas.push([
          apiCod,
          infoUs.IdOriginal || idUs,
          infoUs.Etapa || '',
          infoUs.Processo || '',
          infoUs.Regra || '',
          infoUs.RegraDet || infoUs.FuncDet || '',
          infoUs.Status || '',
          infoUs.PctReal || 0
        ]);
      });
    });
  });

  const header = [
    'API_Codigo',
    'ID_US',
    'Etapa',
    'Processo',
    'Regra',
    'Regra_Detalhada_WBS',
    'Status_Projeto',
    'Pct_Realizado'
  ];

  escreverAba_(ss, SHEET_LK_API_X_REGRA, header, linhas);
}