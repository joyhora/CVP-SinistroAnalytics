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
  montarLkApiXRegra_(ss, mapaUsFinal, mapaVal);
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
 *  - mapaVal: contém para cada ID_US a string de APIs envolvidas (Validação Cruzada EF×WBS)
 *  - mapaUsFinal: resumo das informações de cada US já consolidadas em LK_US_BASE
 *
 * Ou seja, usamos diretamente a mesma origem de APIs que já alimenta a LK_US_BASE,
 * sem depender do layout da aba "Validação Anexos×WBS×APIs".
 */
function montarLkApiXRegra_(ss, mapaUsFinal, mapaVal) {
  // Opcional: mapear status da API a partir do catálogo detalhado
  const mapaStatusApi = {};
  const catSheet = ss.getSheetByName(SHEET_CAT_APIS);
  if (catSheet) {
    const catData = catSheet.getDataRange().getValues();
    if (catData.length > 1) {
      const catHeaderIdx = indexByName_(catData[0]);
      // Na sua aba o código da API está em "API ID" e o status/fase
      // em uma coluna que contém "Fase". Usamos busca por tokens.
      const iCatCod  = findColByTokens_(catHeaderIdx, ['api','id']) ??
                       findColByTokens_(catHeaderIdx, ['api']);
      const iCatStat = findColByTokens_(catHeaderIdx, ['fase']) ??
                       findColByTokens_(catHeaderIdx, ['status']);
      if (iCatCod !== undefined && iCatStat !== undefined) {
        catData.slice(1).forEach(r => {
          const cod = safeTrim_(r[iCatCod]);
          const st  = safeTrim_(r[iCatStat]);
          if (!cod) return;
          const stNorm = st.toLowerCase();
          // Regra de negócio: se contiver "fase 1" considerar "Homologação",
          // caso contrário "Não iniciado".
          let statusApi = 'Não iniciado';
          if (stNorm.indexOf('fase 1') !== -1) {
            statusApi = 'Homologação';
          }
          mapaStatusApi[cod] = statusApi;
        });
      }
    }
  }

  const linhas = [];

  // Para cada US consolidada, olhamos a lista de APIs vinda de mapaVal[idUs].Apis
  Object.keys(mapaUsFinal).forEach(idUs => {
    const infoUs = mapaUsFinal[idUs];
    const v      = mapaVal[idUs] || {};
    const apisStr = safeTrim_(v.Apis || '');
    if (!apisStr) return;

    const apis = apisStr.split(/[,\n;]+/).map(s => safeTrim_(s)).filter(Boolean);
    apis.forEach(apiCod => {
      if (!apiCod) return;
      const statusApi = mapaStatusApi[apiCod] || 'Não iniciado';
      linhas.push([
        apiCod,
        infoUs.IdOriginal || idUs,
        infoUs.Etapa || '',
        infoUs.Processo || '',
        infoUs.Regra || '',
        infoUs.RegraDet || infoUs.FuncDet || '',
        infoUs.Status || '',
        infoUs.PctReal || 0,
        statusApi
      ]);
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
    'Pct_Realizado',
    'Status_API'
  ];

  escreverAba_(ss, SHEET_LK_API_X_REGRA, header, linhas);
}

/*************** DRAFT CRONOGRAMA ***************/

/** Aba existente na planilha (nome com espaço). Fallback: Draft-Cronograma */
const SHEET_DRAFT_CRONOGRAMA_PRI = 'Draft Cronograma';
const SHEET_DRAFT_CRONOGRAMA_ALT = 'Draft-Cronograma';
/** Data limite do cronograma (31/10/2026) */
const CRONO_DATA_FIM = new Date(2026, 9, 31);
/** Início do projeto (01/04/2026) — piso se não houver datas de planejamento na planilha */
const CRONO_DATA_INICIO_PROJETO = new Date(2026, 3, 1);
/** Total de horas do escopo a distribuir entre as linhas classificadas */
const CRONO_HORAS_TOTAL_PROJETO = 6160;
/**
 * Pesos relativos por tipo de atividade (repartição das 6.160 h).
 * Entrega pelo backend = 0 (sem carga alocada neste critério).
 */
const CRONO_PESO_ANALISE = 8;
const CRONO_PESO_ENTREGA = 0;
const CRONO_PESO_DEV = 60;
const CRONO_PESO_TESTE = 8;
const CRONO_PESO_HOMOLOG = 16;

/**
 * Localiza a aba de cronograma já existente (não cria aba nova).
 */
function obterAbaDraftCronograma_(ss) {
  let sh = ss.getSheetByName(SHEET_DRAFT_CRONOGRAMA_PRI);
  if (!sh) sh = ss.getSheetByName(SHEET_DRAFT_CRONOGRAMA_ALT);
  return sh;
}

/**
 * Classifica as 5 atividades do fluxo (normalizado sem acento).
 * Retorna: ANALISE | ENTREGA | DEV | TESTE | HOMOLOG | null
 *
 * Etapas esperadas no cronograma (WBS): Receber Demanda, Gerenciador de Documentos,
 * Analisar, Subsidiar Análise, Concluir Sinistro, Alçada Superior, Regras Gerais —
 * cada uma com as mesmas atividades-tipo abaixo, na ordem das linhas da planilha.
 */
function classificarAtividadeCronograma_(nomeTarefa) {
  const n = safeTrim_(nomeTarefa)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
  if (!n) return null;
  if (/homologa/.test(n)) return 'HOMOLOG';
  if (/teste.*conjunto|testes.*conjunto/.test(n)) return 'TESTE';
  if (/desenvolvimento/.test(n) && /salesforce|api|dev/.test(n)) return 'DEV';
  if (/desenvolvimento.*salesforce/.test(n)) return 'DEV';
  if (/entrega/.test(n) && (/backend|api/.test(n) || /pelo backend/.test(n))) return 'ENTREGA';
  if (/entrega.*api|api.*entrega/.test(n)) return 'ENTREGA';
  // "analise" como palavra (evita confundir o título da etapa ".3 Analisar" com análise de doc)
  if (/\banalise\b/.test(n) && /document|documento|api/.test(n)) return 'ANALISE';
  if (/\banalise\b.*document/.test(n)) return 'ANALISE';
  return null;
}

function obterPesoTipoCronograma_(tipo) {
  if (tipo === 'ANALISE') return CRONO_PESO_ANALISE;
  if (tipo === 'ENTREGA') return CRONO_PESO_ENTREGA;
  if (tipo === 'DEV') return CRONO_PESO_DEV;
  if (tipo === 'TESTE') return CRONO_PESO_TESTE;
  if (tipo === 'HOMOLOG') return CRONO_PESO_HOMOLOG;
  return 0;
}

/** Uma linha curta para o racional (tipo de trabalho). */
function rotuloTipoCronograma_(tipo) {
  if (tipo === 'ANALISE') return 'Análise de documentação das APIs.';
  if (tipo === 'ENTREGA') return 'Entrega das APIs pelo backend.';
  if (tipo === 'DEV') return 'Desenvolvimento Salesforce (APIs).';
  if (tipo === 'TESTE') return 'Testes em conjunto com o backend (APIs).';
  if (tipo === 'HOMOLOG') return 'Homologação com o negócio.';
  return '';
}

/** WBS no início da célula: .3, 3., 3 - … → 1..7 */
function extrairNumeroEtapaCronograma_(nomeTarefa) {
  const n = safeTrim_(nomeTarefa)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
  if (!n) return null;
  const m = n.match(/^\s*\.?\s*([1-7])[\.\)\s:-]/);
  if (m) return parseInt(m[1], 10);
  const m2 = n.match(/^\s*([1-7])\s*[\.\)]\s/);
  if (m2) return parseInt(m2[1], 10);
  return null;
}

/** Etapa pelo nome (linhas só de título, sem atividade classificada). */
function detectarEtapaPorPalavra_(nomeTarefa) {
  const n = safeTrim_(nomeTarefa)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
  if (!n) return null;
  if (/regras\s+gerais/.test(n)) return 7;
  if (/alcada\s+superior|aprovar\s+processo/.test(n)) return 6;
  if (/concluir\s+sinistro/.test(n)) return 5;
  if (/subsidiar\s+analise|subsidiar/.test(n)) return 4;
  if (/gerenciador(\s+de)?\s+documentos/.test(n)) return 2;
  if (/receber\s+demanda|demanda\s+de\s+sinistro/.test(n)) return 1;
  if (/\banalisar\b/.test(n) && !/\banalise\b/.test(n) && !/document/.test(n)) return 3;
  return null;
}

/** Linha de título de etapa (não é uma das cinco atividades). */
function ehTituloEtapaSemAtividade_(nomeTarefa) {
  if (classificarAtividadeCronograma_(nomeTarefa)) return false;
  return extrairNumeroEtapaCronograma_(nomeTarefa) !== null || detectarEtapaPorPalavra_(nomeTarefa) !== null;
}

/** Etapa da linha: número WBS, palavras-chave da etapa ou último título visto. */
function resolverEtapaLinhaCronograma_(nomeTarefa, etapaCursor) {
  const ex = extrairNumeroEtapaCronograma_(nomeTarefa);
  if (ex !== null) return ex;
  const pal = detectarEtapaPorPalavra_(nomeTarefa);
  if (pal !== null) return pal;
  if (etapaCursor > 0) return etapaCursor;
  return 1;
}

/**
 * Reparte um inteiro (ex.: 6.160 h) proporcionalmente a pesos, soma exata = alvo.
 * Pesos com valor 0 recebem 0 (ex.: entrega de APIs com peso 0 no critério).
 */
function distribuirInteirosProporcional_(pesos, alvo) {
  const n = pesos.length;
  if (n === 0) return [];
  const totalP = pesos.reduce((a, b) => a + b, 0);
  if (totalP <= 0) {
    return pesos.map(() => 0);
  }
  const raw = pesos.map(p => (p / totalP) * alvo);
  const floors = raw.map(v => Math.floor(v));
  let soma = floors.reduce((a, b) => a + b, 0);
  let rem = alvo - soma;
  const ordem = raw
    .map((v, i) => ({ i, r: v - Math.floor(v) }))
    .sort((a, b) => b.r - a.r);
  const out = floors.slice();
  for (let k = 0; k < rem; k++) out[ordem[k].i]++;
  return out;
}

/**
 * Detecta índices de colunas a partir da linha de cabeçalho (Task, Datas, Horas, API, Observação).
 */
function detectarColunasCronograma_(headerRow) {
  const raw = headerRow.map(c => safeTrim_(String(c)));
  const h = raw.map(s => s.toLowerCase());
  const norm = s =>
    s
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');

  let colTask = -1;
  for (let i = 0; i < h.length; i++) {
    const t = norm(raw[i]);
    if (t.indexOf('task') !== -1 || (t.indexOf('nome') !== -1 && t.indexOf('recurso') === -1 && t.indexOf('area') === -1))
      colTask = i;
  }
  if (colTask < 0) colTask = 0;

  const idx = pred => {
    for (let i = 0; i < h.length; i++) if (pred(norm(raw[i]), i)) return i;
    return -1;
  };

  const colDur = idx((t, i) => /dura|duracao/.test(t));
  const colIni = idx((t, i) => /inicio|início/.test(t) && !/planej/.test(t));
  const colFim = idx((t, i) => /termino|término|fim/.test(t));
  const colPct = idx((t, i) => /%|conclu/.test(t));
  // Horas: quantidade / estimada / horas (evita coluna "nomes" duplicada)
  let colHor = idx((t, i) => /hora|quantidade|estim/.test(t) && !/recurso/.test(t));
  if (colHor < 0) colHor = 6;
  // Coluna só "API"
  let colApi = idx((t, i) => t === 'api' || (t.indexOf('api') !== -1 && t.indexOf('observa') === -1 && t.indexOf('envolv') === -1));
  const colObs = idx((t, i) => /observa|observacao|observação|racional/.test(t));

  if (colDur < 0) colDur = 1;
  if (colIni < 0) colIni = 2;
  if (colFim < 0) colFim = 3;
  if (colPct < 0) colPct = 4;
  if (colHor < 0) colHor = 6;
  // Se não achou coluna API separada, APIs vêm do texto da observação
  if (colApi < 0) colApi = -1;

  return {
    colTask,
    colDur,
    colIni,
    colFim,
    colPct,
    colHor,
    colApi,
    colObs: colObs >= 0 ? colObs : 7
  };
}

/**
 * Primeira linha com uma das 5 atividades do fluxo; linhas acima = planejamento.
 * Usa a maior data em "Término" antes dessa linha + 1 dia útil, com piso em 01/04/2026.
 */
function inferirDataInicioFaseDev_(data, cols) {
  const colTask = cols.colTask;
  const colFim = cols.colFim;
  let firstDevIdx = -1;
  for (let r = 1; r < data.length; r++) {
    const nome = safeTrim_(data[r][colTask]);
    if (classificarAtividadeCronograma_(nome)) {
      firstDevIdx = r;
      break;
    }
  }
  if (firstDevIdx < 0) return new Date(CRONO_DATA_INICIO_PROJETO.getTime());

  let maxD = null;
  for (let r = 1; r < firstDevIdx; r++) {
    const v = data[r][colFim];
    if (v instanceof Date && !isNaN(v.getTime())) {
      if (!maxD || v.getTime() > maxD.getTime()) maxD = v;
    }
  }
  if (!maxD) return new Date(CRONO_DATA_INICIO_PROJETO.getTime());

  const next = addBusinessDays_(maxD, 1);
  const padrao = new Date(CRONO_DATA_INICIO_PROJETO.getTime());
  return next.getTime() > padrao.getTime() ? next : padrao;
}

/** Quantidade de dias úteis entre duas datas (inclusive). */
function contarDiasUteisInclusive_(inicio, fim) {
  if (inicio > fim) return 0;
  let n = 0;
  const cur = new Date(inicio.getTime());
  cur.setHours(12, 0, 0, 0);
  const lim = new Date(fim.getTime());
  lim.setHours(12, 0, 0, 0);
  while (cur <= lim) {
    const wd = cur.getDay();
    if (wd !== 0 && wd !== 6) n++;
    cur.setDate(cur.getDate() + 1);
  }
  return n;
}

/**
 * Distribui dias úteis entre atividades proporcionalmente ao esforço (horas),
 * para preencher todo o período [dataInicio, CRONO_DATA_FIM] com visão de evolução (ordem das linhas).
 * Usa maior resto (Hamilton) para a soma bater exatamente com diasUteisTotais.
 */
function distribuirDiasUteisProporcional_(itens, diasUteisTotais) {
  const n = itens.length;
  if (n === 0) return [];
  if (diasUteisTotais <= 0) return itens.map(() => 0);

  const totalHoras = itens.reduce((s, x) => s + x.totalHoras, 0);
  if (totalHoras <= 0) {
    const base = Math.floor(diasUteisTotais / n);
    let rem = diasUteisTotais - base * n;
    return itens.map((_, i) => base + (i < rem ? 1 : 0));
  }

  const quotas = itens.map(x => (x.totalHoras / totalHoras) * diasUteisTotais);
  const floors = quotas.map(q => Math.floor(q));
  let soma = floors.reduce((a, b) => a + b, 0);
  let rem = diasUteisTotais - soma;
  const ordem = quotas
    .map((q, i) => ({ i, r: q - Math.floor(q) }))
    .sort((a, b) => b.r - a.r);
  const dias = floors.slice();
  for (let k = 0; k < rem; k++) dias[ordem[k].i]++;
  return dias;
}

/**
 * Gera/atualiza a aba **Draft Cronograma**: horas, datas e observação.
 * Escopo: **6.160 h** entre **01/04/2026** e **31/10/2026**; atividades com peso por tipo.
 * Linhas só de etapa (.1 … .7) recebem **soma de horas** e **início/fim** = união temporal das atividades da etapa.
 */
function gerarDraftCronograma() {
  const ss = SpreadsheetApp.getActive();
  const sh = obterAbaDraftCronograma_(ss);
  if (!sh) {
    const msg =
      'Não encontrei a aba "' +
      SHEET_DRAFT_CRONOGRAMA_PRI +
      '" nem "' +
      SHEET_DRAFT_CRONOGRAMA_ALT +
      '". Renomeie sua aba para um desses nomes ou crie a aba manualmente.';
    try {
      SpreadsheetApp.getUi().alert(msg);
    } catch (e) {
      throw new Error(msg);
    }
    return;
  }

  const data = sh.getDataRange().getValues();
  if (data.length < 2) {
    try {
      SpreadsheetApp.getUi().alert('Aba de cronograma sem linhas de dados além do cabeçalho.');
    } catch (e) {
      /* headless */
    }
    return;
  }

  const cols = detectarColunasCronograma_(data[0]);
  const colTask = cols.colTask;
  const colDur = cols.colDur;
  const colIni = cols.colIni;
  const colFim = cols.colFim;
  const colHor = cols.colHor;
  const colObs = cols.colObs;
  const colApi = cols.colApi;

  const dataInicioProjeto = inferirDataInicioFaseDev_(data, cols);
  const diasUteisDisponiveis = contarDiasUteisInclusive_(dataInicioProjeto, CRONO_DATA_FIM);

  if (diasUteisDisponiveis < 1) {
    try {
      SpreadsheetApp.getUi().alert('A data de início da fase dev está após 31/10/2026. Ajuste datas de planejamento na planilha.');
    } catch (e) {
      throw new Error('Janela de cronograma inválida.');
    }
    return;
  }

  const apiRegex = /\b(API-\d+|BK\d+)\b/gi;
  const itens = [];
  let etapaCursor = 0;

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const nomeTarefa = safeTrim_(row[colTask]);
    if (!nomeTarefa) continue;

    if (ehTituloEtapaSemAtividade_(nomeTarefa)) {
      const ex = extrairNumeroEtapaCronograma_(nomeTarefa);
      if (ex !== null) etapaCursor = ex;
      else {
        const p = detectarEtapaPorPalavra_(nomeTarefa);
        if (p !== null) etapaCursor = p;
      }
      continue;
    }

    const tipo = classificarAtividadeCronograma_(nomeTarefa);
    if (!tipo) continue;

    const etapa = resolverEtapaLinhaCronograma_(nomeTarefa, etapaCursor);
    if (extrairNumeroEtapaCronograma_(nomeTarefa) !== null) {
      etapaCursor = extrairNumeroEtapaCronograma_(nomeTarefa);
    }

    const textoApiCol = colApi >= 0 ? safeTrim_(row[colApi] || '') : '';
    const obsExistente = safeTrim_(row[colObs] || '');
    const sFull = textoApiCol + ' ' + obsExistente + ' ' + nomeTarefa;

    const apis = [];
    let m;
    apiRegex.lastIndex = 0;
    while ((m = apiRegex.exec(sFull)) !== null) {
      const cod = m[1].toUpperCase();
      if (apis.indexOf(cod) === -1) apis.push(cod);
    }

    const peso = obterPesoTipoCronograma_(tipo);
    itens.push({
      rowIndex: r,
      etapa: etapa,
      tipo: tipo,
      peso: peso,
      rotuloTipo: rotuloTipoCronograma_(tipo),
      apis: apis,
      obsExistente: obsExistente
    });
  }

  if (itens.length === 0) {
    try {
      SpreadsheetApp.getUi().alert(
        'Nenhuma linha reconhecida. Use os nomes das cinco atividades (análise documentação, entrega APIs pelo backend, desenvolvimento Salesforce, testes com backend, homologação).'
      );
    } catch (e) {
      /* headless */
    }
    return;
  }

  const somaPesos = itens.reduce((s, x) => s + x.peso, 0);
  if (somaPesos <= 0) {
    try {
      SpreadsheetApp.getUi().alert(
        'A soma dos pesos das linhas é zero (só há tarefas “entrega de APIs pelo backend”, peso 0 neste critério). Inclua linhas com análise, desenvolvimento, testes ou homologação.'
      );
    } catch (e) {
      throw new Error('Soma de pesos zero.');
    }
    return;
  }

  const horasArr = distribuirInteirosProporcional_(
    itens.map(x => x.peso),
    CRONO_HORAS_TOTAL_PROJETO
  );
  for (let i = 0; i < itens.length; i++) itens[i].totalHoras = horasArr[i];

  const diasPorItem = distribuirDiasUteisProporcional_(itens, diasUteisDisponiveis);
  let cursor = new Date(dataInicioProjeto.getTime());
  const agregadoEtapa = {};

  const fmtPeriodoInicio =
    ('0' + dataInicioProjeto.getDate()).slice(-2) +
    '/' +
    ('0' + (dataInicioProjeto.getMonth() + 1)).slice(-2) +
    '/' +
    dataInicioProjeto.getFullYear();
  const fmtFimProjeto =
    ('0' + CRONO_DATA_FIM.getDate()).slice(-2) +
    '/' +
    ('0' + (CRONO_DATA_FIM.getMonth() + 1)).slice(-2) +
    '/' +
    CRONO_DATA_FIM.getFullYear();

  for (let k = 0; k < itens.length; k++) {
    const item = itens[k];
    const dUteis = diasPorItem[k];
    const r = item.rowIndex;
    const inicio = new Date(cursor.getTime());
    let termino;
    if (dUteis <= 0) {
      termino = new Date(inicio.getTime());
    } else {
      termino = addBusinessDays_(inicio, dUteis - 1);
      if (termino.getTime() > CRONO_DATA_FIM.getTime()) {
        termino = new Date(CRONO_DATA_FIM.getTime());
      }
    }

    const durCalendario = diasEntreDatas_(inicio, termino) + 1;

    const blocoRacional =
      '\n\n--- Racional (automático) ---\n' +
      'Janela ' +
      fmtPeriodoInicio +
      '–' +
      fmtFimProjeto +
      '; ' +
      CRONO_HORAS_TOTAL_PROJETO +
      ' h no escopo. Esta linha: ' +
      item.totalHoras +
      ' h — ' +
      item.rotuloTipo.replace(/\.$/, '') +
      '. ' +
      'As horas foram alocadas pelo tipo de atividade em relação às outras linhas; o cronograma segue a ordem da planilha ao longo dos ' +
      diasUteisDisponiveis +
      ' dias úteis. ' +
      (item.apis.length ? 'APIs: ' + item.apis.join(', ') + '.' : '');

    const textoObs = item.obsExistente ? item.obsExistente + blocoRacional : safeTrim_(blocoRacional.replace(/^\s+/, ''));

    item.inicio = inicio;
    item.termino = termino;

    const ep = item.etapa;
    if (!agregadoEtapa[ep]) {
      agregadoEtapa[ep] = { horas: 0, ini: null, fim: null };
    }
    agregadoEtapa[ep].horas += item.totalHoras;
    if (!agregadoEtapa[ep].ini || inicio.getTime() < agregadoEtapa[ep].ini.getTime()) {
      agregadoEtapa[ep].ini = new Date(inicio.getTime());
    }
    if (!agregadoEtapa[ep].fim || termino.getTime() > agregadoEtapa[ep].fim.getTime()) {
      agregadoEtapa[ep].fim = new Date(termino.getTime());
    }

    sh.getRange(r + 1, colHor + 1).setValue(item.totalHoras);
    sh.getRange(r + 1, colIni + 1).setValue(inicio);
    sh.getRange(r + 1, colFim + 1).setValue(termino);
    sh.getRange(r + 1, colDur + 1).setValue(durCalendario);
    sh.getRange(r + 1, colObs + 1).setValue(textoObs);

    if (dUteis > 0) {
      cursor = addBusinessDays_(termino, 1);
      if (cursor.getTime() > CRONO_DATA_FIM.getTime()) {
        cursor = new Date(CRONO_DATA_FIM.getTime());
      }
    }
  }

  const lastRow = sh.getLastRow();
  sh.getRange(2, colIni + 1, lastRow, colFim + 1).setNumberFormat('dd/mm/yyyy');

  for (let r = 1; r < data.length; r++) {
    const nomeTitulo = safeTrim_(data[r][colTask]);
    if (!nomeTitulo || !ehTituloEtapaSemAtividade_(nomeTitulo)) continue;
    let epTit = extrairNumeroEtapaCronograma_(nomeTitulo);
    if (epTit === null) epTit = detectarEtapaPorPalavra_(nomeTitulo);
    if (epTit === null || !agregadoEtapa[epTit]) continue;
    const agg = agregadoEtapa[epTit];
    sh.getRange(r + 1, colHor + 1).setValue(agg.horas);
    sh.getRange(r + 1, colIni + 1).setValue(agg.ini);
    sh.getRange(r + 1, colFim + 1).setValue(agg.fim);
    sh.getRange(r + 1, colDur + 1).setValue(diasEntreDatas_(agg.ini, agg.fim) + 1);
    const obsTit = safeTrim_(String(data[r][colObs] || ''));
    const blocoEtapa =
      '\n\n--- Agrupamento etapa ---\n' +
      'Horas e datas da linha = soma das atividades e do primeiro ao último dia útil desta etapa.';
    sh.getRange(r + 1, colObs + 1).setValue(obsTit ? obsTit + blocoEtapa : safeTrim_(blocoEtapa.replace(/^\s+/, '')));
  }
}

/** Soma dias úteis (seg–sex) a partir de start, incluindo start como dia 0 quando n=0 */
function addBusinessDays_(start, n) {
  const d = new Date(start.getTime());
  let left = n;
  while (left > 0) {
    d.setDate(d.getDate() + 1);
    const wd = d.getDay();
    if (wd !== 0 && wd !== 6) left--;
  }
  return d;
}

function subtractBusinessDays_(end, n) {
  const d = new Date(end.getTime());
  let left = n;
  while (left > 0) {
    d.setDate(d.getDate() - 1);
    const wd = d.getDay();
    if (wd !== 0 && wd !== 6) left--;
  }
  return d;
}

function diasEntreDatas_(a, b) {
  const ms = 24 * 60 * 60 * 1000;
  return Math.round((b.getTime() - a.getTime()) / ms);
}