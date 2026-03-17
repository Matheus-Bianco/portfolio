/**
 * ============================================================
 * SISTEMA DE MONITORAMENTO E GESTAO ESTRATEGICA - SED JOINVILLE
 * PEI 2025-2029
 * ============================================================
 * Arquivo principal: Roteamento e configuracao
 */

// ============ CONFIGURACAO GLOBAL ============
var CONFIG = {
  SPREADSHEET_ID: '1LJApwmT0aH53CLylt7ktP7u6b9UvVhslfasK-A5ICDI', // Preencher com o ID da planilha Google Sheets (banco de dados)
  SPREADSHEET_2026_ID: '1Dd7iUABtgVgKux3MkTMVBFFGp2LtrzLDGQ6JR6vW7gs', // Preencher com o ID da planilha "Acompanhamento de OKR's e Indicadores 2026"
  ANO_CORRENTE: 2026,
  ABAS: {
    OE: 'OE',
    RESULTADOS: 'RESULTADOS',
    OKRS: 'OKRS',
    ATIVIDADES: 'ATIVIDADES',
    USUARIOS: 'USUARIOS',
    LOGS: 'LOGS',
    INDICADORES: 'INDICADORES',
    NOMES_DIRETORIAS: 'Nomes_Diretorias',
    COMENTARIOS: 'COMENTARIOS'
  },
  ABAS_2026: {
    SUMARIO: 'Sumario',
    INDICADORES_METAS: 'IndicadoresMetas',
    RESULT_TABS: [
      'DPE.UIF.R01','DPE.UIF.R02','DPE.UAA.R03','DPE.UAA.R04','DPE.UAA.R05',
      'DIF.UME.R06','DIF.UIN.R07','DIF.UIN.R08','DIF.UIN.R09','DIF.UIN.R10',
      'DGE.UPL.R11','DGE.UGP.R12','DGE.UGP.R13','DGE.UGP.R14','DGE.UAJ.R15','DGE.UAJ.R16',
      'DFI.GIT.R17','DFI.DP.R18','DFI.DP.R19','DFI.DP.R20','DFI.AVA.R21','DFI.AVA.R22'
    ]
  },
  EIXOS: ['ALUNOS', 'PROFESSORES', 'GESTAO ESCOLAR', 'CAPACIDADE INSTITUCIONAL'],
  PERFIS: ['GERENCIA', 'PLANEJAMENTO', 'SECRETARIO'],
  STATUS_ATIVIDADE: ['Nao iniciada', 'No prazo', 'Em andamento', 'Concluido', 'Parcialmente concluido', 'Nao concluido', 'Atrasado', 'Replanejado', 'Removido'],
  STATUS_OKR: ['Em andamento', 'Concluido', 'Parcialmente concluido', 'Nao concluido', 'No prazo', 'Em atraso'],
  MESES: ['Janeiro', 'Fevereiro', 'Marco', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
  TRIMESTRES: ['1o', '2o', '3o', '4o']
};

// ============ HELPERS DE ID ============

/**
 * Preencher numero com zeros a esquerda
 */
function padNum(n, digitos) {
  var s = String(n);
  while (s.length < digitos) s = '0' + s;
  return s;
}

/**
 * Gerar proximo ID sequencial para uma sheet
 * Le todos os IDs existentes, encontra o maior numero, retorna proximo
 */
function gerarProximoId(sheet, prefixo, digitos) {
  var dados = sheet.getDataRange().getValues();
  var maxNum = 0;
  var regex = new RegExp('^' + prefixo + '(\\d+)$');
  for (var i = 1; i < dados.length; i++) {
    var id = String(dados[i][0] || '');
    var m = id.match(regex);
    if (m) {
      var num = parseInt(m[1], 10);
      if (num > maxNum) maxNum = num;
    }
  }
  return prefixo + padNum(maxNum + 1, digitos);
}

/**
 * Mapa de nome de aba (ex: 'DPE.UIF.R01') para ID simples (ex: 'R01')
 * Derivado de RESULT_TABS via regex R(\d+)$
 */
var TAB_TO_ID = {};
(function() {
  var tabs = CONFIG.ABAS_2026.RESULT_TABS;
  for (var i = 0; i < tabs.length; i++) {
    var m = tabs[i].match(/R(\d+)$/);
    if (m) {
      TAB_TO_ID[tabs[i]] = 'R' + padNum(parseInt(m[1], 10), 2);
    }
  }
})();

/**
 * Funcao principal: serve a aplicacao web
 * Verifica sessao primeiro - se logado, mostra App; senao, mostra Login
 */
function doGet(e) {
  try {
    var usuario = getUsuarioLogado();

    if (usuario) {
      return HtmlService.createTemplateFromFile('App')
        .evaluate()
        .setTitle('SED Joinville - Monitoramento Estrategico PEI 2025-2029')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }

    return HtmlService.createTemplateFromFile('Login')
      .evaluate()
      .setTitle('SED Joinville - Sistema de Monitoramento Estrategico')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (erro) {
    return HtmlService.createHtmlOutput(
      '<h2>Erro ao carregar o sistema</h2>' +
      '<p>' + erro.message + '</p>' +
      '<p><a href="' + ScriptApp.getService().getUrl() + '">Tentar novamente</a></p>'
    );
  }
}

/**
 * Retorna a URL real da Web App (para redirecionamento seguro)
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Incluir arquivo HTML parcial (para CSS/JS)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Obter a planilha ativa ou por ID
 */
function getSpreadsheet() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Obter uma aba especifica
 */
function getSheet(nomeAba) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(nomeAba);
  if (!sheet) {
    throw new Error('Aba "' + nomeAba + '" nao encontrada na planilha.');
  }
  return sheet;
}

/**
 * Inicializar estrutura do banco de dados (criar abas se nao existirem)
 */
function inicializarBanco() {
  var ss = getSpreadsheet();

  // Estrutura de cada aba com seus cabecalhos
  var estrutura = {
    'OE': ['id_oe', 'eixo', 'descricao', 'objetivo_impacto', 'tipo'],
    'RESULTADOS': ['id_resultado', 'id_oe', 'descricao', 'meta_2025', 'meta_2029', 'classificacao', 'responsavel', 'macroprocesso', 'projeto_codigo', 'projeto_nome', 'diretoria', 'unidade', 'meta_2026', 'requisitos_qualidade'],
    'OKRS': ['id_okr', 'id_resultado', 'id_oe', 'trimestre', 'descricao', 'situacao', 'percentual', 'requisitos_qualidade', 'ano', 'diretoria', 'unidade', 'eixo'],
    'ATIVIDADES': ['id_atividade', 'id_okr', 'id_resultado', 'id_oe', 'descricao', 'mes_planejado', 'mes_realizado', 'criterio', 'peso', 'responsavel', 'observacoes', 'status', 'diretoria', 'unidade', 'eixo', 'percentual_execucao', 'mes_entrega'],
    'USUARIOS': ['id_usuario', 'email', 'nome', 'perfil', 'diretoria', 'senha'],
    'LOGS': ['timestamp', 'email', 'id_registro', 'tipo_registro', 'campo_alterado', 'valor_antigo', 'valor_novo'],
    'INDICADORES': ['id_indicador', 'eixo', 'id_oe', 'categoria', 'projeto_gerencia', 'cod', 'indicador', 'periodicidade', 'responsavel', 'origem', 'pei', 'formula', 'val_2021', 'val_2022', 'val_2023', 'val_2024', 'val_2025', 'meta_2026', 'polaridade', 'jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez', 'status_coleta', 'pct_evolucao', 'pct_atingimento', 'atingimento_meta', 'observacao'],
    'COMENTARIOS': ['id_comentario', 'timestamp', 'autor_email', 'autor_nome', 'perfil', 'secao', 'texto', 'status_dev', 'resposta_dev', 'data_resposta']
  };

  for (var nomeAba in estrutura) {
    var sheet = ss.getSheetByName(nomeAba);
    var cabecalhos = estrutura[nomeAba];
    if (!sheet) {
      // Aba nao existe: criar do zero
      sheet = ss.insertSheet(nomeAba);
      sheet.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
      sheet.getRange(1, 1, 1, cabecalhos.length)
        .setFontWeight('bold')
        .setBackground('#1B3A5C')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    } else {
      // Aba ja existe: atualizar cabecalhos se diferentes
      var colsAtuais = sheet.getLastColumn() || 0;
      var headersAtuais = colsAtuais > 0 ? sheet.getRange(1, 1, 1, colsAtuais).getValues()[0] : [];
      var primeiro = headersAtuais[0] || '';
      var ultimo = headersAtuais[headersAtuais.length - 1] || '';
      if (primeiro !== cabecalhos[0] || ultimo !== cabecalhos[cabecalhos.length - 1] || colsAtuais !== cabecalhos.length) {
        // Limpar row 1 inteira e reescrever
        if (colsAtuais > 0) sheet.getRange(1, 1, 1, colsAtuais).clear();
        sheet.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
        sheet.getRange(1, 1, 1, cabecalhos.length)
          .setFontWeight('bold')
          .setBackground('#1B3A5C')
          .setFontColor('#FFFFFF');
        sheet.setFrozenRows(1);
      }
    }
  }

  // Inserir usuario admin padrao se aba USUARIOS estiver vazia
  var sheetUsuarios = ss.getSheetByName('USUARIOS');
  if (sheetUsuarios.getLastRow() <= 1) {
    sheetUsuarios.appendRow([
      1,
      'admin@sed.joinville.sc.gov.br',
      'Administrador do Sistema',
      'PLANEJAMENTO',
      'SED',
      '123456'
    ]);
  }

  return 'Banco de dados inicializado com sucesso!';
}

/**
 * DIAGNOSTICO: Executar no editor do Apps Script para verificar problemas
 * Menu: Executar > diagnosticarSistema
 */
function diagnosticarSistema() {
  var resultado = [];

  // 1. Verificar planilha
  try {
    var ss = getSpreadsheet();
    resultado.push('OK - Planilha encontrada: ' + ss.getName());
  } catch (e) {
    resultado.push('ERRO - Planilha: ' + e.message);
    resultado.push('>>> Verifique se CONFIG.SPREADSHEET_ID esta preenchido em Code.gs');
  }

  // 2. Verificar abas
  var abas = ['OE', 'RESULTADOS', 'OKRS', 'ATIVIDADES', 'USUARIOS', 'LOGS', 'INDICADORES'];
  for (var i = 0; i < abas.length; i++) {
    try {
      var sheet = getSheet(abas[i]);
      var linhas = sheet.getLastRow();
      resultado.push('OK - Aba ' + abas[i] + ': ' + (linhas - 1) + ' registros');
    } catch (e) {
      resultado.push('ERRO - Aba ' + abas[i] + ': ' + e.message);
    }
  }

  // 3. Verificar sessao ativa
  try {
    var usuario = getUsuarioLogado();
    if (usuario) {
      resultado.push('OK - Sessao ativa: ' + usuario.email + ' (' + usuario.perfil + ') Diretoria: [' + usuario.diretoria + ']');
    } else {
      resultado.push('INFO - Nenhuma sessao ativa');
    }
  } catch (e) {
    resultado.push('ERRO - Sessao: ' + e.message);
  }

  // 4. Verificar getDashboardData
  try {
    var dashData = getDashboardData();
    if (dashData.status === 'ok') {
      resultado.push('OK - Dashboard: ' + dashData.totalAtividades + ' atividades, ' + dashData.totalOEs + ' OEs, ' + dashData.totalOKRs + ' OKRs');
    } else {
      resultado.push('ERRO - Dashboard: ' + dashData.mensagem);
    }
  } catch (e) {
    resultado.push('ERRO - Dashboard: ' + e.message);
  }

  // 5. Verificar usuarios cadastrados
  try {
    var sheetU = getSheet('USUARIOS');
    var dadosU = sheetU.getDataRange().getValues();
    resultado.push('--- Usuarios cadastrados ---');
    for (var j = 1; j < dadosU.length; j++) {
      resultado.push('  ' + dadosU[j][1] + ' | Perfil: ' + dadosU[j][3] + ' | Diretoria: [' + dadosU[j][4] + ']');
    }
  } catch (e) {
    resultado.push('ERRO - Usuarios: ' + e.message);
  }

  var mensagem = resultado.join('\n');
  Logger.log(mensagem);
  return mensagem;
}

/**
 * Importar dados do PEI para as abas do sistema
 */
function importarDadosPEI() {
  var ss = getSpreadsheet();
  var sheetOE = getSheet('OE');
  var sheetResultados = getSheet('RESULTADOS');

  // Objetivos de Impacto
  var ois = [
    ['OI01', 'ALUNOS', 'Garantir que toda crianca e adolescente tenha seu direito a educacao garantido', '', 'IMPACTO'],
    ['OI02', 'ALUNOS', 'Apoiar todos os estudantes para que avancem na idade certa, sem interrupcoes ou atrasos na sua trajetoria', '', 'IMPACTO'],
    ['OI03', 'ALUNOS', 'Promover alta aprendizagem e desenvolvimento integral de todos os estudantes', '', 'IMPACTO']
  ];

  // Objetivos Estrategicos
  var oes = [
    ['OE01', 'ALUNOS', 'Expandir a oferta de vagas com foco no ensino em tempo integral', 'OI01', 'ESTRATEGICO'],
    ['OE02', 'ALUNOS', 'Estruturar curriculo de excelencia, voltado ao desenvolvimento integral', 'OI03', 'ESTRATEGICO'],
    ['OE03', 'ALUNOS', 'Oferecer apoio e estrategias diferenciadas para estudantes com dificuldades de aprendizagem', 'OI03', 'ESTRATEGICO'],
    ['OE04', 'ALUNOS', 'Atingir altos indices de frequencia e engajamento escolar', 'OI02', 'ESTRATEGICO'],
    ['OE05', 'ALUNOS', 'Ampliar e qualificar o atendimento aos estudantes publico-alvo da Educacao Especial', 'OI01', 'ESTRATEGICO'],
    ['OE06', 'PROFESSORES', 'Aprimorar o processo de selecao, integracao e desenvolvimento de atuais e futuros de professores', '', 'ESTRATEGICO'],
    ['OE07', 'PROFESSORES', 'Disponibilizar materiais de apoio de alta qualidade para a pratica docente', '', 'ESTRATEGICO'],
    ['OE08', 'GESTAO ESCOLAR', 'Promover a saude, a motivacao e o bem-estar no ambiente escolar', '', 'ESTRATEGICO'],
    ['OE09', 'GESTAO ESCOLAR', 'Fortalecer a lideranca e garantir uma gestao escolar agil, efetiva e de excelencia', '', 'ESTRATEGICO'],
    ['OE10', 'CAPACIDADE INSTITUCIONAL', 'Consolidar uma gestao de rede educacional eficiente, moderna e inovadora', '', 'ESTRATEGICO'],
    ['OE11', 'CAPACIDADE INSTITUCIONAL', 'Implementar solucoes eficazes e sustentaveis para a melhoria continua da infraestrutura escolar', '', 'ESTRATEGICO'],
    ['OE12', 'CAPACIDADE INSTITUCIONAL', 'Assegurar que as unidades escolares tenham servidores suficientes, motivados, preparados e presentes', '', 'ESTRATEGICO']
  ];

  // Limpar e inserir OEs
  if (sheetOE.getLastRow() > 1) {
    sheetOE.getRange(2, 1, sheetOE.getLastRow() - 1, sheetOE.getLastColumn()).clear();
  }

  var todosOE = ois.concat(oes);
  if (todosOE.length > 0) {
    sheetOE.getRange(2, 1, todosOE.length, todosOE[0].length).setValues(todosOE);
  }

  return 'Dados do PEI importados com sucesso! ' + todosOE.length + ' objetivos inseridos.';
}

// ============ IMPORTACAO PLANILHA 2026 ============

/**
 * Importar dados da planilha de acompanhamento 2026
 * Le as abas de resultados, OKRs e atividades da planilha externa
 */
function importarDados2026() {
  if (!CONFIG.SPREADSHEET_2026_ID) {
    return 'Erro: SPREADSHEET_2026_ID nao configurado em CONFIG.';
  }

  var ss2026 = SpreadsheetApp.openById(CONFIG.SPREADSHEET_2026_ID);
  var ssBanco = getSpreadsheet();

  // 1. Ler Sumario (lista de resultados com OE, metas, classificacao)
  var sumarioData = lerSumario(ss2026);
  Logger.log('Sumario: ' + sumarioData.linhas.length + ' linhas lidas');

  // 2. Ler IndicadoresMetas (periodicidade, formula, linha_base, etc.)
  var indicadoresData = lerIndicadoresMetas(ss2026);
  Logger.log('IndicadoresMetas: ' + Object.keys(indicadoresData.porChave).length + ' chaves');

  // 2b. Construir mapa OE→eixo a partir do banco
  var oeEixoMap = {};
  try {
    var sheetOE = getSheet(CONFIG.ABAS.OE);
    var dadosOE = sheetOE.getDataRange().getValues();
    for (var oei = 1; oei < dadosOE.length; oei++) {
      if (dadosOE[oei][0]) oeEixoMap[String(dadosOE[oei][0])] = String(dadosOE[oei][1] || '');
    }
  } catch (e) { Logger.log('Erro ao ler OE para mapa eixo: ' + e.message); }

  // 3. Preparar arrays para inserir
  var todosResultados = [];
  var todosOKRs = [];
  var todasAtividades = [];
  var avisos = [];

  // 4. Para cada aba de resultado, parsear OKRs e atividades
  var tabs = CONFIG.ABAS_2026.RESULT_TABS;
  var globalOkrCounter = 0;
  var globalAtvCounter = 0;
  for (var t = 0; t < tabs.length; t++) {
    var tabName = tabs[t];
    var sheet2026 = encontrarAba2026(ss2026, tabName);
    if (!sheet2026) {
      avisos.push('Aba nao encontrada: ' + tabName);
      continue;
    }

    var abaData = lerAbaResultado(sheet2026, tabName, sumarioData, indicadoresData, oeEixoMap, globalOkrCounter, globalAtvCounter);
    if (abaData.resultado) todosResultados.push(abaData.resultado);
    todosOKRs = todosOKRs.concat(abaData.okrs);
    todasAtividades = todasAtividades.concat(abaData.atividades);
    globalOkrCounter += abaData.okrs.length;
    globalAtvCounter += abaData.atividades.length;
    Logger.log('Aba ' + tabName + ': OE=' + (abaData.resultado ? abaData.resultado[1] : '?') + ', ' + abaData.okrs.length + ' OKRs, ' + abaData.atividades.length + ' atividades');
  }

  // 5. Escrever no banco de dados (limpar dados antigos, inserir novos)
  escreverDados(ssBanco, 'RESULTADOS', todosResultados);
  escreverDados(ssBanco, 'OKRS', todosOKRs);
  escreverDados(ssBanco, 'ATIVIDADES', todasAtividades);

  // 6. Importar entidade INDICADORES da aba IndicadoresMetas
  var UNIDADE_DIRETORIA = {
    'UIF': 'DPE', 'UAA': 'DPE',
    'UME': 'DIF', 'UIN': 'DIF',
    'UPL': 'DGE', 'UGP': 'DGE', 'UAJ': 'DGE',
    'GIT': 'DFI', 'DP': 'DFI', 'AVA': 'DFI'
  };

  // Construir mapa OE+unidade→id_resultado para vincular indicadores a resultados
  var oeUnitResultMap = {};
  for (var ri = 0; ri < todosResultados.length; ri++) {
    var resRow = todosResultados[ri];
    var resOE = String(resRow[1] || '');     // id_oe
    var resUnit = String(resRow[11] || '');   // unidade
    if (resOE && resUnit) {
      oeUnitResultMap[resOE + '.' + resUnit] = String(resRow[0] || ''); // id_resultado
    }
  }

  var todosIndicadores = [];
  var indCounter = 1;
  var chaves = Object.keys(indicadoresData.porChave);
  for (var ci = 0; ci < chaves.length; ci++) {
    var indList = indicadoresData.porChave[chaves[ci]];
    for (var ii = 0; ii < indList.length; ii++) {
      var ind = indList[ii];
      var newIndId = 'IND' + padNum(indCounter++, 3);
      todosIndicadores.push([
        newIndId,                       // id_indicador
        ind.eixo,                       // eixo
        ind.id_oe,                      // id_oe
        ind.categoria,                  // categoria
        ind.projeto_gerencia,           // projeto_gerencia
        ind.cod,                        // cod
        ind.indicador,                  // indicador
        ind.periodicidade,              // periodicidade
        ind.responsavel,                // responsavel
        ind.origem,                     // origem
        ind.pei,                        // pei
        ind.formula,                    // formula
        ind.val_2021,                   // val_2021
        ind.val_2022,                   // val_2022
        ind.val_2023,                   // val_2023
        ind.val_2024,                   // val_2024
        ind.val_2025,                   // val_2025
        ind.meta_2026,                  // meta_2026
        ind.polaridade,                 // polaridade
        ind.jan,                        // jan
        ind.fev,                        // fev
        ind.mar,                        // mar
        ind.abr,                        // abr
        ind.mai,                        // mai
        ind.jun,                        // jun
        ind.jul,                        // jul
        ind.ago,                        // ago
        ind.set,                        // set
        ind.out,                        // out
        ind.nov,                        // nov
        ind.dez,                        // dez
        ind.status_coleta,              // status_coleta
        ind.pct_evolucao,               // pct_evolucao
        ind.pct_atingimento,            // pct_atingimento
        ind.atingimento_meta,           // atingimento_meta
        ind.observacao_ind              // observacao
      ]);
    }
  }
  Logger.log('INDICADORES para escrever: ' + todosIndicadores.length + ' linhas');
  if (todosIndicadores.length > 0) {
    escreverDados(ssBanco, 'INDICADORES', todosIndicadores);
    Logger.log('INDICADORES escritos com sucesso: ' + todosIndicadores.length);
  } else {
    Logger.log('ATENCAO: Nenhum indicador para escrever!');
  }

  var msg = 'Importacao 2026 concluida: ' + todosResultados.length + ' resultados, ' +
    todosOKRs.length + ' OKRs, ' + todasAtividades.length + ' atividades, ' +
    todosIndicadores.length + ' indicadores.';
  if (avisos.length > 0) msg += '\nAvisos: ' + avisos.join('; ');
  Logger.log(msg);
  return msg;
}

/**
 * DIAGNOSTICO: Verificar a aba IndicadoresMetas da planilha mae
 * Rodar esta funcao no editor para ver o que esta na planilha sem importar
 */
function diagnosticoIndicadoresMetas() {
  var ss2026 = SpreadsheetApp.openById(CONFIG.SPREADSHEET_2026_ID);
  var sheet = ss2026.getSheetByName(CONFIG.ABAS_2026.INDICADORES_METAS);
  if (!sheet) return 'Aba IndicadoresMetas nao encontrada!';

  var dados = sheet.getDataRange().getValues();
  var totalLinhas = dados.length;
  var totalColunas = dados[0] ? dados[0].length : 0;

  var log = '=== DIAGNOSTICO IndicadoresMetas ===\n';
  log += 'Total linhas: ' + totalLinhas + ', Total colunas: ' + totalColunas + '\n\n';

  // Mostrar primeiras 3 linhas (cabecalhos)
  for (var h = 0; h < Math.min(3, totalLinhas); h++) {
    log += 'Linha ' + (h+1) + ': ';
    for (var c = 0; c < Math.min(totalColunas, 35); c++) {
      var val = String(dados[h][c] || '').trim();
      if (val) log += '[' + c + ']=' + val.substring(0, 20) + '  ';
    }
    log += '\n';
  }

  // Contar linhas com IND no COD e linhas com indicador preenchido
  var comIND = 0, semIND = 0, comNome = 0, totalNaoVazias = 0;
  var formatos = {};
  for (var i = 1; i < totalLinhas; i++) {
    var temConteudo = false;
    for (var cc = 0; cc < Math.min(totalColunas, 10); cc++) {
      if (String(dados[i][cc] || '').trim()) { temConteudo = true; break; }
    }
    if (!temConteudo) continue;
    totalNaoVazias++;

    // Verificar cada coluna para achar CODs com IND
    var achouIND = false;
    for (var ci = 0; ci < totalColunas; ci++) {
      var v = String(dados[i][ci] || '').trim();
      if (v.indexOf('IND') === 0) {
        achouIND = true;
        var fmt = v.split('.').length + ' partes';
        formatos[fmt] = (formatos[fmt] || 0) + 1;
        break;
      }
    }
    if (achouIND) comIND++;
    else semIND++;

    // Verificar se alguma coluna tem nome longo (provavel indicador)
    for (var cn = 0; cn < totalColunas; cn++) {
      var vn = String(dados[i][cn] || '').trim();
      if (vn.length > 20 && vn.indexOf('IND') !== 0) { comNome++; break; }
    }
  }

  log += '\n--- CONTAGENS ---\n';
  log += 'Linhas nao vazias: ' + totalNaoVazias + '\n';
  log += 'Com COD iniciando em IND: ' + comIND + '\n';
  log += 'Sem COD IND (mas com conteudo): ' + semIND + '\n';
  log += 'Com nome de indicador (texto >20 chars): ' + comNome + '\n';
  log += 'Formatos de COD: ' + JSON.stringify(formatos) + '\n';

  // Amostra das linhas sem IND
  if (semIND > 0) {
    log += '\n--- AMOSTRA linhas SEM IND (primeiras 10) ---\n';
    var cnt = 0;
    for (var j = 1; j < totalLinhas && cnt < 10; j++) {
      var temC = false;
      for (var cj = 0; cj < Math.min(totalColunas, 10); cj++) {
        if (String(dados[j][cj] || '').trim()) { temC = true; break; }
      }
      if (!temC) continue;
      var achouI = false;
      for (var ck = 0; ck < totalColunas; ck++) {
        if (String(dados[j][ck] || '').trim().indexOf('IND') === 0) { achouI = true; break; }
      }
      if (!achouI) {
        cnt++;
        log += 'Linha ' + (j+1) + ': ';
        for (var cl = 0; cl < Math.min(totalColunas, 8); cl++) {
          var vl = String(dados[j][cl] || '').trim();
          if (vl) log += '[' + cl + ']=' + vl.substring(0, 25) + '  ';
        }
        log += '\n';
      }
    }
  }

  Logger.log(log);
  return log;
}

/**
 * Encontrar aba na planilha 2026 com tolerancia a espacos e variantes de nome
 */
function encontrarAba2026(ss, tabName) {
  // Tentativa 1: nome exato
  var sheet = ss.getSheetByName(tabName);
  if (sheet) return sheet;

  // Tentativa 2: com espaco no final (ex: "DPE.UAA.R03 ")
  sheet = ss.getSheetByName(tabName + ' ');
  if (sheet) return sheet;

  // Tentativa 3: buscar por nome trimmed em todas as abas
  var allSheets = ss.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName().trim() === tabName.trim()) {
      return allSheets[i];
    }
  }

  return null;
}

/**
 * Ler aba Sumario da planilha 2026
 * Retorna lista plana de linhas para busca por OE + descricao fuzzy
 * Estrutura: Row 2 = cabecalhos, Row 3+ = dados
 * Colunas: A=COD.OE, B=Objetivo, C=COD.Resultado2029, D=seq, E=Resultado2029,
 *          F=COD.Resultado2026, G=seq, H=Resultado2026, I=COD.Resultado2025,
 *          J=seq, K=Resultado2025, L=Classificacao, M=Responsavel R25,
 *          N=COD.Projeto/Unidade, O=Projeto/Setor, P=Responsavel PRJ/Gerencia
 */
function lerSumario(ss2026) {
  var resultado = { linhas: [] };
  try {
    var sheet = ss2026.getSheetByName(CONFIG.ABAS_2026.SUMARIO);
    if (!sheet) return resultado;
    var dados = sheet.getDataRange().getValues();
    var currentOE = '';

    for (var i = 2; i < dados.length; i++) {
      // Coluna A (0) = COD. OE - pode estar vazio (celula mesclada, herda do anterior)
      var oeVal = String(dados[i][0] || '').trim();
      if (oeVal && oeVal.match(/O[EI]\s*\d/)) {
        // Normalizar: "OE3" -> "OE03", "OE 9" -> "OE09", "OI01" -> "OI01"
        var oeMatch = oeVal.match(/(O[EI])\s*(\d+)/);
        if (oeMatch) {
          var num = parseInt(oeMatch[2], 10);
          currentOE = oeMatch[1] + (num < 10 ? '0' : '') + num;
        }
      }

      var descricao2026 = String(dados[i][7] || '').trim();  // H
      var descricao2029 = String(dados[i][4] || '').trim();  // E
      var descricao = descricao2026 || descricao2029;

      // Pular linhas sem descricao ou sem OE
      if (!descricao || descricao === '?' || !currentOE) continue;

      resultado.linhas.push({
        id_oe: currentOE,
        descricao_2026: descricao2026,
        descricao_2029: descricao2029,
        classificacao: String(dados[i][11] || '').trim(),
        responsavel: String(dados[i][15] || dados[i][12] || '').trim(),
        meta_2025: String(dados[i][10] || '').trim(),
        meta_2026: descricao2026,
        meta_2029: descricao2029,
        projeto_codigo: String(dados[i][13] || '').trim(),
        projeto_nome: String(dados[i][14] || '').trim()
      });
    }
  } catch (e) {
    Logger.log('Erro ao ler Sumario: ' + e.message);
  }
  return resultado;
}

/**
 * Ler aba IndicadoresMetas da planilha 2026
 * Indexa por OE+Unidade para cruzar com abas de resultado (ex: "OE09.UIF")
 * Colunas: A=EIXO, B=OE, C=Categoria, D=Projeto/Gerencia, E=COD (IND.OE09.UIF.1),
 *          F=INDICADOR, G=PERIODICIDADE, H=RESPONSAVEL, I=ORIGEM, J=PEI?,
 *          K=FORMULA, L=Data linha base, M=Linha de base, N=2025, O=Meta 2026,
 *          P=Polaridade, Q-AB=Jan-Dez, AC=Status coleta, AD=% Evolucao
 */
function lerIndicadoresMetas(ss2026) {
  var resultado = { porChave: {} };
  try {
    var sheet = ss2026.getSheetByName(CONFIG.ABAS_2026.INDICADORES_METAS);
    if (!sheet) { Logger.log('IndicadoresMetas: aba NAO encontrada'); return resultado; }
    var dados = sheet.getDataRange().getValues();
    Logger.log('IndicadoresMetas: ' + dados.length + ' linhas total, ' + (dados[0] ? dados[0].length : 0) + ' colunas');

    // === MAPEAMENTO DINAMICO DE CABECALHOS ===
    // Buscar linha de cabecalho e mapear por nome (robusto contra colunas ocultas/reordenadas)
    var headerRow = -1;
    var colMap = {};

    // Aliases para reconhecer cabecalhos (normalizado para lowercase sem acentos)
    var HEADER_ALIASES = {
      'eixo': 'eixo', 'oe': 'id_oe', 'o.e': 'id_oe', 'o.e.': 'id_oe', 'objetivo estrategico': 'id_oe',
      'categoria': 'categoria',
      'projeto': 'projeto_gerencia', 'gerencia': 'projeto_gerencia', 'projeto/gerencia': 'projeto_gerencia', 'setor': 'projeto_gerencia',
      'cod': 'cod', 'cod.': 'cod', 'codigo': 'cod',
      'indicador': 'indicador', 'indicadores': 'indicador', 'nome indicador': 'indicador',
      'periodicidade': 'periodicidade',
      'responsavel': 'responsavel', 'responsável': 'responsavel',
      'origem': 'origem', 'fonte': 'origem',
      'pei': 'pei', 'pei?': 'pei',
      'formula': 'formula', 'fórmula': 'formula',
      'data linha base': 'data_linha_base', 'linha de base': 'data_linha_base',
      'linha base': 'val_2021', '2021': 'val_2021', 'baseline': 'val_2021',
      '2025': 'val_2025',
      'meta 2026': 'meta_2026', 'meta2026': 'meta_2026',
      'polaridade': 'polaridade',
      'jan': 'jan', 'janeiro': 'jan',
      'fev': 'fev', 'fevereiro': 'fev',
      'mar': 'mar', 'marco': 'mar', 'março': 'mar',
      'abr': 'abr', 'abril': 'abr',
      'mai': 'mai', 'maio': 'mai',
      'jun': 'jun', 'junho': 'jun',
      'jul': 'jul', 'julho': 'jul',
      'ago': 'ago', 'agosto': 'ago',
      'set': 'set', 'setembro': 'set',
      'out': 'out', 'outubro': 'out',
      'nov': 'nov', 'novembro': 'nov',
      'dez': 'dez', 'dezembro': 'dez',
      'status coleta': 'status_coleta', 'status': 'status_coleta',
      '% evolucao': 'pct_evolucao', 'evolucao': 'pct_evolucao', '% evolução': 'pct_evolucao',
      '% atingimento': 'pct_atingimento', 'atingimento': 'pct_atingimento',
      'atingimento meta': 'atingimento_meta', 'ating. meta': 'atingimento_meta',
      'observacao': 'observacao_ind', 'observações': 'observacao_ind', 'obs': 'observacao_ind', 'observacoes': 'observacao_ind'
    };

    // Normalizar string para comparacao
    function _normHeader(s) {
      return String(s || '').trim().toLowerCase()
        .replace(/á/g,'a').replace(/é/g,'e').replace(/í/g,'i').replace(/ó/g,'o').replace(/ú/g,'u')
        .replace(/ã/g,'a').replace(/õ/g,'o').replace(/ç/g,'c').replace(/ê/g,'e').replace(/â/g,'a')
        .replace(/[^a-z0-9 ./?%]/g, '').trim();
    }

    // Buscar cabecalho nas primeiras 5 linhas
    for (var h = 0; h < Math.min(dados.length, 5); h++) {
      var matchCount = 0;
      for (var c = 0; c < dados[h].length; c++) {
        var norm = _normHeader(dados[h][c]);
        if (HEADER_ALIASES[norm]) matchCount++;
      }
      // Se encontramos pelo menos 5 cabecalhos reconhecidos, esta eh a linha de cabecalho
      if (matchCount >= 5) {
        headerRow = h;
        break;
      }
    }

    if (headerRow >= 0) {
      // Mapear cabecalhos por nome
      for (var c2 = 0; c2 < dados[headerRow].length; c2++) {
        var normH = _normHeader(dados[headerRow][c2]);
        if (HEADER_ALIASES[normH] && colMap[HEADER_ALIASES[normH]] === undefined) {
          colMap[HEADER_ALIASES[normH]] = c2;
        }
      }
      Logger.log('IndicadoresMetas: cabecalho na linha ' + headerRow + ', colunas mapeadas: ' + JSON.stringify(colMap));
    } else {
      // Fallback: buscar coluna COD em qualquer posicao da primeira linha de dados
      Logger.log('IndicadoresMetas: CABECALHO NAO ENCONTRADO - tentando detectar COD');
      // Tentar encontrar a coluna que tem valores IND.*
      var codCol = -1;
      for (var fi = 0; fi < Math.min(dados.length, 10); fi++) {
        for (var fc = 0; fc < dados[fi].length; fc++) {
          if (String(dados[fi][fc] || '').trim().indexOf('IND.') === 0) {
            codCol = fc;
            break;
          }
        }
        if (codCol >= 0) break;
      }
      if (codCol < 0) codCol = 4; // fallback absoluto
      // Montar mapa padrao baseado no offset da coluna COD
      var base = codCol - 4; // COD normalmente era coluna 4 (E)
      if (base < 0) base = 0;
      colMap = {
        'eixo': base+0, 'id_oe': base+1, 'categoria': base+2, 'projeto_gerencia': base+3,
        'cod': codCol, 'indicador': base+5, 'periodicidade': base+6, 'responsavel': base+7,
        'origem': base+8, 'pei': base+9, 'formula': base+10, 'data_linha_base': base+11,
        'val_2021': base+12, 'val_2025': base+13, 'meta_2026': base+14, 'polaridade': base+15,
        'jan': base+16, 'fev': base+17, 'mar': base+18, 'abr': base+19, 'mai': base+20,
        'jun': base+21, 'jul': base+22, 'ago': base+23, 'set': base+24, 'out': base+25,
        'nov': base+26, 'dez': base+27, 'status_coleta': base+28, 'pct_evolucao': base+29,
        'pct_atingimento': base+30, 'atingimento_meta': base+31, 'observacao_ind': base+32
      };
      Logger.log('IndicadoresMetas: fallback colMap com COD em col ' + codCol + ': ' + JSON.stringify(colMap));
    }

    // Helper seguro para ler valor de coluna
    function _col(row, campo) {
      var idx = colMap[campo];
      if (idx === undefined || idx < 0 || idx >= row.length) return '';
      return row[idx];
    }
    function _colStr(row, campo) {
      return String(_col(row, campo) || '').trim();
    }
    function _colVal(row, campo) {
      var v = _col(row, campo);
      return (v !== undefined && v !== null) ? v : '';
    }

    // Linha de inicio dos dados = logo apos cabecalho, ou onde encontramos IND.*
    var startRow = headerRow >= 0 ? headerRow + 1 : 0;
    // Se nao achou cabecalho, buscar primeira linha com IND.* ou conteudo valido
    if (headerRow < 0) {
      for (var s = 0; s < Math.min(dados.length, 10); s++) {
        var testCod = _colStr(dados[s], 'cod');
        if (testCod.indexOf('IND.') === 0 || testCod.indexOf('IND') === 0) { startRow = s; break; }
      }
    }
    Logger.log('IndicadoresMetas: dados comecam na linha ' + startRow);

    var totalLidas = 0;
    var totalAceitas = 0;
    var totalRejeitadas = 0;
    var rejeitadasAmostras = [];

    for (var i = startRow; i < dados.length; i++) {
      var codRes = _colStr(dados[i], 'cod');
      var nomeIndicador = _colStr(dados[i], 'indicador');
      var oeStr = _colStr(dados[i], 'id_oe');

      // Pular linhas totalmente vazias
      if (!codRes && !nomeIndicador && !oeStr) continue;
      totalLidas++;

      // REGRA DE ACEITACAO AMPLIADA:
      // Aceitar se COD comeca com IND (com ou sem ponto)
      // OU se tem nome de indicador preenchido (linha valida mesmo sem cod no formato esperado)
      var codValido = codRes && (codRes.indexOf('IND') === 0 || codRes.indexOf('ind') === 0);
      var temNomeIndicador = nomeIndicador.length > 3;

      if (!codValido && !temNomeIndicador) {
        totalRejeitadas++;
        if (rejeitadasAmostras.length < 5) {
          rejeitadasAmostras.push('Linha ' + (i+1) + ': cod="' + codRes + '", ind="' + nomeIndicador.substring(0,30) + '"');
        }
        continue;
      }
      totalAceitas++;

      // Extrair OE e unidade do codigo: IND.OE09.UIF.1 -> OE=OE09, Unit=UIF
      var codOE = '';
      var codUnit = '';
      if (codRes) {
        // Normalizar separador: IND.OE09.UIF.1, IND-OE09-UIF-1, IND_OE09_UIF_1
        var codNorm = codRes.replace(/[-_]/g, '.');
        var codParts = codNorm.split('.');
        if (codParts.length >= 3) {
          codOE = codParts[1].toUpperCase();
          codUnit = codParts[2].toUpperCase();
        } else if (codParts.length >= 2) {
          codOE = codParts[1].toUpperCase();
        }
      }
      // Fallback: se nao extraiu OE do cod, usar coluna OE diretamente
      if (!codOE && oeStr) {
        var oeMatch = oeStr.match(/(O[EI]\s*\d+)/i);
        if (oeMatch) {
          codOE = oeMatch[1].replace(/\s/g, '').toUpperCase();
          if (codOE.length === 3) codOE = codOE.substring(0,2) + '0' + codOE.substring(2); // OE9 -> OE09
        }
      }

      // Se nao tem COD, gerar um provisorio
      if (!codRes && codOE) {
        codRes = 'IND.' + codOE + '.' + (codUnit || 'GEN') + '.' + totalAceitas;
      }

      var indData = {
        cod: codRes,
        eixo: _colStr(dados[i], 'eixo'),
        id_oe: codOE,
        categoria: _colStr(dados[i], 'categoria'),
        projeto_gerencia: _colStr(dados[i], 'projeto_gerencia'),
        indicador: nomeIndicador,
        periodicidade: _colStr(dados[i], 'periodicidade'),
        responsavel: _colStr(dados[i], 'responsavel'),
        origem: _colStr(dados[i], 'origem'),
        pei: _colStr(dados[i], 'pei'),
        formula: _colStr(dados[i], 'formula'),
        val_2021: '',
        val_2022: '',
        val_2023: '',
        val_2024: '',
        val_2025: _colStr(dados[i], 'val_2025'),
        meta_2026: _colStr(dados[i], 'meta_2026'),
        polaridade: _colStr(dados[i], 'polaridade'),
        jan: _colVal(dados[i], 'jan'),
        fev: _colVal(dados[i], 'fev'),
        mar: _colVal(dados[i], 'mar'),
        abr: _colVal(dados[i], 'abr'),
        mai: _colVal(dados[i], 'mai'),
        jun: _colVal(dados[i], 'jun'),
        jul: _colVal(dados[i], 'jul'),
        ago: _colVal(dados[i], 'ago'),
        set: _colVal(dados[i], 'set'),
        out: _colVal(dados[i], 'out'),
        nov: _colVal(dados[i], 'nov'),
        dez: _colVal(dados[i], 'dez'),
        status_coleta: _colStr(dados[i], 'status_coleta'),
        pct_evolucao: _colStr(dados[i], 'pct_evolucao'),
        pct_atingimento: _colStr(dados[i], 'pct_atingimento'),
        atingimento_meta: _colStr(dados[i], 'atingimento_meta'),
        observacao_ind: _colStr(dados[i], 'observacao_ind')
      };

      var chave = codOE + '.' + codUnit;
      if (!chave || chave === '.') chave = '_SEM_CHAVE_';
      if (!resultado.porChave[chave]) {
        resultado.porChave[chave] = [];
      }
      resultado.porChave[chave].push(indData);
    }

    Logger.log('IndicadoresMetas RESUMO: ' + totalLidas + ' linhas lidas, ' + totalAceitas + ' aceitas, ' + totalRejeitadas + ' rejeitadas');
    if (rejeitadasAmostras.length > 0) {
      Logger.log('IndicadoresMetas REJEITADAS (amostra): ' + rejeitadasAmostras.join(' | '));
    }
    Logger.log('IndicadoresMetas: ' + Object.keys(resultado.porChave).length + ' chaves unicas');
    // Log contagem total de indicadores
    var totalInds = 0;
    for (var ck in resultado.porChave) totalInds += resultado.porChave[ck].length;
    Logger.log('IndicadoresMetas: ' + totalInds + ' indicadores TOTAL no porChave');

  } catch (e) {
    Logger.log('Erro ao ler IndicadoresMetas: ' + e.message + ' | Stack: ' + e.stack);
  }
  return resultado;
}

/**
 * Ler uma aba de resultado da planilha 2026
 * Estrutura:
 *   Row 1 (A1:F1 merged): "OE XX - descricao..."   |  G1:H3 merged: "Responsavel: ..."
 *   Row 2 (A2:F2 merged): "Resultado: descricao..."
 *   Row 3 (A3:F3 merged): "Requisitos de Qualidade: ..."
 *   Row 4: Cabecalhos (Trimestre, OKR, Req.Qualidade, Situacao, %, Atividades, Mes Plan, Realiz, Mes Entrega, Resp, Peso)
 *   Row 5+: Blocos de linhas por OKR (A-E merged verticalmente)
 */
function lerAbaResultado(sheet, tabName, sumarioData, indicadoresData, oeEixoMap, okrOffset, atvOffset) {
  var resultado = { resultado: null, okrs: [], atividades: [] };
  var dados = sheet.getDataRange().getValues();
  if (dados.length < 5) return resultado;

  // Extrair diretoria e unidade do nome da aba (DPE.UIF.R01 -> dir=DPE, unit=UIF)
  var diretoria = tabName.substring(0, 3);
  var parts = tabName.split('.');
  var unidade = parts.length >= 2 ? parts[1] : '';

  // Row 1 (indice 0): OE ref - texto merged em A1:F1
  var textoOE = String(dados[0][0] || '').trim();
  var oeRef = extrairCodigoOE(textoOE);

  // Row 2 (indice 1): "Resultado: descricao..."
  var textoRes = String(dados[1][0] || '').trim();
  var descResultado = textoRes.replace(/^Resultado\s*:\s*/i, '').trim();

  // Row 3 (indice 2): "Requisitos de Qualidade: ..."
  var textoReq = String(dados[2][0] || '').trim();
  var reqQualidade = textoReq.replace(/^Requisitos\s+de\s+Qualidade\s*:\s*/i, '').trim();

  // Responsavel do cabecalho (G1:H3 merged, pode estar em col 6 ou 7)
  var respCabecalho = String(dados[0][6] || dados[0][7] || '').trim();
  respCabecalho = respCabecalho.replace(/^Respons[aá]vel\s*:\s*/i, '').trim();

  // Buscar dados complementares do Sumario (por OE + descricao fuzzy)
  var sumInfo = encontrarSumario(sumarioData, oeRef, descResultado);

  // Buscar dados complementares de IndicadoresMetas (por OE + unidade)
  var indInfo = encontrarIndicador(indicadoresData, oeRef, unidade);

  // Derivar eixo do OE via mapa
  var eixo = (oeEixoMap && oeEixoMap[oeRef]) ? oeEixoMap[oeRef] : '';

  // Criar registro de resultado (14 colunas)
  // id_oe vem PRIMARIAMENTE da aba (Row 1), nao do Sumario
  var simpleResId = TAB_TO_ID[tabName] || tabName;
  resultado.resultado = [
    simpleResId,                                    // id_resultado
    oeRef || sumInfo.id_oe || '',                   // id_oe (PRIMARIO: da aba)
    descResultado || '',                             // descricao (da aba)
    sumInfo.meta_2025 || '',                         // meta_2025
    sumInfo.meta_2029 || '',                         // meta_2029
    sumInfo.classificacao || '',                     // classificacao
    sumInfo.responsavel || respCabecalho || '',      // responsavel
    '',                                              // macroprocesso
    sumInfo.projeto_codigo || '',                    // projeto_codigo
    sumInfo.projeto_nome || '',                      // projeto_nome
    diretoria,                                       // diretoria
    unidade,                                         // unidade (NOVO)
    sumInfo.meta_2026 || '',                         // meta_2026
    reqQualidade                                     // requisitos_qualidade
  ];

  // Parsear OKRs e atividades a partir da Row 5 (indice 4)
  var okrCounter = 0;
  var atvCounter = 0;
  var currentOKR = null;
  var currentTrimestre = '';

  for (var i = 4; i < dados.length; i++) {
    var row = dados[i];

    // Coluna A (0) = Trimestre - so tem valor na primeira linha do bloco (merged)
    var triVal = String(row[0] || '').trim();
    if (triVal) {
      var triMatch = triVal.match(/(\d)/);
      if (triMatch) currentTrimestre = triMatch[1] + 'o';
    }

    // Coluna B (1) = OKR - so tem valor na primeira linha do bloco (merged)
    var okrVal = String(row[1] || '').trim();
    if (okrVal) {
      okrCounter++;
      var idOkr = 'OKR' + padNum(okrOffset + okrCounter, 3);
      var reqQualOKR = String(row[2] || '').trim();  // C = Req. Qualidade do OKR
      var sitOKR = String(row[3] || '').trim();       // D = Situacao
      var pctOKR = parseFloat(row[4]) || 0;           // E = %

      currentOKR = idOkr;
      resultado.okrs.push([
        idOkr,                    // id_okr
        simpleResId,              // id_resultado
        oeRef || '',              // id_oe (NOVO)
        currentTrimestre,         // trimestre
        okrVal,                   // descricao
        sitOKR || 'No prazo',    // situacao
        pctOKR,                   // percentual
        reqQualOKR,               // requisitos_qualidade
        CONFIG.ANO_CORRENTE,      // ano
        diretoria,                // diretoria (NOVO)
        unidade,                  // unidade (NOVO)
        eixo                      // eixo (NOVO)
      ]);
    }

    // Coluna F (5) = Atividade - cada linha do bloco pode ter uma atividade
    var atvVal = String(row[5] || '').trim();
    if (atvVal && currentOKR) {
      atvCounter++;
      var idAtv = 'ATV' + padNum(atvOffset + atvCounter, 3);
      var mesPlan = String(row[6] || '').trim();      // G = Mes Planejado
      var mesReal = String(row[7] || '').trim();      // H = Realizado
      var mesEntrega = String(row[8] || '').trim();   // I = Mes Entrega
      var respAtv = String(row[9] || '').trim();      // J = Responsavel
      var pesoAtv = row[10] || '';                     // K = Peso da Atividade
      // Normalizar status da coluna H (Realizado) para padroes do sistema
      var statusAtv = 'Sem Status';
      if (mesReal) {
        var sLower = mesReal.toLowerCase().replace(/[áàã]/g,'a').replace(/[éè]/g,'e').replace(/[íì]/g,'i').replace(/[óò]/g,'o');
        if (sLower === 'realizado') statusAtv = 'Concluido';
        else if (sLower === 'em andamento') statusAtv = 'Em andamento';
        else if (sLower === 'atrasado') statusAtv = 'Atrasado';
        else if (sLower === 'nao iniciado') statusAtv = 'Nao iniciada';
        else statusAtv = mesReal.trim();
      }

      resultado.atividades.push([
        idAtv,                    // id_atividade
        currentOKR,               // id_okr
        simpleResId,              // id_resultado
        oeRef || '',              // id_oe (NOVO)
        atvVal,                   // descricao
        mesPlan,                  // mes_planejado
        '',                       // mes_realizado (vazio - valor de H agora vai para status)
        'Qualitativo',            // criterio
        pesoAtv,                  // peso
        respAtv,                  // responsavel
        '',                       // observacoes
        statusAtv,                // status (normalizado da coluna H "Realizado")
        diretoria,                // diretoria
        unidade,                  // unidade (NOVO)
        eixo,                     // eixo (NOVO)
        '',                       // percentual_execucao (nao utilizado - % derivado do status)
        mesEntrega                // mes_entrega
      ]);
    }
  }

  return resultado;
}

/**
 * Buscar dados do Sumario que correspondem a um resultado (por OE + descricao fuzzy)
 * Como o Sumario nao tem nomes de abas (DPE.UIF.R01), usamos OE + similaridade de texto
 */
function encontrarSumario(sumarioData, oeRef, descResultado) {
  if (!sumarioData || !sumarioData.linhas || !oeRef || !descResultado) return {};

  var descNorm = normalizarTexto(descResultado);
  if (!descNorm) return {};

  var melhorMatch = {};
  var melhorScore = 0;

  for (var i = 0; i < sumarioData.linhas.length; i++) {
    var linha = sumarioData.linhas[i];
    if (linha.id_oe !== oeRef) continue;

    // Calcular similaridade com descricao 2026 e 2029
    var score = 0;
    var desc2026Norm = normalizarTexto(linha.descricao_2026);
    var desc2029Norm = normalizarTexto(linha.descricao_2029);

    if (desc2026Norm) {
      score = Math.max(score, calcularSimilaridade(descNorm, desc2026Norm));
    }
    if (desc2029Norm) {
      score = Math.max(score, calcularSimilaridade(descNorm, desc2029Norm));
    }

    if (score > melhorScore && score > 0.3) {
      melhorScore = score;
      melhorMatch = linha;
    }
  }

  return melhorMatch;
}

/**
 * Buscar dados de IndicadoresMetas por OE + codigo de unidade
 * Tab DPE.UIF.R01 com OE09 -> busca chave "OE09.UIF" no indice de indicadores
 */
function encontrarIndicador(indicadoresData, oeRef, unidade) {
  if (!indicadoresData || !indicadoresData.porChave || !oeRef || !unidade) return {};

  var chave = oeRef + '.' + unidade;
  var indicadores = indicadoresData.porChave[chave];

  if (indicadores && indicadores.length > 0) {
    return indicadores[0]; // Retorna primeiro match
  }

  return {};
}

/**
 * Normalizar texto para comparacao fuzzy (remove acentos, pontuacao, lowercase)
 */
function normalizarTexto(texto) {
  if (!texto) return '';
  return String(texto).toLowerCase()
    .replace(/[áàãâä]/g, 'a')
    .replace(/[éèêë]/g, 'e')
    .replace(/[íìîï]/g, 'i')
    .replace(/[óòõôö]/g, 'o')
    .replace(/[úùûü]/g, 'u')
    .replace(/[ç]/g, 'c')
    .replace(/[^a-z0-9\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Calcular similaridade entre dois textos normalizados (word overlap)
 * Retorna valor entre 0 e 1
 */
function calcularSimilaridade(a, b) {
  var wordsA = a.split(' ');
  var wordsB = b.split(' ');
  var commonWords = 0;
  var significantWords = 0;

  for (var i = 0; i < wordsA.length; i++) {
    if (wordsA[i].length < 3) continue; // pular palavras curtas (de, do, em, etc.)
    significantWords++;
    for (var j = 0; j < wordsB.length; j++) {
      if (wordsA[i] === wordsB[j]) {
        commonWords++;
        break;
      }
    }
  }

  return significantWords > 0 ? commonWords / significantWords : 0;
}

/**
 * Extrair codigo OE de uma string (ex: "OE 09 - Fortalecer..." -> "OE09")
 * Trata variantes: "OE3", "OE 9", "OE09", "OI01"
 */
function extrairCodigoOE(texto) {
  var match = texto.match(/O[EI]\s*(\d+)/i);
  if (match) {
    var num = parseInt(match[1], 10);
    var prefix = texto.match(/(O[EI])\s*\d/i)[1].toUpperCase();
    return prefix + (num < 10 ? '0' : '') + num;
  }
  return '';
}

/**
 * Escrever dados em uma aba do banco (limpar dados antigos, manter cabecalho)
 */
function escreverDados(ss, nomeAba, dados) {
  var sheet = ss.getSheetByName(nomeAba);
  if (!sheet) return;

  // Limpar dados (manter cabecalho)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  }

  // Inserir novos dados
  if (dados.length > 0) {
    sheet.getRange(2, 1, dados.length, dados[0].length).setValues(dados);
  }
}

/**
 * Configurar triggers para importacao automatica do PEI 2026.
 * Executa importarDados2026() diariamente as 7h, 12h e 17h.
 * Rode esta funcao UMA VEZ manualmente no editor do Apps Script.
 */
function configurarTriggers() {
  // Remover triggers existentes da importacao para evitar duplicatas
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'importarDados2026') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Criar 3 triggers diarios: 7h, 12h e 17h
  var horarios = [7, 12, 17];
  for (var h = 0; h < horarios.length; h++) {
    ScriptApp.newTrigger('importarDados2026')
      .timeBased()
      .everyDays(1)
      .atHour(horarios[h])
      .create();
  }

  Logger.log('Triggers configurados: importarDados2026 rodara diariamente as 7h, 12h e 17h.');
}
