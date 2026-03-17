/**
 * ============================================================
 * SERVICO DE API - ENDPOINTS REST VIA doPost
 * ============================================================
 * Permite que outros sistemas (Apps Script, apps externos)
 * consultem e manipulem dados do Sistema PEI via HTTP POST.
 *
 * CONFIGURACAO:
 * 1. Defina sua API_KEY abaixo (troque o valor padrao)
 * 2. Faca deploy como Web App (Executar como: eu, Acesso: qualquer pessoa)
 * 3. Use a URL do deploy para fazer requisicoes POST
 *
 * FORMATO DA REQUISICAO (JSON no body):
 * {
 *   "apiKey": "SUA_CHAVE_AQUI",
 *   "action": "nome_da_acao",
 *   "params": { ...parametros opcionais... }
 * }
 */

// ============ CONFIGURACAO DA API ============
var API_CONFIG = {
  API_KEY: '@sed@joinville',  // IMPORTANTE: Troque esta chave antes de usar em producao!
  VERSAO: '1.0.0',
  RATE_LIMIT_POR_MINUTO: 60
};

// ============ HANDLER PRINCIPAL ============

/**
 * Handler de requisicoes POST (API REST)
 * Recebe JSON com { apiKey, action, params }
 */
function doPost(e) {
  try {
    // Parsear payload
    if (!e || !e.postData || !e.postData.contents) {
      return _respostaAPI({ error: 'Requisicao vazia ou invalida' }, 400);
    }

    var payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return _respostaAPI({ error: 'JSON invalido no corpo da requisicao' }, 400);
    }

    // Validar API Key
    if (!payload.apiKey || payload.apiKey !== API_CONFIG.API_KEY) {
      return _respostaAPI({ error: 'API Key invalida ou ausente' }, 401);
    }

    // Validar acao
    var action = payload.action;
    if (!action) {
      return _respostaAPI({ error: 'Campo "action" e obrigatorio' }, 400);
    }

    var params = payload.params || {};

    // Rotear para a funcao correta
    var resultado = _rotearAcaoAPI(action, params);
    return _respostaAPI(resultado);

  } catch (erro) {
    return _respostaAPI({
      error: 'Erro interno do servidor',
      detalhes: erro.message
    }, 500);
  }
}

// ============ ROTEADOR DE ACOES ============

/**
 * Roteia a acao recebida para a funcao correspondente
 */
function _rotearAcaoAPI(action, params) {
  switch (action) {

    // ---- STATUS / INFO ----
    case 'ping':
      return { success: true, message: 'API operacional', versao: API_CONFIG.VERSAO, timestamp: new Date().toISOString() };

    case 'info':
      return _apiInfo();

    // ---- OBJETIVOS ESTRATEGICOS ----
    case 'listarOEs':
      return { success: true, data: listarOEs() };

    // ---- RESULTADOS ----
    case 'listarResultados':
      return { success: true, data: listarResultados(params.filtroOE || null) };

    case 'getResultadoPorAba':
      if (!params.idResultado) return { success: false, error: 'Parametro "idResultado" e obrigatorio' };
      return { success: true, data: getResultadoPorAba(params.idResultado) };

    // ---- OKRs ----
    case 'listarOKRs':
      return { success: true, data: listarOKRs(params.filtroResultado || null, params.filtroTrimestre || null) };

    case 'getOKRsAgrupados':
      return { success: true, data: getOKRsAgrupados() };

    // ---- ATIVIDADES ----
    case 'listarAtividadesAPI':
      return { success: true, data: _listarAtividadesSemSessao(params.filtros || null) };

    case 'getAtividadesAgrupadas':
      return { success: true, data: _getAtividadesAgrupadasAPI() };

    case 'getDadosGantt':
      return { success: true, data: _getDadosGanttAPI(params.filtros || null) };

    // ---- DASHBOARD ----
    case 'getDashboardResumo':
      return { success: true, data: _getDashboardResumoAPI() };

    case 'getDadosGraficos':
      return { success: true, data: _getDadosGraficosAPI() };

    // ---- INDICADORES ----
    case 'listarIndicadores':
      return { success: true, data: listarIndicadores(params.filtroOE || null) };

    case 'getIndicadoresAgrupados':
      return { success: true, data: getIndicadoresAgrupados() };

    // ---- VISAO HIERARQUICA ----
    case 'getVisaoHierarquica':
      return { success: true, data: _getVisaoHierarquicaAPI(params.filtros || null) };

    case 'getResultadosAgrupados':
      return { success: true, data: getResultadosAgrupados() };

    // ---- DIRETORIAS ----
    case 'listarDiretorias':
      return { success: true, data: listarDiretorias() };

    // ---- BUSCA ----
    case 'buscar':
      if (!params.termo) return { success: false, error: 'Parametro "termo" e obrigatorio' };
      return { success: true, data: buscarGeral(params.termo) };

    default:
      return { success: false, error: 'Acao nao reconhecida: ' + action, acoesDisponiveis: _listarAcoesDisponiveis() };
  }
}

// ============ FUNCOES AUXILIARES DA API ============

/**
 * Formatar resposta como JSON para ContentService
 */
function _respostaAPI(dados, statusCode) {
  var output = ContentService
    .createTextOutput(JSON.stringify(dados))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

/**
 * Retorna informacoes gerais da API
 */
function _apiInfo() {
  var ss = getSpreadsheet();
  return {
    success: true,
    sistema: 'Sistema de Monitoramento Estrategico - SED Joinville',
    versao: API_CONFIG.VERSAO,
    ano: CONFIG.ANO_CORRENTE,
    planilha: ss.getName(),
    eixos: CONFIG.EIXOS,
    acoesDisponiveis: _listarAcoesDisponiveis()
  };
}

/**
 * Lista todas as acoes disponiveis na API
 */
function _listarAcoesDisponiveis() {
  return [
    'ping',
    'info',
    'listarOEs',
    'listarResultados',
    'getResultadoPorAba',
    'listarOKRs',
    'getOKRsAgrupados',
    'listarAtividadesAPI',
    'getAtividadesAgrupadas',
    'getDadosGantt',
    'getDashboardResumo',
    'getDadosGraficos',
    'listarIndicadores',
    'getIndicadoresAgrupados',
    'getVisaoHierarquica',
    'getResultadosAgrupados',
    'listarDiretorias',
    'buscar'
  ];
}

// ============ WRAPPERS SEM SESSAO ============
// As funcoes originais dependem de getUsuarioLogado() para filtrar por diretoria.
// Na API, nao ha sessao ativa, entao criamos versoes que retornam todos os dados.

/**
 * Listar atividades sem depender de sessao (retorna todas)
 */
function _listarAtividadesSemSessao(filtros) {
  var sheet = getSheet(CONFIG.ABAS.ATIVIDADES);
  var dados = sheet.getDataRange().getValues();
  var resultado = [];

  for (var i = 1; i < dados.length; i++) {
    if (!dados[i][0]) continue;

    // Filtros opcionais
    if (filtros) {
      if (filtros.id_okr && String(dados[i][1]) !== String(filtros.id_okr)) continue;
      if (filtros.id_resultado && String(dados[i][2]) !== String(filtros.id_resultado)) continue;
      if (filtros.status && String(dados[i][11]) !== String(filtros.status)) continue;
      if (filtros.diretoria && String(dados[i][12]) !== String(filtros.diretoria)) continue;
      if (filtros.responsavel && String(dados[i][9]).toLowerCase().indexOf(String(filtros.responsavel).toLowerCase()) === -1) continue;
      if (filtros.eixo && String(dados[i][14]) !== String(filtros.eixo)) continue;
    }

    resultado.push({
      id_atividade: dados[i][0],
      id_okr: dados[i][1],
      id_resultado: dados[i][2],
      id_oe: dados[i][3] || '',
      descricao: dados[i][4],
      mes_planejado: dados[i][5],
      mes_realizado: dados[i][6],
      criterio: dados[i][7],
      peso: dados[i][8],
      responsavel: dados[i][9],
      observacoes: dados[i][10],
      status: dados[i][11],
      diretoria: dados[i][12],
      unidade: dados[i][13] || '',
      eixo: dados[i][14] || '',
      percentual_execucao: dados[i][15],
      mes_entrega: dados[i][16] || ''
    });
  }
  return resultado;
}

/**
 * Dashboard resumo sem depender de sessao
 */
function _getDashboardResumoAPI() {
  var oes = listarOEs();
  var resultados = listarResultados();
  var okrs = listarOKRs();
  var atividades = _listarAtividadesSemSessao(null);

  var totalAtividades = atividades.length;
  var concluidas = 0, emAndamento = 0, atrasadas = 0, noPrazo = 0, naoIniciadas = 0, replanejadas = 0;
  var somaPercentual = 0;

  for (var i = 0; i < atividades.length; i++) {
    var status = String(atividades[i].status || '').toLowerCase();
    var pct = parseFloat(atividades[i].percentual_execucao) || 0;
    somaPercentual += pct;

    if (status === 'nao iniciada') naoIniciadas++;
    else if (status.indexOf('conclu') !== -1 && status.indexOf('nao') === -1 && status.indexOf('parcial') === -1) concluidas++;
    else if (status.indexOf('andamento') !== -1) emAndamento++;
    else if (status.indexOf('atras') !== -1) atrasadas++;
    else if (status.indexOf('prazo') !== -1) noPrazo++;
    else if (status.indexOf('replanej') !== -1) replanejadas++;
  }

  var percentualGeral = totalAtividades > 0 ? Math.round(somaPercentual / totalAtividades) : 0;

  // Percentual por OE
  var resultadoOE = {};
  for (var r = 0; r < resultados.length; r++) {
    resultadoOE[resultados[r].id_resultado] = resultados[r].id_oe;
  }

  var percentualPorOE = [];
  for (var oi = 0; oi < oes.length; oi++) {
    var oe = oes[oi];
    var soma = 0, count = 0;
    for (var ai = 0; ai < atividades.length; ai++) {
      if (resultadoOE[atividades[ai].id_resultado] === oe.id_oe) {
        soma += parseFloat(atividades[ai].percentual_execucao) || 0;
        count++;
      }
    }
    percentualPorOE.push({
      id_oe: oe.id_oe, descricao: oe.descricao, eixo: oe.eixo, tipo: oe.tipo,
      percentual: count > 0 ? Math.round(soma / count) : 0, totalAtividades: count
    });
  }

  // Ranking por diretoria
  var rankingDiretoria = calcularRankingDiretoria(atividades);

  // OKRs atrasados
  var okrsAtrasados = [];
  for (var j = 0; j < okrs.length; j++) {
    var sit = String(okrs[j].situacao || '').toLowerCase();
    if (sit.indexOf('atras') !== -1 || sit.indexOf('nao conclu') !== -1) {
      okrsAtrasados.push({
        id: okrs[j].id_okr, descricao: okrs[j].descricao,
        trimestre: okrs[j].trimestre, situacao: okrs[j].situacao, percentual: okrs[j].percentual
      });
    }
  }

  var distribuicaoStatus = {
    'Concluido': concluidas, 'Em andamento': emAndamento, 'Atrasado': atrasadas,
    'No prazo': noPrazo, 'Nao iniciada': naoIniciadas, 'Replanejado': replanejadas,
    'Outros': totalAtividades - concluidas - emAndamento - atrasadas - noPrazo - replanejadas - naoIniciadas
  };

  return {
    percentualGeral: percentualGeral,
    totalAtividades: totalAtividades,
    concluidas: concluidas, emAndamento: emAndamento, atrasadas: atrasadas,
    noPrazo: noPrazo, naoIniciadas: naoIniciadas, replanejadas: replanejadas,
    percentualPorOE: percentualPorOE,
    rankingDiretoria: rankingDiretoria,
    okrsAtrasados: okrsAtrasados,
    distribuicaoStatus: distribuicaoStatus,
    totalOEs: oes.length,
    totalResultados: resultados.length,
    totalOKRs: okrs.length,
    totalIndicadores: contarIndicadores()
  };
}

/**
 * Dados de graficos sem depender de sessao
 */
function _getDadosGraficosAPI() {
  var dashData = _getDashboardResumoAPI();
  return {
    graficoExecucaoOE: dashData.percentualPorOE.map(function(item) {
      return { label: item.id_oe, valor: item.percentual };
    }),
    graficoStatus: dashData.distribuicaoStatus,
    graficoRanking: dashData.rankingDiretoria.slice(0, 10).map(function(item) {
      return { label: item.diretoria, valor: item.percentual };
    })
  };
}

/**
 * Visao hierarquica sem depender de sessao
 */
function _getVisaoHierarquicaAPI(filtros) {
  var oes = listarOEs();
  var resultados = listarResultados();
  var okrs = listarOKRs();
  var atividades = _listarAtividadesSemSessao(null);

  if (filtros) {
    if (filtros.eixo) oes = oes.filter(function(oe) { return oe.eixo === filtros.eixo; });
    if (filtros.id_oe) oes = oes.filter(function(oe) { return oe.id_oe === filtros.id_oe; });
  }

  var hierarquia = [];
  for (var i = 0; i < oes.length; i++) {
    var oe = oes[i];
    var resultadosOE = resultados.filter(function(r) { return r.id_oe === oe.id_oe; });
    var oeNode = { oe: oe, resultados: [] };

    for (var j = 0; j < resultadosOE.length; j++) {
      var res = resultadosOE[j];
      var okrsRes = okrs.filter(function(o) { return o.id_resultado === res.id_resultado; });
      var resNode = { resultado: res, okrs: [] };

      for (var k = 0; k < okrsRes.length; k++) {
        var okr = okrsRes[k];
        var atvsOKR = atividades.filter(function(a) { return a.id_okr === okr.id_okr; });
        resNode.okrs.push({ okr: okr, atividades: atvsOKR });
      }
      oeNode.resultados.push(resNode);
    }
    hierarquia.push(oeNode);
  }
  return hierarquia;
}

/**
 * Atividades agrupadas sem depender de sessao
 */
function _getAtividadesAgrupadasAPI() {
  var atividades = _listarAtividadesSemSessao(null);
  var okrs = listarOKRs();
  var total = atividades.length;
  var concluidas = 0, emAndamento = 0, atrasadas = 0, somaPercentual = 0;

  var okrMap = {};
  for (var k = 0; k < okrs.length; k++) {
    okrMap[okrs[k].id_okr] = okrs[k];
  }

  var dirMap = {};
  for (var i = 0; i < atividades.length; i++) {
    var atv = atividades[i];
    var st = String(atv.status || '').toLowerCase();
    var pct = parseFloat(atv.percentual_execucao) || 0;
    somaPercentual += pct;

    if (st.indexOf('conclu') !== -1 && st.indexOf('nao') === -1 && st.indexOf('parcial') === -1) concluidas++;
    else if (st.indexOf('andamento') !== -1) emAndamento++;
    else if (st.indexOf('atras') !== -1) atrasadas++;

    var dir = String(atv.diretoria || '').trim() || 'Sem diretoria';
    if (!dirMap[dir]) dirMap[dir] = [];

    var okrInfo = okrMap[atv.id_okr];
    atv.okr_descricao = okrInfo ? okrInfo.descricao : '';
    atv.okr_trimestre = okrInfo ? okrInfo.trimestre : '';
    dirMap[dir].push(atv);
  }

  var grupos = [];
  var dirs = Object.keys(dirMap).sort();
  for (var d = 0; d < dirs.length; d++) {
    grupos.push({ diretoria: dirs[d], atividades: dirMap[dirs[d]] });
  }

  return {
    grupos: grupos, total: total, concluidas: concluidas,
    emAndamento: emAndamento, atrasadas: atrasadas,
    percentualMedio: total > 0 ? Math.round(somaPercentual / total) : 0
  };
}

/**
 * Dados Gantt sem depender de sessao
 */
function _getDadosGanttAPI(filtros) {
  var atividades = _listarAtividadesSemSessao(null);
  var resultados = listarResultados();
  var okrs = listarOKRs();

  var resMap = {};
  for (var r = 0; r < resultados.length; r++) {
    resMap[resultados[r].id_resultado] = resultados[r];
  }
  var okrMap = {};
  for (var k = 0; k < okrs.length; k++) {
    okrMap[okrs[k].id_okr] = okrs[k];
  }

  var resultado = [];
  for (var i = 0; i < atividades.length; i++) {
    var atv = atividades[i];
    var res = resMap[atv.id_resultado] || {};
    var okr = okrMap[atv.id_okr] || {};

    if (filtros) {
      if (filtros.diretoria && String(atv.diretoria) !== String(filtros.diretoria)) continue;
      if (filtros.id_oe && String(res.id_oe) !== String(filtros.id_oe)) continue;
      if (filtros.trimestre && String(okr.trimestre) !== String(filtros.trimestre)) continue;
      if (filtros.status && String(atv.status) !== String(filtros.status)) continue;
    }

    resultado.push({
      id: atv.id_atividade, descricao: atv.descricao,
      mes_planejado: atv.mes_planejado, mes_entrega: atv.mes_entrega || '',
      mes_realizado: atv.mes_realizado || '', status: atv.status,
      percentual: atv.percentual_execucao, peso: atv.peso,
      responsavel: atv.responsavel, diretoria: atv.diretoria,
      id_resultado: atv.id_resultado, resultado_descricao: res.descricao || '',
      id_oe: res.id_oe || '', trimestre: okr.trimestre || '',
      id_okr: atv.id_okr, okr_descricao: okr.descricao || ''
    });
  }

  resultado.sort(function(a, b) {
    return String(a.id_resultado).localeCompare(String(b.id_resultado));
  });

  return resultado;
}

// ============ FUNCAO DE TESTE ============

/**
 * Testar a API localmente (executar no editor do Apps Script)
 * Menu: Executar > testarAPI
 */
function testarAPI() {
  var testes = ['ping', 'info', 'listarOEs', 'getDashboardResumo'];
  var resultados = [];

  for (var i = 0; i < testes.length; i++) {
    try {
      var resultado = _rotearAcaoAPI(testes[i], {});
      resultados.push('OK - ' + testes[i] + ': ' + (resultado.success ? 'sucesso' : resultado.error));
    } catch (e) {
      resultados.push('ERRO - ' + testes[i] + ': ' + e.message);
    }
  }

  var msg = '=== TESTE DA API ===\n' + resultados.join('\n');
  Logger.log(msg);
  return msg;
}
