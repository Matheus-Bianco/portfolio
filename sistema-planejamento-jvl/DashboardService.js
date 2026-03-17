/**
 * ============================================================
 * SERVICO DE DASHBOARD - AGREGACAO DE DADOS
 * ============================================================
 */

var CACHE_TTL = 300; // 5 minutos em segundos

/**
 * Derivar percentual de execucao a partir do status da atividade.
 * Concluido/Realizado = 100%, qualquer outro status = 0%.
 */
function _percentualPorStatus(status) {
  var s = String(status || '').toLowerCase();
  if (s === 'realizado' || s === 'concluido' || (s.indexOf('conclu') !== -1 && s.indexOf('nao') === -1 && s.indexOf('parcial') === -1)) return 100;
  return 0;
}

/**
 * Obter chave de cache por usuario (GERENCIA tem dados filtrados por diretoria)
 */
function _getCacheKey(base) {
  var usuario = getUsuarioLogado();
  if (!usuario) return null;
  if (usuario.perfil === 'GERENCIA') {
    return base + '_' + usuario.diretoria;
  }
  return base + '_geral';
}

/**
 * Tentar ler do cache. Retorna null se nao encontrado ou expirado.
 */
function _lerCache(chave) {
  try {
    var cache = CacheService.getScriptCache();
    var raw = cache.get(chave);
    if (raw) return JSON.parse(raw);
  } catch (e) {}
  return null;
}

/**
 * Salvar no cache. Silencioso se falhar (dados > 100KB).
 */
function _salvarCache(chave, dados) {
  try {
    var cache = CacheService.getScriptCache();
    var json = JSON.stringify(dados);
    if (json.length < 100000) {
      cache.put(chave, json, CACHE_TTL);
    }
  } catch (e) {}
}

/**
 * Invalidar caches de dashboard apos escritas.
 * Chamada pelas funcoes de edicao/criacao/exclusao.
 */
function invalidarCacheDashboard() {
  try {
    var cache = CacheService.getScriptCache();
    cache.removeAll([
      'dashboard_geral', 'visao_geral', 'indicadores_geral',
      'dashboard_' + ((getUsuarioLogado() || {}).diretoria || ''),
      'visao_' + ((getUsuarioLogado() || {}).diretoria || ''),
      'indicadores_' + ((getUsuarioLogado() || {}).diretoria || '')
    ]);
  } catch (e) {}
}

/**
 * Obter dados completos do dashboard
 */
function getDashboardData() {
  try {
    var usuario = getUsuarioLogado();
    if (!usuario) return { status: 'erro', mensagem: 'Nao logado' };

    // Tentar cache
    var cacheKey = _getCacheKey('dashboard');
    if (cacheKey) {
      var cached = _lerCache(cacheKey);
      if (cached) return cached;
    }

    var oes = listarOEs();
    var resultados = listarResultados();
    var okrs = listarOKRs();
    var atividades = listarAtividades(null);

    // Se GERENCIA, filtrar por diretoria (ja vem filtrado de listarAtividades)
    var todasAtividades = atividades || [];

    // Calcular metricas
    var totalAtividades = todasAtividades.length;
    var concluidas = 0;
    var emAndamento = 0;
    var atrasadas = 0;
    var noPrazo = 0;
    var replanejadas = 0;
    var naoIniciadas = 0;
    var somaPercentual = 0;
    var listaAndamento = [];
    var listaAtrasadas = [];

    for (var i = 0; i < todasAtividades.length; i++) {
      var status = String(todasAtividades[i].status || '').toLowerCase();
      var pctExec = parseFloat(todasAtividades[i].percentual_execucao) || 0;
      var pct = pctExec > 0 ? pctExec : _percentualPorStatus(todasAtividades[i].status);
      somaPercentual += pct;

      var resumoAtv = {
        id: todasAtividades[i].id_atividade,
        descricao: todasAtividades[i].descricao,
        responsavel: todasAtividades[i].responsavel,
        percentual: pct,
        status: todasAtividades[i].status,
        diretoria: todasAtividades[i].diretoria,
        id_okr: todasAtividades[i].id_okr || '',
        okr_descricao: todasAtividades[i].okr_descricao || ''
      };

      if (status === 'nao iniciada' || status === 'nao iniciado') {
        naoIniciadas++;
      } else if (status === 'realizado' || (status.indexOf('conclu') !== -1 && status.indexOf('nao') === -1 && status.indexOf('parcial') === -1)) {
        concluidas++;
      } else if (status.indexOf('andamento') !== -1) {
        emAndamento++;
        listaAndamento.push(resumoAtv);
      } else if (status.indexOf('atras') !== -1) {
        atrasadas++;
        listaAtrasadas.push(resumoAtv);
      } else if (status === 'sem status' || status === '' || status === 'no prazo') {
        noPrazo++;
      } else if (status.indexOf('replanej') !== -1) {
        replanejadas++;
      }
    }

    var percentualGeral = totalAtividades > 0 ? Math.round(somaPercentual / totalAtividades) : 0;

    // Percentual por OE
    var percentualPorOE = calcularPercentualPorOE(oes || [], todasAtividades, resultados || []);

    // Ranking por diretoria
    var rankingDiretoria = calcularRankingDiretoria(todasAtividades);

    // OKRs atrasados
    var okrsAtrasados = [];
    var okrsList = okrs || [];
    for (var j = 0; j < okrsList.length; j++) {
      var sit = String(okrsList[j].situacao || '').toLowerCase();
      if (sit.indexOf('atras') !== -1 || sit.indexOf('nao conclu') !== -1) {
        okrsAtrasados.push({
          id: okrsList[j].id_okr,
          descricao: okrsList[j].descricao,
          trimestre: okrsList[j].trimestre,
          situacao: okrsList[j].situacao,
          percentual: okrsList[j].percentual
        });
      }
    }

    // Ranking por responsavel
    var rankingResponsavel = calcularRankingResponsavel(todasAtividades);

    // Dados por eixo
    var dadosPorEixo = calcularDadosPorEixo(oes || [], todasAtividades, resultados || []);

    // Distribuicao de status
    var distribuicaoStatus = {
      'Concluido': concluidas,
      'Em andamento': emAndamento,
      'Atrasado': atrasadas,
      'Sem Status': noPrazo,
      'Nao iniciada': naoIniciadas,
      'Replanejado': replanejadas,
      'Outros': totalAtividades - concluidas - emAndamento - atrasadas - noPrazo - replanejadas - naoIniciadas
    };

    var resultado = {
      status: 'ok',
      percentualGeral: percentualGeral,
      totalAtividades: totalAtividades,
      concluidas: concluidas,
      emAndamento: emAndamento,
      atrasadas: atrasadas,
      noPrazo: noPrazo,
      naoIniciadas: naoIniciadas,
      replanejadas: replanejadas,
      percentualPorOE: percentualPorOE || [],
      rankingDiretoria: rankingDiretoria || [],
      rankingResponsavel: rankingResponsavel || [],
      okrsAtrasados: okrsAtrasados,
      dadosPorEixo: dadosPorEixo || [],
      distribuicaoStatus: distribuicaoStatus,
      totalOEs: (oes || []).length,
      totalResultados: (resultados || []).length,
      totalOKRs: (okrs || []).length,
      totalIndicadores: contarIndicadores(),
      listaAndamento: listaAndamento.slice(0, 30),
      listaAtrasadas: listaAtrasadas.slice(0, 30)
    };

    if (cacheKey) _salvarCache(cacheKey, resultado);
    return resultado;
  } catch (e) {
    return { status: 'erro', mensagem: 'Erro ao montar dashboard: ' + e.message };
  }
}

/**
 * Calcular percentual medio de execucao por OE
 */
function calcularPercentualPorOE(oes, atividades, resultados) {
  var resultado = [];

  // Mapear resultado -> OE (fallback para atividades sem id_oe direto)
  var resultadoOE = {};
  for (var r = 0; r < resultados.length; r++) {
    resultadoOE[resultados[r].id_resultado] = resultados[r].id_oe;
  }

  for (var i = 0; i < oes.length; i++) {
    var oe = oes[i];
    var soma = 0;
    var count = 0;

    for (var j = 0; j < atividades.length; j++) {
      var atv = atividades[j];
      var oeAtv = atv.id_oe || resultadoOE[atv.id_resultado];
      if (oeAtv === oe.id_oe) {
        var pctExec = parseFloat(atv.percentual_execucao) || 0;
        soma += pctExec > 0 ? pctExec : _percentualPorStatus(atv.status);
        count++;
      }
    }

    resultado.push({
      id_oe: oe.id_oe,
      descricao: oe.descricao,
      eixo: oe.eixo,
      tipo: oe.tipo,
      percentual: count > 0 ? Math.round(soma / count) : 0,
      totalAtividades: count
    });
  }

  return resultado;
}

/**
 * Calcular ranking por diretoria
 */
function calcularRankingDiretoria(atividades) {
  var diretorias = {};

  for (var i = 0; i < atividades.length; i++) {
    var dir = String(atividades[i].diretoria).trim();
    if (!dir || dir === 'undefined' || dir === '') continue;

    if (!diretorias[dir]) {
      diretorias[dir] = { total: 0, soma: 0 };
    }
    diretorias[dir].total++;
    var pctExec = parseFloat(atividades[i].percentual_execucao) || 0;
    diretorias[dir].soma += pctExec > 0 ? pctExec : _percentualPorStatus(atividades[i].status);
  }

  var ranking = [];
  for (var d in diretorias) {
    ranking.push({
      diretoria: d,
      percentual: Math.round(diretorias[d].soma / diretorias[d].total),
      total: diretorias[d].total
    });
  }

  ranking.sort(function(a, b) { return b.percentual - a.percentual; });
  return ranking;
}

/**
 * Calcular dados agrupados por eixo
 */
function calcularDadosPorEixo(oes, atividades, resultados) {
  var eixos = {};

  // Mapear resultado -> OE
  var resultadoOE = {};
  for (var r = 0; r < resultados.length; r++) {
    resultadoOE[resultados[r].id_resultado] = resultados[r].id_oe;
  }

  // Mapear OE -> eixo
  var oeEixo = {};
  for (var o = 0; o < oes.length; o++) {
    oeEixo[oes[o].id_oe] = oes[o].eixo;
  }

  for (var i = 0; i < atividades.length; i++) {
    var atv = atividades[i];
    var oeId = atv.id_oe || resultadoOE[atv.id_resultado];
    var eixo = atv.eixo || oeEixo[oeId] || 'NAO CLASSIFICADO';

    if (!eixos[eixo]) {
      eixos[eixo] = { total: 0, soma: 0, concluidas: 0, atrasadas: 0 };
    }
    eixos[eixo].total++;
    var pctExec = parseFloat(atv.percentual_execucao) || 0;
    eixos[eixo].soma += pctExec > 0 ? pctExec : _percentualPorStatus(atv.status);

    var status = String(atv.status).toLowerCase();
    if (status === 'realizado' || (status.indexOf('conclu') !== -1 && status.indexOf('nao') === -1)) {
      eixos[eixo].concluidas++;
    }
    if (status.indexOf('atras') !== -1) {
      eixos[eixo].atrasadas++;
    }
  }

  var resultado = [];
  for (var e in eixos) {
    resultado.push({
      eixo: e,
      percentual: eixos[e].total > 0 ? Math.round(eixos[e].soma / eixos[e].total) : 0,
      total: eixos[e].total,
      concluidas: eixos[e].concluidas,
      atrasadas: eixos[e].atrasadas
    });
  }

  return resultado;
}

/**
 * Calcular ranking por responsavel com detalhamento de atividades
 */
function calcularRankingResponsavel(atividades) {
  var responsaveis = {};

  for (var i = 0; i < atividades.length; i++) {
    var resp = String(atividades[i].responsavel || '').trim();
    if (!resp || resp === 'undefined' || resp === '') continue;

    if (!responsaveis[resp]) {
      responsaveis[resp] = { total: 0, soma: 0, concluidas: 0, atrasadas: 0, emAndamento: 0, diretoria: '', atividades: [] };
    }
    responsaveis[resp].total++;
    var pctExec = parseFloat(atividades[i].percentual_execucao) || 0;
    responsaveis[resp].soma += pctExec > 0 ? pctExec : _percentualPorStatus(atividades[i].status);
    responsaveis[resp].diretoria = atividades[i].diretoria || responsaveis[resp].diretoria;

    var status = String(atividades[i].status || '').toLowerCase();
    if (status === 'realizado' || (status.indexOf('conclu') !== -1 && status.indexOf('nao') === -1 && status.indexOf('parcial') === -1)) {
      responsaveis[resp].concluidas++;
    } else if (status.indexOf('atras') !== -1) {
      responsaveis[resp].atrasadas++;
    } else if (status.indexOf('andamento') !== -1) {
      responsaveis[resp].emAndamento++;
    }

    responsaveis[resp].atividades.push({
      id: atividades[i].id_atividade,
      descricao: atividades[i].descricao,
      status: atividades[i].status,
      percentual: atividades[i].percentual_execucao,
      mes: atividades[i].mes_planejado
    });
  }

  var ranking = [];
  for (var r in responsaveis) {
    ranking.push({
      responsavel: r,
      percentual: responsaveis[r].total > 0 ? Math.round(responsaveis[r].soma / responsaveis[r].total) : 0,
      total: responsaveis[r].total,
      concluidas: responsaveis[r].concluidas,
      atrasadas: responsaveis[r].atrasadas,
      emAndamento: responsaveis[r].emAndamento,
      diretoria: responsaveis[r].diretoria,
      atividades: responsaveis[r].atividades
    });
  }

  ranking.sort(function(a, b) { return b.total - a.total; });
  return ranking;
}

/**
 * Obter dados para graficos do dashboard
 */
function getDadosGraficos() {
  var dashData = getDashboardData();
  if (dashData.status === 'erro') return dashData;

  return {
    status: 'ok',
    graficoExecucaoOE: dashData.percentualPorOE.map(function(item) {
      return { label: item.id_oe, valor: item.percentual };
    }),
    graficoStatus: dashData.distribuicaoStatus,
    graficoEixos: dashData.dadosPorEixo.map(function(item) {
      return { label: item.eixo, valor: item.percentual };
    }),
    graficoRanking: dashData.rankingDiretoria.slice(0, 10).map(function(item) {
      return { label: item.diretoria, valor: item.percentual };
    })
  };
}

/**
 * Obter visao hierarquica completa: OE -> Resultado -> OKR -> Atividades
 */
function getVisaoHierarquica(filtros) {
  var oes = listarOEs();
  var resultados = listarResultados();
  var okrs = listarOKRs();
  var atividades = listarAtividades(null);

  // Aplicar filtros
  if (filtros) {
    if (filtros.eixo) {
      oes = oes.filter(function(oe) { return oe.eixo === filtros.eixo; });
    }
    if (filtros.id_oe) {
      oes = oes.filter(function(oe) { return oe.id_oe === filtros.id_oe; });
    }
  }

  var hierarquia = [];

  for (var i = 0; i < oes.length; i++) {
    var oe = oes[i];
    var resultadosOE = resultados.filter(function(r) { return r.id_oe === oe.id_oe; });

    var oeNode = {
      oe: oe,
      resultados: []
    };

    for (var j = 0; j < resultadosOE.length; j++) {
      var res = resultadosOE[j];
      var okrsRes = okrs.filter(function(o) { return o.id_resultado === res.id_resultado; });

      var resNode = {
        resultado: res,
        okrs: []
      };

      for (var k = 0; k < okrsRes.length; k++) {
        var okr = okrsRes[k];
        var atvsOKR = atividades.filter(function(a) { return a.id_okr === okr.id_okr; });

        resNode.okrs.push({
          okr: okr,
          atividades: atvsOKR
        });
      }

      oeNode.resultados.push(resNode);
    }

    hierarquia.push(oeNode);
  }

  return hierarquia;
}

/**
 * Obter visao estrategica completa agrupada por EIXO com 4 niveis de cascata
 * Eixo -> OE -> Resultado (com Indicadores) -> OKR -> Atividades
 */
function getVisaoEstrategica(filtros) {
  try {
  // Cache apenas se sem filtros (chamada padrao)
  var cacheKeyVisao = null;
  if (!filtros) {
    cacheKeyVisao = _getCacheKey('visao');
    if (cacheKeyVisao) {
      var cachedVisao = _lerCache(cacheKeyVisao);
      if (cachedVisao) return cachedVisao;
    }
  }

  var oes = listarOEs();
  var resultados = listarResultados();
  var okrs = listarOKRs();
  var atividades = listarAtividades(null);
  var indicadores = listarIndicadores();

  // Sanitizar Dates em atividades (google.script.run falha com objetos Date)
  for (var si = 0; si < atividades.length; si++) {
    var a = atividades[si];
    for (var key in a) {
      if (a[key] instanceof Date) a[key] = a[key].toISOString();
    }
  }
  for (var sri = 0; sri < resultados.length; sri++) {
    var r = resultados[sri];
    for (var rkey in r) {
      if (r[rkey] instanceof Date) r[rkey] = r[rkey].toISOString();
    }
  }
  for (var soi = 0; soi < okrs.length; soi++) {
    var ok = okrs[soi];
    for (var okey in ok) {
      if (ok[okey] instanceof Date) ok[okey] = ok[okey].toISOString();
    }
  }

  // Indexar indicadores diretamente por OE (id_oe da aba INDICADORES)
  var indicadoresPorOE = {};
  for (var ii = 0; ii < indicadores.length; ii++) {
    var indOE = String(indicadores[ii].id_oe || '').trim();
    // Fallback: extrair OE do cod (IND.OE09.UIF.1 -> OE09)
    if (!indOE) {
      var indCod = String(indicadores[ii].cod || '');
      var indParts = indCod.replace(/[-_]/g, '.').split('.');
      if (indParts.length >= 2) indOE = indParts[1].toUpperCase();
    }
    if (!indOE) continue;
    if (!indicadoresPorOE[indOE]) indicadoresPorOE[indOE] = [];
    indicadoresPorOE[indOE].push(indicadores[ii]);
  }

  // Indexar resultados por OE
  var resultadosPorOE = {};
  for (var ri = 0; ri < resultados.length; ri++) {
    var oeId = resultados[ri].id_oe;
    if (!resultadosPorOE[oeId]) resultadosPorOE[oeId] = [];
    resultadosPorOE[oeId].push(resultados[ri]);
  }

  // Indexar OKRs por resultado
  var okrsPorResultado = {};
  for (var oi = 0; oi < okrs.length; oi++) {
    var idRes = okrs[oi].id_resultado;
    if (!okrsPorResultado[idRes]) okrsPorResultado[idRes] = [];
    okrsPorResultado[idRes].push(okrs[oi]);
  }

  // Indexar atividades por OKR
  var atvsPorOKR = {};
  for (var ai = 0; ai < atividades.length; ai++) {
    var idOkr = atividades[ai].id_okr;
    if (!atvsPorOKR[idOkr]) atvsPorOKR[idOkr] = [];
    atvsPorOKR[idOkr].push(atividades[ai]);
  }

  // Configuracao dos eixos
  var eixosConfig = [
    { nome: 'ALUNOS', classe: 'alunos' },
    { nome: 'PROFESSORES', classe: 'professores' },
    { nome: 'GESTAO ESCOLAR', classe: 'gestao' },
    { nome: 'CAPACIDADE INSTITUCIONAL', classe: 'institucional' }
  ];

  var eixosResult = [];
  var objetivosImpacto = [];

  for (var ei = 0; ei < eixosConfig.length; ei++) {
    var cfg = eixosConfig[ei];
    var oesEixo = oes.filter(function(o) { return o.eixo === cfg.nome && o.tipo !== 'IMPACTO'; });

    var eixoSoma = 0, eixoCount = 0;
    var eixoOEs = [];

    for (var oi2 = 0; oi2 < oesEixo.length; oi2++) {
      var oe = oesEixo[oi2];
      var resOE = resultadosPorOE[oe.id_oe] || [];

      // Indicadores vinculados diretamente ao OE
      var oeIndicadores = indicadoresPorOE[oe.id_oe] || [];

      var oeSoma = 0, oeCount = 0;
      var oeResultados = [];

      for (var ri2 = 0; ri2 < resOE.length; ri2++) {
        var res = resOE[ri2];
        var okrsRes = okrsPorResultado[res.id_resultado] || [];
        var resSoma = 0, resCount = 0;

        var resOKRs = [];
        for (var ki = 0; ki < okrsRes.length; ki++) {
          var okr = okrsRes[ki];
          var atvsOKR = atvsPorOKR[okr.id_okr] || [];

          for (var ati = 0; ati < atvsOKR.length; ati++) {
            var pctAtv = parseFloat(atvsOKR[ati].percentual_execucao) || 0;
            resSoma += pctAtv;
            resCount++;
          }

          resOKRs.push({ okr: okr, atividades: atvsOKR });
        }

        var resPct = resCount > 0 ? Math.round(resSoma / resCount) : 0;
        oeSoma += resSoma;
        oeCount += resCount;

        oeResultados.push({
          resultado: res,
          percentual: resPct,
          okrs: resOKRs
        });
      }

      var oePct = oeCount > 0 ? Math.round(oeSoma / oeCount) : 0;
      eixoSoma += oeSoma;
      eixoCount += oeCount;

      eixoOEs.push({
        oe: oe,
        percentual: oePct,
        totalIndicadores: oeIndicadores.length,
        indicadores: oeIndicadores,
        resultados: oeResultados
      });
    }

    var eixoPct = eixoCount > 0 ? Math.round(eixoSoma / eixoCount) : 0;

    eixosResult.push({
      nome: cfg.nome,
      classe: cfg.classe,
      percentual: eixoPct,
      oes: eixoOEs
    });
  }

  // Objetivos de Impacto
  var oisEntries = oes.filter(function(o) { return o.tipo === 'IMPACTO'; });
  for (var oii = 0; oii < oisEntries.length; oii++) {
    objetivosImpacto.push({ oe: oisEntries[oii] });
  }

  // Coletar responsaveis e diretorias unicos para filtros (de TODAS as entidades)
  var responsaveisSet = {};
  var diretoriasSet = {};
  // De Resultados
  for (var fri = 0; fri < resultados.length; fri++) {
    var fResp = String(resultados[fri].responsavel || '').trim();
    var fDir = String(resultados[fri].diretoria || '').trim();
    if (fResp && fResp !== 'undefined') responsaveisSet[fResp] = true;
    if (fDir && fDir !== 'undefined') diretoriasSet[fDir] = true;
  }
  // De Atividades
  for (var fai = 0; fai < atividades.length; fai++) {
    var faResp = String(atividades[fai].responsavel || '').trim();
    var faDir = String(atividades[fai].diretoria || '').trim();
    if (faResp && faResp !== 'undefined') responsaveisSet[faResp] = true;
    if (faDir && faDir !== 'undefined') diretoriasSet[faDir] = true;
  }
  // De Indicadores
  for (var fii = 0; fii < indicadores.length; fii++) {
    var fiResp = String(indicadores[fii].responsavel || '').trim();
    if (fiResp && fiResp !== 'undefined') responsaveisSet[fiResp] = true;
  }

  var resultadoVisao = {
    eixos: eixosResult,
    objetivosImpacto: objetivosImpacto,
    responsaveis: Object.keys(responsaveisSet).sort(),
    diretorias: Object.keys(diretoriasSet).sort(),
    totais: {
      oes: oes.filter(function(o) { return o.tipo !== 'IMPACTO'; }).length,
      resultados: resultados.length,
      okrs: okrs.length,
      atividades: atividades.length,
      indicadores: indicadores.length
    }
  };

  if (cacheKeyVisao) _salvarCache(cacheKeyVisao, resultadoVisao);
  return resultadoVisao;
  } catch (e) {
    Logger.log('ERRO getVisaoEstrategica: ' + e.message + ' | Stack: ' + e.stack);
    return { eixos: [], objetivosImpacto: [], erro: e.message };
  }
}

/**
 * Obter resultados agrupados por OE com OKRs vinculados por resultado
 */
/**
 * Gerar HTML standalone da Mandala Estrategica para exportacao
 * Usa o template MandalaExport.html e injeta dados reais
 */
function gerarMandalaHTML() {
  var data = getVisaoEstrategica();
  var dataJSON = JSON.stringify(data);
  var hoje = new Date();
  var dataStr = Utilities.formatDate(hoje, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');

  var template = HtmlService.createHtmlOutputFromFile('MandalaExport').getContent();
  var html = template
    .replace('%%DATA_JSON%%', dataJSON)
    .replace('%%DATA_GERACAO%%', dataStr);
  return html;
}


function getResultadosAgrupados() {
  var oes = listarOEs();
  var resultados = listarResultados();
  var okrs = listarOKRs();

  // Indexar OKRs por id_resultado
  var okrsPorResultado = {};
  for (var k = 0; k < okrs.length; k++) {
    var idRes = okrs[k].id_resultado;
    if (!okrsPorResultado[idRes]) okrsPorResultado[idRes] = [];
    okrsPorResultado[idRes].push(okrs[k]);
  }

  var totalResultados = resultados.length;
  var totalOKRsVinculados = okrs.length;

  var grupos = [];
  var oesEstrategicos = oes.filter(function(o) { return o.tipo === 'ESTRATEGICO'; });

  for (var i = 0; i < oesEstrategicos.length; i++) {
    var oe = oesEstrategicos[i];
    var resOE = resultados.filter(function(r) { return r.id_oe === oe.id_oe; });
    resOE = resOE.map(function(r) {
      var okrsRes = okrsPorResultado[r.id_resultado] || [];
      var somaPct = 0;
      var concluidos = 0;
      for (var j = 0; j < okrsRes.length; j++) {
        somaPct += parseFloat(okrsRes[j].percentual) || 0;
        var sit = String(okrsRes[j].situacao || '').toLowerCase();
        if (sit.indexOf('conclu') !== -1 && sit.indexOf('nao') === -1 && sit.indexOf('parcial') === -1) concluidos++;
      }
      r.okrs = okrsRes;
      r.percentual_medio = okrsRes.length > 0 ? Math.round(somaPct / okrsRes.length) : 0;
      r.okrs_concluidos = concluidos;
      r.meta_2026 = r.meta_2026 || '';
      r.requisitos_qualidade = r.requisitos_qualidade || '';
      return r;
    });
    grupos.push({ oe: oe, resultados: resOE });
  }

  return { grupos: grupos, totalResultados: totalResultados, totalOKRsVinculados: totalOKRsVinculados };
}

/**
 * Obter OKRs agrupados por trimestre com metricas
 */
function getOKRsAgrupados() {
  var okrs = listarOKRs();
  var totalOKRs = okrs.length;
  var concluidos = 0;
  var emAndamento = 0;
  var atrasados = 0;
  var somaPercentual = 0;

  for (var i = 0; i < okrs.length; i++) {
    var sit = String(okrs[i].situacao || '').toLowerCase();
    somaPercentual += parseFloat(okrs[i].percentual) || 0;
    if (sit.indexOf('conclu') !== -1 && sit.indexOf('nao') === -1 && sit.indexOf('parcial') === -1) {
      concluidos++;
    } else if (sit.indexOf('andamento') !== -1 || sit.indexOf('prazo') !== -1) {
      emAndamento++;
    } else if (sit.indexOf('atras') !== -1 || sit.indexOf('nao conclu') !== -1) {
      atrasados++;
    }
  }

  var trimestres = ['1o', '2o', '3o', '4o'];
  var grupos = [];
  for (var t = 0; t < trimestres.length; t++) {
    var tri = trimestres[t];
    var okrsTri = okrs.filter(function(o) { return o.trimestre === tri; });
    if (okrsTri.length > 0) {
      grupos.push({ trimestre: tri, okrs: okrsTri });
    }
  }

  // OKRs sem trimestre definido
  var semTri = okrs.filter(function(o) {
    return trimestres.indexOf(o.trimestre) === -1;
  });
  if (semTri.length > 0) {
    grupos.push({ trimestre: 'Outros', okrs: semTri });
  }

  // Tambem agrupar por resultado
  var resultados = listarResultados();
  var resMap = {};
  for (var ri = 0; ri < resultados.length; ri++) {
    resMap[resultados[ri].id_resultado] = resultados[ri];
  }

  var porResultado = {};
  var ordemResultado = [];
  for (var oi = 0; oi < okrs.length; oi++) {
    var idRes = okrs[oi].id_resultado || 'sem_resultado';
    if (!porResultado[idRes]) {
      var resInfo = resMap[idRes] || {};
      porResultado[idRes] = {
        id_resultado: idRes,
        descricao_resultado: resInfo.descricao || '',
        diretoria: resInfo.diretoria || '',
        okrs: []
      };
      ordemResultado.push(idRes);
    }
    porResultado[idRes].okrs.push(okrs[oi]);
  }

  var gruposPorResultado = [];
  for (var gi = 0; gi < ordemResultado.length; gi++) {
    gruposPorResultado.push(porResultado[ordemResultado[gi]]);
  }

  return {
    grupos: grupos,
    gruposPorResultado: gruposPorResultado,
    totalOKRs: totalOKRs,
    concluidos: concluidos,
    emAndamento: emAndamento,
    atrasados: atrasados,
    percentualMedio: totalOKRs > 0 ? Math.round(somaPercentual / totalOKRs) : 0
  };
}

/**
 * Obter atividades agrupadas por diretoria com metricas
 * Tambem agrupa por OKR dentro de cada diretoria para visao hierarquica
 */
function getAtividadesAgrupadas() {
  var atividades = listarAtividades(null);
  var okrs = listarOKRs();
  var total = atividades.length;
  var concluidas = 0;
  var emAndamento = 0;
  var atrasadas = 0;
  var somaPercentual = 0;

  // Mapa de OKRs para enriquecimento
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

    if (st.indexOf('conclu') !== -1 && st.indexOf('nao') === -1 && st.indexOf('parcial') === -1) {
      concluidas++;
    } else if (st.indexOf('andamento') !== -1) {
      emAndamento++;
    } else if (st.indexOf('atras') !== -1) {
      atrasadas++;
    }

    var dir = String(atv.diretoria || '').trim() || 'Sem diretoria';
    if (!dirMap[dir]) dirMap[dir] = [];

    // Enriquecer atividade com descricao do OKR
    var okrInfo = okrMap[atv.id_okr];
    atv.okr_descricao = okrInfo ? okrInfo.descricao : '';
    atv.okr_trimestre = okrInfo ? okrInfo.trimestre : '';

    dirMap[dir].push(atv);
  }

  var grupos = [];
  var dirs = Object.keys(dirMap).sort();
  for (var d = 0; d < dirs.length; d++) {
    var atvDir = dirMap[dirs[d]];

    // Agrupar por OKR dentro da diretoria
    var okrGrupos = {};
    var okrOrdem = [];
    for (var a = 0; a < atvDir.length; a++) {
      var idOkr = atvDir[a].id_okr || 'sem_okr';
      if (!okrGrupos[idOkr]) {
        okrGrupos[idOkr] = {
          id_okr: idOkr,
          descricao: atvDir[a].okr_descricao || '',
          trimestre: atvDir[a].okr_trimestre || '',
          atividades: []
        };
        okrOrdem.push(idOkr);
      }
      okrGrupos[idOkr].atividades.push(atvDir[a]);
    }

    var okrsList = [];
    for (var o = 0; o < okrOrdem.length; o++) {
      okrsList.push(okrGrupos[okrOrdem[o]]);
    }

    grupos.push({ diretoria: dirs[d], atividades: atvDir, okrGrupos: okrsList });
  }

  return {
    grupos: grupos,
    total: total,
    concluidas: concluidas,
    emAndamento: emAndamento,
    atrasadas: atrasadas,
    percentualMedio: total > 0 ? Math.round(somaPercentual / total) : 0
  };
}

/**
 * Contar total de indicadores no banco
 */
function contarIndicadores() {
  try {
    var sheet = getSheet(CONFIG.ABAS.INDICADORES);
    return Math.max(sheet.getLastRow() - 1, 0);
  } catch (e) {
    return 0;
  }
}

/**
 * Obter indicadores agrupados por OE com contagem de dados aferidos
 */
function getIndicadoresAgrupados() {
  var oes = listarOEs();
  var indicadores = listarIndicadores();

  // Contar dados preenchidos por indicador (colunas jan-dez)
  var mesesCampos = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
  var totalIndicadores = indicadores.length;
  var comDados = 0;
  for (var i = 0; i < indicadores.length; i++) {
    var qtd = 0;
    for (var m = 0; m < mesesCampos.length; m++) {
      var v = indicadores[i][mesesCampos[m]];
      if (v !== undefined && v !== null && String(v).trim() !== '') qtd++;
    }
    indicadores[i].qtd_dados = qtd;
    if (qtd > 0) comDados++;
  }

  var grupos = [];
  var oesEstrategicos = oes.filter(function(o) { return o.tipo === 'ESTRATEGICO'; });

  for (var j = 0; j < oesEstrategicos.length; j++) {
    var oe = oesEstrategicos[j];
    var indsOE = indicadores.filter(function(ind) { return ind.id_oe === oe.id_oe; });
    if (indsOE.length > 0) {
      grupos.push({ oe: oe, indicadores: indsOE });
    }
  }

  return { grupos: grupos, totalIndicadores: totalIndicadores, comDados: comDados };
}

/**
 * Obter dados para a Arvore de Atividades.
 * Retorna atividades enriquecidas com descricoes de OKR e Resultado.
 */
function getArvoreData() {
  try {
    var usuario = getUsuarioLogado();
    if (!usuario) return { status: 'erro', mensagem: 'Nao logado' };

    var atividades = listarAtividades(null);
    var resultados = listarResultados();
    var okrs = listarOKRs();

    // Mapas de descricao
    var resMap = {};
    for (var r = 0; r < resultados.length; r++) {
      resMap[resultados[r].id_resultado] = resultados[r].descricao || '';
    }
    var okrMap = {};
    for (var o = 0; o < okrs.length; o++) {
      okrMap[okrs[o].id_okr] = {
        descricao: okrs[o].descricao || '',
        trimestre: okrs[o].trimestre || '',
        situacao: okrs[o].situacao || ''
      };
    }

    // Enriquecer atividades
    var atvEnriquecidas = [];
    for (var i = 0; i < atividades.length; i++) {
      var atv = atividades[i];
      var okrInfo = okrMap[atv.id_okr] || {};
      atvEnriquecidas.push({
        id_atividade: atv.id_atividade,
        descricao: atv.descricao,
        status: atv.status,
        responsavel: atv.responsavel,
        mes_planejado: atv.mes_planejado,
        mes_entrega: atv.mes_entrega,
        diretoria: atv.diretoria || 'SEM DIRETORIA',
        unidade: atv.unidade || 'SEM UNIDADE',
        id_resultado: atv.id_resultado || '',
        resultado_descricao: resMap[atv.id_resultado] || '',
        id_okr: atv.id_okr || '',
        okr_descricao: okrInfo.descricao || '',
        okr_trimestre: okrInfo.trimestre || '',
        eixo: atv.eixo || ''
      });
    }

    return { status: 'ok', atividades: atvEnriquecidas };
  } catch (e) {
    return { status: 'erro', mensagem: 'Erro na arvore: ' + e.message };
  }
}
