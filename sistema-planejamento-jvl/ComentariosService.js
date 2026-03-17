/**
 * ============================================================
 * SERVICO DE COMENTARIOS E SUGESTOES
 * ============================================================
 * Permite que usuarios registrem comentarios e sugestoes
 * que ficam armazenados na aba COMENTARIOS da planilha.
 */

// ==================== CRUD COMENTARIOS ====================

/**
 * Criar novo comentario/sugestao
 * @param {string} secao - Secao de origem (Dashboard, Mandala, Indicadores, etc.)
 * @param {string} texto - Texto do comentario
 * @return {Object} Resultado da operacao
 */
function criarComentario(secao, texto) {
  var usuario = getUsuarioLogado();
  if (!usuario) return { status: 'erro', mensagem: 'Nao logado' };
  if (!texto || !texto.trim()) return { status: 'erro', mensagem: 'Texto do comentario e obrigatorio' };

  try {
    var sheet = getSheet(CONFIG.ABAS.COMENTARIOS);
    var id = gerarProximoId(sheet, 'CMT', 4);
    var timestamp = new Date();

    sheet.appendRow([
      id,
      timestamp,
      usuario.email,
      usuario.nome,
      usuario.perfil,
      secao || 'Geral',
      texto.trim(),
      'Novo',       // status_dev: Novo, Em analise, Implementado, Descartado, Em desenvolvimento
      '',            // resposta_dev
      ''             // data_resposta
    ]);

    registrarLog('', id, 'COMENTARIO', 'criacao', '', texto.trim().substring(0, 100));
    return { status: 'ok', mensagem: 'Comentario registrado com sucesso!', id: id };
  } catch (e) {
    return { status: 'erro', mensagem: 'Erro ao salvar comentario: ' + e.message };
  }
}

/**
 * Listar todos os comentarios (para administracao)
 * @param {Object} filtros - Filtros opcionais { secao, status_dev, autor }
 * @return {Array} Lista de comentarios
 */
function listarComentarios(filtros) {
  try {
    var sheet = getSheet(CONFIG.ABAS.COMENTARIOS);
    var dados = sheet.getDataRange().getValues();
    var resultado = [];

    for (var i = 1; i < dados.length; i++) {
      if (!dados[i][0]) continue;

      var comentario = {
        id_comentario: dados[i][0],
        timestamp: dados[i][1],
        autor_email: dados[i][2],
        autor_nome: dados[i][3],
        perfil: dados[i][4],
        secao: dados[i][5],
        texto: dados[i][6],
        status_dev: dados[i][7] || 'Novo',
        resposta_dev: dados[i][8] || '',
        data_resposta: dados[i][9] || ''
      };

      // Sanitizar Dates
      if (comentario.timestamp instanceof Date) comentario.timestamp = comentario.timestamp.toISOString();
      if (comentario.data_resposta instanceof Date) comentario.data_resposta = comentario.data_resposta.toISOString();

      // Aplicar filtros
      if (filtros) {
        if (filtros.secao && String(comentario.secao) !== String(filtros.secao)) continue;
        if (filtros.status_dev && String(comentario.status_dev) !== String(filtros.status_dev)) continue;
        if (filtros.autor && String(comentario.autor_email).toLowerCase().indexOf(String(filtros.autor).toLowerCase()) === -1) continue;
      }

      resultado.push(comentario);
    }

    resultado.reverse(); // Mais recentes primeiro
    return resultado;
  } catch (e) {
    return [];
  }
}

/**
 * Atualizar status de um comentario (para administracao/desenvolvedor)
 * @param {string} idComentario - ID do comentario
 * @param {string} statusDev - Novo status (Novo, Em analise, Em desenvolvimento, Implementado, Descartado)
 * @param {string} respostaDev - Resposta do desenvolvedor (opcional)
 */
function atualizarStatusComentario(idComentario, statusDev, respostaDev) {
  var usuario = getUsuarioLogado();
  if (!usuario) return { status: 'erro', mensagem: 'Nao logado' };
  if (usuario.perfil !== 'PLANEJAMENTO') return { status: 'erro', mensagem: 'Sem permissao' };

  try {
    var sheet = getSheet(CONFIG.ABAS.COMENTARIOS);
    var dados = sheet.getDataRange().getValues();

    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(idComentario)) {
        var statusAntigo = dados[i][7];
        sheet.getRange(i + 1, 8).setValue(statusDev);
        if (respostaDev !== undefined) {
          sheet.getRange(i + 1, 9).setValue(respostaDev);
          sheet.getRange(i + 1, 10).setValue(new Date());
        }
        registrarLog('', idComentario, 'COMENTARIO', 'status_dev', String(statusAntigo), statusDev);
        return { status: 'ok', mensagem: 'Status atualizado' };
      }
    }
    return { status: 'erro', mensagem: 'Comentario nao encontrado' };
  } catch (e) {
    return { status: 'erro', mensagem: e.message };
  }
}

/**
 * Excluir comentario
 */
function excluirComentario(idComentario) {
  var usuario = getUsuarioLogado();
  if (!usuario) return { status: 'erro', mensagem: 'Nao logado' };
  if (usuario.perfil !== 'PLANEJAMENTO') return { status: 'erro', mensagem: 'Sem permissao' };

  try {
    var sheet = getSheet(CONFIG.ABAS.COMENTARIOS);
    var dados = sheet.getDataRange().getValues();

    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]) === String(idComentario)) {
        registrarLog('', idComentario, 'COMENTARIO', 'exclusao', dados[i][6], '');
        sheet.deleteRow(i + 1);
        return { status: 'ok', mensagem: 'Comentario excluido' };
      }
    }
    return { status: 'erro', mensagem: 'Comentario nao encontrado' };
  } catch (e) {
    return { status: 'erro', mensagem: e.message };
  }
}

// ==================== INDICADORES PRIORITARIOS ====================

/**
 * Salvar lista de indicadores prioritarios do usuario
 * @param {Array} listaIds - Array de IDs de indicadores marcados como prioritarios
 */
function salvarIndicadoresPrioritarios(listaIds) {
  var usuario = getUsuarioLogado();
  if (!usuario) return { status: 'erro', mensagem: 'Nao logado' };

  var props = PropertiesService.getScriptProperties();
  var chave = 'indicadores_prioritarios_' + usuario.email;
  props.setProperty(chave, JSON.stringify(listaIds || []));
  return { status: 'ok', mensagem: 'Indicadores prioritarios salvos' };
}

/**
 * Obter indicadores prioritarios do usuario
 */
function getIndicadoresPrioritarios() {
  var usuario = getUsuarioLogado();
  if (!usuario) return [];

  var props = PropertiesService.getScriptProperties();
  var chave = 'indicadores_prioritarios_' + usuario.email;
  var raw = props.getProperty(chave);
  if (raw) {
    try { return JSON.parse(raw); } catch(e) {}
  }
  return [];
}

/**
 * Obter dados resumidos dos indicadores para dashboard
 */
function getDashboardIndicadores() {
  try {
    // Tentar cache
    var cacheKeyInd = _getCacheKey('indicadores');
    if (cacheKeyInd) {
      var cachedInd = _lerCache(cacheKeyInd);
      if (cachedInd) return cachedInd;
    }

    var indicadores = listarIndicadores();
    var mesesCampos = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];

    var totalIndicadores = indicadores.length;
    var comDados = 0;
    var comMeta = 0;
    var atingiramMeta = 0;
    var porOE = {};
    var porPeriodicidade = {};
    var porPolaridade = {};
    var ultimoMesComDados = -1;

    for (var i = 0; i < indicadores.length; i++) {
      var ind = indicadores[i];
      var qtdMeses = 0;
      var ultimoValor = null;

      for (var m = 0; m < mesesCampos.length; m++) {
        var v = ind[mesesCampos[m]];
        if (v !== undefined && v !== null && String(v).trim() !== '') {
          qtdMeses++;
          ultimoValor = v;
          if (m > ultimoMesComDados) ultimoMesComDados = m;
        }
      }

      ind.qtd_dados = qtdMeses;
      ind.ultimo_valor = ultimoValor;

      if (qtdMeses > 0) comDados++;
      if (ind.meta_2026 && String(ind.meta_2026).trim() !== '') comMeta++;

      // Verificar atingimento
      var ating = String(ind.atingimento_meta || '').toLowerCase();
      if (ating.indexOf('sim') !== -1 || ating.indexOf('atingi') !== -1) atingiramMeta++;

      // Agrupar por OE
      var oe = String(ind.id_oe || '').trim();
      if (oe) {
        if (!porOE[oe]) porOE[oe] = { total: 0, comDados: 0 };
        porOE[oe].total++;
        if (qtdMeses > 0) porOE[oe].comDados++;
      }

      // Agrupar por periodicidade
      var per = String(ind.periodicidade || 'Nao definida').trim();
      porPeriodicidade[per] = (porPeriodicidade[per] || 0) + 1;

      // Agrupar por polaridade
      var pol = String(ind.polaridade || 'Nao definida').trim();
      porPolaridade[pol] = (porPolaridade[pol] || 0) + 1;
    }

    // Evolucao mensal: quantos indicadores tem dados em cada mes
    var evolucaoMensal = [];
    for (var em = 0; em < mesesCampos.length; em++) {
      var count = 0;
      for (var ei = 0; ei < indicadores.length; ei++) {
        var val = indicadores[ei][mesesCampos[em]];
        if (val !== undefined && val !== null && String(val).trim() !== '') count++;
      }
      evolucaoMensal.push({ mes: mesesCampos[em], count: count });
    }

    var resultadoInd = {
      totalIndicadores: totalIndicadores,
      comDados: comDados,
      semDados: totalIndicadores - comDados,
      comMeta: comMeta,
      atingiramMeta: atingiramMeta,
      porOE: porOE,
      porPeriodicidade: porPeriodicidade,
      porPolaridade: porPolaridade,
      evolucaoMensal: evolucaoMensal,
      ultimoMesComDados: ultimoMesComDados >= 0 ? mesesCampos[ultimoMesComDados] : 'nenhum',
      indicadores: indicadores
    };

    if (cacheKeyInd) _salvarCache(cacheKeyInd, resultadoInd);
    return resultadoInd;
  } catch (e) {
    return { totalIndicadores: 0, erro: e.message };
  }
}

// ==================== APRESENTACAO EXECUTIVA ====================

/**
 * Gerar dados para apresentacao executiva mensal
 */
function gerarDadosApresentacao() {
  try {
    var oes = listarOEs();
    var resultados = listarResultados();
    var okrs = listarOKRs();
    var atividades = listarAtividades(null);
    var indicadores = listarIndicadores();

    // Sanitizar datas
    var todas = [].concat(atividades, resultados, okrs);
    for (var si = 0; si < todas.length; si++) {
      for (var key in todas[si]) {
        if (todas[si][key] instanceof Date) todas[si][key] = todas[si][key].toISOString();
      }
    }

    // Metricas gerais
    var totalAtv = atividades.length;
    var concluidas = 0, emAndamento = 0, atrasadas = 0, noPrazo = 0, replanejadas = 0, naoIniciadas = 0;
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

    var percentualGeral = totalAtv > 0 ? Math.round(somaPercentual / totalAtv) : 0;

    // Percentual por eixo
    var oeEixo = {};
    for (var o = 0; o < oes.length; o++) {
      oeEixo[oes[o].id_oe] = oes[o].eixo;
    }
    var resOE = {};
    for (var r = 0; r < resultados.length; r++) {
      resOE[resultados[r].id_resultado] = resultados[r].id_oe;
    }

    var eixos = {};
    for (var a = 0; a < atividades.length; a++) {
      var oeId = atividades[a].id_oe || resOE[atividades[a].id_resultado];
      var eixo = atividades[a].eixo || oeEixo[oeId] || 'Outros';
      if (!eixos[eixo]) eixos[eixo] = { total: 0, soma: 0, concluidas: 0, atrasadas: 0 };
      eixos[eixo].total++;
      eixos[eixo].soma += parseFloat(atividades[a].percentual_execucao) || 0;
      var st = String(atividades[a].status || '').toLowerCase();
      if (st.indexOf('conclu') !== -1 && st.indexOf('nao') === -1) eixos[eixo].concluidas++;
      if (st.indexOf('atras') !== -1) eixos[eixo].atrasadas++;
    }

    var dadosPorEixo = [];
    for (var e in eixos) {
      dadosPorEixo.push({
        eixo: e,
        percentual: eixos[e].total > 0 ? Math.round(eixos[e].soma / eixos[e].total) : 0,
        total: eixos[e].total,
        concluidas: eixos[e].concluidas,
        atrasadas: eixos[e].atrasadas
      });
    }

    // Top OKRs atrasados
    var okrsAtrasados = [];
    for (var k = 0; k < okrs.length; k++) {
      var sit = String(okrs[k].situacao || '').toLowerCase();
      if (sit.indexOf('atras') !== -1 || sit.indexOf('nao conclu') !== -1) {
        okrsAtrasados.push(okrs[k]);
      }
    }

    // Ranking diretorias
    var ranking = calcularRankingDiretoria(atividades);

    // Indicadores resumo
    var indComDados = 0;
    var mesesCampos = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
    for (var ii = 0; ii < indicadores.length; ii++) {
      for (var mm = 0; mm < mesesCampos.length; mm++) {
        var vv = indicadores[ii][mesesCampos[mm]];
        if (vv !== undefined && vv !== null && String(vv).trim() !== '') { indComDados++; break; }
      }
    }

    var hoje = new Date();
    var mesAtual = hoje.getMonth(); // 0-indexed
    var nomeMes = CONFIG.MESES[mesAtual] || '';

    return {
      dataGeracao: hoje.toISOString(),
      mesReferencia: nomeMes,
      anoReferencia: CONFIG.ANO_CORRENTE,
      percentualGeral: percentualGeral,
      totalAtividades: totalAtv,
      concluidas: concluidas,
      emAndamento: emAndamento,
      atrasadas: atrasadas,
      noPrazo: noPrazo,
      replanejadas: replanejadas,
      naoIniciadas: naoIniciadas,
      totalOEs: oes.filter(function(o) { return o.tipo !== 'IMPACTO'; }).length,
      totalResultados: resultados.length,
      totalOKRs: okrs.length,
      totalIndicadores: indicadores.length,
      indicadoresComDados: indComDados,
      dadosPorEixo: dadosPorEixo,
      okrsAtrasados: okrsAtrasados.slice(0, 10),
      rankingDiretorias: ranking,
      oes: oes.filter(function(o) { return o.tipo !== 'IMPACTO'; })
    };
  } catch (e) {
    return { erro: e.message };
  }
}
