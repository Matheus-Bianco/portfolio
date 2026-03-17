/**
 * ============================================================
 * SERVICO DE LOG - REGISTRO DE ALTERACOES
 * ============================================================
 */

/**
 * Registrar log de alteracao
 * @param {string} email - Email do usuario (automatico se vazio)
 * @param {string} idRegistro - ID do registro alterado
 * @param {string} tipoRegistro - Tipo (OE, RESULTADO, OKR, ATIVIDADE, USUARIO)
 * @param {string} campoAlterado - Campo que foi alterado
 * @param {string} valorAntigo - Valor anterior
 * @param {string} valorNovo - Novo valor
 */
function registrarLog(email, idRegistro, tipoRegistro, campoAlterado, valorAntigo, valorNovo) {
  try {
    if (!email) {
      var usuario = getUsuarioLogado();
      email = usuario ? usuario.email : 'sistema';
    }

    var sheet = getSheet(CONFIG.ABAS.LOGS);
    var timestamp = new Date();

    sheet.appendRow([
      timestamp,
      email,
      idRegistro,
      tipoRegistro,
      campoAlterado,
      valorAntigo,
      valorNovo
    ]);

    // Invalidar cache de dashboard apos alteracoes em dados
    var tiposComCache = ['OE', 'RESULTADO', 'OKR', 'ATIVIDADE', 'INDICADOR'];
    if (tiposComCache.indexOf(tipoRegistro) !== -1) {
      invalidarCacheDashboard();
    }
  } catch (e) {
    // Log silencioso - nao interromper a operacao principal
    Logger.log('Erro ao registrar log: ' + e.message);
  }
}

/**
 * Consultar logs com filtros
 * @param {Object} filtros - Filtros opcionais
 * @return {Array} Lista de logs
 */
function consultarLogs(filtros) {
  var usuario = getUsuarioLogado();
  if (!usuario || usuario.perfil !== 'PLANEJAMENTO') {
    return [];
  }

  var sheet = getSheet(CONFIG.ABAS.LOGS);
  var dados = sheet.getDataRange().getValues();
  var resultado = [];

  for (var i = Math.max(1, dados.length - 500); i < dados.length; i++) {
    if (!dados[i][0]) continue;

    var log = {
      timestamp: dados[i][0],
      email: dados[i][1],
      id_registro: dados[i][2],
      tipo_registro: dados[i][3],
      campo_alterado: dados[i][4],
      valor_antigo: dados[i][5],
      valor_novo: dados[i][6]
    };

    // Aplicar filtros
    if (filtros) {
      if (filtros.email && String(log.email).toLowerCase().indexOf(String(filtros.email).toLowerCase()) === -1) continue;
      if (filtros.tipo && String(log.tipo_registro) !== String(filtros.tipo)) continue;
      if (filtros.id_registro && String(log.id_registro) !== String(filtros.id_registro)) continue;
    }

    resultado.push(log);
  }

  // Retornar os mais recentes primeiro
  resultado.reverse();
  return resultado;
}
