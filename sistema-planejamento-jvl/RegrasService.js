/**
 * ============================================================
 * REGRAS DE NEGOCIO E AUTOMACAO
 * ============================================================
 */

/**
 * Verificar e marcar atividades atrasadas automaticamente
 * Regra: Se mes_atual > mes_planejado e status != Concluido -> Atrasado
 */
function verificarAtrasados() {
  var sheet = getSheet(CONFIG.ABAS.ATIVIDADES);
  var dados = sheet.getDataRange().getValues();
  var mesAtual = new Date().getMonth(); // 0-indexed
  var mesesNomes = CONFIG.MESES;
  var alterados = 0;

  for (var i = 1; i < dados.length; i++) {
    if (!dados[i][0]) continue;

    var status = String(dados[i][11]).toLowerCase();
    var mesPlanejado = String(dados[i][5]).trim();

    // Pular se ja concluido ou removido
    if (status.indexOf('conclu') !== -1 && status.indexOf('nao') === -1 && status.indexOf('parcial') === -1) continue;
    if (status.indexOf('removido') !== -1) continue;

    // Encontrar indice do mes planejado
    var indiceMesPlanejado = -1;
    for (var m = 0; m < mesesNomes.length; m++) {
      if (mesesNomes[m].toLowerCase() === mesPlanejado.toLowerCase()) {
        indiceMesPlanejado = m;
        break;
      }
    }

    // Se mes planejado foi encontrado e ja passou
    if (indiceMesPlanejado !== -1 && mesAtual > indiceMesPlanejado) {
      if (status !== 'atrasado') {
        var statusAntigo = dados[i][11];
        sheet.getRange(i + 1, 12).setValue('Atrasado');
        registrarLog('sistema', String(dados[i][0]), 'ATIVIDADE', 'status',
          statusAntigo, 'Atrasado (automatico)');
        alterados++;
      }
    }
  }

  return alterados + ' atividades marcadas como atrasadas';
}

/**
 * Validar e corrigir percentuais acima de 100%
 */
function validarPercentuais() {
  var sheet = getSheet(CONFIG.ABAS.ATIVIDADES);
  var dados = sheet.getDataRange().getValues();
  var corrigidos = 0;

  for (var i = 1; i < dados.length; i++) {
    if (!dados[i][0]) continue;

    var percentual = parseFloat(dados[i][15]) || 0;
    if (percentual > 100) {
      sheet.getRange(i + 1, 16).setValue(100);
      registrarLog('sistema', String(dados[i][0]), 'ATIVIDADE', 'percentual_execucao',
        String(percentual), '100 (corrigido automaticamente)');
      corrigidos++;
    }
  }

  return corrigidos + ' percentuais corrigidos';
}

/**
 * Fechar mes - bloquear edicao de atividades do mes anterior
 * (Implementado via flag no PropertiesService)
 */
function fecharMes(mes, ano) {
  if (!verificarPermissao('fechar_mes')) {
    return { status: 'erro', mensagem: 'Sem permissao para fechar mes' };
  }

  var props = PropertiesService.getScriptProperties();
  var chave = 'MES_FECHADO_' + ano + '_' + mes;
  props.setProperty(chave, 'true');

  registrarLog('', '', 'SISTEMA', 'fechamento_mes', '', mes + '/' + ano);
  return { status: 'ok', mensagem: 'Mes ' + mes + '/' + ano + ' fechado com sucesso' };
}

/**
 * Verificar se um mes esta fechado
 */
function mesFechado(mes, ano) {
  var props = PropertiesService.getScriptProperties();
  var chave = 'MES_FECHADO_' + ano + '_' + mes;
  return props.getProperty(chave) === 'true';
}

/**
 * Reabrir mes (apenas PLANEJAMENTO)
 */
function reabrirMes(mes, ano) {
  if (!verificarPermissao('fechar_mes')) {
    return { status: 'erro', mensagem: 'Sem permissao' };
  }

  var props = PropertiesService.getScriptProperties();
  var chave = 'MES_FECHADO_' + ano + '_' + mes;
  props.deleteProperty(chave);

  registrarLog('', '', 'SISTEMA', 'reabertura_mes', '', mes + '/' + ano);
  return { status: 'ok', mensagem: 'Mes ' + mes + '/' + ano + ' reaberto' };
}

/**
 * Listar meses fechados
 */
function listarMesesFechados() {
  var props = PropertiesService.getScriptProperties();
  var todas = props.getProperties();
  var mesesFechados = [];

  for (var chave in todas) {
    if (chave.indexOf('MES_FECHADO_') === 0 && todas[chave] === 'true') {
      var partes = chave.replace('MES_FECHADO_', '').split('_');
      mesesFechados.push({ ano: partes[0], mes: partes[1] });
    }
  }

  return mesesFechados;
}

/**
 * Trigger automatico - configurar para rodar diariamente
 * Menu: Ferramentas > Gatilhos > Adicionar gatilho
 */
function rotinaDiaria() {
  verificarAtrasados();
  validarPercentuais();
}

/**
 * Criar trigger automatico
 */
function criarTriggerDiario() {
  // Remover triggers existentes
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'rotinaDiaria') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Criar novo trigger diario as 6h
  ScriptApp.newTrigger('rotinaDiaria')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  return 'Trigger diario criado com sucesso';
}
