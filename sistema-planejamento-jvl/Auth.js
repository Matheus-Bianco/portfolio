/**
 * ============================================================
 * AUTENTICACAO E CONTROLE DE ACESSO
 * ============================================================
 */

var SESSAO_TTL_MS = 8 * 60 * 60 * 1000; // 8 horas

/**
 * Obter prefixo de sessao unico por usuario.
 * Usa Session.getActiveUser().getEmail() para isolar sessoes
 * quando o web app roda como "Execute as: Me".
 */
function getSessionPrefix() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (email) {
      return 'sess_' + email + '_';
    }
  } catch (e) {}
  return 'sess_default_';
}

function setSessionProp(key, value) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(getSessionPrefix() + key, value);
}

function getSessionProp(key) {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty(getSessionPrefix() + key);
}

function clearSessionProps() {
  var props = PropertiesService.getScriptProperties();
  var prefix = getSessionPrefix();
  var all = props.getProperties();
  var keys = Object.keys(all);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].indexOf(prefix) === 0) {
      props.deleteProperty(keys[i]);
    }
  }
}

/**
 * Validar login do usuario
 * @param {string} email - Email do usuario
 * @param {string} senha - Senha do usuario
 * @return {Object} Resultado da validacao
 */
function validarLogin(email, senha) {
  try {
    var sheet = getSheet(CONFIG.ABAS.USUARIOS);
    var dados = sheet.getDataRange().getValues();

    for (var i = 1; i < dados.length; i++) {
      var emailDB = String(dados[i][1]).trim().toLowerCase();
      var senhaDB = String(dados[i][5]).trim();

      if (emailDB === String(email).trim().toLowerCase() && senhaDB === String(senha).trim()) {
        // Salvar sessao (isolada por usuario via ScriptProperties + prefixo)
        setSessionProp('usuario_email', String(dados[i][1]).trim());
        setSessionProp('usuario_nome', String(dados[i][2]).trim());
        setSessionProp('usuario_perfil', String(dados[i][3]).trim());
        setSessionProp('usuario_diretoria', String(dados[i][4]).trim());
        setSessionProp('usuario_id', String(dados[i][0]));
        setSessionProp('sessao_ativa', 'true');
        setSessionProp('sessao_timestamp', String(new Date().getTime()));

        registrarLog(String(dados[i][1]).trim(), String(dados[i][0]), 'ACESSO', 'login', '', String(dados[i][3]).trim());

        return {
          status: 'ok',
          perfil: String(dados[i][3]).trim(),
          nome: String(dados[i][2]).trim(),
          diretoria: String(dados[i][4]).trim(),
          email: String(dados[i][1]).trim()
        };
      }
    }

    return { status: 'erro', mensagem: 'Usuario ou senha invalidos' };
  } catch (e) {
    return { status: 'erro', mensagem: 'Erro ao validar login: ' + e.message };
  }
}

/**
 * Obter usuario logado da sessao (com verificacao de expiracao)
 * @return {Object|null} Dados do usuario ou null
 */
function getUsuarioLogado() {
  try {
    var sessaoAtiva = getSessionProp('sessao_ativa');

    if (sessaoAtiva !== 'true') {
      return null;
    }

    // Verificar expiracao da sessao
    var timestamp = getSessionProp('sessao_timestamp');
    if (timestamp) {
      var idade = new Date().getTime() - parseInt(timestamp);
      if (idade > SESSAO_TTL_MS) {
        clearSessionProps();
        return null;
      }
    }

    return {
      email: getSessionProp('usuario_email'),
      nome: getSessionProp('usuario_nome'),
      perfil: getSessionProp('usuario_perfil'),
      diretoria: getSessionProp('usuario_diretoria'),
      id: getSessionProp('usuario_id')
    };
  } catch (e) {
    return null;
  }
}

/**
 * Encerrar sessao do usuario
 */
function logout() {
  try {
    var usuario = getUsuarioLogado();
    if (usuario) {
      registrarLog(usuario.email, usuario.id || '', 'ACESSO', 'logout', '', '');
    }
    clearSessionProps();
    return { status: 'ok' };
  } catch (e) {
    return { status: 'erro', mensagem: e.message };
  }
}

/**
 * Limpar sessoes expiradas de TODOS os usuarios.
 * Deve ser chamada periodicamente via trigger (a cada 1h).
 */
function limparSessoesExpiradas() {
  try {
    var props = PropertiesService.getScriptProperties();
    var all = props.getProperties();
    var keys = Object.keys(all);
    var agora = new Date().getTime();
    var removidas = 0;

    // Encontrar todos os prefixos de sessao unicos
    var prefixos = {};
    for (var i = 0; i < keys.length; i++) {
      if (keys[i].indexOf('sess_') === 0 && keys[i].indexOf('_sessao_timestamp') !== -1) {
        var prefix = keys[i].replace('sessao_timestamp', '');
        prefixos[prefix] = parseInt(all[keys[i]]) || 0;
      }
    }

    // Remover sessoes expiradas
    var prefixosList = Object.keys(prefixos);
    for (var p = 0; p < prefixosList.length; p++) {
      var pref = prefixosList[p];
      var ts = prefixos[pref];
      if (agora - ts > SESSAO_TTL_MS) {
        // Remover todas as props deste prefixo
        for (var k = 0; k < keys.length; k++) {
          if (keys[k].indexOf(pref) === 0) {
            props.deleteProperty(keys[k]);
          }
        }
        removidas++;
      }
    }

    Logger.log('Limpeza de sessoes: ' + removidas + ' sessao(oes) expirada(s) removida(s) de ' + prefixosList.length + ' total.');
    return { status: 'ok', removidas: removidas, total: prefixosList.length };
  } catch (e) {
    Logger.log('Erro ao limpar sessoes: ' + e.message);
    return { status: 'erro', mensagem: e.message };
  }
}

/**
 * Instalar trigger para limpeza automatica de sessoes (rodar 1x).
 * Cria um trigger horario que executa limparSessoesExpiradas().
 */
function instalarTriggerLimpezaSessoes() {
  // Remover triggers existentes desta funcao
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'limparSessoesExpiradas') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Criar novo trigger horario
  ScriptApp.newTrigger('limparSessoesExpiradas')
    .timeBased()
    .everyHours(1)
    .create();
  Logger.log('Trigger de limpeza de sessoes instalado (a cada 1 hora).');
  return { status: 'ok', mensagem: 'Trigger instalado com sucesso' };
}

/**
 * Alterar senha do usuario
 * @param {string} senhaAtual - Senha atual
 * @param {string} novaSenha - Nova senha
 * @return {Object} Resultado
 */
function alterarSenha(senhaAtual, novaSenha) {
  try {
    var usuario = getUsuarioLogado();
    if (!usuario) {
      return { status: 'erro', mensagem: 'Usuario nao esta logado' };
    }

    var sheet = getSheet(CONFIG.ABAS.USUARIOS);
    var dados = sheet.getDataRange().getValues();

    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][1]).trim().toLowerCase() === usuario.email.toLowerCase()) {
        if (String(dados[i][5]).trim() !== String(senhaAtual).trim()) {
          return { status: 'erro', mensagem: 'Senha atual incorreta' };
        }
        sheet.getRange(i + 1, 6).setValue(novaSenha);
        return { status: 'ok', mensagem: 'Senha alterada com sucesso' };
      }
    }

    return { status: 'erro', mensagem: 'Usuario nao encontrado' };
  } catch (e) {
    return { status: 'erro', mensagem: e.message };
  }
}

/**
 * Verificar se usuario tem permissao para uma acao
 * @param {string} acao - Acao a verificar
 * @return {boolean}
 */
function verificarPermissao(acao) {
  var usuario = getUsuarioLogado();
  if (!usuario) return false;

  var perfil = usuario.perfil;

  var permissoes = {
    'PLANEJAMENTO': [
      'criar_oe', 'editar_oe', 'excluir_oe',
      'criar_resultado', 'editar_resultado', 'excluir_resultado',
      'criar_okr', 'editar_okr', 'excluir_okr',
      'criar_atividade', 'editar_atividade', 'excluir_atividade',
      'gerenciar_usuarios',
      'visualizar_tudo',
      'fechar_mes',
      'importar_dados'
    ],
    'GERENCIA': [
      'editar_status_atividade',
      'editar_percentual_atividade',
      'editar_observacoes_atividade',
      'visualizar_diretoria'
    ],
    'SECRETARIO': [
      'visualizar_tudo'
    ]
  };

  var perfilPermissoes = permissoes[perfil] || [];
  return perfilPermissoes.indexOf(acao) !== -1;
}

/**
 * Obter dados de sessao para o cliente
 */
function getDadosSessao() {
  var usuario = getUsuarioLogado();
  if (!usuario) {
    return { logado: false };
  }

  return {
    logado: true,
    email: usuario.email,
    nome: usuario.nome,
    perfil: usuario.perfil,
    diretoria: usuario.diretoria
  };
}
