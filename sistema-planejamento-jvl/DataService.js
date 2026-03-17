/**
 * ============================================================
 * SERVICO DE DADOS - CRUD COMPLETO
 * ============================================================
 */

// ==================== OBJETIVOS ESTRATEGICOS ====================

/**
 * Listar todos os OEs
 */
function listarOEs() {
  var sheet = getSheet(CONFIG.ABAS.OE);
  var dados = sheet.getDataRange().getValues();
  var resultado = [];

  for (var i = 1; i < dados.length; i++) {
    if (dados[i][0]) {
      resultado.push({
        id_oe: dados[i][0],
        eixo: dados[i][1],
        descricao: dados[i][2],
        objetivo_impacto: dados[i][3],
        tipo: dados[i][4]
      });
    }
  }
  return resultado;
}

/**
 * Criar novo OE
 */
function criarOE(dados) {
  if (!verificarPermissao('criar_oe')) {
    return { status: 'erro', mensagem: 'Sem permissao' };
  }
  var sheet = getSheet(CONFIG.ABAS.OE);
  sheet.appendRow([dados.id_oe, dados.eixo, dados.descricao, dados.objetivo_impacto || '', dados.tipo || 'ESTRATEGICO']);
  registrarLog('', dados.id_oe, 'OE', 'criacao', '', JSON.stringify(dados));
  return { status: 'ok', mensagem: 'OE criado com sucesso' };
}

/**
 * Editar OE
 */
function editarOE(idOE, campo, valor) {
  if (!verificarPermissao('editar_oe')) {
    return { status: 'erro', mensagem: 'Sem permissao' };
  }
  var sheet = getSheet(CONFIG.ABAS.OE);
  var dados = sheet.getDataRange().getValues();
  var campos = dados[0];
  var colIndex = campos.indexOf(campo);

  if (colIndex === -1) return { status: 'erro', mensagem: 'Campo nao encontrado' };

  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(idOE)) {
      var valorAntigo = dados[i][colIndex];
      sheet.getRange(i + 1, colIndex + 1).setValue(valor);
      registrarLog('', idOE, 'OE', campo, String(valorAntigo), String(valor));
      return { status: 'ok', mensagem: 'OE atualizado' };
    }
  }
  return { status: 'erro', mensagem: 'OE nao encontrado' };
}

// ==================== RESULTADOS ====================

/**
 * Listar resultados (com filtro opcional por OE)
 */
function listarResultados(filtroOE) {
  var sheet = getSheet(CONFIG.ABAS.RESULTADOS);
  var dados = sheet.getDataRange().getValues();
  var resultado = [];

  for (var i = 1; i < dados.length; i++) {
    if (dados[i][0]) {
      if (filtroOE && String(dados[i][1]) !== String(filtroOE)) continue;
      resultado.push({
        id_resultado: dados[i][0],
        id_oe: dados[i][1],
        descricao: dados[i][2],
        meta_2025: dados[i][3],
        meta_2029: dados[i][4],
        classificacao: dados[i][5],
        responsavel: dados[i][6],
        macroprocesso: dados[i][7],
        projeto_codigo: dados[i][8],
        projeto_nome: dados[i][9],
        diretoria: dados[i][10],
        unidade: dados[i][11] || '',
        meta_2026: dados[i][12] || '',
        requisitos_qualidade: dados[i][13] || ''
      });
    }
  }
  return resultado;
}

/**
 * Criar resultado
 */
function criarResultado(dados) {
  if (!verificarPermissao('criar_resultado')) {
    return { status: 'erro', mensagem: 'Sem permissao' };
  }
  var sheet = getSheet(CONFIG.ABAS.RESULTADOS);
  var id = gerarProximoId(sheet, 'R', 2);
  sheet.appendRow([
    id, dados.id_oe, dados.descricao, dados.meta_2025 || '',
    dados.meta_2029 || '', dados.classificacao || '', dados.responsavel || '',
    dados.macroprocesso || '', dados.projeto_codigo || '', dados.projeto_nome || '',
    dados.diretoria || '', dados.unidade || '', dados.meta_2026 || '',
    dados.requisitos_qualidade || ''
  ]);
  registrarLog('', id, 'RESULTADO', 'criacao', '', JSON.stringify(dados));
  return { status: 'ok', mensagem: 'Resultado criado', id: id };
}

/**
 * Editar resultado
 */
function editarResultado(idResultado, campo, valor) {
  if (!verificarPermissao('editar_resultado')) {
    return { status: 'erro', mensagem: 'Sem permissao' };
  }
  var sheet = getSheet(CONFIG.ABAS.RESULTADOS);
  var dados = sheet.getDataRange().getValues();
  var campos = dados[0];
  var colIndex = campos.indexOf(campo);

  if (colIndex === -1) return { status: 'erro', mensagem: 'Campo nao encontrado' };

  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(idResultado)) {
      var valorAntigo = dados[i][colIndex];
      sheet.getRange(i + 1, colIndex + 1).setValue(valor);
      registrarLog('', idResultado, 'RESULTADO', campo, String(valorAntigo), String(valor));
      return { status: 'ok', mensagem: 'Resultado atualizado' };
    }
  }
  return { status: 'erro', mensagem: 'Resultado nao encontrado' };
}

// ==================== OKRs ====================

/**
 * Listar OKRs (com filtro opcional)
 */
function listarOKRs(filtroResultado, filtroTrimestre) {
  var sheet = getSheet(CONFIG.ABAS.OKRS);
  var dados = sheet.getDataRange().getValues();
  var resultado = [];

  for (var i = 1; i < dados.length; i++) {
    if (dados[i][0]) {
      if (filtroResultado && String(dados[i][1]) !== String(filtroResultado)) continue;
      if (filtroTrimestre && String(dados[i][3]) !== String(filtroTrimestre)) continue;
      resultado.push({
        id_okr: dados[i][0],
        id_resultado: dados[i][1],
        id_oe: dados[i][2],
        trimestre: dados[i][3],
        descricao: dados[i][4],
        situacao: dados[i][5],
        percentual: dados[i][6],
        requisitos_qualidade: dados[i][7],
        ano: dados[i][8],
        diretoria: dados[i][9] || '',
        unidade: dados[i][10] || '',
        eixo: dados[i][11] || ''
      });
    }
  }
  return resultado;
}

/**
 * Criar OKR
 */
function criarOKR(dados) {
  if (!verificarPermissao('criar_okr')) {
    return { status: 'erro', mensagem: 'Sem permissao' };
  }
  var sheet = getSheet(CONFIG.ABAS.OKRS);
  var id = gerarProximoId(sheet, 'OKR', 3);
  sheet.appendRow([
    id, dados.id_resultado, dados.id_oe || '', dados.trimestre, dados.descricao,
    dados.situacao || 'Em andamento', dados.percentual || 0,
    dados.requisitos_qualidade || '', dados.ano || new Date().getFullYear(),
    dados.diretoria || '', dados.unidade || '', dados.eixo || ''
  ]);
  registrarLog('', id, 'OKR', 'criacao', '', JSON.stringify(dados));
  return { status: 'ok', mensagem: 'OKR criado', id: id };
}

/**
 * Editar OKR
 */
function editarOKR(idOKR, campo, valor) {
  var usuario = getUsuarioLogado();
  if (!usuario) return { status: 'erro', mensagem: 'Nao logado' };

  // PLANEJAMENTO pode editar tudo; GERENCIA apenas situacao e percentual
  var camposPermitidosGerencia = ['situacao', 'percentual'];
  if (usuario.perfil === 'GERENCIA' && camposPermitidosGerencia.indexOf(campo) === -1) {
    return { status: 'erro', mensagem: 'Sem permissao para editar este campo' };
  }
  if (usuario.perfil === 'SECRETARIO') {
    var camposPermitidosSecretario = ['requisitos_qualidade'];
    if (camposPermitidosSecretario.indexOf(campo) === -1) {
      return { status: 'erro', mensagem: 'Sem permissao' };
    }
  }

  var sheet = getSheet(CONFIG.ABAS.OKRS);
  var dados = sheet.getDataRange().getValues();
  var campos = dados[0];
  var colIndex = campos.indexOf(campo);

  if (colIndex === -1) return { status: 'erro', mensagem: 'Campo nao encontrado' };

  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(idOKR)) {
      var valorAntigo = dados[i][colIndex];
      sheet.getRange(i + 1, colIndex + 1).setValue(valor);
      registrarLog('', idOKR, 'OKR', campo, String(valorAntigo), String(valor));
      return { status: 'ok', mensagem: 'OKR atualizado' };
    }
  }
  return { status: 'erro', mensagem: 'OKR nao encontrado' };
}

// ==================== ATIVIDADES ====================

/**
 * Listar atividades com filtros
 */
function listarAtividades(filtros) {
  var sheet = getSheet(CONFIG.ABAS.ATIVIDADES);
  var dados = sheet.getDataRange().getValues();
  var resultado = [];
  var usuario = getUsuarioLogado();

  for (var i = 1; i < dados.length; i++) {
    if (!dados[i][0]) continue;

    // Filtro por diretoria para GERENCIA
    if (usuario && usuario.perfil === 'GERENCIA') {
      if (String(dados[i][12]) !== usuario.diretoria) continue;
    }

    // Filtros opcionais
    if (filtros) {
      if (filtros.id_okr && String(dados[i][1]) !== String(filtros.id_okr)) continue;
      if (filtros.id_resultado && String(dados[i][2]) !== String(filtros.id_resultado)) continue;
      if (filtros.status && String(dados[i][11]) !== String(filtros.status)) continue;
      if (filtros.diretoria && String(dados[i][12]) !== String(filtros.diretoria)) continue;
      if (filtros.responsavel && String(dados[i][9]).toLowerCase().indexOf(String(filtros.responsavel).toLowerCase()) === -1) continue;
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
 * Criar atividade
 */
function criarAtividade(dados) {
  if (!verificarPermissao('criar_atividade')) {
    return { status: 'erro', mensagem: 'Sem permissao' };
  }
  var sheet = getSheet(CONFIG.ABAS.ATIVIDADES);
  var id = gerarProximoId(sheet, 'ATV', 3);
  var percentual = dados.percentual_execucao || 0;
  if (percentual > 100) percentual = 100;

  sheet.appendRow([
    id, dados.id_okr || '', dados.id_resultado || '', dados.id_oe || '',
    dados.descricao, dados.mes_planejado || '', dados.mes_realizado || '',
    dados.criterio || 'Qualitativo', dados.peso || 1,
    dados.responsavel || '', dados.observacoes || '',
    dados.status || 'No prazo', dados.diretoria || '',
    dados.unidade || '', dados.eixo || '',
    percentual, dados.mes_entrega || ''
  ]);
  registrarLog('', id, 'ATIVIDADE', 'criacao', '', JSON.stringify(dados));
  return { status: 'ok', mensagem: 'Atividade criada', id: id };
}

/**
 * Editar atividade (com controle de acesso por perfil)
 */
function editarAtividade(idAtividade, campo, valor) {
  var usuario = getUsuarioLogado();
  if (!usuario) return { status: 'erro', mensagem: 'Nao logado' };

  // Campos que GERENCIA pode editar
  var camposGerencia = ['status', 'percentual_execucao', 'observacoes', 'mes_realizado'];
  if (usuario.perfil === 'GERENCIA' && camposGerencia.indexOf(campo) === -1) {
    return { status: 'erro', mensagem: 'Sem permissao para editar este campo' };
  }
  if (usuario.perfil === 'SECRETARIO') {
    return { status: 'erro', mensagem: 'Perfil nao permite edicao' };
  }

  // Validar percentual maximo 100%
  if (campo === 'percentual_execucao') {
    valor = Math.min(parseFloat(valor) || 0, 100);
  }

  var sheet = getSheet(CONFIG.ABAS.ATIVIDADES);
  var dados = sheet.getDataRange().getValues();
  var campos = dados[0];
  var colIndex = campos.indexOf(campo);

  if (colIndex === -1) return { status: 'erro', mensagem: 'Campo nao encontrado' };

  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(idAtividade)) {
      // GERENCIA: verificar se atividade pertence a sua diretoria
      var dirCol = campos.indexOf('diretoria');
      if (usuario.perfil === 'GERENCIA' && dirCol !== -1 && String(dados[i][dirCol]) !== usuario.diretoria) {
        return { status: 'erro', mensagem: 'Atividade nao pertence a sua diretoria' };
      }

      var valorAntigo = dados[i][colIndex];
      sheet.getRange(i + 1, colIndex + 1).setValue(valor);
      registrarLog('', idAtividade, 'ATIVIDADE', campo, String(valorAntigo), String(valor));
      return { status: 'ok', mensagem: 'Atividade atualizada' };
    }
  }
  return { status: 'erro', mensagem: 'Atividade nao encontrada' };
}

/**
 * Excluir atividade
 */
function excluirAtividade(idAtividade) {
  if (!verificarPermissao('excluir_atividade')) {
    return { status: 'erro', mensagem: 'Sem permissao' };
  }
  var sheet = getSheet(CONFIG.ABAS.ATIVIDADES);
  var dados = sheet.getDataRange().getValues();

  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(idAtividade)) {
      registrarLog('', idAtividade, 'ATIVIDADE', 'exclusao', JSON.stringify(dados[i]), '');
      sheet.deleteRow(i + 1);
      return { status: 'ok', mensagem: 'Atividade excluida' };
    }
  }
  return { status: 'erro', mensagem: 'Atividade nao encontrada' };
}

// ==================== USUARIOS ====================

/**
 * Listar usuarios
 */
function listarUsuarios() {
  if (!verificarPermissao('gerenciar_usuarios')) {
    return { status: 'erro', mensagem: 'Sem permissao' };
  }
  var sheet = getSheet(CONFIG.ABAS.USUARIOS);
  var dados = sheet.getDataRange().getValues();
  var resultado = [];

  for (var i = 1; i < dados.length; i++) {
    if (dados[i][0]) {
      resultado.push({
        id_usuario: dados[i][0],
        email: dados[i][1],
        nome: dados[i][2],
        perfil: dados[i][3],
        diretoria: dados[i][4],
        // nao retornar senha
      });
    }
  }
  return resultado;
}

/**
 * Criar usuario
 */
function criarUsuario(dados) {
  if (!verificarPermissao('gerenciar_usuarios')) {
    return { status: 'erro', mensagem: 'Sem permissao' };
  }

  var sheet = getSheet(CONFIG.ABAS.USUARIOS);
  var todosUsuarios = sheet.getDataRange().getValues();

  // Verificar email duplicado
  for (var i = 1; i < todosUsuarios.length; i++) {
    if (String(todosUsuarios[i][1]).trim().toLowerCase() === String(dados.email).trim().toLowerCase()) {
      return { status: 'erro', mensagem: 'Email ja cadastrado' };
    }
  }

  var novoId = todosUsuarios.length;
  sheet.appendRow([
    novoId, dados.email, dados.nome, dados.perfil,
    dados.diretoria || '', '123456'
  ]);
  registrarLog('', String(novoId), 'USUARIO', 'criacao', '', dados.nome + ' (' + dados.email + ')');
  return { status: 'ok', mensagem: 'Usuario criado com senha padrao: 123456' };
}

/**
 * Editar usuario
 */
function editarUsuario(idUsuario, dados) {
  if (!verificarPermissao('gerenciar_usuarios')) {
    return { status: 'erro', mensagem: 'Sem permissao' };
  }

  var sheet = getSheet(CONFIG.ABAS.USUARIOS);
  var todosUsuarios = sheet.getDataRange().getValues();

  for (var i = 1; i < todosUsuarios.length; i++) {
    if (String(todosUsuarios[i][0]) === String(idUsuario)) {
      if (dados.email) sheet.getRange(i + 1, 2).setValue(dados.email);
      if (dados.nome) sheet.getRange(i + 1, 3).setValue(dados.nome);
      if (dados.perfil) sheet.getRange(i + 1, 4).setValue(dados.perfil);
      if (dados.diretoria !== undefined) sheet.getRange(i + 1, 5).setValue(dados.diretoria);
      registrarLog('', String(idUsuario), 'USUARIO', 'edicao', '', JSON.stringify(dados));
      return { status: 'ok', mensagem: 'Usuario atualizado' };
    }
  }
  return { status: 'erro', mensagem: 'Usuario nao encontrado' };
}

/**
 * Resetar senha do usuario
 */
function resetarSenhaUsuario(idUsuario) {
  if (!verificarPermissao('gerenciar_usuarios')) {
    return { status: 'erro', mensagem: 'Sem permissao' };
  }

  var sheet = getSheet(CONFIG.ABAS.USUARIOS);
  var dados = sheet.getDataRange().getValues();

  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(idUsuario)) {
      sheet.getRange(i + 1, 6).setValue('123456');
      registrarLog('', String(idUsuario), 'USUARIO', 'reset_senha', '', 'Senha resetada');
      return { status: 'ok', mensagem: 'Senha resetada para 123456' };
    }
  }
  return { status: 'erro', mensagem: 'Usuario nao encontrado' };
}

/**
 * Excluir usuario
 */
function excluirUsuario(idUsuario) {
  if (!verificarPermissao('gerenciar_usuarios')) {
    return { status: 'erro', mensagem: 'Sem permissao' };
  }

  var sheet = getSheet(CONFIG.ABAS.USUARIOS);
  var dados = sheet.getDataRange().getValues();

  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(idUsuario)) {
      registrarLog('', String(idUsuario), 'USUARIO', 'exclusao', dados[i][2], '');
      sheet.deleteRow(i + 1);
      return { status: 'ok', mensagem: 'Usuario excluido' };
    }
  }
  return { status: 'erro', mensagem: 'Usuario nao encontrado' };
}

// ==================== BUSCA GERAL ====================

/**
 * Busca geral em todas as entidades
 */
function buscarGeral(termo) {
  var resultados = [];
  var termoLower = String(termo).toLowerCase();

  // Buscar em OEs
  var oes = listarOEs();
  for (var i = 0; i < oes.length; i++) {
    if (oes[i].descricao.toLowerCase().indexOf(termoLower) !== -1 ||
        oes[i].id_oe.toLowerCase().indexOf(termoLower) !== -1) {
      resultados.push({ tipo: 'OE', id: oes[i].id_oe, descricao: oes[i].descricao });
    }
  }

  // Buscar em Atividades
  var atividades = listarAtividades(null);
  for (var j = 0; j < atividades.length; j++) {
    if (atividades[j].descricao.toLowerCase().indexOf(termoLower) !== -1 ||
        atividades[j].responsavel.toLowerCase().indexOf(termoLower) !== -1) {
      resultados.push({ tipo: 'ATIVIDADE', id: atividades[j].id_atividade, descricao: atividades[j].descricao });
    }
  }

  return resultados;
}

/**
 * Obter lista de diretorias unicas
 */
function listarDiretorias() {
  var sheet = getSheet(CONFIG.ABAS.RESULTADOS);
  var dados = sheet.getDataRange().getValues();
  var diretorias = {};

  for (var i = 1; i < dados.length; i++) {
    var dir = String(dados[i][10]).trim();
    if (dir && dir !== 'undefined' && dir !== '') {
      diretorias[dir] = true;
    }
  }

  // Adicionar do sheet de usuarios
  var sheetU = getSheet(CONFIG.ABAS.USUARIOS);
  var dadosU = sheetU.getDataRange().getValues();
  for (var j = 1; j < dadosU.length; j++) {
    var dirU = String(dadosU[j][4]).trim();
    if (dirU && dirU !== 'undefined' && dirU !== '') {
      diretorias[dirU] = true;
    }
  }

  return Object.keys(diretorias).sort();
}

// ==================== NOMES_DIRETORIAS / GLOSSARIO ====================

/**
 * Listar dados da aba Nomes_Diretorias (sigla -> descricao)
 */
function listarNomesDiretorias() {
  try {
    var sheet = getSheet(CONFIG.ABAS.NOMES_DIRETORIAS);
    var dados = sheet.getDataRange().getValues();
    var headers = dados[0].map(function(h) { return String(h).trim().toLowerCase(); });
    var resultado = [];
    for (var i = 1; i < dados.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        row[headers[j]] = String(dados[i][j] || '').trim();
      }
      if (row['unidade'] || row['sigla'] || row['diretoria']) {
        resultado.push(row);
      }
    }
    return resultado;
  } catch (e) {
    Logger.log('Erro listarNomesDiretorias: ' + e.message);
    return [];
  }
}

/**
 * Gerar glossario de siglas comparando Nomes_Diretorias com siglas usadas no sistema
 */
function gerarGlossarioSiglas() {
  var nomes = listarNomesDiretorias();

  // Mapa de siglas oficiais (da aba Nomes_Diretorias)
  var glossarioOficial = {};
  for (var i = 0; i < nomes.length; i++) {
    var sigla = nomes[i]['unidade'] || nomes[i]['sigla'] || '';
    var descricao = nomes[i]['descricao'] || nomes[i]['nome'] || '';
    var diretoria = nomes[i]['diretoria'] || '';
    if (sigla) {
      glossarioOficial[sigla.toUpperCase()] = {
        sigla: sigla,
        descricao: descricao,
        diretoria: diretoria
      };
    }
  }

  // Siglas usadas no sistema (das abas RESULTADOS, ATIVIDADES, INDICADORES)
  var siglasUsadas = {};

  // Resultados: coluna diretoria(10) e unidade(11)
  try {
    var sheetR = getSheet(CONFIG.ABAS.RESULTADOS);
    var dadosR = sheetR.getDataRange().getValues();
    for (var r = 1; r < dadosR.length; r++) {
      var dirR = String(dadosR[r][10] || '').trim();
      var uniR = String(dadosR[r][11] || '').trim();
      if (dirR) siglasUsadas[dirR.toUpperCase()] = { fonte: 'RESULTADOS', campo: 'diretoria' };
      if (uniR) siglasUsadas[uniR.toUpperCase()] = { fonte: 'RESULTADOS', campo: 'unidade' };
    }
  } catch(e) {}

  // Atividades: coluna diretoria(12), unidade(13)
  try {
    var sheetA = getSheet(CONFIG.ABAS.ATIVIDADES);
    var dadosA = sheetA.getDataRange().getValues();
    for (var a = 1; a < dadosA.length; a++) {
      var dirA = String(dadosA[a][12] || '').trim();
      var uniA = String(dadosA[a][13] || '').trim();
      if (dirA) siglasUsadas[dirA.toUpperCase()] = { fonte: 'ATIVIDADES', campo: 'diretoria' };
      if (uniA) siglasUsadas[uniA.toUpperCase()] = { fonte: 'ATIVIDADES', campo: 'unidade' };
    }
  } catch(e) {}

  // Indicadores: coluna projeto_gerencia(4)
  try {
    var sheetI = getSheet(CONFIG.ABAS.INDICADORES);
    var dadosI = sheetI.getDataRange().getValues();
    for (var ii = 1; ii < dadosI.length; ii++) {
      var pg = String(dadosI[ii][4] || '').trim();
      if (pg) siglasUsadas[pg.toUpperCase()] = { fonte: 'INDICADORES', campo: 'projeto_gerencia' };
    }
  } catch(e) {}

  // Comparar
  var resultado = {
    glossario: glossarioOficial,
    siglasUsadas: Object.keys(siglasUsadas).sort(),
    correspondencias: [],
    semCorrespondencia: [],
    siglasNaoDocumentadas: []
  };

  // Siglas usadas que TEM correspondencia no glossario
  var usadasKeys = Object.keys(siglasUsadas).sort();
  for (var u = 0; u < usadasKeys.length; u++) {
    var sk = usadasKeys[u];
    if (glossarioOficial[sk]) {
      resultado.correspondencias.push({
        sigla: sk,
        descricao: glossarioOficial[sk].descricao,
        diretoria: glossarioOficial[sk].diretoria,
        fonte: siglasUsadas[sk].fonte,
        campo: siglasUsadas[sk].campo
      });
    } else {
      resultado.siglasNaoDocumentadas.push({
        sigla: sk,
        fonte: siglasUsadas[sk].fonte,
        campo: siglasUsadas[sk].campo
      });
    }
  }

  // Siglas do glossario que NAO sao usadas no sistema
  var oficialKeys = Object.keys(glossarioOficial).sort();
  for (var g = 0; g < oficialKeys.length; g++) {
    if (!siglasUsadas[oficialKeys[g]]) {
      resultado.semCorrespondencia.push({
        sigla: oficialKeys[g],
        descricao: glossarioOficial[oficialKeys[g]].descricao
      });
    }
  }

  return resultado;
}

// ==================== INDICADORES ====================

/**
 * Listar indicadores (com filtro opcional por OE)
 * Nova estrutura unificada com 35 colunas (inclui dados mensais jan-dez)
 */
function listarIndicadores(filtroOE) {
  try {
    var sheet = getSheet(CONFIG.ABAS.INDICADORES);
    var dados = sheet.getDataRange().getValues();
    var resultado = [];

    for (var i = 1; i < dados.length; i++) {
      if (!dados[i][0]) continue;
      if (filtroOE && String(dados[i][2]) !== String(filtroOE)) continue;
      resultado.push({
        id_indicador: dados[i][0],
        eixo: dados[i][1],
        id_oe: dados[i][2],
        categoria: dados[i][3],
        projeto_gerencia: dados[i][4],
        cod: dados[i][5],
        indicador: dados[i][6],
        periodicidade: dados[i][7],
        responsavel: dados[i][8],
        origem: dados[i][9],
        pei: dados[i][10],
        formula: dados[i][11],
        val_2021: dados[i][12],
        val_2022: dados[i][13],
        val_2023: dados[i][14],
        val_2024: dados[i][15],
        val_2025: dados[i][16],
        meta_2026: dados[i][17],
        polaridade: dados[i][18],
        jan: dados[i][19],
        fev: dados[i][20],
        mar: dados[i][21],
        abr: dados[i][22],
        mai: dados[i][23],
        jun: dados[i][24],
        jul: dados[i][25],
        ago: dados[i][26],
        set: dados[i][27],
        out: dados[i][28],
        nov: dados[i][29],
        dez: dados[i][30],
        status_coleta: dados[i][31],
        pct_evolucao: dados[i][32],
        pct_atingimento: dados[i][33],
        atingimento_meta: dados[i][34],
        observacao: dados[i][35] || ''
      });
    }
    return resultado;
  } catch (e) {
    return [];
  }
}

/**
 * Editar campo de um indicador (inclui jan-dez para edicao inline de valores mensais)
 */
function editarIndicador(idIndicador, campo, valor) {
  var usuario = getUsuarioLogado();
  if (!usuario) return { status: 'erro', mensagem: 'Nao logado' };
  if (usuario.perfil === 'SECRETARIO') return { status: 'erro', mensagem: 'Sem permissao' };

  var sheet = getSheet(CONFIG.ABAS.INDICADORES);
  var dados = sheet.getDataRange().getValues();
  var campos = dados[0];
  var colIndex = campos.indexOf(campo);

  if (colIndex === -1) return { status: 'erro', mensagem: 'Campo nao encontrado: ' + campo };

  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(idIndicador)) {
      var valorAntigo = dados[i][colIndex];
      sheet.getRange(i + 1, colIndex + 1).setValue(valor);
      registrarLog('', idIndicador, 'INDICADOR', campo, String(valorAntigo), String(valor));
      return { status: 'ok', mensagem: 'Indicador atualizado' };
    }
  }
  return { status: 'erro', mensagem: 'Indicador nao encontrado' };
}

// ==================== CHECKS REQUISITOS OKR ====================

/**
 * Obter todos os checks de requisitos de OKRs (armazenados via ScriptProperties)
 */
function getChecksRequisitos() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('checks_requisitos_okr');
  if (raw) {
    try { return JSON.parse(raw); } catch(e) {}
  }
  return {};
}

/**
 * Salvar check de requisito individual de um OKR
 */
function salvarCheckRequisito(idOkr, index, valor) {
  var usuario = getUsuarioLogado();
  if (!usuario) return { status: 'erro', mensagem: 'Nao logado' };
  if (usuario.perfil !== 'SECRETARIO' && usuario.perfil !== 'PLANEJAMENTO') {
    return { status: 'erro', mensagem: 'Sem permissao' };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { status: 'erro', mensagem: 'Sistema ocupado, tente novamente' };
  }

  try {
    var props = PropertiesService.getScriptProperties();
    var checks = {};
    var raw = props.getProperty('checks_requisitos_okr');
    if (raw) {
      try { checks = JSON.parse(raw); } catch(e) {}
    }

    if (!checks[idOkr]) checks[idOkr] = {};
    checks[idOkr][String(index)] = !!valor;

    props.setProperty('checks_requisitos_okr', JSON.stringify(checks));
    return { status: 'ok', mensagem: 'Check salvo' };
  } finally {
    lock.releaseLock();
  }
}

// ==================== COMENTARIOS OKR ====================

/**
 * Obter todos os comentarios de OKRs (armazenados via ScriptProperties)
 */
function getComentariosOKR() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('comentarios_okr');
  if (raw) {
    try { return JSON.parse(raw); } catch(e) {}
  }
  return {};
}

/**
 * Salvar comentario de um OKR
 */
function salvarComentarioOKR(idOkr, texto) {
  var usuario = getUsuarioLogado();
  if (!usuario) return { status: 'erro', mensagem: 'Nao logado' };
  if (usuario.perfil !== 'SECRETARIO' && usuario.perfil !== 'PLANEJAMENTO') {
    return { status: 'erro', mensagem: 'Sem permissao' };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { status: 'erro', mensagem: 'Sistema ocupado, tente novamente' };
  }

  try {
    var props = PropertiesService.getScriptProperties();
    var comentarios = {};
    var raw = props.getProperty('comentarios_okr');
    if (raw) {
      try { comentarios = JSON.parse(raw); } catch(e) {}
    }

    if (texto && texto.trim()) {
      comentarios[idOkr] = texto.trim();
    } else {
      delete comentarios[idOkr];
    }

    props.setProperty('comentarios_okr', JSON.stringify(comentarios));
    registrarLog('', idOkr, 'COMENTARIO_OKR', 'comentario', '', texto || '(removido)');
    return { status: 'ok', mensagem: 'Comentario salvo' };
  } finally {
    lock.releaseLock();
  }
}

// ==================== COMENTARIOS ATIVIDADES ====================

function getComentariosAtividade() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('comentarios_atividade');
  if (raw) {
    try { return JSON.parse(raw); } catch(e) {}
  }
  return {};
}

function salvarComentarioAtividade(idAtividade, texto) {
  var usuario = getUsuarioLogado();
  if (!usuario) return { status: 'erro', mensagem: 'Nao logado' };
  if (usuario.perfil !== 'SECRETARIO' && usuario.perfil !== 'PLANEJAMENTO') {
    return { status: 'erro', mensagem: 'Sem permissao' };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { status: 'erro', mensagem: 'Sistema ocupado, tente novamente' };
  }

  try {
    var props = PropertiesService.getScriptProperties();
    var comentarios = {};
    var raw = props.getProperty('comentarios_atividade');
    if (raw) {
      try { comentarios = JSON.parse(raw); } catch(e) {}
    }

    if (texto && texto.trim()) {
      comentarios[idAtividade] = texto.trim();
    } else {
      delete comentarios[idAtividade];
    }

    props.setProperty('comentarios_atividade', JSON.stringify(comentarios));
    registrarLog('', idAtividade, 'COMENTARIO_ATV', 'comentario', '', texto || '(removido)');
    return { status: 'ok', mensagem: 'Comentario salvo' };
  } finally {
    lock.releaseLock();
  }
}

// ==================== COMENTARIOS RESULTADO ====================

function getComentariosResultado() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('comentarios_resultado');
  if (raw) {
    try { return JSON.parse(raw); } catch(e) {}
  }
  return {};
}

function salvarComentarioResultado(idResultado, texto) {
  var usuario = getUsuarioLogado();
  if (!usuario) return { status: 'erro', mensagem: 'Nao logado' };
  if (usuario.perfil !== 'SECRETARIO' && usuario.perfil !== 'PLANEJAMENTO') {
    return { status: 'erro', mensagem: 'Sem permissao' };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { status: 'erro', mensagem: 'Sistema ocupado, tente novamente' };
  }

  try {
    var props = PropertiesService.getScriptProperties();
    var comentarios = {};
    var raw = props.getProperty('comentarios_resultado');
    if (raw) {
      try { comentarios = JSON.parse(raw); } catch(e) {}
    }

    if (texto && texto.trim()) {
      comentarios[idResultado] = texto.trim();
    } else {
      delete comentarios[idResultado];
    }

    props.setProperty('comentarios_resultado', JSON.stringify(comentarios));
    registrarLog('', idResultado, 'COMENTARIO_RES', 'comentario', '', texto || '(removido)');
    return { status: 'ok', mensagem: 'Comentario salvo' };
  } finally {
    lock.releaseLock();
  }
}

// ==================== GANTT ====================

/**
 * Obter dados de atividades enriquecidos para o cronograma Gantt
 */
function getDadosGantt(filtros) {
  var atividades = listarAtividades(null);
  var resultados = listarResultados();
  var okrs = listarOKRs();

  // Mapas para enriquecimento
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

    // Aplicar filtros
    if (filtros) {
      if (filtros.diretoria && String(atv.diretoria) !== String(filtros.diretoria)) continue;
      if (filtros.id_oe && String(res.id_oe) !== String(filtros.id_oe)) continue;
      if (filtros.trimestre && String(okr.trimestre) !== String(filtros.trimestre)) continue;
      if (filtros.status && String(atv.status) !== String(filtros.status)) continue;
    }

    resultado.push({
      id: atv.id_atividade,
      descricao: atv.descricao,
      mes_planejado: atv.mes_planejado,
      mes_entrega: atv.mes_entrega || '',
      mes_realizado: atv.mes_realizado || '',
      status: atv.status,
      percentual: atv.percentual_execucao,
      peso: atv.peso,
      responsavel: atv.responsavel,
      diretoria: atv.diretoria,
      id_resultado: atv.id_resultado,
      resultado_descricao: res.descricao || '',
      id_oe: res.id_oe || '',
      trimestre: okr.trimestre || '',
      id_okr: atv.id_okr,
      okr_descricao: okr.descricao || ''
    });
  }

  // Ordenar por resultado, depois por mes_planejado
  resultado.sort(function(a, b) {
    if (a.id_resultado !== b.id_resultado) return String(a.id_resultado).localeCompare(String(b.id_resultado));
    return 0;
  });

  return resultado;
}

/**
 * Obter resultado completo no formato da aba da planilha
 */
function getResultadoPorAba(idResultado) {
  var resultados = listarResultados();
  var res = null;
  for (var r = 0; r < resultados.length; r++) {
    if (String(resultados[r].id_resultado) === String(idResultado)) {
      res = resultados[r];
      break;
    }
  }
  if (!res) return null;

  var okrs = listarOKRs(idResultado);
  var atividades = listarAtividades({ id_resultado: idResultado });

  var okrsComAtividades = [];
  for (var k = 0; k < okrs.length; k++) {
    var okr = okrs[k];
    var atvsOKR = atividades.filter(function(a) { return a.id_okr === okr.id_okr; });
    okrsComAtividades.push({
      okr: okr,
      atividades: atvsOKR
    });
  }

  return {
    resultado: res,
    okrs: okrsComAtividades
  };
}
