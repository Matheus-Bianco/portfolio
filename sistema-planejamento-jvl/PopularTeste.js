/**
 * ============================================================
 * POPULAR DADOS DE TESTE - DEMONSTRACAO PARA O SECRETARIO
 * ============================================================
 * Execute esta funcao UMA VEZ para preencher o sistema com
 * dados realistas baseados no PEI 2025-2029 da SED Joinville.
 *
 * Para executar: selecione "popularDadosTeste" e clique em Executar.
 * ============================================================
 */

function popularDadosTeste() {
  var ss = getSpreadsheet();

  // 1. Popular USUARIOS
  popularUsuariosTeste(ss);

  // 2. Popular OEs (ja existente, mas garantir)
  importarDadosPEI();

  // 3. Popular RESULTADOS
  popularResultadosTeste(ss);

  // 4. Popular OKRs
  popularOKRsTeste(ss);

  // 5. Popular ATIVIDADES
  popularAtividadesTeste(ss);

  return 'Dados de teste populados com sucesso! Sistema pronto para demonstracao.';
}

// ==================== USUARIOS DE TESTE ====================
function popularUsuariosTeste(ss) {
  var sheet = ss.getSheetByName('USUARIOS');

  // Limpar dados existentes (manter cabecalho)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  }

  var usuarios = [
    [1, 'admin@sed.joinville.sc.gov.br', 'Administrador do Sistema', 'PLANEJAMENTO', 'UPL', '123456'],
    [2, 'secretario@sed.joinville.sc.gov.br', 'Diego Calegari', 'SECRETARIO', 'SED', '123456'],
    [3, 'gerencia.dpe@sed.joinville.sc.gov.br', 'Giani Magali da Silva', 'GERENCIA', 'DPE', '123456'],
    [4, 'gerencia.dfi@sed.joinville.sc.gov.br', 'Cleberson de Lima Mendes', 'GERENCIA', 'DFI', '123456'],
    [5, 'gerencia.dge@sed.joinville.sc.gov.br', 'Andrei Popovski', 'GERENCIA', 'DGE', '123456'],
    [6, 'gerencia.din@sed.joinville.sc.gov.br', 'Felipe Hardt', 'GERENCIA', 'DIN', '123456'],
    [7, 'planejamento@sed.joinville.sc.gov.br', 'Jose Victor Goncalves', 'PLANEJAMENTO', 'UPL', '123456'],
    [8, 'gerencia.uaa@sed.joinville.sc.gov.br', 'Suelen Rodrigues', 'GERENCIA', 'UAA', '123456'],
    [9, 'gerencia.uit@sed.joinville.sc.gov.br', 'Camilla Borges', 'GERENCIA', 'UIT', '123456'],
    [10, 'gerencia.uef@sed.joinville.sc.gov.br', 'Gabriel Iwaya', 'GERENCIA', 'UEF', '123456'],
    [11, 'gerencia.udp@sed.joinville.sc.gov.br', 'Ckelen Ribeiro', 'GERENCIA', 'UDP', '123456']
  ];

  if (usuarios.length > 0) {
    sheet.getRange(2, 1, usuarios.length, usuarios[0].length).setValues(usuarios);
  }
}

// ==================== RESULTADOS DE TESTE ====================
function popularResultadosTeste(ss) {
  var sheet = ss.getSheetByName('RESULTADOS');

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  }

  // id_resultado, id_oe, descricao, meta_2025, meta_2029, classificacao, responsavel, macroprocesso, projeto_codigo, projeto_nome, diretoria, unidade, meta_2026, requisitos_qualidade
  var resultados = [
    // OI01
    ['R01', 'OI01', 'Ofertar ensino em tempo integral para percentual dos alunos da rede', '11,8% dos alunos', '15% dos alunos', 'Indicador estrategico', 'SED', 'Plano de Expansao', '', '', 'SED', 'SED', '', ''],
    ['R02', 'OI01', 'Reduzir a fila de criancas em espera por vagas', 'Reduzir em 15%', 'Reduzir em 100%', 'Indicador estrategico', 'SED', 'Plano de Expansao', '', '', 'SED', 'SED', '', ''],
    ['R03', 'OI01', 'Ampliar oferta de vagas para criancas de 0 a 3 anos', '40,43% das criancas', '51,02% das criancas', 'Indicador estrategico', 'SED', 'Plano de Expansao', '', '', 'SED', 'SED', '', ''],

    // OI02
    ['R04', 'OI02', 'Obter aprovacao nos Anos Iniciais e Anos Finais', '99,6% AI e 98,5% AF', '99,8% AI e 99% AF', 'Indicador estrategico', 'SED', '', '', '', 'SED', 'SED', '', ''],
    ['R05', 'OI02', 'Diminuir distorcao idade-serie nos AI do EF', '3,8%', '3,3%', 'Indicador estrategico', 'SED', '', '', '', 'SED', 'SED', '', ''],

    // OI03
    ['R06', 'OI03', 'Ter alunos alfabetizados no 2o ano', '75% dos alunos', '90% dos alunos', 'Indicador estrategico', 'SED', '', '', '', 'SED', 'SED', '', ''],
    ['R07', 'OI03', 'Alunos em nivel adequado de aprendizagem em LP', '78,63% 2o ano, 83,56% 5o ano, 62,49% 9o ano', '90% 2o e 5o ano, 75% 9o ano', 'Indicador estrategico', 'SED', '', '', '', 'SED', 'SED', '', ''],

    // OE01 - Expandir oferta de vagas
    ['R08', 'OE01', 'Estruturar o projeto de financiamento do BID', 'Projeto estruturado', 'Financiamento aprovado', 'Resultado de projeto', 'Jose Victor', 'Planejamento estrategico', 'PRJ.UPL.01', 'Plano de Expansao', 'UPL', 'UPL', '', ''],
    ['R09', 'OE01', 'Entregar 7 novos CEIs e ampliar 1 criando vagas de creche', '1.091 vagas parciais e 799 integrais', '6.600 parciais e 5.000 integrais EI', 'Indicador', 'Ademar Stringari', 'Infraestrutura', '', 'Plano de Expansao', 'DIN', 'UIN', '', ''],
    ['R10', 'OE01', 'Entrega do estudo de PPP para manutencao e ampliacao da estrutura fisica escolar', 'Estudo entregue', '33% das obras executadas', 'Resultado de melhoria de processo', 'Felipe Hardt', 'Infraestrutura', 'PRJ.UPL.01', 'Plano de Expansao', 'DIN', 'UIN', '', ''],

    // OE02 - Curriculo de excelencia
    ['R11', 'OE02', 'Realizar revisao piloto com foco em matematica nos Anos Iniciais', 'Revisao piloto realizada', 'Novo curriculo publicado', 'Resultado de projeto', 'Ckelen Ribeiro', 'Curriculo', 'PRJ.DPE.02', 'Curriculo de Excelencia', 'DPE', 'DPE', '', ''],
    ['R12', 'OE02', 'Desenvolver mapas de progressao da aprendizagem da Pre-Escola', 'Mapas desenvolvidos', 'Curriculo completo', 'Resultado de projeto', 'Juliano', 'Curriculo', 'PRJ.DPE.02', 'Curriculo de Excelencia', 'DPE', 'DPE', '', ''],
    ['R13', 'OE02', 'Iniciar atividades do CRIE com oficinas para alunos', 'CRIE iniciado', '2 CRIEs criados', 'Resultado de projeto', 'Denise', 'Tecnologias Educacionais', 'PRJ.UIT.04', 'CRIE', 'UIT', 'UIT', '', ''],
    ['R14', 'OE02', 'Atingir alunos atendidos no contraturno escolar', '4.000 alunos', 'Ampliar atendimento', 'Indicador', 'Juliana Alano', 'Tecnologias Educacionais', '', '', 'UIT', 'UIT', '', ''],

    // OE03 - Apoio a estudantes com dificuldades
    ['R15', 'OE03', 'Efetivar a primeira parceria de reforco escolar', 'Parceria efetivada', '95% das UEs atendidas', 'Resultado de projeto', 'Suelen', 'Reforco escolar', '', '', 'UAA', 'UAA', '', ''],
    ['R16', 'OE03', 'Ter alunos com media acima de 5 em LP e MAT', '99,3% dos alunos', '99,7% dos alunos', 'Indicador', 'Suelen', '', '', '', 'UAA', 'UAA', '', ''],

    // OE04 - Frequencia e engajamento
    ['R17', 'OE04', 'Ter professores lancando frequencia diariamente', '90% dos professores', '100% dos professores', 'Indicador', 'Claudia e Fernanda', 'Gestao escolar', '', '', 'DPE', 'DPE', '', ''],
    ['R18', 'OE04', 'Alcancar indice de abandono escolar', '0,1%', '0,06%', 'Indicador estrategico', 'Gabriel Iwaya', 'Ensino Fundamental', '', '', 'UEF', 'UEF', '', ''],

    // OE05 - Educacao Especial
    ['R19', 'OE05', 'Iniciar funcionamento do polo Pertencer', 'Polo em funcionamento', '99% dos alunos atendidos', 'Resultado de projeto', 'Roberta Engels', '', 'PRJ.UAA.03', 'Atendimento Especializado', 'UAA', 'UAA', '', ''],
    ['R20', 'OE05', 'Aprovar a Politica de Educacao Especial da rede municipal', 'Politica aprovada', 'Politica implementada', 'Resultado de projeto', 'Priscila Deud', '', 'PRJ.UAA.03', 'Atendimento Especializado', 'UAA', 'UAA', '', ''],

    // OE06 - Professores
    ['R21', 'OE06', 'Ofertar programa de formacao para novos professores com mentorias', '85% ciclo formativo completo', '99% ciclo completo, NPS 80', 'Resultado de projeto', 'UDP', 'Formacao continuada', '', '', 'UDP', 'UDP', '', ''],
    ['R22', 'OE06', 'NPS das formacoes continuadas', 'NPS de 75', 'NPS de 80', 'Indicador', 'UDP', 'Formacao continuada', '', '', 'UDP', 'UDP', '', ''],

    // OE07 - Materiais de apoio
    ['R23', 'OE07', 'Modelo de avaliacao formativa de rede concebido (2o ao 4o ano) em LP e MAT', 'Modelo concebido', 'Resultados em 5 dias uteis', 'Resultado de projeto', 'Debora', 'Avaliacoes', '', '', 'UDP', 'UDP', '', ''],
    ['R24', 'OE07', 'Disponibilizar resultados apos avaliacao externa', 'Em 14 dias', 'Em 5 dias uteis', 'Indicador', 'Ckelen', 'Avaliacoes', '', '', 'UDP', 'UDP', '', ''],

    // OE08 - Saude e bem-estar
    ['R25', 'OE08', 'Realizar piloto do Projeto Escola Saudavel em 10 unidades', 'Piloto realizado', 'Indices elevados de bem-estar', 'Resultado de projeto', 'DPE', 'Desenvolvimento integral', '', '', 'DPE', 'DPE', '', ''],

    // OE09 - Lideranca e gestao escolar
    ['R26', 'OE09', 'Ter equipe diretiva formada no programa Escola de Lideres', '90% formados', '100% ciclo completo EL', 'Resultado de projeto', 'DGE', 'Gestao escolar', '', '', 'DGE', 'DGE', '', ''],
    ['R27', 'OE09', 'Garantir execucao dos recursos repassados para as UExs', '75% dos recursos', '90% dos recursos', 'Indicador', 'DGE', 'Gestao financeira', '', '', 'DGE', 'DGE', '', ''],

    // OE10 - Gestao eficiente
    ['R28', 'OE10', 'Implementar sistema integrado de acompanhamento de resultados', 'Sistema implementado', 'Sistema consolidado', 'Resultado de projeto', 'Jose Victor', 'Monitoramento de indicadores', '', '', 'UPL', 'UPL', '', ''],
    ['R29', 'OE10', 'Contratar novo sistema de gestao academica', 'Sistema contratado', 'Integracao completa', 'Resultado de projeto', 'UPL', 'Tecnologia da Informacao', '', '', 'UPL', 'UPL', '', ''],
    ['R30', 'OE10', 'Reduzir tempo medio de processos licitatorios', 'Reduzir em 20%', 'Processo otimizado', 'Indicador', 'DIN', 'Compras', '', '', 'DIN', 'DIN', '', ''],

    // OE11 - Infraestrutura
    ['R31', 'OE11', 'Instalar pisos modulares nas quadras de 6 unidades escolares', '6 quadras', '100% escolas com quadras cobertas', 'Resultado de projeto', 'DIN', 'Infraestrutura', '', '', 'DIN', 'DIN', '', ''],
    ['R32', 'OE11', 'Ampliar unidades com adequacao dos bombeiros', '80 unidades', 'Todas adequadas', 'Indicador', 'DIN', 'Manutencao das UEs', '', '', 'DIN', 'DIN', '', ''],

    // OE12 - Servidores
    ['R33', 'OE12', 'Atingir taxa de ausencia de professores e auxiliares inferior a 1%', 'Inferior a 1%', 'Inferior a 0,5%', 'Indicador', 'DGE', 'Contratacao', '', '', 'DGE', 'DGE', '', ''],
    ['R34', 'OE12', 'Abrir inscricoes para novo concurso publico da SED', 'Inscricoes abertas', '90% concursados', 'Resultado de projeto', 'DGE', 'Contratacao', '', '', 'DGE', 'DGE', '', '']
  ];

  if (resultados.length > 0) {
    sheet.getRange(2, 1, resultados.length, resultados[0].length).setValues(resultados);
  }
}

// ==================== OKRs DE TESTE ====================
function popularOKRsTeste(ss) {
  var sheet = ss.getSheetByName('OKRS');

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  }

  // id_okr, id_resultado, id_oe, trimestre, descricao, situacao, percentual, requisitos_qualidade, ano, diretoria, unidade, eixo
  var okrs = [
    // ---- OE01: Plano de Expansao ----
    ['OKR001', 'R08', 'OE01', '1o', 'Finalizar o PEP e preencher a PACI', 'Concluido', 100, 'Documentacao completa e validada', 2025, 'UPL', 'UPL', 'ALUNOS'],
    ['OKR002', 'R08', 'OE01', '2o', 'Aprovar o POD pelo BID', 'Parcialmente concluido', 80, 'Aprovacao formal do BID', 2025, 'UPL', 'UPL', 'ALUNOS'],
    ['OKR003', 'R08', 'OE01', '3o', 'Aprovar a Proposta de Emprestimo (OPC)', 'Concluido', 100, 'Validacao do conselho', 2025, 'UPL', 'UPL', 'ALUNOS'],
    ['OKR004', 'R08', 'OE01', '4o', 'Aprovar o projeto no STN', 'Nao concluido', 25, 'Aprovacao do Tesouro Nacional', 2025, 'UPL', 'UPL', 'ALUNOS'],
    ['OKR005', 'R10', 'OE01', '1o', 'Definir plano de trabalho junto a SP Parcerias', 'Concluido', 100, 'Plano de trabalho validado', 2025, 'DIN', 'UIN', 'ALUNOS'],
    ['OKR006', 'R10', 'OE01', '2o', 'Elaborar Relatorio de Pre Viabilidade', 'Parcialmente concluido', 50, 'Qualidade tecnica do relatorio', 2025, 'DIN', 'UIN', 'ALUNOS'],
    ['OKR007', 'R10', 'OE01', '3o', 'Elaborar Relatorio de Pre Viabilidade V.Final', 'Concluido', 100, 'Versao definitiva entregue', 2025, 'DIN', 'UIN', 'ALUNOS'],
    ['OKR008', 'R10', 'OE01', '4o', 'Entregar documentos de Consulta Publica', 'No prazo', 100, 'Documentos validados pela diretoria', 2025, 'DIN', 'UIN', 'ALUNOS'],

    // ---- OE02: Curriculo de Excelencia ----
    ['OKR009', 'R11', 'OE02', '2o', 'Aprovacao da reforma administrativa', 'Concluido', 100, 'Reforma aprovada formalmente', 2025, 'DPE', 'DPE', 'ALUNOS'],
    ['OKR010', 'R11', 'OE02', '3o', 'Resolucao preliminar do CME (Ingles, CH e BNCC Computacao)', 'Concluido', 80, 'Resolucao aprovada pelo CME', 2025, 'DPE', 'DPE', 'ALUNOS'],
    ['OKR011', 'R11', 'OE02', '3o', 'Contratacao da Camino', 'Parcialmente concluido', 67, 'Contrato assinado', 2025, 'DPE', 'DPE', 'ALUNOS'],
    ['OKR012', 'R11', 'OE02', '4o', 'Contratacao da Camino - parecer PGM', 'Em andamento', 30, 'Parecer juridico favoravel', 2025, 'DPE', 'DPE', 'ALUNOS'],
    ['OKR013', 'R12', 'OE02', '2o', 'V.F dos indicadores dos marcos de desenvolvimento da EI', 'Concluido', 100, 'Validacao com unidades e equipe', 2025, 'DPE', 'DPE', 'ALUNOS'],
    ['OKR014', 'R12', 'OE02', '4o', 'Producao dos mapas de progressao de linguagens na EI', 'Em atraso', 50, 'Qualidade pedagogica e validacao', 2025, 'DPE', 'DPE', 'ALUNOS'],
    ['OKR015', 'R13', 'OE02', '3o', 'Definir layout de escola em tempo integral', 'Concluido', 100, 'Planta padrao validada', 2025, 'UIT', 'UIT', 'ALUNOS'],
    ['OKR016', 'R13', 'OE02', '4o', 'Enviar processo de contratacao da Motriz para SAP', 'Parcialmente concluido', 25, 'TR elaborado e validado', 2025, 'UIT', 'UIT', 'ALUNOS'],

    // ---- OE03: Apoio a aprendizagem ----
    ['OKR017', 'R15', 'OE03', '1o', 'Estruturar modelo integrado de recomposicao', 'Concluido', 100, 'Modelo conceitual definido', 2025, 'UAA', 'UAA', 'ALUNOS'],
    ['OKR018', 'R15', 'OE03', '2o', 'Implementar programa em unidades piloto', 'Em andamento', 60, 'Atendimento adequado nas UEs piloto', 2025, 'UAA', 'UAA', 'ALUNOS'],
    ['OKR019', 'R15', 'OE03', '3o', 'Expandir para 50% das unidades com direito ao Aprender Mais', 'Em andamento', 45, '50% das UEs atendidas', 2025, 'UAA', 'UAA', 'ALUNOS'],
    ['OKR020', 'R15', 'OE03', '4o', 'Avaliar resultados e ajustar modelo', 'No prazo', 10, 'Relatorio de avaliacao', 2025, 'UAA', 'UAA', 'ALUNOS'],

    // ---- OE04: Frequencia e engajamento ----
    ['OKR021', 'R17', 'OE04', '2o', 'Atingir 70% dos professores lancando frequencia', 'Concluido', 100, 'Adesao mensurada no sistema', 2025, 'DPE', 'DPE', 'ALUNOS'],
    ['OKR022', 'R17', 'OE04', '3o', 'Atingir 80% dos professores lancando frequencia', 'Concluido', 100, 'Monitoramento mensal', 2025, 'DPE', 'DPE', 'ALUNOS'],
    ['OKR023', 'R17', 'OE04', '4o', 'Atingir 90% dos professores lancando frequencia', 'Em andamento', 85, 'Meta anual de 90%', 2025, 'DPE', 'DPE', 'ALUNOS'],

    // ---- OE05: Educacao Especial ----
    ['OKR024', 'R19', 'OE05', '2o', 'Iniciar atendimento do polo Pertencer', 'Parcialmente concluido', 50, 'Atendimento iniciado com profissionais', 2025, 'UAA', 'UAA', 'ALUNOS'],
    ['OKR025', 'R19', 'OE05', '3o', 'Iniciar atendimento aos alunos no Pertencer', 'Concluido', 100, 'Fluxo de atendimento organizado', 2025, 'UAA', 'UAA', 'ALUNOS'],
    ['OKR026', 'R19', 'OE05', '4o', '1o diagnostico de avaliacao do servico', 'No prazo', 100, 'Instrumento de avaliacao aplicado', 2025, 'UAA', 'UAA', 'ALUNOS'],
    ['OKR027', 'R20', 'OE05', '2o', 'V.1 da Politica de Educacao Especial', 'Parcialmente concluido', 67, 'Politica alinhada as diretrizes', 2025, 'UAA', 'UAA', 'ALUNOS'],
    ['OKR028', 'R20', 'OE05', '3o', 'Validar V.1 da Politica', 'Nao concluido', 0, 'Validacao pela comissao', 2025, 'UAA', 'UAA', 'ALUNOS'],
    ['OKR029', 'R20', 'OE05', '4o', 'Validar V.1 com comissao da politica', 'Parcialmente concluido', 50, 'Aprovacao formal', 2025, 'UAA', 'UAA', 'ALUNOS'],

    // ---- OE06: Formacao de professores ----
    ['OKR030', 'R21', 'OE06', '1o', 'Definir estrutura do programa de formacao', 'Concluido', 100, 'Programa estruturado', 2025, 'UDP', 'UDP', 'PROFESSORES'],
    ['OKR031', 'R21', 'OE06', '2o', 'Iniciar ciclo formativo com 1a turma', 'Concluido', 100, 'Turma formada e em andamento', 2025, 'UDP', 'UDP', 'PROFESSORES'],
    ['OKR032', 'R21', 'OE06', '3o', 'Atingir 60% do ciclo formativo completo', 'Em andamento', 72, '85% ate fim do ano', 2025, 'UDP', 'UDP', 'PROFESSORES'],
    ['OKR033', 'R21', 'OE06', '4o', 'Atingir 85% do ciclo formativo completo', 'Em andamento', 55, 'NPS minimo de 75', 2025, 'UDP', 'UDP', 'PROFESSORES'],

    // ---- OE07: Avaliacao formativa ----
    ['OKR034', 'R23', 'OE07', '1o', 'Estruturar instrumentos tecnicos para contratacao da parceria', 'Concluido', 100, 'Aderencia as diretrizes pedagogicas', 2025, 'UDP', 'UDP', 'PROFESSORES'],
    ['OKR035', 'R23', 'OE07', '2o', 'Construir e estruturar banco inicial de questoes', 'Concluido', 100, 'Qualidade pedagogica dos itens', 2025, 'UDP', 'UDP', 'PROFESSORES'],
    ['OKR036', 'R23', 'OE07', '3o', 'Validar sistema por meio de aplicacao piloto', 'Concluido', 100, 'Funcionamento operacional', 2025, 'UDP', 'UDP', 'PROFESSORES'],
    ['OKR037', 'R23', 'OE07', '4o', 'Documentar e institucionalizar o modelo avaliativo', 'Em andamento', 40, 'Completude documental', 2025, 'UDP', 'UDP', 'PROFESSORES'],

    // ---- OE08: Bem-estar ----
    ['OKR038', 'R25', 'OE08', '2o', 'Realizar primeiro diagnostico de clima e bem-estar', 'Em atraso', 30, 'Instrumento validado', 2025, 'DPE', 'DPE', 'GESTAO ESCOLAR'],
    ['OKR039', 'R25', 'OE08', '3o', 'Iniciar piloto Escola Saudavel em 5 unidades', 'Em andamento', 60, 'Adesao das unidades', 2025, 'DPE', 'DPE', 'GESTAO ESCOLAR'],

    // ---- OE09: Gestao escolar ----
    ['OKR040', 'R26', 'OE09', '1o', 'Estruturar modulos da Escola de Lideres', 'Concluido', 100, 'Modulos pedagogicamente consistentes', 2025, 'DGE', 'DGE', 'GESTAO ESCOLAR'],
    ['OKR041', 'R26', 'OE09', '2o', 'Formar 1a turma de 50 gestores na EL', 'Concluido', 100, 'Formacao completa', 2025, 'DGE', 'DGE', 'GESTAO ESCOLAR'],
    ['OKR042', 'R26', 'OE09', '3o', 'Formar 2a turma e atingir 70% do total', 'Em andamento', 78, '90% ate fim do ano', 2025, 'DGE', 'DGE', 'GESTAO ESCOLAR'],
    ['OKR043', 'R27', 'OE09', '3o', 'Atingir 60% de execucao dos recursos das UExs', 'Em andamento', 65, 'Execucao financeira monitorada', 2025, 'DGE', 'DGE', 'GESTAO ESCOLAR'],

    // ---- OE10: Gestao eficiente ----
    ['OKR044', 'R28', 'OE10', '1o', 'Definir requisitos do sistema de monitoramento', 'Concluido', 100, 'Requisitos validados com equipe', 2025, 'UPL', 'UPL', 'CAPACIDADE INSTITUCIONAL'],
    ['OKR045', 'R28', 'OE10', '2o', 'Desenvolver MVP do sistema de monitoramento', 'Concluido', 100, 'Sistema funcional basico', 2025, 'UPL', 'UPL', 'CAPACIDADE INSTITUCIONAL'],
    ['OKR046', 'R28', 'OE10', '3o', 'Implantar sistema com dados reais', 'Em andamento', 70, 'Sistema em producao', 2025, 'UPL', 'UPL', 'CAPACIDADE INSTITUCIONAL'],
    ['OKR047', 'R29', 'OE10', '2o', 'Elaborar TR para novo sistema de gestao academica', 'Concluido', 100, 'TR aprovado', 2025, 'UPL', 'UPL', 'CAPACIDADE INSTITUCIONAL'],
    ['OKR048', 'R29', 'OE10', '3o', 'Processo licitatorio do sistema academico', 'Em andamento', 40, 'Contratacao efetivada', 2025, 'UPL', 'UPL', 'CAPACIDADE INSTITUCIONAL'],

    // ---- OE11: Infraestrutura ----
    ['OKR049', 'R31', 'OE11', '2o', 'Instalar pisos modulares em 3 quadras', 'Concluido', 100, 'Instalacao conforme especificacao', 2025, 'DIN', 'DIN', 'CAPACIDADE INSTITUCIONAL'],
    ['OKR050', 'R31', 'OE11', '4o', 'Instalar pisos modulares nas 3 quadras restantes', 'Em andamento', 50, 'Cronograma cumprido', 2025, 'DIN', 'DIN', 'CAPACIDADE INSTITUCIONAL'],
    ['OKR051', 'R32', 'OE11', '3o', 'Contratar servico de adequacao bombeiros de 10 UEs', 'Parcialmente concluido', 60, 'Contratacao e execucao', 2025, 'DIN', 'DIN', 'CAPACIDADE INSTITUCIONAL'],

    // ---- OE12: Servidores ----
    ['OKR052', 'R33', 'OE12', '2o', 'Implementar sistema de monitoramento de ausencias', 'Concluido', 100, 'Sistema automatizado', 2025, 'DGE', 'DGE', 'CAPACIDADE INSTITUCIONAL'],
    ['OKR053', 'R33', 'OE12', '4o', 'Taxa de ausencia inferior a 1%', 'Em andamento', 70, 'Monitoramento continuo', 2025, 'DGE', 'DGE', 'CAPACIDADE INSTITUCIONAL'],
    ['OKR054', 'R34', 'OE12', '3o', 'Publicar edital do concurso publico', 'Concluido', 100, 'Edital publicado', 2025, 'DGE', 'DGE', 'CAPACIDADE INSTITUCIONAL'],
    ['OKR055', 'R34', 'OE12', '4o', 'Realizar provas do concurso', 'No prazo', 20, 'Processo seletivo justo', 2025, 'DGE', 'DGE', 'CAPACIDADE INSTITUCIONAL']
  ];

  if (okrs.length > 0) {
    sheet.getRange(2, 1, okrs.length, okrs[0].length).setValues(okrs);
  }
}

// ==================== ATIVIDADES DE TESTE ====================
function popularAtividadesTeste(ss) {
  var sheet = ss.getSheetByName('ATIVIDADES');

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
  }

  // id_atividade, id_okr, id_resultado, id_oe, descricao, mes_planejado, mes_realizado, criterio, peso, responsavel, observacoes, status, diretoria, unidade, eixo, percentual_execucao, mes_entrega
  var atividades = [
    // === OE01: Plano de Expansao (UPL / DIN) ===
    ['ATV001', 'OKR001', 'R08', 'OE01', 'Finalizacao do PEP', 'Abril', 'Abril', 'Qualitativo', 0.2, 'Jose Victor', '', 'Concluido', 'UPL', 'UPL', 'ALUNOS', 100, ''],
    ['ATV002', 'OKR001', 'R08', 'OE01', 'Preencher a PACI', 'Maio', 'Maio', 'Qualitativo', 0.2, 'Jose Victor', '', 'Concluido', 'UPL', 'UPL', 'ALUNOS', 100, ''],
    ['ATV003', 'OKR001', 'R08', 'OE01', 'Minuta de lei autorizativa', 'Maio', 'Maio', 'Qualitativo', 0.2, 'Jose Victor', '', 'Concluido', 'UPL', 'UPL', 'ALUNOS', 100, ''],
    ['ATV004', 'OKR001', 'R08', 'OE01', 'Realizacao de consulta publica', 'Junho', 'Junho', 'Qualitativo', 0.2, 'Jose Victor', '', 'Concluido', 'UPL', 'UPL', 'ALUNOS', 100, ''],
    ['ATV005', 'OKR002', 'R08', 'OE01', 'Enviar o POD para aprovacao do BID', 'Junho', '', 'Qualitativo', 0.2, 'Jose Victor', 'Replanejado para agosto por ajustes solicitados pelo BID', 'Replanejado', 'UPL', 'UPL', 'ALUNOS', 80, ''],
    ['ATV006', 'OKR003', 'R08', 'OE01', 'Aprovacao da Lei Autorizativa', 'Agosto', 'Julho', 'Qualitativo', 0.33, 'Juliana', '', 'Concluido', 'UPL', 'UPL', 'ALUNOS', 100, ''],
    ['ATV007', 'OKR003', 'R08', 'OE01', 'Aprovacao do POD (QRR)', 'Setembro', 'Setembro', 'Qualitativo', 0.33, 'Juliana', '', 'Concluido', 'UPL', 'UPL', 'ALUNOS', 100, ''],
    ['ATV008', 'OKR004', 'R08', 'OE01', 'Ter a minuta final do contrato de financiamento', 'Outubro', 'Outubro', 'Qualitativo', 0.25, 'Jose Victor', '', 'Concluido', 'UPL', 'UPL', 'ALUNOS', 100, ''],
    ['ATV009', 'OKR004', 'R08', 'OE01', 'Validar minuta final com conselho do BID', 'Outubro', '', 'Qualitativo', 0.25, 'Jose Victor', 'Aguardando retorno do BID', 'Atrasado', 'UPL', 'UPL', 'ALUNOS', 20, ''],
    ['ATV010', 'OKR004', 'R08', 'OE01', 'Solicitar verificacao de limite junto ao Tesouro Nacional', 'Outubro', '', 'Qualitativo', 0.25, 'Jose Victor', 'Depende da validacao anterior', 'Atrasado', 'UPL', 'UPL', 'ALUNOS', 0, ''],
    ['ATV011', 'OKR004', 'R08', 'OE01', 'Aprovar o projeto no STN', 'Dezembro', '', 'Qualitativo', 0.25, 'Jose Victor', '', 'No prazo', 'UPL', 'UPL', 'ALUNOS', 0, ''],

    // PPP (DIN)
    ['ATV012', 'OKR005', 'R10', 'OE01', 'Definir plano de trabalho com SP Parcerias', 'Fevereiro', 'Fevereiro', 'Qualitativo', 0.5, 'Felipe Hardt', '', 'Concluido', 'DIN', 'UIN', 'ALUNOS', 100, ''],
    ['ATV013', 'OKR005', 'R10', 'OE01', 'Realizar oficinas para relatorio de pre viabilidade', 'Marco', 'Marco', 'Qualitativo', 0.5, 'Felipe Hardt', '', 'Concluido', 'DIN', 'UIN', 'ALUNOS', 100, ''],
    ['ATV014', 'OKR007', 'R10', 'OE01', 'Revisao e validacao junto a diretoria de Infraestrutura', 'Agosto', 'Agosto', 'Qualitativo', 0.5, 'Felipe Hardt', '', 'Concluido', 'DIN', 'UIN', 'ALUNOS', 100, ''],
    ['ATV015', 'OKR007', 'R10', 'OE01', 'Entregar versao definitiva do Relatorio de Pre Viabilidade', 'Agosto', 'Agosto', 'Qualitativo', 0.5, 'Felipe Hardt', '', 'Concluido', 'DIN', 'UIN', 'ALUNOS', 100, ''],
    ['ATV016', 'OKR008', 'R10', 'OE01', 'Elaborar V.0 dos documentos para consulta publica', 'Novembro', 'Outubro', 'Qualitativo', 0.33, 'Felipe Hardt', 'Entregue antes do prazo', 'Concluido', 'DIN', 'UIN', 'ALUNOS', 100, ''],
    ['ATV017', 'OKR008', 'R10', 'OE01', 'Validar documentos junto a diretoria de Infraestrutura', 'Novembro', 'Outubro', 'Qualitativo', 0.33, 'Felipe Hardt', '', 'Concluido', 'DIN', 'UIN', 'ALUNOS', 100, ''],

    // === OE02: Curriculo (DPE / UIT) ===
    ['ATV018', 'OKR009', 'R11', 'OE02', 'Aprovacao da reforma administrativa', 'Junho', 'Junho', 'Qualitativo', 1, 'Adilson', '', 'Concluido', 'DPE', 'DPE', 'ALUNOS', 100, ''],
    ['ATV019', 'OKR010', 'R11', 'OE02', 'Apresentar propostas de alteracoes ao CME', 'Julho', 'Julho', 'Qualitativo', 0.2, 'Ckelen', '', 'Concluido', 'DPE', 'DPE', 'ALUNOS', 100, ''],
    ['ATV020', 'OKR010', 'R11', 'OE02', 'Entregar V.1 do curriculo complementar BNCC Computacao', 'Julho', 'Julho', 'Qualitativo', 0.2, 'Camilla', '', 'Concluido', 'UIT', 'UIT', 'ALUNOS', 100, ''],
    ['ATV021', 'OKR010', 'R11', 'OE02', 'Aprovar curriculo BNCC de Computacao no CME', 'Agosto', 'Agosto', 'Qualitativo', 0.2, 'Ckelen', '', 'Concluido', 'DPE', 'DPE', 'ALUNOS', 100, ''],
    ['ATV022', 'OKR011', 'R11', 'OE02', 'Validar proposta comercial com equipe de curriculo', 'Julho', 'Julho', 'Qualitativo', 0.33, 'Cleberson', '', 'Concluido', 'DPE', 'DPE', 'ALUNOS', 100, ''],
    ['ATV023', 'OKR011', 'R11', 'OE02', 'Elaborar termo de referencia', 'Agosto', 'Agosto', 'Qualitativo', 0.33, 'Cleberson', '', 'Concluido', 'DPE', 'DPE', 'ALUNOS', 100, ''],
    ['ATV024', 'OKR011', 'R11', 'OE02', 'Enviar processo para SAP', 'Agosto', '', 'Qualitativo', 0.33, 'Cleberson', 'Faltou priorizacao da equipe para o TR', 'Replanejado', 'DPE', 'DPE', 'ALUNOS', 50, ''],
    ['ATV025', 'OKR012', 'R11', 'OE02', 'Enviar processo para SAP', 'Outubro', 'Outubro', 'Qualitativo', 0.33, 'Cleberson', '', 'Concluido', 'DPE', 'DPE', 'ALUNOS', 100, ''],
    ['ATV026', 'OKR012', 'R11', 'OE02', 'Ter o parecer da PGM', 'Dezembro', '', 'Qualitativo', 0.33, 'Cleberson', '', 'Em andamento', 'DPE', 'DPE', 'ALUNOS', 30, ''],
    ['ATV027', 'OKR012', 'R11', 'OE02', 'Assinar o contrato', 'Dezembro', '', 'Qualitativo', 0.33, 'Cleberson', '', 'No prazo', 'DPE', 'DPE', 'ALUNOS', 0, ''],

    // Mapas de progressao (DPE)
    ['ATV028', 'OKR013', 'R12', 'OE02', 'V.2 dos indicadores dos marcos de desenvolvimento', 'Abril', 'Abril', 'Qualitativo', 0.25, 'Juliano', '', 'Concluido', 'DPE', 'DPE', 'ALUNOS', 100, ''],
    ['ATV029', 'OKR013', 'R12', 'OE02', 'Validacao com unidades escolares', 'Maio', 'Maio', 'Qualitativo', 0.25, 'Juliano', '', 'Concluido', 'DPE', 'DPE', 'ALUNOS', 100, ''],
    ['ATV030', 'OKR013', 'R12', 'OE02', 'V.F dos indicadores', 'Junho', 'Junho', 'Qualitativo', 0.25, 'Juliano', '', 'Concluido', 'DPE', 'DPE', 'ALUNOS', 100, ''],
    ['ATV031', 'OKR014', 'R12', 'OE02', 'Definir cronograma de producao dos mapas', 'Outubro', 'Outubro', 'Qualitativo', 0.25, 'Juliano', '', 'Concluido', 'DPE', 'DPE', 'ALUNOS', 100, ''],
    ['ATV032', 'OKR014', 'R12', 'OE02', 'Entregar V.1 dos mapas de progressao', 'Novembro', 'Novembro', 'Qualitativo', 0.25, 'Juliano', '', 'Concluido', 'DPE', 'DPE', 'ALUNOS', 100, ''],
    ['ATV033', 'OKR014', 'R12', 'OE02', 'Validar V.1 com especialistas', 'Novembro', '', 'Qualitativo', 0.25, 'Juliano', 'Em andamento - aguardando agenda dos profissionais', 'Em andamento', 'DPE', 'DPE', 'ALUNOS', 60, ''],
    ['ATV034', 'OKR014', 'R12', 'OE02', 'Entregar V.2 diagramada', 'Dezembro', '', 'Qualitativo', 0.25, 'Juliano', '', 'No prazo', 'DPE', 'DPE', 'ALUNOS', 0, ''],

    // CRIE (UIT)
    ['ATV035', 'OKR015', 'R13', 'OE02', 'Revisao da planta da unidade', 'Julho', 'Julho', 'Qualitativo', 0.33, 'Denise', '', 'Concluido', 'UIT', 'UIT', 'ALUNOS', 100, ''],
    ['ATV036', 'OKR015', 'R13', 'OE02', 'Elaborar planta padrao de escola em tempo integral', 'Julho', 'Julho', 'Qualitativo', 0.33, 'Denise', '', 'Concluido', 'UIT', 'UIT', 'ALUNOS', 100, ''],
    ['ATV037', 'OKR015', 'R13', 'OE02', 'Validar planta com o Secretario', 'Julho', 'Julho', 'Qualitativo', 0.33, 'Denise', '', 'Concluido', 'UIT', 'UIT', 'ALUNOS', 100, ''],
    ['ATV038', 'OKR016', 'R13', 'OE02', 'Definir escopo da contratacao Motriz', 'Outubro', 'Outubro', 'Qualitativo', 0.25, 'Juliana Alano', '', 'Concluido', 'UIT', 'UIT', 'ALUNOS', 100, ''],
    ['ATV039', 'OKR016', 'R13', 'OE02', 'Validar plano de trabalho com Secretario', 'Outubro', '', 'Qualitativo', 0.25, 'Juliana Alano', 'Em analise', 'Em andamento', 'UIT', 'UIT', 'ALUNOS', 50, ''],
    ['ATV040', 'OKR016', 'R13', 'OE02', 'Elaborar o TR', 'Dezembro', '', 'Qualitativo', 0.25, 'Juliana Alano', '', 'No prazo', 'UIT', 'UIT', 'ALUNOS', 0, ''],
    ['ATV041', 'OKR016', 'R13', 'OE02', 'Enviar processo para SAP', 'Dezembro', '', 'Qualitativo', 0.25, 'Juliana Alano', '', 'No prazo', 'UIT', 'UIT', 'ALUNOS', 0, ''],

    // === OE03: Reforco escolar (UAA) ===
    ['ATV042', 'OKR017', 'R15', 'OE03', 'Elaborar documento-base do modelo integrado', 'Janeiro', 'Janeiro', 'Qualitativo', 0.33, 'Suelen', '', 'Concluido', 'UAA', 'UAA', 'ALUNOS', 100, ''],
    ['ATV043', 'OKR017', 'R15', 'OE03', 'Definir criterios pedagogicos de classificacao', 'Janeiro', 'Janeiro', 'Qualitativo', 0.33, 'Suelen', '', 'Concluido', 'UAA', 'UAA', 'ALUNOS', 100, ''],
    ['ATV044', 'OKR017', 'R15', 'OE03', 'Construir fluxograma de atendimento', 'Janeiro', 'Fevereiro', 'Qualitativo', 0.33, 'Suelen', 'Entregue com pequeno atraso', 'Concluido', 'UAA', 'UAA', 'ALUNOS', 100, ''],
    ['ATV045', 'OKR018', 'R15', 'OE03', 'Definir instrumentos de diagnostico', 'Fevereiro', 'Fevereiro', 'Qualitativo', 0.33, 'Suelen', '', 'Concluido', 'UAA', 'UAA', 'ALUNOS', 100, ''],
    ['ATV046', 'OKR018', 'R15', 'OE03', 'Realizar formacao para uso do material didatico', 'Fevereiro', 'Marco', 'Qualitativo', 0.33, 'Suelen', '', 'Concluido', 'UAA', 'UAA', 'ALUNOS', 100, ''],
    ['ATV047', 'OKR018', 'R15', 'OE03', 'Selecionar materiais didaticos por nivel', 'Marco', '', 'Qualitativo', 0.33, 'Suelen', 'Processo em andamento', 'Em andamento', 'UAA', 'UAA', 'ALUNOS', 40, ''],
    ['ATV048', 'OKR019', 'R15', 'OE03', 'Expandir atendimento para novas UEs', 'Agosto', '', 'Qualitativo', 0.5, 'Suelen', 'Meta de 50% das UEs', 'Em andamento', 'UAA', 'UAA', 'ALUNOS', 45, ''],
    ['ATV049', 'OKR019', 'R15', 'OE03', 'Monitorar resultados das UEs atendidas', 'Setembro', '', 'Qualitativo', 0.5, 'Suelen', '', 'Em andamento', 'UAA', 'UAA', 'ALUNOS', 30, ''],

    // === OE05: Educacao Especial (UAA) ===
    ['ATV050', 'OKR024', 'R19', 'OE05', 'Inicio das formacoes dos profissionais do polo', 'Abril', 'Abril', 'Qualitativo', 0.25, 'Priscila Deud', '', 'Concluido', 'UAA', 'UAA', 'ALUNOS', 100, ''],
    ['ATV051', 'OKR024', 'R19', 'OE05', 'Organizacao da infraestrutura do espaco', 'Maio', 'Maio', 'Qualitativo', 0.25, 'Priscila Deud', '', 'Concluido', 'UAA', 'UAA', 'ALUNOS', 100, ''],
    ['ATV052', 'OKR024', 'R19', 'OE05', 'Publicar protocolos e documentacao', 'Maio', '', 'Qualitativo', 0.25, 'Priscila Deud', 'Replanejado por falta de RH', 'Replanejado', 'UAA', 'UAA', 'ALUNOS', 40, ''],
    ['ATV053', 'OKR025', 'R19', 'OE05', 'Organizar fluxo de atendimento', 'Agosto', 'Agosto', 'Qualitativo', 0.33, 'Priscila Deud', '', 'Concluido', 'UAA', 'UAA', 'ALUNOS', 100, ''],
    ['ATV054', 'OKR025', 'R19', 'OE05', 'Definir priorizacao dos atendimentos', 'Agosto', 'Agosto', 'Qualitativo', 0.33, 'Priscila Deud', '', 'Concluido', 'UAA', 'UAA', 'ALUNOS', 100, ''],
    ['ATV055', 'OKR025', 'R19', 'OE05', 'Iniciar atendimentos', 'Setembro', 'Setembro', 'Qualitativo', 0.33, 'Priscila Deud', '', 'Concluido', 'UAA', 'UAA', 'ALUNOS', 100, ''],
    ['ATV056', 'OKR027', 'R20', 'OE05', 'V.0 da politica para analise do setor de politicas', 'Abril', 'Abril', 'Qualitativo', 0.33, 'Priscila Deud', '', 'Concluido', 'UAA', 'UAA', 'ALUNOS', 100, ''],
    ['ATV057', 'OKR027', 'R20', 'OE05', 'Criar cronograma da elaboracao da politica', 'Maio', 'Maio', 'Qualitativo', 0.33, 'Priscila Deud', '', 'Concluido', 'UAA', 'UAA', 'ALUNOS', 100, ''],
    ['ATV058', 'OKR027', 'R20', 'OE05', 'Elaborar V.01 da politica', 'Junho', '', 'Qualitativo', 0.33, 'Priscila Deud', 'Dificuldade por volume de demandas', 'Atrasado', 'UAA', 'UAA', 'ALUNOS', 30, ''],
    ['ATV059', 'OKR029', 'R20', 'OE05', 'Elaborar V.01 da politica', 'Novembro', 'Novembro', 'Qualitativo', 0.5, 'Priscila Deud', '', 'Concluido', 'UAA', 'UAA', 'ALUNOS', 100, ''],
    ['ATV060', 'OKR029', 'R20', 'OE05', 'Validar V.01 com comissao da politica', 'Dezembro', '', 'Qualitativo', 0.5, 'Priscila Deud', 'Aguardando validacao', 'Em andamento', 'UAA', 'UAA', 'ALUNOS', 20, ''],

    // === OE07: Avaliacao formativa (UDP) ===
    ['ATV061', 'OKR034', 'R23', 'OE07', 'Analise comparativa tecnica-pedagogica das plataformas', 'Janeiro', 'Janeiro', 'Qualitativo', 0.2, 'Nicolas', '', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV062', 'OKR034', 'R23', 'OE07', 'Definicao da plataforma avaliativa', 'Fevereiro', 'Fevereiro', 'Qualitativo', 0.2, 'Ckelen', '', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV063', 'OKR034', 'R23', 'OE07', 'Elaboracao do DFDi', 'Fevereiro', 'Fevereiro', 'Qualitativo', 0.2, 'Nicolas', '', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV064', 'OKR034', 'R23', 'OE07', 'Elaboracao do Termo de Referencia', 'Fevereiro', 'Fevereiro', 'Qualitativo', 0.2, 'Nicolas', '', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV065', 'OKR034', 'R23', 'OE07', 'Encaminhamento do processo ao setor de compras', 'Marco', 'Marco', 'Qualitativo', 0.2, 'Nicolas', '', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV066', 'OKR035', 'R23', 'OE07', 'Elaboracao do construto das avaliacoes diagnosticas e formativas', 'Abril', 'Abril', 'Qualitativo', 0.25, 'Debora', '', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV067', 'OKR035', 'R23', 'OE07', 'Elaboracao das questoes de LP e MT para avaliacao diagnostica', 'Maio', 'Maio', 'Qualitativo', 0.25, 'Daiane', '', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV068', 'OKR035', 'R23', 'OE07', 'Revisao tecnica e pedagogica das questoes', 'Junho', 'Junho', 'Qualitativo', 0.25, 'Silvana', '', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV069', 'OKR035', 'R23', 'OE07', 'Definicao do modelo de composicao das notas trimestrais', 'Junho', 'Junho', 'Qualitativo', 0.25, 'Ckelen', '', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV070', 'OKR036', 'R23', 'OE07', 'Aplicacao piloto', 'Agosto', 'Agosto', 'Qualitativo', 0.33, 'Nicolas', '', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV071', 'OKR036', 'R23', 'OE07', 'Analise dos resultados e fluxos operacionais', 'Agosto', 'Agosto', 'Qualitativo', 0.33, 'Nicolas', '', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV072', 'OKR036', 'R23', 'OE07', 'Ajustes tecnicos e pedagogicos', 'Agosto', 'Setembro', 'Qualitativo', 0.33, 'Debora', '', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV073', 'OKR037', 'R23', 'OE07', 'Elaboracao do documento orientador', 'Setembro', '', 'Qualitativo', 0.25, 'Debora', 'Em producao', 'Em andamento', 'UDP', 'UDP', 'PROFESSORES', 60, ''],
    ['ATV074', 'OKR037', 'R23', 'OE07', 'Sistematizacao de fluxos, protocolos e criterios', 'Setembro', '', 'Qualitativo', 0.25, 'Nicolas', 'Parcialmente concluido', 'Em andamento', 'UDP', 'UDP', 'PROFESSORES', 50, ''],
    ['ATV075', 'OKR037', 'R23', 'OE07', 'Formalizacao normativa', 'Outubro', '', 'Qualitativo', 0.25, 'Ckelen', '', 'No prazo', 'UDP', 'UDP', 'PROFESSORES', 10, ''],
    ['ATV076', 'OKR037', 'R23', 'OE07', 'Acoes formativas e comunicacionais', 'Novembro', '', 'Qualitativo', 0.25, 'Ckelen', '', 'No prazo', 'UDP', 'UDP', 'PROFESSORES', 0, ''],

    // === OE09: Gestao escolar (DGE) ===
    ['ATV077', 'OKR040', 'R26', 'OE09', 'Estruturar modulos do programa Escola de Lideres', 'Marco', 'Marco', 'Qualitativo', 0.5, 'Andrei Popovski', '', 'Concluido', 'DGE', 'DGE', 'GESTAO ESCOLAR', 100, ''],
    ['ATV078', 'OKR040', 'R26', 'OE09', 'Validar conteudo programatico com especialistas', 'Marco', 'Abril', 'Qualitativo', 0.5, 'Andrei Popovski', '', 'Concluido', 'DGE', 'DGE', 'GESTAO ESCOLAR', 100, ''],
    ['ATV079', 'OKR041', 'R26', 'OE09', 'Selecionar e convocar 50 gestores para 1a turma', 'Abril', 'Abril', 'Qualitativo', 0.33, 'Andrei Popovski', '', 'Concluido', 'DGE', 'DGE', 'GESTAO ESCOLAR', 100, ''],
    ['ATV080', 'OKR041', 'R26', 'OE09', 'Executar modulos formativos da 1a turma', 'Junho', 'Junho', 'Qualitativo', 0.33, 'Andrei Popovski', '', 'Concluido', 'DGE', 'DGE', 'GESTAO ESCOLAR', 100, ''],
    ['ATV081', 'OKR041', 'R26', 'OE09', 'Aplicar avaliacao de desempenho da 1a turma', 'Junho', 'Junho', 'Qualitativo', 0.33, 'Andrei Popovski', '', 'Concluido', 'DGE', 'DGE', 'GESTAO ESCOLAR', 100, ''],
    ['ATV082', 'OKR042', 'R26', 'OE09', 'Iniciar formacao da 2a turma', 'Agosto', 'Agosto', 'Qualitativo', 0.33, 'Andrei Popovski', '', 'Concluido', 'DGE', 'DGE', 'GESTAO ESCOLAR', 100, ''],
    ['ATV083', 'OKR042', 'R26', 'OE09', 'Executar modulos formativos da 2a turma', 'Outubro', '', 'Qualitativo', 0.33, 'Andrei Popovski', 'Em andamento - 3 de 5 modulos concluidos', 'Em andamento', 'DGE', 'DGE', 'GESTAO ESCOLAR', 60, ''],
    ['ATV084', 'OKR042', 'R26', 'OE09', 'Avaliacao e certificacao da 2a turma', 'Novembro', '', 'Qualitativo', 0.33, 'Andrei Popovski', '', 'No prazo', 'DGE', 'DGE', 'GESTAO ESCOLAR', 0, ''],
    ['ATV085', 'OKR043', 'R27', 'OE09', 'Monitorar execucao orcamentaria das UExs', 'Setembro', '', 'Quantitativo', 0.5, 'Andrei Popovski', '65% de execucao ate setembro', 'Em andamento', 'DGE', 'DGE', 'GESTAO ESCOLAR', 65, ''],
    ['ATV086', 'OKR043', 'R27', 'OE09', 'Emitir relatorios de execucao por UEx', 'Outubro', '', 'Quantitativo', 0.5, 'Andrei Popovski', '', 'Em andamento', 'DGE', 'DGE', 'GESTAO ESCOLAR', 40, ''],

    // === OE10: Gestao eficiente (UPL) ===
    ['ATV087', 'OKR044', 'R28', 'OE10', 'Levantar requisitos com todas as gerencias', 'Janeiro', 'Janeiro', 'Qualitativo', 0.5, 'Jose Victor', '', 'Concluido', 'UPL', 'UPL', 'CAPACIDADE INSTITUCIONAL', 100, ''],
    ['ATV088', 'OKR044', 'R28', 'OE10', 'Documento de requisitos validado', 'Fevereiro', 'Marco', 'Qualitativo', 0.5, 'Jose Victor', '', 'Concluido', 'UPL', 'UPL', 'CAPACIDADE INSTITUCIONAL', 100, ''],
    ['ATV089', 'OKR045', 'R28', 'OE10', 'Desenvolver MVP do sistema de monitoramento', 'Maio', 'Junho', 'Qualitativo', 0.5, 'Jose Victor', '', 'Concluido', 'UPL', 'UPL', 'CAPACIDADE INSTITUCIONAL', 100, ''],
    ['ATV090', 'OKR045', 'R28', 'OE10', 'Testar com equipe de planejamento', 'Junho', 'Junho', 'Qualitativo', 0.5, 'Jose Victor', '', 'Concluido', 'UPL', 'UPL', 'CAPACIDADE INSTITUCIONAL', 100, ''],
    ['ATV091', 'OKR046', 'R28', 'OE10', 'Migrar dados reais das planilhas para o sistema', 'Agosto', 'Setembro', 'Qualitativo', 0.33, 'Jose Victor', 'Processo complexo de migracao', 'Concluido', 'UPL', 'UPL', 'CAPACIDADE INSTITUCIONAL', 100, ''],
    ['ATV092', 'OKR046', 'R28', 'OE10', 'Treinar gerencias no uso do sistema', 'Setembro', '', 'Qualitativo', 0.33, 'Jose Victor', 'Treinamentos em andamento', 'Em andamento', 'UPL', 'UPL', 'CAPACIDADE INSTITUCIONAL', 60, ''],
    ['ATV093', 'OKR046', 'R28', 'OE10', 'Go-live com todas as diretorias', 'Outubro', '', 'Qualitativo', 0.33, 'Jose Victor', '', 'Em andamento', 'UPL', 'UPL', 'CAPACIDADE INSTITUCIONAL', 40, ''],

    // === OE11: Infraestrutura (DIN) ===
    ['ATV094', 'OKR049', 'R31', 'OE11', 'Licitacao para pisos modulares - lote 1', 'Marco', 'Abril', 'Qualitativo', 0.33, 'Felipe Hardt', '', 'Concluido', 'DIN', 'DIN', 'CAPACIDADE INSTITUCIONAL', 100, ''],
    ['ATV095', 'OKR049', 'R31', 'OE11', 'Instalacao pisos em 3 quadras', 'Maio', 'Junho', 'Qualitativo', 0.67, 'Felipe Hardt', '', 'Concluido', 'DIN', 'DIN', 'CAPACIDADE INSTITUCIONAL', 100, ''],
    ['ATV096', 'OKR050', 'R31', 'OE11', 'Licitacao para pisos modulares - lote 2', 'Setembro', 'Outubro', 'Qualitativo', 0.33, 'Felipe Hardt', '', 'Concluido', 'DIN', 'DIN', 'CAPACIDADE INSTITUCIONAL', 100, ''],
    ['ATV097', 'OKR050', 'R31', 'OE11', 'Instalacao pisos em 3 quadras restantes', 'Novembro', '', 'Qualitativo', 0.67, 'Felipe Hardt', 'Instalacao em andamento - 1 de 3 concluida', 'Em andamento', 'DIN', 'DIN', 'CAPACIDADE INSTITUCIONAL', 33, ''],
    ['ATV098', 'OKR051', 'R32', 'OE11', 'Contratar servico adequacao bombeiros', 'Julho', 'Agosto', 'Qualitativo', 0.5, 'Felipe Hardt', '', 'Concluido', 'DIN', 'DIN', 'CAPACIDADE INSTITUCIONAL', 100, ''],
    ['ATV099', 'OKR051', 'R32', 'OE11', 'Executar adequacoes em 10 UEs', 'Setembro', '', 'Qualitativo', 0.5, 'Felipe Hardt', '6 de 10 UEs concluidas', 'Em andamento', 'DIN', 'DIN', 'CAPACIDADE INSTITUCIONAL', 60, ''],

    // === OE12: Servidores (DGE) ===
    ['ATV100', 'OKR052', 'R33', 'OE12', 'Implementar dashboard de monitoramento de ausencias', 'Abril', 'Maio', 'Qualitativo', 0.5, 'Andrei Popovski', '', 'Concluido', 'DGE', 'DGE', 'CAPACIDADE INSTITUCIONAL', 100, ''],
    ['ATV101', 'OKR052', 'R33', 'OE12', 'Treinar gestores no uso do sistema de ausencias', 'Maio', 'Junho', 'Qualitativo', 0.5, 'Andrei Popovski', '', 'Concluido', 'DGE', 'DGE', 'CAPACIDADE INSTITUCIONAL', 100, ''],
    ['ATV102', 'OKR053', 'R33', 'OE12', 'Monitorar mensalmente a taxa de ausencia', 'Outubro', '', 'Quantitativo', 0.5, 'Andrei Popovski', 'Taxa atual: 1.2%', 'Em andamento', 'DGE', 'DGE', 'CAPACIDADE INSTITUCIONAL', 70, ''],
    ['ATV103', 'OKR053', 'R33', 'OE12', 'Implementar acoes corretivas nas UEs com maior indice', 'Novembro', '', 'Qualitativo', 0.5, 'Andrei Popovski', '', 'Em andamento', 'DGE', 'DGE', 'CAPACIDADE INSTITUCIONAL', 40, ''],
    ['ATV104', 'OKR054', 'R34', 'OE12', 'Elaborar e publicar edital do concurso', 'Agosto', 'Agosto', 'Qualitativo', 0.5, 'Andrei Popovski', '', 'Concluido', 'DGE', 'DGE', 'CAPACIDADE INSTITUCIONAL', 100, ''],
    ['ATV105', 'OKR054', 'R34', 'OE12', 'Periodo de inscricoes', 'Setembro', 'Setembro', 'Qualitativo', 0.5, 'Andrei Popovski', 'Mais de 5.000 inscritos', 'Concluido', 'DGE', 'DGE', 'CAPACIDADE INSTITUCIONAL', 100, ''],
    ['ATV106', 'OKR055', 'R34', 'OE12', 'Aplicar provas do concurso', 'Novembro', '', 'Qualitativo', 0.5, 'Andrei Popovski', 'Data marcada para 20/nov', 'No prazo', 'DGE', 'DGE', 'CAPACIDADE INSTITUCIONAL', 20, ''],
    ['ATV107', 'OKR055', 'R34', 'OE12', 'Divulgar resultado preliminar', 'Dezembro', '', 'Qualitativo', 0.5, 'Andrei Popovski', '', 'No prazo', 'DGE', 'DGE', 'CAPACIDADE INSTITUCIONAL', 0, ''],

    // === OE06/08: Formacao e Bem-estar (UDP/DPE) ===
    ['ATV108', 'OKR030', 'R21', 'OE06', 'Definir estrutura curricular do programa de formacao', 'Fevereiro', 'Fevereiro', 'Qualitativo', 0.5, 'Ckelen', '', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV109', 'OKR030', 'R21', 'OE06', 'Selecionar formadores e mentores', 'Marco', 'Marco', 'Qualitativo', 0.5, 'Ckelen', '', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV110', 'OKR031', 'R21', 'OE06', 'Iniciar 1a turma do ciclo formativo', 'Abril', 'Abril', 'Qualitativo', 0.5, 'Ckelen', '', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV111', 'OKR031', 'R21', 'OE06', 'Aplicar avaliacao NPS da 1a turma', 'Junho', 'Junho', 'Qualitativo', 0.5, 'Ckelen', 'NPS: 78', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV112', 'OKR032', 'R21', 'OE06', 'Iniciar 2a e 3a turmas do ciclo formativo', 'Julho', 'Julho', 'Qualitativo', 0.33, 'Ckelen', '', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV113', 'OKR032', 'R21', 'OE06', 'Monitorar avanco do ciclo formativo (meta 60%)', 'Setembro', '', 'Quantitativo', 0.33, 'Ckelen', '72% atingido', 'Concluido', 'UDP', 'UDP', 'PROFESSORES', 100, ''],
    ['ATV114', 'OKR033', 'R21', 'OE06', 'Atingir 85% do ciclo formativo completo', 'Dezembro', '', 'Quantitativo', 0.5, 'Ckelen', 'Atual: 72% - em progresso', 'Em andamento', 'UDP', 'UDP', 'PROFESSORES', 55, ''],
    ['ATV115', 'OKR038', 'R25', 'OE08', 'Desenvolver instrumento de diagnostico de clima', 'Abril', 'Maio', 'Qualitativo', 0.5, 'DPE', '', 'Concluido', 'DPE', 'DPE', 'GESTAO ESCOLAR', 100, ''],
    ['ATV116', 'OKR038', 'R25', 'OE08', 'Aplicar diagnostico piloto em 5 unidades', 'Junho', '', 'Qualitativo', 0.5, 'DPE', 'Atraso na validacao do instrumento', 'Atrasado', 'DPE', 'DPE', 'GESTAO ESCOLAR', 30, ''],
    ['ATV117', 'OKR039', 'R25', 'OE08', 'Selecionar 5 UEs para piloto Escola Saudavel', 'Julho', 'Julho', 'Qualitativo', 0.33, 'DPE', '', 'Concluido', 'DPE', 'DPE', 'GESTAO ESCOLAR', 100, ''],
    ['ATV118', 'OKR039', 'R25', 'OE08', 'Iniciar acoes do projeto nas UEs selecionadas', 'Agosto', 'Setembro', 'Qualitativo', 0.33, 'DPE', '', 'Em andamento', 'DPE', 'DPE', 'GESTAO ESCOLAR', 60, ''],
    ['ATV119', 'OKR039', 'R25', 'OE08', 'Coletar dados de IMC inicial dos alunos', 'Setembro', '', 'Quantitativo', 0.33, 'DPE', 'Coleta em andamento', 'Em andamento', 'DPE', 'DPE', 'GESTAO ESCOLAR', 40, '']
  ];

  if (atividades.length > 0) {
    sheet.getRange(2, 1, atividades.length, atividades[0].length).setValues(atividades);
  }
}
