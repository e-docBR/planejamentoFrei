// ============================================================
//  SISTEMA DE PLANEJAMENTO v9 — COLÉGIO MUNICIPAL DE ITABATAN
//  Google Apps Script — Código Principal
//  CORREÇÕES DE BUGS + TRIGGER CONFIGURÁVEL + ANO LETIVO DINÂMICO
// ============================================================
//
//  v9 — CORREÇÕES:
//  [V9-1] Bug salvarConfiguracoes: índice de linha corrigido (re-leitura
//         da planilha após cada escrita evita deslocamento de índice)
//  [V9-2] Bug instalarTriggerSemanal: lê configurações antes de instalar
//         (garantia de que lembrete_dia_semana e lembrete_hora existem)
//  [V9-3] Ano letivo dinâmico nos PDFs: usa _cfg('ano_letivo') com fallback
//         para CONFIG.ANO_LETIVO em vez de string fixa
//  [V9-4] Cache de _cfg() invalidado junto com cache do painel ao salvar
//         configurações — evita PDFs com ano desatualizado
//  [V9-5] lembreteSemanaPlanejamento: envio de resumo para TODOS os e-mails
//         de coordenação (não apenas o primeiro)
//  [V9-6] exportarParaDocs: pasta individual do professor para Docs editáveis
//  [V9-7] _paginaVerificacao: tratamento seguro quando aba Verificacao tem
//         colunas extras (desestruturação defensiva)
//
//  v9 (patch 2026-04) — CORREÇÕES ADICIONAIS:
//  [V9-8]  _validarConfig(): bloqueia operações se CONFIG ainda tem placeholders;
//          doGet() exibe mensagem clara de configuração pendente
//  [V9-9]  enviarEmailConfirmacao: envia notificação para TODOS os e-mails de
//          coordenação (BUG-4: antes só enviava para o primeiro)
//  [V9-10] processarFilaOffline: valida estrutura de cada item da fila antes de
//          processar (BUG-6: itens malformados antes quebravam silenciosamente)
//  [V9-11] _buscarBNCC_API: cache key usa MD5 da query completa em vez de
//          substring(0,40) — evita colisões em queries longas (PERF-6)
//  [V9-12] _buscarBNCC_API: filtra itens null/não-objeto antes do .map()
//          (BUG-9: crash com resposta inesperada da API)
//  [V9-13] salvarCadastros: aplica .trim() em nomes de professores/turmas/
//          componentes para evitar duplicatas por espaços acidentais (QUALITY-3)
//  [V9-14] _provisionarPastaProfessor: coordenação recebe Viewer (não Editor)
//          nas pastas individuais dos professores (SEC-6)
//  [V9-15] SHEET_COLS: constantes de índice de coluna por tipo de planilha —
//          elimina magic numbers e facilita manutenção (QUALITY-2)
//  [V9-16] healthCheck(): testa Drive, Sheets, e-mail, API BNCC e CacheService;
//          getStatusSistema() expõe para o painel da coordenação
//  [V9-17] getStatusSistema(): wrapper autenticado para healthCheck()
//  [V9-18] getHistoricoProfessor(): suporte a paginação (page, pageSize)
//          e uso de SHEET_COLS (PERF-5 / QUALITY-2)
//  [V9-19] salvarAprovacao(): verificação explícita de perfil coordenação
//          antes de registrar aprovação/revisão (SEC-4)
//  [V9-20] ✨ CRIAÇÃO AUTOMÁTICA DA PLANILHA: CONFIG.PLANILHA_ID agora é
//          OPCIONAL! No primeiro acesso, a planilha é criada automaticamente
//          e o ID é salvo nas propriedades do script. Facilita instalação!
//          Use infosPlanilha() para ver status e resetarIdPlanilha() se precisar.
// ============================================================

// ─── CONFIGURAÇÕES — ÚNICO LUGAR PARA EDITAR ───────────────
const CONFIG = {
  PASTA_ID:           "1oElOAG9HwoPTyEatRNskrgBHLPBXC3j1",
  
  // [V9-20] PLANILHA_ID agora é OPCIONAL!
  // Se deixar vazio ou com placeholder, a planilha será criada automaticamente
  // no primeiro acesso e o ID será salvo nas propriedades do script.
  // Para melhor performance, você pode preencher manualmente após a criação:
  PLANILHA_ID:        "COLE_AQUI_O_ID_DA_PLANILHA_REGISTRO",  // OPCIONAL
  
  // [V7-1] E-mails com acesso de coordenação (pode ter mais de um)
  EMAILS_COORDENACAO: [
    "josival.lima@edu.mucuri.ba.gov.br",
    // "josival.lima@edu.mucuri.ba.gov.br",  // adicione quantos precisar
  ],
  FUSO:               "America/Bahia",
  ANO_LETIVO:         new Date().getFullYear(),
  CACHE_PAINEL_SEG:   300,
};

// [V9-8] Validação de configuração — bloqueia operações se placeholders não foram substituídos
function _validarConfig() {
  const erros = [];
  
  // [V9-20] PLANILHA_ID não é mais obrigatório - será criada automaticamente no primeiro acesso
  // if (!CONFIG.PLANILHA_ID || CONFIG.PLANILHA_ID.startsWith('COLE_')) {
  //   erros.push('PLANILHA_ID não configurado. Edite CONFIG.PLANILHA_ID com o ID real da planilha.');
  // }
  
  if (!CONFIG.PASTA_ID || CONFIG.PASTA_ID.startsWith('COLE_')) {
    erros.push('PASTA_ID não configurado. Edite CONFIG.PASTA_ID com o ID real da pasta no Drive.');
  }
  const emailsInvalidos = CONFIG.EMAILS_COORDENACAO.filter(e => !e || e.startsWith('COLE_'));
  if (emailsInvalidos.length > 0) {
    erros.push('EMAILS_COORDENACAO contém placeholders. Substitua pelo(s) e-mail(s) real(is) da coordenação.');
  }
  if (erros.length > 0) {
    throw new Error('⚠️ Configuração incompleta:\n• ' + erros.join('\n• '));
  }
}

// Atalho para verificar coordenação com múltiplos e-mails
function _ehCoordenacao(email) {
  if (!email) return false;
  return CONFIG.EMAILS_COORDENACAO.some(e => e.toLowerCase() === email.toLowerCase());
}

// Atalho retrocompatível para código que ainda usa CONFIG.EMAIL_COORDENACAO
Object.defineProperty(CONFIG, 'EMAIL_COORDENACAO', {
  get() { return CONFIG.EMAILS_COORDENACAO[0] || ''; }
});

// Nomes das subpastas de arquivo
const SP = {
  TRIMESTRAL: "Planejamento Trimestral",
  SEMANAL:    "Planejamento Semanal",
  ANUAL:      "Planejamento Anual",
  DOCS:       "Docs Editáveis",
};

// Nomes das abas da planilha
const ABA = {
  TRIMESTRAL:       "Trimestral",
  SEMANAL:          "Semanal",
  ANUAL:            "Anual",
  CADASTROS:        "Cadastros",
  LOGS:             "Logs",
  CALENDARIO:       "Calendario",
  BNCC:             "BNCC",
  APROVACOES:       "Aprovacoes",
  CONFIGURACOES:    "Configuracoes",    // [V8-2] Página de configurações
  TURMAS_PROFESSOR: "TurmasProfessor",  // [V8-4] Turmas atribuídas por professor
};

// [V9-15] Constantes de índice de coluna por tipo de planilha — evita magic numbers
// Ao adicionar colunas na planilha, atualize aqui e o código todo reflete automaticamente
const SHEET_COLS = {
  TRIMESTRAL: { DATA: 0, PROFESSOR: 1, ANO_TURMA: 2, COMPONENTE: 3, TRIMESTRE: 4, URL: 5 },
  SEMANAL:    { DATA: 0, PROFESSOR: 1, TURMAS: 2, COMPONENTE: 3, MES: 4, SEMANA_INICIO: 5, URL: 7 },
  ANUAL:      { DATA: 0, PROFESSOR: 1, ANO_TURMA: 2, COMPONENTE: 3, URL: 4 },
  VERIFICACAO:{ ID: 0, TIPO: 1, PROFESSOR: 2, COMPONENTE: 3, PERIODO: 4, DATA_HORA: 5, URL_PDF: 6, VALIDO: 7 },
  CADASTROS:  { TIPO: 0, NOME: 1, EMAIL: 2 },
  APROVACOES: { DATA_HORA: 0, TIPO: 1, PROFESSOR: 2, COMPONENTE: 3, PERIODO: 4, DECISAO: 5, COMENTARIO: 6, URL_PDF: 7, REVISOR: 8 },
};

// ─── PONTO DE ENTRADA ──────────────────────────────────────
function doGet(e) {
  try {
    // [F4-3] Rota de verificação pública — sem autenticação
    if (e && e.parameter && e.parameter.verificar) {
      return _paginaVerificacao(e.parameter.verificar);
    }

    // [V9-LANDING] Landing page pública — servida quando ?app=1 não está presente
    // Usa createTemplateFromFile para injetar a URL real do exec server-side,
    // pois window.location dentro do sandbox GAS retorna googleusercontent.com, não o /exec
    if (!e || !e.parameter || !e.parameter.app) {
      const tpl = HtmlService.createTemplateFromFile('LandingPage_Sites');
      tpl.appUrl = ScriptApp.getService().getUrl() + '?app=1';
      return tpl.evaluate()
        .setTitle('Planejamento Escolar — Colégio Itabatan')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }

    // [V9-8] Valida configuração antes de entrar no app
    try { _validarConfig(); } catch(configErr) {
      return HtmlService.createHtmlOutput(
        '<pre style="font-family:monospace;padding:24px;color:#c0392b">' + configErr.message + '</pre>'
      ).setTitle('Configuração necessária');
    }

    // [V7-5] PORTÃO DE AUTENTICAÇÃO
    // Verifica se o usuário está cadastrado antes de servir qualquer página
    const perfil = getPerfil();
    if (perfil.perfil === 'desconhecido') {
      return _paginaAcessoNegado(perfil.email);
    }

    // [F4-4] Roteamento por módulo
    // NOTA: usa createHtmlOutputFromFile (não createTemplateFromFile + evaluate)
    // porque o HTML usa template literals JS (`${var}`) que conflitam com o
    // compilador de templates do GAS — os backticks fecham o template literal
    // interno do GAS, causando SyntaxError "Unexpected identifier '$'" no browser.
    const modulo = e.parameter.modulo;
    const output = modulo
      ? HtmlService.createHtmlOutputFromFile(modulo)
      : HtmlService.createHtmlOutputFromFile('Index');

    return output
      .setTitle('Planejamento Escolar – Colégio Itabatan')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      
  } catch(erroGlobal) {
    // [V9-20] Captura qualquer erro não tratado e exibe mensagem detalhada
    log('ERRO', 'doGet', `Erro não tratado: ${erroGlobal.message}\n${erroGlobal.stack}`);
    
    return HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html lang="pt-BR">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width,initial-scale=1">
        <title>Erro no Sistema</title>
        <style>
          body { font-family: 'Segoe UI', Arial, sans-serif; background: #f8f9fa; padding: 20px; }
          .erro-container { max-width: 800px; margin: 40px auto; background: white; 
                           border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
          .erro-header { background: #dc3545; color: white; padding: 20px; 
                        border-radius: 8px 8px 0 0; }
          .erro-body { padding: 30px; }
          .erro-mensagem { background: #f8d7da; border: 1px solid #f5c6cb; 
                          border-radius: 4px; padding: 15px; margin: 20px 0; 
                          color: #721c24; font-family: monospace; white-space: pre-wrap; }
          .erro-ajuda { background: #d1ecf1; border: 1px solid #bee5eb; 
                       border-radius: 4px; padding: 15px; margin: 20px 0; 
                       color: #0c5460; }
          .btn-diagnostico { background: #007bff; color: white; padding: 10px 20px; 
                            border: none; border-radius: 4px; cursor: pointer; 
                            font-size: 14px; margin-top: 10px; }
          .btn-diagnostico:hover { background: #0056b3; }
        </style>
      </head>
      <body>
        <div class="erro-container">
          <div class="erro-header">
            <h1>⚠️ Erro no Sistema de Planejamento</h1>
          </div>
          <div class="erro-body">
            <p><strong>Ocorreu um erro inesperado ao carregar o sistema:</strong></p>
            
            <div class="erro-mensagem">${erroGlobal.message}

Stack trace:
${erroGlobal.stack}</div>
            
            <div class="erro-ajuda">
              <h3>📋 Como resolver:</h3>
              <ol>
                <li>Abra o <strong>Apps Script Editor</strong> deste projeto</li>
                <li>Na lista de funções, selecione <strong>diagnosticoSistema</strong></li>
                <li>Clique em <strong>Executar</strong> (ícone de play ▶)</li>
                <li>Veja o relatório completo no painel de <strong>Execuções</strong></li>
                <li>Siga as recomendações apresentadas no diagnóstico</li>
              </ol>
              
              <p><strong>Possíveis causas comuns:</strong></p>
              <ul>
                <li>📁 ID da pasta do Drive incorreto ou sem permissão</li>
                <li>📊 Erro ao criar/acessar a planilha de registros</li>
                <li>🔐 Faltam permissões de acesso ao Drive/Sheets</li>
                <li>⚙️ Configuração incompleta no objeto CONFIG</li>
              </ul>
            </div>
            
            <p style="margin-top: 20px; padding-top: 20px; border-top: 1px solid #dee2e6; color: #6c757d; font-size: 12px;">
              Sistema de Planejamento Escolar v9.20 - Colégio Municipal de Itabatan
            </p>
          </div>
        </div>
      </body>
      </html>
    `).setTitle('Erro no Sistema');
  }
}

// [V7-5] Página de acesso negado
function _paginaAcessoNegado(email) {
  const html = `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Acesso não autorizado</title>
  <style>
    body{font-family:Arial,sans-serif;background:#f4f6fb;display:flex;align-items:center;
      justify-content:center;min-height:100vh;margin:0}
    .box{background:#fff;border-radius:14px;padding:40px 44px;max-width:460px;width:100%;
      text-align:center;box-shadow:0 8px 40px rgba(0,0,0,.12)}
    .icon{font-size:3.5rem;margin-bottom:14px}
    h1{font-size:1.25rem;color:#1a3a6b;margin-bottom:8px}
    p{color:#555;font-size:.95rem;line-height:1.6;margin-bottom:6px}
    .email{font-family:monospace;font-size:.85rem;background:#f4f6fb;
      padding:4px 12px;border-radius:6px;display:inline-block;margin:8px 0}
    .instrucao{margin-top:18px;padding:14px;background:#ebf0fa;border-radius:8px;
      font-size:.85rem;color:#1a3a6b;text-align:left}
  </style></head><body>
  <div class="box">
    <div class="icon">🔒</div>
    <h1>Acesso não autorizado</h1>
    <p>Você tentou acessar o sistema com o e-mail:</p>
    <span class="email">${esc(email || 'não identificado')}</span>
    <p>Este e-mail não está cadastrado no sistema de planejamento escolar.</p>
    <div class="instrucao">
      <strong>O que fazer?</strong><br>
      Entre em contato com a coordenação pedagógica para solicitar o cadastro
      do seu e-mail no sistema. Após o cadastro, tente acessar novamente.
    </div>
    <p style="margin-top:18px;font-size:.8rem;color:#888">
      Colégio Municipal de 1º e 2º Graus de Itabatan — Mucuri/BA
    </p>
  </div>
  </body></html>`;
  return HtmlService.createHtmlOutput(html).setTitle('Acesso não autorizado');
}

// [F4-4] Inclui arquivos parciais nos templates HTML
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ═══════════════════════════════════════════════════════════
//  [BUG-1] CADASTROS LIDOS DA PLANILHA (não mais hardcoded)
// ═══════════════════════════════════════════════════════════
function getCadastros() {
  try {
    const ss  = abrirPlanilha();
    let aba   = ss.getSheetByName(ABA.CADASTROS);

    // Cria a aba com dados iniciais se não existir
    if (!aba) {
      aba = _criarAbaCadastros(ss);
    }

    const vals = aba.getDataRange().getValues();
    const professores = [], componentes = [], turmas = [], emails = {};

    // Formato da aba: coluna A = tipo (PROFESSOR/COMPONENTE/TURMA)
    //                 coluna B = valor
    //                 coluna C = e-mail (apenas para PROFESSOR)
    vals.slice(1).forEach(row => {
      // Validação robusta de dados
      if (!row || row.length < 2) return;
      const tipo  = String(row[0] || '').trim().toUpperCase();
      const valor = String(row[1] || '').trim();
      const email = String(row[2] || '').trim();
      if (!valor || !tipo) return;
      
      // Normaliza espaços duplicados e caracteres especiais invisíveis
      const valorNormalizado = valor.replace(/\s+/g, ' ').trim();
      
      if (tipo === 'PROFESSOR') { 
        professores.push(valorNormalizado); 
        if (email && email.includes('@')) emails[valorNormalizado] = email.toLowerCase(); 
      }
      if (tipo === 'COMPONENTE') componentes.push(valorNormalizado);
      if (tipo === 'TURMA')      turmas.push(valorNormalizado);
    });

    // Remove duplicatas mantendo ordem
    const profUnicos = [...new Set(professores)];
    const compUnicos = [...new Set(componentes)];
    const turmUnicos = [...new Set(turmas)];

    return { 
      professores: profUnicos, 
      componentes: compUnicos, 
      turmas: turmUnicos, 
      emails 
    };
  } catch(e) {
    log('ERRO', 'getCadastros', e.message);
    // Fallback mínimo para o sistema não parar
    return {
      professores: ['— Configure a aba Cadastros na planilha —'],
      componentes: ['— Configure a aba Cadastros na planilha —'],
      turmas:      ['— Configure a aba Cadastros na planilha —'],
      emails:      {}
    };
  }
}

function _criarAbaCadastros(ss) {
  const aba = ss.insertSheet(ABA.CADASTROS);
  const header = ["Tipo", "Nome/Valor", "E-mail (só Professor)"];
  aba.appendRow(header);
  aba.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#1a3a6b").setFontColor("white");

  // Dados iniciais de exemplo — coordenação edita diretamente na planilha
  const exemplos = [
    ["PROFESSOR",  "Ana Lima",             "ana.lima@escola.edu.br"],
    ["PROFESSOR",  "Bruno Santos",         "bruno.santos@escola.edu.br"],
    ["PROFESSOR",  "Carla Oliveira",       "carla.oliveira@escola.edu.br"],
    ["COMPONENTE", "Língua Portuguesa",    ""],
    ["COMPONENTE", "Matemática",           ""],
    ["COMPONENTE", "Ciências",             ""],
    ["COMPONENTE", "História",             ""],
    ["COMPONENTE", "Geografia",            ""],
    ["COMPONENTE", "Arte",                 ""],
    ["COMPONENTE", "Educação Física",      ""],
    ["COMPONENTE", "Inglês",               ""],
    ["COMPONENTE", "Ensino Religioso",     ""],
    ["COMPONENTE", "EJA — Língua Portuguesa", ""],
    ["COMPONENTE", "EJA — Matemática",    ""],
    ["TURMA", "6ºA",""], ["TURMA", "6ºB",""], ["TURMA", "6ºC",""], ["TURMA", "6ºD",""],
    ["TURMA", "6ºE",""], ["TURMA", "6ºF",""], ["TURMA", "6ºG",""], ["TURMA", "6ºH",""],
    ["TURMA", "7ºA",""], ["TURMA", "7ºB",""], ["TURMA", "7ºC",""], ["TURMA", "7ºD",""], ["TURMA", "7ºE",""],
    ["TURMA", "8ºA",""], ["TURMA", "8ºB",""], ["TURMA", "8ºC",""], ["TURMA", "8ºD",""],
    ["TURMA", "9ºA",""], ["TURMA", "9ºB",""], ["TURMA", "9ºC",""],
  ];
  aba.getRange(2, 1, exemplos.length, 3).setValues(exemplos);
  aba.autoResizeColumns(1, 3);

  // Nota de instrução
  aba.getRange(exemplos.length + 3, 1).setValue(
    "INSTRUÇÃO: Adicione linhas com Tipo = PROFESSOR, COMPONENTE ou TURMA. " +
    "Não remova o cabeçalho. Alterações refletem imediatamente no sistema."
  ).setFontStyle("italic").setFontColor("#555555");

  return aba;
}

// ═══════════════════════════════════════════════════════════
//  SALVAR PDFs  [V7-4] pasta individual do professor
// ═══════════════════════════════════════════════════════════
function salvarTrimestral(dados) {
  try {
    _verificarPermissao('salvarTrimestral');                        // [V7-6]
    const pasta  = _pastaDestino('T', dados.professor);             // [V7-4]
    const nome   = nomearArquivo("T", dados);
    const tempPdf = toPDF(htmlTrimestral(dados, { qrSrc: '', urlVerif: '' }), nome + '_tmp');
    const tempArq = pasta.createFile(tempPdf);
    const urlPDF  = tempArq.getUrl();

    const verif  = _gerarRegistroVerificacao(ABA.TRIMESTRAL, dados, urlPDF);
    const pdfFinal = toPDF(htmlTrimestral(dados, verif), nome);
    tempArq.setTrashed(true);
    const arq = pasta.createFile(pdfFinal);

    _atualizarUrlVerificacao(verif.id, arq.getUrl());
    registrar(ABA.TRIMESTRAL, dados, arq.getUrl());
    log('INFO', 'salvarTrimestral', `PDF gerado na pasta do professor: ${nome}`);
    return { sucesso: true, url: arq.getUrl(), nome: nome + ".pdf", urlVerif: verif.urlVerif };
  } catch(e) {
    log('ERRO', 'salvarTrimestral', e.message, dados.professor);
    return { sucesso: false, erro: e.message };
  }
}

function salvarSemanal(dados) {
  try {
    _verificarPermissao('salvarSemanal');                           // [V7-6]
    const pasta  = _pastaDestino('S', dados.professor);             // [V7-4]
    const nome   = nomearArquivo("S", dados);
    const tempPdf = toPDF(htmlSemanal(dados, { qrSrc: '', urlVerif: '' }), nome + '_tmp');
    const tempArq = pasta.createFile(tempPdf);
    const urlPDF  = tempArq.getUrl();

    const verif   = _gerarRegistroVerificacao(ABA.SEMANAL, dados, urlPDF);
    const pdfFinal = toPDF(htmlSemanal(dados, verif), nome);
    tempArq.setTrashed(true);
    const arq = pasta.createFile(pdfFinal);
    _atualizarUrlVerificacao(verif.id, arq.getUrl());

    registrar(ABA.SEMANAL, dados, arq.getUrl());
    log('INFO', 'salvarSemanal', `PDF gerado na pasta do professor: ${nome}`);
    return { sucesso: true, url: arq.getUrl(), nome: nome + ".pdf", urlVerif: verif.urlVerif };
  } catch(e) {
    log('ERRO', 'salvarSemanal', e.message, dados.professor);
    return { sucesso: false, erro: e.message };
  }
}

function salvarAnual(dados) {
  try {
    _verificarPermissao('salvarAnual');                             // [V7-6]
    const pasta   = _pastaDestino('A', dados.professor);            // [V7-4]
    const nome    = nomearArquivo("A", dados);
    const tempPdf = toPDF(htmlAnual(dados, { qrSrc: '', urlVerif: '' }), nome + '_tmp');
    const tempArq = pasta.createFile(tempPdf);
    const urlPDF  = tempArq.getUrl();

    const verif   = _gerarRegistroVerificacao(ABA.ANUAL, dados, urlPDF);
    const pdfFinal = toPDF(htmlAnual(dados, verif), nome);
    tempArq.setTrashed(true);
    const arq = pasta.createFile(pdfFinal);
    _atualizarUrlVerificacao(verif.id, arq.getUrl());

    registrar(ABA.ANUAL, dados, arq.getUrl());
    log('INFO', 'salvarAnual', `PDF gerado na pasta do professor: ${nome}`);
    return { sucesso: true, url: arq.getUrl(), nome: nome + ".pdf", urlVerif: verif.urlVerif };
  } catch(e) {
    log('ERRO', 'salvarAnual', e.message, dados.professor);
    return { sucesso: false, erro: e.message };
  }
}

function _atualizarUrlVerificacao(id, urlFinal) {
  try {
    const ss  = abrirPlanilha();
    const aba = ss.getSheetByName("Verificacao");
    if (!aba) return;
    const vals = aba.getDataRange().getValues();
    vals.slice(1).forEach((row, i) => {
      if (String(row[0]).trim() === id) {
        aba.getRange(i + 2, 7).setValue(urlFinal); // coluna G = Link PDF
      }
    });
  } catch(e) { /* não crítico */ }
}

// ═══════════════════════════════════════════════════════════
//  EXPORTAR PARA GOOGLE DOCS
// ═══════════════════════════════════════════════════════════
function exportarParaDocs(payload) {
  try {
    const { tipo, dados } = payload;
    // [V9-6] Usa a pasta individual do professor (subpasta Docs) em vez da pasta global
    let pasta;
    try {
      pasta = _pastaDestino('Docs', dados.professor);
    } catch(e2) {
      // Fallback para pasta global se a individual não existir
      pasta = subpasta(CONFIG.PASTA_ID, SP.DOCS);
    }
    const nome  = nomearArquivo(tipo.charAt(0).toUpperCase(), dados) + "_EDITAVEL";
    const doc   = DocumentApp.create(nome);
    const body  = doc.getBody();

    const hdr = body.appendParagraph("COLÉGIO MUNICIPAL DE 1º E 2º GRAUS DE ITABATAN");
    hdr.setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph("Criação 062/94 · Av. Itapetinga, 305 – Itabatan – Mucuri/BA")
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    if (tipo === 'trimestral') _docTrimestral(body, dados);
    else if (tipo === 'semanal') _docSemanal(body, dados);
    else if (tipo === 'anual')   _docAnual(body, dados);

    doc.saveAndClose();
    const arqDoc = DriveApp.getFileById(doc.getId());
    pasta.addFile(arqDoc);
    DriveApp.getRootFolder().removeFile(arqDoc);
    log('INFO', 'exportarParaDocs', `Doc criado: ${nome}`);
    return { sucesso: true, url: doc.getUrl() };
  } catch(e) {
    log('ERRO', 'exportarParaDocs', e.message);
    return { sucesso: false, erro: e.message };
  }
}

function _docTrimestral(body, d) {
  body.appendParagraph(`Professor(a): ${d.professor}  |  Componente: ${d.componente}  |  Turma: ${d.anoTurma}`);
  body.appendParagraph(`PLANEJAMENTO ${d.trimestre} TRIMESTRE ${CONFIG.ANO_LETIVO}`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  const tab = body.appendTable();
  const hr  = tab.appendTableRow();
  ["UNIDADE TEMÁTICA", "HABILIDADES", "METODOLOGIA / RECURSOS", "AVALIAÇÃO"]
    .forEach(h => hr.appendTableCell(h).editAsText().setBold(true));
  (d.linhas || []).forEach(l => {
    const row = tab.appendTableRow();
    [l.unidade||'', l.habilidades||'', l.metodologia||'', l.avaliacao||'']
      .forEach(v => row.appendTableCell(v));
  });
  if (d.obs) body.appendParagraph(`\nObservações: ${d.obs}`);
}

function _docSemanal(body, d) {
  body.appendParagraph(`Professor(a): ${d.professor}  |  Componente: ${d.componente}  |  Turmas: ${d.turmas}`);
  body.appendParagraph(`PLANEJAMENTO SEMANAL — ${d.semanaInicio} a ${d.semanaFim}`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  const tab = body.appendTable();
  const hr  = tab.appendTableRow();
  hr.appendTableCell('').editAsText().setBold(true);
  ["Segunda","Terça","Quarta","Quinta","Sexta"].forEach(d2 => hr.appendTableCell(d2).editAsText().setBold(true));
  ['habilidade','estrategia','recursos','observacao'].forEach(campo => {
    const row = tab.appendTableRow();
    row.appendTableCell(campo.toUpperCase()).editAsText().setBold(true);
    [0,1,2,3,4].forEach(c => row.appendTableCell(d.semana&&d.semana[c] ? d.semana[c][campo]||'' : ''));
  });
  if (d.obs) body.appendParagraph(`\nObservações: ${d.obs}`);
}

function _docAnual(body, d) {
  body.appendParagraph(`Professor(a): ${d.professor}  |  Componente: ${d.componente}  |  Turma: ${d.anoTurma}`);
  body.appendParagraph(`PLANEJAMENTO ANUAL ${CONFIG.ANO_LETIVO}`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  if (d.objetivo) body.appendParagraph(`Objetivo Geral: ${d.objetivo}`);
  if (d.bncc)     body.appendParagraph(`Competências BNCC: ${d.bncc}`);
  const meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const tab = body.appendTable();
  const hr  = tab.appendTableRow();
  hr.appendTableCell('MÊS').editAsText().setBold(true);
  hr.appendTableCell('CONTEÚDOS / UNIDADES / HABILIDADES').editAsText().setBold(true);
  meses.forEach((mes, i) => {
    const row = tab.appendTableRow();
    row.appendTableCell(mes).editAsText().setBold(true);
    row.appendTableCell(d.meses&&d.meses[i] ? d.meses[i] : '');
  });
  if (d.obs) body.appendParagraph(`\nObservações: ${d.obs}`);
}

// ═══════════════════════════════════════════════════════════
//  NOTIFICAÇÃO POR E-MAIL
// ═══════════════════════════════════════════════════════════
function enviarEmailConfirmacao(payload) {
  try {
    // Validação de entrada
    if (!payload || typeof payload !== 'object') {
      log('AVISO', 'enviarEmailConfirmacao', 'Payload inválido');
      return;
    }
    
    const { professor, tipo, urlPDF } = payload;
    
    if (!professor || !tipo) {
      log('AVISO', 'enviarEmailConfirmacao', 'Dados insuficientes para envio de email');
      return;
    }
    
    const agora  = ts('dd/MM/yyyy HH:mm');
    const cadastros = getCadastros();
    const emailProf = cadastros.emails && cadastros.emails[professor];

    const assunto = `✅ Planejamento ${esc(tipo)} salvo — ${esc(professor)}`;
    const corpo = [
      `Olá, ${esc(professor)}!`,
      ``,
      `Seu Planejamento ${esc(tipo)} foi salvo com sucesso no sistema.`,
      ``,
      `📅 Data/Hora: ${agora}`,
      urlPDF ? `📄 Link para o PDF: ${urlPDF}` : '',
      ``,
      `Colégio Municipal de 1º e 2º Graus de Itabatan`,
      `Coordenação Pedagógica`,
    ].filter(Boolean).join('\n');

    if (emailProf && emailProf.includes('@')) {
      try {
        MailApp.sendEmail(emailProf, assunto, corpo);
        log('INFO', 'email', `Confirmação enviada para ${emailProf}`);
      } catch(eMailProf) {
        log('AVISO', 'enviarEmailConfirmacao', `Falha ao enviar para professor ${emailProf}: ${eMailProf.message}`);
      }
    }
    
    // [V9-9] Envia para TODOS os e-mails de coordenação (corrige BUG-4: antes enviava só para o primeiro)
    const corpoCoord = `Novo planejamento entregue:\nProfessor: ${esc(professor)}\nTipo: ${esc(tipo)}\nData: ${agora}${urlPDF ? '\nLink: ' + urlPDF : ''}`;
    CONFIG.EMAILS_COORDENACAO.filter(e => e && !e.startsWith('COLE_') && e.includes('@')).forEach(emailCoord => {
      try {
        MailApp.sendEmail(emailCoord, `[Coordenação] ${assunto}`, corpoCoord);
      } catch(eCoord) {
        log('AVISO', 'enviarEmailConfirmacao', `Falha ao enviar para coordenação ${emailCoord}: ${eCoord.message}`);
      }
    });
  } catch(e) {
    log('AVISO', 'enviarEmailConfirmacao', `Falha ao enviar e-mail: ${e.message}`);
  }
}

// ═══════════════════════════════════════════════════════════
//  HISTÓRICO DO PROFESSOR
// ═══════════════════════════════════════════════════════════
// [V9-18] Suporte a paginação — page (1-based), pageSize (default 30)
function getHistoricoProfessor(professor, page, pageSize) {
  try {
    // Validação de entrada
    if (!professor || typeof professor !== 'string' || !professor.trim()) {
      return { itens: [], total: 0, pagina: 1, tamPagina: 30, totalPaginas: 0, erro: 'Professor não especificado' };
    }
    
    const pagAtual  = Math.max(1, parseInt(page  || 1));
    const tamPagina = Math.min(100, Math.max(1, parseInt(pageSize || 30)));
    const C_T = SHEET_COLS.TRIMESTRAL;
    const C_S = SHEET_COLS.SEMANAL;
    const C_A = SHEET_COLS.ANUAL;

    const ss  = abrirPlanilha();
    const res = [];
    [ABA.TRIMESTRAL, ABA.SEMANAL, ABA.ANUAL].forEach(tipo => {
      const aba = ss.getSheetByName(tipo);
      if (!aba || aba.getLastRow() < 2) return; // Verifica se tem dados
      const C = tipo === ABA.TRIMESTRAL ? C_T : tipo === ABA.SEMANAL ? C_S : C_A;
      aba.getDataRange().getValues().slice(1).forEach(row => {
        // Validação robusta de dados da linha
        if (!row || row.length === 0 || !row[C.PROFESSOR]) return;
        if (String(row[C.PROFESSOR]).trim() !== String(professor).trim()) return;
        
        const r = { 
          tipo, 
          data: row[C.DATA] || '', 
          professor: String(row[C.PROFESSOR] || '').trim() 
        };
        
        if (tipo === ABA.TRIMESTRAL) {
          Object.assign(r, { 
            anoTurma: String(row[C_T.ANO_TURMA] || '').trim(), 
            componente: String(row[C_T.COMPONENTE] || '').trim(), 
            trimestre: String(row[C_T.TRIMESTRE] || '').trim(), 
            url: String(row[C_T.URL] || '').trim() 
          });
        } else if (tipo === ABA.SEMANAL) {
          Object.assign(r, { 
            turmas: String(row[C_S.TURMAS] || '').trim(), 
            componente: String(row[C_S.COMPONENTE] || '').trim(), 
            mes: String(row[C_S.MES] || '').trim(), 
            semanaInicio: String(row[C_S.SEMANA_INICIO] || '').trim(), 
            url: String(row[C_S.URL] || '').trim() 
          });
        } else {
          Object.assign(r, { 
            anoTurma: String(row[C_A.ANO_TURMA] || '').trim(), 
            componente: String(row[C_A.COMPONENTE] || '').trim(), 
            url: String(row[C_A.URL] || '').trim() 
          });
        }
        res.push(r);
      });
    });
    
    // Ordenação segura
    res.sort((a, b) => {
      const dataA = a.data || '';
      const dataB = b.data || '';
      return dataB > dataA ? 1 : -1;
    });

    const total  = res.length;
    const inicio = (pagAtual - 1) * tamPagina;
    const itens  = res.slice(inicio, inicio + tamPagina);
    return { itens, total, pagina: pagAtual, tamPagina, totalPaginas: Math.ceil(total / tamPagina) };
  } catch(e) {
    log('ERRO', 'getHistoricoProfessor', e.message);
    return { itens: [], total: 0, pagina: 1, tamPagina: 30, totalPaginas: 0, erro: e.message };
  }
}

// ═══════════════════════════════════════════════════════════
//  [BUG-3] PAINEL DA COORDENAÇÃO COM CACHE
// ═══════════════════════════════════════════════════════════
function getDadosPainel() {
  // Cache de 5 minutos via CacheService para evitar timeout na planilha
  const cache = CacheService.getScriptCache();
  const KEY   = 'painel_dados_v3';
  const hit   = cache.get(KEY);
  if (hit) {
    try { return JSON.parse(hit); } catch(e) { /* cache inválido, recalcula */ }
  }

  try {
    const ss  = abrirPlanilha();
    const cadastros = getCadastros();
    const registros = [];
    const comEntrega = new Set();
    let semanaisNoMes = 0;
    const mesAtual = Utilities.formatDate(new Date(), CONFIG.FUSO, "MM");

    [ABA.TRIMESTRAL, ABA.SEMANAL, ABA.ANUAL].forEach(tipo => {
      const aba = ss.getSheetByName(tipo);
      if (!aba || aba.getLastRow() < 2) return; // Verifica se tem dados
      
      const vals = aba.getDataRange().getValues();
      vals.slice(1).forEach(row => {
        // Validação robusta de dados
        if (!row || row.length === 0 || !row[0]) return;
        
        const professor = String(row[1] || '').trim();
        if (professor) comEntrega.add(professor);
        
        // Conta semanais do mês atual
        if (tipo === ABA.SEMANAL) {
          const dataCel = String(row[0] || '');
          const mesReg  = dataCel.length >= 5 ? dataCel.substring(3, 5) : '';
          if (mesReg === mesAtual) semanaisNoMes++;
        }
        
        const r = { 
          tipo, 
          data: row[0] || '', 
          professor: professor, 
          status: "Entregue" 
        };
        
        if (tipo === ABA.TRIMESTRAL) {
          Object.assign(r, { 
            anoTurma: String(row[2] || '').trim(), 
            componente: String(row[3] || '').trim(), 
            trimestre: String(row[4] || '').trim(), 
            url: String(row[5] || '').trim() 
          });
        } else if (tipo === ABA.SEMANAL) {
          Object.assign(r, { 
            turmas: String(row[2] || '').trim(), 
            componente: String(row[3] || '').trim(), 
            mes: String(row[4] || '').trim(), 
            semanaInicio: String(row[5] || '').trim(), 
            url: String(row[7] || '').trim() 
          });
        } else {
          Object.assign(r, { 
            anoTurma: String(row[2] || '').trim(), 
            componente: String(row[3] || '').trim(), 
            url: String(row[4] || '').trim() 
          });
        }
        registros.push(r);
      });
    });

    const totalProfessores = cadastros.professores.filter(p => p !== '— Configure a aba Cadastros na planilha —').length;
    const totalPendentes   = cadastros.professores.filter(p => !comEntrega.has(p) && p !== '— Configure a aba Cadastros na planilha —').length;

    const dados = {
      registros,
      totalEntregues:   registros.length,
      totalPendentes,
      totalProfessores,
      semanaisNoMes,
      geradoEm: ts('HH:mm'),
    };

    // Armazena no cache (máx 100KB por entrada do CacheService)
    try { cache.put(KEY, JSON.stringify(dados), CONFIG.CACHE_PAINEL_SEG); } catch(e) { /* dados > 100KB, ignora cache */ }

    return dados;
  } catch(e) {
    log('ERRO', 'getDadosPainel', e.message);
    return {
      registros: [],
      totalEntregues: 0,
      totalPendentes: 0,
      totalProfessores: 0,
      semanaisNoMes: 0,
      geradoEm: ts('HH:mm'),
      erro: e.message
    };
  }
}

// Invalida o cache quando novos planejamentos são salvos
function _invalidarCachePainel() {
  try { CacheService.getScriptCache().remove('painel_dados_v3'); } catch(e) {}
}

// ═══════════════════════════════════════════════════════════
//  RELATÓRIO DE PENDÊNCIAS
// ═══════════════════════════════════════════════════════════
function gerarRelatorioPendencias() {
  try {
    const ss       = abrirPlanilha();
    const cadastros = getCadastros();
    const comEntrega = new Set();

    [ABA.TRIMESTRAL, ABA.SEMANAL, ABA.ANUAL].forEach(tipo => {
      const aba = ss.getSheetByName(tipo);
      if (!aba) return;
      aba.getDataRange().getValues().slice(1).forEach(row => { if (row[1]) comEntrega.add(row[1]); });
    });

    const pendentes = cadastros.professores.filter(p => !comEntrega.has(p));
    const nome      = `Relatorio_Pendencias_${ts('yyyyMMdd_HHmm')}`;
    const ssPend    = SpreadsheetApp.create(nome);
    const aba       = ssPend.getActiveSheet();
    aba.setName("Pendências");
    aba.appendRow(["Professor(a)", "Status", "E-mail"]);
    aba.getRange(1,1,1,3).setFontWeight("bold").setBackground("#c0392b").setFontColor("white");
    pendentes.forEach(p => aba.appendRow([p, "⚠️ Sem planejamento entregue", cadastros.emails[p] || '']));
    aba.appendRow([]);
    aba.appendRow([`Gerado em: ${ts('dd/MM/yyyy HH:mm')} | Total pendente: ${pendentes.length} de ${cadastros.professores.length}`]);
    aba.autoResizeColumns(1, 3);

    DriveApp.getFolderById(CONFIG.PASTA_ID).addFile(DriveApp.getFileById(ssPend.getId()));
    log('INFO', 'gerarRelatorioPendencias', `${pendentes.length} pendências encontradas`);
    return { sucesso: true, url: ssPend.getUrl() };
  } catch(e) {
    log('ERRO', 'gerarRelatorioPendencias', e.message);
    return { sucesso: false, erro: e.message };
  }
}

// ═══════════════════════════════════════════════════════════
//  EXPORTAR PAINEL XLSX
// ═══════════════════════════════════════════════════════════
function exportarPainelXLSX() {
  try {
    const dados = getDadosPainel();
    if (!dados) return { sucesso: false, erro: "Sem dados." };

    const nome  = `Painel_Planejamentos_${ts('yyyyMMdd_HHmm')}`;
    const ss    = SpreadsheetApp.create(nome);
    const aba   = ss.getActiveSheet();
    aba.setName("Registros");
    aba.appendRow(["Tipo","Data/Hora","Professor(a)","Componente","Turma/Turmas","Período","Link PDF","Status"]);
    aba.getRange(1,1,1,8).setFontWeight("bold").setBackground("#1a3a6b").setFontColor("white");
    dados.registros.forEach(r => aba.appendRow([
      r.tipo, r.data, r.professor, r.componente,
      r.anoTurma||r.turmas||'', r.trimestre||r.semanaInicio||'',
      r.url||'', r.status
    ]));
    aba.autoResizeColumns(1, 8);

    DriveApp.getFolderById(CONFIG.PASTA_ID).addFile(DriveApp.getFileById(ss.getId()));
    return { sucesso: true, url: ss.getUrl() };
  } catch(e) {
    log('ERRO', 'exportarPainelXLSX', e.message);
    return { sucesso: false, erro: e.message };
  }
}

// ═══════════════════════════════════════════════════════════
//  [F4-2] QR CODE DE AUTENTICIDADE NOS PDFs
//  [F4-3] URL PERMANENTE DE VERIFICAÇÃO
// ═══════════════════════════════════════════════════════════

/**
 * Gera um ID único de verificação e grava na aba "Verificacao".
 * Retorna a URL de verificação pública e o src do QR Code via Google Charts.
 */
function _gerarRegistroVerificacao(tipo, dados, urlPDF) {
  try {
    const id  = Utilities.getUuid().substring(0, 13).toUpperCase(); // ex: A1B2-C3D4-E5F6
    const ss  = abrirPlanilha();
    let aba   = ss.getSheetByName("Verificacao");

    if (!aba) {
      aba = ss.insertSheet("Verificacao");
      aba.appendRow(["ID","Tipo","Professor(a)","Componente","Período","Data/Hora","Link PDF","Válido"]);
      aba.getRange(1,1,1,8).setFontWeight("bold").setBackground("#1a3a6b").setFontColor("white");
    }

    const periodo = tipo === ABA.TRIMESTRAL ? (dados.trimestre + 'º Tri')
                  : tipo === ABA.SEMANAL    ? dados.semanaInicio
                  : 'Anual';

    aba.appendRow([id, tipo, dados.professor, dados.componente, periodo, ts('dd/MM/yyyy HH:mm'), urlPDF, "SIM"]);

    // URL de verificação pública: mesma Web App com ?verificar=ID
    const urlBase  = ScriptApp.getService().getUrl();
    const urlVerif = `${urlBase}?verificar=${id}`;

    // QR Code via Google Charts API (gratuita, sem autenticação)
    const qrSrc = `https://chart.googleapis.com/chart?chs=80x80&cht=qr&chl=${encodeURIComponent(urlVerif)}&choe=UTF-8`;

    return { id, urlVerif, qrSrc };
  } catch(e) {
    log('AVISO', '_gerarRegistroVerificacao', `Falha ao gerar QR: ${e.message}`);
    return { id: '', urlVerif: '', qrSrc: '' };
  }
}

/**
 * Página HTML de verificação de documento.
 * Acessada via URL pública com ?verificar=ID
 */
function _paginaVerificacao(id) {
  try {
    // Validação rigorosa do ID para evitar injeção
    if (!id || typeof id !== 'string') {
      return HtmlService.createHtmlOutput(_htmlVerifErro('', 'ID de verificação inválido.'));
    }
    
    const idSanitizado = String(id).trim().substring(0, 50); // Limita tamanho
    if (!/^[A-Z0-9\-]+$/i.test(idSanitizado)) {
      return HtmlService.createHtmlOutput(_htmlVerifErro(id, 'ID contém caracteres inválidos.'));
    }
    
    const ss  = abrirPlanilha();
    const aba = ss.getSheetByName("Verificacao");

    if (!aba || aba.getLastRow() < 2) {
      return HtmlService.createHtmlOutput(_htmlVerifErro(id, 'Registro de verificação não encontrado.'));
    }

    const vals = aba.getDataRange().getValues();
    const reg  = vals.slice(1).find(row => String(row[0] || '').trim().toUpperCase() === idSanitizado.toUpperCase());

    if (!reg || !reg[0]) {
      return HtmlService.createHtmlOutput(_htmlVerifErro(id, 'Documento não encontrado no registro da escola.'));
    }

    // [V9-7][V9-15] Desestruturação via constantes de coluna — resistente a colunas extras
    const C        = SHEET_COLS.VERIFICACAO;
    const rid      = String(reg[C.ID] || '').trim();
    const tipo     = String(reg[C.TIPO] || '').trim();
    const professor  = String(reg[C.PROFESSOR] || '').trim();
    const componente = String(reg[C.COMPONENTE] || '').trim();
    const periodo  = String(reg[C.PERIODO] || '').trim();
    const dataHora = String(reg[C.DATA_HORA] || '').trim();
    const urlPDF   = String(reg[C.URL_PDF] || '').trim();
    const valido   = String(reg[C.VALIDO] || 'NÃO').trim().toUpperCase();
    
    const html = `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">
    <meta name="viewport" content="width=device-width,initial-scale=1">
    <title>Verificação de Documento — Colégio Itabatan</title>
    <style>
      body{font-family:Arial,sans-serif;background:#f4f6fb;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0}
      .box{background:#fff;border-radius:14px;padding:36px 40px;max-width:480px;width:100%;box-shadow:0 8px 40px rgba(0,0,0,.12);text-align:center}
      .icon{font-size:3.5rem;margin-bottom:12px}
      h1{font-size:1.2rem;color:#1a3a6b;margin-bottom:6px}
      .badge{display:inline-block;background:#dcfce7;color:#166534;padding:4px 16px;border-radius:20px;font-weight:700;font-size:.88rem;margin-bottom:20px}
      .badge.invalido{background:#fee2e2;color:#991b1b}
      table{width:100%;border-collapse:collapse;text-align:left;font-size:.92rem;margin-bottom:20px}
      td{padding:8px 10px;border-bottom:1px solid #e2e8f0}
      td:first-child{font-weight:700;color:#1a3a6b;width:40%}
      .id{font-family:monospace;font-size:.82rem;color:#64748b;margin-top:8px}
      a.btn{display:inline-block;background:#1a3a6b;color:#fff;padding:10px 24px;border-radius:8px;text-decoration:none;font-weight:700;font-size:.92rem}
    </style></head><body>
    <div class="box">
      <div class="icon">${valido === 'SIM' ? '✅' : '⚠️'}</div>
      <h1>Verificação de Documento Escolar</h1>
      <div class="badge ${valido === 'SIM' ? '' : 'invalido'}">${valido === 'SIM' ? 'DOCUMENTO AUTÊNTICO' : 'DOCUMENTO INVÁLIDO'}</div>
      <table>
        <tr><td>Tipo</td><td>${esc(tipo)}</td></tr>
        <tr><td>Professor(a)</td><td>${esc(professor)}</td></tr>
        <tr><td>Componente</td><td>${esc(componente)}</td></tr>
        <tr><td>Período</td><td>${esc(periodo)}</td></tr>
        <tr><td>Gerado em</td><td>${esc(dataHora)}</td></tr>
      </table>
      ${urlPDF && urlPDF.startsWith('http') ? `<a class="btn" href="${esc(urlPDF)}" target="_blank">📄 Abrir PDF Original</a>` : ''}
      <div class="id">ID: ${esc(rid)}</div>
      <div class="id" style="margin-top:6px">Colégio Municipal de 1º e 2º Graus de Itabatan — Mucuri/BA</div>
    </div>
    </body></html>`;

    return HtmlService.createHtmlOutput(html).setTitle('Verificação de Documento');
  } catch(e) {
    log('ERRO', '_paginaVerificacao', e.message);
    return HtmlService.createHtmlOutput(_htmlVerifErro(id, e.message));
  }
}

function _htmlVerifErro(id, motivo) {
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Verificação — Erro</title>
  <style>body{font-family:Arial,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;background:#f4f6fb}
  .box{background:#fff;border-radius:14px;padding:36px;max-width:420px;text-align:center;box-shadow:0 8px 32px rgba(0,0,0,.1)}
  h1{color:#c0392b;font-size:1.1rem} .id{font-family:monospace;color:#888;font-size:.8rem;margin-top:10px}</style></head>
  <body><div class="box"><div style="font-size:3rem">❌</div><h1>Documento não verificado</h1>
  <p style="color:#555;font-size:.9rem;margin-top:8px">${esc(motivo)}</p>
  <div class="id">ID consultado: ${esc(id)}</div></div></body></html>`;
}

// ═══════════════════════════════════════════════════════════
//  [F4-1] PROCESSAMENTO DA FILA OFFLINE
// ═══════════════════════════════════════════════════════════

/**
 * Processa um item da fila offline enviado pelo frontend.
 * Idêntico às funções de salvar normais, mas recebe o tipo explicitamente.
 */
function processarFilaOffline(itens) {
  const resultados = [];
  (itens || []).forEach(item => {
    try {
      // [V9-10] Valida estrutura do item antes de processar (corrige BUG-6)
      if (!item || typeof item !== 'object') {
        resultados.push({ id: null, sucesso: false, erro: 'Item de fila inválido (não é objeto)' });
        return;
      }
      if (!item.id) {
        resultados.push({ id: null, sucesso: false, erro: 'Item sem campo "id"' });
        return;
      }
      if (!item.tipo) {
        resultados.push({ id: item.id, sucesso: false, erro: 'Item sem campo "tipo"' });
        return;
      }
      if (!item.dados || typeof item.dados !== 'object') {
        resultados.push({ id: item.id, sucesso: false, erro: 'Item sem campo "dados" válido' });
        return;
      }

      let resultado;
      if      (item.tipo === 'trimestral') resultado = salvarTrimestral(item.dados);
      else if (item.tipo === 'semanal')    resultado = salvarSemanal(item.dados);
      else if (item.tipo === 'anual')      resultado = salvarAnual(item.dados);
      else resultado = { sucesso: false, erro: `Tipo desconhecido: ${item.tipo}` };

      resultados.push({ id: item.id, ...resultado });
      log('INFO', 'processarFilaOffline', `Item ${item.id} (${item.tipo}) — ${resultado.sucesso ? 'OK' : 'ERRO: ' + resultado.erro}`);
    } catch(e) {
      resultados.push({ id: item.id, sucesso: false, erro: e.message });
      log('ERRO', 'processarFilaOffline', `Item ${item.id}: ${e.message}`);
    }
  });
  return resultados;
}

// ═══════════════════════════════════════════════════════════
//  GERADORES DE HTML PARA PDF
// ═══════════════════════════════════════════════════════════
// ─── CSS BASE atualizado com estilos para QR Code ─────────
const CSS_BASE = `
  @page { size: A4 landscape; margin: 15mm 12mm; }
  * { box-sizing: border-box; font-family: Arial, sans-serif; }
  body { margin: 0; padding: 0; font-size: 10pt; }
  .cab { text-align: center; margin-bottom: 8px; }
  .cab strong { font-size: 12pt; display: block; }
  .cab small  { font-size: 8pt; color: #555; }
  .linha-id { display: flex; gap: 20px; font-size: 9pt; flex-wrap: wrap;
    border-top: 1pt solid #000; border-bottom: 1pt solid #000; padding: 4px 0; margin-bottom: 6px; }
  .tit { text-align: center; font-weight: bold; font-size: 11pt;
    text-decoration: underline; margin: 6px 0; }
  table { width: 100%; border-collapse: collapse; margin-top: 4px; table-layout: fixed; }
  th { background: #1a3a6b; color: white; padding: 6px 8px; font-size: 9pt; text-align: left; border: 1pt solid #555; word-break: break-word; overflow-wrap: break-word; }
  td { border: 1pt solid #aaa; padding: 6px 8px; vertical-align: top; font-size: 9pt; word-break: break-word; overflow-wrap: break-word; }
  .rod { display: flex; justify-content: space-between; align-items: flex-end; margin-top: 16px; font-size: 9pt; }
  .ass { border-top: 1pt solid #000; padding-top: 2px; min-width: 200px; text-align: center; }
  .obs { margin-top: 10px; font-size: 8.5pt; color: #333; padding: 6px;
    background: #f7f9fc; border-left: 3pt solid #1a3a6b; word-break: break-word; overflow-wrap: break-word; }
  .qr-bloco { text-align: center; }
  .qr-bloco img { width: 70px; height: 70px; display: block; margin: 0 auto 3px; }
  .qr-bloco small { font-size: 6pt; color: #888; display: block; }
`;

function _cab() {
  const nome    = _cfg('escola_nome',    'COLÉGIO MUNICIPAL DE 1º E 2º GRAUS DE ITABATAN');
  const criacao = _cfg('escola_criacao', 'Criação 062/94 - Aut. CEE 236/95 DO 25/01/1996');
  const end     = _cfg('escola_endereco','Av. Itapetinga, 305 - Gazzinelândia - Itabatan - Mucuri / BA');
  const cep     = _cfg('escola_cep',     '45.936-000');
  return `<div class="cab">
    <strong>${esc(nome)}</strong>
    <small>${esc(criacao)}</small><br>
    <small>${esc(end)} — CEP.: ${esc(cep)}</small>
  </div>`;
}

function _rodapeComQR(verif, coordNome) {
  const coord = coordNome || _cfg('coordenacao_nome', 'Coordenadora Pedagógica');
  if (!verif || !verif.qrSrc) {
    return `<div class="rod">
      <div class="ass">Assinatura do(a) Professor(a)</div>
      <div class="ass">Assinatura da ${esc(coord)}</div>
    </div>`;
  }
  return `<div class="rod">
    <div class="ass">Assinatura do(a) Professor(a)</div>
    <div class="ass">Assinatura da ${esc(coord)}</div>
    <div class="qr-bloco">
      <img src="${verif.qrSrc}" alt="QR de autenticidade">
      <small>Verifique a autenticidade</small>
      <small style="font-size:5.5pt">${verif.id || ''}</small>
    </div>
  </div>`;
}

function htmlTrimestral(d, verif) {
  // [V9-3] Usa _cfg() para buscar o ano letivo configurado pela coordenação
  const ano      = _cfg('ano_letivo', String(CONFIG.ANO_LETIVO)) || String(CONFIG.ANO_LETIVO);
  const avalGeral = _cfg('avaliacao_geral',   '12,0');
  const avalParcial = _cfg('avaliacao_parcial', '18,0');
  const coordNome = _cfg('coordenacao_nome', 'Coordenadora Pedagógica');

  const linhas = (d.linhas||[]).map((l, i) => `<tr>
    <td style="width:22%;background:#f7f9fc">${esc(l.unidade)}</td>
    <td style="width:35%">${nl2br(esc(l.habilidades))}</td>
    <td style="width:20%">${nl2br(esc(l.metodologia))}</td>
    <td style="width:23%">
      ${i===0 ? `<div style="font-size:8pt;color:#333;margin-bottom:4px">OBS.: Descrever processo, critérios e distribuição da pontuação</div>
      <div style="font-weight:bold">AVALIAÇÃO GERAL: ${esc(avalGeral)}</div>
      <div style="font-weight:bold">AVALIAÇÃO PARCIAL: ${esc(avalParcial)} SENDO:</div>` : ''}
      ${nl2br(esc(l.avaliacao))}
    </td></tr>`).join('') || `<tr><td style="height:120px"></td><td></td><td></td>
    <td><div style="font-size:8pt;color:#333">OBS.: Descrever processo, critérios e distribuição da pontuação</div>
    <div style="font-weight:bold">AVALIAÇÃO GERAL: ${esc(avalGeral)}</div>
    <div style="font-weight:bold">AVALIAÇÃO PARCIAL: ${esc(avalParcial)}</div></td></tr>`;

  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>${CSS_BASE}</style></head><body>
  ${_cab()}
  <div class="linha-id">
    <span><strong>Ano/Turma:</strong> ${esc(d.anoTurma)}</span>
    <span><strong>Componente:</strong> ${esc(d.componente)}</span>
    <span><strong>Professor(a):</strong> ${esc(d.professor)}</span>
  </div>
  <div class="tit">PLANEJAMENTO ${esc(d.trimestre)} TRIMESTRE ${ano}</div>
  <table><thead><tr>
    <th>UNIDADE TEMÁTICA / OBJETOS DE CONHECIMENTO</th>
    <th>HABILIDADES</th><th>METODOLOGIA / RECURSOS</th><th>AVALIAÇÃO</th>
  </tr></thead><tbody>${linhas}</tbody></table>
  ${d.obs ? `<div class="obs"><strong>Observações:</strong> ${nl2br(esc(d.obs))}</div>` : ''}
  ${_rodapeComQR(verif, coordNome)}
  </body></html>`;
}

function htmlSemanal(d, verif) {
  const diasL  = ["SEGUNDA-FEIRA","TERÇA-FEIRA","QUARTA-FEIRA","QUINTA-FEIRA","SEXTA-FEIRA"];
  const datas  = (d.datas||[]).map(dt => dt || '___/___');
  const campos = ['habilidade','estrategia','recursos','observacao'];
  const labC   = ['HABILIDADE','ESTRATÉGIA','RECURSOS','OBSERVAÇÃO'];
  const hdrDias = diasL.map((dia, i) =>
    `<th style="text-align:center">${dia}<br><span style="font-weight:normal;font-size:8pt">${datas[i]}</span></th>`
  ).join('');
  const linhas = campos.map((campo, ci) => {
    const cols = [0,1,2,3,4].map(di =>
      `<td>${nl2br(esc(d.semana&&d.semana[di] ? d.semana[di][campo]||'' : ''))}</td>`
    ).join('');
    return `<tr><th style="background:#2c5282;font-size:8pt;width:80px">${labC[ci]}</th>${cols}</tr>`;
  }).join('');

  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>${CSS_BASE}</style></head><body>
  ${_cab()}
  <div class="linha-id">
    <span><strong>Componente:</strong> ${esc(d.componente)}</span>
    <span><strong>Professor(a):</strong> ${esc(d.professor)}</span>
    <span><strong>Turmas:</strong> ${esc(d.turmas)}</span>
    <span><strong>Mês:</strong> ${esc(d.mes)}</span>
  </div>
  <div class="tit">PLANEJAMENTO SEMANAL – INÍCIO: ${esc(d.semanaInicio)} | TÉRMINO: ${esc(d.semanaFim)}</div>
  <table><thead><tr><th style="width:80px"></th>${hdrDias}</tr></thead><tbody>${linhas}</tbody></table>
  ${d.obs ? `<div class="obs"><strong>Observações:</strong> ${nl2br(esc(d.obs))}</div>` : ''}
  ${_rodapeComQR(verif)}
  </body></html>`;
}

function htmlAnual(d, verif) {
  // [V9-3] Usa _cfg() para buscar o ano letivo configurado pela coordenação
  // em vez de usar CONFIG.ANO_LETIVO fixo do código
  const ano   = _cfg('ano_letivo', String(CONFIG.ANO_LETIVO)) || String(CONFIG.ANO_LETIVO);
  const coordNome = _cfg('coordenacao_nome', 'Coordenadora Pedagógica');
  const meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const linhas = meses.map((mes, i) =>
    `<tr><td style="width:12%;background:#f7f9fc;font-weight:bold">${mes}</td>
    <td>${nl2br(esc(d.meses&&d.meses[i]?d.meses[i]:''))}</td></tr>`
  ).join('');

  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>${CSS_BASE}</style></head><body>
  ${_cab()}
  <div class="linha-id">
    <span><strong>Componente:</strong> ${esc(d.componente)}</span>
    <span><strong>Professor(a):</strong> ${esc(d.professor)}</span>
    <span><strong>Ano/Turma:</strong> ${esc(d.anoTurma)}</span>
  </div>
  <div class="tit">PLANEJAMENTO ANUAL ${ano}</div>
  ${d.objetivo ? `<p style="font-size:9pt;margin-bottom:4px"><strong>Objetivo Geral:</strong> ${esc(d.objetivo)}</p>` : ''}
  ${d.bncc     ? `<p style="font-size:9pt;margin-bottom:6px"><strong>Competências BNCC:</strong> ${esc(d.bncc)}</p>`  : ''}
  <table><thead><tr><th style="width:12%">MÊS</th><th>CONTEÚDOS / UNIDADES TEMÁTICAS / HABILIDADES</th></tr></thead>
  <tbody>${linhas}</tbody></table>
  ${d.obs ? `<div class="obs"><strong>Observações:</strong> ${nl2br(esc(d.obs))}</div>` : ''}
  ${_rodapeComQR(verif, coordNome)}
  </body></html>`;
}

// ═══════════════════════════════════════════════════════════
//  REGISTRO NA PLANILHA
// ═══════════════════════════════════════════════════════════
function registrar(tipo, dados, url) {
  // LockService: evita conflito quando múltiplos professores enviam simultaneamente
  const lock = LockService.getScriptLock();
  let lockObtido = false;
  try {
    // Validação de entrada
    if (!tipo || !dados || typeof dados !== 'object') {
      log('ERRO', 'registrar', 'Tipo ou dados inválidos');
      return;
    }

    lockObtido = lock.tryLock(10000); // espera até 10 segundos
    if (!lockObtido) {
      log('AVISO', 'registrar', `Timeout ao aguardar lock para ${tipo} - ${dados.professor}`);
      return;
    }

    const ss  = abrirPlanilha();
    let aba   = ss.getSheetByName(tipo);
    const ago = ts('dd/MM/yyyy HH:mm');

    if (!aba) {
      aba = ss.insertSheet(tipo);
      const hdrs = {
        [ABA.TRIMESTRAL]: ["Data/Hora","Professor(a)","Ano/Turma","Componente","Trimestre","Link PDF"],
        [ABA.SEMANAL]:    ["Data/Hora","Professor(a)","Turmas","Componente","Mês","Semana Início","Semana Fim","Link PDF"],
        [ABA.ANUAL]:      ["Data/Hora","Professor(a)","Ano/Turma","Componente","Link PDF"],
      };

      const header = hdrs[tipo];
      if (!header) {
        log('ERRO', 'registrar', `Tipo de aba desconhecido: ${tipo}`);
        return;
      }

      aba.appendRow(header);
      aba.getRange(1, 1, 1, header.length).setFontWeight("bold").setBackground("#1a3a6b").setFontColor("white");
    }

    // Prepara dados com valores padrão para evitar undefined
    const prof = String(dados.professor || '').trim();
    const comp = String(dados.componente || '').trim();
    const urlFinal = String(url || '').trim();

    if (tipo === ABA.TRIMESTRAL) {
      const anoTurma = String(dados.anoTurma || '').trim();
      const trim = String(dados.trimestre || '').replace('º', '') + 'º';
      aba.appendRow([ago, prof, anoTurma, comp, trim, urlFinal]);
    } else if (tipo === ABA.SEMANAL) {
      const turmas = String(dados.turmas || '').trim();
      const mes = String(dados.mes || '').trim();
      const semIni = String(dados.semanaInicio || '').trim();
      const semFim = String(dados.semanaFim || '').trim();
      aba.appendRow([ago, prof, turmas, comp, mes, semIni, semFim, urlFinal]);
    } else {
      const anoTurma = String(dados.anoTurma || '').trim();
      aba.appendRow([ago, prof, anoTurma, comp, urlFinal]);
    }

    // Invalida cache do painel após novo registro
    _invalidarCachePainel();

    log('INFO', 'registrar', `Registro salvo na aba ${tipo} - ${prof}`);
  } catch(e) {
    log('ERRO', 'registrar', `Falha ao registrar na aba ${tipo}: ${e.message}`);
    // Não lança exceção para não impedir salvamento do PDF
  } finally {
    if (lockObtido) lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════
//  [MELHORIA] LOG ESTRUTURADO NA PLANILHA
// ═══════════════════════════════════════════════════════════
function log(nivel, funcao, mensagem, usuario) {
  // LockService: evita sobreposição de linhas no log sob carga concorrente
  const lock = LockService.getScriptLock();
  let lockObtido = false;
  try {
    const ss  = abrirPlanilha();
    let aba   = ss.getSheetByName(ABA.LOGS);
    if (!aba) {
      aba = ss.insertSheet(ABA.LOGS);
      aba.appendRow(["Data/Hora", "Nível", "Função", "Mensagem", "Usuário"]);
      aba.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#2c3e50").setFontColor("white");
      aba.setFrozenRows(1);
    }
    lockObtido = lock.tryLock(5000); // espera até 5 segundos
    const cores = { 'ERRO': '#c0392b', 'AVISO': '#d97706', 'INFO': '#2d7a4f' };
    const linha = aba.getLastRow() + 1;
    aba.appendRow([ts('dd/MM/yyyy HH:mm:ss'), nivel, funcao, mensagem, usuario || '']);
    if (cores[nivel]) {
      aba.getRange(linha, 1, 1, 5).setBackground(nivel === 'ERRO' ? '#fee2e2' : nivel === 'AVISO' ? '#fef9c3' : '#f0fdf4');
      aba.getRange(linha, 2).setFontWeight('bold').setFontColor(cores[nivel]);
    }
    Logger.log(`[${nivel}] ${funcao}: ${mensagem}`);
  } catch(e) {
    Logger.log(`[LOG FAIL] ${e.message}`);
  } finally {
    if (lockObtido) lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════
//  [V7-1..V7-7] PERFIS DE ACESSO + PASTA INDIVIDUAL DO PROFESSOR
// ═══════════════════════════════════════════════════════════

/**
 * Retorna o perfil do usuário atual.
 * Para professores: cria a pasta individual no 1º acesso e retorna o pastaId.
 * { nome, email, perfil, pastaId, pastaUrl }
 */
// [V9-LOGOUT] Retorna a URL base do exec (sem parâmetros) para redirecionar à landing page
function getExecUrl() {
  return ScriptApp.getService().getUrl();
}

function getPerfil() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email || !email.includes('@')) {
      return { nome: 'Visitante', email: '', perfil: 'desconhecido' };
    }
    
    const emailNormalizado = email.toLowerCase().trim();

    // [V7-5] Coordenação — acesso total, sem pasta individual
    if (_ehCoordenacao(emailNormalizado)) {
      return { nome: 'Coordenação', email: emailNormalizado, perfil: 'coordenacao', pastaId: null };
    }

    // Verifica na aba de cadastros
    const cadastros = getCadastros();
    const profEntry = Object.entries(cadastros.emails || {})
      .find(([, em]) => String(em || '').toLowerCase().trim() === emailNormalizado);

    if (profEntry) {
      const nomeProfessor = profEntry[0];

      // [V7-1] Provisiona pasta individual no 1º login
      let pasta = null;
      try {
        pasta = _provisionarPastaProfessor(nomeProfessor, emailNormalizado);
      } catch(ePasta) {
        log('AVISO', 'getPerfil', `Falha ao provisionar pasta para ${nomeProfessor}: ${ePasta.message}`);
      }

      // [V8-1] Busca turmas e componentes atribuídos ao professor
      const atribuicoes = _getAtribuicoesProfessor(nomeProfessor);

      log('INFO', 'getPerfil', `Login: ${nomeProfessor} (${emailNormalizado})${pasta ? ' — pasta: ' + pasta.getId() : ''}`);
      return {
        nome:        nomeProfessor,
        email:       emailNormalizado,
        perfil:      'professor',
        pastaId:     pasta ? pasta.getId() : null,
        pastaUrl:    pasta ? pasta.getUrl() : null,
        turmas:      atribuicoes.turmas,      // [V8-1] pré-seleção automática
        componentes: atribuicoes.componentes, // [V8-1] componentes deste professor
      };
    }

    // E-mail Google reconhecido mas não cadastrado
    log('AVISO', 'getPerfil', `Acesso negado: ${emailNormalizado} não está na aba Cadastros`);
    return { nome: email.split('@')[0], email: emailNormalizado, perfil: 'desconhecido' };

  } catch(e) {
    log('AVISO', 'getPerfil', e.message);
    return { nome: 'Modo Dev', email: '', perfil: 'dev', pastaId: null };
  }
}

/**
 * [V7-1..V7-3] Cria (ou recupera) a pasta individual do professor no Drive.
 * Estrutura: PASTA_RAIZ/Professores/NomeProfessor/Trimestral | Semanal | Anual | Docs
 * Permissão: professor (Editor) + coordenação (Editor) — somente eles.
 */
function _provisionarPastaProfessor(nome, emailProfessor) {
  const raiz         = DriveApp.getFolderById(CONFIG.PASTA_ID);
  const pastaProfRef = subpasta(raiz.getId(), 'Professores');

  // Normaliza nome para nome de pasta (remove acentos e caracteres especiais)
  const nomeDir = nome.normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/[^a-zA-Z0-9\s]/g,'').trim().replace(/\s+/g,'_');

  // Verifica se já existe para não recriar a cada login
  const existentes = pastaProfRef.getFoldersByName(nomeDir);
  if (existentes.hasNext()) {
    return existentes.next(); // já provisionada
  }

  // [V7-2] Cria pasta com compartilhamento restrito
  const pastaProfessor = pastaProfRef.createFolder(nomeDir);

  // Cria subpastas internas
  pastaProfessor.createFolder('Trimestral');
  pastaProfessor.createFolder('Semanal');
  pastaProfessor.createFolder('Anual');
  pastaProfessor.createFolder('Docs');

  // [V7-2] Concede acesso ao professor como Editor
  try {
    pastaProfessor.addEditor(emailProfessor);
  } catch(e) {
    log('AVISO', '_provisionarPastaProfessor', `Falha ao conceder acesso a ${emailProfessor}: ${e.message}`);
  }

  // [V9-14] Concede acesso a coordenação como Viewer (não Editor) para evitar alteração acidental de PDFs
  CONFIG.EMAILS_COORDENACAO.forEach(emailCoord => {
    try {
      if (emailCoord && !emailCoord.startsWith('COLE_')) {
        pastaProfessor.addViewer(emailCoord);
      }
    } catch(e2) {
      log('AVISO', '_provisionarPastaProfessor', `Falha ao conceder acesso à coordenação (${emailCoord}): ${e2.message}`);
    }
  });

  log('INFO', '_provisionarPastaProfessor',
    `Pasta criada: ${nomeDir} — acesso: ${emailProfessor} + coordenação`);

  return pastaProfessor;
}

/**
 * [V7-4] Retorna a pasta correta para salvar um PDF de um professor.
 * Se o professor tiver pasta individual → usa ela + subpasta do tipo.
 * Fallback (modo dev / coordenação testando) → usa subpasta global.
 */
function _pastaDestino(tipo, nomeProfessor) {
  // Tenta encontrar a pasta individual do professor
  try {
    const raiz    = DriveApp.getFolderById(CONFIG.PASTA_ID);
    const profDir = subpasta(raiz.getId(), 'Professores');

    const nomeDir = (nomeProfessor || '').normalize('NFD')
      .replace(/[\u0300-\u036f]/g,'')
      .replace(/[^a-zA-Z0-9\s]/g,'').trim().replace(/\s+/g,'_');

    if (!nomeDir) throw new Error('Nome inválido');

    const profExist = profDir.getFoldersByName(nomeDir);
    if (!profExist.hasNext()) throw new Error('Pasta do professor não existe ainda');

    const pastaProfessor = profExist.next();
    const nomeSub        = { T:'Trimestral', S:'Semanal', A:'Anual', Docs:'Docs' }[tipo] || tipo;
    return subpasta(pastaProfessor.getId(), nomeSub);

  } catch(e) {
    // Fallback: usa subpastas globais (comportamento anterior)
    const nomeSub = { T: SP.TRIMESTRAL, S: SP.SEMANAL, A: SP.ANUAL }[tipo] || tipo;
    return subpasta(CONFIG.PASTA_ID, nomeSub);
  }
}

/**
 * [V7-6] Guard de segurança no backend.
 * Verifica se o usuário autenticado tem permissão para executar a operação.
 * Lança exceção se não tiver — impede acesso via chamada direta à API.
 */
function _verificarPermissao(operacao) {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) throw new Error('Sessão não identificada.');

    // Coordenação tem tudo liberado
    if (_ehCoordenacao(email)) return { email, perfil: 'coordenacao' };

    // Professor cadastrado — operações restritas ao seu próprio contexto
    const cadastros  = getCadastros();
    const profEntry  = Object.entries(cadastros.emails || {})
      .find(([, em]) => em.toLowerCase() === email.toLowerCase());

    if (profEntry) return { email, perfil: 'professor', nome: profEntry[0] };

    // Operações públicas permitidas mesmo sem cadastro (verificação de QR)
    const opsPublicas = ['verificarDocumento'];
    if (opsPublicas.includes(operacao)) return { email, perfil: 'publico' };

    throw new Error(`Usuário ${email} não está autorizado para "${operacao}".`);
  } catch(e) {
    log('AVISO', '_verificarPermissao', e.message);
    throw e; // re-lança para que o frontend receba o erro
  }
}

/**
 * [V7-6] Provisiona pastas para todos os professores cadastrados.
 * Execute manualmente no editor quando quiser pré-criar pastas em lote.
 */
function provisionarTodasAsPastas() {
  const cadastros = getCadastros();
  let criadas = 0, existentes = 0;

  cadastros.professores.forEach(nome => {
    const email = cadastros.emails && cadastros.emails[nome];
    if (!email) {
      log('AVISO', 'provisionarTodasAsPastas', `Professor sem e-mail: ${nome} — pulado`);
      return;
    }
    try {
      const raiz    = DriveApp.getFolderById(CONFIG.PASTA_ID);
      const profDir = subpasta(raiz.getId(), 'Professores');
      const nomeDir = nome.normalize('NFD').replace(/[\u0300-\u036f]/g,'')
        .replace(/[^a-zA-Z0-9\s]/g,'').trim().replace(/\s+/g,'_');
      const existe  = profDir.getFoldersByName(nomeDir).hasNext();

      if (!existe) {
        _provisionarPastaProfessor(nome, email);
        criadas++;
        Logger.log(`✅ Criada: ${nome}`);
      } else {
        existentes++;
      }
    } catch(e2) {
      log('ERRO', 'provisionarTodasAsPastas', `${nome}: ${e2.message}`);
    }
  });

  const msg = `Concluído: ${criadas} pasta(s) criada(s), ${existentes} já existia(m).`;
  log('INFO', 'provisionarTodasAsPastas', msg);
  Logger.log(msg);
  return msg;
}

// ═══════════════════════════════════════════════════════════
//  [F2-2] CALENDÁRIO ESCOLAR
// ═══════════════════════════════════════════════════════════

/**
 * Retorna lista de eventos do calendário escolar.
 * Aba "Calendario" — colunas: Data (dd/MM/yyyy), Tipo, Descrição
 * Tipos: FERIADO | EVENTO | RECESSO | LETIVO_ESPECIAL
 */
function getCalendario() {
  try {
    const ss  = abrirPlanilha();
    let aba   = ss.getSheetByName(ABA.CALENDARIO);
    if (!aba) aba = _criarAbaCalendario(ss);

    const vals = aba.getDataRange().getValues();
    return vals.slice(1)
      .filter(row => row[0])
      .map(row => ({
        data:      _normalizarData(row[0]),
        tipo:      String(row[1]).trim().toUpperCase(),
        descricao: String(row[2]).trim(),
      }))
      .filter(e => e.data);
  } catch(e) {
    log('ERRO', 'getCalendario', e.message);
    return [];
  }
}

function _normalizarData(val) {
  // Aceita Date do Sheets, string dd/MM/yyyy ou yyyy-MM-dd
  if (val instanceof Date) {
    return Utilities.formatDate(val, CONFIG.FUSO, 'dd/MM/yyyy');
  }
  const s = String(val).trim();
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const [y,m,d] = s.split('-');
    return `${d}/${m}/${y}`;
  }
  return '';
}

function _criarAbaCalendario(ss) {
  const aba = ss.insertSheet(ABA.CALENDARIO);
  aba.appendRow(["Data (dd/mm/aaaa)", "Tipo", "Descrição"]);
  aba.getRange(1,1,1,3).setFontWeight("bold").setBackground("#1a3a6b").setFontColor("white");

  // Feriados nacionais fixos e exemplos de eventos escolares
  const ano = CONFIG.ANO_LETIVO;
  const eventos = [
    [`01/01/${ano}`, "FERIADO",  "Confraternização Universal"],
    [`03/03/${ano}`, "FERIADO",  "Carnaval (ponto facultativo)"],
    [`04/03/${ano}`, "FERIADO",  "Carnaval"],
    [`05/03/${ano}`, "FERIADO",  "Quarta-feira de Cinzas (meio expediente)"],
    [`18/04/${ano}`, "FERIADO",  "Sexta-feira Santa"],
    [`21/04/${ano}`, "FERIADO",  "Tiradentes"],
    [`01/05/${ano}`, "FERIADO",  "Dia do Trabalho"],
    [`19/06/${ano}`, "FERIADO",  "Corpus Christi"],
    [`07/09/${ano}`, "FERIADO",  "Independência do Brasil"],
    [`12/10/${ano}`, "FERIADO",  "Nossa Senhora Aparecida"],
    [`02/11/${ano}`, "FERIADO",  "Finados"],
    [`15/11/${ano}`, "FERIADO",  "Proclamação da República"],
    [`25/12/${ano}`, "FERIADO",  "Natal"],
    [`14/07/${ano}`, "RECESSO",  "Início do recesso escolar de julho"],
    [`18/07/${ano}`, "RECESSO",  "Fim do recesso escolar de julho"],
    [`10/03/${ano}`, "EVENTO",   "Reunião de pais e mestres — 1º trimestre"],
    [`15/07/${ano}`, "EVENTO",   "Conselho de classe — 2º trimestre"],
  ];
  aba.getRange(2, 1, eventos.length, 3).setValues(eventos);
  aba.getRange(2, 1, eventos.length, 1).setNumberFormat('@STRING@'); // força texto nas datas
  aba.autoResizeColumns(1, 3);

  aba.getRange(eventos.length + 3, 1).setValue(
    "INSTRUÇÃO: Adicione feriados e eventos. Tipo: FERIADO | RECESSO | EVENTO | LETIVO_ESPECIAL. " +
    "Datas no formato dd/mm/aaaa. Dias com tipo FERIADO ou RECESSO aparecem bloqueados no sistema."
  ).setFontStyle("italic").setFontColor("#555555");

  return aba;
}

// ═══════════════════════════════════════════════════════════
//  [F2-3] BANCO DE HABILIDADES BNCC
// ═══════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════
//  [F2-3] BANCO DE HABILIDADES BNCC — ver implementação v8 abaixo
// ═══════════════════════════════════════════════════════════
// A função buscarHabilidadesBNCC está na seção [V8-3] com API oficial.

function _criarAbaBNCC(ss) {
  const aba = ss.insertSheet(ABA.BNCC);
  aba.appendRow(["Código", "Componente", "Ano/Segmento", "Descrição da Habilidade"]);
  aba.getRange(1,1,1,4).setFontWeight("bold").setBackground("#1a3a6b").setFontColor("white");

  // Amostra representativa de habilidades para Anos Finais (6º–9º)
  // A coordenação deve completar com todas as habilidades relevantes
  const habilidades = [
    // Língua Portuguesa
    ["EF06LP01","Língua Portuguesa","6º ano","Identificar as marcas linguísticas da identidade sociocultural dos sujeitos na escuta de variantes do português."],
    ["EF06LP05","Língua Portuguesa","6º ano","Reconhecer as variedades da língua falada, o conceito de norma-padrão e o preconceito linguístico."],
    ["EF07LP01","Língua Portuguesa","7º ano","Identificar, em textos, o uso de recursos expressivos para estabelecer a persuasão do leitor."],
    ["EF08LP01","Língua Portuguesa","8º ano","Analisar os efeitos de sentido produzidos pelo uso de recursos expressivos em diferentes textos."],
    ["EF09LP01","Língua Portuguesa","9º ano","Produzir textos de diferentes gêneros, considerando a situação comunicativa, o tema, a finalidade."],
    // Matemática
    ["EF06MA01","Matemática","6º ano","Comparar, ordenar, ler e escrever números naturais e números racionais em diferentes contextos."],
    ["EF06MA12","Matemática","6º ano","Resolver e elaborar problemas com números racionais, envolvendo as quatro operações fundamentais."],
    ["EF07MA01","Matemática","7º ano","Resolver e elaborar problemas com números inteiros, envolvendo as quatro operações fundamentais."],
    ["EF07MA17","Matemática","7º ano","Calcular a probabilidade de eventos, expressando-a por número racional."],
    ["EF08MA01","Matemática","8º ano","Resolver e elaborar problemas com potências de base 10."],
    ["EF08MA15","Matemática","8º ano","Aplicar o Teorema de Pitágoras para resolver situações-problema envolvendo figuras planas."],
    ["EF09MA01","Matemática","9º ano","Reconhecer e utilizar números irracionais e números reais em situações de aprendizagem."],
    // Ciências
    ["EF06CI01","Ciências","6º ano","Classificar como reagente ou produto as substâncias envolvidas nas transformações químicas."],
    ["EF06CI11","Ciências","6º ano","Identificar as diferentes camadas que compõem a Terra (da crosta ao núcleo)."],
    ["EF07CI01","Ciências","7º ano","Classificar os seres vivos em grandes grupos — animais, plantas, fungos, protistas e moneras."],
    ["EF07CI09","Ciências","7º ano","Interpretar as condições de saúde da comunidade, considerando saneamento básico, energia elétrica."],
    ["EF08CI01","Ciências","8º ano","Analisar e representar as transformações e conservações em sistemas que envolvam quantidade de matéria."],
    ["EF09CI03","Ciências","9º ano","Identificar modelos que descrevem a estrutura do átomo."],
    // História
    ["EF06HI01","História","6º ano","Identificar diferentes formas de compreensão da noção de tempo e de periodização."],
    ["EF07HI01","História","7º ano","Explicar o significado de Modernidade e suas lógicas de inclusão e exclusão."],
    ["EF08HI01","História","8º ano","Identificar os mecanismos e as dinâmicas de exploração colonial que incidiram sobre as sociedades americanas."],
    ["EF09HI01","História","9º ano","Descrever e contextualizar os principais aspectos sociais, culturais, econômicos e políticos da emergência da modernidade."],
    // Geografia
    ["EF06GE01","Geografia","6º ano","Comparar modificações das paisagens nos lugares de vivência dos estudantes."],
    ["EF07GE01","Geografia","7º ano","Avaliar, por meio de exemplos extraídos dos contextos locais e global, a importância da biodiversidade."],
    ["EF08GE01","Geografia","8º ano","Descrever as rotas de dispersão da população humana pelo planeta."],
    ["EF09GE01","Geografia","9º ano","Analisar a distribuição territorial da população mundial e do Brasil."],
    // Arte
    ["EF69AR01","Arte","6º ao 9º ano","Pesquisar, apreciar e analisar formas distintas das artes visuais tradicionais e contemporâneas."],
    ["EF69AR31","Arte","6º ao 9º ano","Relacionar as práticas artísticas às diferentes dimensões da vida social, cultural, política, histórica."],
    // Educação Física
    ["EF67EF01","Educação Física","6º e 7º ano","Experimentar e fruir, na escola e fora dela, jogos eletrônicos diversos."],
    ["EF89EF01","Educação Física","8º e 9º ano","Experimentar e fruir diferentes papéis (atuação, organização, arbitragem) na vivência de esportes de invasão."],
    // Inglês
    ["EF06LI01","Inglês","6º ano","Explorar formas de interação verbal em língua inglesa."],
    ["EF09LI13","Inglês","9º ano","Selecionar texto em língua inglesa de maior extensão para aprofundar um tema de pesquisa."],
  ];

  aba.getRange(2, 1, habilidades.length, 4).setValues(habilidades);
  aba.autoResizeColumns(1, 2);
  aba.setColumnWidth(4, 500);

  aba.getRange(habilidades.length + 3, 1).setValue(
    "INSTRUÇÃO: Adicione habilidades da BNCC. Formato: Código | Componente | Ano/Segmento | Descrição completa. " +
    "O sistema busca por qualquer combinação dessas colunas."
  ).setFontStyle("italic").setFontColor("#555555");

  return aba;
}

// ═══════════════════════════════════════════════════════════
//  [F2-1] LEMBRETE AUTOMÁTICO SEMANAL
// ═══════════════════════════════════════════════════════════

/**
 * CONFIGURAÇÃO DO TRIGGER:
 * No Apps Script → Gatilhos → Adicionar gatilho:
 *   Função: lembreteSemanaPlanejamento
 *   Evento: Controlado por tempo → Semanal → Sexta-feira → Entre 15h–16h
 *
 * Verifica quem não enviou o planejamento semanal na semana corrente e envia lembrete.
 */
function lembreteSemanaPlanejamento() {
  try {
    const ss       = abrirPlanilha();
    const cadastros = getCadastros();
    const aba       = ss.getSheetByName(ABA.SEMANAL);

    // Descobre a semana atual (segunda-feira como início)
    const hoje       = new Date();
    const diaSemana  = hoje.getDay(); // 0=Dom, 1=Seg ... 6=Sáb
    const diffSeg    = diaSemana === 0 ? -6 : 1 - diaSemana;
    const segunda    = new Date(hoje);
    segunda.setDate(hoje.getDate() + diffSeg);
    const semanaStr  = Utilities.formatDate(segunda, CONFIG.FUSO, 'dd/MM/yyyy');

    // Professores que já entregaram esta semana
    const jaEntregaram = new Set();
    if (aba) {
      aba.getDataRange().getValues().slice(1).forEach(row => {
        // Coluna 5 (índice 5) = Semana Início
        if (String(row[5]).trim() === semanaStr) jaEntregaram.add(row[1]);
      });
    }

    const pendentes = cadastros.professores.filter(p => !jaEntregaram.has(p));
    if (pendentes.length === 0) {
      log('INFO', 'lembreteSemanaPlanejamento', 'Todos entregaram o semanal desta semana — nenhum lembrete enviado.');
      return;
    }

    // Envia e-mail para cada pendente
    let enviados = 0;
    pendentes.forEach(professor => {
      const emailProf = cadastros.emails && cadastros.emails[professor];
      if (!emailProf) return;

      const assunto = `⏰ Lembrete: planejamento semanal não entregue — ${professor}`;
      const corpo = [
        `Olá, ${professor}!`,
        ``,
        `Este é um lembrete automático do sistema de planejamento escolar.`,
        ``,
        `Identificamos que o seu planejamento semanal da semana de ${semanaStr} ainda não foi enviado.`,
        ``,
        `📋 Acesse o sistema para preencher e salvar o planejamento:`,
        `   ${ScriptApp.getService().getUrl()}`,
        ``,
        `Se já enviou o planejamento por outro meio, desconsidere este e-mail.`,
        ``,
        `Colégio Municipal de 1º e 2º Graus de Itabatan`,
        `Coordenação Pedagógica — Sistema Automático`,
      ].join('\n');

      MailApp.sendEmail(emailProf, assunto, corpo);
      enviados++;
    });

    // [V9-5] Resumo para TODOS os e-mails de coordenação (não apenas o primeiro)
    const emailsCoord = CONFIG.EMAILS_COORDENACAO.filter(e => e && !e.startsWith('COLE_'));
    emailsCoord.forEach(emailCoord => {
      try {
        MailApp.sendEmail(
          emailCoord,
          `📊 Resumo semanal de pendências — ${semanaStr}`,
          `Semana de ${semanaStr}\n\nPendentes (${pendentes.length}):\n${pendentes.join('\n')}\n\nLembretes enviados: ${enviados} de ${pendentes.length}\n(Professores sem e-mail cadastrado não recebem lembrete.)`
        );
      } catch(eCoord) {
        log('AVISO', 'lembreteSemanaPlanejamento', `Falha ao notificar coordenação ${emailCoord}: ${eCoord.message}`);
      }
    });

    log('INFO', 'lembreteSemanaPlanejamento', `${pendentes.length} pendentes, ${enviados} lembretes enviados — semana ${semanaStr}`);
  } catch(e) {
    log('ERRO', 'lembreteSemanaPlanejamento', e.message);
  }
}

/**
 * Instala o trigger semanal automaticamente.
 * Execute esta função UMA VEZ manualmente no editor do Apps Script.
 */
function instalarTriggerSemanal() {
  // Remove triggers antigos da mesma função para evitar duplicatas
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'lembreteSemanaPlanejamento')
    .forEach(t => ScriptApp.deleteTrigger(t));

  // [V9-2] Garante que a aba de configurações existe antes de ler
  // (caso seja a primeira execução, _cfg() pode retornar o fallback
  //  corretamente graças ao segundo parâmetro)
  const diaStr = (_cfg('lembrete_dia_semana', 'FRIDAY') || 'FRIDAY').toUpperCase().trim();
  const horaStr = _cfg('lembrete_hora', '15') || '15';
  const hora   = Math.max(0, Math.min(23, parseInt(horaStr, 10) || 15));

  const diasMap = {
    MONDAY:    ScriptApp.WeekDay.MONDAY,
    TUESDAY:   ScriptApp.WeekDay.TUESDAY,
    WEDNESDAY: ScriptApp.WeekDay.WEDNESDAY,
    THURSDAY:  ScriptApp.WeekDay.THURSDAY,
    FRIDAY:    ScriptApp.WeekDay.FRIDAY,
    SATURDAY:  ScriptApp.WeekDay.SATURDAY,
    SUNDAY:    ScriptApp.WeekDay.SUNDAY,
    // Aliases em português para facilitar o uso
    SEGUNDA:   ScriptApp.WeekDay.MONDAY,
    TERCA:     ScriptApp.WeekDay.TUESDAY,
    QUARTA:    ScriptApp.WeekDay.WEDNESDAY,
    QUINTA:    ScriptApp.WeekDay.THURSDAY,
    SEXTA:     ScriptApp.WeekDay.FRIDAY,
    SABADO:    ScriptApp.WeekDay.SATURDAY,
    DOMINGO:   ScriptApp.WeekDay.SUNDAY,
  };
  const dia = diasMap[diaStr] || ScriptApp.WeekDay.FRIDAY;

  ScriptApp.newTrigger('lembreteSemanaPlanejamento')
    .timeBased()
    .onWeekDay(dia)
    .atHour(hora)
    .create();

  const msg = `Trigger instalado — ${diaStr} às ${hora}h (fuso ${CONFIG.FUSO})`;
  log('INFO', 'instalarTriggerSemanal', msg);
  Logger.log(`✅ ${msg}`);
}

// ═══════════════════════════════════════════════════════════
//  [F3-2] CARREGAR PLANEJAMENTO SALVO PARA EDIÇÃO
// ═══════════════════════════════════════════════════════════

/**
 * Busca os dados de um registro pelo índice de linha em sua aba,
 * para o professor preencher o formulário de edição.
 * Retorna objeto com tipo + dados suficientes para recarregar o form.
 */
function carregarParaEdicao(tipo, linhaIdx) {
  try {
    const ss  = abrirPlanilha();
    const aba = ss.getSheetByName(tipo);
    if (!aba) return { sucesso: false, erro: 'Aba não encontrada.' };

    const row = aba.getRange(linhaIdx, 1, 1, aba.getLastColumn()).getValues()[0];
    if (!row || !row[0]) return { sucesso: false, erro: 'Linha não encontrada.' };

    let dados = {};
    if (tipo === ABA.TRIMESTRAL) {
      dados = {
        professor: row[1], anoTurma: row[2], componente: row[3],
        trimestre: String(row[4]).replace('º',''), obs: '',
        linhas: [{ unidade:'', habilidades:'', metodologia:'', avaliacao:'' }],
        _linhaOrigem: linhaIdx,
        _urlOrigem:   row[5],
      };
    } else if (tipo === ABA.SEMANAL) {
      dados = {
        professor: row[1], turmas: row[2], componente: row[3],
        mes: row[4], semanaInicio: row[5], semanaFim: row[6], obs: '',
        semana: Array(5).fill(null).map(() => ({ habilidade:'', estrategia:'', recursos:'', observacao:'' })),
        datas: ['','','','',''],
        _linhaOrigem: linhaIdx,
        _urlOrigem:   row[7],
      };
    } else {
      dados = {
        professor: row[1], anoTurma: row[2], componente: row[3],
        objetivo: '', bncc: '', obs: '',
        meses: Array(12).fill(''),
        _linhaOrigem: linhaIdx,
        _urlOrigem:   row[4],
      };
    }

    log('INFO', 'carregarParaEdicao', `Carregando ${tipo} linha ${linhaIdx} para edição — ${dados.professor}`);
    return { sucesso: true, tipo, dados };
  } catch(e) {
    log('ERRO', 'carregarParaEdicao', e.message);
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Salva nova versão de um planejamento editado.
 * Adiciona uma nova linha na planilha (mantém histórico) e marca a antiga como "(editado)".
 */
function salvarEdicao(payload) {
  try {
    const { tipo, dados } = payload;
    let resultado;
    if      (tipo === ABA.TRIMESTRAL) resultado = salvarTrimestral(dados);
    else if (tipo === ABA.SEMANAL)    resultado = salvarSemanal(dados);
    else                              resultado = salvarAnual(dados);

    if (!resultado.sucesso) return resultado;

    // Marca linha original como editada (coluna de status se existir)
    if (dados._linhaOrigem) {
      try {
        const ss  = abrirPlanilha();
        const aba = ss.getSheetByName(tipo);
        if (aba) {
          const ultimaCol = aba.getLastColumn();
          aba.getRange(dados._linhaOrigem, ultimaCol, 1, 1)
             .setValue(`[Substituído em ${ts('dd/MM/yyyy HH:mm')}]`);
        }
      } catch(e2) { /* não crítico */ }
    }

    log('INFO', 'salvarEdicao', `Nova versão salva para ${dados.professor} — ${tipo}`);
    return resultado;
  } catch(e) {
    log('ERRO', 'salvarEdicao', e.message);
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Retorna todos os registros de um professor com número de linha,
 * para o histórico oferecer botão "Editar".
 */
function getHistoricoComLinhas(professor) {
  try {
    const ss  = abrirPlanilha();
    const res = [];
    [ABA.TRIMESTRAL, ABA.SEMANAL, ABA.ANUAL].forEach(tipo => {
      const aba = ss.getSheetByName(tipo);
      if (!aba) return;
      const vals = aba.getDataRange().getValues();
      vals.slice(1).forEach((row, i) => {
        if (!row[1] || row[1] !== professor) return;
        const linhaIdx = i + 2; // +1 header, +1 base-1
        const r = { tipo, data: row[0], professor: row[1], linhaIdx };
        if (tipo === ABA.TRIMESTRAL) Object.assign(r, { anoTurma:row[2], componente:row[3], trimestre:row[4], url:row[5] });
        else if (tipo === ABA.SEMANAL) Object.assign(r, { turmas:row[2], componente:row[3], mes:row[4], semanaInicio:row[5], url:row[7] });
        else Object.assign(r, { anoTurma:row[2], componente:row[3], url:row[4] });
        res.push(r);
      });
    });
    res.sort((a, b) => (b.data > a.data ? 1 : -1));
    return res;
  } catch(e) {
    log('ERRO', 'getHistoricoComLinhas', e.message);
    return [];
  }
}

// ═══════════════════════════════════════════════════════════
//  [F3-1] FLUXO DE APROVAÇÃO
// ═══════════════════════════════════════════════════════════

/**
 * Salva uma decisão de aprovação na aba "Aprovacoes".
 * payload: { tipo, linhaIdx, professor, componente, periodo, decisao, comentario, urlPDF }
 * decisao: "APROVADO" | "REVISAO"
 */
function salvarAprovacao(payload) {
  // [V9-19] Aprovação é exclusiva da coordenação — verifica explicitamente
  const sess = _verificarPermissao('salvarAprovacao');
  if (sess.perfil !== 'coordenacao') {
    return { sucesso: false, erro: 'Apenas a coordenação pode aprovar ou solicitar revisão de planejamentos.' };
  }
  try {
    const { tipo, linhaIdx, professor, componente, periodo, decisao, comentario, urlPDF } = payload;
    const ss  = abrirPlanilha();
    let aba   = ss.getSheetByName(ABA.APROVACOES);

    if (!aba) {
      aba = ss.insertSheet(ABA.APROVACOES);
      aba.appendRow(["Data/Hora","Tipo","Professor(a)","Componente","Período","Decisão","Comentário","Link PDF","Aprovado por"]);
      aba.getRange(1,1,1,9).setFontWeight("bold").setBackground("#1a3a6b").setFontColor("white");
      aba.setFrozenRows(1);
    }

    const agora   = ts('dd/MM/yyyy HH:mm');
    const revisor = Session.getActiveUser().getEmail() || CONFIG.EMAIL_COORDENACAO;
    aba.appendRow([agora, tipo, professor, componente, periodo, decisao, comentario || '', urlPDF || '', revisor]);

    // Coloriza a linha conforme decisão
    const linha = aba.getLastRow();
    const cor   = decisao === 'APROVADO' ? '#f0fdf4' : '#fef9c3';
    aba.getRange(linha, 1, 1, 9).setBackground(cor);
    aba.getRange(linha, 6).setFontWeight('bold')
       .setFontColor(decisao === 'APROVADO' ? '#166534' : '#713f12');

    // Notifica o professor por e-mail
    _notificarDecisao(professor, componente, periodo, decisao, comentario);

    // Invalida cache do painel
    _invalidarCachePainel();

    log('INFO', 'salvarAprovacao', `${decisao} — ${professor} / ${componente} — ${revisor}`);
    return { sucesso: true };
  } catch(e) {
    log('ERRO', 'salvarAprovacao', e.message);
    return { sucesso: false, erro: e.message };
  }
}

function _notificarDecisao(professor, componente, periodo, decisao, comentario) {
  try {
    const cadastros = getCadastros();
    const emailProf = cadastros.emails && cadastros.emails[professor];
    if (!emailProf) return;

    const emoji  = decisao === 'APROVADO' ? '✅' : '🔄';
    const status = decisao === 'APROVADO' ? 'APROVADO' : 'AGUARDA REVISÃO';
    const corpo  = [
      `Olá, ${professor}!`,
      ``,
      `Seu planejamento de ${componente} — ${periodo} foi analisado pela coordenação.`,
      ``,
      `${emoji} Status: ${status}`,
      comentario ? `💬 Comentário da coordenação:\n   "${comentario}"` : '',
      ``,
      `Acesse o sistema para visualizar detalhes ou fazer ajustes:`,
      `   ${ScriptApp.getService().getUrl()}`,
      ``,
      `Colégio Municipal de 1º e 2º Graus de Itabatan — Coordenação Pedagógica`,
    ].filter(Boolean).join('\n');

    MailApp.sendEmail(emailProf, `${emoji} Planejamento ${status} — ${componente}`, corpo);
  } catch(e) {
    log('AVISO', '_notificarDecisao', `Falha ao notificar ${professor}: ${e.message}`);
  }
}

/**
 * Retorna decisões de aprovação para um professor específico ou todos (coordenação).
 */
function getAprovacoes(filtro) {
  try {
    const ss  = abrirPlanilha();
    const aba = ss.getSheetByName(ABA.APROVACOES);
    if (!aba) return [];

    const vals = aba.getDataRange().getValues();
    return vals.slice(1)
      .filter(row => !filtro || row[2] === filtro)
      .map(row => ({
        data: row[0], tipo: row[1], professor: row[2], componente: row[3],
        periodo: row[4], decisao: row[5], comentario: row[6],
        url: row[7], revisor: row[8],
      }))
      .sort((a, b) => (b.data > a.data ? 1 : -1));
  } catch(e) {
    log('ERRO', 'getAprovacoes', e.message);
    return [];
  }
}

// ═══════════════════════════════════════════════════════════
//  [F3-4] ESTATÍSTICAS PESSOAIS DO PROFESSOR
// ═══════════════════════════════════════════════════════════

function getEstatisticasProfessor(professor) {
  try {
    const ss   = abrirPlanilha();
    let totalTrimestral = 0, totalSemanal = 0, totalAnual = 0;
    let aprovados = 0, revisoes = 0;
    const semanasPorMes = {};
    let sequenciaAtual = 0, maiorSequencia = 0;

    // Conta por tipo
    [ABA.TRIMESTRAL, ABA.SEMANAL, ABA.ANUAL].forEach(tipo => {
      const aba = ss.getSheetByName(tipo);
      if (!aba) return;
      aba.getDataRange().getValues().slice(1).forEach(row => {
        if (row[1] !== professor) return;
        if (tipo === ABA.TRIMESTRAL) totalTrimestral++;
        else if (tipo === ABA.ANUAL) totalAnual++;
        else {
          totalSemanal++;
          // Agrupa semanais por mês para calcular sequência
          const mes = String(row[4]).trim();
          semanasPorMes[mes] = (semanasPorMes[mes] || 0) + 1;
        }
      });
    });

    // Conta aprovações
    const abaAprov = ss.getSheetByName(ABA.APROVACOES);
    if (abaAprov) {
      abaAprov.getDataRange().getValues().slice(1).forEach(row => {
        if (row[2] !== professor) return;
        if (row[5] === 'APROVADO') aprovados++;
        else if (row[5] === 'REVISAO') revisoes++;
      });
    }

    // Calcula sequência de semanas consecutivas (simplificado: semanas no mês corrente e anterior)
    sequenciaAtual  = semanasPorMes[_mesLabel(0)] || 0;
    maiorSequencia  = Math.max(...Object.values(semanasPorMes), 0);

    const taxaPontualidade = totalSemanal > 0
      ? Math.round((aprovados / Math.max(totalSemanal, 1)) * 100)
      : null;

    return {
      sucesso: true,
      totalTrimestral, totalSemanal, totalAnual,
      aprovados, revisoes,
      sequenciaAtual, maiorSequencia,
      taxaPontualidade,
      semanaisPorMes: semanasPorMes,
    };
  } catch(e) {
    log('ERRO', 'getEstatisticasProfessor', e.message);
    return { sucesso: false, erro: e.message };
  }
}

function _mesLabel(delta) {
  const d = new Date();
  d.setMonth(d.getMonth() + delta);
  const meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
                 'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  return meses[d.getMonth()];
}

// ═══════════════════════════════════════════════════════════
//  [F3-3] DADOS DOS GRÁFICOS DO PAINEL
// ═══════════════════════════════════════════════════════════

/**
 * Retorna dados agregados para os gráficos da coordenação.
 * Encapsulado separado do getDadosPainel para não inflar o cache principal.
 */
function getDadosGraficos() {
  try {
    const ss  = abrirPlanilha();
    const meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];

    // Entregas por mês (todos os tipos)
    const porMes = Array(12).fill(0);
    [ABA.TRIMESTRAL, ABA.SEMANAL, ABA.ANUAL].forEach(tipo => {
      const aba = ss.getSheetByName(tipo);
      if (!aba) return;
      aba.getDataRange().getValues().slice(1).forEach(row => {
        if (!row[0]) return;
        const dataStr = String(row[0]);
        // Extrai mês do formato dd/MM/yyyy
        const m = parseInt(dataStr.substring(3,5), 10);
        if (m >= 1 && m <= 12) porMes[m-1]++;
      });
    });

    // Entregas por tipo
    const porTipo = {
      Trimestral: 0,
      Semanal:    0,
      Anual:      0,
    };
    [ABA.TRIMESTRAL, ABA.SEMANAL, ABA.ANUAL].forEach(tipo => {
      const aba = ss.getSheetByName(tipo);
      if (!aba) return;
      const count = Math.max(0, aba.getLastRow() - 1);
      porTipo[tipo] = count;
    });

    // Aprovações vs pendentes
    const cadastros   = getCadastros();
    const comEntrega  = new Set();
    [ABA.TRIMESTRAL, ABA.SEMANAL, ABA.ANUAL].forEach(tipo => {
      const aba = ss.getSheetByName(tipo);
      if (!aba) return;
      aba.getDataRange().getValues().slice(1).forEach(row => { if (row[1]) comEntrega.add(row[1]); });
    });
    const totalProf  = cadastros.professores.length;
    const comPlano   = comEntrega.size;
    const semPlano   = Math.max(0, totalProf - comPlano);

    // Top 5 professores mais pontuais (mais entregas)
    const contProf = {};
    [ABA.TRIMESTRAL, ABA.SEMANAL, ABA.ANUAL].forEach(tipo => {
      const aba = ss.getSheetByName(tipo);
      if (!aba) return;
      aba.getDataRange().getValues().slice(1).forEach(row => {
        if (!row[1]) return;
        contProf[row[1]] = (contProf[row[1]] || 0) + 1;
      });
    });
    const top5 = Object.entries(contProf)
      .sort((a,b) => b[1]-a[1])
      .slice(0,5)
      .map(([nome, total]) => ({ nome: nome.split(' ')[0], total }));

    return {
      sucesso:   true,
      meses,
      porMes,
      porTipo,
      comPlano,
      semPlano,
      top5,
    };
  } catch(e) {
    log('ERRO', 'getDadosGraficos', e.message);
    return { sucesso: false, erro: e.message };
  }
}

// ═══════════════════════════════════════════════════════════
//  [V8-4] ATRIBUIÇÕES DE TURMA/COMPONENTE POR PROFESSOR
// ═══════════════════════════════════════════════════════════

/**
 * Retorna turmas e componentes atribuídos a um professor específico.
 * Aba "TurmasProfessor" — colunas: Professor | Componente | Turma
 * Se não existir a aba ou o professor não tiver atribuições,
 * devolve todas as turmas/componentes disponíveis (fallback).
 */
function _getAtribuicoesProfessor(nomeProfessor) {
  try {
    const ss  = abrirPlanilha();
    const aba = ss.getSheetByName(ABA.TURMAS_PROFESSOR);
    if (!aba) return { turmas: [], componentes: [] };

    const vals = aba.getDataRange().getValues();
    const turmas      = new Set();
    const componentes = new Set();

    vals.slice(1).forEach(row => {
      if (String(row[0]).trim() !== nomeProfessor) return;
      if (row[1]) componentes.add(String(row[1]).trim());
      if (row[2]) turmas.add(String(row[2]).trim());
    });

    return {
      turmas:      [...turmas],
      componentes: [...componentes],
    };
  } catch(e) {
    return { turmas: [], componentes: [] };
  }
}

// ═══════════════════════════════════════════════════════════
//  [V8-2] CONFIGURAÇÕES DO SISTEMA
// ═══════════════════════════════════════════════════════════

/**
 * Carrega todas as configurações da planilha.
 * Retorna objeto com todos os dados editáveis pela coordenação.
 */
function getConfiguracoes() {
  _verificarPermissao('getConfiguracoes');
  try {
    const ss  = abrirPlanilha();
    let aba   = ss.getSheetByName(ABA.CONFIGURACOES);
    if (!aba) aba = _criarAbaConfiguracoes(ss);

    const vals = aba.getDataRange().getValues();
    const cfg  = {};
    vals.slice(1).forEach(row => {
      if (!row[0]) return;
      cfg[String(row[0]).trim()] = String(row[1] ?? '').trim();
    });
    return { sucesso: true, cfg };
  } catch(e) {
    log('ERRO', 'getConfiguracoes', e.message);
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Salva uma ou mais configurações de volta na planilha.
 * payload: { chave: valor, chave2: valor2, ... }
 */
function salvarConfiguracoes(payload) {
  _verificarPermissao('salvarConfiguracoes');
  try {
    const ss  = abrirPlanilha();
    let aba   = ss.getSheetByName(ABA.CONFIGURACOES);
    if (!aba) aba = _criarAbaConfiguracoes(ss);

    // [V9-1] BUG CORRIGIDO: re-lê a planilha a cada iteração para garantir
    // índices corretos após escritas anteriores no mesmo lote.
    // O bug original usava um único `vals` estático; se duas chaves
    // distintas fossem salvas no mesmo payload, a segunda poderia gravar
    // na linha errada porque `i + 2` estava deslocado pelo header não skipado.
    Object.entries(payload).forEach(([chave, valor]) => {
      // Re-lê para ter o índice atualizado a cada iteração
      const valsAtual = aba.getDataRange().getValues();
      let linhaEncontrada = -1;

      // slice(1) para pular o cabeçalho; o índice real na planilha = i + 2
      valsAtual.slice(1).forEach((row, i) => {
        if (String(row[0]).trim() === chave) {
          linhaEncontrada = i + 2; // +1 do slice(1) + +1 do 1-based do Sheets
        }
      });

      if (linhaEncontrada > 0) {
        aba.getRange(linhaEncontrada, 2).setValue(valor);
      } else {
        aba.appendRow([chave, valor]);
      }

      // [V9-4] Invalida cache individual desta chave para que PDFs gerados
      // após salvar usem o valor atualizado imediatamente
      try {
        CacheService.getScriptCache().remove('cfg_' + chave);
      } catch(e2) { /* não crítico */ }
    });

    log('INFO', 'salvarConfiguracoes', `Configurações atualizadas: ${Object.keys(payload).join(', ')}`);
    return { sucesso: true };
  } catch(e) {
    log('ERRO', 'salvarConfiguracoes', e.message);
    return { sucesso: false, erro: e.message };
  }
}

/**
 * Salva alterações em lote na aba Cadastros (professores, turmas, componentes).
 * payload: { professores: [...], turmas: [...], componentes: [...], emails: {...} }
 */
function salvarCadastros(payload) {
  _verificarPermissao('salvarCadastros');
  try {
    const ss  = abrirPlanilha();
    let aba   = ss.getSheetByName(ABA.CADASTROS);
    if (!aba) aba = _criarAbaCadastros(ss);

    // Reconstrói a aba inteiramente a partir dos dados recebidos
    const linhas = [["Tipo", "Nome/Valor", "E-mail (só Professor)"]];

    // [V9-13] Normaliza espaços para evitar entradas duplicadas por espaços acidentais (corrige QUALITY-3)
    (payload.professores || []).map(n => String(n).trim()).filter(Boolean).forEach(nome => {
      const email = String((payload.emails || {})[nome] || '').trim();
      linhas.push(["PROFESSOR", nome, email]);
    });
    (payload.componentes || []).map(c => String(c).trim()).filter(Boolean).forEach(comp => linhas.push(["COMPONENTE", comp, ""]));
    (payload.turmas      || []).map(t => String(t).trim()).filter(Boolean).forEach(turm => linhas.push(["TURMA",      turm, ""]));

    aba.clearContents();
    aba.getRange(1, 1, linhas.length, 3).setValues(linhas);
    aba.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#1a3a6b").setFontColor("white");
    aba.autoResizeColumns(1, 3);

    // Atualiza atribuições de turma/componente por professor
    if (payload.atribuicoes) _salvarAtribuicoes(ss, payload.atribuicoes);

    // Invalida cache do painel
    _invalidarCachePainel();
    log('INFO', 'salvarCadastros', `Cadastros atualizados — ${(payload.professores||[]).length} professores`);
    return { sucesso: true };
  } catch(e) {
    log('ERRO', 'salvarCadastros', e.message);
    return { sucesso: false, erro: e.message };
  }
}

function _salvarAtribuicoes(ss, atribuicoes) {
  let aba = ss.getSheetByName(ABA.TURMAS_PROFESSOR);
  if (!aba) {
    aba = ss.insertSheet(ABA.TURMAS_PROFESSOR);
    aba.appendRow(["Professor(a)", "Componente", "Turma"]);
    aba.getRange(1,1,1,3).setFontWeight("bold").setBackground("#1a3a6b").setFontColor("white");
  }

  // Mantém apenas a linha de header e re-escreve tudo
  const ultimaLinha = aba.getLastRow();
  if (ultimaLinha > 1) aba.deleteRows(2, ultimaLinha - 1);

  // atribuicoes: [{ professor, componente, turma }, ...]
  atribuicoes.forEach(a => {
    if (a.professor && a.componente && a.turma) {
      aba.appendRow([a.professor, a.componente, a.turma]);
    }
  });
}

function _criarAbaConfiguracoes(ss) {
  const aba = ss.insertSheet(ABA.CONFIGURACOES);
  aba.appendRow(["Configuração", "Valor", "Descrição"]);
  aba.getRange(1,1,1,3).setFontWeight("bold").setBackground("#1a3a6b").setFontColor("white");

  const cfgs = [
    ["escola_nome",          "Colégio Municipal de 1º e 2º Graus de Itabatan", "Nome completo da escola"],
    ["escola_endereco",      "Av. Itapetinga, 305 - Gazzinelândia - Itabatan - Mucuri/BA", "Endereço da escola"],
    ["escola_cep",           "45.936-000",       "CEP"],
    ["escola_criacao",       "Criação 062/94 - Aut. CEE 236/95 DO 25/01/1996", "Portaria de criação"],
    ["ano_letivo",           String(new Date().getFullYear()), "Ano letivo atual"],
    ["trimestres",           "3",                "Número de trimestres (2 ou 3)"],
    ["avaliacao_geral",      "12,0",             "Pontuação da avaliação geral"],
    ["avaliacao_parcial",    "18,0",             "Pontuação total das avaliações parciais"],
    ["dias_letivos_anuais",  "200",              "Total de dias letivos no ano"],
    ["hora_aula_minutos",    "50",               "Duração da aula em minutos"],
    ["lembrete_dia_semana",  "FRIDAY",           "Dia do lembrete semanal (MONDAY..SUNDAY)"],
    ["lembrete_hora",        "15",               "Hora do lembrete semanal (0-23)"],
    ["coordenacao_nome",     "Coordenação Pedagógica", "Nome que aparece nos documentos"],
  ];

  aba.getRange(2, 1, cfgs.length, 3).setValues(cfgs);
  aba.setColumnWidth(2, 360);
  aba.setColumnWidth(3, 280);

  return aba;
}

/**
 * Retorna configuração individual ou todas como objeto.
 * Função pública (sem guard) para uso nos geradores de PDF.
 */
function _cfg(chave, fallback) {
  try {
    // Validação de entrada
    if (!chave || typeof chave !== 'string') return String(fallback || '');
    
    const cache = CacheService.getScriptCache();
    const cKey  = 'cfg_' + chave;
    const hit   = cache.get(cKey);
    if (hit !== null) return hit;

    const ss  = abrirPlanilha();
    const aba = ss.getSheetByName(ABA.CONFIGURACOES);
    if (!aba || aba.getLastRow() < 2) return String(fallback || '');

    const vals = aba.getDataRange().getValues();
    for (const row of vals.slice(1)) {
      if (!row || row.length < 2) continue;
      if (String(row[0] || '').trim() === chave) {
        const val = String(row[1] ?? '').trim();
        // Cache por 5 minutos
        try { cache.put(cKey, val, 300); } catch(eCa) { /* ignora erro de cache */ }
        return val || String(fallback || '');
      }
    }
    return String(fallback || '');
  } catch(e) {
    log('AVISO', '_cfg', `Erro ao buscar configuração '${chave}': ${e.message}`);
    return String(fallback || '');
  }
}

// ═══════════════════════════════════════════════════════════
//  [V8-3] BNCC OFICIAL — API pública do MEC (bncc.ninja)
// ═══════════════════════════════════════════════════════════

/**
 * Busca habilidades BNCC pela API pública bncc.ninja (dados oficiais do MEC).
 * Com cache local na planilha para evitar chamadas repetidas.
 *
 * A API retorna habilidades filtradas por:
 *   - código (ex: EF07CI01)
 *   - texto da descrição (palavra-chave)
 *
 * Endpoint: https://bncc.ninja/api/habilidades?q=QUERY&limit=10
 * (API pública, sem autenticação, mantida por educadores brasileiros)
 *
 * Fallback: se a API estiver indisponível, busca na aba BNCC local.
 */
function buscarHabilidadesBNCC(query) {
  if (!query || query.trim().length < 2) return [];

  try {
    // 1. Tenta a API oficial primeiro
    const resultado = _buscarBNCC_API(query.trim());
    if (resultado && resultado.length > 0) return resultado;

    // 2. Fallback: busca local na planilha
    return _buscarBNCC_local(query.trim());
  } catch(e) {
    log('AVISO', 'buscarHabilidadesBNCC', `API indisponível, usando local: ${e.message}`);
    return _buscarBNCC_local(query.trim());
  }
}

function _buscarBNCC_API(query) {
  // Cache de 10 minutos por query para não sobrecarregar a API
  // [V9-11] Usa hash da query completa para evitar colisões em queries longas (corrige PERF-6)
  const cacheKey = 'bncc_' + Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    query.toLowerCase()
  ).map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2,'0')).join('');
  const cache    = CacheService.getScriptCache();
  const hit      = cache.get(cacheKey);
  if (hit) {
    try { return JSON.parse(hit); } catch(e) { /* continua */ }
  }

  // Tenta a API bncc.ninja (dados oficiais MEC)
  const url = 'https://bncc.ninja/api/habilidades?q=' + encodeURIComponent(query) + '&limit=15';
  let resp;
  try {
    resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'Accept': 'application/json' },
      deadline: 10,
    });
  } catch(eFetch) {
    // DNS error ou timeout — lança para o caller usar fallback local
    throw new Error('API bncc.ninja indisponivel (DNS/timeout): ' + eFetch.message);
  }

  if (resp.getResponseCode() !== 200) throw new Error('API retornou ' + resp.getResponseCode());

  const data = JSON.parse(resp.getContentText());
  const itens = (data.habilidades || data.data || data || []).slice(0, 15);

  // [V9-12] Ignora itens null/não-objeto antes de mapear (corrige BUG-9)
  const resultado = itens.filter(h => h && typeof h === 'object').map(h => ({
    codigo:     h.codigo    || h.code    || '',
    componente: h.componente || h.area   || h.component || '',
    segmento:   h.ano       || h.segment || h.serie     || '',
    descricao:  h.descricao || h.description || h.text  || '',
    fonte:      'API BNCC oficial',
  })).filter(h => h.codigo && h.descricao);

  // Salva no cache
  try { cache.put(cacheKey, JSON.stringify(resultado), 600); } catch(e) { /* ignora */ }

  // Salva novos resultados na aba BNCC local para uso offline futuro
  _salvarHabilidadesLocal(resultado);

  return resultado;
}

function _buscarBNCC_local(query) {
  try {
    const ss   = abrirPlanilha();
    let aba    = ss.getSheetByName(ABA.BNCC);
    if (!aba) aba = _criarAbaBNCC(ss);

    const q    = query.toLowerCase();
    const vals = aba.getDataRange().getValues();

    return vals.slice(1)
      .filter(row => {
        const cod  = String(row[0]).toLowerCase();
        const comp = String(row[1]).toLowerCase();
        const desc = String(row[3]).toLowerCase();
        return cod.includes(q) || comp.includes(q) || desc.includes(q);
      })
      .slice(0, 15)
      .map(row => ({
        codigo:     String(row[0]).trim(),
        componente: String(row[1]).trim(),
        segmento:   String(row[2]).trim(),
        descricao:  String(row[3]).trim(),
        fonte:      'cache local',
      }));
  } catch(e) {
    return [];
  }
}

/**
 * Salva habilidades retornadas pela API na aba BNCC local (evita duplicatas).
 */
function _salvarHabilidadesLocal(habilidades) {
  if (!habilidades || habilidades.length === 0) return;
  try {
    const ss  = abrirPlanilha();
    let aba   = ss.getSheetByName(ABA.BNCC);
    if (!aba) aba = _criarAbaBNCC(ss);

    // Lê códigos já existentes para evitar duplicata
    const existentes = new Set(
      aba.getDataRange().getValues().slice(1).map(r => String(r[0]).trim())
    );

    habilidades.forEach(h => {
      if (!h.codigo || existentes.has(h.codigo)) return;
      aba.appendRow([h.codigo, h.componente, h.segmento, h.descricao]);
      existentes.add(h.codigo);
    });
  } catch(e) { /* não crítico — operação em background */ }
}

/**
 * Importa habilidades em lote pela API para um componente/segmento inteiro.
 * Execute manualmente para pré-popular o banco local.
 * componente: "LP", "MA", "CI", "HI", "GE", "AR", "EF", "LI", "ER"
 * segmento:   "EF" (Ensino Fundamental), "EM" (Ensino Médio)
 *
 * Se chamada sem parâmetros (ex: do editor), importa todos os componentes padrão.
 * Se a API externa estiver indisponível, registra aviso e retorna dados locais existentes.
 */
function importarHabilidadesBNCC(componente, segmento) {
  _verificarPermissao('importarHabilidadesBNCC');

  // [FIX] Se chamada sem parâmetros do editor, importa todos os componentes padrão
  if (!componente || String(componente) === 'undefined') {
    Logger.log('importarHabilidadesBNCC: nenhum componente informado — importando todos os padrão.');
    const comps = ['LP', 'MA', 'CI', 'HI', 'GE', 'AR', 'EF', 'LI', 'ER'];
    let totalImportado = 0;
    const erros = [];
    comps.forEach(comp => {
      const res = importarHabilidadesBNCC(comp, segmento || 'EF');
      if (res.sucesso) totalImportado += (res.importadas || 0);
      else if (!res.apiIndisponivel) erros.push(comp + ': ' + res.erro);
    });
    const msg = 'Importacao em lote: ' + totalImportado + ' habilidade(s).' +
                (erros.length ? ' Erros: ' + erros.join('; ') : '');
    log('INFO', 'importarHabilidadesBNCC', msg);
    Logger.log(msg);
    return { sucesso: true, importadas: totalImportado, msg };
  }

  const seg = segmento || 'EF';
  try {
    const url = 'https://bncc.ninja/api/habilidades?componente=' + encodeURIComponent(componente) +
                '&segmento=' + encodeURIComponent(seg) + '&limit=200';
    let resp;
    try {
      resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, deadline: 15 });
    } catch(eFetch) {
      // DNS error ou timeout — API indisponível
      const msg = 'API bncc.ninja indisponivel (' + eFetch.message + '). ' +
                  'Os dados locais continuam disponiveis para busca offline.';
      log('AVISO', 'importarHabilidadesBNCC', msg);
      Logger.log('AVISO: ' + msg);
      return { sucesso: false, importadas: 0, erro: msg, apiIndisponivel: true };
    }

    if (resp.getResponseCode() !== 200) {
      throw new Error('HTTP ' + resp.getResponseCode());
    }

    const data = JSON.parse(resp.getContentText());
    const itens = (data.habilidades || data.data || data || []);
    const resultado = itens
      .filter(h => h && typeof h === 'object')
      .map(h => ({
        codigo:     h.codigo     || h.code        || '',
        componente: h.componente || h.area        || '',
        segmento:   h.ano        || h.serie       || '',
        descricao:  h.descricao  || h.description || h.text || '',
      }))
      .filter(h => h.codigo && h.descricao);

    _salvarHabilidadesLocal(resultado);
    const msg = 'Importadas ' + resultado.length + ' habilidades de ' + componente + '/' + seg;
    log('INFO', 'importarHabilidadesBNCC', msg);
    Logger.log(msg);
    return { sucesso: true, importadas: resultado.length, msg };
  } catch(e) {
    log('ERRO', 'importarHabilidadesBNCC', e.message);
    Logger.log('ERRO importarHabilidadesBNCC (' + componente + '): ' + e.message);
    return { sucesso: false, erro: e.message };
  }
}

// ═══════════════════════════════════════════════════════════
//  UTILITÁRIOS
// ═══════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════
//  GERENCIAMENTO DE ID DA PLANILHA
// ═══════════════════════════════════════════════════════════

/**
 * [V9-20] Obtém o ID da planilha de registros.
 * Prioridade: 1) CONFIG.PLANILHA_ID, 2) Propriedades do Script, 3) Busca/Criação
 */
function _obterIdPlanilha() {
  // 1. Verifica CONFIG primeiro
  if (CONFIG.PLANILHA_ID && CONFIG.PLANILHA_ID !== "COLE_AQUI_O_ID_DA_PLANILHA_REGISTRO") {
    return CONFIG.PLANILHA_ID;
  }
  
  // 2. Verifica propriedades do script (salvo automaticamente)
  try {
    const props = PropertiesService.getScriptProperties();
    const idSalvo = props.getProperty('PLANILHA_REGISTRO_ID');
    if (idSalvo) {
      Logger.log(`✓ ID da planilha recuperado das propriedades: ${idSalvo}`);
      return idSalvo;
    }
  } catch(e) {
    Logger.log(`Aviso: Não foi possível acessar propriedades do script: ${e.message}`);
  }
  
  return null;
}

/**
 * [V9-20] Salva o ID da planilha nas propriedades do script
 */
function _salvarIdPlanilha(id) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('PLANILHA_REGISTRO_ID', id);
    Logger.log(`✓ ID da planilha salvo automaticamente: ${id}`);
    Logger.log(`💡 OPCIONAL: Para melhor performance, defina CONFIG.PLANILHA_ID = "${id}" no código.`);
  } catch(e) {
    Logger.log(`Aviso: Não foi possível salvar ID nas propriedades: ${e.message}`);
  }
}

// [BUG-2][V9-20] Planilha aberta sempre por ID — criação automática no primeiro acesso
function abrirPlanilha() {
  // NOTA: usa Logger.log (NÃO log()) para evitar recursão infinita
  // log() chama abrirPlanilha() → não usar log() aqui
  function _L(msg) { try { Logger.log('[abrirPlanilha] ' + msg); } catch(e) {} }
  try {
    // 1. Tenta abrir por ID (CONFIG ou propriedades salvas)
    const idExistente = _obterIdPlanilha();
    if (idExistente) {
      try {
        const ss = SpreadsheetApp.openById(idExistente);
        _L('Planilha aberta por ID: ' + ss.getName());
        return ss;
      } catch(e) {
        _L('ID salvo inválido (' + idExistente + '): ' + e.message);
        // ID inválido, continua para buscar/criar
      }
    }

    // 2. Busca planilha existente pelo nome na pasta
    let pasta;
    try {
      pasta = DriveApp.getFolderById(CONFIG.PASTA_ID);
    } catch(e) {
      throw new Error('Não foi possível acessar a pasta do Drive (ID: ' + CONFIG.PASTA_ID + '). Verifique se o ID está correto e se você tem permissão. Erro: ' + e.message);
    }

    const nomeArquivo = 'Registro_Planejamentos_' + CONFIG.ANO_LETIVO;
    const arqs = pasta.getFilesByName(nomeArquivo);

    if (arqs.hasNext()) {
      const arq = arqs.next();
      const id = arq.getId();
      _salvarIdPlanilha(id);
      const ss = SpreadsheetApp.openById(id);
      _L('Planilha encontrada: ' + nomeArquivo + ' (' + id + ')');
      return ss;
    }

    // 3. Cria nova planilha (primeiro acesso)
    _L('Criando nova planilha: ' + nomeArquivo);
    let ss;
    try {
      ss = SpreadsheetApp.create(nomeArquivo);
    } catch(e) {
      throw new Error('Não foi possível criar a planilha. Erro: ' + e.message);
    }

    const arqId = ss.getId();

    // Move para a pasta correta
    try {
      const arqDrive = DriveApp.getFileById(arqId);
      pasta.addFile(arqDrive);
      DriveApp.getRootFolder().removeFile(arqDrive);
    } catch(e) {
      _L('Aviso ao mover arquivo: ' + e.message);
    }

    _salvarIdPlanilha(arqId);
    _L('Planilha criada com sucesso. ID: ' + arqId);
    return ss;
  } catch(e) {
    if (e.message && e.message.includes('Não foi possível')) throw e;
    throw new Error('Erro ao acessar planilha de registros: ' + e.message);
  }
}

function toPDF(html, nome) {
  try {
    // Validação de entrada
    if (!html || typeof html !== 'string') {
      throw new Error('Conteúdo HTML inválido');
    }
    if (!nome || typeof nome !== 'string') {
      nome = `documento_${Utilities.formatDate(new Date(), CONFIG.FUSO, 'yyyyMMdd_HHmmss')}`;
    }
    
    // Sanitiza nome do arquivo
    const nomeLimpo = String(nome).replace(/[^a-zA-Z0-9_\-]/g, '_').substring(0, 100);
    
    const blob = Utilities.newBlob(html, "text/html", nomeLimpo + ".html");
    const tmp  = DriveApp.createFile(blob);
    
    let pdf;
    try {
      pdf = tmp.getAs("application/pdf");
      pdf.setName(nomeLimpo + ".pdf");
    } finally {
      // Garante que o arquivo temporário seja sempre removido
      try { tmp.setTrashed(true); } catch(e) { /* ignora */ }
    }
    
    return pdf;
  } catch(e) {
    log('ERRO', 'toPDF', `Erro ao gerar PDF: ${e.message}`);
    throw new Error(`Falha ao converter HTML para PDF: ${e.message}`);
  }
}

function subpasta(raizId, nome) {
  const raiz   = DriveApp.getFolderById(raizId);
  const pastas = raiz.getFoldersByName(nome);
  return pastas.hasNext() ? pastas.next() : raiz.createFolder(nome);
}

function nomearArquivo(prefixo, dados) {
  const lim = s => {
    if (!s) return 'sem_nome';
    return String(s)
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '') // Remove acentos
      .replace(/[^a-zA-Z0-9\s]/g, '')   // Remove caracteres especiais
      .replace(/\s+/g, '_')             // Substitui espaços por underscore
      .substring(0, 25)                 // Limita tamanho
      .replace(/^_+|_+$/g, '')          // Remove underscores do início/fim
      || 'sem_nome';
  };
  
  const t = ts('yyyyMMdd_HHmm');
  
  try {
    if (prefixo === 'T') {
      return `T_${lim(dados.professor)}_${lim(dados.componente)}_${lim(dados.anoTurma)}_${dados.trimestre || 'X'}_${t}`;
    }
    if (prefixo === 'S') {
      return `S_${lim(dados.professor)}_${lim(dados.componente)}_${lim(dados.mes)}_${t}`;
    }
    if (prefixo === 'A') {
      return `A_${lim(dados.professor)}_${lim(dados.componente)}_${t}`;
    }
    return `${prefixo}_${t}`;
  } catch(e) {
    log('AVISO', 'nomearArquivo', `Erro ao gerar nome: ${e.message}`);
    return `${prefixo}_${t}_erro`;
  }
}

// [BUG-5] Usa CONFIG.ANO_LETIVO, não string fixa
function ts(fmt) {
  return Utilities.formatDate(new Date(), CONFIG.FUSO, fmt);
}

function esc(str) {
  if (str === null || str === undefined) return '';
  // Converte para string e trata todos os casos especiais
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function nl2br(str) { 
  if (str === null || str === undefined) return '';
  return String(str).replace(/\r\n/g,'<br>').replace(/\n/g,'<br>').replace(/\r/g,'<br>'); 
}

// ═══════════════════════════════════════════════════════════
//  [V9-16] HEALTH CHECK — testa todos os serviços críticos
//  Execute manualmente no editor para diagnosticar problemas
// ═══════════════════════════════════════════════════════════
function healthCheck() {
  const resultado = { ok: true, servicos: {}, ts: new Date().toISOString() };

  // 1. Validação de configuração
  try {
    _validarConfig();
    resultado.servicos.config = { status: 'OK' };
  } catch(e) {
    resultado.servicos.config = { status: 'ERRO', detalhe: e.message };
    resultado.ok = false;
  }

  // 2. Acesso ao Google Drive
  try {
    const pasta = DriveApp.getFolderById(CONFIG.PASTA_ID);
    resultado.servicos.drive = { status: 'OK', pasta: pasta.getName() };
  } catch(e) {
    resultado.servicos.drive = { status: 'ERRO', detalhe: e.message };
    resultado.ok = false;
  }

  // 3. Acesso à planilha (Sheets)
  try {
    const ss  = abrirPlanilha();
    const abas = ss.getSheets().map(a => a.getName());
    resultado.servicos.sheets = { status: 'OK', planilha: ss.getName(), abas };
  } catch(e) {
    resultado.servicos.sheets = { status: 'ERRO', detalhe: e.message };
    resultado.ok = false;
  }

  // 4. Serviço de e-mail (cota disponível)
  try {
    const cotaRestante = MailApp.getRemainingDailyQuota();
    resultado.servicos.email = { status: cotaRestante > 0 ? 'OK' : 'AVISO', cotaRestante };
    if (cotaRestante === 0) resultado.ok = false;
  } catch(e) {
    resultado.servicos.email = { status: 'ERRO', detalhe: e.message };
    resultado.ok = false;
  }

  // 5. API BNCC externa
  try {
    const resp = UrlFetchApp.fetch('https://bncc.ninja/api/habilidades?q=leitura&limit=1', {
      muteHttpExceptions: true,
      headers: { 'Accept': 'application/json' },
    });
    const code = resp.getResponseCode();
    resultado.servicos.bnccApi = { status: code === 200 ? 'OK' : 'AVISO', httpStatus: code };
    if (code !== 200) resultado.ok = false;
  } catch(e) {
    resultado.servicos.bnccApi = { status: 'AVISO', detalhe: 'API inacessível — modo offline ativo' };
  }

  // 6. CacheService
  try {
    const cache = CacheService.getScriptCache();
    cache.put('_healthcheck_ping', '1', 10);
    const pong  = cache.get('_healthcheck_ping');
    resultado.servicos.cache = { status: pong === '1' ? 'OK' : 'AVISO' };
    cache.remove('_healthcheck_ping');
  } catch(e) {
    resultado.servicos.cache = { status: 'AVISO', detalhe: e.message };
  }

  // Log do resultado
  const icone = resultado.ok ? '✅' : '❌';
  Logger.log(`${icone} healthCheck — ${resultado.ok ? 'TUDO OK' : 'PROBLEMAS DETECTADOS'}`);
  Object.entries(resultado.servicos).forEach(([svc, info]) => {
    const ico = info.status === 'OK' ? '  ✓' : info.status === 'AVISO' ? '  ⚠' : '  ✗';
    Logger.log(`${ico} ${svc}: ${info.status}${info.detalhe ? ' — ' + info.detalhe : ''}`);
  });

  return resultado;
}

/**
 * [V9-17] Retorna uma versão resumida do health check para o painel da coordenação.
 */
function getStatusSistema() {
  _verificarPermissao('getStatusSistema');
  return healthCheck();
}

// ═══════════════════════════════════════════════════════════
//  [V9-20] FERRAMENTAS DE GERENCIAMENTO DA PLANILHA
// ═══════════════════════════════════════════════════════════

/**
 * Exibe informações sobre a planilha de registros.
 * Execute manualmente no editor para ver o ID atual.
 */
function infosPlanilha() {
  try {
    console.log('════════════════════════════════════════════════');
    console.log('   INFORMAÇÕES DA PLANILHA DE REGISTROS');
    console.log('════════════════════════════════════════════════');
    console.log('');
    
    // ID do CONFIG
    const idConfig = CONFIG.PLANILHA_ID;
    const temIdConfig = idConfig && idConfig !== 'COLE_AQUI_O_ID_DA_PLANILHA_REGISTRO';
    console.log(`📋 CONFIG.PLANILHA_ID: ${temIdConfig ? idConfig : '(não configurado)'}`);
    
    // ID das propriedades
    const props = PropertiesService.getScriptProperties();
    const idSalvo = props.getProperty('PLANILHA_REGISTRO_ID');
    console.log(`💾 Propriedades do Script: ${idSalvo || '(nenhum ID salvo)'}`);
    
    // ID efetivo que será usado
    const idEfetivo = _obterIdPlanilha();
    console.log(`✅ ID Efetivo (usado): ${idEfetivo || '(será criado no próximo acesso)'}`);
    console.log('');
    
    // Tenta abrir a planilha
    if (idEfetivo) {
      try {
        const ss = SpreadsheetApp.openById(idEfetivo);
        console.log(`📊 PLANILHA ENCONTRADA:`);
        console.log(`   Nome: ${ss.getName()}`);
        console.log(`   URL: ${ss.getUrl()}`);
        console.log(`   Abas: ${ss.getSheets().map(s => s.getName()).join(', ')}`);
      } catch(e) {
        console.log(`❌ ERRO: ID salvo não é válido (${e.message})`);
        console.log(`   A planilha será recriada no próximo acesso.`);
      }
    } else {
      console.log(`ℹ️  Nenhuma planilha configurada ainda.`);
      console.log(`   Será criada automaticamente no primeiro acesso ao sistema.`);
    }
    
    console.log('');
    console.log('════════════════════════════════════════════════');
  } catch(e) {
    console.log(`Erro ao obter informações: ${e.message}`);
  }
}

/**
 * Reseta o ID da planilha salvo nas propriedades do script.
 * Use se precisar forçar a criação de uma nova planilha.
 * Execute manualmente no editor quando necessário.
 */
function resetarIdPlanilha() {
  try {
    const props = PropertiesService.getScriptProperties();
    const idAnterior = props.getProperty('PLANILHA_REGISTRO_ID');
    
    if (idAnterior) {
      props.deleteProperty('PLANILHA_REGISTRO_ID');
      console.log(`✅ ID da planilha resetado com sucesso!`);
      console.log(`   ID anterior: ${idAnterior}`);
      console.log('');
      console.log(`ℹ️  No próximo acesso, o sistema irá:`);
      console.log(`   1. Buscar uma planilha existente pelo nome`);
      console.log(`   2. Ou criar uma nova se não encontrar`);
    } else {
      console.log(`ℹ️  Nenhum ID estava salvo nas propriedades.`);
    }
  } catch(e) {
    console.log(`Erro ao resetar ID: ${e.message}`);
  }
}

/**
 * [V9-20] Função de diagnóstico completo do sistema
 * Executa uma bateria de testes e retorna um relatório detalhado
 * Útil para troubleshooting quando o sistema não funciona
 * 
 * Como usar:
 * 1. Abra o Apps Script Editor
 * 2. Selecione a função "diagnosticoSistema" na lista de funções
 * 3. Clique em "Executar"
 * 4. Veja o resultado no painel de Logs
 */
function diagnosticoSistema() {
  const relatorio = [];
  relatorio.push('═══════════════════════════════════════════════════════');
  relatorio.push('  DIAGNÓSTICO DO SISTEMA - v9.20');
  relatorio.push('═══════════════════════════════════════════════════════\n');
  
  // 1. TESTE DE AUTENTICAÇÃO
  relatorio.push('1️⃣  TESTE DE AUTENTICAÇÃO');
  relatorio.push('─────────────────────────');
  try {
    const email = Session.getActiveUser().getEmail();
    relatorio.push(`✓ Email detectado: ${email || 'NENHUM'}`);
    
    if (!email || !email.includes('@')) {
      relatorio.push(`❌ PROBLEMA: Email inválido ou não detectado`);
      relatorio.push(`   Solução: Faça deploy como Web App e acesse a URL gerada\n`);
    } else {
      const emailNormalizado = email.toLowerCase().trim();
      const ehCoord = _ehCoordenacao(emailNormalizado);
      
      relatorio.push(`✓ Email normalizado: ${emailNormalizado}`);
      relatorio.push(`✓ É coordenação? ${ehCoord ? 'SIM ✅' : 'NÃO ❌'}`);
      
      if (ehCoord) {
        relatorio.push(`✓ Status: AUTORIZADO como coordenação\n`);
      } else {
        relatorio.push(`⚠️  Email não está na lista de coordenação`);
        relatorio.push(`   Emails autorizados: ${CONFIG.EMAILS_COORDENACAO.join(', ')}\n`);
      }
    }
  } catch(e) {
    relatorio.push(`❌ ERRO ao verificar autenticação: ${e.message}\n`);
  }
  
  // 2. TESTE DE CONFIGURAÇÃO
  relatorio.push('2️⃣  TESTE DE CONFIGURAÇÃO');
  relatorio.push('─────────────────────────');
  try {
    relatorio.push(`✓ PASTA_ID: ${CONFIG.PASTA_ID || 'NÃO CONFIGURADO'}`);
    relatorio.push(`✓ PLANILHA_ID (CONFIG): ${CONFIG.PLANILHA_ID || 'NÃO CONFIGURADO'}`);
    relatorio.push(`✓ Emails coordenação: ${CONFIG.EMAILS_COORDENACAO.join(', ')}`);
    
    // Tenta validar config
    try {
      _validarConfig();
      relatorio.push(`✓ Validação de config: OK ✅\n`);
    } catch(configErr) {
      relatorio.push(`⚠️  Validação de config falhou: ${configErr.message}`);
      relatorio.push(`   Isso pode ser normal se a planilha ainda não foi criada\n`);
    }
  } catch(e) {
    relatorio.push(`❌ ERRO ao verificar configuração: ${e.message}\n`);
  }
  
  // 3. TESTE DE ACESSO À PASTA DO DRIVE
  relatorio.push('3️⃣  TESTE DE ACESSO À PASTA DO DRIVE');
  relatorio.push('─────────────────────────────────────');
  try {
    if (!CONFIG.PASTA_ID || CONFIG.PASTA_ID.includes('COLE_AQUI')) {
      relatorio.push(`❌ PASTA_ID não configurado!`);
      relatorio.push(`   Configure o ID da pasta no Drive em CONFIG.PASTA_ID\n`);
    } else {
      const pasta = DriveApp.getFolderById(CONFIG.PASTA_ID);
      relatorio.push(`✓ Pasta encontrada: ${pasta.getName()}`);
      relatorio.push(`✓ ID da pasta: ${pasta.getId()}`);
      relatorio.push(`✓ URL: ${pasta.getUrl()}\n`);
    }
  } catch(e) {
    relatorio.push(`❌ ERRO ao acessar pasta: ${e.message}`);
    relatorio.push(`   Verifique se:`);
    relatorio.push(`   1. O ID da pasta está correto`);
    relatorio.push(`   2. Você tem permissão de acesso à pasta`);
    relatorio.push(`   3. A pasta não foi excluída\n`);
  }
  
  // 4. TESTE DE PLANILHA
  relatorio.push('4️⃣  TESTE DE ACESSO/CRIAÇÃO DA PLANILHA');
  relatorio.push('───────────────────────────────────────');
  try {
    const idAtual = _obterIdPlanilha();
    relatorio.push(`✓ ID atual da planilha: ${idAtual || 'NENHUM (será criada)'}`);
    
    const ss = abrirPlanilha();
    relatorio.push(`✓ Planilha acessada/criada com sucesso!`);
    relatorio.push(`✓ Nome: ${ss.getName()}`);
    relatorio.push(`✓ ID: ${ss.getId()}`);
    relatorio.push(`✓ URL: ${ss.getUrl()}`);
    
    // Verifica abas
    const abas = ss.getSheets().map(s => s.getName());
    relatorio.push(`✓ Abas encontradas (${abas.length}): ${abas.join(', ')}`);
    
    // Verifica se tem as abas essenciais
    const abasEssenciais = [ABA.CADASTROS, ABA.TRIMESTRAL, ABA.SEMANAL, ABA.ANUAL];
    const faltando = abasEssenciais.filter(aba => !abas.includes(aba));
    
    if (faltando.length > 0) {
      relatorio.push(`⚠️  Abas faltando: ${faltando.join(', ')}`);
      relatorio.push(`   As abas serão criadas automaticamente se necessário\n`);
    } else {
      relatorio.push(`✓ Todas as abas essenciais existem ✅\n`);
    }
  } catch(e) {
    relatorio.push(`❌ ERRO ao acessar/criar planilha: ${e.message}`);
    relatorio.push(`   Stack trace: ${e.stack}\n`);
  }
  
  // 5. TESTE DE CADASTROS
  relatorio.push('5️⃣  TESTE DE CADASTROS');
  relatorio.push('──────────────────────');
  try {
    const cadastros = getCadastros();
    const qtdProfs = Object.keys(cadastros.emails || {}).length;
    const qtdTurmas = cadastros.turmas ? cadastros.turmas.length : 0;
    const qtdComps = cadastros.componentes ? cadastros.componentes.length : 0;
    
    relatorio.push(`✓ Professores cadastrados: ${qtdProfs}`);
    relatorio.push(`✓ Turmas cadastradas: ${qtdTurmas}`);
    relatorio.push(`✓ Componentes cadastrados: ${qtdComps}`);
    
    if (qtdProfs === 0) {
      relatorio.push(`⚠️  Nenhum professor cadastrado ainda`);
      relatorio.push(`   Use o menu "Cadastros" para adicionar professores\n`);
    } else {
      relatorio.push(`✓ Cadastros OK ✅\n`);
    }
  } catch(e) {
    relatorio.push(`❌ ERRO ao verificar cadastros: ${e.message}\n`);
  }
  
  // 6. TESTE DE PERFIL
  relatorio.push('6️⃣  TESTE DE PERFIL DO USUÁRIO');
  relatorio.push('──────────────────────────────');
  try {
    const perfil = getPerfil();
    relatorio.push(`✓ Nome: ${perfil.nome}`);
    relatorio.push(`✓ Email: ${perfil.email}`);
    relatorio.push(`✓ Perfil: ${perfil.perfil}`);
    
    if (perfil.perfil === 'desconhecido') {
      relatorio.push(`❌ ACESSO NEGADO!`);
      relatorio.push(`   Motivo: Email não está cadastrado e não é coordenação`);
      relatorio.push(`   Solução:`);
      relatorio.push(`   1. Adicione o email à lista de coordenação, OU`);
      relatorio.push(`   2. Cadastre o professor na aba Cadastros\n`);
    } else if (perfil.perfil === 'coordenacao') {
      relatorio.push(`✓ Acesso TOTAL como coordenação ✅\n`);
    } else if (perfil.perfil === 'professor') {
      relatorio.push(`✓ Acesso como professor ✅`);
      relatorio.push(`✓ Pasta individual: ${perfil.pastaUrl || 'Não criada ainda'}\n`);
    }
  } catch(e) {
    relatorio.push(`❌ ERRO ao verificar perfil: ${e.message}\n`);
  }
  
  // 7. RESUMO FINAL
  relatorio.push('═══════════════════════════════════════════════════════');
  relatorio.push('  RESUMO E RECOMENDAÇÕES');
  relatorio.push('═══════════════════════════════════════════════════════');
  
  // Verifica se todos os testes passaram
  const temErros = relatorio.some(linha => linha.includes('❌'));
  
  if (!temErros) {
    relatorio.push('✅ SISTEMA FUNCIONANDO CORRETAMENTE!');
    relatorio.push('');
    relatorio.push('Se os menus não aparecem, verifique:');
    relatorio.push('1. Console do navegador (F12) para erros JavaScript');
    relatorio.push('2. Se fez deploy como Web App com acesso apropriado');
    relatorio.push('3. Se está acessando a URL correta do Web App');
  } else {
    relatorio.push('⚠️  PROBLEMAS DETECTADOS - Veja os itens marcados com ❌ acima');
    relatorio.push('');
    relatorio.push('CHECKLIST DE INSTALAÇÃO:');
    relatorio.push('[ ] 1. Configure CONFIG.PASTA_ID com ID da pasta do Drive');
    relatorio.push('[ ] 2. Configure CONFIG.EMAILS_COORDENACAO com emails autorizados');
    relatorio.push('[ ] 3. Faça deploy como Web App (Implantar > Nova implantação)');
    relatorio.push('[ ] 4. Configure "Executar como: Eu"');
    relatorio.push('[ ] 5. Configure "Quem tem acesso: Qualquer pessoa"');
    relatorio.push('[ ] 6. Copie a URL do Web App e acesse pelo navegador');
    relatorio.push('[ ] 7. Autorize as permissões solicitadas pelo Google');
  }
  
  relatorio.push('═══════════════════════════════════════════════════════\n');
  
  // Imprime o relatório
  const textoCompleto = relatorio.join('\n');
  console.log(textoCompleto);
  
  // Retorna também como objeto para poder ser chamado via google.script.run
  return textoCompleto;
}
