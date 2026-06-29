// ================================================================
// SENTINELA v2.0 — SETUP INICIAL (Execute UMA vez no editor)
// ================================================================
// Este script cria toda a infraestrutura necessária no Google Drive:
// - Planilha "Sentinela_DB" com 5 abas
// - Pasta "Sentinela_Fotos_Mandados"
// - Pasta "Sentinela_Fotos_Usuarios"
// ================================================================

/**
 * FUNÇÃO PRINCIPAL — Execute esta função no editor para criar tudo.
 */
function executarSetupInicial() {
  const ui = SpreadsheetApp.getUi();

  try {
    // ============================================================
    // 1. CRIAR OU LOCALIZAR A PLANILHA PRINCIPAL
    // ============================================================
    let planilha = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log("✅ Planilha ativa: " + planilha.getName());

    // ============================================================
    // 2. CRIAR ABAS NECESSÁRIAS
    // ============================================================
    criarAbaSeNaoExiste(planilha, "Mandados", getHeadersGeral());
    criarAbaSeNaoExiste(planilha, "Usuarios", getHeadersUsuarios());
    criarAbaSeNaoExiste(planilha, "Config", null); // Populada separadamente
    criarAbaSeNaoExiste(planilha, "Gamificacao", getHeadersGamificacao());
    criarAbaSeNaoExiste(planilha, "Poligonos", getHeadersPoligonos());
    criarAbaSeNaoExiste(planilha, "Conferencia", getHeadersConferencia());
    criarAbaSeNaoExiste(planilha, "Historico", getHeadersHistorico());
    criarAbaSeNaoExiste(planilha, "Leis", getHeadersLeis());
    criarAbaSeNaoExiste(planilha, "Notificacoes", getHeadersNotificacoes());

    // Popular leis padrão se aba estiver vazia
    popularLeisPadrao(planilha);

    // Remover aba padrão "Sheet1" / "Página1" se existir e houver outras abas
    const abaPadrao = planilha.getSheetByName("Sheet1") || planilha.getSheetByName("Página1") || planilha.getSheetByName("Plan1");
    if (abaPadrao && planilha.getSheets().length > 1) {
      try { planilha.deleteSheet(abaPadrao); } catch(e) {}
    }

    // ============================================================
    // 3. POPULAR ABA CONFIG COM VALORES PADRÃO
    // ============================================================
    popularConfigPadrao(planilha);

    // ============================================================
    // 4. CRIAR PASTAS NO GOOGLE DRIVE
    // ============================================================
    criarPastaSeNaoExiste("Sentinela_Fotos_Mandados");
    criarPastaSeNaoExiste("Sentinela_Fotos_Usuarios");

    // ============================================================
    // 5. REGISTRAR ADMIN INICIAL (quem executou o setup)
    // ============================================================
    registrarAdminInicial(planilha);

    Logger.log("SETUP COMPLETO! Infraestrutura do Sentinela v2.5 criada com sucesso.");
    ui.alert("SETUP COMPLETO!\n\nToda a infraestrutura foi criada:\n- Abas: Mandados, Usuarios, Config, Gamificacao, Poligonos, Conferencia, Historico, Leis, Notificacoes\n- Pastas: Sentinela_Fotos_Mandados, Sentinela_Fotos_Usuarios\n- Leis/Artigos padrão populados\n- Seu e-mail foi registrado como Admin.\n\nAgora publique o Web App.");

  } catch (erro) {
    Logger.log("❌ ERRO NO SETUP: " + erro.message);
    ui.alert("❌ ERRO NO SETUP:\n\n" + erro.message);
  }
}

// ================================================================
// FUNÇÕES DE CRIAÇÃO DE ABAS
// ================================================================

function criarAbaSeNaoExiste(planilha, nomeAba, headers) {
  let aba = planilha.getSheetByName(nomeAba);
  if (!aba) {
    aba = planilha.insertSheet(nomeAba);
    Logger.log("📄 Aba criada: " + nomeAba);
  } else {
    Logger.log("📄 Aba já existe: " + nomeAba);
  }

  if (headers && headers.length > 0) {
    const headerExistente = aba.getRange(1, 1, 1, headers.length).getValues()[0];
    const estaVazio = headerExistente.every(cell => cell === "" || cell === null);

    if (estaVazio) {
      aba.getRange(1, 1, 1, headers.length).setValues([headers]);
      aba.getRange(1, 1, 1, headers.length)
        .setFontWeight("bold")
        .setBackground("#1e293b")
        .setFontColor("#e2e8f0");
      aba.setFrozenRows(1);
      Logger.log("  → Headers inseridos em: " + nomeAba);
    }
  }
  return aba;
}

// ================================================================
// HEADERS DE CADA ABA
// ================================================================

function getHeadersGeral() {
  return [
    "Data de Lancamento",     // A (0)
    "Data de Conferencia",    // B (1)
    "Mandado",                // C (2)
    "Artigo",                 // D (3)
    "Nome",                   // E (4)
    "CPF",                    // F (5)
    "RG",                     // G (6)
    "Nascimento",             // H (7)
    "Naturalidade",           // I (8)
    "Sexo",                   // J (9)
    "Cor",                    // K (10)
    "Filiacao",               // L (11)
    "Foto URL",               // M (12)
    "Batalhao",               // N (13)
    "Endereco Principal",     // O (14)
    "Outros Enderecos",       // P (15)
    "Status",                 // Q (16)
    "Validade",               // R (17)
    "Info Processuais",       // S (18)
    "Geodata Secundarios",    // T (19)
    "Dados Extras JSON",      // U (20)
    "Observacoes",            // V (21)
    "Latitude",               // W (22)
    "Longitude",              // X (23)
    "CPI",                    // Y (24)
    "BPM Area",               // Z (25)
    "CIA Area",               // AA (26)
    "DP Area",                // AB (27)
    "Cidade",                 // AC (28)
    "TipoImportacao"          // AD (29)
  ];
}

/**
 * v4.1.0: Garante que a aba Mandados existe e possui o cabeçalho correto.
 * - Se a aba não existir, cria com cabeçalho completo.
 * - Se existir mas o cabeçalho estiver vazio, insere os headers.
 * - Se existir com cabeçalho, verifica colunas faltantes e adiciona ao final.
 * 
 * Pode ser executada diretamente pelo editor ou chamada por outras funções.
 */
function garantirCabecalhoMandados() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName("Mandados");
  const headersEsperados = getHeadersGeral();

  // Se a aba não existir, cria do zero
  if (!aba) {
    aba = planilha.insertSheet("Mandados");
    aba.getRange(1, 1, 1, headersEsperados.length).setValues([headersEsperados]);
    aba.getRange(1, 1, 1, headersEsperados.length)
      .setFontWeight("bold")
      .setBackground("#1e293b")
      .setFontColor("#e2e8f0");
    aba.setFrozenRows(1);
    Logger.log("✅ Aba Mandados criada com " + headersEsperados.length + " colunas.");
    return { criada: true, colunasFaltantes: [], total: headersEsperados.length };
  }

  // Se existir, verificar cabeçalho atual
  const lastCol = aba.getLastColumn();
  const colunasFaltantes = [];

  if (lastCol === 0) {
    // Aba existe mas está totalmente vazia
    aba.getRange(1, 1, 1, headersEsperados.length).setValues([headersEsperados]);
    aba.getRange(1, 1, 1, headersEsperados.length)
      .setFontWeight("bold")
      .setBackground("#1e293b")
      .setFontColor("#e2e8f0");
    aba.setFrozenRows(1);
    Logger.log("✅ Cabeçalho inserido na aba Mandados vazia (" + headersEsperados.length + " colunas).");
    return { criada: false, colunasFaltantes: headersEsperados, total: headersEsperados.length };
  }

  // Aba existe com dados — verificar quais colunas estão faltando
  const headersAtuais = aba.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });

  headersEsperados.forEach(function(h) {
    if (headersAtuais.indexOf(h) === -1) {
      colunasFaltantes.push(h);
    }
  });

  // Adicionar colunas faltantes ao final
  if (colunasFaltantes.length > 0) {
    const startCol = lastCol + 1;
    aba.getRange(1, startCol, 1, colunasFaltantes.length).setValues([colunasFaltantes]);
    aba.getRange(1, startCol, 1, colunasFaltantes.length)
      .setFontWeight("bold")
      .setBackground("#1e293b")
      .setFontColor("#e2e8f0");
    Logger.log("✅ " + colunasFaltantes.length + " coluna(s) adicionada(s): " + colunasFaltantes.join(", "));
  } else {
    Logger.log("✅ Cabeçalho da aba Mandados está completo. Nenhuma coluna faltante.");
  }

  return { criada: false, colunasFaltantes: colunasFaltantes, total: headersAtuais.length + colunasFaltantes.length };
}

function getHeadersUsuarios() {
  return [
    "Email",              // A (0)
    "Nome",               // B (1)
    "Unidade",            // C (2)
    "Perfil",             // D (3)
    "Status",             // E (4)
    "Foto URL",           // F (5)
    "Pontos",             // G (6)
    "Badges",             // H (7)
    "Data Cadastro",      // I (8)
    "Ultimo acesso",      // J (9)
    "Qtd de acesso",      // K (10)
    "Status de Conexao"   // L (11)
  ];
}

function getHeadersGamificacao() {
  return [
    "Data",           // A
    "Email",          // B
    "Acao",           // C
    "Pontos",         // D
    "Mandado Ref",    // E
    "Descricao"       // F
  ];
}

function getHeadersPoligonos() {
  return [
    "Batalhao",       // A
    "Cia",            // B
    "Setor",          // C
    "Cidade",         // D
    "Delegacia",      // E
    "GeoJSON",        // F
    "Cor",            // G
    "Ativo"           // H
  ];
}

function getHeadersConferencia() {
  return [
    "Data Importacao", // A
    "Mandado",         // B
    "Status no CSV",   // C
    "Status Sentinela",// D
    "Situacao",        // E
    "Nome Procurado",  // F
    "Batalhao",        // G
    "Conferido Por"    // H
  ];
}

function getHeadersHistorico() {
  return [
    "Data",            // A
    "Email",           // B
    "Nome",            // C
    "Acao",            // D
    "Detalhes",        // E
    "IP/Dispositivo"   // F
  ];
}

// ================================================================
// CONFIGURAÇÕES PADRÃO (Aba Config — Chave/Valor)
// ================================================================

function popularConfigPadrao(planilha) {
  const aba = planilha.getSheetByName("Config");
  if (!aba) return;

  // Verifica se já foi populada
  const primeiraChave = aba.getRange("A1").getValue();
  if (primeiraChave === "CHAVE") {
    Logger.log("⚙️ Config já populada, pulando...");
    return;
  }

  const configs = [
    // ---- IDENTIDADE ----
    ["CHAVE", "VALOR", "CATEGORIA", "DESCRICAO"],
    ["app_nome", "Sentinela", "Identidade", "Nome do aplicativo exibido no header e sidebar"],
    ["app_slogan", "Sistema Tático de Inteligência", "Identidade", "Slogan exibido abaixo do nome"],
    ["app_icone", "local_police", "Identidade", "Ícone Material Symbols do logo (usado se não houver imagem)"],
    ["app_logo_url", "", "Identidade", "URL da imagem de logo (sobrepõe o ícone se preenchida)"],

    // ---- APARÊNCIA ----
    ["cor_primaria", "#6c63ff", "Aparencia", "Cor principal do sistema (botões, destaques)"],
    ["cor_fundo", "#2d2d3f", "Aparencia", "Cor de fundo base"],
    ["cor_superficie", "#353548", "Aparencia", "Cor das superfícies elevadas (cards, painéis)"],
    ["cor_texto", "#e8e8f0", "Aparencia", "Cor do texto principal"],
    ["cor_sucesso", "#2ecc71", "Aparencia", "Cor de sucesso (capturado, aprovado)"],
    ["cor_perigo", "#e74c3c", "Aparencia", "Cor de perigo (urgente, bloqueado)"],
    ["cor_alerta", "#f39c12", "Aparencia", "Cor de alerta (pendente, atenção)"],

    // ---- MAPA ----
    ["mapa_api_key", "AIzaSyCGSU6lmAWHszLK92vZRgeKHTSd07kXha4", "Mapa", "Chave da API Google Maps"],
    ["mapa_zoom_padrao", "10", "Mapa", "Zoom inicial do mapa"],
    ["mapa_lat_padrao", "-22.9056", "Mapa", "Latitude do centro padrão"],
    ["mapa_lng_padrao", "-47.0608", "Mapa", "Longitude do centro padrão"],
    ["mapa_tipo", "roadmap", "Mapa", "Tipo do mapa: roadmap, satellite, hybrid, terrain"],

    // ---- VALIDADE ----
    ["validade_verde_dias", "91", "Validade", "Acima deste valor = verde (seguro)"],
    ["validade_amarelo_dias", "90", "Validade", "Abaixo ou igual = amarelo (atenção)"],
    ["validade_laranja_dias", "30", "Validade", "Abaixo ou igual = laranja (alerta)"],
    ["validade_vermelho_dias", "15", "Validade", "Abaixo ou igual = vermelho piscante (urgente)"],

    // ---- WHATSAPP ----
    ["whatsapp_numero", "5519992693763", "WhatsApp", "Número do Admin para notificações WhatsApp"],
    ["whatsapp_ativo", "true", "WhatsApp", "Habilitar opção de WhatsApp no cadastro"],

    // ---- PERMISSÕES DO PATRULHEIRO ----
    ["perm_patr_mudar_status", "true", "Permissoes", "Patrulheiro pode mudar status (Procurado/Capturado)"],
    ["perm_patr_add_obs", "true", "Permissoes", "Patrulheiro pode adicionar observações"],
    ["perm_patr_ver_cpf", "true", "Permissoes", "Patrulheiro pode ver CPF do procurado"],
    ["perm_patr_ver_processo", "true", "Permissoes", "Patrulheiro pode ver dados processuais"],
    ["perm_patr_ver_foto", "true", "Permissoes", "Patrulheiro pode ver foto do procurado"],
    ["perm_patr_ver_filiacao", "true", "Permissoes", "Patrulheiro pode ver filiação"],
    ["perm_patr_ver_endereco_sec", "true", "Permissoes", "Patrulheiro pode ver endereços secundários"],

    // ---- GAMIFICAÇÃO ----
    ["gamif_ativo", "true", "Gamificacao", "Habilitar sistema de gamificação"],
    ["gamif_pontos_relato", "10", "Gamificacao", "Pontos por relato de abordagem"],
    ["gamif_pontos_captura", "100", "Gamificacao", "Pontos por captura confirmada"],
    ["gamif_pontos_intel", "25", "Gamificacao", "Pontos por informação de inteligência"],

    // ---- TEXTOS DE TOOLTIPS ----
    ["tooltip_data_lancamento", "Data em que o administrador inseriu este mandado no sistema.", "Tooltips", "Texto do tooltip da data de lançamento"],
    ["tooltip_data_conferencia", "Última data em que o administrador verificou se este mandado ainda está ativo.", "Tooltips", "Texto do tooltip da data de conferência"],
    ["tooltip_validade", "Data de validade do mandado conforme consta no documento oficial.", "Tooltips", "Texto do tooltip da validade"],
    ["tooltip_data_documento", "Data e local de emissão do documento oficial.", "Tooltips", "Texto do tooltip da data do documento"],
    ["tooltip_enderecos", "Endereços conhecidos do procurado, geocodificados e mapeados. O endereço principal é onde há maior probabilidade de localização.", "Tooltips", "Texto do tooltip dos endereços"],
    ["tooltip_area_interativa", "Área para registro de abordagens e mudança de status. Qualquer alteração é registrada no banco de dados.", "Tooltips", "Texto do tooltip da área interativa"],
    ["tooltip_distancia", "Distância aproximada em linha reta da sua posição atual até o endereço indicado.", "Tooltips", "Texto do tooltip da distância"]
  ];

  aba.getRange(1, 1, configs.length, configs[0].length).setValues(configs);

  // Estilizar header
  aba.getRange(1, 1, 1, 4)
    .setFontWeight("bold")
    .setBackground("#1e293b")
    .setFontColor("#e2e8f0");
  aba.setFrozenRows(1);

  // Ajustar larguras
  aba.setColumnWidth(1, 250);
  aba.setColumnWidth(2, 350);
  aba.setColumnWidth(3, 120);
  aba.setColumnWidth(4, 400);

  Logger.log("⚙️ Configurações padrão populadas com sucesso.");
}

// ================================================================
// CRIAÇÃO DE PASTAS NO DRIVE
// ================================================================

function criarPastaSeNaoExiste(nomePasta) {
  const pastas = DriveApp.getFoldersByName(nomePasta);
  if (pastas.hasNext()) {
    Logger.log("📁 Pasta já existe: " + nomePasta);
    return pastas.next();
  }
  const novaPasta = DriveApp.createFolder(nomePasta);
  Logger.log("📁 Pasta criada: " + nomePasta);
  return novaPasta;
}

// ================================================================
// REGISTRAR ADMIN INICIAL
// ================================================================

function registrarAdminInicial(planilha) {
  const emailAdmin = Session.getActiveUser().getEmail();
  if (!emailAdmin) {
    Logger.log("⚠️ Não foi possível obter o e-mail do usuário ativo.");
    return;
  }

  const abaUsuarios = planilha.getSheetByName("Usuarios");
  const dados = abaUsuarios.getDataRange().getValues();

  // Verifica se já existe
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0].toString().trim().toLowerCase() === emailAdmin.toLowerCase()) {
      Logger.log("👤 Admin já registrado: " + emailAdmin);
      return;
    }
  }

  abaUsuarios.appendRow([
    emailAdmin,       // A - Email
    "Administrador",  // B - Nome
    "Comando",        // C - Unidade
    "Admin",          // D - Perfil
    "ATIVO",          // E - Status
    "",               // F - Foto URL
    "0",              // G - Pontos
    ""                // H - Badges
  ]);
  Logger.log("👤 Admin registrado: " + emailAdmin);
}

// ================================================================
// VERIFICAÇÃO DE INTEGRIDADE (Execute para diagnosticar problemas)
// ================================================================

function verificarIntegridade() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const resultados = [];

  // Verificar abas
  const abasNecessarias = ["Mandados", "Usuarios", "Config", "Gamificacao", "Poligonos", "Conferencia", "Historico", "Leis", "Notificacoes"];
  abasNecessarias.forEach(nome => {
    const aba = planilha.getSheetByName(nome);
    resultados.push(aba ? "✅ Aba '" + nome + "' OK" : "❌ Aba '" + nome + "' AUSENTE");
  });

  // Verificar pastas
  const pastasNecessarias = ["Sentinela_Fotos_Mandados", "Sentinela_Fotos_Usuarios"];
  pastasNecessarias.forEach(nome => {
    const pastas = DriveApp.getFoldersByName(nome);
    resultados.push(pastas.hasNext() ? "✅ Pasta '" + nome + "' OK" : "❌ Pasta '" + nome + "' AUSENTE");
  });

  // Verificar Config
  const abaConfig = planilha.getSheetByName("Config");
  if (abaConfig) {
    const numConfigs = abaConfig.getLastRow() - 1;
    resultados.push("⚙️ Config: " + numConfigs + " parâmetros cadastrados");
  }

  const relatorio = resultados.join("\n");
  Logger.log("\n=== RELATORIO DE INTEGRIDADE ===\n" + relatorio);
  SpreadsheetApp.getUi().alert("RELATORIO DE INTEGRIDADE\n\n" + relatorio);
}

// ================================================================
// HEADERS: LEIS E NOTIFICAÇÕES
// ================================================================

function getHeadersLeis() {
  return [
    "Categoria",        // A — Nome da categoria (ex: Homicídio)
    "Palavras Chave",   // B — Palavras separadas por | (ex: HOMIC|121)
    "Cor",              // C — Cor hex (ex: #8b0000)
    "Icone SVG",        // D — SVG path (legado)
    "Ordem",            // E — Ordem de exibição
    "Ativo",            // F — SIM/NAO
    "PinoTexto",        // G — Texto curto para o pino do mapa (ex: ROUBO, JACK, 171)
    "Lei Nome",         // H — Nome extenso da lei (ex: Código Penal)
    "Numero Lei",       // I — Número (ex: 11.343)
    "Artigo",           // J — Artigo (ex: 33)
    "Paragrafo",        // K — Parágrafo (ex: 1)
    "Inciso",           // L — Inciso (ex: I)
    "Tipificacao Completa" // M — Descrição completa do crime
  ];
}

function getHeadersNotificacoes() {
  return [
    "Data",             // A — Data/hora da notificação
    "Titulo",           // B — Título curto
    "Mensagem",         // C — Corpo da notificação
    "Tipo",             // D — info/warning/success/danger
    "Para",             // E — Email destinatário ou "TODOS"
    "Lida",             // F — SIM/NAO
    "De"                // G — Email de quem criou
  ];
}

/**
 * Popula a aba Leis com categorias de crimes padrão.
 * Só executa se a aba estiver vazia (apenas headers).
 */
function popularLeisPadrao(planilha) {
  var aba = planilha.getSheetByName("Leis");
  if (!aba) return;
  if (aba.getLastRow() > 1) return; // Já tem dados

  // Coluna D (Icone SVG) mantida como legado — preenchida com vazio
  var leis = [
    // [Categoria, Palavras Chave, Cor, Icone SVG (legado), Ordem, Ativo, PinoTexto]
    ["Homicídio",         "HOMIC|121|LATROCÍNIO|LATROCINIO",                                         "#8b0000", "", 1,  "SIM", "HOMIC"],
    ["Estupro",           "ESTUPRO|213|217|DIGNIDADE SEXUAL",                                        "#991b1b", "", 2,  "SIM", "JACK"],
    ["V. Doméstica",      "MARIA DA PENHA|VIOLÊNCIA DOMÉSTICA|VIOLENCIA DOMESTICA|11340",             "#e11d48", "", 3,  "SIM", "PENHA"],
    ["Roubo",             "ROUBO|157",                                                                "#dc2626", "", 4,  "SIM", "ROUBO"],
    ["Furto",             "FURTO|155",                                                                "#ea580c", "", 5,  "SIM", "FURTO"],
    ["Tráfico",           "TRÁFICO|TRAFICO|DROGAS|ENTORPECENTE",                                      "#7c3aed", "", 6,  "SIM", "TRÁF"],
    ["Estelionato",       "ESTELIONATO|171",                                                          "#ca8a04", "", 7,  "SIM", "171"],
    ["Receptação",        "RECEPTAÇÃO|RECEPTACAO|180",                                                 "#b45309", "", 8,  "SIM", "180"],
    ["Alimentícia",       "CIVIL|ALIMENTO|PENSÃO",                                                    "#2563eb", "", 9,  "SIM", "CIVIL"],
    ["Porte de Arma",     "ARMA|PORTE|POSSE|244|10826",                                               "#475569", "", 10, "SIM", "ARMA"],
    ["Trânsito",          "TRÂNSITO|TRANSITO|CTB|9503|EMBRIAGUEZ",                                    "#0891b2", "", 11, "SIM", "CTB"],
    ["Fiscal",            "FISCAL|TRIBUTÁRIO|TRIBUTARIO|SONEGAÇÃO|SONEGACAO",                          "#059669", "", 12, "SIM", "FISCAL"],
    ["Organização Crim.", "ASSOCIAÇÃO|ASSOCIACAO|ORGANIZAÇÃO CRIMINOSA|ORGANIZACAO CRIMINOSA|288",     "#4338ca", "", 13, "SIM", "288"],
    ["Ameaça",            "AMEAÇA|AMEACA|147",                                                        "#9f1239", "", 14, "SIM", "147"],
    ["Sequestro",         "SEQUESTRO|CÁRCERE|CARCERE|148",                                            "#7f1d1d", "", 15, "SIM", "148"],
    ["Lesão Corporal",    "LESÃO CORPORAL|LESAO CORPORAL|129",                                        "#b91c1c", "", 16, "SIM", "LESÃO"],
    ["Outros",            "",                                                                         "#6b7280", "", 99, "SIM", "CRIME"]
  ];

  aba.getRange(2, 1, leis.length, 7).setValues(leis);
  Logger.log("Leis padrao populadas: " + leis.length + " categorias.");
}

// ================================================================
// INSTALAR TRIGGERS AUTOMÁTICOS
// ================================================================
// Execute esta função UMA VEZ para instalar os triggers necessários.
// O onChange é OBRIGATÓRIO para detectar exclusão de linhas na planilha.

/**
 * Instala triggers onChange e onEdit na planilha ativa.
 * Remove triggers antigos para evitar duplicatas.
 */
function instalarTriggers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Remover triggers antigos do projeto para esta planilha
  var triggers = ScriptApp.getProjectTriggers();
  var removidos = 0;
  triggers.forEach(function(trigger) {
    var fn = trigger.getHandlerFunction();
    // Inclui auditarIntegridadePlanilha para limpar triggers antigos do Antigravity (removido na v4.6.0)
    if (fn === 'onChange' || fn === 'onEdit' || fn === 'executarCeifador' || fn === 'auditarIntegridadePlanilha') {
      ScriptApp.deleteTrigger(trigger);
      removidos++;
    }
  });
  
  // 2. Criar trigger onChange (captura INSERT_ROW, REMOVE_ROW, EDIT, OTHER)
  ScriptApp.newTrigger('onChange')
    .forSpreadsheet(ss)
    .onChange()
    .create();
  
  // 3. Criar trigger onEdit (captura edições simples de células)
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
    
  // 4. Criar trigger de Tempo para o Ceifador (executa a cada 1 minuto)
  ScriptApp.newTrigger('executarCeifador')
    .timeBased()
    .everyMinutes(1)
    .create();
  
  Logger.log("Triggers instalados com sucesso! (Removidos " + removidos + " antigos)");
  Logger.log("  onChange: Detecta inserção/exclusão de linhas");
  Logger.log("  onEdit: Detecta edições em células");
  Logger.log("  executarCeifador: A cada 1 minuto");
  
  try {
    SpreadsheetApp.getUi().alert(
      "TRIGGERS INSTALADOS!\n\n" +
      "onChange — Detecta inserção e exclusão de linhas\n" +
      "onEdit — Detecta edições em células\n" +
      "executarCeifador — Verifica inatividade a cada 1 min\n\n" +
      (removidos > 0 ? "(" + removidos + " triggers antigos foram removidos para evitar duplicatas)" : "") +
      "\nAgora o sistema atualiza automaticamente quando\nalguém exclui ou edita direto na planilha."
    );
  } catch(e) {}
}

/**
 * ATUALIZADOR AUTOMÁTICO DA ABA LEIS (V3.9.76+)
 * Execute essa função para recriar a aba Leis com os novos 13 cabeçalhos.
 */
function atualizarAbaLeisParaV3() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var abaAntiga = planilha.getSheetByName("Leis");
  
  if (abaAntiga) {
    var dataAgora = new Date().getTime();
    abaAntiga.setName("Leis_BKP_" + dataAgora);
    Logger.log("Aba antiga renomeada para Leis_BKP_" + dataAgora);
  }
  
  criarAbaSeNaoExiste(planilha, "Leis", getHeadersLeis());
  popularLeisPadrao(planilha);
  
  SpreadsheetApp.getUi().alert("✅ ABA LEIS ATUALIZADA!\n\nA nova aba 'Leis' foi criada com as 13 colunas necessárias para a v3.9.76.\nSua aba antiga foi renomeada para 'Leis_BKP_...'.\n\nAgora você pode copiar seus dados da planilha CSV e colar na nova aba.");
}
