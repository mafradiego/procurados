// ================================================================
// SENTINELA v2.0 — CONFIGURAÇÃO WEB & MOTOR DE CONFIGURAÇÕES
// ================================================================

/**
 * Ponto de entrada do Web App.
 */
function doGet(e) {
  // Invalidar cache do usuário ao recarregar a página para refletir mudanças manuais na planilha instantaneamente
  try {
    const emailAtivo = Session.getActiveUser().getEmail();
    if (emailAtivo) {
      const cache = CacheService.getScriptCache();
      cache.remove("AUTH_V2_" + emailAtivo.toLowerCase());
    }
  } catch (err) {
    // Ignorar falhas de leitura de e-mail no escopo de carregamento
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('PAPA-ROMEUS — Sistema Tático')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/**
 * Inclui arquivos HTML parciais (CSS, JS).
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Menu na barra da planilha.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('PAPA-ROMEUS')
    .addItem('Executar Setup Inicial', 'executarSetupInicial')
    .addItem('Geocodificar Pendentes', 'processarPendentes')
    .addItem('Parar Geocod. Automática', 'pararGeocodificacaoAutomatica')
    .addItem('Gerar Dados Fictícios', 'gerarDadosFicticios')
    .addItem('Remover Endereços Duplicados', 'removerEnderecosDuplicados')
    .addSeparator()
    .addItem('Atualizar Headers v2.5 (Migracao)', 'atualizarHeadersMandados')
    .addItem('Verificar Integridade', 'verificarIntegridade')
    .addToUi();
}

// ================================================================
// MOTOR DE CONFIGURAÇÕES (Leitura e Escrita da aba Config)
// ================================================================

/**
 * Lê TODAS as configurações da aba Config e retorna como objeto JSON.
 * Chamada pelo frontend para aplicar configs dinâmicas.
 */
function obterConfiguracoes() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  if (!aba) return {};

  const dados = aba.getDataRange().getValues();
  const configs = {};

  for (let i = 1; i < dados.length; i++) {
    const chave = dados[i][0].toString().trim();
    const valor = dados[i][1].toString().trim();
    const categoria = dados[i][2].toString().trim();

    if (chave) {
      configs[chave] = {
        valor: valor,
        categoria: categoria,
        descricao: dados[i][3] ? dados[i][3].toString().trim() : ""
      };
    }
  }

  // Cota dinâmica para geocodificação se não existir na planilha
  if (!configs["geocodificacao_limite_mensal"]) {
    configs["geocodificacao_limite_mensal"] = {
      valor: "40000",
      categoria: "Mapa",
      descricao: "Limite mensal de endereços geocodificados na API do Google Maps."
    };
  }

  return configs;
}

/**
 * Retorna apenas os valores das configs (sem metadados).
 * Versão leve para o frontend usar diretamente.
 */
function obterConfiguracoesSimples() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  if (!aba) return {};

  const dados = aba.getDataRange().getValues();
  const configs = {};

  for (let i = 1; i < dados.length; i++) {
    const chave = dados[i][0].toString().trim();
    const valor = dados[i][1].toString().trim();
    if (chave) configs[chave] = valor;
  }

  if (!configs["geocodificacao_limite_mensal"]) {
    configs["geocodificacao_limite_mensal"] = "40000";
  }

  return configs;
}

/**
 * Salva múltiplas configurações de uma vez (Admin only).
 * @param {Object} novasConfigs — { chave: novoValor, ... }
 */
function salvarConfiguracoes(novasConfigs) {
  // Verificar permissão Admin
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    throw new Error("Acesso negado: Apenas administradores podem alterar configurações.");
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
    if (!aba) throw new Error("Aba Config não encontrada.");

    const dados = aba.getDataRange().getValues();
    let atualizados = 0;
    const chavesExistentes = new Set(dados.map(r => r[0].toString().trim()));

    // Atualizar linhas existentes
    for (let i = 1; i < dados.length; i++) {
      const chaveAtual = dados[i][0].toString().trim();
      if (novasConfigs.hasOwnProperty(chaveAtual)) {
        aba.getRange(i + 1, 2).setValue(novasConfigs[chaveAtual]);
        atualizados++;
      }
    }

    // Inserir novas configurações que não existiam na planilha
    Object.keys(novasConfigs).forEach(chave => {
      if (!chavesExistentes.has(chave) && novasConfigs[chave] !== undefined) {
        aba.appendRow([chave, novasConfigs[chave], "Mapa", "Configuração adicionada dinamicamente"]);
        atualizados++;
      }
    });

    return { sucesso: true, mensagem: atualizados + " configuração(ões) atualizada(s) com sucesso." };

  } catch (erro) {
    return { sucesso: false, mensagem: "Erro ao salvar configurações: " + erro.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Retorna a versão atual dos dados para o Cache do IndexedDB (Frontend).
 * Baseado em PropertiesService e fallback estrutural para evitar loops ou permissões Extras (DriveApp).
 */
function obterVersaoBancos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const abaPoligonos = ss.getSheetByName("Poligonos");
  const linhasPoligonos = abaPoligonos ? abaPoligonos.getLastRow() : 0;
  
  const abaMandados = ss.getSheetByName("Mandados");
  const linhasMandados = abaMandados ? abaMandados.getLastRow() : 0;
  
  // Lemos um timestamp no Script Properties, que nós vamos atualizar
  // sempre que o sistema gravar algo novo.
  const props = PropertiesService.getScriptProperties();
  let ts = props.getProperty('DB_UPDATE_TIMESTAMP');
  if (!ts) {
    ts = new Date().getTime().toString();
    props.setProperty('DB_UPDATE_TIMESTAMP', ts);
  }

  return {
    poligonosVersao: "v_" + linhasPoligonos + "_" + ts,
    mandadosVersao: "v_" + linhasMandados + "_" + ts
  };
}

/**
 * Função utilitária para o backend invalidar o cache dos clientes.
 * Deve ser chamada em toda operação de salvar/editar Mandados.
 */
function invalidarCacheMandados() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('DB_UPDATE_TIMESTAMP', new Date().getTime().toString());
}

/**
 * Retorna a API Key do Google Maps (chamada apenas pelo backend
 * para injetar no template de forma segura, sem expor no HTML source).
 */
function obterApiKeyMaps() {
  const configs = obterConfiguracoesSimples();
  return configs["mapa_api_key"] || "";
}
