// ================================================================
// SENTINELA v2.0 — CONFIGURAÇÃO WEB & MOTOR DE CONFIGURAÇÕES
// ================================================================

/**
 * Ponto de entrada do Web App.
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sentinela — Sistema Tático')
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
  ui.createMenu('Sentinela')
    .addItem('Executar Setup Inicial', 'executarSetupInicial')
    .addItem('Geocodificar Pendentes', 'processarPendentes')
    .addItem('Importar Poligonos CPI-2', 'importarPoligonosCPI2')
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

  return configs;
}

/**
 * Salva múltiplas configurações de uma vez (Admin only).
 * @param {Object} novasConfigs — { chave: novoValor, chave2: novoValor2, ... }
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

    for (let i = 1; i < dados.length; i++) {
      const chaveAtual = dados[i][0].toString().trim();
      if (novasConfigs.hasOwnProperty(chaveAtual)) {
        aba.getRange(i + 1, 2).setValue(novasConfigs[chaveAtual]);
        atualizados++;
      }
    }

    return { sucesso: true, mensagem: atualizados + " configuração(ões) atualizada(s) com sucesso." };

  } catch (erro) {
    return { sucesso: false, mensagem: "Erro ao salvar configurações: " + erro.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Retorna a API Key do Google Maps (chamada apenas pelo backend
 * para injetar no template de forma segura, sem expor no HTML source).
 */
function obterApiKeyMaps() {
  const configs = obterConfiguracoesSimples();
  return configs["mapa_api_key"] || "";
}
