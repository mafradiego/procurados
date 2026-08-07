// ================================================================
// SENTINELA — MÓDULO FIREBASE SYNC (FirebaseSync.gs)
// Comunicação com Firebase Realtime Database via REST API
// Usa Service Account para autenticação JWT/OAuth2
// ================================================================

// --- CONFIGURAÇÃO (Server-side only - não visível no frontend) ---
var FIREBASE_DATABASE_URL_ = 'https://gen-lang-client-0691254724-default-rtdb.firebaseio.com';
var FIREBASE_CLIENT_EMAIL_ = 'firebase-adminsdk-fbsvc@gen-lang-client-0691254724.iam.gserviceaccount.com';
var FIREBASE_PRIVATE_KEY_ = '-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQC0eIb+PP6FyuP+\nn8RFoofexPYkd30jgZM0FfOhCglyFFKdpUVdacbhrXWYwz7QxfmwyZLoQXi5Qjf3\no18IXLZdtxGwqrhHrPU/MviH/UY0vdoMVO02FKss2XeQBg1EXBRhbV505yFTWztI\nfpqE1WJ7sqk0FVH99dfyNrREkRwzLMZ96MUYABojY1lTqh/CTUyZ9lBb4O+hRmG5\nrnUvGxv5VKYP+HgtwucMiL35N44cBZLScK/8ot9i3wUjpU/jAFud5sTc6ij2p6Do\njiLV643iiwY1hNeBC+5yNNyAs7JPRan4V3ZDrRFcr8VQRkyUhAVCCjsUL+DWzZG0\nZZfHV5HlAgMBAAECggEATzufYN2cjVb/dFnmjQAHkQl+W5JOUZlg7Hw6w1gpv3km\nsGnQI2v+UjSaraAC4xwYPMKzCd88xnFJKyMB5ST3N/vY+2xojadk4QNa1DGhciRs\nJ//dSWIXqRAtCwFtL7cRh5jB4KddHUCvrWtQIcOhgld1LBW0MsCNt5zG1lezPs/K\n0HvCXBsqgQfK6t4fKEftvEhpnustSRJTmYNdNu7VPnhskZ0yhW+ai0rjIeSFWKUJ\nogTTbyXdLfKI7jidCsMjC/439P6gLxG1sA4USiEvEJ0vBMc6LhOq3ZPAam0uoBiU\nJPkBv6Qnmaz36AqE94wi57t8y7/wTyRyPCVz1uZUXQKBgQDv67US1bjFGZ+NtKuV\nuiVUG1oq/jyNI2iZEK7UIQ53MRuIYm1vY5se6RyRjX0CvROkKttg+b57TjIJxQtg\nR7Pw4An+t2c/6kFmfE+ZfBd99082tUMj8NYT9rIyyNNLbUHBQSg5jKYx1Ipui5b9\nl5G3xTWFTyisQ1eWS1dfoqG5cwKBgQDAkNguSZisk4SblDy4wavoVddt1vUpLrF/\nEE37BWIQfUIMY+celQ487uhCaGd6TfTyKN1j1L+bXGYZSMD3PLOZjsahFnFzz/Yd\nyyvAXFC8wh4I6AFLdEGyOGbpndjLVsojHvgoKlzM65pZohRmF/UjCkREiFDY46vs\nMNrCAwGRRwKBgAJreS8QcsWiJYbTXKus7fV0NSub1taBlet9TJYdIz02hJWSkJIA\nCi98oGojaDBf4dPVDtNikXZC6qRIlX3KG1mBmPg1t1fgr/otpMvigYYmuWjO2TeC\nlVYsNc3nUqI/HtupIZO7BG4aO68zvzkdjz2wQoLusnVSVI7SgbYYONNbAoGAHGwE\njGhucWJgmzn48GgdgHBTGLI2gzqSFigI6njz7W0fd4azUS885793Zn3UcBfbhHaF\nnruDAGJV63tyh8tc0thg9tvCKQ7Ty7f2IDE/9WxNVJExx7pEDBMJFYgnvHN16FMk\n24DNK63GnV4v1pXIp8BXLSpVjH2mHXlkgGhDLfcCgYEA2OXwU59CpuUagp3vKgPH\nKGTejMA9M5fTSsdJDA7iHX3bmrn2Geph/zB/lZ4NY0bpAFhoADL0+ucCI79LsaTI\ngNc073BiRqqmrRgCcuOEwT6sx+h5gqG3JXum44XjXkb7AAARQzlAqWkoaDDi4Hyj\ncx/kLNBCBqi6CfRENdlTwUc=\n-----END PRIVATE KEY-----\n';

// ================================================================
// JWT & OAUTH2 — Geração de Token de Acesso
// ================================================================

/**
 * Gera um JWT assinado com a chave privada do Service Account.
 * @returns {string} JWT assinado (header.payload.signature)
 */
function gerarJwtFirebase_() {
  var header = {
    alg: 'RS256',
    typ: 'JWT'
  };

  var agora = Math.floor(Date.now() / 1000);
  var payload = {
    iss: FIREBASE_CLIENT_EMAIL_,
    sub: FIREBASE_CLIENT_EMAIL_,
    aud: 'https://oauth2.googleapis.com/token',
    iat: agora,
    exp: agora + 3600,
    scope: 'https://www.googleapis.com/auth/firebase.database https://www.googleapis.com/auth/userinfo.email'
  };

  var headerB64 = Utilities.base64EncodeWebSafe(JSON.stringify(header));
  var payloadB64 = Utilities.base64EncodeWebSafe(JSON.stringify(payload));
  var toSign = headerB64 + '.' + payloadB64;

  var signatureBytes = Utilities.computeRsaSha256Signature(toSign, FIREBASE_PRIVATE_KEY_);
  var signatureB64 = Utilities.base64EncodeWebSafe(signatureBytes);

  return toSign + '.' + signatureB64;
}

/**
 * Obtém um access_token OAuth2 trocando o JWT.
 * Cache de 50 minutos (token vale 60 min).
 * @returns {string} Access token válido
 */
function obterAccessTokenFirebase_() {
  var cache = CacheService.getScriptCache();
  var tokenCacheado = cache.get('FIREBASE_ACCESS_TOKEN');
  if (tokenCacheado) return tokenCacheado;

  var jwt = gerarJwtFirebase_();

  var response = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt
    },
    muteHttpExceptions: true
  });

  var resultado = JSON.parse(response.getContentText());

  if (resultado.error) {
    console.error('Erro ao obter token Firebase: ' + resultado.error_description);
    throw new Error('Falha na autenticação Firebase: ' + resultado.error_description);
  }

  // Cache por 50 minutos (token vale 60)
  cache.put('FIREBASE_ACCESS_TOKEN', resultado.access_token, 3000);
  return resultado.access_token;
}

// ================================================================
// REST API — Operações CRUD no Firebase RTDB
// ================================================================

/**
 * Sanitiza uma chave para ser válida no Firebase (remove . $ # [ ] /)
 * @param {string} key Chave original
 * @returns {string} Chave sanitizada
 */
function sanitizarChaveFirebase_(key) {
  if (!key) return 'sem_chave';
  return String(key).replace(/[.#$\[\]\/]/g, '_').replace(/\s+/g, '_').trim();
}

/**
 * Executa uma requisição REST ao Firebase RTDB.
 * @param {string} path Caminho no banco (ex: 'mandados/001_2026')
 * @param {string} method GET, PUT, PATCH, DELETE
 * @param {Object} [data] Dados para PUT/PATCH
 * @returns {Object|null} Resposta parseada ou null
 */
function firebaseRequest_(path, method, data) {
  try {
    var token = obterAccessTokenFirebase_();
    var url = FIREBASE_DATABASE_URL_ + '/' + path + '.json?access_token=' + token;

    var options = {
      method: method.toLowerCase(),
      contentType: 'application/json',
      muteHttpExceptions: true
    };

    if (data && (method === 'PUT' || method === 'PATCH' || method === 'POST')) {
      options.payload = JSON.stringify(data);
    }

    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();

    if (code >= 200 && code < 300) {
      var content = response.getContentText();
      return content ? JSON.parse(content) : null;
    } else {
      var errorMsg = 'Firebase REST erro ' + code + ': ' + response.getContentText();
      console.error(errorMsg);
      throw new Error(errorMsg);
    }
  } catch (e) {
    console.error('Firebase request falhou (' + method + ' ' + path + '): ' + e.message);
    throw e;
  }
}

/**
 * Escreve/substitui dados num caminho.
 */
function firebasePut_(path, data) {
  return firebaseRequest_(path, 'PUT', data);
}

/**
 * Atualiza parcialmente dados num caminho.
 */
function firebasePatch_(path, data) {
  return firebaseRequest_(path, 'PATCH', data);
}

/**
 * Lê dados de um caminho.
 */
function firebaseGet_(path) {
  return firebaseRequest_(path, 'GET');
}

/**
 * Remove dados de um caminho.
 */
function firebaseDelete_(path) {
  return firebaseRequest_(path, 'DELETE');
}

// ================================================================
// SYNC — Funções de Sincronização de Mandados
// ================================================================

/**
 * Converte um objeto mandado do formato Sheets para o formato Firebase.
 * @param {Object} mandadoObj Objeto com os dados do mandado
 * @returns {Object} Objeto limpo para gravar no Firebase
 */
function mandadoParaFirebase_(mandadoObj) {
  return {
    dataLancamento: mandadoObj.dataLancamento || '',
    dataConferencia: mandadoObj.dataConferencia || '',
    mandado: mandadoObj.mandado || '',
    artigo: mandadoObj.artigo || '',
    nome: mandadoObj.nome || '',
    cpf: mandadoObj.cpf || '',
    rg: mandadoObj.rg || '',
    nascimento: mandadoObj.nascimento || '',
    naturalidade: mandadoObj.naturalidade || '',
    sexo: mandadoObj.sexo || '',
    cor: mandadoObj.cor || '',
    filiacao: mandadoObj.filiacao || '',
    fotoUrl: mandadoObj.fotoUrl || '',
    batalhao: mandadoObj.batalhao || '',
    enderecoPrincipal: mandadoObj.enderecoPrincipal || '',
    outrosEnderecos: mandadoObj.outrosEnderecos || '',
    status: mandadoObj.status || 'Procurado',
    validade: mandadoObj.validade || '',
    infoProcessuais: mandadoObj.infoProcessuais || '',
    geodataSecundarios: mandadoObj.geodataSecundarios || '[]',
    dadosExtrasJSON: mandadoObj.dadosExtrasJSON || '{}',
    observacoes: mandadoObj.observacoes || '',
    historicoObservacoes: mandadoObj.historicoObservacoes || '[]',
    lat: mandadoObj.lat || null,
    lng: mandadoObj.lng || null,
    cpi: mandadoObj.cpi || mandadoObj.cpiArea || '',
    bpmArea: mandadoObj.bpmArea || '',
    ciaArea: mandadoObj.ciaArea || '',
    dpArea: mandadoObj.dpArea || '',
    cidade: mandadoObj.cidade || '',
    tipoImportacao: mandadoObj.tipoImportacao || '',
    idLinha: mandadoObj.idLinha || 0,
    semEndereco: mandadoObj.semEndereco || false,
    _ultimaAtualizacao: new Date().getTime()
  };
}

/**
 * Sincroniza um único mandado no Firebase (upsert).
 * @param {Object} mandadoObj Dados do mandado
 */
function sincronizarMandadoFirebase(mandadoObj) {
  try {
    if (!mandadoObj || !mandadoObj.mandado) return;
    var chave = sanitizarChaveFirebase_(mandadoObj.mandado);
    var dados = mandadoParaFirebase_(mandadoObj);
    firebasePut_('mandados/' + chave, dados);
    sinalizarMudancaFirebase_();
  } catch (e) {
    console.error('Erro ao sincronizar mandado no Firebase: ' + e.message);
  }
}

/**
 * Atualiza parcialmente um mandado no Firebase.
 * @param {string} mandadoNum Número do mandado
 * @param {Object} camposAtualizados Objeto com os campos a atualizar
 */
function atualizarMandadoFirebase(mandadoNum, camposAtualizados) {
  if (!mandadoNum) return;
  var chave = sanitizarChaveFirebase_(mandadoNum);
  camposAtualizados._ultimaAtualizacao = new Date().getTime();
  firebasePatch_('mandados/' + chave, camposAtualizados);
  sinalizarMudancaFirebase_();
}

/**
 * Remove um mandado do Firebase.
 * @param {string} mandadoNum Número do mandado
 */
function removerMandadoFirebase(mandadoNum) {
  try {
    if (!mandadoNum) return;
    var chave = sanitizarChaveFirebase_(mandadoNum);
    firebaseDelete_('mandados/' + chave);
    sinalizarMudancaFirebase_();
  } catch (e) {
    console.error('Erro ao remover mandado do Firebase: ' + e.message);
  }
}

/**
 * Atualiza o timestamp global de última modificação no Firebase.
 */
function sinalizarMudancaFirebase_() {
  try {
    firebasePatch_('metadata', {
      lastUpdate: new Date().getTime(),
      version: 'v5.5.0'
    });
  } catch (e) {
    console.error('Erro ao sinalizar mudança no Firebase: ' + e.message);
  }
}

// ================================================================
// MIGRAÇÃO — Transferência Completa Sheets → Firebase
// ================================================================

/**
 * Migra TODOS os mandados do Google Sheets para o Firebase.
 * Execute esta função UMA VEZ para popular o Firebase com os dados existentes.
 * Pode ser re-executada sem problemas (sobrescreve dados existentes).
 */
function migrarSheetsParaFirebase() {
  var inicio = new Date();
  console.log('🔥 Iniciando migração Sheets → Firebase...');

  try {
    // Testar conexão primeiro
    var teste = firebasePut_('metadata/migrationTest', { timestamp: new Date().getTime() });
    if (teste === null) {
      console.error('❌ Falha no teste de conexão com Firebase. Verifique se o Realtime Database está criado.');
      return { sucesso: false, mensagem: 'Falha na conexão com Firebase. Crie o Realtime Database no console.' };
    }

    // Obter todos os mandados via função existente
    var mandados = obterDados();
    console.log('📊 Total de mandados a migrar: ' + mandados.length);

    // Preparar batch (Firebase aceita escrita em lote via PUT no nó raiz)
    var batch = {};
    mandados.forEach(function(m) {
      var chave = sanitizarChaveFirebase_(m.mandado);
      batch[chave] = mandadoParaFirebase_(m);
    });

    // Gravar tudo de uma vez (mais eficiente que um por um)
    firebasePut_('mandados', batch);

    // Atualizar metadata
    firebasePut_('metadata', {
      lastUpdate: new Date().getTime(),
      totalMandados: mandados.length,
      migradoEm: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
      version: 'v5.5.0'
    });

    var duracao = ((new Date() - inicio) / 1000).toFixed(1);
    console.log('✅ Migração concluída! ' + mandados.length + ' mandados em ' + duracao + 's');

    return {
      sucesso: true,
      mensagem: 'Migração concluída! ' + mandados.length + ' mandados transferidos para o Firebase em ' + duracao + ' segundos.'
    };
  } catch (e) {
    console.error('❌ Erro na migração: ' + e.message);
    return { sucesso: false, mensagem: 'Erro na migração: ' + e.message };
  }
}

// ================================================================
// VERIFICAÇÃO — Integridade Sheets vs Firebase
// ================================================================

/**
 * Compara a quantidade de registros entre Sheets e Firebase.
 * @returns {Object} Relatório de integridade
 */
function verificarIntegridadeFirebase() {
  try {
    var mandadosSheets = obterDados();
    var mandadosFirebase = firebaseGet_('mandados');

    var totalSheets = mandadosSheets.length;
    var totalFirebase = mandadosFirebase ? Object.keys(mandadosFirebase).length : 0;

    var resultado = {
      totalSheets: totalSheets,
      totalFirebase: totalFirebase,
      sincronizado: totalSheets === totalFirebase,
      diferenca: Math.abs(totalSheets - totalFirebase)
    };

    console.log('📊 Integridade: Sheets=' + totalSheets + ', Firebase=' + totalFirebase +
      (resultado.sincronizado ? ' ✅ SINCRONIZADO' : ' ⚠️ DIFERENÇA: ' + resultado.diferenca));

    return resultado;
  } catch (e) {
    console.error('Erro ao verificar integridade: ' + e.message);
    return { erro: e.message };
  }
}

// ================================================================
// TESTE — Testar Conexão com Firebase
// ================================================================

/**
 * Testa a conexão com o Firebase RTDB.
 * Execute esta função no editor do Apps Script para validar.
 */
function testarConexaoFirebase() {
  console.log('🔥 Testando conexão com Firebase...');

  try {
    // 1. Testar obtenção de token
    var token = obterAccessTokenFirebase_();
    console.log('✅ Token obtido: ' + token.substring(0, 20) + '...');

    // 2. Testar escrita
    var resultado = firebasePut_('_teste', {
      mensagem: 'Conexão OK!',
      timestamp: new Date().getTime(),
      origem: 'Apps Script'
    });
    console.log('✅ Escrita: ' + JSON.stringify(resultado));

    // 3. Testar leitura
    var leitura = firebaseGet_('_teste');
    console.log('✅ Leitura: ' + JSON.stringify(leitura));

    // 4. Limpar teste
    firebaseDelete_('_teste');
    console.log('✅ Limpeza: nó de teste removido');

    console.log('🎉 Todos os testes passaram! Firebase está pronto.');
    return { sucesso: true, mensagem: 'Firebase conectado e funcionando!' };
  } catch (e) {
    console.error('❌ Falha no teste: ' + e.message);
    return { sucesso: false, mensagem: 'Erro: ' + e.message };
  }
}

// ================================================================
// CONFIG FRONTEND — Expor configuração do Firebase para o JS frontend
// ================================================================

/**
 * Retorna a configuração pública do Firebase para inicializar o SDK no frontend.
 * Chamada pelo frontend via google.script.run
 */
function getFirebaseConfig() {
  return {
    apiKey: 'AIzaSyBG0Jwp_S531tvTQxZqA9VudHvsoBPfXzw',
    authDomain: 'gen-lang-client-0691254724.firebaseapp.com',
    databaseURL: 'https://gen-lang-client-0691254724-default-rtdb.firebaseio.com',
    projectId: 'gen-lang-client-0691254724',
    storageBucket: 'gen-lang-client-0691254724.firebasestorage.app',
    messagingSenderId: '966277008479',
    appId: '1:966277008479:web:05fded49287e9dac7d9fea'
  };
}

/**
 * Registra um log de consumo de geocodificação no Firebase.
 * @param {string} endereco Endereço geocodificado
 * @param {string} tipo 'Principal' ou 'Secundario'
 * @param {string} fonte 'Google', 'Mapbox', 'Gemini'
 * @param {boolean} sucesso Se a geocodificação teve sucesso
 */
/**
 * Registra o consumo de uma requisição de geocodificação no Firebase e na planilha persistentemente.
 */
function registrarConsumoGeocodificacao(endereco, tipo, fonte, sucesso, mandado) {
  try {
    var hojeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var agoraTimestamp = new Date().getTime();
    var prov = fonte || "Google";
    var isSuc = sucesso === true;
    var mandadoStr = mandado ? String(mandado).trim() : "N/A";

    var logData = {
      timestamp: agoraTimestamp,
      data: hojeStr,
      mandado: mandadoStr,
      endereco: endereco || "Não informado",
      tipo: tipo || "Principal",
      fonte: prov,
      provedor: prov,
      sucesso: isSuc
    };

    // 1. Enviar para o Firebase RTDB
    try {
      firebaseRequest_('geocodificacao_consumo', 'POST', logData);
    } catch(errFb) {
      console.error("Erro ao registrar consumo no Firebase:", errFb);
    }

    // 2. Gravação persistente na planilha Geocodificacao_Consumo (Google Sheets)
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("Geocodificacao_Consumo");
    var headers = ["Timestamp", "Data", "Mandado", "Provedor", "Tipo", "Endereço", "Sucesso", "Usuario"];
    
    if (!aba) {
      aba = ss.insertSheet("Geocodificacao_Consumo");
      aba.getRange(1, 1, 1, headers.length).setValues([headers]);
      aba.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1e293b").setFontColor("#e2e8f0");
      aba.setFrozenRows(1);
    } else {
      var hData = aba.getRange(1, 1, 1, Math.max(aba.getLastColumn(), headers.length)).getValues()[0];
      var hasMandadoHeader = hData.some(function(h){ return String(h).toLowerCase() === 'mandado'; });
      if (!hasMandadoHeader) {
        aba.getRange(1, 1, 1, headers.length).setValues([headers]);
        aba.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1e293b").setFontColor("#e2e8f0");
      }
    }

    var emailUser = "";
    try {
      var user = verificarAcessoUsuario();
      if (user && user.email) emailUser = user.email;
    } catch(e){}

    aba.appendRow([agoraTimestamp, hojeStr, mandadoStr, prov, tipo || "Principal", endereco || "Não informado", isSuc ? "SIM" : "NAO", emailUser]);
  } catch(e) {
    console.error("Erro ao registrar consumo de geocodificação:", e);
  }
}

/**
 * Retorna todo o histórico de logs de geocodificação gravados na planilha Geocodificacao_Consumo
 * (com fallback para Firebase ou inicialização automática a partir de mandados se vazia).
 */
function obterHistoricoConsumo() {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado) return [];

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("Geocodificacao_Consumo");
    
    if (!aba || aba.getLastRow() <= 1) {
      // 1. Tentar ler do Firebase se disponível
      var fbData = firebaseRequest_('geocodificacao_consumo', 'GET');
      var logsFb = [];
      if (fbData && typeof fbData === 'object') {
        Object.keys(fbData).forEach(function(k) {
          if (fbData[k]) {
            var item = fbData[k];
            item.mandado = item.mandado || "N/A";
            item.provedor = item.provedor || item.fonte || "Google";
            item.sucesso = item.sucesso === true;
            logsFb.push(item);
          }
        });
      }

      if (logsFb.length > 0) return logsFb;

      // 2. Se Firebase estiver vazio e houver mandados ativos, inicializa histórico inicial
      inicializarHistoricoConsumo();
      
      aba = ss.getSheetByName("Geocodificacao_Consumo");
      if (!aba || aba.getLastRow() <= 1) return [];
    }

    var data = aba.getDataRange().getValues();
    if (data.length <= 1) return [];

    var colMap = {};
    var hRow = data[0];
    for (var c = 0; c < hRow.length; c++) {
      colMap[String(hRow[c]).trim().toLowerCase()] = c;
    }

    var idxTs = colMap['timestamp'] !== undefined ? colMap['timestamp'] : 0;
    var idxData = colMap['data'] !== undefined ? colMap['data'] : 1;
    var idxMandado = colMap['mandado'] !== undefined ? colMap['mandado'] : -1;
    var idxProv = colMap['provedor'] !== undefined ? colMap['provedor'] : (colMap['fonte'] !== undefined ? colMap['fonte'] : 2);
    var idxTipo = colMap['tipo'] !== undefined ? colMap['tipo'] : 3;
    var idxEnd = colMap['endereço'] !== undefined ? colMap['endereço'] : (colMap['endereco'] !== undefined ? colMap['endereco'] : 4);
    var idxSuc = colMap['sucesso'] !== undefined ? colMap['sucesso'] : 5;
    var idxUser = colMap['usuario'] !== undefined ? colMap['usuario'] : 6;

    var logs = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[idxTs] || row[idxData]) {
        var provVal = String(row[idxProv] || 'Google');
        var sucVal = row[idxSuc] === true || String(row[idxSuc]).toUpperCase() === 'SIM' || String(row[idxSuc]).toUpperCase() === 'TRUE';
        logs.push({
          timestamp: row[idxTs],
          data: String(row[idxData] || ''),
          mandado: idxMandado >= 0 ? String(row[idxMandado] || 'N/A') : 'N/A',
          provedor: provVal,
          fonte: provVal,
          tipo: String(row[idxTipo] || 'Principal'),
          endereco: String(row[idxEnd] || ''),
          sucesso: sucVal,
          usuario: String(row[idxUser] || '')
        });
      }
    }
    return logs;
  } catch(e) {
    console.error("Erro ao obter histórico de consumo:", e);
    return [];
  }
}

/**
 * Inicializa o histórico de geocodificação no Firebase e na aba Geocodificacao_Consumo.
 */
function inicializarHistoricoConsumo() {
  var dados = obterDados();
  if (!dados || dados.length === 0) return "Nenhum mandado na planilha.";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("Geocodificacao_Consumo");
  if (!aba) {
    aba = ss.insertSheet("Geocodificacao_Consumo");
    var headers = ["Timestamp", "Data", "Mandado", "Provedor", "Tipo", "Endereço", "Sucesso", "Usuario"];
    aba.getRange(1, 1, 1, headers.length).setValues([headers]);
    aba.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1e293b").setFontColor("#e2e8f0");
    aba.setFrozenRows(1);
  }

  var contador = 0;
  var novasLinhas = [];

  dados.forEach(function(d) {
    var dataStr = d.dataLancamento || "";
    var dataIso = "";
    if (dataStr && dataStr.split('/').length === 3) {
      var p = dataStr.split('/');
      dataIso = p[2] + "-" + p[1] + "-" + p[0];
    } else {
      dataIso = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    // Principal
    if (d.lat && d.lng && d.enderecoPrincipal) {
      var ts = new Date().getTime();
      var logData = {
        timestamp: ts,
        data: dataIso,
        endereco: d.enderecoPrincipal,
        tipo: "Principal",
        fonte: "Google",
        provedor: "Google",
        sucesso: true
      };
      try { firebaseRequest_('geocodificacao_consumo', 'POST', logData); } catch(e){}
      novasLinhas.push([ts, dataIso, "Google", "Principal", d.enderecoPrincipal, "SIM", "Sistema"]);
      contador++;
    }

    // Secundários
    var secLista = [];
    try { secLista = JSON.parse(d.geodataSecundarios || "[]"); } catch(e){}
    if (secLista && secLista.length > 0) {
      secLista.forEach(function(sec) {
        if (sec.lat && sec.lng && sec.endereco) {
          var ts = new Date().getTime();
          var logData = {
            timestamp: ts,
            data: dataIso,
            endereco: sec.endereco,
            tipo: "Secundario",
            fonte: "Google",
            provedor: "Google",
            sucesso: true
          };
          try { firebaseRequest_('geocodificacao_consumo', 'POST', logData); } catch(e){}
          novasLinhas.push([ts, dataIso, "Google", "Secundario", sec.endereco, "SIM", "Sistema"]);
          contador++;
        }
      });
    }
  });

  if (novasLinhas.length > 0) {
    aba.getRange(aba.getLastRow() + 1, 1, novasLinhas.length, 7).setValues(novasLinhas);
    SpreadsheetApp.flush();
  }

  return "Inicialização concluída. Registrados " + contador + " logs de geocodificação históricos.";
}
