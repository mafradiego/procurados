// ================================================================
// SENTINELA v4.2.0 — MÓDULO DE GEOCODIFICAÇÃO
// Motor: Gemini Map Grounding + Fallback Sedes Táticas
// ================================================================

const COORDENADAS_SEDES = {
  "08ª BPM/I": { lat: -22.9056, lng: -47.0608 },
  "11ª BPM/I": { lat: -23.1857, lng: -46.8978 },
  "26ª BPM/I": { lat: -22.4332, lng: -46.9476 },
  "34ª BPM/I": { lat: -22.9056, lng: -47.0608 },
  "35ª BPM/I": { lat: -22.9056, lng: -47.0608 },
  "47ª BPM/I": { lat: -22.9056, lng: -47.0608 },
  "48ª BPM/I": { lat: -22.9056, lng: -47.0608 },
  "49ª BPM/I": { lat: -22.9056, lng: -47.0608 },
  "DEFAULT": { lat: -22.9056, lng: -47.0608 }
};

const UFS_BRASIL = "AC|AL|AP|AM|BA|CE|DF|ES|GO|MA|MT|MS|MG|PA|PB|PR|PE|PI|RJ|RN|RS|RO|RR|SC|SP|SE|TO";

// ================================================================
// MOTOR DE GEOCODIFICAÇÃO VIA GEMINI MAP GROUNDING (v4.2.0)
// Substitui o Google Maps Geocoder clássico.
// Usa a ferramenta google_maps do Gemini para resolver endereços
// informais e incompletos com precisão superior.
// Cota: 500 RPD (Request Per Day) / 15 RPM (Request Per Minute).
// ================================================================

/**
 * Geocodifica um endereço usando Gemini 3.1 Flash Lite com Map Grounding.
 * Retorna { lat, lng } ou null se não localizar.
 *
 * @param {string} endereco O endereço a geocodificar.
 * @return {{ lat: number, lng: number } | null}
 */
function geocodificarComGeminiGrounding(endereco) {
  if (!endereco || endereco.length < 5) return null;

  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) {
    Logger.log("[GEO-GROUNDING] GEMINI_API_KEY não configurada.");
    return null;
  }

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key=" + apiKey;

  const payload = {
    "contents": [{
      "parts": [{
        "text": "Retorne APENAS as coordenadas geográficas (latitude e longitude) do seguinte endereço brasileiro em JSON puro, sem explicações: " + endereco
      }]
    }],
    "tools": [{
      "google_maps": {}
    }],
    "generationConfig": {
      "responseMimeType": "application/json",
      "responseSchema": {
        "type": "OBJECT",
        "properties": {
          "latitude": { "type": "NUMBER" },
          "longitude": { "type": "NUMBER" }
        },
        "required": ["latitude", "longitude"]
      }
    }
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    // CADÊNCIA v4.2.0: 4.2s de delay para respeitar teto de 15 RPM
    Utilities.sleep(4200);

    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var content = response.getContentText();

    // Retry para erros temporários (429/503)
    if (code === 429 || code === 503) {
      Logger.log("[GEO-GROUNDING] Erro " + code + " — aguardando 10s para retry...");
      Utilities.sleep(10000);
      response = UrlFetchApp.fetch(url, options);
      code = response.getResponseCode();
      content = response.getContentText();
    }

    if (code !== 200) {
      Logger.log("[GEO-GROUNDING] Erro HTTP " + code + ": " + content.substring(0, 200));
      return null;
    }

    var parsed = JSON.parse(content);
    if (!parsed.candidates || !parsed.candidates[0] || !parsed.candidates[0].content ||
        !parsed.candidates[0].content.parts || parsed.candidates[0].content.parts.length === 0) {
      Logger.log("[GEO-GROUNDING] Resposta vazia do Gemini para: " + endereco);
      return null;
    }

    var jsonStr = parsed.candidates[0].content.parts[0].text;
    var coords = JSON.parse(jsonStr);

    if (coords.latitude && coords.longitude &&
        !isNaN(coords.latitude) && !isNaN(coords.longitude) &&
        Math.abs(coords.latitude) > 0.1 && Math.abs(coords.longitude) > 0.1) {
      Logger.log("[GEO-GROUNDING] Sucesso: " + endereco + " → " + coords.latitude + ", " + coords.longitude);
      incrementarCotaGemini_("GEOCODIFICACAO");
      return { lat: coords.latitude, lng: coords.longitude };
    }

    Logger.log("[GEO-GROUNDING] Coordenadas inválidas retornadas para: " + endereco);
    return null;

  } catch (e) {
    var msg = e.message ? e.message.toLowerCase() : "";
    if (msg.includes("quota") || msg.includes("limit") || msg.includes("429")) {
      throw new Error("QUOTA_LIMIT");
    }
    Logger.log("[GEO-GROUNDING] Erro ao geocodificar: " + e.message);
    return null;
  }
}

// ================================================================
// VALIDAÇÃO DE CIDADES E CEP (Economia de Cota)
// ================================================================

var _cacheCidades = null;
function carregarCidadesCache_() {
  if (_cacheCidades) return _cacheCidades;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cidades");
  if (!sheet) {
    _cacheCidades = [];
    return [];
  }
  var data = sheet.getDataRange().getValues();
  var cidades = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) cidades.push(data[i][0].toString().trim().toLowerCase());
  }
  _cacheCidades = cidades;
  return cidades;
}

function validarEnderecoAntesDeGeocodificar_(endereco) {
  if (!endereco) return false;
  // Se tiver CEP explícito, permite
  if (/\b\d{5}-?\d{3}\b/.test(endereco) || /\bCEP\s*\d+\b/i.test(endereco)) return true;
  
  var cidades = carregarCidadesCache_();
  // Se a aba Cidades não existir ou estiver vazia, ignora a validação restrita
  if (cidades.length === 0) return true;
  
  var endLower = endereco.toLowerCase();
  for (var i = 0; i < cidades.length; i++) {
    // Tratar regex seguro para nome da cidade
    var cidRegex = cidades[i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    var regex = new RegExp("\\b" + cidRegex + "\\b", "i");
    if (regex.test(endLower)) return true;
  }
  return false;
}

// ================================================================
// IDENTIFICAÇÃO DE ÁREA POR COORDENADAS (Point-in-Polygon)
// ================================================================

/**
 * Cache dos polígonos para evitar leituras repetidas durante batch.
 */
var _cachePoligonos = null;

function carregarPoligonosCache_() {
  if (_cachePoligonos) return _cachePoligonos;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Poligonos");
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues();
  var poligonos = [];

  for (var i = 1; i < data.length; i++) {
    var ativo = String(data[i][7] || "").trim().toUpperCase();
    if (ativo !== "SIM") continue;

    var geoJsonStr = String(data[i][5] || "");
    for (var col = 8; col < data[i].length; col++) {
      if (data[i][col]) geoJsonStr += String(data[i][col]);
    }
    if (!geoJsonStr) continue;

    try {
      var geo = JSON.parse(geoJsonStr);
      if (!geo || !geo.coordinates || !geo.coordinates[0]) continue;

      var outerRings = [];

      if (geo.type === "MultiPolygon") {
        geo.coordinates.forEach(function(polygon) {
          if (polygon.length > 0) outerRings.push(polygon[0]);
        });
      } else if (geo.type === "Polygon") {
        if (geo.coordinates.length > 0) outerRings.push(geo.coordinates[0]);
      }

      if (outerRings.length === 0) continue;

      outerRings.forEach(function(ring) {
        var coords = ring.map(function(c) {
          return { lat: c[1], lng: c[0] };
        });

        poligonos.push({
          cpi: String(data[i][0] || ""),
          batalhao: String(data[i][1] || ""),
          cia: String(data[i][2] || ""),
          cidade: String(data[i][3] || ""),
          delegacia: String(data[i][4] || ""),
          coords: coords
        });
      });
    } catch (e) {
      // Polígono inválido, pular
    }
  }

  _cachePoligonos = poligonos;
  return poligonos;
}

/**
 * Algoritmo Ray Casting — verifica se um ponto está dentro de um polígono.
 * @param {number} lat - Latitude do ponto
 * @param {number} lng - Longitude do ponto
 * @param {Array} polygon - Array de {lat, lng}
 * @returns {boolean}
 */
function pontoNoPoligono_(lat, lng, polygon) {
  var inside = false;
  var n = polygon.length;

  for (var i = 0, j = n - 1; i < n; j = i++) {
    var yi = polygon[i].lat, xi = polygon[i].lng;
    var yj = polygon[j].lat, xj = polygon[j].lng;

    if (((yi > lat) !== (yj > lat)) &&
        (lng < (xj - xi) * (lat - yi) / (yj - yi) + xi)) {
      inside = !inside;
    }
  }

  return inside;
}

/**
 * Identifica em qual área (BPM/CIA/DP) um ponto com lat/lng se encontra.
 * @param {number} lat - Latitude
 * @param {number} lng - Longitude
 * @returns {Object|null} - {cpi, batalhao, cia, delegacia} ou null
 */
function identificarAreaPorCoordenadas(lat, lng) {
  if (!lat || !lng || isNaN(lat) || isNaN(lng)) return null;

  var poligonos = carregarPoligonosCache_();
  if (poligonos.length === 0) return null;

  for (var i = 0; i < poligonos.length; i++) {
    if (pontoNoPoligono_(lat, lng, poligonos[i].coords)) {
      return {
        cpi: poligonos[i].cpi,
        batalhao: poligonos[i].batalhao,
        cia: poligonos[i].cia,
        delegacia: poligonos[i].delegacia,
        cidade: poligonos[i].cidade
      };
    }
  }

  return null;
}

/**
 * Grava os dados de área nas colunas Y-AB de uma linha da planilha.
 * @param {Sheet} sheet - A aba Mandados
 * @param {number} linhaIdx - Índice da linha (1-based)
 * @param {number} lat - Latitude
 * @param {number} lng - Longitude
 */
function gravarAreaNaLinha_(sheet, linhaIdx, lat, lng) {
  var area = identificarAreaPorCoordenadas(lat, lng);
  if (area) {
    sheet.getRange(linhaIdx, 25).setValue(area.cpi);        // Y = CPI
    sheet.getRange(linhaIdx, 26).setValue(area.batalhao);    // Z = BPM_Area
    sheet.getRange(linhaIdx, 27).setValue(area.cia);         // AA = CIA_Area
    sheet.getRange(linhaIdx, 28).setValue(area.delegacia);   // AB = DP_Area
    sheet.getRange(linhaIdx, 29).setValue(area.cidade);      // AC = Cidade
    if (area.batalhao) {
      sheet.getRange(linhaIdx, 14).setValue(area.batalhao);  // N = Batalhão Principal
    }
  }
}

// ================================================================
// GEOCODIFICAÇÃO
// ================================================================

/**
 * Trigger onEdit — geocodifica ao editar a coluna F.
 */
function geocodificarAoEditar(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Mandados") return;

  const colInicio = e.range.getColumn();
  const colFim = e.range.getLastColumn();
  const linhaInicio = e.range.getRow();
  const linhaFim = e.range.getLastRow();

  if (colInicio > 15 || colFim < 15) return;
  if (linhaInicio === 1 && linhaFim === 1) return;

  for (let i = Math.max(2, linhaInicio); i <= linhaFim; i++) {
    executarGeocodificacaoLinha(sheet, i, false);
  }
}

/**
 * Processamento batch de todos os endereços pendentes.
 */
function processarPendentes() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Aba 'Mandados' não encontrada.");
    return;
  }

  // Limpar cache para forçar leitura fresca dos polígonos e cidades
  _cachePoligonos = null;
  _cacheCidades = null;

  const dataRange = sheet.getDataRange();
  const numRows = dataRange.getNumRows();
  let atualizados = 0;
  
  const startTime = Date.now();
  const TIME_LIMIT = 4.5 * 60 * 1000; // 4.5 minutos
  let lastPropCheck = Date.now();
  
  // Limpar flag antiga se houver (para permitir que ele rode de novo se o usuário der play)
  PropertiesService.getScriptProperties().deleteProperty('STOP_GEO');

  for (let i = 2; i <= numRows; i++) {
    // A cada 5 segundos, verifica se o usuário apertou o botão "Parar"
    if (Date.now() - lastPropCheck > 5000) {
      if (PropertiesService.getScriptProperties().getProperty('STOP_GEO') === 'true') {
        PropertiesService.getScriptProperties().deleteProperty('STOP_GEO');
        try { SpreadsheetApp.getUi().alert("Processo cancelado pelo usuário!"); } catch(e){}
        return;
      }
      lastPropCheck = Date.now();
    }
    try {
      if (executarGeocodificacaoLinha(sheet, i, true)) {
        atualizados++;
      }
    } catch (erro) {
      if (erro.message === "QUOTA_LIMIT") {
        pararGeocodificacaoAutomatica(true);
        try {
          SpreadsheetApp.getUi().alert("⛔ LIMITE DIÁRIO DO GOOGLE MAPS ATINGIDO!\nO Google permite cerca de 1.000 buscas de endereço por dia.\n\n" + atualizados + " mandados foram processados com sucesso agora.\nAguarde 24h para processar o resto.");
        } catch(e){}
        return;
      }
    }
    
    // Evitar erro de limite de tempo de execução do Google (6 minutos)
    if (Date.now() - startTime > TIME_LIMIT) {
      // Criar o gatilho para continuar rodando automaticamente em 30 segundos
      ScriptApp.newTrigger('processarPendentes').timeBased().after(30 * 1000).create();
      
      try {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          atualizados + " mandados processados. Retomando automaticamente em 30 segundos...",
          "⏱️ Pausa de Segurança", 
          10
        );
      } catch (e) {
        // Ignora erro caso esteja rodando por trigger (UI indisponível)
      }
      return;
    }
  }
  
  pararGeocodificacaoAutomatica(true); // Deleta triggers anteriores porque já acabou tudo
  try {
    SpreadsheetApp.getUi().alert("Varredura Total Concluída!\n" + atualizados + " mandados foram processados na rodada final.");
  } catch (e) {}
}

/**
 * Interrompe qualquer execução engatilhada pendente.
 */
function pararGeocodificacaoAutomatica(silencioso) {
  PropertiesService.getScriptProperties().setProperty('STOP_GEO', 'true');
  
  const triggers = ScriptApp.getProjectTriggers();
  let deletados = 0;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processarPendentes') {
      ScriptApp.deleteTrigger(triggers[i]);
      deletados++;
    }
  }
  
  if (!silencioso) {
    try {
      SpreadsheetApp.getUi().alert("Sinal de parada enviado!\nSe o sistema estava rodando, ele vai abortar imediatamente.\nTriggers pendentes cancelados: " + deletados);
    } catch (e) {}
  }
}

/**
 * Motor central de geocodificação (usado por ambos os modos).
 */
function executarGeocodificacaoLinha(sheet, linhaIdx, isBatchMode) {
  const batalhao = sheet.getRange(linhaIdx, 14).getValue();  // N = Batalhao
  let textoCompleto = sheet.getRange(linhaIdx, 15).getValue(); // O = Endereco Principal
  const latAtual = sheet.getRange(linhaIdx, 23).getValue();   // W = Latitude

  if (!textoCompleto || latAtual) return false;

  let lat = null;
  let lng = null;
  let pinoNaSede = false;
  let mensagemObs = "";
  let enderecoPrincipal = "";
  let outrosEnderecos = "";

  if (textoCompleto.toString().toLowerCase().includes("não sabe informar")) {
    pinoNaSede = true;
    mensagemObs = "Endereço não informado.";
    enderecoPrincipal = "Não informado";
  } else {
    textoCompleto = textoCompleto.toString().replace(/^Endereços\s*/i, "").trim();
    textoCompleto = textoCompleto.replace(/(Telefone|Celular).*?(?=\s[A-Z]|$)/ig, "").trim();

    const regexCorte = new RegExp(`(.*?\\b(?:${UFS_BRASIL})\\b(?:\\s*-\\s*CEP:\\s*\\d+|\\s*CEP\\s*\\d+-?\\d*|\\s*\\d{5}-?\\d{3}|\\s*\\d{8})?)(.*)`, "i");
    const match = textoCompleto.match(regexCorte);

    if (match) {
      enderecoPrincipal = match[1].trim();
      outrosEnderecos = match[2].replace(/^[\s\/,\-]+/, '').trim();
    } else {
      enderecoPrincipal = textoCompleto;
    }

    if (enderecoPrincipal.length < 5) {
      pinoNaSede = true;
      mensagemObs = "Endereço inválido.";
    } else if (!validarEnderecoAntesDeGeocodificar_(enderecoPrincipal)) {
      pinoNaSede = true;
      mensagemObs = "Cidade não coberta na tabela Cidades e sem CEP. Pino na Sede.";
    }
  }

  if (pinoNaSede) {
    const coords = COORDENADAS_SEDES[batalhao] || COORDENADAS_SEDES["DEFAULT"];
    lat = coords.lat;
    lng = coords.lng;
  } else {
    try {
      // v4.2.0: Motor de geocodificação via Gemini Map Grounding
      const geoResult = geocodificarComGeminiGrounding(enderecoPrincipal);

      if (geoResult) {
        lat = geoResult.lat;
        lng = geoResult.lng;
      } else {
        const coords = COORDENADAS_SEDES[batalhao] || COORDENADAS_SEDES["DEFAULT"];
        lat = coords.lat;
        lng = coords.lng;
        mensagemObs = "Endereço principal não localizado pela IA. Pino na sede.";
      }
    } catch (erro) {
      const msg = erro.message ? erro.message.toLowerCase() : "";
      if (msg === "quota_limit" || msg.includes("quota") || msg.includes("limit")) {
        throw new Error("QUOTA_LIMIT");
      }
      
      const coords = COORDENADAS_SEDES[batalhao] || COORDENADAS_SEDES["DEFAULT"];
      lat = coords.lat;
      lng = coords.lng;
      mensagemObs = "Erro Geo IA: " + (erro.message ? erro.message.substring(0, 50) : "Desconhecido") + " Pino Sede.";
    }
  }

  if (lat && lng) {
    sheet.getRange(linhaIdx, 15).setValue(enderecoPrincipal);  // O = Endereco Principal
    sheet.getRange(linhaIdx, 16).setValue(outrosEnderecos);    // P = Outros Enderecos
    sheet.getRange(linhaIdx, 23).setValue(lat);                // W = Latitude
    sheet.getRange(linhaIdx, 24).setValue(lng);                // X = Longitude
    sheet.getRange(linhaIdx, 17).setValue("Procurado");        // Q = Status

    // Identificar área (BPM/CIA/DP) pelas coordenadas
    gravarAreaNaLinha_(sheet, linhaIdx, lat, lng);

    if (mensagemObs !== "") {
      const obsAtual = sheet.getRange(linhaIdx, 22).getValue(); // V = Observacoes
      const novaObs = obsAtual ? obsAtual + " | " + mensagemObs : mensagemObs;
      sheet.getRange(linhaIdx, 22).setValue(novaObs);
    }

    // Se não tiver data de lançamento, inserir agora
    const dataLanc = sheet.getRange(linhaIdx, 1).getValue();  // A = Data de Lancamento
    if (!dataLanc) {
      const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
      sheet.getRange(linhaIdx, 1).setValue(dataAtual);
    }

    return true;
  }
  return false;
}
