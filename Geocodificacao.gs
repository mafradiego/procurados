// ================================================================
// SENTINELA v2.0 — MÓDULO DE GEOCODIFICAÇÃO
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
    if (!geoJsonStr) continue;

    try {
      var geo = JSON.parse(geoJsonStr);
      if (!geo || !geo.coordinates || !geo.coordinates[0]) continue;

      // Converter GeoJSON [lng, lat] para array de {lat, lng}
      var coords = geo.coordinates[0].map(function(c) {
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

  // Limpar cache para forçar leitura fresca dos polígonos
  _cachePoligonos = null;

  const dataRange = sheet.getDataRange();
  const numRows = dataRange.getNumRows();
  let atualizados = 0;

  for (let i = 2; i <= numRows; i++) {
    if (executarGeocodificacaoLinha(sheet, i, true)) {
      atualizados++;
    }
  }
  SpreadsheetApp.getUi().alert("Varredura Concluída!\n" + atualizados + " mandados foram processados.");
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
    }
  }

  if (pinoNaSede) {
    const coords = COORDENADAS_SEDES[batalhao] || COORDENADAS_SEDES["DEFAULT"];
    lat = coords.lat;
    lng = coords.lng;
  } else {
    try {
      const geocoder = Maps.newGeocoder();
      const response = geocoder.geocode(enderecoPrincipal);

      if (response.status === 'OK' && response.results.length > 0) {
        lat = response.results[0].geometry.location.lat;
        lng = response.results[0].geometry.location.lng;
      } else {
        const coords = COORDENADAS_SEDES[batalhao] || COORDENADAS_SEDES["DEFAULT"];
        lat = coords.lat;
        lng = coords.lng;
        mensagemObs = "Endereço principal não localizado no mapa.";
      }
    } catch (erro) {
      if (isBatchMode) Utilities.sleep(2000);
      return false;
    }
    if (isBatchMode) Utilities.sleep(1000);
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
