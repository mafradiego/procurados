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
 *//**
 * ETAPA 1: Usa o Gemini para sanitizar e padronizar o texto do endereço extraído do PDF.
 * Não calcula coordenadas para evitar alucinação de mapas.
 * 
 * @param {string} enderecoBruto Texto bruto do endereço.
 * @return {string|null} Endereço formatado e limpo ou null se inválido.
 * @private
 */
function padronizarEnderecoComGemini_(enderecoBruto) {
  if (!enderecoBruto || enderecoBruto.length < 5) return null;

  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) {
    Logger.log("[GEO-PADRONIZAR] GEMINI_API_KEY não configurada.");
    return null;
  }

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key=" + apiKey;

  const prompt = `Você é um especialista em geolocalização militar e análise de dados.
Sua única tarefa é analisar o texto bruto abaixo (extraído de um mandado de prisão) e convertê-lo em um endereço limpo, completo e padronizado para busca no Google Maps.

REGRAS RÍGIDAS:
1. Formato esperado: "Logradouro, Número - Bairro, Cidade - UF, CEP" (ex: "Rua das Flores, 123 - Centro, Campinas - SP, 13010-000").
2. Remova termos inúteis como "Telefone", "Celular", "Fixo", "Desconhecido", "Residencial", "Aos cuidados de".
3. Se o texto indicar expressamente que a pessoa não sabe o endereço, ou for completamente ininteligível, responda APENAS "INVALIDO".
4. Responda APENAS o texto do endereço limpo. Não adicione saudações, explicações ou marcadores markdown.

Texto bruto do documento:
"${enderecoBruto}"`;

  const payload = {
    "contents": [{
      "parts": [{ "text": prompt }]
    }]
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    // Manutenção da cadência do sistema (460ms)
    Utilities.sleep(460);

    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var content = response.getContentText();

    // Retry automático para instabilidades do Gemini (429/503)
    if (code === 429 || code === 503) {
      Logger.log("[GEO-PADRONIZAR] Erro temporário " + code + " — aguardando 5s...");
      Utilities.sleep(5000);
      response = UrlFetchApp.fetch(url, options);
      code = response.getResponseCode();
      content = response.getContentText();
    }

    if (code !== 200) {
      Logger.log("[GEO-PADRONIZAR] Erro HTTP " + code + ": " + content.substring(0, 150));
      return null;
    }

    var parsed = JSON.parse(content);
    if (!parsed.candidates || !parsed.candidates[0] || !parsed.candidates[0].content ||
        !parsed.candidates[0].content.parts || parsed.candidates[0].content.parts.length === 0) {
      return null;
    }

    var textoLimpo = parsed.candidates[0].content.parts[0].text.trim();
    textoLimpo = textoLimpo.replace(/```/g, '').trim();

    if (textoLimpo.toUpperCase().includes("INVALIDO") || textoLimpo.length < 5) {
      Logger.log("[GEO-PADRONIZAR] Endereço rejeitado pela IA: " + enderecoBruto);
      return null;
    }

    Logger.log("[GEO-PADRONIZAR] Sucesso na sanitização: '" + enderecoBruto + "' → '" + textoLimpo + "'");
    return textoLimpo;

  } catch (e) {
    Logger.log("[GEO-PADRONIZAR] Erro na requisição ao Gemini: " + e.message);
    return null;
  }
}

/**
 * ETAPA 2: Geocodificação Segura de Fallback.
 * Sanitiza o endereço via Gemini e depois calcula a coordenada REAL via Google Maps Geocoder nativo.
 * Mantém o mesmo nome e assinatura para não quebrar chamadas legadas do sistema.
 * 
 * @param {string} endereco O endereço bruto a geocodificar.
 * @return {{ lat: number, lng: number, enderecoFormatado: string, precisao: string } | null}
 */
function geocodificarComGeminiGrounding(endereco) {
  if (!endereco || endereco.length < 5) return null;

  // 1. Usa a IA para limpar e padronizar o texto do endereço
  var enderecoSanitizado = padronizarEnderecoComGemini_(endereco);
  if (!enderecoSanitizado) return null;

  // 2. Envia o endereço limpo para o motor Geocoder REAL do Google Apps Script
  try {
    var geocoder = Maps.newGeocoder();
    var response = geocoder.geocode(enderecoSanitizado);

    if (response.status === 'OK' && response.results && response.results.length > 0) {
      var res = response.results[0];
      var locType = res.geometry.location_type;

      // Rejeita resultados imprecisos (centro geométrico de cidade/bairro)
      if (locType === "APPROXIMATE" || locType === "GEOMETRIC_CENTER") {
        Logger.log("[GEO-GROUNDING] Endereço padronizado retornado com precisão insuficiente (" + locType + "): " + enderecoSanitizado);
        return null;
      }

      Logger.log("[GEO-GROUNDING] Sucesso Real: " + enderecoSanitizado + " → " + res.geometry.location.lat + ", " + res.geometry.location.lng + " [" + locType + "]");
      
      // Incrementa a cota diária de geocodificação por IA
      incrementarCotaGemini_("GEOCODIFICACAO");

      return {
        lat: res.geometry.location.lat,
        lng: res.geometry.location.lng,
        enderecoFormatado: res.formatted_address || enderecoSanitizado,
        precisao: locType
      };
    } else {
      Logger.log("[GEO-GROUNDING] Geocoder nativo não localizou o endereço sanitizado: " + enderecoSanitizado);
    }
  } catch (e) {
    var msg = e.message ? e.message.toLowerCase() : "";
    if (msg.includes("quota") || msg.includes("limit") || msg.includes("429")) {
      throw new Error("QUOTA_LIMIT");
    }
    Logger.log("[GEO-GROUNDING] Erro ao calcular coordenadas nativas: " + e.message);
  }

  return null;
}

/**
 * Busca coordenadas pré-existentes na base de dados para evitar consumo de requisições de API de geocodificação.
 */
function buscarCoordenadasEmCache_(endereco) {
  if (!endereco || endereco.length < 5) return null;
  var normBuscada = normalizarTextoParaComparacao_(endereco);
  if (!normBuscada) return null;

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetMandados = ss.getSheetByName("Mandados");
    if (sheetMandados && sheetMandados.getLastRow() > 1) {
      var col = obterMapaColunas(sheetMandados);
      var idxEnd = col['Endereco Principal'] !== undefined ? col['Endereco Principal'] : col['Endereço Principal'];
      var idxLat = col['Latitude'];
      var idxLng = col['Longitude'];
      var idxBpm = col['Batalhao'] !== undefined ? col['Batalhao'] : col['Batalhão'];

      if (idxEnd !== undefined && idxLat !== undefined && idxLng !== undefined) {
        var dadosM = sheetMandados.getDataRange().getValues();
        for (var i = 1; i < dadosM.length; i++) {
          var row = dadosM[i];
          var endBase = row[idxEnd];
          var latVal = parseFloat(row[idxLat]);
          var lngVal = parseFloat(row[idxLng]);

          if (endBase && !isNaN(latVal) && !isNaN(lngVal) && latVal !== 0 && lngVal !== 0) {
            var normBase = normalizarTextoParaComparacao_(endBase);
            if (normBase === normBuscada || (normBase.length > 10 && normBuscada.length > 10 && normBase.includes(normBuscada))) {
              Logger.log("[GEO-CACHE] Coordenadas reutilizadas do banco local para: " + endereco + " → (" + latVal + ", " + lngVal + ")");
              return {
                lat: latVal,
                lng: lngVal,
                enderecoFormatado: String(endBase),
                fonte: "Cache Local",
                precisao: "ROOFTOP",
                batalhao: idxBpm !== undefined ? String(row[idxBpm] || "") : ""
              };
            }
          }
        }
      }
    }
  } catch (e) {
    Logger.log("[GEO-CACHE] Erro ao consultar cache local: " + e.message);
  }
  return null;
}

/**
 * Geocodificação Híbrida para Painel de Inconsistências
 * Tenta Mapbox primeiro (se forceGoogle = false). Se falhar, vai para Google/Gemini.
 */
function geocodificarHibrido(endereco, method, mandadoNum) {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") return null;

  // Checagem prévia no cache local de coordenadas
  var cacheRes = buscarCoordenadasEmCache_(endereco);
  if (cacheRes) {
    try { registrarConsumoGeocodificacao(endereco, "Manual (Cache)", cacheRes.fonte, true, mandadoNum); } catch(e) {}
    return cacheRes;
  }

  if (method === true) method = "Google";
  if (!method || method === false) method = "Hibrido";

  var resFinal = null;

  if (method === "Google") {
    Logger.log("[GEO-HIBRIDO] Forçando Google Maps API para: " + endereco);
    var resGoogle = geocodificarGoogleMaps_(endereco); // Tenta Maps API nativa
    if (!resGoogle) resGoogle = geocodificarComGeminiGrounding(endereco); // Fallback Gemini
    if (resGoogle) resGoogle.fonte = "Google";
    resFinal = resGoogle;
  } else if (method === "Mapbox" || method === "Hibrido") {
    // Tentar Mapbox Geocoding API
    try {
      var mapboxKey = PropertiesService.getScriptProperties().getProperty("MAPBOX_API_KEY");
      if (!mapboxKey) {
        Logger.log("[GEO-HIBRIDO] MAPBOX_API_KEY não configurada nas Script Properties.");
      } else {
        var query = encodeURIComponent(endereco);
        var url = "https://api.mapbox.com/geocoding/v5/mapbox.places/" + query + ".json?access_token=" + mapboxKey + "&country=br&limit=1&language=pt&types=address,poi,place,locality,neighborhood";
        var options = {
          "method": "GET",
          "muteHttpExceptions": true
        };
        var response = UrlFetchApp.fetch(url, options);
        
        if (response.getResponseCode() === 200) {
          var json = JSON.parse(response.getContentText());
          if (json && json.features && json.features.length > 0) {
            var feat = json.features[0];
            var coords = feat.center; // [lng, lat]
            var lat = coords[1];
            var lon = coords[0];
            var relevance = feat.relevance || 0;
            var placeType = (feat.place_type && feat.place_type[0]) || "";
            
            // Mapear relevância e tipo para nível de precisão
            var precisao = "APPROXIMATE";
            if (placeType === "address" && relevance >= 0.9) {
              precisao = "ROOFTOP";
            } else if (placeType === "address" && relevance >= 0.7) {
              precisao = "RANGE_INTERPOLATED";
            } else if (placeType === "poi" && relevance >= 0.8) {
              precisao = "ROOFTOP";
            } else if (placeType === "neighborhood" || placeType === "locality") {
              precisao = "GEOMETRIC_CENTER";
            } else if (placeType === "place") {
              precisao = "GEOMETRIC_CENTER";
            }

            // --- TRAVA RÍGIDA DE MUNICÍPIO E UF ---
            var cidadeMatchOk = true;
            var ufMatchOk = true;
            var cidadeBuscadaInfo = extrairCidadeUfDaBusca_(endereco);

            if (cidadeBuscadaInfo) {
              if (cidadeBuscadaInfo.cidade) {
                var normBuscada = normalizarTextoParaComparacao_(cidadeBuscadaInfo.cidade);
                var placeNameNorm = normalizarTextoParaComparacao_(feat.place_name || "");
                
                var achouNoContexto = false;
                if (feat.context && Array.isArray(feat.context)) {
                  achouNoContexto = feat.context.some(function(ctx) {
                    return normalizarTextoParaComparacao_(ctx.text || "").includes(normBuscada);
                  });
                }

                if (!placeNameNorm.includes(normBuscada) && !achouNoContexto) {
                  cidadeMatchOk = false;
                  Logger.log("[GEO-MAPBOX] REJEITADO por divergência de município: Solicitado '" + cidadeBuscadaInfo.cidade + "', Retornado '" + feat.place_name + "'");
                }
              }

              if (cidadeBuscadaInfo.uf) {
                var normUfBuscada = normalizarTextoParaComparacao_(cidadeBuscadaInfo.uf);
                var placeNameNormUf = normalizarTextoParaComparacao_(feat.place_name || "");
                var achouUfContexto = false;
                if (feat.context && Array.isArray(feat.context)) {
                  achouUfContexto = feat.context.some(function(ctx) {
                    var txt = normalizarTextoParaComparacao_(ctx.text || "");
                    var code = normalizarTextoParaComparacao_(ctx.short_code || "");
                    return txt.includes(normUfBuscada) || code.includes(normUfBuscada);
                  });
                }

                if (!placeNameNormUf.includes(normUfBuscada) && !achouUfContexto) {
                  ufMatchOk = false;
                  Logger.log("[GEO-MAPBOX] REJEITADO por divergência de UF: Solicitado '" + cidadeBuscadaInfo.uf + "', Retornado '" + feat.place_name + "'");
                }
              }
            }

            // EXIGÊNCIA RIGOROSA DE PRECISÃO: Mapbox NUNCA aceita GEOMETRIC_CENTER ou APPROXIMATE em nenhum modo (incluindo modo Mapbox)
            var precisaoAceita = (precisao === "ROOFTOP" || precisao === "RANGE_INTERPOLATED");

            if (cidadeMatchOk && ufMatchOk && precisaoAceita) {
              Logger.log("[GEO-MAPBOX] Sucesso: " + lat + ", " + lon + " | Precisão: " + precisao + " | Relevância: " + relevance);
              resFinal = { lat: lat, lng: lon, enderecoFormatado: feat.place_name || "", fonte: 'Mapbox', precisao: precisao };
              try { registrarConsumoGeocodificacao(endereco, "Manual", "Mapbox", true, mandadoNum); } catch(e) {}
            } else {
              Logger.log("[GEO-MAPBOX] Mapbox REJEITADO: cidadeMatch=" + cidadeMatchOk + ", ufMatch=" + ufMatchOk + ", precisaoAceita=" + precisaoAceita + " (precisão: " + precisao + ", relevância: " + relevance + ")");
            }
          }
        }
        if (!resFinal) Logger.log("[GEO-MAPBOX] Mapbox sem resultados válidos com precisão suficiente.");
      }
    } catch(e) {
      Logger.log("[GEO-MAPBOX] Erro Mapbox: " + e.message);
    }
    
    // Se for apenas Mapbox, retorna null se falhar
    if (method === "Mapbox" && !resFinal) return null;

    if (!resFinal) {
      // Fallback Google Maps API / Gemini (Se for Híbrido)
      Logger.log("[GEO-HIBRIDO] Fallback para Google Maps API...");
      var googleRes = geocodificarGoogleMaps_(endereco);
      if (!googleRes) {
        Logger.log("[GEO-HIBRIDO] Fallback para Google Gemini...");
        googleRes = geocodificarComGeminiGrounding(endereco);
      }
      if (googleRes) {
        if (googleRes.precisao === "APPROXIMATE" || googleRes.precisao === "GEOMETRIC_CENTER") {
          Logger.log("[GEO-HIBRIDO] Google/Gemini retornado rejeitado por precisão imprecisa: " + googleRes.precisao);
          return null;
        }
        googleRes.fonte = 'Google';
        try { registrarConsumoGeocodificacao(endereco, "Manual", googleRes.fonte || "Google", true, mandadoNum); } catch(e) {}
        resFinal = googleRes;
      }
    }
  }

  // Enriquecer resultado com cruzamento de polígonos territoriais (CPI, BTL, CIA, DP, Cidade)
  if (resFinal && resFinal.lat && resFinal.lng) {
    try {
      var areaDetectada = identificarAreaPorCoordenadas(resFinal.lat, resFinal.lng);
      if (areaDetectada) {
        resFinal.cpi = areaDetectada.cpi || "";
        resFinal.batalhao = areaDetectada.batalhao || "";
        resFinal.cia = areaDetectada.cia || "";
        resFinal.delegacia = areaDetectada.delegacia || "";
        resFinal.cidade = areaDetectada.cidade || "";
      } else {
        resFinal.cpi = "";
        resFinal.batalhao = "";
        resFinal.cia = "";
        resFinal.delegacia = "";
        resFinal.cidade = "";
      }
    } catch (errArea) {
      console.warn("Erro ao identificar área para coordenadas (" + resFinal.lat + ", " + resFinal.lng + "):", errArea);
    }
  }

  return resFinal;
}

/**
 * Geocodificação nativa usando Google Maps API (quando a chave está disponível)
 */
function geocodificarGoogleMaps_(endereco) {
  try {
    if (typeof obterApiKeyMaps !== 'function') return null;
    var apiKey = obterApiKeyMaps();
    if (!apiKey) return null;
    
    var url = "https://maps.googleapis.com/maps/api/geocode/json?address=" + encodeURIComponent(endereco) + "&key=" + apiKey;
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    
    if (response.getResponseCode() === 200) {
      var json = JSON.parse(response.getContentText());
      if (json.status === "OK" && json.results.length > 0) {
        var res = json.results[0];
        return {
          lat: res.geometry.location.lat,
          lng: res.geometry.location.lng,
          enderecoFormatado: res.formatted_address,
          precisao: res.geometry.location_type
        };
      }
    }
  } catch(e) {}
  return null;
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
    if (data[i][0]) {
      cidades.push({
        municipio: removerAcentos_(data[i][0].toString().trim()),
        uf: data[i][1] ? data[i][1].toString().trim().toUpperCase() : ""
      });
    }
  }
  _cacheCidades = cidades;
  return cidades;
}

function obterCidadesCacheFrontend() {
  return carregarCidadesCache_();
}

function removerAcentos_(texto) {
  if (!texto) return "";
  return texto.toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
}

function validarEnderecoAntesDeGeocodificar_(endereco) {
  if (!endereco) return false;
  // Se tiver CEP explícito, permite
  if (/\b\d{5}-?\d{3}\b/.test(endereco) || /\bCEP\s*\d+\b/i.test(endereco)) return true;
  
  // Se tiver UF explícita, permite (evita bloqueios desnecessários para endereços válidos)
  if (new RegExp("\\b(" + UFS_BRASIL + ")\\b", "i").test(endereco)) return true;
  
  var cidades = carregarCidadesCache_();
  if (cidades.length === 0) return true;
  
  var endLower = removerAcentos_(endereco);
  for (var i = 0; i < cidades.length; i++) {
    var cidRegex = cidades[i].municipio.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    var regex = new RegExp("\\b" + cidRegex + "\\b", "i");
    if (regex.test(endLower)) return true;
  }
  return false;
}

/**
 * Validação rigorosa: Verifica se a cidade (do banco de cidades) consta explicitamente no endereço.
 * Usado na importação em lote para desviar para a aba Inconsistencias.
 */
function validarCidadeNoEnderecoInconsistencia_(endereco) {
  if (!endereco) return false;
  
  var cidades = carregarCidadesCache_();
  if (cidades.length === 0) return true; // fallback se não houver cidades
  
  var endLower = removerAcentos_(endereco);
  for (var i = 0; i < cidades.length; i++) {
    var cidRegex = cidades[i].municipio.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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
    var ativo = String(data[i][7] || "SIM").trim().toUpperCase();
    if (ativo !== "SIM" && ativo !== "") continue;

    var geoJsonStr = String(data[i][5] || "");
    for (var col = 8; col < data[i].length; col++) {
      if (data[i][col]) geoJsonStr += String(data[i][col]);
    }
    if (!geoJsonStr) continue;

    try {
      var geo = JSON.parse(geoJsonStr);
      if (!geo || !geo.coordinates) continue;

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
function gravarAreaNaLinha_(sheet, linhaIdx, lat, lng, mapCol) {
  const col = mapCol || obterMapaColunas(sheet);
  
  if (lat && lng) {
    var area = identificarAreaPorCoordenadas(lat, lng);
    if (area) {
      if (col['CPI'] >= 0) sheet.getRange(linhaIdx, col['CPI'] + 1).setValue(area.cpi);
      if (col['BPM Area'] >= 0) sheet.getRange(linhaIdx, col['BPM Area'] + 1).setValue(area.batalhao);
      if (col['CIA Area'] >= 0) sheet.getRange(linhaIdx, col['CIA Area'] + 1).setValue(area.cia);
      if (col['DP Area'] >= 0) sheet.getRange(linhaIdx, col['DP Area'] + 1).setValue(area.delegacia);
      if (col['Cidade'] >= 0) sheet.getRange(linhaIdx, col['Cidade'] + 1).setValue(area.cidade);
      if (area.batalhao) {
        if (col['Batalhao'] >= 0) sheet.getRange(linhaIdx, col['Batalhao'] + 1).setValue(area.batalhao);
      }
    }
  }

  // Marcar como não mais pendente de re-verificação de área (v4.9.8)
  var obsCol = col['Observacoes'] !== undefined ? col['Observacoes'] : -1;
  if (obsCol >= 0) {
    var obsAtual = String(sheet.getRange(linhaIdx, obsCol + 1).getValue() || "");
    if (obsAtual.includes("RECALCULAR_AREA")) {
      obsAtual = obsAtual.replace(/\|?\s*RECALCULAR_AREA/g, "").replace(/^[\s|]+|[\s|]+$/g, "").trim();
      sheet.getRange(linhaIdx, obsCol + 1).setValue(obsAtual);
    }
  }
}

/**
 * Função utilitária para corrigir o passado: faz o reverse geocoding 
 * de todos os mandados que têm lat/lng mas não têm "enderecoGeocodificado" no JSON.
 */
function corrigirEnderecosLegados() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
  if (!sheet) return;
  var col = obterMapaColunas(sheet);
  if (col['Latitude'] < 0 || col['Dados Extras JSON'] < 0) return;
  
  var data = sheet.getDataRange().getValues();
  var limit = 0;
  var geocoder = Maps.newGeocoder();
  
  for (var i = 1; i < data.length; i++) {
    var lat = parseFloat(data[i][col['Latitude']]);
    var lng = parseFloat(data[i][col['Longitude']]);
    
    if (lat && lng && !isNaN(lat) && !isNaN(lng)) {
      var jsonStr = String(data[i][col['Dados Extras JSON']] || "{}");
      try {
        var obj = JSON.parse(jsonStr);
        if (!obj.enderecoGeocodificado) {
          // Precisamos fazer reverse geocode
          if (limit > 800) break; // Limite de cota diária do Maps nativo
          
          try {
            var response = geocoder.reverseGeocode(lat, lng);
            if (response.status === 'OK' && response.results.length > 0) {
              obj.enderecoGeocodificado = response.results[0].formatted_address;
              sheet.getRange(i + 1, col['Dados Extras JSON'] + 1).setValue(JSON.stringify(obj));
              limit++;
              Utilities.sleep(100); // 100ms para evitar rate limit
            }
          } catch(e) {
            Logger.log("Erro no reverse geocode para linha " + (i+1));
          }
        }
      } catch(e) {}
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
function executarGeocodificacaoLinha(sheet, linhaIdx, isBatchMode, mapCol) {
  const col = mapCol || obterMapaColunas(sheet);
  const mandadoNum = sheet.getRange(linhaIdx, col['Mandado'] + 1).getValue();
  const batalhao = sheet.getRange(linhaIdx, col['Batalhao'] + 1).getValue();
  let textoCompleto = sheet.getRange(linhaIdx, col['Endereco Principal'] + 1).getValue();
  const latAtual = sheet.getRange(linhaIdx, col['Latitude'] + 1).getValue();
  let obsColIdx = col['Observações'] !== undefined ? col['Observações'] : col['Observacoes'];
  const observacao = obsColIdx !== undefined ? sheet.getRange(linhaIdx, obsColIdx + 1).getValue() : "";

  if (!textoCompleto || latAtual) return false;
  if (String(observacao).includes("(Geocodificação retida)")) return false;

  let lat = null;
  let lng = null;
  let enderecoFormatado = "";
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
      mensagemObs = "ERRO_GEOCODIFICACAO: Endereço inválido.";
    } else if (!validarEnderecoAntesDeGeocodificar_(enderecoPrincipal)) {
      pinoNaSede = true;
      mensagemObs = "ERRO_GEOCODIFICACAO: Cidade não coberta na tabela Cidades e sem CEP.";
    }
  }

  if (pinoNaSede) {
    lat = null;
    lng = null;
  } else {
    // 1. Checagem em Cache Local (Economia de Requisições de API)
    var cacheRes = buscarCoordenadasEmCache_(enderecoPrincipal);
    if (cacheRes && cacheRes.lat && cacheRes.lng) {
      lat = cacheRes.lat;
      lng = cacheRes.lng;
      enderecoFormatado = cacheRes.enderecoFormatado || "";
      try {
        registrarConsumoGeocodificacao(enderecoPrincipal, "Principal (Cache)", cacheRes.fonte, true, mandadoNum);
      } catch(e) {}
    } else {
      let geocoderNativoDisponivel = true;
      try {
        const geocoder = Maps.newGeocoder();
        const response = geocoder.geocode(enderecoPrincipal);
        
        if (response.status === 'OK' && response.results.length > 0) {
          var locType = response.results[0].geometry.location_type;
          if (locType === "APPROXIMATE" || locType === "GEOMETRIC_CENTER") {
            lat = null;
            lng = null;
            mensagemObs = "ERRO_GEOCODIFICACAO: Endereço impreciso (" + locType + " - Nível de cidade ou bairro). Requer revisão manual.";
          } else {
            lat = response.results[0].geometry.location.lat;
            lng = response.results[0].geometry.location.lng;
            enderecoFormatado = response.results[0].formatted_address || "";
            Logger.log("[GEO-NATIVO] Sucesso: " + enderecoPrincipal + " -> " + lat + ", " + lng);
          }
        } else {
          lat = null;
          lng = null;
          mensagemObs = "ERRO_GEOCODIFICACAO: Endereço não localizado com precisão.";
        }
      } catch (eGeo) {
        // Se o Geocoder nativo falhou (cota ou erro de rede), usamos o Gemini como fallback
        geocoderNativoDisponivel = false;
        Logger.log("[GEO-NATIVO] Falha no geocoder nativo, tentando Gemini: " + eGeo.message);
      }

      if (!geocoderNativoDisponivel) {
        try {
          const geoResult = geocodificarComGeminiGrounding(enderecoPrincipal);
          if (geoResult) {
            if (geoResult.precisao === "APPROXIMATE" || geoResult.precisao === "GEOMETRIC_CENTER") {
              lat = null;
              lng = null;
              mensagemObs = "ERRO_GEOCODIFICACAO: Endereço impreciso (" + geoResult.precisao + "). Requer revisão manual.";
            } else {
              // LLM pode alucinar ROOFTOP, mas é o melhor que temos se o nativo estiver fora do ar
              lat = geoResult.lat;
              lng = geoResult.lng;
              enderecoFormatado = geoResult.enderecoFormatado || "";
              mensagemObs = ""; // Limpa qualquer erro prévio
            }
          } else {
            lat = null;
            lng = null;
            mensagemObs = "ERRO_GEOCODIFICACAO: Gemini não conseguiu geocodificar.";
          }
        } catch (erroGemini) {
          const msg = erroGemini.message ? erroGemini.message.toLowerCase() : "";
          if (msg === "quota_limit" || msg.includes("quota") || msg.includes("limit")) {
            throw new Error("QUOTA_LIMIT");
          }
          lat = null;
          lng = null;
        }
      }
    }
  }

  // Sempre gravar o endereço principal e outros endereços na planilha
  sheet.getRange(linhaIdx, col['Endereco Principal'] + 1).setValue(enderecoPrincipal);
  sheet.getRange(linhaIdx, col['Outros Enderecos'] + 1).setValue(outrosEnderecos);

  if (lat && lng) {
    sheet.getRange(linhaIdx, col['Latitude'] + 1).setValue(lat);
    sheet.getRange(linhaIdx, col['Longitude'] + 1).setValue(lng);
    sheet.getRange(linhaIdx, col['Status'] + 1).setValue("Procurado");
    try {
      registrarConsumoGeocodificacao(enderecoPrincipal, "Principal", geocoderNativoDisponivel ? "Google" : "Gemini", true, mandadoNum);
    } catch(e) {}

    // Identificar área (BPM/CIA/DP) pelas coordenadas
    gravarAreaNaLinha_(sheet, linhaIdx, lat, lng, col);

    // Gravar o endereço formatado no Dados Extras JSON
    if (enderecoFormatado) {
      try {
        var extrasColIdx = col['Dados Extras JSON'];
        var extrasStr = String(sheet.getRange(linhaIdx, extrasColIdx + 1).getValue() || "{}");
        var extras = JSON.parse(extrasStr);
        extras.enderecoGeocodificado = enderecoFormatado;
        sheet.getRange(linhaIdx, extrasColIdx + 1).setValue(JSON.stringify(extras));
      } catch(e) {}
    }

    // v4.9.85: Validar endereços secundários contra tabela de cidades
    if (outrosEnderecos && outrosEnderecos.trim().length > 5) {
      var secParts = outrosEnderecos.split(/[;\/]/).map(function(s){ return s.trim(); }).filter(function(s){ return s.length > 3; });
      var secInvalidos = [];
      for (var si = 0; si < secParts.length; si++) {
        if (!validarEnderecoAntesDeGeocodificar_(secParts[si])) {
          secInvalidos.push(secParts[si]);
        }
      }
      if (secInvalidos.length > 0) {
        mensagemObs += (mensagemObs ? " | " : "") + "INCONSISTÊNCIA: Endereço(s) secundário(s) sem cidade/UF válida: " + secInvalidos.join("; ");
      }
    }

    // Se obteve sucesso, limpar erros de geocodificação antigos nas observações
    if (obsColIdx !== undefined) {
      var obsAtual = String(sheet.getRange(linhaIdx, obsColIdx + 1).getValue() || "");
      if (obsAtual.includes("ERRO_GEOCODIFICACAO")) {
        obsAtual = obsAtual.replace(/\|?\s*ERRO_GEOCODIFICACAO:[^|]+/g, "").replace(/^[\s|]+|[\s|]+$/g, "").trim();
        sheet.getRange(linhaIdx, obsColIdx + 1).setValue(obsAtual);
      }
    }

    if (mensagemObs !== "") {
      const obsAtualAgora = String(sheet.getRange(linhaIdx, col['Observacoes'] + 1).getValue() || "");
      if (!obsAtualAgora.includes(mensagemObs)) {
        const novaObs = obsAtualAgora ? obsAtualAgora + " | " + mensagemObs : mensagemObs;
        sheet.getRange(linhaIdx, col['Observacoes'] + 1).setValue(novaObs);
      }
    }

    // Se não tiver data de lançamento, inserir agora
    const dataLanc = sheet.getRange(linhaIdx, col['Data de Lancamento'] + 1).getValue();
    if (!dataLanc) {
      const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
      sheet.getRange(linhaIdx, col['Data de Lancamento'] + 1).setValue(dataAtual);
    }

    return true;
  } else {
    // Se falhar na geocodificação, limpar coordenadas e áreas para sumir do mapa e forçar inconsistência
    sheet.getRange(linhaIdx, col['Latitude'] + 1).clearContent();
    sheet.getRange(linhaIdx, col['Longitude'] + 1).clearContent();
    sheet.getRange(linhaIdx, col['CPI'] + 1).clearContent();
    sheet.getRange(linhaIdx, col['CIA Area'] + 1).clearContent();
    sheet.getRange(linhaIdx, col['DP Area'] + 1).clearContent();
    sheet.getRange(linhaIdx, col['BPM Area'] + 1).clearContent();

    // Limpar o endereço formatado do Dados Extras JSON
    try {
      var extrasColIdx = col['Dados Extras JSON'];
      var extrasStr = String(sheet.getRange(linhaIdx, extrasColIdx + 1).getValue() || "{}");
      var extras = JSON.parse(extrasStr);
      if (extras.enderecoGeocodificado) {
        delete extras.enderecoGeocodificado;
        sheet.getRange(linhaIdx, extrasColIdx + 1).setValue(JSON.stringify(extras));
      }
    } catch(e) {}

    // Sempre salvar a observação de erro
    if (mensagemObs && obsColIdx !== undefined) {
      var obsAtual = String(sheet.getRange(linhaIdx, obsColIdx + 1).getValue() || "");
      if (!obsAtual.includes(mensagemObs)) {
        sheet.getRange(linhaIdx, obsColIdx + 1).setValue((obsAtual ? obsAtual + " | " : "") + mensagemObs);
      }
    }

    // Copiar para a aba de Inconsistências se ainda não estiver lá
    try {
      var ss = sheet.getParent();
      var sheetIncons = ss.getSheetByName("Inconsistencias");
      if (!sheetIncons) {
        sheetIncons = ss.insertSheet("Inconsistencias");
        var headers = sheet.getRange(1, 1, 1, 30).getValues();
        sheetIncons.getRange(1, 1, 1, 30).setValues(headers)
          .setFontWeight("bold").setBackground("#1e293b").setFontColor("#e2e8f0");
        sheetIncons.setFrozenRows(1);
      }
      
      var dadosI = sheetIncons.getDataRange().getValues();
      var colMapI = obterMapaCabecalhos(dadosI[0]);
      var mandado = String(sheet.getRange(linhaIdx, col['Mandado'] + 1).getValue() || "").trim();
      var colMandadoI = colMapI['Mandado'] !== undefined ? colMapI['Mandado'] : colMapI['Nº Mandado'];
      
      var existeIncons = false;
      if (colMandadoI !== undefined) {
        for (var j = 1; j < dadosI.length; j++) {
          if (String(dadosI[j][colMandadoI] || "").trim() === mandado) {
            existeIncons = true;
            break;
          }
        }
      }
      
      if (!existeIncons) {
        var rowData = sheet.getRange(linhaIdx, 1, 1, 30).getValues()[0];
        sheetIncons.appendRow(rowData);
      }
    } catch(eInc) {
      Logger.log("Erro ao registrar na aba de Inconsistências: " + eInc.message);
    }
    return false;
  }
}

/**
 * Solicita a chave do Mapbox via caixa de diálogo e salva nas ScriptProperties.
 * Evita expor tokens privados diretamente no código-fonte.
 */
function configurarMapboxApiKey() {
  var ui = SpreadsheetApp.getUi();
  var resposta = ui.prompt(
    "Configuração de Segurança", 
    "Cole a sua MAPBOX_API_KEY no campo abaixo:", 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (resposta.getSelectedButton() === ui.Button.OK) {
    var chave = resposta.getResponseText().trim();
    if (chave && chave.length > 10) {
      PropertiesService.getScriptProperties().setProperty("MAPBOX_API_KEY", chave);
      ui.alert("✅ MAPBOX_API_KEY salva com sucesso nas Script Properties!");
    } else {
      ui.alert("⚠️ Nenhuma chave válida foi informada.");
    }
  }
}

/**
 * Extrai Cidade e UF do texto de busca de endereço
 */
function extrairCidadeUfDaBusca_(str) {
  if (!str) return null;
  var s = String(str).trim();

  // 1. Padrão ", Cidade - UF" ou ", Cidade/UF" ou " - Cidade - UF"
  var m = s.match(/(?:,|\s-)\s*([A-ZÁÀÂÃÉÈÊÍÏÓÒÔÕÚÜÇ\s.-]+?)\s*[-/]\s*([A-Z]{2})/i);
  if (m) {
    return {
      cidade: m[1].trim(),
      uf: m[2].trim().toUpperCase()
    };
  }

  // 2. Padrão "de Cidade - UF" ou "em Cidade - UF" ou "de Cidade SP"
  var m2 = s.match(/(?:de|em|\b)\s+([A-ZÁÀÂÃÉÈÊÍÏÓÒÔÕÚÜÇ\s.-]+?)\s+(SP|RJ|MG|PR|SC|RS|BA|GO|MT|MS|PE|CE|PA|MA|PB|RN|AL|SE|PI|AM|AP|AC|RO|RR|TO|DF)\b/i);
  if (m2) {
    return {
      cidade: m2[1].trim(),
      uf: m2[2].trim().toUpperCase()
    };
  }

  // 3. Padrão "de <Cidade>"
  var m3 = s.match(/(?:de|em)\s+([A-ZÁÀÂÃÉÈÊÍÏÓÒÔÕÚÜÇ]+)\b/i);
  if (m3) {
    return {
      cidade: m3[1].trim(),
      uf: "SP"
    };
  }

  return { cidade: null, uf: "SP" };
}

/**
 * Normaliza texto para comparação sem acentos nem caracteres especiais
 */
function normalizarTextoParaComparacao_(str) {
  if (!str) return '';
  return String(str)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .trim();
}
