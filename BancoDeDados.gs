// ================================================================
// SENTINELA v4.1.4 — MÓDULO DE BANCO DE DADOS (CRUD + GAMIFICAÇÃO)
// Refatoração: Índices dinâmicos por cabeçalho para resiliência.
// ================================================================

/**
 * Helper v4.1.0: Captura os cabeçalhos da aba e retorna um mapa
 * { nomeCabeçalho: índice0based }. Usado por todas as funções CRUD
 * para eliminar acoplamento rígido com posições fixas de colunas.
 */
function obterMapaColunas(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  let needsUpdate = false;
  if (headers.indexOf('Dados Extras JSON') === -1) {
    headers.push('Dados Extras JSON');
    needsUpdate = true;
  }
  if (headers.indexOf('Observacoes') === -1) {
    headers.push('Observacoes');
    needsUpdate = true;
  }
  if (needsUpdate) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const mapa = {};
  headers.forEach(function(h, i) {
    mapa[String(h).trim()] = i;
  });
  return mapa;
}

/**
 * Lê todos os dados da aba Mandados e retorna como array de objetos.
 * v4.1.0: Usa cabeçalhos dinâmicos em vez de índices hardcoded.
 */
function obterDados() {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado) {
    throw new Error("Acesso negado: Operação não autorizada.");
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const linhas = data.slice(1);

  // Mapear índices dinamicamente a partir dos cabeçalhos
  const idx = {};
  const nomesEsperados = [
    'Data de Lancamento', 'Data de Conferencia', 'Mandado', 'Artigo', 'Nome',
    'CPF', 'RG', 'Nascimento', 'Naturalidade', 'Sexo', 'Cor', 'Filiacao',
    'Foto URL', 'Batalhao', 'Endereco Principal', 'Outros Enderecos', 'Status', 'Validade',
    'Info Processuais', 'Geodata Secundarios', 'Dados Extras JSON', 'Observacoes',
    'Latitude', 'Longitude', 'CPI', 'BPM Area', 'CIA Area', 'DP Area', 'Cidade',
    'TipoImportacao', 'REVISÃO', 'Revisão', 'REVISAO', 'Revisao', 'Revisado'
  ];
  nomesEsperados.forEach(function(nome) {
    idx[nome] = headers.indexOf(nome);
  });

  const procurados = [];

  linhas.forEach((linha, index) => {
    const latRaw = idx.Latitude >= 0 ? linha[idx.Latitude] : null;
    const lngRaw = idx.Longitude >= 0 ? linha[idx.Longitude] : null;
    
    let lat = null;
    let lng = null;
    let temCoordenadas = false;

    if (latRaw !== undefined && latRaw !== null && latRaw !== "" &&
        lngRaw !== undefined && lngRaw !== null && lngRaw !== "") {
      const parsedLat = parseFloat(latRaw);
      const parsedLng = parseFloat(lngRaw);
      if (!isNaN(parsedLat) && !isNaN(parsedLng)) {
        lat = parsedLat;
        lng = parsedLng;
        temCoordenadas = true;
      }
    }

    const formatarData = (valor) => {
      if (valor instanceof Date) {
        return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      return String(valor || "").trim();
    };

    const formatarDataHora = (valor) => {
      if (valor instanceof Date) {
        return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      }
      return String(valor || "").trim();
    };

    // Helper seguro: retorna valor da coluna ou fallback
    const col = (nome, fallback) => idx[nome] >= 0 ? linha[idx[nome]] : (fallback !== undefined ? fallback : "");

    procurados.push({
      idLinha: index + 2,
      dataLancamento: formatarData(col('Data de Lancamento')),
      dataConferencia: formatarDataHora(col('Data de Conferencia')),
      mandado: String(col('Mandado') || ""),
      artigo: String(col('Artigo') || ""),
      nome: String(col('Nome') || ""),
      cpf: String(col('CPF') || ""),
      rg: String(col('RG') || "N/A"),
      nascimento: formatarData(col('Nascimento')),
      naturalidade: String(col('Naturalidade') || "N/A"),
      sexo: String(col('Sexo') || "N/A"),
      cor: String(col('Cor') || "N/A"),
      filiacao: String(col('Filiacao') || "N/A"),
      fotoUrl: (function() {
        var v = String(col('Foto URL') || "").trim();
        // Sanitiza: descarta valores inválidos como strings ERRO:, base64, etc.
        if (!v || v === "" || v === "N/A" || !v.startsWith("http")) return "N/A";
        return v;
      })(),
      batalhao: String(col('Batalhao') || ""),
      enderecoPrincipal: String(col('Endereco Principal') || ""),
      outrosEnderecos: String(col('Outros Enderecos') || ""),
      status: String(col('Status') || "Procurado"),
      validade: formatarData(col('Validade')),
      infoProcessuais: String(col('Info Processuais') || ""),
      geodataSecundarios: (function() {
        try {
          const rawSec = String(col('Geodata Secundarios') || "[]");
          if (!rawSec || rawSec === "[]") return "[]";
          const secArr = JSON.parse(rawSec);
          if (Array.isArray(secArr) && secArr.length > 0) {
            let alterou = false;
            secArr.forEach(function(sec) {
              if (sec.lat && sec.lng && (!sec.cpi || !sec.batalhao || !sec.cia)) {
                const areaSec = identificarAreaPorCoordenadas(Number(sec.lat), Number(sec.lng));
                if (areaSec) {
                  if (areaSec.cpi) { sec.cpi = areaSec.cpi; alterou = true; }
                  if (areaSec.batalhao) { sec.batalhao = areaSec.batalhao; alterou = true; }
                  if (areaSec.cia) { sec.cia = areaSec.cia; alterou = true; }
                  if (areaSec.delegacia) { sec.delegacia = areaSec.delegacia; alterou = true; }
                  if (areaSec.cidade) { sec.cidade = areaSec.cidade; alterou = true; }
                }
              }
            });
            if (alterou) return JSON.stringify(secArr);
          }
        } catch(e) {}
        return String(col('Geodata Secundarios') || "[]");
      })(),
      dadosExtrasJSON: String(col('Dados Extras JSON') || "{}"),
      observacoes: "", // Textarea ficará vazio por padrão
      historicoObservacoes: (function() {
        try {
          const obsStr = String(col('Observacoes') || "").trim();
          if (obsStr.startsWith('[')) {
             try { JSON.parse(obsStr); return obsStr; } catch(e) {}
          } else if (obsStr !== "") {
             return JSON.stringify([{ data: "Legado", usuario: "Sistema", texto: obsStr }]);
          }
        } catch(e) {}
        return "[]";
      })(),
      bnmpConferencia: (function() {
        try {
          const extraStr = col('Dados Extras JSON');
          if (extraStr) {
            const extra = JSON.parse(extraStr);
            return extra.bnmpConferencia || "";
          }
        } catch(e) {}
        return "";
      })(),
      dataBNMP: (function() {
        try {
          const extraStr = col('Dados Extras JSON');
          if (extraStr) {
            const extra = JSON.parse(extraStr);
            return extra.emissão || "";
          }
        } catch(e) {}
        return "";
      })(),
      tipoMandado: (function() {
        try {
          const extraStr = col('Dados Extras JSON');
          if (extraStr) {
            const extra = JSON.parse(extraStr);
            return extra.titulo || "";
          }
        } catch(e) {}
        return "";
      })(),
      lat: lat,
      lng: lng,
      cpi: String(col('CPI') || ""),
      bpmArea: String(col('BPM Area') || ""),
      ciaArea: String(col('CIA Area') || ""),
      dpArea: String(col('DP Area') || ""),
      cidade: String(col('Cidade') || ""),
      tipoImportacao: String(col('TipoImportacao') || ""),
      semEndereco: !temCoordenadas,
      enderecoGeocodificado: (function() {
        try {
          const extraStr = col('Dados Extras JSON');
          if (extraStr) {
            const extra = JSON.parse(extraStr);
            return extra.enderecoGeocodificado || "";
          }
        } catch(e) {}
        return "";
      })(),
      // Alias de compatibilidade
      cpiArea: String(col('CPI') || ""),
      data: formatarData(col('Data de Lancamento'))
    });
  });

  return procurados;
}

/**
 * Cadastra um novo mandado no banco de dados.
 * v4.1.0: Usa cabeçalhos dinâmicos para verificar duplicidade e posicionar NumberFormat.
 */
function cadastrarMandadoWebAppPreview(dados) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil === "Patrulheiro") {
    throw new Error("Acesso negado: Patrulheiros não podem cadastrar mandados.");
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
    const col = obterMapaColunas(sheet);

    // Verificar duplicidade de mandado
    const idxMandado = col['Mandado'];
    const todasAsLinhas = sheet.getDataRange().getValues();
    for (let i = 1; i < todasAsLinhas.length; i++) {
      if (todasAsLinhas[i][idxMandado] === dados.mandado) {
        return { sucesso: false, mensagem: "⚠️ REJEITADO: Mandado já cadastrado." };
      }
    }

    const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    // Processar foto
    let urlFotoSalva = "N/A";
    let fotoInput = dados.fotoBase64 || dados.fotoUrl || dados.foto || "";
    if (fotoInput && fotoInput !== "" && fotoInput !== "N/A") {
      if (!fotoInput.startsWith("http")) {
        urlFotoSalva = processarEDespacharFotoNoDrive(dados.mandado, fotoInput, dados.nome);
      } else {
        urlFotoSalva = fotoInput;
      }
    }

    // Forçar dados geográficos e de área como nulos/vazios (o sistema calcula posteriormente)
    var areaInfo = null;

    // Limpar quaisquer coordenadas dos endereços secundários no cadastro
    (dados.secundarios || []).forEach(sec => {
      sec.lat = null;
      sec.lng = null;
      sec.cpi = "";
      sec.batalhao = "";
      sec.cia = "";
      sec.delegacia = "";
      sec.cidade = "";
    });

    const jsonSecundarios = JSON.stringify(dados.secundarios || []);
    const textoSecundarios = (dados.secundarios || []).map(s => s.endereco).join("\n");
    dados.extras = dados.extras || {};
    dados.extras.Criado_Por = checagem.email;
    if (dados.enderecoGeocodificado) {
      dados.extras.enderecoGeocodificado = dados.enderecoGeocodificado;
    }
    const jsonExtras = JSON.stringify(dados.extras);

    // Montar linha na ordem dos cabeçalhos usando mapa dinâmico
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const novaLinha = new Array(headers.length).fill("");
    
    // Indice resiliente para Foto / Foto URL
    let idxFotoCol = col['Foto URL'];
    if (idxFotoCol === undefined) idxFotoCol = col['Foto'];
    if (idxFotoCol === undefined) {
      for (let k in col) {
        if (k.toLowerCase().includes('foto')) { idxFotoCol = col[k]; break; }
      }
    }

    novaLinha[col['Data de Lancamento']] = dataAtual;
    novaLinha[col['Data de Conferencia']] = "";
    novaLinha[col['Mandado']] = dados.mandado;
    novaLinha[col['Artigo']] = dados.artigo || "Vide Info Proc.";
    novaLinha[col['Nome']] = dados.nome;
    novaLinha[col['CPF']] = dados.cpf;
    novaLinha[col['RG']] = dados.rg;
    novaLinha[col['Nascimento']] = dados.nascimento;
    novaLinha[col['Naturalidade']] = dados.naturalidade;
    novaLinha[col['Sexo']] = dados.sexo;
    novaLinha[col['Cor']] = dados.cor;
    novaLinha[col['Filiacao']] = dados.filiacao;
    if (idxFotoCol !== undefined) {
      novaLinha[idxFotoCol] = urlFotoSalva;
    }
    novaLinha[col['Batalhao']] = "A DEFINIR (GEO)";
    novaLinha[col['Endereco Principal']] = dados.enderecoPrincipal;
    novaLinha[col['Outros Enderecos']] = textoSecundarios;
    novaLinha[col['Status']] = "Procurado";
    novaLinha[col['Validade']] = dados.validade;
    novaLinha[col['Info Processuais']] = dados.infoProcessuais;
    novaLinha[col['Geodata Secundarios']] = jsonSecundarios;
    novaLinha[col['Dados Extras JSON']] = jsonExtras;
    novaLinha[col['Observacoes']] = "";
    novaLinha[col['Latitude']] = null;
    novaLinha[col['Longitude']] = null;
    novaLinha[col['CPI']] = "";
    novaLinha[col['BPM Area']] = "";
    novaLinha[col['CIA Area']] = "";
    novaLinha[col['DP Area']] = "";
    novaLinha[col['Cidade']] = "";
    novaLinha[col['TipoImportacao']] = dados.tipoImportacao || "REGEX";

    sheet.appendRow(novaLinha);
    // Formatar colunas de coordenadas (1-based = idx + 1)
    sheet.getRange(sheet.getLastRow(), col['Latitude'] + 1, 1, 2).setNumberFormat("0.00000000");

    // Registrar no Historico
    registrarHistorico(checagem.email, checagem.nome || "Admin", "CADASTRO", "Cadastrou mandado: " + dados.nome + " (" + dados.mandado + ")");

    // Registrar pontos de gamificação para o Admin
    registrarPontosGamificacao(checagem.email, "CADASTRO", 0, dados.mandado, "Cadastrou mandado no sistema");
    SpreadsheetApp.flush();
    sinalizarMudancaMandados();

    // FIREBASE SYNC: Sincronizar o novo mandado
    try {
      sincronizarMandadoFirebase({
        mandado: dados.mandado, nome: dados.nome, cpf: dados.cpf, rg: dados.rg,
        nascimento: dados.nascimento, naturalidade: dados.naturalidade, sexo: dados.sexo,
        cor: dados.cor, filiacao: dados.filiacao, fotoUrl: urlFotoSalva,
        batalhao: 'A DEFINIR (GEO)', enderecoPrincipal: dados.enderecoPrincipal,
        outrosEnderecos: textoSecundarios, status: 'Procurado', validade: dados.validade,
        infoProcessuais: dados.infoProcessuais, geodataSecundarios: jsonSecundarios,
        dadosExtrasJSON: jsonExtras, dataLancamento: dataAtual,
        tipoImportacao: dados.tipoImportacao || 'REGEX'
      });
    } catch(fbErr) { console.error('Firebase sync falhou no cadastro: ' + fbErr.message); }

    return { sucesso: true, mensagem: "Alvo e todos os endereços salvos no mapa e no banco!" };
  } catch (erro) {
    return { sucesso: false, mensagem: "Erro no servidor: " + erro.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cadastra múltiplos mandados em lote de uma só vez usando setValues
 */
function cadastrarMandadosEmLote(listaDados) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil === "Patrulheiro") {
    return { sucesso: false, mensagem: "Acesso negado: Patrulheiros não podem cadastrar mandados." };
  }
  
  if (!listaDados || listaDados.length === 0) {
    return { sucesso: false, mensagem: "Nenhum dado enviado." };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = planilha.getSheetByName("Mandados");
    if (!sheet) return { sucesso: false, mensagem: "Aba Mandados não encontrada." };
    
    // Obter ou criar aba de Inconsistencias
    let sheetInconsistencias = planilha.getSheetByName("Inconsistencias");
    if (!sheetInconsistencias) {
      sheetInconsistencias = planilha.insertSheet("Inconsistencias");
      const headers = sheet.getRange(1, 1, 1, 30).getValues();
      sheetInconsistencias.getRange(1, 1, 1, 30).setValues(headers)
        .setFontWeight("bold").setBackground("#1e293b").setFontColor("#e2e8f0");
      sheetInconsistencias.setFrozenRows(1);
    }
    
    // Filtro para evitar duplicidades que possam ter passado
    const todasAsLinhas = sheet.getDataRange().getValues();
    const mandadosExistentes = todasAsLinhas.slice(1).map(function(r) { return String(r[2] || ""); });
    
    const matrizParaSalvar = [];
    const matrizInconsistencias = [];
    let pontos = 0;
    
    listaDados.forEach(dados => {
      if (mandadosExistentes.includes(dados.mandado)) return; // Pula se já existir
      
      const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

      // Preencher campos de área (CPI/BTL/CIA) nos endereços secundários via cruzamento de polígonos
      (dados.secundarios || []).forEach(sec => {
        if (sec.lat && sec.lng) {
          var areaSec = identificarAreaPorCoordenadas(sec.lat, sec.lng);
          if (areaSec) {
            sec.cpi = areaSec.cpi || "";
            sec.batalhao = areaSec.batalhao || "";
            sec.cia = areaSec.cia || "";
            sec.delegacia = areaSec.delegacia || "";
            sec.cidade = areaSec.cidade || "";
          } else {
            sec.cpi = sec.cpi || "";
            sec.batalhao = sec.batalhao || "";
            sec.cia = sec.cia || "";
            sec.delegacia = sec.delegacia || "";
            sec.cidade = sec.cidade || "";
          }
        } else {
          sec.cpi = sec.cpi || "";
          sec.batalhao = sec.batalhao || "";
          sec.cia = sec.cia || "";
          sec.delegacia = sec.delegacia || "";
          sec.cidade = sec.cidade || "";
        }
      });

      const jsonSecundarios = JSON.stringify(dados.secundarios || []);
      const textoSecundarios = (dados.secundarios || []).map(s => s.endereco).join("\\n");
      dados.extras = dados.extras || {};
      dados.extras.Criado_Por = checagem.email;
      if (dados.enderecoGeocodificado) {
        dados.extras.enderecoGeocodificado = dados.enderecoGeocodificado;
      }
      const jsonExtras = JSON.stringify(dados.extras);

      // IDENTIFICAÇÃO ESPACIAL DE ÁREAS (CPI/BTL/CIA)
      let areaInfo = { cpi: "", batalhao: "", cia: "", delegacia: "", cidade: "" };
      if (dados.latPrincipal && dados.lngPrincipal) {
        const areaDetectada = identificarAreaPorCoordenadas(dados.latPrincipal, dados.lngPrincipal);
        if (areaDetectada) {
          areaInfo = areaDetectada;
        }
      }

      // PROCESSAMENTO DE FOTO EM LOTE
      let urlFotoSalva = "N/A";
      let fotoInput = dados.fotoBase64 || dados.fotoUrl || dados.foto || "";
      if (fotoInput && fotoInput !== "" && fotoInput !== "N/A") {
        if (!fotoInput.startsWith("http")) {
          urlFotoSalva = processarEDespacharFotoNoDrive(dados.mandado, fotoInput, dados.nome);
        } else {
          urlFotoSalva = fotoInput;
        }
      }

      const linhaArray = [
        dataAtual,                // A — Data de Lançamento
        "",                       // B — Data de Conferência
        dados.mandado,            // C — Mandado
        dados.artigo || "Vide Info Proc.", // D — Artigo
        dados.nome,               // E — Nome
        dados.cpf,                // F — CPF
        dados.rg,                 // G — RG
        dados.nascimento,         // H — Nascimento
        dados.naturalidade,       // I — Naturalidade
        dados.sexo,               // J — Sexo
        dados.cor,                // K — Cor
        dados.filiacao,           // L — Filiação
        urlFotoSalva,             // M — Foto URL (agora suporta crop de IA/Regex)
        areaInfo.batalhao || "A DEFINIR (GEO)", // N — Batalhão (Sempre puxa do GeoJSON ou fallback)
        dados.enderecoPrincipal,  // O — Endereço Principal
        textoSecundarios,         // P — Outros Endereços
        "Procurado",              // Q — Status
        dados.validade,           // R — Validade
        dados.infoProcessuais,    // S — Info Processuais
        jsonSecundarios,          // T — Geodata Secundários
        jsonExtras,               // U — Dados Extras JSON
        "Importado em Lote",      // V — Observações
        dados.latPrincipal || null, // W — Latitude
        dados.lngPrincipal || null, // X — Longitude
        areaInfo.cpi,             // Y — CPI
        areaInfo.batalhao,        // Z — BPM_Area
        areaInfo.cia,             // AA — CIA_Area
        areaInfo.delegacia,       // AB — DP_Area
        areaInfo.cidade,          // AC — Cidade
        dados.tipoImportacao || "REGEX" // AD — Tipo de Importação
      ];

      // Validação rigorosa: tem cidade no endereço principal e geocodificação foi bem-sucedida?
      var temCidade = validarCidadeNoEnderecoInconsistencia_(dados.enderecoPrincipal);
      var temCoordenadas = dados.latPrincipal && dados.lngPrincipal;

      if (temCidade && temCoordenadas) {
        matrizParaSalvar.push(linhaArray);
        pontos++;
      } else {
        if (!temCidade) {
          linhaArray[21] = "Falha: Cidade não identificada (Geocodificação retida)"; // Observações
        } else {
          linhaArray[21] = "Falha: Endereço não localizado (Zero resultados)"; // Observações
        }
        matrizParaSalvar.push(linhaArray);
        matrizInconsistencias.push(linhaArray);
        pontos++;
      }
    });
    
    if (matrizParaSalvar.length > 0) {
      const startRow = Math.max(sheet.getLastRow() + 1, 2);
      sheet.getRange(startRow, 1, matrizParaSalvar.length, 30).setValues(matrizParaSalvar);
      sheet.getRange(startRow, 23, matrizParaSalvar.length, 2).setNumberFormat("0.00000000");
    }

    if (matrizInconsistencias.length > 0) {
      const startRowInc = Math.max(sheetInconsistencias.getLastRow() + 1, 2);
      sheetInconsistencias.getRange(startRowInc, 1, matrizInconsistencias.length, 30).setValues(matrizInconsistencias);
    }
    
    SpreadsheetApp.flush();
    
    if (matrizParaSalvar.length > 0 || matrizInconsistencias.length > 0) {
      if (matrizParaSalvar.length > 0) {
        registrarPontosGamificacao(checagem.email, "CADASTRO", 0, "LOTE_" + pontos, "Cadastrou " + pontos + " mandados em lote");
        sinalizarMudancaMandados();

        // FIREBASE SYNC: Batch Insert
        try {
          var batchData = {};
          matrizParaSalvar.forEach(function(linhaArray) {
            var mObj = {
              dataLancamento: linhaArray[0], dataConferencia: linhaArray[1], mandado: linhaArray[2],
              artigo: linhaArray[3], nome: linhaArray[4], cpf: linhaArray[5], rg: linhaArray[6],
              nascimento: linhaArray[7], naturalidade: linhaArray[8], sexo: linhaArray[9], cor: linhaArray[10],
              filiacao: linhaArray[11], fotoUrl: linhaArray[12], batalhao: linhaArray[13],
              enderecoPrincipal: linhaArray[14], outrosEnderecos: linhaArray[15], status: linhaArray[16],
              validade: linhaArray[17], infoProcessuais: linhaArray[18], geodataSecundarios: linhaArray[19],
              dadosExtrasJSON: linhaArray[20], observacoes: linhaArray[21], lat: linhaArray[22], lng: linhaArray[23],
              cpi: linhaArray[24], bpmArea: linhaArray[25], ciaArea: linhaArray[26], dpArea: linhaArray[27],
              cidade: linhaArray[28], tipoImportacao: linhaArray[29]
            };
            batchData[sanitizarChaveFirebase_(mObj.mandado)] = mandadoParaFirebase_(mObj);
          });
          firebasePatch_('mandados', batchData);
        } catch(fbErr) { console.error('Firebase sync lote falhou: ' + fbErr.message); }
      }
      return { 
        sucesso: true, 
        mensagem: "Lote processado! Salvos: " + matrizParaSalvar.length + ". Inconsistentes: " + matrizInconsistencias.length 
      };
    } else {
      return { sucesso: false, mensagem: "Todos os mandados do lote já existiam no banco." };
    }
  } catch (erro) {
    return { sucesso: false, mensagem: "Erro ao gravar lote: " + erro.message };
  } finally {
    lock.releaseLock();
  }
}

function encontrarLinhaPorMandado_(sheet, col, numMandado) {
  if (!numMandado) return -1;
  const colMandado = (col && col['Mandado'] !== undefined) ? col['Mandado'] + 1 : 3;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const vals = sheet.getRange(2, colMandado, lastRow - 1, 1).getValues();
  const numLimpo = String(numMandado).trim();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === numLimpo) {
      return i + 2;
    }
  }
  return -1;
}

/**
 * Edição completa de um mandado (somente Admin).
 * v4.1.0: Usa cabeçalhos dinâmicos para gravação e geocodificação.
 */
function editarMandadoCompleto(dados) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado) {
    return { sucesso: false, mensagem: "Acesso negado." };
  }

  if (checagem.perfil === "Patrulheiro") {
    return { sucesso: false, mensagem: "Acesso negado. Patrulheiros não podem editar estrutura." };
  }

  if (!dados) {
    return { sucesso: false, mensagem: "Dados não informados." };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
    if (!sheet) return { sucesso: false, mensagem: "Aba Mandados não encontrada." };

    const col = obterMapaColunas(sheet);

    let row = parseInt(dados.idLinha);
    if ((!row || isNaN(row) || row < 2) && dados.mandado) {
      row = encontrarLinhaPorMandado_(sheet, col, dados.mandado);
    }

    if (!row || isNaN(row) || row < 2) {
      return { sucesso: false, mensagem: "Linha do registro não encontrada na planilha." };
    }
    dados.idLinha = row;

    // Obter índice flexível da coluna Foto URL
    let idxFotoCol = (function() {
      if (col['Foto URL'] !== undefined) return col['Foto URL'];
      if (col['Foto'] !== undefined) return col['Foto'];
      if (col['FotoUrl'] !== undefined) return col['FotoUrl'];
      if (col['FOTO'] !== undefined) return col['FOTO'];
      for (let k in col) {
        if (k.toLowerCase().includes('foto')) return col[k];
      }
      return -1;
    })();

    // Se foto em Base64 foi enviada na edição, processar e salvar no Google Drive
    let fotoInput = dados.fotoBase64 || dados.foto || "";
    // Ignora se já é uma URL válida (não reprocessa)
    if (fotoInput && !fotoInput.startsWith("http") && fotoInput !== "" && fotoInput !== "N/A") {
      const numMandado = dados.mandado || sheet.getRange(row, col['Mandado'] + 1).getValue() || "Mandado";
      const nomeIndividuo = dados.nome || sheet.getRange(row, col['Nome'] + 1).getValue() || "Procurado";
      const urlFotoDrive = processarEDespacharFotoNoDrive(numMandado, fotoInput, nomeIndividuo);
      // Só salva se recebeu URL válida
      if (urlFotoDrive && urlFotoDrive.startsWith("http")) {
        dados.fotoUrl = urlFotoDrive;
        if (idxFotoCol >= 0) {
          sheet.getRange(row, idxFotoCol + 1).setValue(urlFotoDrive);
        }
      }
    } else if (dados.fotoUrl && dados.fotoUrl.startsWith("http") && idxFotoCol >= 0) {
      // Atualiza apenas se a URL existente for válida
      sheet.getRange(row, idxFotoCol + 1).setValue(dados.fotoUrl);
    }

    // Mapeamento campo frontend → cabeçalho da planilha (dinâmico)
    var campos = {
      nome: col['Nome'],
      cpf: col['CPF'],
      rg: col['RG'],
      nascimento: col['Nascimento'],
      naturalidade: col['Naturalidade'],
      sexo: col['Sexo'],
      cor: col['Cor'],
      filiacao: col['Filiacao'],
      batalhao: col['Batalhao'],
      enderecoPrincipal: col['Endereco Principal'],
      outrosEnderecos: col['Outros Enderecos'],
      status: col['Status'],
      validade: col['Validade'],
      infoProcessuais: col['Info Processuais'],
      observacoes: col['Observacoes'],
      cpiArea: col['CPI'],
      ciaArea: col['CIA Area'],
      cidade: col['Cidade']
    };

    // Gravar cada campo editado (col é 0-based, getRange é 1-based: +1)
    for (var chave in campos) {
      if (dados.hasOwnProperty(chave) && dados[chave] !== undefined && campos[chave] >= 0) {
        if (chave === 'observacoes') {
          var novaObs = String(dados[chave]).trim();
          if (novaObs !== "") {
            var cellObs = sheet.getRange(row, campos[chave] + 1);
            var histAtual = String(cellObs.getValue() || "").trim();
            var arrayHist = [];
            if (histAtual.startsWith('[')) {
               try { arrayHist = JSON.parse(histAtual); } catch(e) {}
            } else if (histAtual !== "") {
               arrayHist.push({ data: "Legado", usuario: "Sistema", texto: histAtual });
            }
            var dataHoraStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
            arrayHist.unshift({ data: dataHoraStr, usuario: checagem.email || "Usuário", texto: novaObs });
            cellObs.setValue(JSON.stringify(arrayHist));
          }
        } else {
          sheet.getRange(row, campos[chave] + 1).setValue(dados[chave]);
        }
      }
    }

    // Se endereço foi editado, tentar geocodificar o Principal
    if (dados.enderecoPrincipal && dados.enderecoPrincipal !== "Não informado" && dados.enderecoPrincipal.length > 5) {
      try {
        var response = Maps.newGeocoder().geocode(dados.enderecoPrincipal);
        if (response.results && response.results.length > 0) {
          var loc = response.results[0].geometry.location;
          sheet.getRange(row, col['Latitude'] + 1).setValue(loc.lat);
          sheet.getRange(row, col['Longitude'] + 1).setValue(loc.lng);
          
          var areaInfo = identificarAreaPorCoordenadas(loc.lat, loc.lng);
          if (areaInfo) {
            if (areaInfo.cpi && col['CPI'] >= 0) sheet.getRange(row, col['CPI'] + 1).setValue(areaInfo.cpi);
            if (areaInfo.batalhao && col['Batalhao'] >= 0) sheet.getRange(row, col['Batalhao'] + 1).setValue(areaInfo.batalhao);
            if (areaInfo.batalhao && col['BPM Area'] >= 0) sheet.getRange(row, col['BPM Area'] + 1).setValue(areaInfo.batalhao);
            if (areaInfo.cia && col['CIA Area'] >= 0) sheet.getRange(row, col['CIA Area'] + 1).setValue(areaInfo.cia);
            if (areaInfo.delegacia && col['DP Area'] >= 0) sheet.getRange(row, col['DP Area'] + 1).setValue(areaInfo.delegacia);
            if (areaInfo.cidade && col['Cidade'] >= 0) sheet.getRange(row, col['Cidade'] + 1).setValue(areaInfo.cidade);
          }
        }
      } catch(geoErr) {
        Logger.log("Geocodificação falhou para principal: " + geoErr.message);
      }
    }

    // Geocodificar Endereços Secundários se houver
    if (dados.outrosEnderecos && dados.outrosEnderecos.trim() !== "") {
      var arraySec = dados.outrosEnderecos.split("||").map(s => s.trim()).filter(s => s.length > 5);
      var geoSecundarios = [];
      var geocoder = Maps.newGeocoder();
      for (var i = 0; i < arraySec.length; i++) {
        try {
          var respSec = geocoder.geocode(arraySec[i]);
          if (respSec.results && respSec.results.length > 0) {
            var lsec = respSec.results[0].geometry.location;
            var areaSec = identificarAreaPorCoordenadas(lsec.lat, lsec.lng);
            geoSecundarios.push({ 
              endereco: arraySec[i], 
              lat: lsec.lat, 
              lng: lsec.lng,
              cpi: areaSec ? areaSec.cpi : "",
              batalhao: areaSec ? areaSec.batalhao : "",
              cia: areaSec ? areaSec.cia : "",
              delegacia: areaSec ? areaSec.delegacia : "",
              cidade: areaSec ? areaSec.cidade : ""
            });
          } else {
            geoSecundarios.push({ endereco: arraySec[i], lat: 0, lng: 0 });
          }
          Utilities.sleep(500); // Pausa para não estourar a cota da API Google
        } catch(errSec) {
          geoSecundarios.push({ endereco: arraySec[i], lat: 0, lng: 0 });
        }
      }
      sheet.getRange(row, col['Geodata Secundarios'] + 1).setValue(JSON.stringify(geoSecundarios));
    } else {
      sheet.getRange(row, col['Geodata Secundarios'] + 1).setValue("[]");
    }

    registrarPontosGamificacao(checagem.email, "EDICAO", 0, dados.mandado || "N/A", "Editou registro de mandado");
    sinalizarMudancaMandados();

    SpreadsheetApp.flush();

    // FIREBASE SYNC: Re-sincronizar mandado editado (ler linha atualizada e enviar)
    try {
      var mandadoNumero = dados.mandado || sheet.getRange(row, col['Mandado'] + 1).getValue();
      var fbPayload = {
        nome: dados.nome, cpf: dados.cpf, rg: dados.rg, nascimento: dados.nascimento,
        sexo: dados.sexo, cor: dados.cor, filiacao: dados.filiacao,
        enderecoPrincipal: dados.enderecoPrincipal, status: dados.status,
        validade: dados.validade, batalhao: dados.batalhao, cidade: dados.cidade
      };
      if (dados.fotoUrl) fbPayload.fotoUrl = dados.fotoUrl;
      atualizarMandadoFirebase(mandadoNumero, fbPayload);
    } catch(fbErr) { console.error('Firebase sync falhou na edição: ' + fbErr.message); }

    return { sucesso: true, mensagem: "Mandado atualizado com sucesso!", fotoUrl: dados.fotoUrl || null };
  } catch (erro) {
    return { sucesso: false, mensagem: "Erro ao editar: " + erro.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Exclui um mandado da planilha (somente Admin).
 * v4.1.0: Usa cabeçalhos dinâmicos.
 */
function excluirMandado(idLinha, numeroMandado) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil === "Patrulheiro") {
    return { sucesso: false, mensagem: "Acesso negado. Patrulheiros não podem excluir." };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
    if (!sheet) return { sucesso: false, mensagem: "Aba Mandados não encontrada." };

    const col = obterMapaColunas(sheet);

    let row = parseInt(idLinha);
    if ((!row || isNaN(row) || row < 2) && numeroMandado) {
      row = encontrarLinhaPorMandado_(sheet, col, numeroMandado);
    }

    if (!row || isNaN(row) || row < 2) {
      return { sucesso: false, mensagem: "Linha do registro não encontrada na planilha para exclusão." };
    }
    idLinha = row;

    if (checagem.perfil === "Colaborador") {
      const jsonExtrasBanco = sheet.getRange(idLinha, col['Dados Extras JSON'] + 1).getValue();
      try {
        const extrasBanco = JSON.parse(jsonExtrasBanco || "{}");
        if (extrasBanco.Criado_Por !== checagem.email) {
          return { sucesso: false, mensagem: "Acesso negado. Você só pode excluir mandados que você mesmo cadastrou." };
        }
      } catch (e) {
        return { sucesso: false, mensagem: "Acesso negado. Autoria do mandado não pôde ser confirmada." };
      }
    }
    
    const mandadoNome = sheet.getRange(idLinha, col['Nome'] + 1).getValue();
    const mandadoNum = sheet.getRange(idLinha, col['Mandado'] + 1).getValue();
    sheet.deleteRow(idLinha);

    // Remover também da aba Inconsistencias (se existia lá)
    removerMandadoDaAbaInconsistencias_(mandadoNum || numeroMandado);
    
    registrarPontosGamificacao(checagem.email, "EXCLUSAO", 0, "", "Excluiu mandado: " + mandadoNome);
    sinalizarMudancaMandados();

    // FIREBASE SYNC: Remover mandado do Firebase
    try { removerMandadoFirebase(String(mandadoNum || numeroMandado)); } catch(fbErr) { console.error('Firebase sync falhou na exclusão: ' + fbErr.message); }

    return { sucesso: true, mensagem: "Mandado excluído com sucesso." };
  } catch (erro) {
    return { sucesso: false, mensagem: "Erro ao excluir: " + erro.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Remove todas as ocorrências de um mandado da aba Inconsistencias.
 */
function removerMandadoDaAbaInconsistencias_(numeroMandado) {
  if (!numeroMandado) return;
  try {
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var sheetIncons = planilha.getSheetByName("Inconsistencias");
    if (!sheetIncons || sheetIncons.getLastRow() <= 1) return;

    var dadosI = sheetIncons.getDataRange().getValues();
    var hI = dadosI[0];
    var idxMandI = hI.indexOf("Mandado");
    if (idxMandI === -1) idxMandI = hI.indexOf("Nº Mandado");
    if (idxMandI >= 0) {
      var numClean = String(numeroMandado).trim();
      for (var j = dadosI.length - 1; j >= 1; j--) {
        if (String(dadosI[j][idxMandI] || "").trim() === numClean) {
          sheetIncons.deleteRow(j + 1);
          console.log("[INCONSISTENCIAS] Mandado " + numClean + " removido da aba Inconsistencias (linha " + (j + 1) + ").");
        }
      }
    }
  } catch(e) {
    console.error("Erro ao remover da aba Inconsistencias: " + e.message);
  }
}

/**
 * Atualiza status e observações de um mandado.
 * v4.1.0: Usa cabeçalhos dinâmicos em vez de índices hardcoded.
 */
function atualizarRegistro(mandado, novoStatus, novaObs) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado) {
    throw new Error("Acesso negado: Operação restrita a operadores ativos.");
  }

  // Verificar permissões do patrulheiro
  if (checagem.perfil !== "Admin") {
    const configs = obterConfiguracoesSimples();
    if (novoStatus !== undefined && configs["perm_patr_mudar_status"] !== "true") {
      throw new Error("Sem permissão para alterar status.");
    }
    if (novaObs !== undefined && configs["perm_patr_add_obs"] !== "true") {
      throw new Error("Sem permissão para adicionar observações.");
    }
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
    const col = obterMapaColunas(sheet);
    const data = sheet.getDataRange().getValues();
    let atualizados = 0;

    const idxMandado = col['Mandado'];
    const idxStatus = col['Status'];
    const idxObs = col['Observacoes'];
    const idxJSON = col['Dados Extras JSON'];

    let dataCapturaOriginal = "";
    let histJSONOriginal = "";

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxMandado]).trim() === String(mandado).trim()) {
        if (novoStatus) {
          sheet.getRange(i + 1, idxStatus + 1).setValue(novoStatus);
          if (novoStatus.toUpperCase() === "CAPTURADO") {
            const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
            dataCapturaOriginal = dataAtual;
            const idxConferencia = col['Data de Conferencia'];
            if (idxConferencia !== undefined && idxConferencia >= 0) {
              sheet.getRange(i + 1, idxConferencia + 1).setValue(dataAtual);
            }
          }
        }
        
        if (novaObs && novaObs.trim() !== "") {
          if (idxObs !== undefined && idxObs >= 0) {
            const obsAntiga = data[i][idxObs];
            let hist = [];
            if (obsAntiga && String(obsAntiga).trim() !== "") {
               const strAntiga = String(obsAntiga).trim();
               if (strAntiga.startsWith('[')) {
                 try {
                   hist = JSON.parse(strAntiga);
                 } catch(e) {
                   hist = [{ data: "Legado", usuario: "Sistema", texto: strAntiga }];
                 }
               } else {
                 hist = [{ data: "Legado", usuario: "Sistema", texto: strAntiga }];
               }
            }
            
            const dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yy - HH:mm");
            const usuarioNome = checagem.nome || checagem.email || "Usuário";
            hist.unshift({ data: dataHoje, usuario: usuarioNome, texto: novaObs });
            
            const histStr = JSON.stringify(hist);
            histJSONOriginal = histStr;
            sheet.getRange(i + 1, idxObs + 1).setValue(histStr);
          }
        }
        
        atualizados++;
      }
    }

    // Gamificação: registrar ação
    if (novoStatus && novoStatus.toUpperCase() === "CAPTURADO") {
      const configs = obterConfiguracoesSimples();
      const pontos = parseInt(configs["gamif_pontos_captura"]) || 100;
      registrarPontosGamificacao(checagem.email, "CAPTURA", pontos, mandado, "Captura confirmada");
      atualizarPontosUsuario(checagem.email, pontos);
    } else if (novaObs && novaObs.trim() !== "") {
      const configs = obterConfiguracoesSimples();
      const pontos = parseInt(configs["gamif_pontos_relato"]) || 10;
      registrarPontosGamificacao(checagem.email, "RELATO", pontos, mandado, "Relato de abordagem");
      atualizarPontosUsuario(checagem.email, pontos);
    }

    let fbMsg = "";
    if (atualizados > 0) {
      sinalizarMudancaMandados();
      // FIREBASE SYNC: Atualizar status/obs no Firebase
      try {
        var campos = {};
        if (novoStatus) campos.status = novoStatus;
        if (novoStatus && novoStatus.toUpperCase() === "CAPTURADO" && dataCapturaOriginal) {
          campos.dataConferencia = dataCapturaOriginal;
        }
        if (histJSONOriginal) {
          campos.historicoObservacoes = histJSONOriginal;
          campos.observacoes = novaObs;
        }
        atualizarMandadoFirebase(mandado, campos);
      } catch(fbErr) { 
        console.error('Firebase sync falhou no atualizarRegistro: ' + fbErr.message); 
        fbMsg = " (Aviso: Falha de sincronização com o Firebase: " + fbErr.message + ")";
      }
    }

    return { sucesso: true, mensagem: atualizados + ' pino(s) atualizado(s) com sucesso.' + fbMsg };
  } catch (erro) {
    throw new Error('Erro na atualização: ' + erro.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Admin marca a data de conferência de um mandado.
 * v4.1.0: Usa cabeçalhos dinâmicos.
 */
function conferirMandado(mandado) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    throw new Error("Acesso negado: Apenas administradores podem conferir mandados.");
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
    const col = obterMapaColunas(sheet);
    const data = sheet.getDataRange().getValues();
    const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    const idxMandado = col['Mandado'];
    const idxConferencia = col['Data de Conferencia'];

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxMandado]).trim() === String(mandado).trim()) {
        sheet.getRange(i + 1, idxConferencia + 1).setValue(dataAtual);
        sinalizarMudancaMandados();
        
        // FIREBASE SYNC
        try { atualizarMandadoFirebase(mandado, { dataConferencia: dataAtual }); } catch(fbErr) {}
        
        return { sucesso: true, mensagem: "Mandado conferido em " + dataAtual + "." };
      }
    }
    return { sucesso: false, mensagem: "Mandado não encontrado." };
  } catch (erro) {
    return { sucesso: false, mensagem: "Erro ao conferir: " + erro.message };
  } finally {
    lock.releaseLock();
  }
}

function salvarConferenciaBNMP(mandado, statusBnmp) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado) throw new Error("Acesso negado.");
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
    const col = obterMapaColunas(sheet);
    const data = sheet.getDataRange().getValues();
    const idxMandado = col['Mandado'];
    const idxJSON = col['Dados Extras JSON'];
    if (idxJSON === undefined) throw new Error("Coluna Dados Extras JSON não encontrada.");
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idxMandado]).trim() === String(mandado).trim()) {
        let extra = {};
        try { extra = JSON.parse(data[i][idxJSON] || '{}'); } catch(e){}
        let dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        let usuario = checagem.email;
        extra.bnmpConferencia = statusBnmp + " em " + dataHoje + " por " + usuario;
        sheet.getRange(i + 1, idxJSON + 1).setValue(JSON.stringify(extra));
        sinalizarMudancaMandados();
        
        // FIREBASE SYNC
        try { atualizarMandadoFirebase(mandado, { dadosExtrasJSON: JSON.stringify(extra) }); } catch(fbErr) {}

        return { mensagem: "Status do BNMP salvo com sucesso." };
      }
    }
    throw new Error("Mandado não encontrado.");
  } catch (e) {
    throw new Error("Erro ao salvar BNMP: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ================================================================
// HELPER: SALVA FOTO NO GOOGLE DRIVE
// ================================================================

function processarEDespacharFotoNoDrive(idMandado, base64Completo, nomeProcurado) {
  if (!base64Completo || base64Completo === "" || base64Completo === "N/A") return "N/A";

  try {
    let base64Limpo = base64Completo;
    if (base64Completo.includes(",")) {
      base64Limpo = base64Completo.split(",")[1];
    }
    // Remove qualquer espaço ou quebra de linha que possa quebrar o decode
    base64Limpo = (base64Limpo || "").trim().replace(/\s/g, '');
    if (!base64Limpo || base64Limpo.length < 100) {
      console.error("[FOTO] base64 muito curto ou vazio. Tamanho: " + (base64Limpo || "").length);
      return "ERRO: Base64 muito curto";
    }

    const numClean = String(idMandado || "SemNum").trim().replace(/[^a-zA-Z0-9_\-.]/g, "_");
    const nomeClean = String(nomeProcurado || "").trim().replace(/[^a-zA-Z0-9_\-.]/g, "_");
    let nomeArquivo = numClean;
    if (nomeClean && !numClean.toLowerCase().includes(nomeClean.toLowerCase())) {
      nomeArquivo = numClean + "_" + nomeClean;
    }
    if (!nomeArquivo.toLowerCase().endsWith(".jpg")) {
      nomeArquivo += ".jpg";
    }

    console.log("[FOTO] Salvando arquivo: " + nomeArquivo + " (" + base64Limpo.length + " chars base64)");

    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Limpo),
      "image/jpeg",
      nomeArquivo
    );

    const targetFolderId = "1QHK_Bc-XXIrm4dSWj7XjhfmUrWLydw--";
    const nomePasta = "Sentinela_Fotos_Mandados";
    let pasta = null;
    let origemPasta = "";

    try {
      pasta = DriveApp.getFolderById(targetFolderId);
      origemPasta = "ID";
      console.log("[FOTO] Pasta localizada por ID: " + pasta.getName());
    } catch (errId) {
      console.warn("[FOTO] Pasta por ID inacessível (" + errId.message + "). Tentando por nome...");
    }

    if (!pasta) {
      try {
        const pastas = DriveApp.getFoldersByName(nomePasta);
        if (pastas.hasNext()) {
          pasta = pastas.next();
          origemPasta = "NOME";
          console.log("[FOTO] Pasta localizada por nome: " + pasta.getName());
        } else {
          pasta = DriveApp.createFolder(nomePasta);
          origemPasta = "CRIADA";
          console.log("[FOTO] Pasta criada: " + pasta.getName() + " ID: " + pasta.getId());
        }
      } catch (errNome) {
        console.error("[FOTO] Falha ao buscar/criar pasta por nome: " + errNome.message);
        throw new Error("Falha ao acessar pastas do Drive: " + errNome.message);
      }
    }

    const arquivo = pasta.createFile(blob);

    // Tenta definir compartilhamento público
    try {
      arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (errShare) {
      console.warn("[FOTO] setSharing bloqueado (" + errShare.message + ").");
    }

    const fileId = arquivo.getId();
    const url = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w400";
    return url;

  } catch (e) {
    console.error("[FOTO] ERRO CRITICO ao salvar foto no Drive: " + e.message);
    return "ERRO: " + e.message; 
  }
}

/**
 * Função de diagnóstico — rode diretamente no editor Apps Script para testar.
 * Menu: Executar > testarSalvarFotoDrive
 */
function testarSalvarFotoDrive() {
  // Mini imagem 1x1 pixel JPEG em base64 para teste
  const base64Teste = "/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAABAAEDASIAAhEBAxEB/8QAFAABAAAAAAAAAAAAAAAAAAAACf/EABQQAQAAAAAAAAAAAAAAAAAAAAD/xAAUAQEAAAAAAAAAAAAAAAAAAAAA/8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAwDAQACEQMRAD8AJQAB/9k=";
  
  Logger.log("=== TESTE DRIVE ===");
  Logger.log("Tentando acessar pasta ID: 1QHK_Bc-XXIrm4dSWj7XjhfmUrWLydw--");
  
  try {
    const pasta = DriveApp.getFolderById("1QHK_Bc-XXIrm4dSWj7XjhfmUrWLydw--");
    Logger.log("✅ Pasta acessada com sucesso: " + pasta.getName());
  } catch(e) {
    Logger.log("❌ Erro ao acessar pasta por ID: " + e.message);
  }
  
  const resultado = processarEDespacharFotoNoDrive("TESTE_001", base64Teste, "DIAGNOSTICO");
  Logger.log("Resultado processarEDespacharFotoNoDrive: " + resultado);
  
  if (resultado && resultado !== "N/A" && !resultado.startsWith("ERRO:")) {
    Logger.log("✅ SUCESSO! URL gerada: " + resultado);
  } else {
    Logger.log("❌ FALHOU! Resultado: " + resultado);
  }
}

/**
 * Vincula fotos extraídas de PDFs em lote aos mandados já cadastrados no sistema.
 */
function vincularFotosLotePdfDrive(listaFotos) {
  if (!Array.isArray(listaFotos) || listaFotos.length === 0) {
    return { sucesso: false, mensagem: "Nenhuma foto fornecida." };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
  if (!sheet) return { sucesso: false, mensagem: "Aba Mandados não encontrada." };

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { sucesso: false, mensagem: "Planilha sem dados." };

  const headers = data[0].map(function(h) { return String(h || "").trim(); });
  const idxMandado = headers.indexOf('Mandado');
  const idxNome = headers.indexOf('Nome');
  
  let idxFoto = headers.indexOf('Foto URL');
  if (idxFoto === -1) idxFoto = headers.indexOf('Foto');
  if (idxFoto === -1) {
    for (let k = 0; k < headers.length; k++) {
      if (String(headers[k] || "").trim().toLowerCase().includes('foto')) {
        idxFoto = k;
        break;
      }
    }
  }

  if (idxMandado === -1 || idxFoto === -1) {
    return { sucesso: false, mensagem: "Coluna Mandado ou Foto não localizada na planilha." };
  }

  let vinculadas = 0;
  let naoEncontradas = 0;
  let resultados = [];

  for (let f = 0; f < listaFotos.length; f++) {
    const item = listaFotos[f];
    const numMandadoClean = String(item.mandado || "").trim();
    const base64 = item.fotoBase64;
    
    if (!numMandadoClean || !base64 || base64.length < 100) continue;

    let linhaEncontrada = -1;
    let nomePessoa = item.nome || "";

    for (let r = 1; r < data.length; r++) {
      const mandadoRow = String(data[r][idxMandado] || "").trim();
      if (mandadoRow === numMandadoClean) {
        linhaEncontrada = r + 1;
        if (!nomePessoa && idxNome !== -1) {
          nomePessoa = String(data[r][idxNome] || "").trim();
        }
        break;
      }
    }

    if (linhaEncontrada !== -1) {
      const urlDrive = processarEDespacharFotoNoDrive(numMandadoClean, base64, nomePessoa);
      if (urlDrive && urlDrive !== "N/A") {
        sheet.getRange(linhaEncontrada, idxFoto + 1).setValue(urlDrive);
        
        try {
          if (typeof atualizarMandadoNoFirebase === 'function') {
            atualizarMandadoNoFirebase(numMandadoClean, { fotoUrl: urlDrive });
          }
        } catch (errFb) { console.warn("[VINCULAR FOTO] Erro ao sincronizar com Firebase:", errFb.message); }

        vinculadas++;
        resultados.push({ mandado: numMandadoClean, fotoUrl: urlDrive, sucesso: true });
      }
    } else {
      naoEncontradas++;
      resultados.push({ mandado: numMandadoClean, sucesso: false, motivo: "Mandado não cadastrado" });
    }
  }

  SpreadsheetApp.flush();

  return {
    sucesso: true,
    mensagem: `Fotos vinculadas: ${vinculadas}. Mandados não localizados: ${naoEncontradas}.`,
    resultados: resultados,
    vinculadas: vinculadas
  };
}

/**
 * Limpa células da coluna 'Foto URL' que contenham strings de erro ou valores inválidos.
 * Execute diretamente no editor Apps Script: Executar > limparFotoUrlCorrompidas
 */
function limparFotoUrlCorrompidas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
  if (!sheet) { Logger.log("Aba Mandados não encontrada."); return; }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idxFoto = headers.indexOf("Foto URL");
  const idxMandado = headers.indexOf("Mandado");
  if (idxFoto === -1) { Logger.log("Coluna 'Foto URL' não encontrada."); return; }
  
  const data = sheet.getDataRange().getValues();
  let corrigidas = 0;
  
  for (let i = 1; i < data.length; i++) {
    const val = String(data[i][idxFoto] || "").trim();
    if (val && val !== "" && val !== "N/A" && !val.startsWith("http")) {
      sheet.getRange(i + 1, idxFoto + 1).setValue("N/A");
      Logger.log("Linha " + (i + 1) + " corrigida: '" + val.substring(0, 60) + "...' → N/A");
      
      // Sincroniza limpeza no Firebase também
      if (idxMandado !== -1) {
        const numMandado = String(data[i][idxMandado] || "").trim();
        if (numMandado) {
          try {
            atualizarMandadoFirebase(numMandado, { fotoUrl: "N/A" });
          } catch(fbErr) {
            Logger.log("Firebase sync falhou para mandado " + numMandado + ": " + fbErr.message);
          }
        }
      }
      corrigidas++;
    }
  }
  
  SpreadsheetApp.flush();
  Logger.log("Total de células corrigidas: " + corrigidas);
}

// ================================================================
// GAMIFICAÇÃO
// ================================================================

/**
 * Registra uma ação na aba de gamificação (log de pontos).
 */
function registrarPontosGamificacao(email, acao, pontos, mandadoRef, descricao) {
  try {
    const configs = obterConfiguracoesSimples();
    if (configs["gamif_ativo"] !== "true") return;

    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Gamificacao");
    const dados = abaUsuarios.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0].toString().trim().toLowerCase() === email.toLowerCase()) {
        const pontosAtuais = parseInt(dados[i][6]) || 0;
        const novoTotal = pontosAtuais + pontosAdicionais;
        abaUsuarios.getRange(i + 1, 7).setValue(novoTotal);

        // Verificar e atribuir badges
        const badgesAtuais = dados[i][7] ? dados[i][7].toString() : "";
        const novasBadges = calcularBadges(novoTotal, badgesAtuais);
        if (novasBadges !== badgesAtuais) {
          abaUsuarios.getRange(i + 1, 8).setValue(novasBadges);
        }
        return;
      }
    }
  } catch (e) {
    console.error("Erro ao atualizar pontos: " + e.message);
  }
}

/**
 * Calcula badges baseado na pontuação total.
 */
function calcularBadges(pontosTotais, badgesAtuais) {
  const listaBadges = [];

  if (pontosTotais >= 10) listaBadges.push("🔰 Iniciante");
  if (pontosTotais >= 100) listaBadges.push("⭐ Operador");
  if (pontosTotais >= 500) listaBadges.push("🏅 Veterano");
  if (pontosTotais >= 1000) listaBadges.push("🎖️ Elite");
  if (pontosTotais >= 5000) listaBadges.push("👑 Lendário");

  return listaBadges.join(" | ");
}

/**
 * Obtém o ranking de gamificação (top 20 operadores).
 */
function obterRankingGamificacao() {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado) {
    throw new Error("Acesso negado.");
  }

  const configs = obterConfiguracoesSimples();
  if (configs["gamif_ativo"] !== "true") return [];

  const abaUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
  if (!abaUsuarios) return [];

  const dados = abaUsuarios.getDataRange().getValues();
  const ranking = [];

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][4].toString().trim().toUpperCase() === "ATIVO") {
      ranking.push({
        nome: dados[i][1],
        unidade: dados[i][2],
        pontos: parseInt(dados[i][6]) || 0,
        badges: dados[i][7] ? dados[i][7].toString() : "",
        fotoUrl: dados[i][5] ? dados[i][5].toString() : ""
      });
    }
  }

  // Ordenar por pontos (maior primeiro)
  ranking.sort((a, b) => b.pontos - a.pontos);

  return ranking.slice(0, 20);
}

// ================================================================
// NOVO: CONFERÊNCIA VIA CSV (BNMP) E LOGO
// ================================================================

/**
 * Salva a logo no Google Drive e retorna a URL.
 */
function salvarLogoDrive(base64Completo) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    throw new Error("Acesso negado.");
  }
  return processarEDespacharFotoNoDrive("AppLogo_" + new Date().getTime(), base64Completo);
}

function salvarFotoIndividuoDrive(nome, mandado, base64Completo) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || (checagem.perfil !== "Admin" && checagem.perfil !== "Colaborador")) {
    throw new Error("Acesso negado.");
  }
  return processarEDespacharFotoNoDrive(mandado, base64Completo, nome);
}

/**
 * Processa cruzamento de dados com CSV do BNMP
 */
function processarCruzamentoCSV(dadosCsv) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    throw new Error("Acesso negado: Apenas administradores podem importar CSV.");
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaConferencia = planilha.getSheetByName("Conferencia");
    const abaGeral = planilha.getSheetByName("Mandados");
    
    if (!abaConferencia || !abaGeral) throw new Error("Abas necessárias não encontradas (Execute o Setup inicial).");

    // Limpar aba Conferencia (mantendo cabeçalhos)
    const ultimaLinha = abaConferencia.getLastRow();
    if (ultimaLinha > 1) {
      abaConferencia.getRange(2, 1, ultimaLinha - 1, abaConferencia.getLastColumn()).clearContent();
    }

    // Ler DB atual do Sentinela
    const dadosSentinela = abaGeral.getDataRange().getValues();
    const mandadosSentinela = {};
    
    for(let i=1; i < dadosSentinela.length; i++) {
      const mandado = String(dadosSentinela[i][2]).trim();  // C = Mandado (index 2)
      if (mandado) {
        mandadosSentinela[mandado] = { 
          status: String(dadosSentinela[i][16]).trim(),   // Q = Status
          nome: String(dadosSentinela[i][4]).trim(),      // E = Nome
          batalhao: String(dadosSentinela[i][13]).trim(), // N = Batalhao
          linhaObj: dadosSentinela[i]
        };
      }
    }

    // Identificar coluna do mandado no CSV
    const headersCsv = dadosCsv[0] || [];
    let idxMandado = -1;
    for(let i=0; i < headersCsv.length; i++) {
      const th = String(headersCsv[i]).toLowerCase();
      if (th.includes("mandado") || th.includes("número") || th.includes("numero") || th.includes("documento")) {
        idxMandado = i;
        break;
      }
    }
    if (idxMandado === -1) idxMandado = 0; // fallback para primeira coluna

    // Ler mandados do CSV
    const mandadosCsv = {};
    for (let i=1; i < dadosCsv.length; i++) {
      const linha = dadosCsv[i];
      if (linha && linha.length > idxMandado) {
        const m = String(linha[idxMandado]).trim();
        if (m) mandadosCsv[m] = true;
      }
    }

    const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    const relatorio = [];
    let contBaixados = 0;
    let contPendentes = 0;

    // Cruzar Sentinela -> CSV
    for (let mandado in mandadosSentinela) {
      const s = mandadosSentinela[mandado];
      let situacao = "PENDENTE";
      let statusCsv = "Ativo no BNMP";

      if (!mandadosCsv[mandado]) {
        situacao = "BAIXADO NO BNMP";
        statusCsv = "Não encontrado no CSV";
        contBaixados++;
      } else {
        contPendentes++;
      }

      relatorio.push([
        dataAtual,          // A - Data
        mandado,            // B - Mandado
        statusCsv,          // C - Status CSV
        s.status,           // D - Status Sentinela
        situacao,           // E - Situacao
        s.nome,             // F - Nome
        s.batalhao,         // G - Batalhao
        checagem.nome       // H - Conferido Por
      ]);
    }

    if (relatorio.length > 0) {
      abaConferencia.getRange(2, 1, relatorio.length, 8).setValues(relatorio);
    }

    return { 
      sucesso: true, 
      mensagem: "Cruzamento concluído com sucesso!", 
      baixados: contBaixados, 
      pendentes: contPendentes 
    };

  } catch (erro) {
    return { sucesso: false, mensagem: "Erro no cruzamento: " + erro.message };
  } finally {
    lock.releaseLock();
  }
}

// ================================================================
// POLÍGONOS — LEITURA DA ABA "Poligonos"
// ================================================================

/**
 * Lê todos os polígonos ativos da aba Poligonos e retorna como array de objetos.
 * Colunas: A=Setor(CPI), B=Batalhao, C=Cia, D=Cidade, E=Delegacia, F=GeoJSON, G=Cor, H=Ativo
 */
function obterPoligonos() {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado) {
    throw new Error("Acesso negado: Operação não autorizada.");
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Poligonos");
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getDataRange().getValues();
  const poligonos = [];

  for (var i = 1; i < data.length; i++) {
    var ativo = String(data[i][7] || "SIM").trim().toUpperCase();
    if (ativo !== "SIM" && ativo !== "") continue;

    var geoJsonStr = String(data[i][5] || "");
    for (var col = 8; col < data[i].length; col++) {
      if (data[i][col]) geoJsonStr += String(data[i][col]);
    }
    if (!geoJsonStr || geoJsonStr === "") continue;

    poligonos.push({
      cpi: String(data[i][0] || ""),
      batalhao: String(data[i][1] || ""),
      cia: String(data[i][2] || ""),
      cidade: String(data[i][3] || ""),
      delegacia: String(data[i][4] || ""),
      geoJson: geoJsonStr,
      cor: String(data[i][6] || "#3388ff")
    });
  }

  return poligonos;
}

// ================================================================
// HISTÓRICO DE AÇÕES
// ================================================================

/**
 * Registra uma ação do usuário na aba Historico.
 */
function registrarHistorico(email, nome, acao, detalhes) {
  try {
    var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Historico");
    if (!aba) return;

    var dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    aba.appendRow([
      dataAtual,       // A — Data
      email,           // B — Email
      nome || "",      // C — Nome
      acao,            // D — Ação
      detalhes || "",  // E — Detalhes
      ""               // F — IP/Dispositivo (preenchido via frontend se possível)
    ]);
  } catch (e) {
    Logger.log("Erro ao registrar historico: " + e.message);
  }
}

/**
 * Retorna as últimas 100 ações do histórico.
 */
function obterHistorico() {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    throw new Error("Acesso negado.");
  }

  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Historico");
  if (!aba || aba.getLastRow() <= 1) return [];

  var data = aba.getDataRange().getValues();
  var historico = [];

  for (var i = data.length - 1; i >= 1; i--) {
    historico.push({
      data: String(data[i][0] || ""),
      email: String(data[i][1] || ""),
      nome: String(data[i][2] || ""),
      acao: String(data[i][3] || ""),
      detalhes: String(data[i][4] || ""),
      dispositivo: String(data[i][5] || "")
    });
    if (historico.length >= 100) break;
  }

  return historico;
}

// ================================================================
// MIGRAÇÃO: Renomear headers e adicionar colunas faltantes
// ================================================================

/**
 * Execute esta função UMA VEZ para atualizar os headers da aba Mandados 
 * e adicionar a coluna Cidade (AC) em planilhas existentes.
 */
function atualizarHeadersMandados() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName("Mandados");
  if (!aba) {
    Logger.log("Aba Mandados nao encontrada.");
    return;
  }

  var headersNovos = [
    "Data de Lancamento", "Data de Conferencia", "Mandado", "Artigo", "Nome",
    "CPF", "RG", "Nascimento", "Naturalidade", "Sexo", "Cor", "Filiacao",
    "Foto URL", "Batalhao", "Endereco Principal", "Outros Enderecos",
    "Status", "Validade", "Info Processuais", "Geodata Secundarios",
    "Dados Extras JSON", "Observacoes", "Latitude", "Longitude",
    "CPI", "BPM Area", "CIA Area", "DP Area", "Cidade", "TipoImportacao"
  ];

  aba.getRange(1, 1, 1, headersNovos.length).setValues([headersNovos]);
  aba.getRange(1, 1, 1, headersNovos.length)
    .setFontWeight("bold")
    .setBackground("#1e293b")
    .setFontColor("#e2e8f0");
  aba.setFrozenRows(1);

  // Criar aba Historico se não existir
  var abaHistorico = planilha.getSheetByName("Historico");
  if (!abaHistorico) {
    abaHistorico = planilha.insertSheet("Historico");
    var headersHist = ["Data", "Email", "Nome", "Acao", "Detalhes", "IP/Dispositivo"];
    abaHistorico.getRange(1, 1, 1, headersHist.length).setValues([headersHist]);
    abaHistorico.getRange(1, 1, 1, headersHist.length)
      .setFontWeight("bold")
      .setBackground("#1e293b")
      .setFontColor("#e2e8f0");
    abaHistorico.setFrozenRows(1);
  }

  // Criar aba Leis se não existir
  var abaLeis = planilha.getSheetByName("Leis");
  if (!abaLeis) {
    abaLeis = planilha.insertSheet("Leis");
    var headersLeis = ["Categoria", "Palavras Chave", "Cor", "Icone SVG", "Ordem", "Ativo"];
    abaLeis.getRange(1, 1, 1, headersLeis.length).setValues([headersLeis]);
    abaLeis.getRange(1, 1, 1, headersLeis.length)
      .setFontWeight("bold")
      .setBackground("#1e293b")
      .setFontColor("#e2e8f0");
    abaLeis.setFrozenRows(1);
  }

  // Criar aba Notificacoes se não existir
  var abaNotif = planilha.getSheetByName("Notificacoes");
  if (!abaNotif) {
    abaNotif = planilha.insertSheet("Notificacoes");
    var headersNotif = ["Data", "Titulo", "Mensagem", "Tipo", "Para", "Lida", "De"];
    abaNotif.getRange(1, 1, 1, headersNotif.length).setValues([headersNotif]);
    abaNotif.getRange(1, 1, 1, headersNotif.length)
      .setFontWeight("bold")
      .setBackground("#1e293b")
      .setFontColor("#e2e8f0");
    abaNotif.setFrozenRows(1);
  }

  Logger.log("Headers atualizados e abas criadas com sucesso!");
  SpreadsheetApp.getUi().alert("Migracao v2.5 concluida!\n\n- Coluna AC (Cidade) adicionada\n- Aba Historico criada\n- Aba Leis criada\n- Aba Notificacoes criada\n- Todos os headers renomeados");
}

// ================================================================
// LEIS — CRUD (Admin)
// ================================================================

/**
 * Retorna todas as leis/artigos cadastrados (Admin).
 */
function obterMapaCabecalhos(headers) {
  var mapa = {};
  for (var i = 0; i < headers.length; i++) {
    var nome = String(headers[i]).trim();
    if (nome) {
      mapa[nome] = i;
    }
  }
  return mapa;
}

function garantirColunaPinoPiscando(aba) {
  if (!aba || aba.getLastRow() < 1) return;
  var headers = aba.getRange(1, 1, 1, aba.getLastColumn()).getValues()[0];
  var mapa = obterMapaCabecalhos(headers);
  if (mapa["PinoPiscando"] === undefined) {
    var colIdx = headers.length + 1;
    aba.getRange(1, colIdx).setValue("PinoPiscando").setFontWeight("bold").setBackground("#1e293b").setFontColor("#e2e8f0");
  }
}

function obterLeis() {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado) throw new Error("Acesso negado.");

  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base_Leis");
  if (!aba) {
    aba = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Base_Leis");
    var headers = ["ID", "Categoria", "Palavras Chave", "Cor", "Ativo", "PinoTexto", "Lei Nome", "Numero Lei", "Artigo", "Paragrafo", "Inciso", "Tipificacao Completa", "PinoPiscando"];
    aba.getRange(1, 1, 1, headers.length).setValues([headers]);
    aba.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#1e293b").setFontColor("#e2e8f0");
    aba.setFrozenRows(1);
  }

  garantirColunaPinoPiscando(aba);

  if (aba.getLastRow() <= 1) {
    var prepopulatedLeis = obterPrepopulatedLeisLocal();
    if (prepopulatedLeis && prepopulatedLeis.length > 0) {
      aba.getRange(2, 1, prepopulatedLeis.length, 12).setValues(prepopulatedLeis);
      SpreadsheetApp.flush();
    }
  }

  var data = aba.getDataRange().getValues();
  var headers = data[0];
  var mapa = obterMapaCabecalhos(headers);

  var getVal = function(linha, colNome, fallback) {
    var idx = mapa[colNome];
    if (idx === undefined) return fallback !== undefined ? fallback : "";
    var v = linha[idx];
    return v !== undefined ? String(v) : (fallback !== undefined ? fallback : "");
  };

  var leis = [];

  for (var i = 1; i < data.length; i++) {
    var linha = data[i];
    leis.push({
      id: getVal(linha, "ID"),
      categoria: getVal(linha, "Categoria"),
      palavrasChave: getVal(linha, "Palavras Chave"),
      cor: getVal(linha, "Cor", "#6b7280"),
      ativo: getVal(linha, "Ativo", "SIM").toUpperCase() === "SIM",
      pinoTexto: getVal(linha, "PinoTexto"),
      leiNome: getVal(linha, "Lei Nome"),
      numeroLei: getVal(linha, "Numero Lei"),
      artigo: getVal(linha, "Artigo"),
      paragrafo: getVal(linha, "Paragrafo"),
      inciso: getVal(linha, "Inciso"),
      tipificacaoCompleta: getVal(linha, "Tipificacao Completa"),
      pinoPiscando: getVal(linha, "PinoPiscando", "NAO").toUpperCase() === "SIM"
    });
  }

  leis.sort(function(a, b) { return a.categoria.localeCompare(b.categoria); });
  return leis;
}

/**
 * Retorna tabela simplificada de classificação para o frontend.
 * Qualquer usuário logado pode chamar (não precisa ser Admin).
 * Carregada 1x no login e cacheada em window._tabelaLeis.
 */
function obterTabelaClassificacao() {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado) throw new Error("Acesso negado.");

  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base_Leis");
  if (!aba) return [];
  if (aba.getLastRow() <= 1) return [];

  garantirColunaPinoPiscando(aba);
  var data = aba.getDataRange().getValues();
  var headers = data[0];
  var mapa = obterMapaCabecalhos(headers);

  var getVal = function(linha, colNome, fallback) {
    var idx = mapa[colNome];
    if (idx === undefined) return fallback !== undefined ? fallback : "";
    var v = linha[idx];
    return v !== undefined ? String(v) : (fallback !== undefined ? fallback : "");
  };

  var tabela = [];

  for (var i = 1; i < data.length; i++) {
    var linha = data[i];
    var ativo = getVal(linha, "Ativo", "SIM").toUpperCase();
    if (ativo !== "SIM") continue;

    tabela.push({
      categoria: getVal(linha, "Categoria"),
      palavrasChave: getVal(linha, "Palavras Chave"),
      cor: getVal(linha, "Cor", "#6b7280"),
      pinoTexto: getVal(linha, "PinoTexto"),
      leiNome: getVal(linha, "Lei Nome"),
      numeroLei: getVal(linha, "Numero Lei"),
      artigo: getVal(linha, "Artigo"),
      paragrafo: getVal(linha, "Paragrafo"),
      inciso: getVal(linha, "Inciso"),
      tipificacaoCompleta: getVal(linha, "Tipificacao Completa"),
      pinoPiscando: getVal(linha, "PinoPiscando", "NAO").toUpperCase() === "SIM"
    });
  }

  tabela.sort(function(a, b) { return a.categoria.localeCompare(b.categoria); });
  return tabela;
}

/**
 * Alterna o status de REVISADO de um mandado no banco de dados.
 */
function alternarRevisadoMandado(mandadoOuId, revisado) {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado) throw new Error("Acesso negado.");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
  if (!sheet) throw new Error("Aba Mandados não encontrada.");

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colMap = obterMapaCabecalhos(headers);
  var idxJSON = colMap['Dados Extras JSON'];
  var idxMandado = colMap['Mandado'];

  if (idxJSON === undefined) throw new Error("Coluna Dados Extras JSON não encontrada.");

  var linhaEncontrada = -1;
  var mandadoNum = "";
  for (var i = 1; i < data.length; i++) {
    var idL = i + 1;
    var m = String(data[i][idxMandado] || "").trim();
    if (String(idL) === String(mandadoOuId) || m === String(mandadoOuId).trim()) {
      linhaEncontrada = idL;
      mandadoNum = m;
      break;
    }
  }

  if (linhaEncontrada === -1) throw new Error("Mandado não encontrado.");

  var currentJsonStr = String(sheet.getRange(linhaEncontrada, idxJSON + 1).getValue() || "{}");
  var obj = {};
  try { obj = JSON.parse(currentJsonStr); } catch(e) {}

  obj.revisado = !!revisado;
  obj.revisadoPor = checagem.nome || checagem.email || "Admin";
  obj.dataRevisado = Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss");

  var jsonStr = JSON.stringify(obj);
  sheet.getRange(linhaEncontrada, idxJSON + 1).setValue(jsonStr);

  // Garantir a coluna AE (Coluna 31 - "REVISÃO") e escrever "REVISADO" ou "REVISAR"
  var idxRevisao = colMap['REVISÃO'] !== undefined ? colMap['REVISÃO'] :
                   (colMap['Revisão'] !== undefined ? colMap['Revisão'] :
                   (colMap['REVISAO'] !== undefined ? colMap['REVISAO'] :
                   (colMap['Revisao'] !== undefined ? colMap['Revisao'] :
                   (colMap['Revisado'] !== undefined ? colMap['Revisado'] : undefined))));

  if (idxRevisao === undefined) {
    idxRevisao = 30; // Coluna 31 (AE)
    var headerCell = sheet.getRange(1, 31);
    if (String(headerCell.getValue() || "").trim() === "") {
      headerCell.setValue("REVISÃO").setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
    }
  }

  sheet.getRange(linhaEncontrada, idxRevisao + 1).setValue(revisado ? "REVISADO" : "REVISAR");

  if (typeof atualizarMandadoFirebase === 'function' && mandadoNum) {
    try {
      atualizarMandadoFirebase(mandadoNum, {
        dadosExtrasJSON: jsonStr,
        revisado: !!revisado,
        revisadoPor: obj.revisadoPor,
        dataRevisado: obj.dataRevisado
      });
    } catch(eFb) {
      console.warn("Erro ao atualizar Firebase em alternarRevisadoMandado:", eFb);
    }
  }

  try {
    registrarHistorico(checagem.email, checagem.nome, "REVISAO", (revisado ? "Marcou" : "Desmarcou") + " mandado como revisado: " + (mandadoNum || mandadoOuId));
  } catch(eHist) {}

  return { sucesso: true, revisado: obj.revisado, usuario: obj.revisadoPor, data: obj.dataRevisado };
}

/**
 * Salva ou atualiza uma lei (Admin only).
 * Se id > 0, atualiza. Se id === 0, cria nova.
 */
function salvarLei(dados) {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    return { sucesso: false, mensagem: "Acesso negado." };
  }

  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base_Leis");
  if (!aba) return { sucesso: false, mensagem: "Aba Base_Leis nao encontrada." };

  garantirColunaPinoPiscando(aba);
  var data = aba.getDataRange().getValues();
  var headers = data[0];
  var mapa = obterMapaCabecalhos(headers);

  var id = dados.id;
  var isNovo = false;
  if (!id || id === "0" || id === "") {
    id = Utilities.getUuid();
    isNovo = true;
  }

  // Localizar linha correspondente se for edição
  var linhaNum = -1;
  if (!isNovo) {
    var idColIdx = mapa["ID"];
    if (idColIdx !== undefined) {
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][idColIdx]).trim() === String(id).trim()) {
          linhaNum = i + 1;
          break;
        }
      }
    }
  }

  // Construir a linha com base no posicionamento atual dos cabeçalhos na planilha
  var novaLinha = new Array(headers.length).fill("");
  var setVal = function(colNome, val) {
    var idx = mapa[colNome];
    if (idx !== undefined) {
      novaLinha[idx] = val;
    }
  };

  setVal("ID", id);
  setVal("Categoria", dados.categoria || "");
  setVal("Palavras Chave", dados.palavrasChave || "");
  setVal("Cor", dados.cor || "#6b7280");
  setVal("Ativo", dados.ativo ? "SIM" : "NAO");
  setVal("PinoTexto", dados.pinoTexto || "");
  setVal("Lei Nome", dados.leiNome || "");
  setVal("Numero Lei", dados.numeroLei || "");
  setVal("Artigo", dados.artigo || "");
  setVal("Paragrafo", dados.paragrafo || "");
  setVal("Inciso", dados.inciso || "");
  setVal("Tipificacao Completa", dados.tipificacaoCompleta || "");
  setVal("PinoPiscando", dados.pinoPiscando ? "SIM" : "NAO");

  if (linhaNum !== -1) {
    aba.getRange(linhaNum, 1, 1, headers.length).setValues([novaLinha]);
  } else {
    aba.appendRow(novaLinha);
  }

  registrarHistorico(checagem.email, checagem.nome, "LEIS", "Salvou lei: " + dados.categoria + " (ID: " + id + ")");
  return { sucesso: true, mensagem: "Lei salva com sucesso!", id: id };
}

/**
 * Adiciona palavras-chave a uma lei existente (Admin only).
 * Usado pelo mini-modal de classificação rápida.
 */
function adicionarPalavrasChaveLei(categoria, novasPalavras) {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    return { sucesso: false, mensagem: "Acesso negado." };
  }

  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base_Leis");
  if (!aba) return { sucesso: false, mensagem: "Aba Base_Leis nao encontrada." };

  var data = aba.getDataRange().getValues();
  var headers = data[0];
  var mapa = obterMapaCabecalhos(headers);
  var catColIdx = mapa["Categoria"];
  var kwColIdx = mapa["Palavras Chave"];

  if (catColIdx === undefined || kwColIdx === undefined) {
    return { sucesso: false, mensagem: "Colunas Categoria ou Palavras Chave nao encontradas." };
  }

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][catColIdx]).trim().toUpperCase() === String(categoria).trim().toUpperCase()) {
      var atual = String(data[i][kwColIdx] || "");
      var novas = String(novasPalavras || "").split("|").map(function(p) { return p.trim().toUpperCase(); }).filter(function(p) { return p; });
      var existentes = atual.toUpperCase().split("|").map(function(p) { return p.trim(); });
      
      // Só adicionar palavras que não existem ainda
      novas.forEach(function(n) {
        if (existentes.indexOf(n) === -1) {
          atual += (atual ? "|" : "") + n;
        }
      });
      
      aba.getRange(i + 1, kwColIdx + 1).setValue(atual);
      registrarHistorico(checagem.email, checagem.nome, "LEIS", "Adicionou keywords em " + categoria + ": " + novasPalavras);
      return { sucesso: true, mensagem: "Palavras-chave adicionadas!" };
    }
  }

  return { sucesso: false, mensagem: "Categoria nao encontrada: " + categoria };
}

/**
 * Exclui uma lei (Admin only).
 */
function excluirLei(idLinha) {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    return { sucesso: false, mensagem: "Acesso negado." };
  }

  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base_Leis");
  if (!aba || idLinha < 2) return { sucesso: false, mensagem: "Linha invalida." };

  var headers = aba.getDataRange().getValues()[0];
  var mapa = obterMapaCabecalhos(headers);
  var catColIdx = mapa["Categoria"] !== undefined ? mapa["Categoria"] : 1;

  var nome = aba.getRange(idLinha, catColIdx + 1).getValue();
  aba.deleteRow(idLinha);

  registrarHistorico(checagem.email, checagem.nome, "LEIS", "Excluiu categoria: " + nome);
  return { sucesso: true, mensagem: "Lei excluida com sucesso!" };
}

function deletarLei(id) {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    return { sucesso: false, mensagem: "Acesso negado." };
  }

  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base_Leis");
  if (!aba) return { sucesso: false, mensagem: "Aba Base_Leis nao encontrada." };

  var data = aba.getDataRange().getValues();
  var headers = data[0];
  var mapa = obterMapaCabecalhos(headers);
  var idColIdx = mapa["ID"];
  var catColIdx = mapa["Categoria"];

  if (idColIdx === undefined) return { sucesso: false, mensagem: "Coluna ID nao encontrada." };

  var linhaNum = -1;
  var categoria = "";
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idColIdx]).trim() === String(id).trim()) {
      linhaNum = i + 1;
      if (catColIdx !== undefined) categoria = String(data[i][catColIdx]);
      break;
    }
  }

  if (linhaNum !== -1) {
    aba.deleteRow(linhaNum);
    registrarHistorico(checagem.email, checagem.nome, "LEIS", "Excluiu lei: " + categoria + " (ID: " + id + ")");
    return { sucesso: true, message: "Lei excluída com sucesso!", sucesso: true, mensagem: "Lei excluída com sucesso!" };
  }

  return { sucesso: false, mensagem: "Lei não encontrada." };
}

function obterFiltrosUnicos() {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado) throw new Error("Acesso negado.");

  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base_Leis");
  if (!aba) return { categorias: [], leisNomes: [] };

  var data = aba.getDataRange().getValues();
  var headers = data[0];
  var mapa = obterMapaCabecalhos(headers);
  var catColIdx = mapa["Categoria"];
  var leiColIdx = mapa["Lei Nome"];

  var categoriasSet = new Set();
  var leisNomesSet = new Set();

  for (var i = 1; i < data.length; i++) {
    var cat = catColIdx !== undefined ? String(data[i][catColIdx] || "").trim() : "";
    var lei = leiColIdx !== undefined ? String(data[i][leiColIdx] || "").trim() : "";
    if (cat) categoriasSet.add(cat);
    if (lei) leisNomesSet.add(lei);
  }

  return {
    categorias: Array.from(categoriasSet).sort(),
    leisNomes: Array.from(leisNomesSet).sort()
  };
}

// ================================================================
// NOTIFICAÇÕES — CRUD
// ================================================================

/**
 * Cria uma notificação (Admin only).
 */
function criarNotificacao(titulo, mensagem, tipo, para) {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    return { sucesso: false, mensagem: "Acesso negado." };
  }

  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Notificacoes");
  if (!aba) return { sucesso: false, mensagem: "Aba Notificacoes nao encontrada." };

  var dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  aba.appendRow([
    dataAtual,
    titulo || "Aviso",
    mensagem || "",
    tipo || "info",
    para || "TODOS",
    "NAO",
    checagem.email
  ]);

  return { sucesso: true, mensagem: "Notificacao criada!" };
}

/**
 * Retorna notificações do usuário logado (últimas 20, não lidas primeiro).
 */
function obterNotificacoes() {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado) return [];

  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Notificacoes");
  if (!aba || aba.getLastRow() <= 1) return [];

  var data = aba.getDataRange().getValues();
  var notificacoes = [];
  var contNaoLidas = 0;

  for (var i = data.length - 1; i >= 1; i--) {
    var para = String(data[i][4] || "").trim().toUpperCase();
    if (para !== "TODOS" && para.toLowerCase() !== checagem.email.toLowerCase()) continue;

    var lida = String(data[i][5] || "").trim().toUpperCase() === "SIM";
    if (!lida) contNaoLidas++;

    notificacoes.push({
      id: i + 1,
      data: String(data[i][0] || ""),
      titulo: String(data[i][1] || ""),
      mensagem: String(data[i][2] || ""),
      tipo: String(data[i][3] || "info"),
      lida: lida,
      de: String(data[i][6] || "")
    });

    if (notificacoes.length >= 20) break;
  }

  return { lista: notificacoes, naoLidas: contNaoLidas };
}

/**
 * Marca uma notificação como lida.
 */
function marcarNotificacaoLida(idLinha) {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado) return;

  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Notificacoes");
  if (!aba || idLinha < 2) return;

  aba.getRange(idLinha, 6).setValue("SIM");
}

/**
 * Marca todas as notificações como lidas para o usuário logado.
 */
function marcarTodasNotificacoesLidas() {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado) return;

  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Notificacoes");
  if (!aba || aba.getLastRow() <= 1) return;

  var data = aba.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var para = String(data[i][4] || "").trim().toUpperCase();
    var lida = String(data[i][5] || "").trim().toUpperCase();
    if ((para === "TODOS" || para.toLowerCase() === checagem.email.toLowerCase()) && lida !== "SIM") {
      aba.getRange(i + 1, 6).setValue("SIM");
    }
  }
}

// ================================================================
// GAMIFICAÇÃO
// ================================================================

/**
 * Registra pontos para o usuário em um ranking local e atualiza o histórico global
 */
function registrarPontosGamificacao(email, acao, pontosExtras, referencia, obs) {
  try {
    const configs = obterConfiguracoesSimples();
    if (configs["gamif_ativo"] !== "true") return;

    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Gamificacao");
    if (!aba) return;

    const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

    aba.appendRow([
      dataAtual,      // A — Data
      email,          // B — Email
      acao,           // C — Ação
      pontosExtras,   // D — Pontos
      referencia,     // E — Mandado Ref
      obs             // F — Descrição
    ]);
  } catch (e) {
    console.error("Erro ao registrar gamificação: " + e.message);
  }
}

/**
 * Atualiza o placar de pontos do usuário logado na aba "Usuarios"
 */
function atualizarPontosUsuario(email, pontosGanhos) {
  try {
    const abaUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
    const dados = abaUsuarios.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][1] === email) {  // Coluna B = Email
        const ptAtual = parseInt(dados[i][5]) || 0; // Coluna F = Pontos
        abaUsuarios.getRange(i + 1, 6).setValue(ptAtual + pontosGanhos);
        break;
      }
    }
  } catch (e) {
    console.error("Erro ao atualizar pontos na aba Usuários: " + e.message);
  }
}

// ================================================================
// SISTEMA DE CACHE E SINCRONIZAÇÃO INTELIGENTE DE MANDADOS
// ================================================================

/**
 * Atualiza a versão interna dos mandados, disparada em cada alteração física.
 */
function sinalizarMudancaMandados() {
  try {
    const ts = new Date().getTime().toString();
    const props = PropertiesService.getScriptProperties();
    props.setProperty('MandadosLastUpdate', ts);
    props.setProperty('DB_UPDATE_TIMESTAMP', ts);
  } catch(e) {}
  // Sync Firebase metadata (não bloqueia se falhar)
  try { sinalizarMudancaFirebase_(); } catch(e) { console.error('Firebase metadata sync falhou: ' + e.message); }
}

function limparCacheBD() {
  CacheService.getScriptCache().remove('marcadoresAtivos');
  CacheService.getScriptCache().remove('dadosDashboard');
}

// ==============================================================================
// MÓDULO 6: EDIÇÃO COMPLETA E UPLOAD DE FOTOS (ADMIN/COLAB)
// ==============================================================================

/**
 * Salva a edição completa do mandado feita pelo front-end (Administrador/Colaborador)
 * @param {Object} mandado Objeto com os dados a serem atualizados (deve conter numeroMandado).
 * @param {boolean} marcarConferido Se true, registra a conferência com a data de hoje.
 */
function salvarEdicaoCompletaCard(mandado, marcarConferido) {
  if (!mandado || !mandado.numeroMandado) {
    throw new Error("Dados inválidos. Número do mandado ausente.");
  }

  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || (checagem.perfil !== "Admin" && checagem.perfil !== "Colaborador")) {
    throw new Error("Sem permissão para editar os mandados.");
  }
  
  const usuario = checagem.nome;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const idxMandado = headers.indexOf('Mandado') !== -1 ? headers.indexOf('Mandado') : headers.indexOf('Nº Mandado');
  const idxInfoProcessuais = headers.indexOf('Info Processuais');
  const idxDadosExtrasJSON = headers.indexOf('Dados Extras JSON');
  const idxOutrosEnderecos = headers.indexOf('Outros Enderecos');
  const idxGeodataSec = headers.indexOf('Geodata Secundarios');
  
  if (idxMandado === -1) throw new Error("Coluna 'Mandado' não encontrada na base.");

  let linhaAtualizar = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxMandado]).trim() === String(mandado.numeroMandado).trim()) {
      linhaAtualizar = i + 1;
      break;
    }
  }

  if (linhaAtualizar === -1) {
    throw new Error("Mandado " + mandado.numeroMandado + " não encontrado na base de dados.");
  }
  
  // Se foi enviada foto em Base64 na edição, salvar no Drive e atualizar fotoUrl
  const idxFotoCol = (function() {
    let i = headers.indexOf('Foto URL');
    if (i !== -1) return i;
    i = headers.indexOf('Foto');
    if (i !== -1) return i;
    for (let k = 0; k < headers.length; k++) {
      if (String(headers[k] || "").trim().toLowerCase().includes('foto')) return k;
    }
    return -1;
  })();

  if (mandado.fotoBase64 && mandado.fotoBase64 !== "" && !mandado.fotoBase64.startsWith("http")) {
    let novaUrlFoto = processarEDespacharFotoNoDrive(mandado.numeroMandado, mandado.fotoBase64, mandado.nome);
    if (novaUrlFoto && novaUrlFoto !== "N/A") {
      mandado.fotoUrl = novaUrlFoto;
    }
  }

  if (mandado.fotoUrl && idxFotoCol !== -1) {
    sheet.getRange(linhaAtualizar, idxFotoCol + 1).setValue(mandado.fotoUrl);
  }

  SpreadsheetApp.flush();

  // Mapeamento de colunas para dados simples.
  const colunasSimples = {
    'Nome': mandado.nome,
    'CPF': mandado.cpf,
    'RG': mandado.rg,
    'Nascimento': mandado.nascimento,
    'Sexo': mandado.sexo,
    'Cor': mandado.cor,
    'Filiacao': mandado.filiacao,
    'Naturalidade': mandado.naturalidade,
    'Foto URL': mandado.fotoUrl,
    'Endereco Principal': mandado.enderecoPrincipal,
    'Status': mandado.status,
    'Validade': mandado.validade,
    'Batalhao': mandado.batalhao,
    'CPI': mandado.cpiArea,
    'Cidade': mandado.cidade
  };
  
  for (const [colName, val] of Object.entries(colunasSimples)) {
    const idx = headers.indexOf(colName);
    if (idx !== -1 && val !== undefined) {
      sheet.getRange(linhaAtualizar, idx + 1).setValue(val);
    }
  }

  // Atualizar Infos Processuais
  if (mandado.infoProcessuais !== undefined && idxInfoProcessuais !== -1) {
    sheet.getRange(linhaAtualizar, idxInfoProcessuais + 1).setValue(mandado.infoProcessuais);
  }

  // Atualizar Endereços Secundários (Outros Endereços e Geodata)
  if (mandado.enderecosSecundarios && Array.isArray(mandado.enderecosSecundarios)) {
    const strSecundarios = mandado.enderecosSecundarios.join('\n');
    if (idxOutrosEnderecos !== -1) {
      sheet.getRange(linhaAtualizar, idxOutrosEnderecos + 1).setValue(strSecundarios);
    }
    if (idxGeodataSec !== -1) {
      const geoList = mandado.enderecosSecundarios.map(e => ({ endereco: e, enderecoOriginal: e, lat: null, lng: null }));
      sheet.getRange(linhaAtualizar, idxGeodataSec + 1).setValue(JSON.stringify(geoList));
    }
  }

  // Atualizar Dados Extras JSON e Marcação de Conferência
  if (idxDadosExtrasJSON !== -1) {
    let extraObj = {};
    try {
      const extraStr = sheet.getRange(linhaAtualizar, idxDadosExtrasJSON + 1).getValue();
      if (extraStr) extraObj = JSON.parse(extraStr);
    } catch(e) {}
    
    if (mandado.tipoMandado !== undefined) extraObj.titulo = mandado.tipoMandado;
    if (mandado.dataBNMP !== undefined) extraObj.emissão = mandado.dataBNMP;
    
    // Preservar infos extras do painel se mandadas
    if (mandado.cpf !== undefined) extraObj.cpf = mandado.cpf;
    if (mandado.rg !== undefined) extraObj.rg = mandado.rg;
    if (mandado.naturalidade !== undefined) extraObj.naturalidade = mandado.naturalidade;
    if (mandado.sexo !== undefined) extraObj.sexo = mandado.sexo;
    if (mandado.cor !== undefined) extraObj.cor = mandado.cor;
    
    if (marcarConferido) {
      const hoje = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
      extraObj.bnmpConferencia = "Conferido/Editado em " + hoje + " por " + usuario;
      const idxConf = headers.indexOf('Data de Conferencia');
      if (idxConf !== -1) {
        sheet.getRange(linhaAtualizar, idxConf + 1).setValue(hoje.split(' ')[0]);
      }
    }
    
    sheet.getRange(linhaAtualizar, idxDadosExtrasJSON + 1).setValue(JSON.stringify(extraObj));
  }
  
  sinalizarMudancaMandados();
  limparCacheBD();

  let fbMsg = "";
  // FIREBASE SYNC: Re-sincronizar mandado via objeto atualizado
  try {
    var fbData = {
      nome: mandado.nome, cpf: mandado.cpf, rg: mandado.rg, nascimento: mandado.nascimento,
      sexo: mandado.sexo, cor: mandado.cor, filiacao: mandado.filiacao,
      naturalidade: mandado.naturalidade, fotoUrl: mandado.fotoUrl,
      enderecoPrincipal: mandado.enderecoPrincipal, status: mandado.status,
      validade: mandado.validade, batalhao: mandado.batalhao, cpi: mandado.cpiArea,
      cidade: mandado.cidade
    };
    if (mandado.infoProcessuais !== undefined) fbData.infoProcessuais = mandado.infoProcessuais;
    if (mandado.enderecosSecundarios && Array.isArray(mandado.enderecosSecundarios)) {
      fbData.outrosEnderecos = mandado.enderecosSecundarios.join('\n');
      fbData.geodataSecundarios = JSON.stringify(mandado.enderecosSecundarios.map(e => ({ endereco: e, enderecoOriginal: e, lat: null, lng: null })));
    }
    if (idxDadosExtrasJSON !== -1 && typeof extraObj !== 'undefined') fbData.dadosExtrasJSON = JSON.stringify(extraObj);
    if (marcarConferido) fbData.dataConferencia = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm").split(' ')[0];
    
    atualizarMandadoFirebase(mandado.numeroMandado, fbData);
  } catch(e) { 
    console.error('Firebase sync falhou no salvarEdicaoCompletaCard: ' + e.message); 
    fbMsg = " (Aviso: Falha de sincronização com o Firebase: " + e.message + ")";
  }

  return { mensagem: "Mandado " + mandado.numeroMandado + " atualizado com sucesso!" + fbMsg };
}

/**
 * Faz o upload de uma imagem em base64 para a pasta específica do Google Drive
 * Retorna a URL pública de compartilhamento para ser setada no card.
 */
function uploadFotoIndividuoDrive(base64Data, nomeIndividuo, numeroMandado) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || (checagem.perfil !== "Admin" && checagem.perfil !== "Colaborador")) {
    throw new Error("Sem permissão para fazer upload de fotos.");
  }

  return processarEDespacharFotoNoDrive(numeroMandado, base64Data, nomeIndividuo);
}


/**
 * Trigger simples do Google Sheets. Caso alguém altere direto na planilha.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  
  if (sheetName === "Mandados") {
    sinalizarMudancaMandados();
    
    var row = e.range.getRow();
    if (row > 1) {
      try {
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        const idx = {};
        const nomesEsperados = [
          'Data de Lancamento', 'Data de Conferencia', 'Mandado', 'Artigo', 'Nome',
          'CPF', 'RG', 'Nascimento', 'Naturalidade', 'Sexo', 'Cor', 'Filiacao',
          'Foto URL', 'Batalhao', 'Endereco Principal', 'Outros Enderecos', 'Status', 'Validade',
          'Info Processuais', 'Geodata Secundarios', 'Dados Extras JSON', 'Observacoes',
          'Latitude', 'Longitude', 'CPI', 'BPM Area', 'CIA Area', 'DP Area', 'Cidade',
          'TipoImportacao'
        ];
        nomesEsperados.forEach(function(nome) {
          idx[nome] = headers.indexOf(nome);
        });

        const col = (nome, fallback) => idx[nome] >= 0 ? values[idx[nome]] : (fallback !== undefined ? fallback : "");
        
        const formatarData = (valor) => {
          if (valor instanceof Date) {
            return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
          }
          return String(valor || "").trim();
        };

        const formatarDataHora = (valor) => {
          if (valor instanceof Date) {
            return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
          }
          return String(valor || "").trim();
        };

        const latRaw = idx.Latitude >= 0 ? values[idx.Latitude] : null;
        const lngRaw = idx.Longitude >= 0 ? values[idx.Longitude] : null;
        let lat = null, lng = null;
        if (latRaw !== undefined && latRaw !== null && latRaw !== "") {
          let pLat = parseFloat(latRaw);
          if (!isNaN(pLat)) lat = pLat;
        }
        if (lngRaw !== undefined && lngRaw !== null && lngRaw !== "") {
          let pLng = parseFloat(lngRaw);
          if (!isNaN(pLng)) lng = pLng;
        }

        const mandadoObj = {
          idLinha: row,
          dataLancamento: formatarData(col('Data de Lancamento')),
          dataConferencia: formatarDataHora(col('Data de Conferencia')),
          mandado: String(col('Mandado') || ""),
          artigo: String(col('Artigo') || ""),
          nome: String(col('Nome') || ""),
          cpf: String(col('CPF') || ""),
          rg: String(col('RG') || "N/A"),
          nascimento: formatarData(col('Nascimento')),
          naturalidade: String(col('Naturalidade') || "N/A"),
          sexo: String(col('Sexo') || "N/A"),
          cor: String(col('Cor') || "N/A"),
          filiacao: String(col('Filiacao') || "N/A"),
          fotoUrl: String(col('Foto URL') || "N/A"),
          batalhao: String(col('Batalhao') || ""),
          enderecoPrincipal: String(col('Endereco Principal') || ""),
          outrosEnderecos: String(col('Outros Enderecos') || ""),
          status: String(col('Status') || "Procurado"),
          validade: formatarData(col('Validade')),
          infoProcessuais: String(col('Info Processuais') || ""),
          geodataSecundarios: String(col('Geodata Secundarios') || "[]"),
          dadosExtrasJSON: String(col('Dados Extras JSON') || "{}"),
          observacoes: (function() {
            try {
              const obsStr = String(col('Observacoes') || "").trim();
              if (obsStr.startsWith('[')) {
                try {
                  const arr = JSON.parse(obsStr);
                  if (arr.length > 0) return arr[0].texto || "";
                } catch(e) {}
              }
              return obsStr;
            } catch(e){}
            return "";
          })(),
          historicoObservacoes: (function() {
            try {
              const obsStr = String(col('Observacoes') || "").trim();
              if (obsStr.startsWith('[')) {
                 try { JSON.parse(obsStr); return obsStr; } catch(e) {}
              } else if (obsStr !== "") {
                 return JSON.stringify([{ data: "Legado", usuario: "Sistema", texto: obsStr }]);
              }
            } catch(e) {}
            return "[]";
          })(),
          lat: lat,
          lng: lng,
          cpi: String(col('CPI') || ""),
          bpmArea: String(col('BPM Area') || ""),
          ciaArea: String(col('CIA Area') || ""),
          dpArea: String(col('DP Area') || ""),
          cidade: String(col('Cidade') || ""),
          tipoImportacao: String(col('TipoImportacao') || ""),
          semEndereco: (col('Endereco Principal') || "").trim().toUpperCase() === "SEM ENDERECO" || (col('Endereco Principal') || "").trim().toUpperCase() === "SEM ENDEREÇO"
        };
        
        if (mandadoObj.mandado) {
          sincronizarMandadoFirebase(mandadoObj);
        }
      } catch (err) {
        console.error("Erro ao sincronizar onEdit para o Firebase: " + err.message);
      }
    }
  } else if (sheetName === "Poligonos") {
    sinalizarMudancaMandados();
  }
}

/**
 * Trigger onChange: captura inserção/exclusão de linhas e colunas.
 * O onEdit simples NÃO dispara para deleção de linhas.
 * IMPORTANTE: Este trigger precisa ser instalado manualmente via
 * Apps Script > Triggers > Add Trigger > onChange (ou via Setup.gs criarTriggers)
 */
function onChange(e) {
  if (!e) return;
  // Disparar para qualquer mudança estrutural na planilha
  if (e.changeType === 'REMOVE_ROW' || e.changeType === 'INSERT_ROW' || 
      e.changeType === 'EDIT' || e.changeType === 'OTHER') {
    sinalizarMudancaMandados();
    // Se a linha foi inserida ou excluída diretamente na planilha Google Sheets, re-sincronizar Firebase
    if (e.changeType === 'REMOVE_ROW' || e.changeType === 'INSERT_ROW') {
      try {
        migrarSheetsParaFirebase();
      } catch (errFb) {
        console.error('Erro ao sincronizar Firebase no onChange: ' + errFb.message);
      }
    }
  }
}

/**
 * Verifica se a memória do celular do policial está atualizada.
 * Se houver mandados mais recentes, baixa o banco todo. Senão, responde "atualizado".
 */
function verificarAtualizacaoMandados(timestampLocal) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado) {
    throw new Error("Acesso negado: Operação não autorizada.");
  }
  var props = PropertiesService.getScriptProperties();
  var ultimo = props.getProperty('MandadosLastUpdate') || "1"; // "1" garante fetch na 1a vez
  
  if (parseInt(ultimo) > parseInt(timestampLocal || 0)) {
    return { atualizado: true, timestamp: ultimo, dados: obterDados() };
  } else {
    return { atualizado: false, timestamp: ultimo };
  }
}

/**
 * Busca todos os mandados sem verificar cache.
 * Chamada pelo frontend quando não há cache local (primeiro acesso).
 */
function obterTodosOsMandados() {
  var props = PropertiesService.getScriptProperties();
  var timestamp = props.getProperty('MandadosLastUpdate') || new Date().getTime().toString();
  return { atualizado: true, timestamp: timestamp, dados: obterDados() };
}

/**
 * Altera o status de vários mandados em lote para "Baixado" e define a data de conferência.
 * Somente disponível para perfis com permissão (Administrador e Colaborador).
 */
function baixarMandadosEmMassa(listaMandados, listaValidados) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil === "Patrulheiro") {
    throw new Error("Acesso negado: Patrulheiros não podem inativar mandados.");
  }

  const mandadosParaBaixar = listaMandados || [];
  const mandadosParaValidar = listaValidados || [];

  if (mandadosParaBaixar.length === 0 && mandadosParaValidar.length === 0) {
    return { sucesso: false, mensagem: "Nenhum mandado informado para baixa ou conferência." };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
    if (!sheet) return { sucesso: false, mensagem: "Aba Mandados não encontrada." };

    const dados = sheet.getDataRange().getValues();
    const dataAtual = Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss");
    
    const setBaixas = new Set(mandadosParaBaixar.map(m => String(m).trim()));
    const setValidados = new Set(mandadosParaValidar.map(m => String(m).trim()));
    let contadorBaixas = 0;
    let contadorValidados = 0;

    for (let i = 1; i < dados.length; i++) {
      const mandadoNum = String(dados[i][2] || "").trim(); // Coluna C = Mandado (índice 2)
      const statusAtual = String(dados[i][16] || "").trim().toLowerCase(); // Coluna Q = Status (índice 16)
      
      // Pula se já estiver inativo (capturado ou baixado)
      if (statusAtual === "capturado" || statusAtual === "baixado") {
        continue;
      }

      const linha = i + 1;

      if (setBaixas.has(mandadoNum)) {
        sheet.getRange(linha, 17).setValue("Capturado"); // Coluna Q = Status (índice 17)
        sheet.getRange(linha, 2).setValue(dataAtual); // Coluna B = DataConferencia (índice 2)
        contadorBaixas++;
      } else if (setValidados.has(mandadoNum)) {
        sheet.getRange(linha, 2).setValue(dataAtual); // Coluna B = DataConferencia (índice 2)
        contadorValidados++;
      }
    }

    if (contadorBaixas > 0 || contadorValidados > 0) {
      sinalizarMudancaMandados();
      return { 
        sucesso: true, 
        mensagem: `Operação concluída. ${contadorBaixas} mandado(s) marcado(s) como Capturado e ${contadorValidados} mandado(s) validado(s).`,
        quantidade: contadorBaixas 
      };
    } else {
      return { sucesso: false, mensagem: "Nenhum mandado correspondente ativo foi encontrado na base local." };
    }
  } catch (erro) {
    return { sucesso: false, mensagem: "Erro ao processar alteração em massa: " + erro.message };
  } finally {
    lock.releaseLock();
  }
}

// ================================================================
// AUDITORIA DE INCONSISTENCIAS - v3.9.67
// ================================================================

/**
 * Corrige campos de um mandado inconsistente diretamente na planilha.
 * payload: {
 *   mandado,             // nº atual (chave de busca)
 *   novoNumeroMandado,   // se diferente, corrige coluna C
 *   novoNumeroProcesso,  // atualiza "Nº do processo:" dentro de infoProcessuais
 *   novoTitulo,          // atualiza dadosExtrasJSON.titulo
 *   artigo, infoProcessuais, enderecoPrincipal, cidade
 * }
 * Admin only.
 */
function corrigirInconsistenciaMandado(payload) {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    return { sucesso: false, mensagem: "Acesso negado: apenas administradores podem corrigir mandados." };
  }
  if (!payload || !payload.mandado) {
    return { sucesso: false, mensagem: "Mandado nao informado." };
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
    if (!sheet) return { sucesso: false, mensagem: "Aba Mandados nao encontrada." };
    
    var col = obterMapaColunas(sheet);
    var dados = sheet.getDataRange().getValues();
    var mandadoBusca = String(payload.mandado).trim();
    var linhaEncontrada = -1;
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][col['Mandado']] || "").trim() === mandadoBusca) { linhaEncontrada = i + 1; break; }
    }
    if (linhaEncontrada < 0) return { sucesso: false, mensagem: "Mandado nao encontrado: " + mandadoBusca };

    var detalhes = [];

    // Nº do Mandado
    if (payload.novoNumeroMandado && payload.novoNumeroMandado.trim() && payload.novoNumeroMandado.trim() !== mandadoBusca) {
      sheet.getRange(linhaEncontrada, col['Mandado'] + 1).setValue(payload.novoNumeroMandado.trim());
      detalhes.push("NovoMandado");
    }

    // Artigo
    if (payload.artigo && payload.artigo.trim()) {
      sheet.getRange(linhaEncontrada, col['Artigo'] + 1).setValue(payload.artigo.trim());
      detalhes.push("Artigo");
    }

    // Lei
    if (payload.lei && payload.lei.trim()) {
      if (col['Lei'] !== undefined) {
        sheet.getRange(linhaEncontrada, col['Lei'] + 1).setValue(payload.lei.trim());
      } else {
        // Se a coluna Lei não existir, salva no Dados Extras JSON
        var extrasColIdx = col['Dados Extras JSON'];
        try {
          var extrasStr = String(dados[linhaEncontrada - 1][extrasColIdx] || "{}");
          var extras = JSON.parse(extrasStr);
          extras.lei = payload.lei.trim();
          sheet.getRange(linhaEncontrada, extrasColIdx + 1).setValue(JSON.stringify(extras));
        } catch(e) {
          sheet.getRange(linhaEncontrada, extrasColIdx + 1).setValue(JSON.stringify({ lei: payload.lei.trim() }));
        }
      }
      detalhes.push("Lei");
    }

    // Endereço Principal
    if (payload.enderecoPrincipal && payload.enderecoPrincipal.trim()) {
      sheet.getRange(linhaEncontrada, col['Endereco Principal'] + 1).setValue(payload.enderecoPrincipal.trim());
      
      // Limpar erros de observação e resetar status para permitir geocodificação
      var obsCol = col['Observações'] !== undefined ? col['Observações'] : col['Observacoes'];
      if (obsCol !== undefined) {
        var obs = String(dados[linhaEncontrada - 1][obsCol] || "");
        obs = obs.replace("Falha: Cidade não identificada (Geocodificação retida)", "").trim();
        sheet.getRange(linhaEncontrada, obsCol + 1).setValue(obs);
      }
      if (col['Status'] !== undefined) {
        sheet.getRange(linhaEncontrada, col['Status'] + 1).setValue("Procurado");
      }

      // Limpar Lat/Lng para forçar geocodificação
      sheet.getRange(linhaEncontrada, col['Latitude'] + 1).clearContent();
      sheet.getRange(linhaEncontrada, col['Longitude'] + 1).clearContent();
      try {
        executarGeocodificacaoLinha(sheet, linhaEncontrada, true, col);
      } catch(e) {}
      
      detalhes.push("Endereco");
    }

    // Info Processual
    var infoBase = payload.infoProcessuais && payload.infoProcessuais.trim()
      ? payload.infoProcessuais.trim()
      : String(dados[linhaEncontrada - 1][col['Info Processuais']] || "");

    var infoAlterada = false;

    if (payload.novoNumeroProcesso && payload.novoNumeroProcesso.trim()) {
      // Substituir "Nº do processo: XXX" dentro do infoProcessuais
      var novoProc = payload.novoNumeroProcesso.trim();
      if (/N[º°]?\s*do processo:/i.test(infoBase)) {
        infoBase = infoBase.replace(/N[º°]?\s*do processo:\s*([^|]+)/i, "Nº do processo: " + novoProc);
      } else {
        infoBase = "Nº do processo: " + novoProc + (infoBase ? " | " + infoBase : "");
      }
      infoAlterada = true;
      detalhes.push("NumProcesso");
    }

    if (payload.lei && payload.lei.trim()) {
      infoBase = atualizarCampoInfoProcessual_(infoBase, "Lei", payload.lei.trim());
      infoAlterada = true;
    }

    if (payload.artigo && payload.artigo.trim()) {
      infoBase = atualizarCampoInfoProcessual_(infoBase, "Artigo", payload.artigo.trim());
      infoAlterada = true;
    }

    if (payload.infoProcessuais && payload.infoProcessuais.trim()) {
      infoAlterada = true;
      detalhes.push("InfoProcessual");
    }

    if (infoAlterada) {
      sheet.getRange(linhaEncontrada, col['Info Processuais'] + 1).setValue(infoBase);
    }

    // Título — atualiza dadosExtrasJSON.titulo
    if (payload.novoTitulo && payload.novoTitulo.trim()) {
      var extrasColIdx = col['Dados Extras JSON'];
      try {
        var extrasStr = String(dados[linhaEncontrada - 1][extrasColIdx] || "{}");
        var extras = JSON.parse(extrasStr);
        extras.titulo = payload.novoTitulo.trim();
        sheet.getRange(linhaEncontrada, extrasColIdx + 1).setValue(JSON.stringify(extras));
        detalhes.push("Titulo");
      } catch(e) {
        // Se JSON inválido, cria novo
        sheet.getRange(linhaEncontrada, extrasColIdx + 1).setValue(JSON.stringify({ titulo: payload.novoTitulo.trim() }));
        detalhes.push("Titulo");
      }
    }

    // Cidade
    if (payload.cidade && payload.cidade.trim()) {
      sheet.getRange(linhaEncontrada, col['Cidade'] + 1).setValue(payload.cidade.trim());
      detalhes.push("Cidade");
    }

    if (detalhes.length === 0) return { sucesso: false, mensagem: "Nenhum campo valido fornecido." };

    sinalizarMudancaMandados();
    registrarHistorico(checagem.email, checagem.nome, "AUDITORIA",
      "Corrigiu mandado " + mandadoBusca + ": [" + detalhes.join(", ") + "]");
    var novaLat = sheet.getRange(linhaEncontrada, col['Latitude'] + 1).getValue();
    var novaLng = sheet.getRange(linhaEncontrada, col['Longitude'] + 1).getValue();
    var novoBtl = sheet.getRange(linhaEncontrada, col['Batalhao'] + 1).getValue();
    var novoCpi = sheet.getRange(linhaEncontrada, col['CPI'] + 1).getValue();
    var novoBpm = sheet.getRange(linhaEncontrada, col['BPM Area'] + 1).getValue();
    var novaCia = sheet.getRange(linhaEncontrada, col['CIA Area'] + 1).getValue();
    var novaDp  = sheet.getRange(linhaEncontrada, col['DP Area'] + 1).getValue();
    var novaCid = sheet.getRange(linhaEncontrada, col['Cidade'] + 1).getValue();

    // FIREBASE SYNC
    try {
      var obsAtual = String(sheet.getRange(linhaEncontrada, col['Observacoes'] + 1).getValue() || "");
      var extrasStr = String(sheet.getRange(linhaEncontrada, col['Dados Extras JSON'] + 1).getValue() || "{}");
      
      var leiValue = "";
      if (col['Lei'] !== undefined) {
        leiValue = sheet.getRange(linhaEncontrada, col['Lei'] + 1).getValue();
      } else {
        try {
          var extObj = JSON.parse(extrasStr);
          leiValue = extObj.lei || "";
        } catch(e){}
      }

      atualizarMandadoFirebase(mandadoBusca, {
        lei: leiValue,
        artigo: sheet.getRange(linhaEncontrada, col['Artigo'] + 1).getValue(),
        enderecoPrincipal: sheet.getRange(linhaEncontrada, col['Endereco Principal'] + 1).getValue(),
        infoProcessuais: sheet.getRange(linhaEncontrada, col['Info Processuais'] + 1).getValue(),
        cidade: novaCid,
        lat: novaLat !== "" ? novaLat : null,
        lng: novaLng !== "" ? novaLng : null,
        batalhao: novoBtl,
        cpi: novoCpi,
        bpmArea: novoBpm,
        ciaArea: novaCia,
        dpArea: novaDp,
        dadosExtrasJSON: extrasStr
      });
    } catch(fbErr) {
      Logger.log("Erro no sync Firebase: " + fbErr);
    }

    return { 
      sucesso: true, 
      mensagem: "Mandado " + mandadoBusca + " corrigido. Campos: " + detalhes.join(", ") + ".",
      lat: novaLat !== "" ? novaLat : null,
      lng: novaLng !== "" ? novaLng : null,
      batalhao: novoBtl,
      cpi: novoCpi,
      bpmArea: novoBpm,
      ciaArea: novaCia,
      dpArea: novaDp,
      cidade: novaCid,
      infoProcessuais: infoBase
    };
  } catch (erro) {
    return { sucesso: false, mensagem: "Erro: " + erro.message };
  } finally {
    lock.releaseLock();
  }
}

// =================================================================
// GESTÃO DE INCONSISTÊNCIAS (NOVO MÓDULO HÍBRIDO)
// =================================================================

/**
 * Retorna a lista de registros com erro de endereço (Inconsistencias + Mandados com endereços não geocodificados).
 */
function buscarInconsistencias() {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    return { sucesso: false, mensagem: "Acesso negado." };
  }
  
  var lista = [];

  // Mapear todos os mandados válidos ativos na aba Mandados
  var setMandadosValidos = {};
  var sheetMandados = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
  if (sheetMandados && sheetMandados.getLastRow() > 1) {
    var dadosM = sheetMandados.getDataRange().getValues();
    if (dadosM.length > 1) {
      var hM = dadosM[0];
      var idxMandM = hM.indexOf("Mandado");
      if (idxMandM === -1) idxMandM = hM.indexOf("Nº Mandado");
      if (idxMandM >= 0) {
        for (var m = 1; m < dadosM.length; m++) {
          var valM = String(dadosM[m][idxMandM] || "").trim();
          if (valM) setMandadosValidos[valM] = true;
        }
      }
    }
  }

  // 1. Puxar da aba "Inconsistencias" (erros oriundos da importação em lote)
  var sheetIncons = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inconsistencias");
  if (sheetIncons && sheetIncons.getLastRow() > 1) {
    var dadosInc = sheetIncons.getDataRange().getValues();
    if (dadosInc.length > 1) {
      var hI = dadosInc[0];
      var idxMandI = hI.indexOf("Mandado");
      if (idxMandI === -1) idxMandI = hI.indexOf("Nº Mandado");
      var idxNomeI = hI.indexOf("Nome");
      var idxEndI = hI.indexOf("Endereco Principal");
      if (idxEndI === -1) idxEndI = hI.indexOf("Endereço Principal");
      var idxObsI = hI.indexOf("Observacoes");
      if (idxObsI === -1) idxObsI = hI.indexOf("Observações");

      if (idxMandI >= 0) {
        var linhasOrfas = [];
        for (var i = 1; i < dadosInc.length; i++) {
          var linha = dadosInc[i];
          var mandNum = String(linha[idxMandI] || "").trim();
          if (!mandNum) continue;

          // Se o mandado não existe mais na aba Mandados (foi excluído)
          if (!setMandadosValidos[mandNum]) {
            linhasOrfas.push(i + 1);
            continue;
          }
          
          var endText = idxEndI >= 0 ? String(linha[idxEndI] || "").trim() : "";
          if (!endText) continue;
          
          var statusErro = idxObsI >= 0 ? String(linha[idxObsI] || "Inconsistente (Principal)") : "Inconsistente (Principal)";
          if (statusErro.indexOf("IGNORAR_ERRO") !== -1 || statusErro.indexOf("Erro Ignorado") !== -1) continue;
          
          lista.push({
            mandado: mandNum,
            nome: idxNomeI >= 0 ? String(linha[idxNomeI] || "") : "",
            enderecoExtendido: endText,
            statusErro: statusErro,
            tipoEndereco: "Principal",
            idLinhaInconsistencia: i + 1
          });
        }

        // Eliminar registros órfãos da aba Inconsistencias
        if (linhasOrfas.length > 0) {
          for (var lo = linhasOrfas.length - 1; lo >= 0; lo--) {
            try { sheetIncons.deleteRow(linhasOrfas[lo]); } catch(eLo){}
          }
        }
      }
    }
  }

  // 2. Varrer a aba "Mandados" identificando erros nos endereços Principais e Secundários
  var sheetMandados = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
  if (sheetMandados) {
    var dadosM = sheetMandados.getDataRange().getValues();
    if (dadosM.length > 1) {
      var hM = dadosM[0];
      var idxMandM = hM.indexOf("Mandado");
      if (idxMandM === -1) idxMandM = hM.indexOf("Nº Mandado");
      var idxNomeM = hM.indexOf("Nome");
      var idxEndM = hM.indexOf("Endereco Principal");
      if (idxEndM === -1) idxEndM = hM.indexOf("Endereço Principal");
      var idxLatM = hM.indexOf("Latitude");
      if (idxLatM === -1) idxLatM = hM.indexOf("Lat");
      var idxLngM = hM.indexOf("Longitude");
      if (idxLngM === -1) idxLngM = hM.indexOf("Lng");
      var idxObsM = hM.indexOf("Observacoes");
      if (idxObsM === -1) idxObsM = hM.indexOf("Observações");
      var idxStatusM = hM.indexOf("Status");
      var idxSecM = hM.indexOf("Geodata Secundarios");
      if (idxSecM === -1) idxSecM = hM.indexOf("Geodata Secundários");

      if (idxMandM >= 0) {
        for (var j = 1; j < dadosM.length; j++) {
          var r = dadosM[j];
          var mandadoNum = String(r[idxMandM] || "").trim();
          if (!mandadoNum) continue;

          var nomeM = idxNomeM >= 0 ? String(r[idxNomeM] || "") : "";
          var endPrin = idxEndM >= 0 ? String(r[idxEndM] || "").trim() : "";
          var latP = idxLatM >= 0 ? r[idxLatM] : null;
          var lngP = idxLngM >= 0 ? r[idxLngM] : null;
          var obsM = idxObsM >= 0 ? String(r[idxObsM] || "") : "";
          var statusM = idxStatusM >= 0 ? String(r[idxStatusM] || "") : "";
          var geoSec = idxSecM >= 0 ? String(r[idxSecM] || "[]") : "[]";

          var jaNaListaP = lista.some(function(item) { return item.mandado === mandadoNum && item.tipoEndereco === "Principal"; });

          if (!jaNaListaP && endPrin && endPrin.length > 2) {
            var ehIgnoradoP = obsM.indexOf("IGNORAR_ERRO") !== -1 || obsM.indexOf("Erro Ignorado") !== -1;
            if (!ehIgnoradoP) {
              if (statusM === "Sem Endereço" || obsM.indexOf("ERRO_GEOCODIFICACAO") !== -1 || obsM.indexOf("Geocodificação retida") !== -1 || obsM.indexOf("Pino Sede") !== -1 || !latP || !lngP) {
                lista.push({
                  mandado: mandadoNum,
                  nome: nomeM,
                  enderecoExtendido: endPrin,
                  statusErro: obsM.indexOf("Geocodificação retida") !== -1 ? "Falha: Cidade não identificada" : (statusM === "Sem Endereço" ? "Sem Endereço" : "Falha na Geocodificação (Principal)"),
                  tipoEndereco: "Principal",
                  idLinhaMandado: j + 1
                });
              }
            }
          }

          try {
            var secArr = JSON.parse(geoSec);
            if (Array.isArray(secArr)) {
              secArr.forEach(function(sec, secIdx) {
                var endSecText = String(sec.endereco || sec.texto || "").trim();
                if (!endSecText) return;

                var sLat = sec.lat;
                var sLng = sec.lng;
                var sStatus = String(sec.statusGeocodificacao || sec.status || "").toUpperCase();
                var sIgnorar = sec.ignorarErro === true || sStatus === "IGNORADO" || sStatus === "OK";

                if (!sIgnorar) {
                  if (!sLat || !sLng || sLat === 0 || sLng === 0 || sStatus.indexOf("ERRO") !== -1 || sStatus.indexOf("FALHA") !== -1) {
                    lista.push({
                      mandado: mandadoNum,
                      nome: nomeM,
                      enderecoExtendido: endSecText,
                      statusErro: "Falha na Geocodificação (Secundário #" + (secIdx + 1) + ")",
                      tipoEndereco: "Secundário",
                      secIndex: secIdx,
                      idLinhaMandado: j + 1
                    });
                  }
                }
              });
            }
          } catch(eSec) {}
        }
      }
    }
  }

  return { sucesso: true, dados: lista };
}

/**
 * Salva a nova coordenada e endereço no banco de Mandados, e remove da aba de Inconsistencias.
 */
function salvarInconsistencia(payload) {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    return { sucesso: false, mensagem: "Acesso negado." };
  }
  
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var sheetMandados = planilha.getSheetByName("Mandados");
    var sheetIncons = planilha.getSheetByName("Inconsistencias");
    
    if (!sheetMandados || !sheetIncons) return { sucesso: false, mensagem: "Abas necessárias não encontradas." };
    
    // 1. Procurar no Mandados
    var dadosM = sheetMandados.getDataRange().getValues();
    var colM = obterMapaCabecalhos(dadosM[0]);
    var mandadoBusca = String(payload.mandado).trim();
    var colMandado = colM['Mandado'] !== undefined ? colM['Mandado'] : colM['Nº Mandado'];
    var colEndPrin = colM['Endereco Principal'] !== undefined ? colM['Endereco Principal'] : colM['Endereço Principal'];
    var colLat = colM['Latitude'] !== undefined ? colM['Latitude'] : colM['Lat'];
    var colLng = colM['Longitude'] !== undefined ? colM['Longitude'] : colM['Lng'];
    var colBat = colM['Batalhao'] !== undefined ? colM['Batalhao'] : colM['Batalhão'];
    var colObs = colM['Observações'] !== undefined ? colM['Observações'] : colM['Observacoes'];
    var colStatus = colM['Status'];

    if (colMandado === undefined) return { sucesso: false, mensagem: "Coluna Mandado não encontrada." };
    
    var linhaM = -1;
    for (var i = 1; i < dadosM.length; i++) {
      if (String(dadosM[i][colMandado] || "").trim() === mandadoBusca) {
        linhaM = i + 1;
        break;
      }
    }
    
    if (linhaM > -1) {
      if (payload.tipoEndereco === 'Secundário' || payload.secIndex !== undefined) {
        // Atualizar Endereço Secundário no Geodata Secundarios
        var colSec = colM['Geodata Secundarios'] !== undefined ? colM['Geodata Secundarios'] : colM['Geodata Secundários'];
        if (colSec !== undefined) {
          var rawSec = String(dadosM[linhaM - 1][colSec] || "[]");
          var arrSec = [];
          try { arrSec = JSON.parse(rawSec); } catch(e) {}
          var secIdx = typeof payload.secIndex === 'number' ? payload.secIndex : 0;
          
          if (!arrSec[secIdx]) arrSec[secIdx] = {};
          arrSec[secIdx].endereco = payload.enderecoFormatado;
          if (payload.lat && payload.lng) {
            arrSec[secIdx].lat = payload.lat;
            arrSec[secIdx].lng = payload.lng;
            arrSec[secIdx].statusGeocodificacao = "OK";
            arrSec[secIdx].ignorarErro = false;
            
            var areaSec = identificarAreaPorCoordenadas(payload.lat, payload.lng);
            if (areaSec) {
              if (areaSec.batalhao) arrSec[secIdx].batalhao = areaSec.batalhao;
              if (areaSec.cpi) arrSec[secIdx].cpi = areaSec.cpi;
              if (areaSec.cia) arrSec[secIdx].cia = areaSec.cia;
              if (areaSec.delegacia) arrSec[secIdx].delegacia = areaSec.delegacia;
              if (areaSec.cidade) arrSec[secIdx].cidade = areaSec.cidade;
            }
          }
          sheetMandados.getRange(linhaM, colSec + 1).setValue(JSON.stringify(arrSec));
        }
      } else {
        // Atualizar Mandados Principal
        if (colEndPrin !== undefined) sheetMandados.getRange(linhaM, colEndPrin + 1).setValue(payload.enderecoFormatado);
        
        if (payload.lat && payload.lng && colLat !== undefined && colLng !== undefined) {
          sheetMandados.getRange(linhaM, colLat + 1).setValue(payload.lat);
          sheetMandados.getRange(linhaM, colLng + 1).setValue(payload.lng);
          
          var areaDetectada = identificarAreaPorCoordenadas(payload.lat, payload.lng);
          if (areaDetectada) {
            if(areaDetectada.batalhao && colBat !== undefined) sheetMandados.getRange(linhaM, colBat + 1).setValue(areaDetectada.batalhao);
            if(areaDetectada.cpi && colM['CPI'] !== undefined) sheetMandados.getRange(linhaM, colM['CPI'] + 1).setValue(areaDetectada.cpi);
            if(areaDetectada.batalhao && colM['BPM Area'] !== undefined) sheetMandados.getRange(linhaM, colM['BPM Area'] + 1).setValue(areaDetectada.batalhao);
            if(areaDetectada.cia && colM['CIA Area'] !== undefined) sheetMandados.getRange(linhaM, colM['CIA Area'] + 1).setValue(areaDetectada.cia);
            if(areaDetectada.delegacia && colM['DP Area'] !== undefined) sheetMandados.getRange(linhaM, colM['DP Area'] + 1).setValue(areaDetectada.delegacia);
            if(areaDetectada.cidade && colM['Cidade'] !== undefined) sheetMandados.getRange(linhaM, colM['Cidade'] + 1).setValue(areaDetectada.cidade);
          }
        }
      }
      
      // Remover observações de bloqueio
      if (colObs !== undefined) {
        var obs = String(dadosM[linhaM - 1][colObs] || "");
        obs = obs.replace("Falha: Cidade não identificada (Geocodificação retida)", "").replace("IGNORAR_ERRO", "").replace("Erro Ignorado", "").trim();
        if (!obs) obs = "Corrigido via Painel";
        sheetMandados.getRange(linhaM, colObs + 1).setValue(obs);
      }
      
      if (colStatus !== undefined) sheetMandados.getRange(linhaM, colStatus + 1).setValue("Procurado");
    }
    
    // 2. Remover da aba Inconsistencias
    var dadosI = sheetIncons.getDataRange().getValues();
    var colI = obterMapaCabecalhos(dadosI[0]);
    var colIMandado = colI['Mandado'] !== undefined ? colI['Mandado'] : colI['Nº Mandado'];
    var linhaI = -1;
    if (colIMandado !== undefined) {
      for (var j = 1; j < dadosI.length; j++) {
        if (String(dadosI[j][colIMandado] || "").trim() === mandadoBusca) {
          linhaI = j + 1;
          break;
        }
      }
    }
    if (linhaI > -1) {
      sheetIncons.deleteRow(linhaI);
    }
    
    limparCacheBD();
    sinalizarMudancaMandados();
    registrarHistorico(checagem.email, checagem.nome, "AUDITORIA", "Corrigiu inconsistência do mandado " + mandadoBusca);
    
    // FIREBASE SYNC
    if (linhaM > -1) {
      try {
        var obsStr = "Corrigido via Painel";
        if (colObs !== undefined) {
           obsStr = String(dadosM[linhaM - 1][colObs] || "").replace("Falha: Cidade não identificada (Geocodificação retida)", "").trim();
           if (!obsStr) obsStr = "Corrigido via Painel";
        }
        var updateFb = {
          enderecoPrincipal: payload.enderecoFormatado, lat: payload.lat, lng: payload.lng,
          status: "Procurado", observacoes: obsStr
        };
        var areaD = identificarAreaPorCoordenadas(payload.lat, payload.lng);
        if (areaD) {
          updateFb.batalhao = areaD.batalhao; updateFb.cpi = areaD.cpi;
          updateFb.bpmArea = areaD.batalhao; updateFb.ciaArea = areaD.cia;
          updateFb.dpArea = areaD.delegacia; updateFb.cidade = areaD.cidade;
        }
        atualizarMandadoFirebase(mandadoBusca, updateFb);
      } catch(e) {}
    }
    
    return { sucesso: true, mensagem: "Inconsistência resolvida com sucesso!" };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Limpa o pino e define o mandado como Mapeamento Inviável.
 */
function descartarInconsistencia(mandado) {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") return { sucesso: false };
  
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var sheetM = planilha.getSheetByName("Mandados");
    var sheetI = planilha.getSheetByName("Inconsistencias");
    
    var dadosM = sheetM.getDataRange().getValues();
    var colM = obterMapaCabecalhos(dadosM[0]);
    var mandadoBusca = String(mandado).trim();
    
    var colMandado = colM['Mandado'] !== undefined ? colM['Mandado'] : colM['Nº Mandado'];
    var colLat = colM['Latitude'] !== undefined ? colM['Latitude'] : colM['Lat'];
    var colLng = colM['Longitude'] !== undefined ? colM['Longitude'] : colM['Lng'];
    var colStatus = colM['Status'];
    var colObs = colM['Observações'] !== undefined ? colM['Observações'] : colM['Observacoes'];

    if (colMandado !== undefined) {
      for (var i = 1; i < dadosM.length; i++) {
        if (String(dadosM[i][colMandado] || "").trim() === mandadoBusca) {
          var l = i + 1;
          if (colLat !== undefined) sheetM.getRange(l, colLat + 1).clearContent();
          if (colLng !== undefined) sheetM.getRange(l, colLng + 1).clearContent();
          if (colStatus !== undefined) sheetM.getRange(l, colStatus + 1).setValue("Mapeamento Inviável");
          if (colObs !== undefined) sheetM.getRange(l, colObs + 1).setValue("Mapeamento Inviável definido pelo Admin");
          
          // FIREBASE SYNC
          try {
            atualizarMandadoFirebase(mandadoBusca, {
              lat: null, lng: null,
              status: "Mapeamento Inviável",
              observacoes: "Mapeamento Inviável definido pelo Admin"
            });
          } catch(e) {}
          break;
        }
      }
    }
    
    if (sheetI) {
      var dadosI = sheetI.getDataRange().getValues();
      var colI = obterMapaCabecalhos(dadosI[0]);
      var colIMandado = colI['Mandado'] !== undefined ? colI['Mandado'] : colI['Nº Mandado'];
      if (colIMandado !== undefined) {
        for (var j = 1; j < dadosI.length; j++) {
          if (String(dadosI[j][colIMandado] || "").trim() === mandadoBusca) {
            sheetI.deleteRow(j + 1);
            break;
          }
        }
      }
    }
    
    sinalizarMudancaMandados();
    registrarHistorico(checagem.email, checagem.nome, "AUDITORIA", "Descartou pino do mandado " + mandadoBusca);
    return { sucesso: true };
  } catch(e) {
    return { sucesso: false, mensagem: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Apaga o endereço incorreto do banco de dados (e remove da lista de inconsistências), mantendo o mandado intacto.
 */
function excluirInconsistencia(payload) {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    return { sucesso: false, mensagem: "Acesso negado." };
  }
  
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var planilha = SpreadsheetApp.getActiveSpreadsheet();
    var sheetMandados = planilha.getSheetByName("Mandados");
    var sheetIncons = planilha.getSheetByName("Inconsistencias");

    var mandadoBusca = "";
    var tipoEndereco = "Principal";
    var secIndex = undefined;

    if (typeof payload === 'object' && payload !== null) {
      mandadoBusca = String(payload.mandado || "").trim();
      tipoEndereco = payload.tipoEndereco || "Principal";
      secIndex = payload.secIndex;
    } else {
      mandadoBusca = String(payload || "").trim();
    }

    if (!mandadoBusca) return { sucesso: false, mensagem: "Mandado não especificado." };

    // 1. Deletar linha correspondente da aba "Inconsistencias" (se existir lá)
    if (sheetIncons) {
      var dadosI = sheetIncons.getDataRange().getValues();
      if (dadosI.length > 1) {
        var hI = dadosI[0];
        var idxMandI = hI.indexOf("Mandado");
        if (idxMandI === -1) idxMandI = hI.indexOf("Nº Mandado");
        if (idxMandI >= 0) {
          for (var j = dadosI.length - 1; j >= 1; j--) {
            if (String(dadosI[j][idxMandI] || "").trim() === mandadoBusca) {
              sheetIncons.deleteRow(j + 1);
            }
          }
        }
      }
    }

    // 2. Apagar o endereço no banco de dados na aba "Mandados"
    if (sheetMandados) {
      var dadosM = sheetMandados.getDataRange().getValues();
      if (dadosM.length > 1) {
        var hM = dadosM[0];
        var idxMandM = hM.indexOf("Mandado");
        if (idxMandM === -1) idxMandM = hM.indexOf("Nº Mandado");

        var idxEndPrin = hM.indexOf("Endereco Principal");
        if (idxEndPrin === -1) idxEndPrin = hM.indexOf("Endereço Principal");

        var idxLat = hM.indexOf("Latitude");
        if (idxLat === -1) idxLat = hM.indexOf("Lat");

        var idxLng = hM.indexOf("Longitude");
        if (idxLng === -1) idxLng = hM.indexOf("Lng");

        var idxObs = hM.indexOf("Observacoes");
        if (idxObs === -1) idxObs = hM.indexOf("Observações");

        var idxStatus = hM.indexOf("Status");
        var idxSec = hM.indexOf("Geodata Secundarios");
        if (idxSec === -1) idxSec = hM.indexOf("Geodata Secundários");

        if (idxMandM >= 0) {
          for (var i = 1; i < dadosM.length; i++) {
            if (String(dadosM[i][idxMandM] || "").trim() === mandadoBusca) {
              var linhaM = i + 1;

              if (tipoEndereco === 'Secundário' || secIndex !== undefined) {
                // Apagar endereço secundário específico do array JSON em Geodata Secundarios
                if (idxSec >= 0) {
                  var rawSec = String(dadosM[i][idxSec] || "[]");
                  var arrSec = [];
                  try { arrSec = JSON.parse(rawSec); } catch(e) {}
                  var sIdx = typeof secIndex === 'number' ? secIndex : 0;
                  if (Array.isArray(arrSec) && arrSec.length > sIdx) {
                    arrSec.splice(sIdx, 1); // Remove o endereço secundário
                    sheetMandados.getRange(linhaM, idxSec + 1).setValue(JSON.stringify(arrSec));
                  }
                }
              } else {
                // Apagar Endereço Principal e Coordenadas do Mandado no banco de dados
                if (idxEndPrin >= 0) sheetMandados.getRange(linhaM, idxEndPrin + 1).setValue("");
                if (idxLat >= 0) sheetMandados.getRange(linhaM, idxLat + 1).clearContent();
                if (idxLng >= 0) sheetMandados.getRange(linhaM, idxLng + 1).clearContent();

                if (idxObs >= 0) {
                  var obs = String(dadosM[i][idxObs] || "");
                  obs = obs.replace(/Falha:[^|]*/gi, "")
                           .replace("ERRO_GEOCODIFICACAO", "")
                           .replace("Pino Sede", "")
                           .replace("Pino na Sede", "")
                           .replace("Geocodificação retida", "")
                           .trim();
                  if (obs.indexOf("IGNORAR_ERRO") === -1) {
                    obs = (obs ? obs + " | " : "") + "IGNORAR_ERRO";
                  }
                  sheetMandados.getRange(linhaM, idxObs + 1).setValue(obs);
                }
                
                if (idxStatus >= 0) sheetMandados.getRange(linhaM, idxStatus + 1).setValue("Sem Endereço");

                // Sincronizar remoção no Firebase
                try {
                  atualizarMandadoFirebase(mandadoBusca, {
                    enderecoPrincipal: "",
                    lat: null,
                    lng: null,
                    status: "Sem Endereço",
                    observacoes: obs
                  });
                } catch(eFb) {}
              }
              break;
            }
          }
        }
      }
    }

    limparCacheBD();
    sinalizarMudancaMandados();
    registrarHistorico(checagem.email, checagem.nome, "AUDITORIA", "Apagou endereço incorreto do mandado " + mandadoBusca);

    return { sucesso: true, mensagem: "Endereço apagado do banco de dados com sucesso." };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao excluir: " + e.message };
  } finally {
    lock.releaseLock();
  }
}



/**
 * Invalida/reprova a geocodificação de um mandado, enviando-o para inconsistências
 */
function reprovarGeocodificacaoMandado(mandado) {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    return { sucesso: false, mensagem: "Acesso negado: apenas administradores podem invalidar geocodificações." };
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
    if (!sheet) return { sucesso: false, mensagem: "Aba Mandados não encontrada." };
    
    var col = obterMapaColunas(sheet);
    var dados = sheet.getDataRange().getValues();
    var mandadoBusca = String(mandado).trim();
    var linhaEncontrada = -1;
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][col['Mandado']] || "").trim() === mandadoBusca) {
        linhaEncontrada = i + 1;
        break;
      }
    }
    if (linhaEncontrada < 0) return { sucesso: false, mensagem: "Mandado não encontrado: " + mandadoBusca };

    // Limpar coordenadas e CPI/Btl/Cia/DP
    sheet.getRange(linhaEncontrada, col['Latitude'] + 1).clearContent();
    sheet.getRange(linhaEncontrada, col['Longitude'] + 1).clearContent();
    sheet.getRange(linhaEncontrada, col['CPI'] + 1).clearContent();
    sheet.getRange(linhaEncontrada, col['CIA Area'] + 1).clearContent();
    sheet.getRange(linhaEncontrada, col['DP Area'] + 1).clearContent();
    sheet.getRange(linhaEncontrada, col['BPM Area'] + 1).clearContent();

    // Adicionar observação de erro
    var obsAtual = String(sheet.getRange(linhaEncontrada, col['Observacoes'] + 1).getValue() || "");
    var msgErro = "ERRO_GEOCODIFICACAO: Reprovado pelo Admin no relatório.";
    if (!obsAtual.includes(msgErro)) {
      sheet.getRange(linhaEncontrada, col['Observacoes'] + 1).setValue((obsAtual ? obsAtual + " | " : "") + msgErro);
    }

    limparCacheBD();
    sinalizarMudancaMandados();
    registrarHistorico(checagem.email, checagem.nome, "AUDITORIA", "Invalidou geocodificação do mandado " + mandadoBusca);
    
    // FIREBASE SYNC
    try {
      atualizarMandadoFirebase(mandadoBusca, {
        lat: null, lng: null,
        cpi: "", ciaArea: "", dpArea: "", bpmArea: "",
        observacoes: (obsAtual ? obsAtual + " | " : "") + msgErro
      });
    } catch(fbErr) {}

    return { sucesso: true, mensagem: "Geocodificação reprovada. O pino foi removido e o mandado enviado para inconsistências." };
  } catch (e) {
    return { sucesso: false, mensagem: e.message };
  } finally {
    lock.releaseLock();
  }
}



function obterPrepopulatedLeisLocal() {
  return [
  [
    "69888488-d618-437f-8f65-93c74af0c6be",
    "PENSÃO",
    "",
    "#DAE051",
    "SIM",
    "PENSÃO",
    "Código de Processo Civil - 2015",
    "13105",
    "528",
    "3",
    "-",
    "Prisão civil por inadimplemento voluntário e inescusável de obrigação alimentar (Pensão Alimentícia)"
  ],
  [
    "7fa78152-9f41-4bd9-b3d3-293e1e5b2c6a",
    "PENSÃO",
    "",
    "#DAE052",
    "SIM",
    "PENSÃO",
    "Código de Processo Civil - 2016",
    "13105",
    "528",
    "",
    "",
    "Prisão civil por débito alimentar - execução de prestações anteriores"
  ],
  [
    "4dfe95fa-887a-451e-8de8-92d318619daa",
    "PENSÃO",
    "comprovação do pagamento|ônus|onus| mas não comprovou",
    "#DAE051",
    "SIM",
    "PENSÃO",
    "Código de Processo Civil - 2015",
    "13105",
    "",
    "7",
    "-",
    "Prisão civil por débito alimentar - execução de prestações anteriores"
  ],
  [
    "bddbd375-b397-4d7c-be2a-486da0364b28",
    "INTERNAÇÃO",
    "",
    "#3DF4B5",
    "SIM",
    "ECA",
    "Estatuto da Criança e do Adolescente - 1990",
    "8069",
    "122",
    "1",
    "-",
    "Internação por ato infracional cometido mediante violência ou grave ameaça a pessoa"
  ],
  [
    "4070ed73-7df4-4b7f-aca6-5e15779952d1",
    "INTERNAÇÃO",
    "",
    "#3DF4B5",
    "SIM",
    "ECA",
    "Estatuto da Criança e do Adolescente - 1990",
    "8069",
    "122",
    "2",
    "-",
    "Internação por reiteração no cometimento de outras infrações graves"
  ],
  [
    "1448c3c4-3237-460f-a960-d07b4a974e6a",
    "INTERNAÇÃO",
    "",
    "#3DF4B5",
    "SIM",
    "ECA",
    "Estatuto da Criança e do Adolescente - 1990",
    "8069",
    "122",
    "3",
    "-",
    "Internação por descumprimento injustificado de medida anteriormente imposta"
  ],
  [
    "32ebb6d9-c735-41fc-80c2-a2a0fc32f123",
    "PEDOFILIA",
    "",
    "#47EAD7",
    "SIM",
    "PEDOFILIA",
    "Estatuto da Criança e do Adolescente - 1990",
    "8069",
    "240",
    "Caput",
    "-",
    "Produzir cena de sexo explícito ou pornográfica envolvendo criança ou adolescente"
  ],
  [
    "880bf746-6b79-4666-8e45-40d4b519f98b",
    "PEDOFILIA",
    "",
    "#47EAD7",
    "SIM",
    "PEDOFILIA",
    "Estatuto da Criança e do Adolescente - 1990",
    "8069",
    "241-A",
    "Caput",
    "-",
    "Oferecer ou disponibilizar material pornográfico envolvendo criança ou adolescente"
  ],
  [
    "9e1714be-18af-421a-85f5-777b6bdf934b",
    "HOM.TRÂNSITO",
    "",
    "#E08351",
    "SIM",
    "HOM.TRÂNSITO",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "302",
    "Caput",
    "-",
    "Homicídio culposo na direção de veículo automotor"
  ],
  [
    "708e3804-cc13-4c70-9dfa-57692d8ba409",
    "HOM.TRÂNSITO",
    "",
    "#E08351",
    "SIM",
    "HOM.TRÂNSITO",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "302",
    "1",
    "I",
    "Homicídio culposo majorado: Não possuir Permissão para Dirigir ou CNH"
  ],
  [
    "1e26d2b8-1817-4ae7-9b52-2b3bf5c74735",
    "HOM.TRÂNSITO",
    "",
    "#E08351",
    "SIM",
    "HOM.TRÂNSITO",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "302",
    "1",
    "II",
    "Homicídio culposo majorado: Praticá-lo em faixa de pedestres ou calçada"
  ],
  [
    "8f4d88eb-99e6-495d-86cc-a5c3bac41a5e",
    "HOM.TRÂNSITO",
    "",
    "#E08351",
    "SIM",
    "HOM.TRÂNSITO",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "302",
    "1",
    "III",
    "Homicídio culposo majorado: Deixar de prestar socorro à vítima do acidente"
  ],
  [
    "08cffb33-df84-48a5-9b09-69b95cdc1fbe",
    "HOM.TRÂNSITO",
    "",
    "#E08351",
    "SIM",
    "HOM.TRÂNSITO",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "302",
    "1",
    "IV",
    "Homicídio culposo majorado: No exercício de profissão transportando passageiros"
  ],
  [
    "284cbd84-5703-4350-bc83-558f798bafb8",
    "HOM.TRÂNSITO",
    "",
    "#E08351",
    "SIM",
    "HOM.TRÂNSITO",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "302",
    "3",
    "-",
    "Homicídio culposo qualificado: Conduzir sob influência de álcool ou substância psicoativa"
  ],
  [
    "4f965b7e-d2c1-43fb-a0ea-708dba54a08b",
    "LESÃO TRÂNSITO",
    "",
    "#51E08E",
    "SIM",
    "LESÃO TRÂNSITO",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "303",
    "Caput",
    "-",
    "Lesão corporal culposa na direção de veículo automotor"
  ],
  [
    "2801901a-4bc9-4d1b-85f0-e09e4b829a88",
    "LESÃO TRÂNSITO",
    "",
    "#51E08E",
    "SIM",
    "LESÃO TRÂNSITO",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "303",
    "2",
    "-",
    "Lesão corporal culposa qualificada: Sob influência de álcool gerando lesão grave/gravíssima"
  ],
  [
    "f149aceb-708f-45ad-859e-253f83db83ce",
    "CRIME TRÂNSITO",
    "",
    "#500AC1",
    "SIM",
    "CRIME TRÂNSITO",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "304",
    "Caput",
    "-",
    "Omissão de socorro no trânsito"
  ],
  [
    "a87c054a-d90a-48c9-9dd5-841a36fa913d",
    "CRIME TRÂNSITO",
    "",
    "#500AC1",
    "SIM",
    "CRIME TRÂNSITO",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "305",
    "Caput",
    "-",
    "Fuga do local do acidente para furtar-se à responsabilidade penal ou civil"
  ],
  [
    "70f9c51d-3468-48da-ba80-bc93cd77b076",
    "EMBRIAGUEZ",
    "",
    "#2DB714",
    "SIM",
    "EMBRIAGUEZ",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "306",
    "Caput",
    "-",
    "Embriaguez ao volante: Conduzir com capacidade psicomotora alterada por álcool/drogas"
  ],
  [
    "c0a0647b-05ad-4f0b-82fa-98330fc85668",
    "CRIME TRÂNSITO",
    "",
    "#500AC1",
    "SIM",
    "CRIME TRÂNSITO",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "307",
    "Caput",
    "-",
    "Violação da suspensão da habilitação"
  ],
  [
    "b9053055-78b4-42ee-9bc8-4f0b29f5a95c",
    "RACHA",
    "",
    "#9851E0",
    "SIM",
    "RACHA",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "308",
    "Caput",
    "-",
    "Participar de racha ou competição automobilística não autorizada gerando risco"
  ],
  [
    "81381473-a6d4-4407-b71d-89ff559aad9d",
    "RACHA",
    "",
    "#9851E0",
    "SIM",
    "RACHA",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "308",
    "1",
    "-",
    "Racha com resultado de lesão corporal grave"
  ],
  [
    "0bfa1adf-e610-43f2-9882-f8fbb3103497",
    "RACHA",
    "",
    "#9851E0",
    "SIM",
    "RACHA",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "308",
    "2",
    "-",
    "Racha com resultado morte"
  ],
  [
    "e9121d7e-17cf-4bd4-870b-b2f4c22d370d",
    "CRIME TRÂNSITO",
    "",
    "#500AC1",
    "SIM",
    "CRIME TRÂNSITO",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "309",
    "Caput",
    "-",
    "Dirigir sem CNH gerando perigo de dano"
  ],
  [
    "51603342-cc7c-4f3d-99e5-d8ee05b84f6b",
    "CRIME TRÂNSITO",
    "",
    "#500AC1",
    "SIM",
    "CRIME TRÂNSITO",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "310",
    "Caput",
    "-",
    "Entregar direção a pessoa não habilitada ou sem condições físicas/mentais (embriaguez)"
  ],
  [
    "aeafdd75-2ba5-44e7-a4b8-0e90ce150746",
    "CRIME TRÂNSITO",
    "",
    "#500AC1",
    "SIM",
    "CRIME TRÂNSITO",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "311",
    "Caput",
    "-",
    "Trafegar em velocidade incompatível nas proximidades de escolas ou aglomerações"
  ],
  [
    "364f4cdf-8ecb-429f-a96c-65a63322e3d8",
    "CRIME TRÂNSITO",
    "",
    "#500AC1",
    "SIM",
    "CRIME TRÂNSITO",
    "Código de Trânsito Brasileiro - 1997",
    "9503",
    "312",
    "Caput",
    "-",
    "Inovação artificiosa em caso de acidente com vítima (Fraude de trânsito)"
  ],
  [
    "07420ead-1cc2-4168-9f66-9a4184d9f1c9",
    "MAUS-TRATOS",
    "",
    "#47EA66",
    "SIM",
    "MAUS-TRATOS",
    "Lei de Crimes Ambientais - 1998",
    "9605",
    "32",
    "1-A",
    "-",
    "Maus-tratos quando se tratar de cão ou gato"
  ],
  [
    "63474458-22fa-4f27-947d-66ecf004281d",
    "CRIME AMBIENTAL",
    "",
    "#C3F43D",
    "SIM",
    "CRIME AMBIENTAL",
    "Lei de Crimes Ambientais - 1998",
    "9605",
    "54",
    "Caput",
    "-",
    "Causar poluição que resulte ou possa resultar em danos à saúde humana"
  ],
  [
    "006657e9-1332-4aa5-a5ba-03fa0b505e37",
    "TRÁFICO",
    "",
    "#B71485",
    "SIM",
    "TRÁFICO",
    "Lei de Drogas - 2006",
    "11343",
    "33",
    "Caput",
    "-",
    "Tráfico de Drogas: Importar, exportar, remeter, preparar, produzir, fabricar, adquirir, vender"
  ],
  [
    "7da92fe0-696e-4e82-a030-55e72ff7eb1e",
    "TRÁFICO",
    "",
    "#B71485",
    "SIM",
    "TRÁFICO",
    "Lei de Drogas - 2006",
    "11343",
    "33",
    "1",
    "I",
    "Tráfico de insumos: Fabricar ou comercializar matéria-prima para preparação de drogas"
  ],
  [
    "5aca6852-356a-4aee-afa4-26b65849a2c6",
    "TRÁFICO",
    "",
    "#B71485",
    "SIM",
    "TRÁFICO",
    "Lei de Drogas - 2006",
    "11343",
    "33",
    "1",
    "II",
    "Tráfico: Semear, cultivar ou colher plantas destinadas à preparação de drogas"
  ],
  [
    "e8bb4994-0698-4546-a49a-6f3cfc676d0f",
    "ASSOCIAÇÃO",
    "",
    "#1EAD8C",
    "SIM",
    "ASSOCIAÇÃO",
    "Lei de Drogas - 2006",
    "11343",
    "35",
    "Caput",
    "-",
    "Associação para o tráfico de drogas"
  ],
  [
    "456135af-3256-4ed1-88aa-f6f02b5634cd",
    "FINAN.TRÁFICO",
    "",
    "#516DE0",
    "SIM",
    "FINA.TRÁFICO",
    "Lei de Drogas - 2006",
    "11343",
    "36",
    "Caput",
    "-",
    "Financiamento do tráfico de drogas"
  ],
  [
    "fc624d82-8a25-401f-98a8-660fe9d71dc8",
    "ARMAS",
    "",
    "#3DB5F4",
    "SIM",
    "ARMAS",
    "Estatuto do Desarmamento - 2003",
    "10826",
    "14",
    "Caput",
    "-",
    "Porte ilegal de arma de fogo de uso permitido"
  ],
  [
    "490bd53c-0a3a-4ebf-983d-c0304975c075",
    "ARMAS",
    "",
    "#3DB5F4",
    "SIM",
    "ARMAS",
    "Estatuto do Desarmamento - 2003",
    "10826",
    "15",
    "Caput",
    "-",
    "Disparo de arma de fogo em via pública ou local habitado"
  ],
  [
    "1d046d5e-25af-4f67-814a-51d23faf3f57",
    "ARMAS",
    "",
    "#3DB5F4",
    "SIM",
    "ARMAS",
    "Estatuto do Desarmamento - 2003",
    "10826",
    "16",
    "Caput",
    "-",
    "Posse ou porte ilegal de arma de fogo de uso restrito/proibido"
  ],
  [
    "d66328c8-e57d-47dc-aac8-7b4164134738",
    "ARMAS",
    "",
    "#3DB5F4",
    "SIM",
    "ARMAS",
    "Estatuto do Desarmamento - 2003",
    "10826",
    "16",
    "1",
    "I",
    "Posse/porte ilegal de arma com numeração raspada ou adulterada"
  ],
  [
    "84cdf6ec-243b-462c-a959-082a2e214c42",
    "ARMAS",
    "",
    "#3DB5F4",
    "SIM",
    "ARMAS",
    "Estatuto do Desarmamento - 2003",
    "10826",
    "17",
    "Caput",
    "-",
    "Comércio ilegal de arma de fogo"
  ],
  [
    "23310db9-d0d6-40b8-8d28-c37073ad5a0f",
    "ARMAS",
    "",
    "#3DB5F4",
    "SIM",
    "ARMAS",
    "Estatuto do Desarmamento - 2003",
    "10826",
    "18",
    "Caput",
    "-",
    "Tráfico internacional de arma de fogo"
  ],
  [
    "6248ae5f-aa1c-475c-9946-5a806598a733",
    "M.PROTETIVA",
    "",
    "#478CEA",
    "SIM",
    "M.PROTETIVA",
    "\"Lei Maria da Penha - 2006",
    " 11.340\"",
    "11340",
    "24-A",
    "Caput",
    "-"
  ],
  [
    "9097f82c-73c0-4c1d-9ad1-9ee8143e64f5",
    "ORG.CRIMINOSA",
    "",
    "#E447EA",
    "SIM",
    "ORG.CRIMINOSA",
    "Lei de Organizações Criminosas - 2013",
    "12850",
    "2",
    "Caput",
    "-",
    "Promover, constituir, financiar ou integrar organização criminosa"
  ],
  [
    "6c3df597-9a82-4662-b9db-78f7bcfd9929",
    "LAVAGEM",
    "",
    "#B75F14",
    "SIM",
    "LAVAGEM",
    "Lei de Lavagem de Dinheiro - 1998",
    "9613",
    "1",
    "Caput",
    "-",
    "Ocultar ou dissimular bens, direitos ou valores provenientes de infração penal"
  ],
  [
    "4854d697-a09e-43dd-8a19-2e552485c9d8",
    "TORTURA",
    "",
    "#1491B7",
    "SIM",
    "TORTURA",
    "Lei de Tortura - 1997",
    "9455",
    "1",
    "Caput",
    "I",
    "Tortura para obter informação, declaração ou confissão"
  ],
  [
    "482b7478-28df-42d6-88a8-121cdf97ded4",
    "TORTURA",
    "",
    "#1491B7",
    "SIM",
    "TORTURA",
    "Lei de Tortura - 1997",
    "9455",
    "1",
    "Caput",
    "II",
    "Tortura para provocar ação ou omissão de natureza criminosa"
  ],
  [
    "68085536-cb64-4f34-aaaf-40a3a5bb790d",
    "TORTURA",
    "",
    "#1491B7",
    "SIM",
    "TORTURA",
    "Lei de Tortura - 1997",
    "9455",
    "1",
    "Caput",
    "III",
    "Tortura por discriminação racial ou religiosa"
  ],
  [
    "3a1f49ce-2057-44aa-96e1-055722851107",
    "HOMICÍDIO",
    "",
    "#9EB714",
    "SIM",
    "HOMICÍDIO",
    "Código Penal - 1940",
    "2848",
    "121",
    "Caput",
    "-",
    "Homicídio simples"
  ],
  [
    "534fda54-6f02-425c-b339-bbc0533c8bbc",
    "HOMICÍDIO",
    "",
    "#9EB715",
    "SIM",
    "HOMICÍDIO",
    "Código Penal - 1941",
    "2848",
    "121",
    "2",
    "-",
    "Homicídio qualificado: Se o homicídio é cometido:"
  ],
  [
    "e7714c2b-2bd1-48d9-998e-41278f2e5135",
    "HOMICÍDIO",
    "",
    "#9EB714",
    "SIM",
    "HOMICÍDIO",
    "Código Penal - 1940",
    "2848",
    "121",
    "2",
    "I",
    "Homicídio qualificado: Paga, promessa de recompensa ou motivo torpe"
  ],
  [
    "7fa906ed-2224-4d2a-8efa-721258c62c42",
    "HOMICÍDIO",
    "",
    "#9EB714",
    "SIM",
    "HOMICÍDIO",
    "Código Penal - 1940",
    "2848",
    "121",
    "2",
    "II",
    "Homicídio qualificado: Por motivo fútil"
  ],
  [
    "8239cb36-c8f0-44d1-b768-dcf8533acda1",
    "HOMICÍDIO",
    "",
    "#9EB714",
    "SIM",
    "HOMICÍDIO",
    "Código Penal - 1940",
    "2848",
    "121",
    "2",
    "III",
    "Homicídio qualificado: Veneno, fogo, explosivo, asfixia, tortura ou meio cruel"
  ],
  [
    "6f20d44f-c7a5-4dd0-ad0d-b015643b986d",
    "HOMICÍDIO",
    "",
    "#9EB714",
    "SIM",
    "HOMICÍDIO",
    "Código Penal - 1940",
    "2848",
    "121",
    "2",
    "IV",
    "Homicídio qualificado: À traição, de emboscada ou dissimulação"
  ],
  [
    "866ef753-a912-474f-b1de-1bb933cbf9c0",
    "HOMICÍDIO",
    "",
    "#9EB714",
    "SIM",
    "HOMICÍDIO",
    "Código Penal - 1940",
    "2848",
    "121",
    "2",
    "V",
    "Homicídio qualificado: Para assegurar execução/ocultação de outro crime"
  ],
  [
    "5fd421cc-985f-4324-8af0-a029fdabf48b",
    "FEMINICÍDIO",
    "",
    "#0AC1C1",
    "SIM",
    "FEMINICÍDIO",
    "Código Penal - 1940",
    "2848",
    "121",
    "2",
    "VI",
    "Feminicídio: Contra a mulher por razões da condição de sexo feminino"
  ],
  [
    "301187de-6f51-43ba-b0d8-a68d6fa2bc23",
    "HOMICÍDIO",
    "",
    "#9EB714",
    "SIM",
    "HOMICÍDIO",
    "Código Penal - 1940",
    "2848",
    "121",
    "2",
    "VII",
    "Homicídio qualificado: Contra autoridade ou agente de segurança"
  ],
  [
    "d51157cd-b646-4bc6-9b10-3169cdcc9239",
    "HOMICÍDIO",
    "",
    "#9EB714",
    "SIM",
    "HOMICÍDIO",
    "Código Penal - 1940",
    "2848",
    "121",
    "2",
    "VIII",
    "Homicídio qualificado: Com emprego de arma de fogo de uso restrito/proibido"
  ],
  [
    "69e189e2-74a6-4eb0-888d-1e3465affff7",
    "HOMICÍDIO",
    "",
    "#9EB714",
    "SIM",
    "HOMICÍDIO",
    "Código Penal - 1940",
    "2848",
    "121",
    "2",
    "IX",
    "Homicídio qualificado: Contra menor de 14 anos"
  ],
  [
    "9b12c295-415e-44ec-8e00-239bd8bc726c",
    "ABORTO",
    "",
    "#341EAD",
    "SIM",
    "ABORTO",
    "Código Penal - 1940",
    "2848",
    "125",
    "Caput",
    "-",
    "Aborto provocado por terceiro sem consentimento da gestante"
  ],
  [
    "8d08e4ad-f3b6-4ddf-94ad-0a5317e03401",
    "LESÃO CORPORAL",
    "",
    "#AD1E1E",
    "SIM",
    "LESÃO CORPORAL",
    "Código Penal - 1940",
    "2848",
    "129",
    "1",
    "I",
    "Lesão corporal grave: Incapacidade para ocupações habituais por mais de 30 dias"
  ],
  [
    "85feff89-be5c-4b55-9725-b2af2084f895",
    "LESÃO CORPORAL",
    "",
    "#AD1E1E",
    "SIM",
    "LESÃO CORPORAL",
    "Código Penal - 1940",
    "2848",
    "129",
    "1",
    "II",
    "Lesão corporal grave: Perigo de vida"
  ],
  [
    "ebb8695b-f350-43a5-a538-bb559d2edd1d",
    "LESÃO CORPORAL",
    "",
    "#AD1E1E",
    "SIM",
    "LESÃO CORPORAL",
    "Código Penal - 1940",
    "2848",
    "129",
    "1",
    "III",
    "Lesão corporal grave: Debilidade permanente de membro, sentido ou função"
  ],
  [
    "ab8b4e5a-bbab-4380-ad33-410943fa54d4",
    "LESÃO CORPORAL",
    "",
    "#AD1E1E",
    "SIM",
    "LESÃO CORPORAL",
    "Código Penal - 1940",
    "2848",
    "129",
    "2",
    "I",
    "Lesão corporal gravíssima: Incapacidade permanente para o trabalho"
  ],
  [
    "59bc5459-6b32-4f6d-8e0f-fbffa4fb6080",
    "LESÃO CORPORAL",
    "",
    "#AD1E1E",
    "SIM",
    "LESÃO CORPORAL",
    "Código Penal - 1940",
    "2848",
    "129",
    "2",
    "II",
    "Lesão corporal gravíssima: Enfermidade incurável"
  ],
  [
    "8b3664c9-b116-4607-a904-06b694b18e35",
    "LESÃO CORPORAL",
    "",
    "#AD1E1E",
    "SIM",
    "LESÃO CORPORAL",
    "Código Penal - 1940",
    "2848",
    "129",
    "2",
    "III",
    "Lesão corporal gravíssima: Perda ou inutilização de membro"
  ],
  [
    "0c632fcd-8fa8-4fbd-806a-a6175388e732",
    "LESÃO CORPORAL",
    "",
    "#AD1E1E",
    "SIM",
    "LESÃO CORPORAL",
    "Código Penal - 1940",
    "2848",
    "129",
    "2",
    "IV",
    "Lesão corporal gravíssima: Deformidade permanente"
  ],
  [
    "5217d61b-7725-4b0a-b6f4-95919bbe50e8",
    "LESÃO CORPORAL",
    "",
    "#AD1E1E",
    "SIM",
    "LESÃO CORPORAL",
    "Código Penal - 1940",
    "2848",
    "129",
    "3",
    "-",
    "Lesão corporal seguida de morte"
  ],
  [
    "0f56bfd0-b17b-4f86-a41e-917fa6146893",
    "V.DOMÉSTICA",
    "",
    "#C10AB3",
    "SIM",
    "V.DOMÉSTICA",
    "Código Penal - 1940",
    "2848",
    "129",
    "9",
    "-",
    "Lesão corporal - Violência doméstica contra cônjuge, companheiro ou ascendente/descendente"
  ],
  [
    "5fe2a0f5-04de-46f5-af25-9d72e67ef358",
    "CÁRCERE",
    "",
    "#1EAD29",
    "SIM",
    "CÁRCERE",
    "Código Penal - 1940",
    "2848",
    "148",
    "Caput",
    "-",
    "Sequestro e cárcere privado"
  ],
  [
    "394f747d-6eda-43e9-82e3-f5bf6e83c61b",
    "TRAB.ESCRAVO",
    "",
    "#971EAD",
    "SIM",
    "TRAB.ESCRAVO",
    "Código Penal - 1940",
    "2848",
    "149",
    "Caput",
    "-",
    "Redução a condição análoga à de escravo"
  ],
  [
    "24f370ff-b23d-412b-9fc3-75f8479c4bfb",
    "TRÁF. PESSOA",
    "",
    "#C10A34",
    "SIM",
    "TRÁF. PESSOA",
    "Código Penal - 1940",
    "2848",
    "149-A",
    "Caput",
    "-",
    "Tráfico de pessoas"
  ],
  [
    "a8a4d5c6-b34b-490b-b30e-47f452f8075c",
    "FURTO",
    "",
    "#443DF4",
    "SIM",
    "FURTO",
    "Código Penal - 1940",
    "2848",
    "155",
    "Caput",
    "-",
    "Furto simples"
  ],
  [
    "4724e546-bbe6-4556-a24a-6065d3f90fd6",
    "FURTO",
    "",
    "#443DF4",
    "SIM",
    "FURTO",
    "Código Penal - 1940",
    "2848",
    "155",
    "4",
    "",
    "Furto qualificado"
  ],
  [
    "fdbe0b16-ca16-4f91-9858-cbf51afb7b4f",
    "FURTO",
    "",
    "#443DF4",
    "SIM",
    "FURTO",
    "Código Penal - 1940",
    "2848",
    "155",
    "4",
    "I",
    "Furto qualificado: Destruição ou rompimento de obstáculo"
  ],
  [
    "0e762883-07d9-46b6-a1ff-eb853c7b8abe",
    "FURTO",
    "",
    "#443DF4",
    "SIM",
    "FURTO",
    "Código Penal - 1940",
    "2848",
    "155",
    "4",
    "II",
    "Furto qualificado: Abuso de confiança, fraude, escalada ou destreza"
  ],
  [
    "7299ab1d-2ed8-4319-a201-1ac3d9da6eb6",
    "FURTO",
    "",
    "#443DF4",
    "SIM",
    "FURTO",
    "Código Penal - 1940",
    "2848",
    "155",
    "4",
    "III",
    "Furto qualificado: Emprego de chave falsa"
  ],
  [
    "cb67c715-412b-4896-b7e5-a913ebc2795a",
    "FURTO",
    "",
    "#443DF4",
    "SIM",
    "FURTO",
    "Código Penal - 1940",
    "2848",
    "155",
    "4",
    "IV",
    "Furto qualificado: Concurso de duas ou mais pessoas"
  ],
  [
    "562d4e41-56fd-4a80-92a3-aae818f5d646",
    "FURTO",
    "",
    "#443DF4",
    "SIM",
    "FURTO",
    "Código Penal - 1940",
    "2848",
    "155",
    "4-A",
    "-",
    "Furto qualificado: Com emprego de explosivo ou artefato análogo"
  ],
  [
    "afff66bf-5f7d-4019-b181-00d757808b17",
    "ROUBO",
    "",
    "#AD1E60",
    "SIM",
    "ROUBO",
    "Código Penal - 1940",
    "2848",
    "157",
    "Caput",
    "-",
    "Roubo simples: Subtração mediante violência ou grave ameaça"
  ],
  [
    "2f6c30a1-0c54-4f3a-91ff-af87f245dcb5",
    "ROUBO",
    "§ 1º",
    "#AD1E60",
    "SIM",
    "ROUBO",
    "Código Penal - 1941",
    "2848",
    "157",
    "1",
    "-",
    "Roubo simples: Subtração mediante violência ou grave ameaçaNa mesma pena incorre quem, logo depois de subtraída a coisa, emprega violência contra pessoa ou grave ameaça, a fim de assegurar a impunidade do crime ou a detenção da coisa para si ou para terceiro."
  ],
  [
    "57da5a09-4523-41de-9c83-c442129f0247",
    "ROUBO",
    "",
    "#AD1E60",
    "SIM",
    "ROUBO",
    "Código Penal - 1940",
    "2848",
    "157",
    "2",
    "II",
    "Roubo majorado: Concurso de duas ou mais pessoas"
  ],
  [
    "f03af8e8-461d-4cdc-b060-4d0d65cfd402",
    "ROUBO",
    "",
    "#AD1E60",
    "SIM",
    "ROUBO",
    "Código Penal - 1940",
    "2848",
    "157",
    "2",
    "III",
    "Roubo majorado: Vítima em serviço de transporte de valores"
  ],
  [
    "95fdb309-eb7d-4bf4-916c-66bce275d890",
    "ROUBO",
    "",
    "#AD1E60",
    "SIM",
    "ROUBO",
    "Código Penal - 1940",
    "2848",
    "157",
    "2",
    "IV",
    "Roubo majorado: Restrição da liberdade da vítima"
  ],
  [
    "07e374b4-149d-4a7e-8ca1-8172a32d5688",
    "ROUBO",
    "",
    "#AD1E60",
    "SIM",
    "ROUBO",
    "Código Penal - 1940",
    "2848",
    "157",
    "2",
    "V",
    "Roubo majorado: Subtração de veículo para outro Estado ou exterior"
  ],
  [
    "2c91acb3-212e-4c9e-816e-53cb359661b0",
    "ROUBO",
    "",
    "#AD1E60",
    "SIM",
    "ROUBO",
    "Código Penal - 1940",
    "2848",
    "157",
    "2",
    "VII",
    "Roubo majorado: Emprego de arma branca"
  ],
  [
    "f480625f-a555-4a78-b4b4-f4e357b6b48e",
    "ROUBO",
    "",
    "#AD1E60",
    "SIM",
    "ROUBO",
    "Código Penal - 1940",
    "2848",
    "157",
    "2-A",
    "I",
    "Roubo majorado: Emprego de arma de fogo"
  ],
  [
    "9135e53b-9ae4-4800-ac57-a19fea0ec7e4",
    "ROUBO",
    "",
    "#AD1E60",
    "SIM",
    "ROUBO",
    "Código Penal - 1940",
    "2848",
    "157",
    "2-A",
    "II",
    "Roubo majorado: Destruição ou rompimento de obstáculo com explosivo"
  ],
  [
    "83557922-d57f-471d-871c-6f345f111d8b",
    "ROUBO",
    "",
    "#AD1E60",
    "SIM",
    "ROUBO",
    "Código Penal - 1940",
    "2848",
    "157",
    "3",
    "I",
    "Roubo qualificado pela lesão corporal grave"
  ],
  [
    "b07adf44-8e95-4357-8923-289d596f8058",
    "LATROCÍNIO",
    "",
    "#1E6BAD",
    "SIM",
    "LATROCÍNIO",
    "Código Penal - 1940",
    "2848",
    "157",
    "3",
    "II",
    "Latrocínio: Roubo seguido de morte"
  ],
  [
    "628466e2-e7bb-462b-8363-1284c9da8178",
    "EXTORSÃO",
    "",
    "#0AC142",
    "SIM",
    "EXTORSÃO",
    "Código Penal - 1940",
    "2848",
    "158",
    "Caput",
    "-",
    "Extorsão simples"
  ],
  [
    "85c9137f-14d8-4cdb-b095-c25941fa6b9a",
    "SEQUESTRO",
    "",
    "#14B76C",
    "SIM",
    "SEQUESTRO",
    "Código Penal - 1940",
    "2848",
    "159",
    "Caput",
    "-",
    "Extorsão mediante sequestro"
  ],
  [
    "9db77f01-9f5a-4d24-a4db-a8046d1ca60d",
    "APROP. PREVID.",
    "",
    "#76AD1E",
    "SIM",
    "APROP. PREVID.",
    "Código Penal - 1940",
    "2848",
    "168-A",
    "1",
    "I",
    "Apropriação indébita previdenciária: Não recolher contribuições descontadas"
  ],
  [
    "423208e2-8a46-487a-ad26-900dd66581ab",
    "ESTELIONATO",
    "",
    "#E05162",
    "SIM",
    "ESTELIONATO",
    "Código Penal - 1940",
    "2848",
    "171",
    "Caput",
    "-",
    "Estelionato simples"
  ],
  [
    "7fe12e73-662b-4964-947f-846bbac39dea",
    "ESTELIONATO",
    "",
    "#E05162",
    "SIM",
    "ESTELIONATO",
    "Código Penal - 1940",
    "2848",
    "171",
    "2",
    "I",
    "Estelionato: Disposição de coisa alheia como própria"
  ],
  [
    "948a851e-cd36-4d92-989f-7a079ce362a2",
    "ESTELIONATO",
    "",
    "#E05162",
    "SIM",
    "ESTELIONATO",
    "Código Penal - 1940",
    "2848",
    "171",
    "2",
    "II",
    "Estelionato: Alienação ou oneração fraudulenta de coisa própria"
  ],
  [
    "b986ee3a-e3c9-4cf8-8c42-baf85b683a96",
    "ESTELIONATO",
    "",
    "#E05162",
    "SIM",
    "ESTELIONATO",
    "Código Penal - 1940",
    "2848",
    "171",
    "2",
    "VI",
    "Estelionato: Fraude no pagamento por cheque"
  ],
  [
    "7bbb60c9-e5df-4bc1-85f3-48bab8e8a5c7",
    "RECEPTAÇÃO",
    "",
    "#0A42C1",
    "SIM",
    "RECEPTAÇÃO",
    "Código Penal - 1940",
    "2848",
    "180",
    "caput",
    "-",
    "Adquirir, receber, transportar, conduzir ou ocultar, em proveito próprio ou alheio, coisa que sabe ser produto de crime, ou influir para que terceiro, de boa-fé, a adquira, receba ou oculte"
  ],
  [
    "ce68d21a-6614-4ec8-ac28-66f6e9ec6b86",
    "RECEPTAÇÃO",
    "",
    "#0A42C1",
    "SIM",
    "RECEPTAÇÃO",
    "Código Penal - 1940",
    "2848",
    "180",
    "1",
    "-",
    "Receptação qualificada: Adquirir, receber, transportar, conduzir, ocultar, ter em depósito, desmontar, montar, remontar, vender, expor à venda, ou de qualquer forma utilizar, em proveito próprio ou alheio, no exercício de atividade comercial ou industrial, coisa que deve saber ser produto de crime"
  ],
  [
    "96ffc297-a7f2-4e0b-8769-1fb67a0319d9",
    "RECEPTAÇÃO",
    "",
    "#0A42C2",
    "SIM",
    "RECEPTAÇÃO",
    "Código Penal - 1941",
    "2848",
    "180",
    "2",
    "-",
    "Receptação qualificada: Equipara-se à atividade comercial, para efeito do parágrafo anterior, qualquer forma de comércio irregular ou clandestino, inclusive o exercício em residência"
  ],
  [
    "524e8d06-1751-4d70-a46e-a792a8b8e3a6",
    "RECEPTAÇÃO",
    "",
    "#0A42C3",
    "SIM",
    "RECEPTAÇÃO",
    "Código Penal - 1942",
    "2848",
    "180",
    "3",
    "-",
    "Receptação qualificada: Adquirir ou receber coisa que, por sua natureza ou pela desproporção entre o valor e o preço, ou pela condição de quem a oferece, deve presumir-se obtida por meio criminoso"
  ],
  [
    "19d05c40-e361-459b-9a13-58d1aa70d3ee",
    "RECEPTAÇÃO",
    "",
    "#0A42C4",
    "SIM",
    "RECEPTAÇÃO",
    "Código Penal - 1943",
    "2848",
    "180",
    "4",
    "-",
    "Receptação qualificada: A receptação é punível, ainda que desconhecido ou isento de pena o autor do crime de que proveio a coisa."
  ],
  [
    "f89609a3-3489-455f-90cf-480714d11176",
    "RECEPTAÇÃO",
    "",
    "#0A42C5",
    "SIM",
    "RECEPTAÇÃO",
    "Código Penal - 1944",
    "2848",
    "180",
    "5",
    "-",
    "Receptação qualificada: Na hipótese do § 3º, se o criminoso é primário, pode o juiz, tendo em consideração as circunstâncias, deixar de aplicar a pena. Na receptação dolosa aplica-se o disposto no § 2º do art. 155"
  ],
  [
    "bd265ef8-9184-4d1a-9932-62caf5d1f70b",
    "RECEPTAÇÃO",
    "",
    "#0A42C6",
    "SIM",
    "RECEPTAÇÃO",
    "Código Penal - 1945",
    "2848",
    "180",
    "6",
    "-",
    "Receptação qualificada: Tratando-se de bens do patrimônio da União, de Estado, do Distrito Federal, de Município ou de autarquia, fundação pública, empresa pública, sociedade de economia mista ou empresa concessionária de serviços públicos, aplica-se em dobro a pena prevista no caput deste artigo"
  ],
  [
    "78e3a805-7543-442e-bdcd-177750678c14",
    "ESTUPRO",
    "",
    "#E051C4",
    "SIM",
    "ESTUPRO",
    "Código Penal - 1940",
    "2848",
    "213",
    "Caput",
    "-",
    "Estupro: Constranger mediante violência a conjunção carnal ou ato libidinoso"
  ],
  [
    "753fa5c1-deda-4a65-a4be-814f8d5437ad",
    "EST.VUNERA",
    "217A|217|217-A",
    "#99EA47",
    "SIM",
    "E.VUNERAVEL",
    "Código Penal - 1940",
    "2848",
    "217-A",
    "Caput",
    "-",
    "Estupro de vulnerável (menor de 14 anos ou pessoa sem discernimento)"
  ],
  [
    "ae6ad54d-802e-452c-8f62-113c5b5aebd9",
    "INCÊNDIO",
    "",
    "#C33DF4",
    "SIM",
    "INCÊNDIO",
    "Código Penal - 1940",
    "2848",
    "250",
    "Caput",
    "-",
    "Incêndio doloso expondo a perigo a vida ou patrimônio"
  ],
  [
    "1527c550-ea0f-43e5-a361-04d49fdeb0c5",
    "EPIDEMIA",
    "",
    "#EA477F",
    "SIM",
    "EPIDEMIA",
    "Código Penal - 1940",
    "2848",
    "267",
    "Caput",
    "-",
    "Epidemia: Causar epidemia mediante a propagação de germes patogênicos"
  ],
  [
    "755313fb-72df-4b54-a93e-38c7ab2d303f",
    "MED. FALSO",
    "",
    "#C1340A",
    "SIM",
    "MED. FALSO",
    "Código Penal - 1940",
    "2848",
    "273",
    "Caput",
    "-",
    "Falsificação, corrupção ou adulteração de produto terapêutico/medicinal"
  ],
  [
    "03ae9aa1-fc2d-49fa-9f6f-bc0eb3692af2",
    "ASSOC. CRIMINOSA",
    "",
    "#51CFE0",
    "SIM",
    "ASSOC. CRIMINOSA",
    "Código Penal - 1940",
    "2848",
    "288",
    "Caput",
    "-",
    "Associação criminosa"
  ],
  [
    "5ba65999-72f6-4b2f-a540-a3710abec6a0",
    "MILÍCIA",
    "",
    "#7347EA",
    "SIM",
    "MILÍCIA",
    "Código Penal - 1940",
    "2848",
    "288-A",
    "Caput",
    "-",
    "Constituição de milícia privada"
  ],
  [
    "321956c3-8aca-428c-9c31-e321c94b7619",
    "MOEDA FALSA",
    "",
    "#F4A73D",
    "SIM",
    "MOEDA FALSA",
    "Código Penal - 1940",
    "2848",
    "289",
    "Caput",
    "-",
    "Moeda falsa: Falsificar, fabricando ou alterando moeda"
  ],
  [
    "9ffa73b5-759c-4160-8ca3-9f9f776bbf91",
    "DOC. FALSO",
    "",
    "#1420B7",
    "SIM",
    "DOC. FALSO",
    "Código Penal - 1940",
    "2848",
    "297",
    "Caput",
    "-",
    "Falsificação de documento público"
  ]
];
}

/**
 * Utilitário para atualizar ou inserir um rótulo do tipo "Rotulo: Valor" dentro do texto de Info Processuais
 */
function atualizarCampoInfoProcessual_(texto, rotulo, novoValor) {
  if (!texto) texto = "";
  var regex = new RegExp("(" + rotulo + "\\s*:\\s*)([^\\n|]+)", "i");
  if (regex.test(texto)) {
    return texto.replace(regex, "$1" + novoValor);
  } else {
    var separador = texto.includes("\n") ? "\n" : " | ";
    return texto.trim() + (texto ? separador : "") + rotulo + ": " + novoValor;
  }
}

/**
 * Reprocessa todas as linhas da planilha Mandados preenchendo os dados de área (CPI, BTL, CIA, DP, Cidade)
 * na coluna Geodata Secundarios para todos os endereços secundários que possuem coordenadas.
 */
function reprocessarAreasGeodataSecundarios() {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    return { sucesso: false, mensagem: "Acesso negado." };
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Mandados");
  if (!sheet) return { sucesso: false, mensagem: "Aba Mandados não encontrada." };

  var dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return { sucesso: true, mensagem: "Nenhum mandado para reprocessar." };

  var colMap = obterMapaCabecalhos(dados[0]);
  var idxSec = colMap['Geodata Secundarios'] !== undefined ? colMap['Geodata Secundarios'] : colMap['Geodata Secundários'];
  if (idxSec === undefined) return { sucesso: false, mensagem: "Coluna Geodata Secundarios não encontrada." };

  var alterados = 0;

  for (var i = 1; i < dados.length; i++) {
    var rawSec = String(dados[i][idxSec] || "[]").trim();
    if (!rawSec || rawSec === "[]") continue;

    try {
      var arrSec = JSON.parse(rawSec);
      if (Array.isArray(arrSec) && arrSec.length > 0) {
        var modificado = false;
        arrSec.forEach(function(sec) {
          var lat = Number(sec.lat);
          var lng = Number(sec.lng);
          if (!isNaN(lat) && lat !== 0 && !isNaN(lng) && lng !== 0) {
            var area = identificarAreaPorCoordenadas(lat, lng);
            if (area) {
              if (area.cpi) { sec.cpi = area.cpi; modificado = true; }
              if (area.batalhao) { sec.batalhao = area.batalhao; modificado = true; }
              if (area.cia) { sec.cia = area.cia; modificado = true; }
              if (area.delegacia) { sec.delegacia = area.delegacia; modificado = true; }
              if (area.cidade) { sec.cidade = area.cidade; modificado = true; }
            }
          }
        });

        if (modificado) {
          sheet.getRange(i + 1, idxSec + 1).setValue(JSON.stringify(arrSec));
          alterados++;
        }
      }
    } catch(e) {
      Logger.log("Erro ao reprocessar linha " + (i + 1) + ": " + e.message);
    }
  }

  SpreadsheetApp.flush();
  try {
    if (typeof migrarSheetsParaFirebase === 'function') {
      migrarSheetsParaFirebase();
    }
  } catch(eFb) {
    Logger.log("Erro ao re-sincronizar Firebase: " + eFb.message);
  }
  sinalizarMudancaMandados();
  return { sucesso: true, mensagem: "Reprocessamento concluído! " + alterados + " mandados atualizados na Planilha e sincronizados automaticamente com o Firebase." };
}

