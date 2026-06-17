// ================================================================
// SENTINELA v2.0 — MÓDULO DE BANCO DE DADOS (CRUD + GAMIFICAÇÃO)
// ================================================================

/**
 * Lê todos os dados da aba Mandados e retorna como array de objetos.
 * Nova ordem de colunas: A=DataLancamento, B=DataConferencia, C=Mandado, D=Artigo, E=Nome,
 * F=CPF, G=RG, H=Nascimento, I=Naturalidade, J=Sexo, K=Cor, L=Filiacao,
 * M=FotoURL, N=Batalhao, O=Endereco, P=OutrosEnderecos, Q=Status, R=Validade,
 * S=InfoProcessuais, T=GeodataSecundarios, U=DadosExtrasJSON, V=Observacoes,
 * W=Lat, X=Lng, Y=CPI, Z=BPM_Area, AA=CIA_Area, AB=DP_Area, AC=Cidade
 */
function obterDados() {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado) {
    throw new Error("Acesso negado: Operação não autorizada.");
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
  const data = sheet.getDataRange().getValues();
  const linhas = data.slice(1);
  const procurados = [];

  linhas.forEach((linha, index) => {
    const lat = linha[22]; // W = Latitude
    const lng = linha[23]; // X = Longitude
    const temCoordenadas = lat && lng && lat !== "" && lng !== "";

    const formatarData = (valor) => {
      if (valor instanceof Date) {
        return Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      return String(valor || "").trim();
    };

    procurados.push({
      idLinha: index + 2,
      dataLancamento: formatarData(linha[0]),     // A = Data de Lancamento
      dataConferencia: formatarData(linha[1]),     // B = Data de Conferencia
      mandado: String(linha[2] || ""),             // C = Mandado
      artigo: String(linha[3] || ""),              // D = Artigo
      nome: String(linha[4] || ""),                // E = Nome
      cpf: String(linha[5] || ""),                 // F = CPF
      rg: String(linha[6] || "N/A"),               // G = RG
      nascimento: formatarData(linha[7]),           // H = Nascimento
      naturalidade: String(linha[8] || "N/A"),     // I = Naturalidade
      sexo: String(linha[9] || "N/A"),             // J = Sexo
      cor: String(linha[10] || "N/A"),             // K = Cor
      filiacao: String(linha[11] || "N/A"),        // L = Filiacao
      fotoUrl: String(linha[12] || "N/A"),         // M = Foto URL
      batalhao: String(linha[13] || ""),            // N = Batalhao
      enderecoPrincipal: String(linha[14] || ""),  // O = Endereco Principal
      outrosEnderecos: String(linha[15] || ""),     // P = Outros Enderecos
      status: String(linha[16] || "Procurado"),     // Q = Status
      validade: formatarData(linha[17]),             // R = Validade
      infoProcessuais: String(linha[18] || ""),     // S = Info Processuais
      geodataSecundarios: String(linha[19] || "[]"),// T = Geodata Secundarios
      dadosExtrasJSON: String(linha[20] || "{}"),   // U = Dados Extras JSON
      observacoes: String(linha[21] || ""),          // V = Observacoes
      lat: temCoordenadas ? parseFloat(lat) : null,  // W = Latitude
      lng: temCoordenadas ? parseFloat(lng) : null,  // X = Longitude
      cpi: String(linha[24] || ""),                  // Y = CPI
      bpmArea: String(linha[25] || ""),              // Z = BPM_Area
      ciaArea: String(linha[26] || ""),              // AA = CIA_Area
      dpArea: String(linha[27] || ""),               // AB = DP_Area
      cidade: String(linha[28] || ""),               // AC = Cidade
      semEndereco: !temCoordenadas,                  // Flag para frontend
      // Alias de compatibilidade
      cpiArea: String(linha[24] || ""),              // Alias: cpiArea = CPI
      data: formatarData(linha[0])
    });
  });

  return procurados;
}

/**
 * Cadastra um novo mandado no banco de dados.
 * Inclui DataLancamento automática e dados extras em JSON.
 */
function cadastrarMandadoWebAppPreview(dados) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    throw new Error("Acesso negado: Operação não autorizada.");
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");

    // Verificar duplicidade de mandado
    const todasAsLinhas = sheet.getDataRange().getValues();
    for (let i = 1; i < todasAsLinhas.length; i++) {
      if (todasAsLinhas[i][2] === dados.mandado) {  // C = Mandado (index 2)
        return { sucesso: false, mensagem: "⚠️ REJEITADO: Mandado já cadastrado." };
      }
    }

    const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    // Processar foto
    let urlFotoSalva = "N/A";
    if (dados.fotoBase64 && dados.fotoBase64 !== "") {
      urlFotoSalva = processarEDespacharFotoNoDrive(dados.mandado, dados.fotoBase64);
    }

    // Identificar área (BPM/CIA/DP) pelas coordenadas (Principal)
    var areaInfo = null;
    if (dados.latPrincipal && dados.lngPrincipal) {
      if (dados.ufPrincipal && dados.ufPrincipal !== "SP" && dados.ufPrincipal !== "") {
        areaInfo = { cpi: "FORA DO ESTADO", batalhao: "", cia: "", delegacia: "", cidade: "" };
      } else {
        areaInfo = identificarAreaPorCoordenadas(parseFloat(dados.latPrincipal), parseFloat(dados.lngPrincipal));
      }
    }

    // Identificar área (BPM/CIA/DP) para os secundários
    (dados.secundarios || []).forEach(sec => {
      if (sec.lat && sec.lng) {
        if (sec.uf && sec.uf !== "SP" && sec.uf !== "") {
          sec.cpi = "FORA DO ESTADO";
          sec.batalhao = ""; sec.cia = ""; sec.delegacia = ""; sec.cidade = "";
        } else {
          var areaSec = identificarAreaPorCoordenadas(parseFloat(sec.lat), parseFloat(sec.lng));
          if (areaSec) {
            sec.cpi = areaSec.cpi;
            sec.batalhao = areaSec.batalhao;
            sec.cia = areaSec.cia;
            sec.delegacia = areaSec.delegacia;
            sec.cidade = areaSec.cidade;
          }
        }
      }
    });

    const jsonSecundarios = JSON.stringify(dados.secundarios || []);
    const textoSecundarios = (dados.secundarios || []).map(s => s.endereco).join("\n");
    const jsonExtras = JSON.stringify(dados.extras || {});

    // GRAVAÇÃO NO BANCO (Colunas A até AB)
    sheet.appendRow([
      dataAtual,                // A — Data de Lançamento (automática)
      "",                       // B — Data de Conferência (vazio até Admin conferir)
      dados.mandado,            // C — Mandado
      "Vide Info Proc.",        // D — Artigo
      dados.nome,               // E — Nome
      dados.cpf,                // F — CPF
      dados.rg,                 // G — RG
      dados.nascimento,         // H — Nascimento
      dados.naturalidade,       // I — Naturalidade
      dados.sexo,               // J — Sexo
      dados.cor,                // K — Cor
      dados.filiacao,           // L — Filiação
      urlFotoSalva,             // M — Foto URL
      dados.batalhao,           // N — Batalhão
      dados.enderecoPrincipal,  // O — Endereço Principal
      textoSecundarios,         // P — Outros Endereços
      "Procurado",              // Q — Status
      dados.validade,           // R — Validade
      dados.infoProcessuais,    // S — Info Processuais
      jsonSecundarios,          // T — Geodata Secundários
      jsonExtras,               // U — Dados Extras JSON
      "",                       // V — Observações
      dados.latPrincipal,       // W — Latitude
      dados.lngPrincipal,       // X — Longitude
      areaInfo ? areaInfo.cpi : "",       // Y — CPI
      areaInfo ? areaInfo.batalhao : "",  // Z — BPM_Area
      areaInfo ? areaInfo.cia : "",       // AA — CIA_Area
      areaInfo ? areaInfo.delegacia : "", // AB — DP_Area
      areaInfo ? areaInfo.cidade : ""     // AC — Cidade
    ]);

    // Registrar no Historico
    registrarHistorico(checagem.email, checagem.nome || "Admin", "CADASTRO", "Cadastrou mandado: " + dados.nome + " (" + dados.mandado + ")");

    // Registrar pontos de gamificação para o Admin
    registrarPontosGamificacao(checagem.email, "CADASTRO", 0, dados.mandado, "Cadastrou mandado no sistema");
    sinalizarMudancaMandados();

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
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    return { sucesso: false, mensagem: "Acesso negado." };
  }
  
  if (!listaDados || listaDados.length === 0) {
    return { sucesso: false, mensagem: "Nenhum dado enviado." };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
    if (!sheet) return { sucesso: false, mensagem: "Aba Mandados não encontrada." };
    
    // Filtro para evitar duplicidades que possam ter passado
    const todasAsLinhas = sheet.getDataRange().getValues();
    const mandadosExistentes = todasAsLinhas.slice(1).map(function(r) { return String(r[2] || ""); });
    
    const matrizParaSalvar = [];
    let pontos = 0;
    
    listaDados.forEach(dados => {
      if (mandadosExistentes.includes(dados.mandado)) return; // Pula se já existir
      
      const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      // Identificar área pelas coordenadas (Principal)
      var areaLote = null;
      if (dados.latPrincipal && dados.lngPrincipal) {
        if (dados.ufPrincipal && dados.ufPrincipal !== "SP" && dados.ufPrincipal !== "") {
          areaLote = { cpi: "FORA DO ESTADO", batalhao: "", cia: "", delegacia: "", cidade: "" };
        } else {
          areaLote = identificarAreaPorCoordenadas(parseFloat(dados.latPrincipal), parseFloat(dados.lngPrincipal));
        }
      }

      // Identificar área (BPM/CIA/DP) para os secundários
      (dados.secundarios || []).forEach(sec => {
        if (sec.lat && sec.lng) {
          if (sec.uf && sec.uf !== "SP" && sec.uf !== "") {
            sec.cpi = "FORA DO ESTADO";
            sec.batalhao = ""; sec.cia = ""; sec.delegacia = ""; sec.cidade = "";
          } else {
            var areaSec = identificarAreaPorCoordenadas(parseFloat(sec.lat), parseFloat(sec.lng));
            if (areaSec) {
              sec.cpi = areaSec.cpi;
              sec.batalhao = areaSec.batalhao;
              sec.cia = areaSec.cia;
              sec.delegacia = areaSec.delegacia;
              sec.cidade = areaSec.cidade;
            }
          }
        }
      });

      const jsonSecundarios = JSON.stringify(dados.secundarios || []);
      const textoSecundarios = (dados.secundarios || []).map(s => s.endereco).join("\\n");
      const jsonExtras = JSON.stringify(dados.extras || {});

      matrizParaSalvar.push([
        dataAtual,                // A — Data de Lançamento
        "",                       // B — Data de Conferência
        dados.mandado,            // C — Mandado
        "Vide Info Proc.",        // D — Artigo
        dados.nome,               // E — Nome
        dados.cpf,                // F — CPF
        dados.rg,                 // G — RG
        dados.nascimento,         // H — Nascimento
        dados.naturalidade,       // I — Naturalidade
        dados.sexo,               // J — Sexo
        dados.cor,                // K — Cor
        dados.filiacao,           // L — Filiação
        "N/A",                    // M — Foto URL (lote não envia foto via crop)
        dados.batalhao,           // N — Batalhão
        dados.enderecoPrincipal,  // O — Endereço Principal
        textoSecundarios,         // P — Outros Endereços
        "Procurado",              // Q — Status
        dados.validade,           // R — Validade
        dados.infoProcessuais,    // S — Info Processuais
        jsonSecundarios,          // T — Geodata Secundários
        jsonExtras,               // U — Dados Extras JSON
        "Importado em Lote",      // V — Observações
        dados.latPrincipal,       // W — Latitude
        dados.lngPrincipal,       // X — Longitude
        areaLote ? areaLote.cpi : "",       // Y — CPI
        areaLote ? areaLote.batalhao : "",  // Z — BPM_Area
        areaLote ? areaLote.cia : "",       // AA — CIA_Area
        areaLote ? areaLote.delegacia : "", // AB — DP_Area
        areaLote ? areaLote.cidade : ""     // AC — Cidade
      ]);
      pontos++;
    });
    
    if (matrizParaSalvar.length > 0) {
      const startRow = Math.max(sheet.getLastRow() + 1, 2);
      sheet.getRange(startRow, 1, matrizParaSalvar.length, 29).setValues(matrizParaSalvar);
      
      registrarPontosGamificacao(checagem.email, "CADASTRO", 0, "LOTE_" + pontos, "Cadastrou " + pontos + " mandados em lote");
      sinalizarMudancaMandados();
      return { sucesso: true, mensagem: "Lote salvo com sucesso!" };
    } else {
      return { sucesso: false, mensagem: "Todos os mandados do lote já existiam no banco." };
    }
  } catch (erro) {
    return { sucesso: false, mensagem: "Erro ao gravar lote: " + erro.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Edição completa de um mandado (somente Admin).
 * Recebe objeto com idLinha e campos editados.
 */
function editarMandadoCompleto(dados) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    return { sucesso: false, mensagem: "Acesso negado. Somente Admin pode editar." };
  }

  if (!dados || !dados.idLinha) {
    return { sucesso: false, mensagem: "Linha do registro não informada." };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
    if (!sheet) return { sucesso: false, mensagem: "Aba Mandados não encontrada." };

    const row = dados.idLinha;

    // Mapeamento coluna → index (1-based)
    // Mapeamento coluna → index (1-based)
    var campos = {
      nome: 5, cpf: 6, rg: 7, nascimento: 8, naturalidade: 9,
      sexo: 10, cor: 11, filiacao: 12, fotoBase64: 13, batalhao: 14,
      enderecoPrincipal: 15, outrosEnderecos: 16, status: 17,
      validade: 18, infoProcessuais: 19, observacoes: 22,
      cpiArea: 25, ciaArea: 27, cidade: 29
    };

    // Gravar cada campo editado
    for (var chave in campos) {
      if (dados.hasOwnProperty(chave) && dados[chave] !== undefined) {
        sheet.getRange(row, campos[chave]).setValue(dados[chave]);
      }
    }

    // Se endereço foi editado, tentar geocodificar o Principal
    if (dados.enderecoPrincipal && dados.enderecoPrincipal !== "Não informado" && dados.enderecoPrincipal.length > 5) {
      try {
        var response = Maps.newGeocoder().geocode(dados.enderecoPrincipal);
        if (response.results && response.results.length > 0) {
          var loc = response.results[0].geometry.location;
          sheet.getRange(row, 23).setValue(loc.lat);  // W = Latitude
          sheet.getRange(row, 24).setValue(loc.lng);   // X = Longitude
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
            geoSecundarios.push({ endereco: arraySec[i], lat: lsec.lat, lng: lsec.lng });
          } else {
            geoSecundarios.push({ endereco: arraySec[i], lat: 0, lng: 0 });
          }
          Utilities.sleep(500); // Pausa para não estourar a cota da API Google
        } catch(errSec) {
          geoSecundarios.push({ endereco: arraySec[i], lat: 0, lng: 0 });
        }
      }
      sheet.getRange(row, 20).setValue(JSON.stringify(geoSecundarios)); // T = Geodata Secundários
    } else {
      sheet.getRange(row, 20).setValue("[]");
    }

    registrarPontosGamificacao(checagem.email, "EDICAO", 0, dados.mandado || "N/A", "Editou registro de mandado");
    sinalizarMudancaMandados();

    return { sucesso: true, mensagem: "Mandado atualizado com sucesso!" };
  } catch (erro) {
    return { sucesso: false, mensagem: "Erro ao editar: " + erro.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Exclui um mandado da planilha (somente Admin).
 */
function excluirMandado(idLinha) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    return { sucesso: false, mensagem: "Acesso negado. Somente Admin pode excluir." };
  }
  if (!idLinha || idLinha < 2) {
    return { sucesso: false, mensagem: "Linha inválida." };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
    if (!sheet) return { sucesso: false, mensagem: "Aba Mandados não encontrada." };
    
    const mandadoNome = sheet.getRange(idLinha, 5).getValue(); // E = Nome
    sheet.deleteRow(idLinha);
    
    registrarPontosGamificacao(checagem.email, "EXCLUSAO", 0, "", "Excluiu mandado: " + mandadoNome);
    sinalizarMudancaMandados();
    return { sucesso: true, mensagem: "Mandado excluído com sucesso." };
  } catch (erro) {
    return { sucesso: false, mensagem: "Erro ao excluir: " + erro.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Atualiza status e observações de um mandado.
 * Qualquer usuário Ativo pode usar.
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
    const data = sheet.getDataRange().getValues();
    let atualizados = 0;

    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === mandado) {          // C = Mandado (index 2)
        sheet.getRange(i + 1, 17).setValue(novoStatus);  // Q = Status
        sheet.getRange(i + 1, 22).setValue(novaObs);     // V = Observacoes
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

    if (atualizados > 0) sinalizarMudancaMandados();

    return { sucesso: true, mensagem: atualizados + ' pino(s) atualizado(s) com sucesso.' };
  } catch (erro) {
    return { sucesso: false, mensagem: 'Erro na atualização: ' + erro.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * NOVO: Admin marca a data de conferência de um mandado.
 * Confirma que verificou se o mandado ainda está vigente.
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
    const data = sheet.getDataRange().getValues();
    const dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");

    for (let i = 1; i < data.length; i++) {
      if (data[i][2] === mandado) {          // C = Mandado (index 2)
        sheet.getRange(i + 1, 2).setValue(dataAtual); // Coluna B = Data de Conferência
        sinalizarMudancaMandados();
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

// ================================================================
// HELPER: SALVA FOTO NO GOOGLE DRIVE
// ================================================================

function processarEDespacharFotoNoDrive(idMandado, base64Completo) {
  if (!base64Completo || base64Completo === "" || !base64Completo.startsWith("data:image")) return "N/A";

  try {
    const base64Limpo = base64Completo.split(",")[1];
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Limpo),
      "image/jpeg",
      "Foto_" + idMandado + ".jpg"
    );

    const nomePasta = "Sentinela_Fotos_Mandados";
    let pasta;
    const pastas = DriveApp.getFoldersByName(nomePasta);
    if (pastas.hasNext()) {
      pasta = pastas.next();
    } else {
      pasta = DriveApp.createFolder(nomePasta);
    }

    const arquivo = pasta.createFile(blob);
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return "https://drive.google.com/thumbnail?id=" + arquivo.getId() + "&sz=w500";

  } catch (e) {
    console.error("Erro ao salvar foto no Drive: " + e.message);
    return "N/A";
  }
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
    "CPI", "BPM Area", "CIA Area", "DP Area", "Cidade"
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

/**
 * Migração v2.6 — Adiciona coluna PinoTexto (G) na aba Leis.
 * Executar 1x pelo editor de scripts.
 */
function migrarLeisPinoTexto() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var aba = planilha.getSheetByName("Leis");
  if (!aba) {
    SpreadsheetApp.getUi().alert("Aba Leis nao encontrada.");
    return;
  }

  // Adicionar header PinoTexto se não existir
  var headers = aba.getRange(1, 1, 1, 7).getValues()[0];
  if (String(headers[6] || "").toUpperCase() !== "PINOTEXTO") {
    aba.getRange(1, 7).setValue("PinoTexto");
    aba.getRange(1, 7).setFontWeight("bold").setBackground("#1e293b").setFontColor("#e2e8f0");
  }

  // Popular PinoTexto padrão para linhas existentes sem valor
  var mapaDefault = {
    "HOMICÍDIO": "HOMIC", "HOMICIDIO": "HOMIC",
    "ESTUPRO": "JACK",
    "V. DOMÉSTICA": "PENHA", "V. DOMESTICA": "PENHA",
    "ROUBO": "ROUBO",
    "FURTO": "FURTO",
    "TRÁFICO": "TRÁF", "TRAFICO": "TRÁF",
    "ESTELIONATO": "171",
    "RECEPTAÇÃO": "180", "RECEPTACAO": "180",
    "ALIMENTÍCIA": "CIVIL", "PRISAO CIVIL": "CIVIL",
    "PORTE DE ARMA": "ARMA",
    "TRÂNSITO": "CTB", "TRANSITO": "CTB",
    "FISCAL": "FISCAL",
    "ORGANIZAÇÃO CRIM.": "288", "ORGANIZACAO CRIM.": "288",
    "AMEAÇA": "147", "AMEACA": "147",
    "SEQUESTRO": "148",
    "LESÃO CORPORAL": "LESÃO", "LESAO CORPORAL": "LESÃO",
    "OUTROS": "CRIME"
  };

  var lastRow = aba.getLastRow();
  if (lastRow > 1) {
    var data = aba.getRange(2, 1, lastRow - 1, 7).getValues();
    for (var i = 0; i < data.length; i++) {
      if (!data[i][6]) {
        var cat = String(data[i][0] || "").toUpperCase().trim();
        data[i][6] = mapaDefault[cat] || cat.substring(0, 5);
      }
    }
    aba.getRange(2, 1, data.length, 7).setValues(data);
  }

  SpreadsheetApp.getUi().alert("Migracao v2.6 concluida!\n\n- Coluna PinoTexto (G) adicionada na aba Leis\n- Valores padrão populados");
}

// ================================================================
// LEIS — CRUD (Admin)
// ================================================================

/**
 * Retorna todas as leis/artigos cadastrados (Admin).
 */
function obterLeis() {
  var checagem = verificarAcessoUsuario();
  if (!checagem.autorizado) throw new Error("Acesso negado.");

  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leis");
  if (!aba || aba.getLastRow() <= 1) return [];

  var data = aba.getDataRange().getValues();
  var leis = [];

  for (var i = 1; i < data.length; i++) {
    leis.push({
      id: i + 1,
      categoria: String(data[i][0] || ""),
      palavrasChave: String(data[i][1] || ""),
      cor: String(data[i][2] || "#6b7280"),
      svgPath: String(data[i][3] || ""),
      ordem: parseInt(data[i][4]) || 99,
      ativo: String(data[i][5] || "SIM").toUpperCase() === "SIM",
      pinoTexto: String(data[i][6] || "")
    });
  }

  leis.sort(function(a, b) { return a.ordem - b.ordem; });
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

  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leis");
  if (!aba || aba.getLastRow() <= 1) return [];

  var data = aba.getDataRange().getValues();
  var tabela = [];

  for (var i = 1; i < data.length; i++) {
    var ativo = String(data[i][5] || "SIM").toUpperCase();
    if (ativo !== "SIM") continue;

    tabela.push({
      categoria: String(data[i][0] || ""),
      palavrasChave: String(data[i][1] || ""),
      cor: String(data[i][2] || "#6b7280"),
      pinoTexto: String(data[i][6] || ""),
      ordem: parseInt(data[i][4]) || 99
    });
  }

  tabela.sort(function(a, b) { return a.ordem - b.ordem; });
  return tabela;
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

  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leis");
  if (!aba) return { sucesso: false, mensagem: "Aba Leis nao encontrada." };

  var linha = [
    dados.categoria || "",
    dados.palavrasChave || "",
    dados.cor || "#6b7280",
    dados.svgPath || "",
    parseInt(dados.ordem) || 99,
    dados.ativo ? "SIM" : "NAO",
    dados.pinoTexto || ""
  ];

  if (dados.id && dados.id > 1) {
    // Atualizar existente
    aba.getRange(dados.id, 1, 1, 7).setValues([linha]);
  } else {
    // Nova lei
    aba.appendRow(linha);
  }

  registrarHistorico(checagem.email, checagem.nome, "LEIS", "Salvou categoria: " + dados.categoria);
  return { sucesso: true, mensagem: "Lei salva com sucesso!" };
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

  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leis");
  if (!aba) return { sucesso: false, mensagem: "Aba Leis nao encontrada." };

  var data = aba.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toUpperCase() === String(categoria).trim().toUpperCase()) {
      var atual = String(data[i][1] || "");
      var novas = String(novasPalavras || "").split("|").map(function(p) { return p.trim().toUpperCase(); }).filter(function(p) { return p; });
      var existentes = atual.toUpperCase().split("|").map(function(p) { return p.trim(); });
      
      // Só adicionar palavras que não existem ainda
      novas.forEach(function(n) {
        if (existentes.indexOf(n) === -1) {
          atual += (atual ? "|" : "") + n;
        }
      });
      
      aba.getRange(i + 1, 2).setValue(atual);
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

  var aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leis");
  if (!aba || idLinha < 2) return { sucesso: false, mensagem: "Linha invalida." };

  var nome = aba.getRange(idLinha, 1).getValue();
  aba.deleteRow(idLinha);

  registrarHistorico(checagem.email, checagem.nome, "LEIS", "Excluiu categoria: " + nome);
  return { sucesso: true, mensagem: "Lei excluida com sucesso!" };
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
    PropertiesService.getScriptProperties().setProperty('MandadosLastUpdate', new Date().getTime().toString());
  } catch(e) {}
}

/**
 * Trigger simples do Google Sheets. Caso alguém altere direto na planilha.
 */
function onEdit(e) {
  if (e && e.range) {
    var sheet = e.range.getSheet();
    if (sheet.getName() === "Mandados") {
      sinalizarMudancaMandados();
    }
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
