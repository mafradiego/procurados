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
    'TipoImportacao'
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
      fotoUrl: String(col('Foto URL') || "N/A"),
      batalhao: String(col('Batalhao') || ""),
      enderecoPrincipal: String(col('Endereco Principal') || ""),
      outrosEnderecos: String(col('Outros Enderecos') || ""),
      status: String(col('Status') || "Procurado"),
      validade: formatarData(col('Validade')),
      infoProcessuais: String(col('Info Processuais') || ""),
      geodataSecundarios: String(col('Geodata Secundarios') || "[]"),
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
    if (dados.fotoBase64 && dados.fotoBase64 !== "") {
      urlFotoSalva = processarEDespacharFotoNoDrive(dados.mandado, dados.fotoBase64);
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
    const jsonExtras = JSON.stringify(dados.extras);

    // Montar linha na ordem dos cabeçalhos usando mapa dinâmico
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const novaLinha = new Array(headers.length).fill("");
    
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
    novaLinha[col['Foto URL']] = urlFotoSalva;
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

      // Limpar campos de área, mas MANTER coordenadas vindas do front-end
      (dados.secundarios || []).forEach(sec => {
        sec.cpi = "";
        sec.batalhao = "";
        sec.cia = "";
        sec.delegacia = "";
        sec.cidade = "";
      });

      const jsonSecundarios = JSON.stringify(dados.secundarios || []);
      const textoSecundarios = (dados.secundarios || []).map(s => s.endereco).join("\\n");
      dados.extras = dados.extras || {};
      dados.extras.Criado_Por = checagem.email;
      const jsonExtras = JSON.stringify(dados.extras);

      // IDENTIFICAÇÃO ESPACIAL DE ÁREAS (CPI/BTL/CIA)
      let areaInfo = { cpi: "", batalhao: "", cia: "", delegacia: "", cidade: "" };
      if (dados.latPrincipal && dados.lngPrincipal) {
        const areaDetectada = identificarAreaPorCoordenadas(dados.latPrincipal, dados.lngPrincipal);
        if (areaDetectada) {
          areaInfo = areaDetectada;
        }
      }

      matrizParaSalvar.push([
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
        "N/A",                    // M — Foto URL (lote não envia foto via crop)
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
      ]);
      pontos++;
    });
    
    if (matrizParaSalvar.length > 0) {
      const startRow = Math.max(sheet.getLastRow() + 1, 2);
      sheet.getRange(startRow, 1, matrizParaSalvar.length, 30).setValues(matrizParaSalvar);
      sheet.getRange(startRow, 23, matrizParaSalvar.length, 2).setNumberFormat("0.00000000");
      SpreadsheetApp.flush();
      
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

  if (!dados || !dados.idLinha) {
    return { sucesso: false, mensagem: "Linha do registro não informada." };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
    if (!sheet) return { sucesso: false, mensagem: "Aba Mandados não encontrada." };

    const col = obterMapaColunas(sheet);

    if (checagem.perfil === "Colaborador") {
      const idxExtras = col['Dados Extras JSON'];
      const jsonExtrasBanco = sheet.getRange(dados.idLinha, idxExtras + 1).getValue();
      try {
        const extrasBanco = JSON.parse(jsonExtrasBanco || "{}");
        if (extrasBanco.Criado_Por !== checagem.email) {
          return { sucesso: false, mensagem: "Acesso negado. Você só pode editar mandados que você mesmo cadastrou." };
        }
      } catch (e) {
        return { sucesso: false, mensagem: "Acesso negado. Autoria do mandado não pôde ser confirmada." };
      }
    }

    const row = dados.idLinha;

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
      fotoBase64: col['Foto URL'],
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
        sheet.getRange(row, campos[chave] + 1).setValue(dados[chave]);
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
      sheet.getRange(row, col['Geodata Secundarios'] + 1).setValue(JSON.stringify(geoSecundarios));
    } else {
      sheet.getRange(row, col['Geodata Secundarios'] + 1).setValue("[]");
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
 * v4.1.0: Usa cabeçalhos dinâmicos.
 */
function excluirMandado(idLinha) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil === "Patrulheiro") {
    return { sucesso: false, mensagem: "Acesso negado. Patrulheiros não podem excluir." };
  }
  if (!idLinha || idLinha < 2) {
    return { sucesso: false, mensagem: "Linha inválida." };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
    if (!sheet) return { sucesso: false, mensagem: "Aba Mandados não encontrada." };

    const col = obterMapaColunas(sheet);

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

    for (let i = 1; i < data.length; i++) {
      if (data[i][idxMandado] === mandado) {
        if (novoStatus) sheet.getRange(i + 1, idxStatus + 1).setValue(novoStatus);
        
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
            hist.push({ data: dataHoje, usuario: usuarioNome, texto: novaObs });
            
            sheet.getRange(i + 1, idxObs + 1).setValue(JSON.stringify(hist));
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

    if (atualizados > 0) sinalizarMudancaMandados();

    return { sucesso: true, mensagem: atualizados + ' pino(s) atualizado(s) com sucesso.' };
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
      if (data[i][idxMandado] === mandado) {
        sheet.getRange(i + 1, idxConferencia + 1).setValue(dataAtual);
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
      if (data[i][idxMandado] === mandado) {
        let extra = {};
        try { extra = JSON.parse(data[i][idxJSON] || '{}'); } catch(e){}
        let dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        let usuario = checagem.email;
        extra.bnmpConferencia = statusBnmp + " em " + dataHoje + " por " + usuario;
        sheet.getRange(i + 1, idxJSON + 1).setValue(JSON.stringify(extra));
        sinalizarMudancaMandados();
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
      ordem: parseInt(data[i][4]) || 99,
      leiNome: String(data[i][7] || ""),
      numeroLei: String(data[i][8] || ""),
      artigo: String(data[i][9] || ""),
      paragrafo: String(data[i][10] || ""),
      inciso: String(data[i][11] || ""),
      tipificacaoCompleta: String(data[i][12] || "")
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
    const ts = new Date().getTime().toString();
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

  const permissoes = verificarPermissoes(Session.getActiveUser().getEmail());
  if (permissoes.nivel !== "administrador" && permissoes.nivel !== "colaborador") {
    throw new Error("Sem permissão para editar os mandados.");
  }
  
  const usuario = permissoes.nome;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mandados");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const idxNumeroMandado = headers.indexOf('Numero Mandado');
  const idxInfoProcessuais = headers.indexOf('Info Processuais');
  const idxDadosExtrasJSON = headers.indexOf('Dados Extras JSON');
  
  if (idxNumeroMandado === -1) throw new Error("Coluna 'Numero Mandado' não encontrada.");

  let linhaAtualizar = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxNumeroMandado]).trim() === String(mandado.numeroMandado).trim()) {
      linhaAtualizar = i + 1;
      break;
    }
  }

  if (linhaAtualizar === -1) {
    throw new Error("Mandado " + mandado.numeroMandado + " não encontrado na base de dados.");
  }
  
  // Mapeamento de colunas para dados simples.
  const colunasSimples = {
    'Nome': mandado.nome,
    'Nome da Mae': mandado.nomeMae,
    'Nascimento': mandado.nascimento,
    'Sexo': mandado.sexo,
    'Cor': mandado.cor,
    'Filiacao': mandado.filiacao,
    'Naturalidade': mandado.naturalidade,
    'Foto URL': mandado.fotoUrl,
    'Endereco Principal': mandado.enderecoPrincipal
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

  // Atualizar Dados Extras JSON e Marcação de Conferência
  if (idxDadosExtrasJSON !== -1) {
    let extraObj = {};
    try {
      const extraStr = sheet.getRange(linhaAtualizar, idxDadosExtrasJSON + 1).getValue();
      if (extraStr) extraObj = JSON.parse(extraStr);
    } catch(e) {}
    
    if (mandado.tipoMandado !== undefined) extraObj.titulo = mandado.tipoMandado;
    if (mandado.dataBNMP !== undefined) extraObj.emissão = mandado.dataBNMP;
    
    if (marcarConferido) {
      const hoje = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
      extraObj.bnmpConferencia = "Conferido/Editado em " + hoje + " por " + usuario;
    }
    
    sheet.getRange(linhaAtualizar, idxDadosExtrasJSON + 1).setValue(JSON.stringify(extraObj));
  }
  
  limparCacheBD();
  return { mensagem: "Mandado " + mandado.numeroMandado + " atualizado com sucesso!" };
}

/**
 * Faz o upload de uma imagem em base64 para a pasta específica do Google Drive
 * Retorna a URL pública de compartilhamento para ser setada no card.
 */
function uploadFotoIndividuoDrive(base64Data, nomeIndividuo, numeroMandado) {
  const permissoes = verificarPermissoes(Session.getActiveUser().getEmail());
  if (permissoes.nivel !== "administrador" && permissoes.nivel !== "colaborador") {
    throw new Error("Sem permissão para fazer upload de fotos.");
  }

  const folderId = "1QHK_Bc-XXIrm4dSWj7XjhfmUrWLydw--";
  const folder = DriveApp.getFolderById(folderId);
  
  // Remove o cabeçalho base64, se houver
  const mimeType = "image/jpeg";
  const base64Str = base64Data.replace(/^data:image\/\w+;base64,/, "");
  
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Str), mimeType, nomeIndividuo.replace(/\s+/g, '_') + "_" + numeroMandado + ".jpg");
  
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return file.getDownloadUrl().replace("&gd=true", "");
}
    const props = PropertiesService.getScriptProperties();
    props.setProperty('MandadosLastUpdate', ts);
    props.setProperty('DB_UPDATE_TIMESTAMP', ts);
  } catch(e) {}
}

/**
 * Trigger simples do Google Sheets. Caso alguém altere direto na planilha.
 */
function onEdit(e) {
  if (e && e.range) {
    var sheet = e.range.getSheet();
    if (sheet.getName() === "Mandados" || sheet.getName() === "Poligonos") {
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
    var dados = sheet.getDataRange().getValues();
    var mandadoBusca = String(payload.mandado).trim();
    var linhaEncontrada = -1;
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][2] || "").trim() === mandadoBusca) { linhaEncontrada = i + 1; break; }
    }
    if (linhaEncontrada < 0) return { sucesso: false, mensagem: "Mandado nao encontrado: " + mandadoBusca };

    var detalhes = [];

    // Nº do Mandado (coluna C = 3)
    if (payload.novoNumeroMandado && payload.novoNumeroMandado.trim() && payload.novoNumeroMandado.trim() !== mandadoBusca) {
      sheet.getRange(linhaEncontrada, 3).setValue(payload.novoNumeroMandado.trim());
      detalhes.push("NovoMandado");
    }

    // Artigo (coluna D = 4)
    if (payload.artigo && payload.artigo.trim()) {
      sheet.getRange(linhaEncontrada, 4).setValue(payload.artigo.trim());
      detalhes.push("Artigo");
    }

    // Endereço Principal (coluna O = 15)
    if (payload.enderecoPrincipal && payload.enderecoPrincipal.trim()) {
      sheet.getRange(linhaEncontrada, 15).setValue(payload.enderecoPrincipal.trim());
      
      // Limpar Lat/Lng para forçar geocodificação
      sheet.getRange(linhaEncontrada, 23).clearContent();
      sheet.getRange(linhaEncontrada, 24).clearContent();
      try {
        executarGeocodificacaoLinha(sheet, linhaEncontrada, false);
      } catch(e) {}
      
      detalhes.push("Endereco");
    }

    // Info Processual (coluna S = 19) - pode conter novoNumeroProcesso embutido
    var infoBase = payload.infoProcessuais && payload.infoProcessuais.trim()
      ? payload.infoProcessuais.trim()
      : String(dados[linhaEncontrada - 1][18] || "");

    if (payload.novoNumeroProcesso && payload.novoNumeroProcesso.trim()) {
      // Substituir "Nº do processo: XXX" dentro do infoProcessuais
      var novoProc = payload.novoNumeroProcesso.trim();
      if (/N[º°]?\s*do processo:/i.test(infoBase)) {
        infoBase = infoBase.replace(/N[º°]?\s*do processo:\s*([^|]+)/i, "Nº do processo: " + novoProc);
      } else {
        infoBase = "Nº do processo: " + novoProc + (infoBase ? " | " + infoBase : "");
      }
      detalhes.push("NumProcesso");
    }

    if (payload.infoProcessuais && payload.infoProcessuais.trim()) {
      sheet.getRange(linhaEncontrada, 19).setValue(infoBase);
      detalhes.push("InfoProcessual");
    } else if (payload.novoNumeroProcesso && payload.novoNumeroProcesso.trim()) {
      sheet.getRange(linhaEncontrada, 19).setValue(infoBase);
    }

    // Título — atualiza dadosExtrasJSON.titulo (coluna U = 21)
    if (payload.novoTitulo && payload.novoTitulo.trim()) {
      try {
        var extrasStr = String(dados[linhaEncontrada - 1][20] || "{}");
        var extras = JSON.parse(extrasStr);
        extras.titulo = payload.novoTitulo.trim();
        sheet.getRange(linhaEncontrada, 21).setValue(JSON.stringify(extras));
        detalhes.push("Titulo");
      } catch(e) {
        // Se JSON inválido, cria novo
        sheet.getRange(linhaEncontrada, 21).setValue(JSON.stringify({ titulo: payload.novoTitulo.trim() }));
        detalhes.push("Titulo");
      }
    }

    // Cidade (coluna AC = 29)
    if (payload.cidade && payload.cidade.trim()) {
      sheet.getRange(linhaEncontrada, 29).setValue(payload.cidade.trim());
      detalhes.push("Cidade");
    }

    if (detalhes.length === 0) return { sucesso: false, mensagem: "Nenhum campo valido fornecido." };

    sinalizarMudancaMandados();
    registrarHistorico(checagem.email, checagem.nome, "AUDITORIA",
      "Corrigiu mandado " + mandadoBusca + ": [" + detalhes.join(", ") + "]");
    return { sucesso: true, mensagem: "Mandado " + mandadoBusca + " corrigido. Campos: " + detalhes.join(", ") + "." };
  } catch (erro) {
    return { sucesso: false, mensagem: "Erro: " + erro.message };
  } finally {
    lock.releaseLock();
  }
}
