// ================================================================
// SENTINELA v2.0 — MÓDULO DE SEGURANÇA E CONTROLE DE ACESSO
// ================================================================

/**
 * Verifica o perfil e status de quem está acessando.
 * Retorna: { autorizado, perfil, nome, unidade, email, fotoUrl, pontos }
 */
function verificarAcessoUsuario() {
  const emailAtivo = Session.getActiveUser().getEmail();

  if (!emailAtivo) {
    return { autorizado: false, motivo: "Sessão Google não identificada." };
  }

  // 1. Tentar ler do Cache do Servidor para evitar abrir a planilha repetidas vezes
  const cache = CacheService.getScriptCache();
  const cacheKey = "AUTH_V2_" + emailAtivo.toLowerCase();
  const acessoCacheado = cache.get(cacheKey);

  if (acessoCacheado) {
    return JSON.parse(acessoCacheado);
  }

  // 2. Se não tem cache, abre a planilha
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaUsuarios = planilha.getSheetByName("Usuarios");

  if (!abaUsuarios) {
    return { autorizado: false, motivo: "Erro crítico: Tabela de controle ausente. Execute o Setup." };
  }

  const dados = abaUsuarios.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    const emailPlanilha = dados[i][0].toString().trim().toLowerCase();
    if (emailPlanilha === emailAtivo.toLowerCase()) {
      const nome = dados[i][1];
      const unidade = dados[i][2];
      
      // Normalizar o perfil para evitar problemas de maiúsculo/minúsculo ou espaços
      let perfil = dados[i][3] ? dados[i][3].toString().trim() : "";
      const perfilLower = perfil.toLowerCase();
      if (perfilLower === "admin") {
        perfil = "Admin";
      } else if (perfilLower === "colaborador") {
        perfil = "Colaborador";
      } else {
        perfil = "Patrulheiro"; // Default fallback
      }

      const status = dados[i][4].toString().trim().toUpperCase();
      const fotoUrl = dados[i][5] ? dados[i][5].toString().trim() : "";
      const pontos = parseInt(dados[i][6]) || 0;
      const badges = dados[i][7] ? dados[i][7].toString().trim() : "";

      if (status === "ATIVO") {
        // Atualiza a coluna J (Último Login), K (Qtd Acessos) e L (Status de Conexão)
        let qtdAcessos = parseInt(dados[i][10]) || 0;
        qtdAcessos++;
        abaUsuarios.getRange(i + 1, 10).setValue(new Date());
        abaUsuarios.getRange(i + 1, 11).setValue(qtdAcessos);
        abaUsuarios.getRange(i + 1, 12).setValue("Online");

        let checagem = {
          autorizado: true,
          perfil: perfil,
          nome: nome,
          unidade: unidade,
          email: emailAtivo,
          fotoUrl: fotoUrl,
          pontos: pontos,
          badges: badges
        };
        // Salva as credenciais no cache do Google por 5 minutos (300 segundos) para propagação rápida
        cache.put(cacheKey, JSON.stringify(checagem), 300);
        return checagem;
      } else if (status === "PENDENTE") {
        return {
          autorizado: false,
          motivo: "VALIDAÇÃO PENDENTE! AGUARDE LIBERAÇÃO DO ADMINISTRADOR",
          status: "PENDENTE"
        };
      } else if (status === "BLOQUEADO" || status === "INATIVO") {
        return {
          autorizado: false,
          motivo: "ACESSO RESTRITO / BLOQUEADO. ENTRAR EM CONTATO COM O ADMINISTRADOR.",
          status: "BLOQUEADO"
        };
      }
    }
  }

  // Não cadastrado
  return { autorizado: false, status: "NAO_CADASTRADO", email: emailAtivo };
}

/**
 * Registra um novo usuário como Pendente.
 * Separando Nome de Guerra e Unidade.
 */
function registrarNovoUsuario(nomeGuerra, unidade) {
  try {
    const emailAtivo = Session.getActiveUser().getEmail();
    if (!emailAtivo) return { sucesso: false, mensagem: "E-mail não identificado." };

    const planilha = SpreadsheetApp.getActiveSpreadsheet();
    const abaUsuarios = planilha.getSheetByName("Usuarios");

    const dados = abaUsuarios.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0].toString().trim().toLowerCase() === emailAtivo.toLowerCase()) {
        return { sucesso: false, mensagem: "Você já possui uma solicitação em andamento." };
      }
    }

    const dataAtual = Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss");
    abaUsuarios.appendRow([
      emailAtivo,         // A - Email
      nomeGuerra,         // B - Nome
      unidade,            // C - Unidade
      "Patrulheiro",      // D - Perfil (padrão)
      "Pendente",         // E - Status (padrão)
      "",                 // F - Foto URL
      "0",                // G - Pontos
      "",                 // H - Badges
      dataAtual,          // I - Data de Cadastro
      "",                 // J - Último Login
      "0",                // K - Qtd Acessos
      "Offline"           // L - Status de Conexão
    ]);

    return { sucesso: true, mensagem: "Cadastro realizado com sucesso!" };

  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao registrar: " + e.message };
  }
}

/**
 * Busca lista de operadores pendentes (Admin only).
 */
function obterUsuariosPendentes() {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") throw new Error("Acesso negado.");

  const abaUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
  if (!abaUsuarios) return [];

  const dados = abaUsuarios.getDataRange().getValues();
  const pendentes = [];

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][4].toString().trim().toUpperCase() === "PENDENTE") {
      pendentes.push({
        email: dados[i][0],
        nome: dados[i][1],
        unidade: dados[i][2]
      });
    }
  }
  return pendentes;
}

/**
 * Busca lista de todos os operadores cadastrados (Admin only).
 */
function obterTodosUsuarios() {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") throw new Error("Acesso negado.");

  const abaUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
  if (!abaUsuarios) return [];

  // Use a range up to column 12 (L) to ensure Status de Conexão is fetched
  const ultimoRow = abaUsuarios.getLastRow();
  if (ultimoRow < 1) return [];
  const dados = abaUsuarios.getRange(1, 1, ultimoRow, 12).getValues();
  const todos = [];

  const limiteMs = 5 * 60 * 1000; // 5 minutos de inatividade máxima (limite de produção)
  const agoraMs = new Date().getTime();

  const isDate = function(obj) {
    return obj && (Object.prototype.toString.call(obj) === '[object Date]' || typeof obj.getTime === 'function');
  };

  function parseLocalStringToTime(str) {
      if (!str) return 0;
      let s = str.toString().replace(" (Offline)", "").replace("Offline", "").trim();
      s = s.replace(/\s*\(Hora.*?\)/g, "").trim(); // Remove timezone strings like "(Hora padrão de Brasília)"
      
      if (s.includes("/")) {
         let parts = s.split(' ');
         if(parts.length < 2) return 0;
         let d = parts[0].split('/');
         let t = parts[1].split(':');
         if(d.length < 3 || t.length < 3) return 0;
         // Use Date.UTC to construct timezone-agnostic timestamp, adjusting for America/Sao_Paulo (GMT-3)
         return Date.UTC(parseInt(d[2], 10), parseInt(d[1], 10) - 1, parseInt(d[0], 10), parseInt(t[0], 10), parseInt(t[1], 10), parseInt(t[2], 10)) + (3 * 60 * 60 * 1000);
      }
      
      let dt = new Date(s);
      return isNaN(dt.getTime()) ? 0 : dt.getTime();
  }

  function formatarDataBr(d, isDataHora) {
    if (!d) return isDataHora ? "Nunca" : "Desconhecida";
    if (isDate(d)) {
      return Utilities.formatDate(d, "America/Sao_Paulo", isDataHora ? "dd/MM/yyyy HH:mm:ss" : "dd/MM/yyyy HH:mm");
    }
    let str = d.toString().trim();
    if (str.includes("/")) {
      return str;
    }
    let cleanStr = str.replace(" (Offline)", "").replace("Offline", "").replace(/\s*\(Hora.*?\)/g, "").trim();
    let parsed = new Date(cleanStr);
    if (!isNaN(parsed.getTime())) {
       let formated = Utilities.formatDate(parsed, "America/Sao_Paulo", isDataHora ? "dd/MM/yyyy HH:mm:ss" : "dd/MM/yyyy HH:mm");
       if (str.includes("(Offline)")) formated += " (Offline)";
       return formated;
    }
    return str;
  }

  for (let i = 1; i < dados.length; i++) {
    let ultimoLoginReal = dados[i][9];
    let statusConexaoReal = dados[i][11] ? dados[i][11].toString().trim() : "Offline";
    
    let isOffline = (statusConexaoReal === "Offline") || (!ultimoLoginReal) || (ultimoLoginReal === "Nunca");
    
    // Normalize to a string in the America/Sao_Paulo timezone
    let dateStr = "";
    if (isDate(ultimoLoginReal)) {
      dateStr = Utilities.formatDate(ultimoLoginReal, "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss");
    } else {
      dateStr = ultimoLoginReal ? ultimoLoginReal.toString().trim() : "";
    }

    // O Ceifador: Se está lido como Online na planilha, mas passou de X segundos...
    if (!isOffline) {
       let epochUltimo = 0;
       if (isDate(ultimoLoginReal)) {
         epochUltimo = ultimoLoginReal.getTime();
       } else {
         epochUltimo = parseLocalStringToTime(ultimoLoginReal);
       }

       if (epochUltimo > 0 && (agoraMs - epochUltimo) > limiteMs) {
           statusConexaoReal = "Offline";
           abaUsuarios.getRange(i + 1, 12).setValue("Offline"); // Escreve permanentemente o status Offline na coluna L
           isOffline = true;
       }
    }

    todos.push({
      email: dados[i][0],
      nome: dados[i][1],
      unidade: dados[i][2],
      perfil: dados[i][3],
      status: dados[i][4].toString().trim().toUpperCase(),
      fotoUrl: dados[i][5] ? dados[i][5].toString().trim() : "",
      dataCadastro: formatarDataBr(dados[i][8], false),
      ultimoLogin: formatarDataBr(dateStr, true),
      qtdAcessos: parseInt(dados[i][10]) || 0,
      isOnline: !isOffline
    });
  }
  return todos;
}

/**
 * Aprova todos os usuários pendentes (Admin only).
 */
function aprovarTodosUsuariosPendentes() {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") throw new Error("Acesso negado.");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const abaUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
    const dados = abaUsuarios.getDataRange().getValues();
    let cont = 0;

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][4].toString().trim().toUpperCase() === "PENDENTE") {
        abaUsuarios.getRange(i + 1, 5).setValue("ATIVO");
        CacheService.getScriptCache().remove("AUTH_V2_" + dados[i][0].toString().trim().toLowerCase());
        cont++;
      }
    }
    SpreadsheetApp.flush();
    return { sucesso: true, mensagem: cont + " usuário(s) aprovado(s) com sucesso." };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro no servidor: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Altera o status de um operador (ATIVO ou BLOQUEADO). Admin only.
 */
function alterarStatusOperador(emailAlvo, novoStatus) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") throw new Error("Acesso negado.");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const abaUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
    const dados = abaUsuarios.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0].toString().trim().toLowerCase() === emailAlvo.toLowerCase()) {
        
        if (dados[i][3] === "Admin" && novoStatus !== "ATIVO") {
          let qtdAdmins = 0;
          for (let j = 1; j < dados.length; j++) {
            if (dados[j][3] === "Admin" && dados[j][4].toString().trim().toUpperCase() === "ATIVO") qtdAdmins++;
          }
          if (qtdAdmins <= 1) {
            return { sucesso: false, mensagem: "Ação negada: Deve haver pelo menos 1 Administrador ATIVO no sistema." };
          }
        }

        abaUsuarios.getRange(i + 1, 5).setValue(novoStatus);
        CacheService.getScriptCache().remove("AUTH_V2_" + emailAlvo.toLowerCase());
        SpreadsheetApp.flush();
        return { sucesso: true, mensagem: "Status alterado para " + novoStatus + " com sucesso." };
      }
    }
    return { sucesso: false, mensagem: "Operador não localizado no banco de dados." };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro no servidor: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Edita Nome e Unidade de um operador (Admin only).
 */
function editarDadosOperador(emailAlvo, novoNome, novaUnidade) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") throw new Error("Acesso negado.");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const abaUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
    const dados = abaUsuarios.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0].toString().trim().toLowerCase() === emailAlvo.toLowerCase()) {
        abaUsuarios.getRange(i + 1, 2).setValue(novoNome);
        abaUsuarios.getRange(i + 1, 3).setValue(novaUnidade);
        CacheService.getScriptCache().remove("AUTH_V2_" + emailAlvo.toLowerCase());
        SpreadsheetApp.flush();
        return { sucesso: true, mensagem: "Dados atualizados com sucesso." };
      }
    }
    return { sucesso: false, mensagem: "Operador não localizado." };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro no servidor: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Adiciona um usuário diretamente pelo Admin (Gatekeeper).
 */
function adicionarUsuarioDireto(emailAdd, perfilAdd, nomeAdd) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") {
    return { sucesso: false, mensagem: "Acesso negado." };
  }
  
  if (!emailAdd || !perfilAdd || !nomeAdd) return { sucesso: false, mensagem: "Email, Nome e Perfil são obrigatórios." };
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const abaUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
    const dados = abaUsuarios.getDataRange().getValues();
    
    // Verifica se já existe
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0].toString().trim().toLowerCase() === emailAdd.toLowerCase()) {
        return { sucesso: false, mensagem: "Usuário já existe no banco de dados." };
      }
    }
    
    const dataAtual = Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss");
    abaUsuarios.appendRow([
      emailAdd.trim().toLowerCase(),
      nomeAdd.trim(),
      "Unidade a Definir",
      perfilAdd,
      "ATIVO",
      "",
      "0",
      "",
      dataAtual,          // I - Data de Cadastro
      "",                 // J - Último Login
      "0",                // K - Qtd Acessos
      "Offline"           // L - Status de Conexão
    ]);
    
    SpreadsheetApp.flush();
    return { sucesso: true, mensagem: "Usuário adicionado com sucesso." };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro no servidor: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Altera o perfil de um operador (Admin only).
 */
function alterarPerfilUsuario(emailAlvo, novoPerfil) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") throw new Error("Acesso negado.");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const abaUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
    const dados = abaUsuarios.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0].toString().trim().toLowerCase() === emailAlvo.toLowerCase()) {
        
        if (dados[i][3] === "Admin" && novoPerfil !== "Admin") {
          let qtdAdmins = 0;
          for (let j = 1; j < dados.length; j++) {
            if (dados[j][3] === "Admin" && dados[j][4].toString().trim().toUpperCase() === "ATIVO") qtdAdmins++;
          }
          if (qtdAdmins <= 1) {
            return { sucesso: false, mensagem: "Ação negada: Deve haver pelo menos 1 Administrador ATIVO no sistema." };
          }
        }

        abaUsuarios.getRange(i + 1, 4).setValue(novoPerfil); // Coluna D = 4
        CacheService.getScriptCache().remove("AUTH_V2_" + emailAlvo.toLowerCase());
        SpreadsheetApp.flush();
        return { sucesso: true, mensagem: "Perfil alterado para " + novoPerfil };
      }
    }
    return { sucesso: false, mensagem: "Operador não localizado." };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro no servidor: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Kill Switch: Exclui o usuário definitivamente (Admin only).
 */
function excluirUsuarioDefinitivo(emailAlvo) {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil !== "Admin") throw new Error("Acesso negado.");

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const abaUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
    const dados = abaUsuarios.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0].toString().trim().toLowerCase() === emailAlvo.toLowerCase()) {
        
        if (dados[i][3] === "Admin") {
          let qtdAdmins = 0;
          for (let j = 1; j < dados.length; j++) {
            if (dados[j][3] === "Admin" && dados[j][4].toString().trim().toUpperCase() === "ATIVO") qtdAdmins++;
          }
          if (qtdAdmins <= 1) {
            return { sucesso: false, mensagem: "Ação negada: Deve haver pelo menos 1 Administrador ATIVO no sistema." };
          }
        }

        abaUsuarios.deleteRow(i + 1);
        CacheService.getScriptCache().remove("AUTH_V2_" + emailAlvo.toLowerCase());
        SpreadsheetApp.flush();
        return { sucesso: true, mensagem: "Acesso revogado instantaneamente." };
      }
    }
    return { sucesso: false, mensagem: "Operador não localizado." };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro no servidor: " + e.message };
  } finally {
    lock.releaseLock();
  }
}


/**
 * Atualiza a foto do perfil do usuário logado.
 * Salva a foto na pasta Sentinela_Fotos_Usuarios.
 */
function atualizarFotoUsuario(fotoBase64) {
  const emailAtivo = Session.getActiveUser().getEmail();
  if (!emailAtivo) return { sucesso: false, mensagem: "E-mail não identificado." };

  try {
    let urlFoto = "";

    if (fotoBase64 && (fotoBase64.startsWith("http://") || fotoBase64.startsWith("https://"))) {
      urlFoto = fotoBase64; // Salva o link direto do avatar, não baixa no Drive.
    } else if (fotoBase64 && fotoBase64.startsWith("data:image")) {
      const base64Limpo = fotoBase64.split(",")[1];
      const blob = Utilities.newBlob(
        Utilities.base64Decode(base64Limpo),
        "image/jpeg",
        "Foto_" + emailAtivo.replace(/[@.]/g, "_") + ".jpg"
      );

      const nomePasta = "Sentinela_Fotos_Usuarios";
      let pasta;
      const pastas = DriveApp.getFoldersByName(nomePasta);
      if (pastas.hasNext()) {
        pasta = pastas.next();
      } else {
        pasta = DriveApp.createFolder(nomePasta);
      }

      // Procurar e excluir foto antiga
      const arquivosAntigos = pasta.getFilesByName("Foto_" + emailAtivo.replace(/[@.]/g, "_") + ".jpg");
      while (arquivosAntigos.hasNext()) {
        arquivosAntigos.next().setTrashed(true);
      }

      const arquivo = pasta.createFile(blob);
      arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      urlFoto = "https://drive.google.com/thumbnail?id=" + arquivo.getId() + "&sz=w200";
    }

    // Atualizar na planilha
    const abaUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
    const dados = abaUsuarios.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0].toString().trim().toLowerCase() === emailAtivo.toLowerCase()) {
        abaUsuarios.getRange(i + 1, 6).setValue(urlFoto);
        CacheService.getScriptCache().remove("AUTH_V2_" + emailAtivo.toLowerCase());
        SpreadsheetApp.flush();
        return { sucesso: true, mensagem: "Foto atualizada com sucesso!", fotoUrl: urlFoto };
      }
    }

    return { sucesso: false, mensagem: "Usuário não encontrado." };
  } catch (e) {
    return { sucesso: false, mensagem: "Erro ao atualizar foto: " + e.message };
  }
}

/**
 * Atualiza o nome de guerra do usuário logado.
 */
function atualizarNomeUsuario(novoNome) {
  const emailAtivo = Session.getActiveUser().getEmail();
  if (!emailAtivo) return { sucesso: false, mensagem: "E-mail não identificado." };

  const abaUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
  const dados = abaUsuarios.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0].toString().trim().toLowerCase() === emailAtivo.toLowerCase()) {
      abaUsuarios.getRange(i + 1, 2).setValue(novoNome);
      CacheService.getScriptCache().remove("AUTH_V2_" + emailAtivo.toLowerCase());
      SpreadsheetApp.flush();
      return { sucesso: true, mensagem: "Nome atualizado com sucesso!" };
    }
  }
  return { sucesso: false, mensagem: "Usuário não encontrado." };
}

/**
 * Registra a saída do usuário, gravando "Offline" na coluna de Último Login (J)
 */
function registrarSaidaSistema() {
  const sessao = Session.getActiveUser().getEmail();
  if (!sessao) return { sucesso: false };
  
  const abaUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
  if (!abaUsuarios) return { sucesso: false };

  const dados = abaUsuarios.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0].toString().trim().toLowerCase() === sessao.toLowerCase()) {
      abaUsuarios.getRange(i + 1, 12).setValue("Offline");
      return { sucesso: true };
    }
  }
  return { sucesso: false };
}

/**
 * Recebe o sinal de vida do usuário (Heartbeat) baseado em movimento de mouse
 * Grava a data/hora exata atual na aba Usuarios e limpa qualquer status "(Offline)".
 */
function registrarHeartbeat() {
  const sessao = Session.getActiveUser().getEmail();
  if (!sessao) return;
  
  const abaUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
  if (!abaUsuarios) return;

  const dados = abaUsuarios.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0].toString().trim().toLowerCase() === sessao.toLowerCase()) {
      abaUsuarios.getRange(i + 1, 10).setValue(new Date());
      abaUsuarios.getRange(i + 1, 12).setValue("Online");
      return;
    }
  }
}

/**
 * O Ceifador em Background: Executado a cada 1 minuto pelo Google Apps Script.
 * Garante que usuários inativos fiquem offline na planilha mesmo com abas fechadas.
 */
function executarCeifador() {
  const abaUsuarios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
  if (!abaUsuarios) return;

  const ultimoRow = abaUsuarios.getLastRow();
  if (ultimoRow < 1) return;
  const dados = abaUsuarios.getRange(1, 1, ultimoRow, 12).getValues();
  const agoraMs = new Date().getTime();
  const limiteMs = 5 * 60 * 1000; // 5 minutos de inatividade (limite de produção)

  const isDate = function(obj) {
    return obj && (Object.prototype.toString.call(obj) === '[object Date]' || typeof obj.getTime === 'function');
  };

  function parseLocalStringToTime(str) {
      if (!str) return 0;
      let s = str.toString().replace(" (Offline)", "").replace("Offline", "").trim();
      s = s.replace(/\s*\(Hora.*?\)/g, "").trim();
      if (s.includes("/")) {
         let parts = s.split(' ');
         if(parts.length < 2) return 0;
         let d = parts[0].split('/');
         let t = parts[1].split(':');
         if(d.length < 3 || t.length < 3) return 0;
         return Date.UTC(parseInt(d[2], 10), parseInt(d[1], 10) - 1, parseInt(d[0], 10), parseInt(t[0], 10), parseInt(t[1], 10), parseInt(t[2], 10)) + (3 * 60 * 60 * 1000);
      }
      let dt = new Date(s);
      return isNaN(dt.getTime()) ? 0 : dt.getTime();
  }

  for (let i = 1; i < dados.length; i++) {
    let ultimoLoginReal = dados[i][9];
    let statusConexaoReal = dados[i][11] ? dados[i][11].toString().trim() : "Offline";
    
    let isOffline = (statusConexaoReal === "Offline") || (!ultimoLoginReal) || (ultimoLoginReal === "Nunca");
    
    if (!isOffline) {
       let epochUltimo = 0;
       if (isDate(ultimoLoginReal)) {
         epochUltimo = ultimoLoginReal.getTime();
       } else {
         epochUltimo = parseLocalStringToTime(ultimoLoginReal);
       }
       
       if (epochUltimo > 0 && (agoraMs - epochUltimo) > limiteMs) {
           abaUsuarios.getRange(i + 1, 12).setValue("Offline");
       }
    }
  }
}

/**
 * Analisa o texto bruto de um PDF BNMP usando a API do Gemini 3.1 Flash Lite.
 * v4.2.0: Guarda de perfil (somente Admin/Colaborador) + cadência de 4.2s (15 RPM).
 * 
 * @param {string} textoBruto O texto extraído do PDF.
 * @return {Object} O objeto contendo o status do processamento e os dados extraídos.
 */
function analisarTextoComGemini(textoBruto) {
  // ── ISOLAMENTO v4.2.0: Bloquear patrulheiros para preservar cota de 500 RPD ──
  const checagemIA = verificarAcessoUsuario();
  if (!checagemIA.autorizado || checagemIA.perfil === "Patrulheiro") {
    return {
      sucesso: false,
      mensagem: "Acesso negado: Extração por IA é restrita a Admins e Colaboradores."
    };
  }

  Logger.log("[DEBUG GEMINI] Iniciando análise de texto. Tamanho do texto bruto: " + (textoBruto ? textoBruto.length : 0) + " caracteres.");
  
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-flash-lite:generateContent?key=" + apiKey;
  
  const prompt = `Você é um analista jurídico sênior especializado em extração de dados do Banco Nacional de Mandados de Prisão (BNMP).
Sua função é ler o texto bruto extraído de um PDF e converter as informações ESTRITAMENTE para o formato JSON exigido.

DIRETRIZES GERAIS E CRÍTICAS:
1. PREENCHIMENTO OBRIGATÓRIO: Se a informação não existir no texto, use EXATAMENTE "Não informado" ou "N/A". NUNCA deixe vazio.
2. QUEBRAS DE LINHA: O texto do PDF contém quebras de linha incorretas. Você DEVE unir frases cortadas em um texto contínuo e lógico.

MAPEAMENTO DE CAMPOS:
- 'mandado': Capture o número completo no formato (ex: NNNNNNN-NN.NNNN.N.NN.NNNN.NN.NNNN-NN).
- 'titulo' (dentro de extras): Extraia apenas a classificação principal (ex: "MANDADO DE PRISÃO", "MANDADO DE INTERNAÇÃO"), omitindo continuações longas como 'DEFINITIVA DECORRENTE DE CONDENAÇÃO'.
- 'infoProcessuais': Formate a string EXATAMENTE com este molde, separando por " | ":
  "Processo: [Nº] | Órgão Judicial: [Órgão] | Espécie de prisão: [Espécie] | Lei: [Lei] | Artigo: [Artigo EXATAMENTE COMO NO PDF INCLUINDO LETRAS ex: 217-A, 217A] | Parágrafo: [Parágrafo ou N/A] | Inciso: [Inciso ou N/A] | Pena restante: [Pena] | Regime Prisional: [Regime] | Motivo: [Resumo da Síntese]"

TEXTO DO DOCUMENTO A SER ANALISADO:
---INICIO DO DOCUMENTO---
${textoBruto}
---FIM DO DOCUMENTO---`;

  const responseSchema = {
    "type": "OBJECT",
    "properties": {
      "nome": { "type": "STRING" },
      "cpf": { "type": "STRING" },
      "mandado": { "type": "STRING" },
      "validade": { "type": "STRING" },
      "rg": { "type": "STRING" },
      "naturalidade": { "type": "STRING" },
      "nascimento": { "type": "STRING" },
      "sexo": { "type": "STRING" },
      "cor": { "type": "STRING" },
      "filiacao": { "type": "STRING" },
      "batalhao": { "type": "STRING" },
      "infoProcessuais": { "type": "STRING", "description": "Resumo unindo Processo, Órgão, Espécie, Lei, Artigo, Motivo" },
      "artigo": { "type": "STRING", "description": "Mantenha as letras exatas do PDF, ex: 217-A" },
      "lei": { "type": "STRING" },
      "enderecoPrincipal": { "type": "STRING" },
      "latPrincipal": { "type": "NUMBER" },
      "lngPrincipal": { "type": "NUMBER" },
      "secundarios": {
        "type": "ARRAY",
        "items": {
          "type": "OBJECT",
          "properties": {
            "endereco": { "type": "STRING" },
            "lat": { "type": "NUMBER" },
            "lng": { "type": "NUMBER" }
          },
          "required": ["endereco"]
        }
      },
      "extras": {
        "type": "OBJECT",
        "properties": {
          "titulo": { "type": "STRING" },
          "emissao": { "type": "STRING" },
          "nomeSocial": { "type": "STRING" },
          "rji": { "type": "STRING" },
          "alcunha": { "type": "STRING" },
          "marcas": { "type": "STRING" },
          "sintese": { "type": "STRING" },
          "obsPdf": { "type": "STRING" }
        },
        "required": ["titulo", "emissao", "nomeSocial", "rji", "alcunha", "marcas", "sintese", "obsPdf"]
      }
    },
    "required": [
      "nome", "cpf", "mandado", "validade", "rg", "naturalidade", "nascimento", "sexo", 
      "cor", "filiacao", "batalhao", "infoProcessuais", "artigo", "lei", "enderecoPrincipal", 
      "secundarios", "extras"
    ]
  };

  const payload = {
    "contents": [
      {
        "parts": [
          {
            "text": prompt
          }
        ]
      }
    ],
    "generationConfig": {
      "responseMimeType": "application/json",
      "responseSchema": responseSchema
    }
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    Logger.log("[DEBUG GEMINI] Enviando requisição para a API do Gemini...");
    
    // ── RETRY AUTOMÁTICO com backoff para erros 503/429 ──
    var MAX_TENTATIVAS = 3;
    var response = null;
    var code = 0;
    var content = "";
    
    for (var tentativa = 1; tentativa <= MAX_TENTATIVAS; tentativa++) {
      response = UrlFetchApp.fetch(url, options);
      code = response.getResponseCode();
      content = response.getContentText();
      
      Logger.log("[DEBUG GEMINI] Tentativa " + tentativa + "/" + MAX_TENTATIVAS + " - Status HTTP: " + code);
      
      if (code === 200) {
        break; // Sucesso!
      }
      
      // Erros temporários (503 = sobrecarga, 429 = rate limit)
      if ((code === 503 || code === 429) && tentativa < MAX_TENTATIVAS) {
        var espera = 5000 * Math.pow(2, tentativa - 1); // 5s, 10s, 20s
        Logger.log("[DEBUG GEMINI] Erro temporário " + code + ". Aguardando " + (espera/1000) + "s antes de tentar novamente...");
        Utilities.sleep(espera);
        continue;
      }
      
      // Erro definitivo (não é 503 nem 429, ou esgotou tentativas)
      throw new Error("Erro na API do Gemini (Status " + code + "): " + content);
    }
    
    // Se chegou aqui após todas as tentativas e ainda não é 200
    if (code !== 200) {
      throw new Error("Erro na API do Gemini após " + MAX_TENTATIVAS + " tentativas (Status " + code + "): " + content);
    }
    
    Logger.log("[DEBUG GEMINI] Resposta bruta da API: " + content);
    
    const parsedResponse = JSON.parse(content);
    
    if (!parsedResponse.candidates || parsedResponse.candidates.length === 0 || 
        !parsedResponse.candidates[0].content || !parsedResponse.candidates[0].content.parts || 
        parsedResponse.candidates[0].content.parts.length === 0) {
      throw new Error("Resposta vazia ou inválida da API do Gemini.");
    }
    
    const jsonString = parsedResponse.candidates[0].content.parts[0].text;
    const extraido = JSON.parse(jsonString);

    // Defesa de dados e Pós-processamento
    extraido.batalhao = "A DEFINIR (GEO)";
    extraido.latPrincipal = null;
    extraido.lngPrincipal = null;

    if (!Array.isArray(extraido.secundarios)) {
      extraido.secundarios = [];
    }
    
    extraido.secundarios = extraido.secundarios.map(function(item) {
      return {
        endereco: (item && item.endereco) ? item.endereco.trim() : "Não informado",
        lat: null,
        lng: null
      };
    });
    
    if (!extraido.extras) {
      extraido.extras = {
        titulo: "Não informado",
        emissao: "Não informado",
        nomeSocial: "Não informado",
        rji: "Não informado",
        alcunha: "Não informado",
        marcas: "Não informado",
        sintese: "Não informado",
        obsPdf: "Não informado"
      };
    }

    Logger.log("[DEBUG GEMINI] Dados estruturados extraídos com sucesso: " + JSON.stringify(extraido));
    extraido.tipoImportacao = "GEMINI";

    // ── CADÊNCIA v4.2.0: 4.2s de delay para respeitar teto de 15 RPM ──
    Utilities.sleep(4200);

    // ── COTA v4.2.0: Registrar chamada bem-sucedida no contador diário ──
    incrementarCotaGemini_("EXTRACAO");

    return {
      sucesso: true,
      dados: extraido
    };
  } catch (e) {
    Logger.log("[DEBUG GEMINI] ERRO na análise com Gemini: " + e.message);
    return {
      sucesso: false,
      mensagem: "Erro ao analisar documento com IA: " + e.message
    };
  }
}

// ================================================================
// MONITORAMENTO DE COTA GEMINI (v4.2.0)
// Rastreia chamadas diárias ao Gemini 3.1 Flash Lite.
// Limites Free Tier: 500 RPD / 15 RPM.
// ================================================================

/**
 * Incrementa o contador diário de chamadas Gemini.
 * @param {string} tipo "EXTRACAO" ou "GEOCODIFICACAO"
 * @private
 */
function incrementarCotaGemini_(tipo) {
  try {
    const props = PropertiesService.getScriptProperties();
    const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const chaveTotal = "GEMINI_COTA_" + hoje;
    const chaveTipo = "GEMINI_COTA_" + tipo + "_" + hoje;

    // Incrementar total
    const totalAtual = parseInt(props.getProperty(chaveTotal)) || 0;
    props.setProperty(chaveTotal, String(totalAtual + 1));

    // Incrementar por tipo
    const tipoAtual = parseInt(props.getProperty(chaveTipo)) || 0;
    props.setProperty(chaveTipo, String(tipoAtual + 1));

    // Limpar contadores de dias anteriores (manter últimos 7 dias)
    if (totalAtual === 0) {
      limparCotasAntigas_();
    }
  } catch (e) {
    Logger.log("[COTA GEMINI] Erro ao incrementar contador: " + e.message);
  }
}

/**
 * Remove contadores de cota com mais de 7 dias.
 * @private
 */
function limparCotasAntigas_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const todas = props.getProperties();
    const agora = new Date();

    for (var chave in todas) {
      if (chave.indexOf("GEMINI_COTA_") === 0) {
        // Extrair data da chave (formato: GEMINI_COTA_yyyy-MM-dd ou GEMINI_COTA_TIPO_yyyy-MM-dd)
        var partes = chave.split("_");
        var dataStr = partes[partes.length - 1];
        if (/^\d{4}-\d{2}-\d{2}$/.test(dataStr)) {
          var dataCota = new Date(dataStr + "T00:00:00");
          var diffDias = (agora.getTime() - dataCota.getTime()) / (1000 * 60 * 60 * 24);
          if (diffDias > 7) {
            props.deleteProperty(chave);
          }
        }
      }
    }
  } catch (e) {
    Logger.log("[COTA GEMINI] Erro ao limpar cotas antigas: " + e.message);
  }
}

/**
 * Retorna estatísticas de uso da cota Gemini para o frontend.
 * Chamada pelo dashboard de Inteligência Operacional.
 * 
 * @return {Object} { hoje: { total, extracao, geocodificacao, limite, percentual, status },
 *                     historico: [{ data, total, extracao, geocodificacao }] }
 */
function obterCotaGemini() {
  const checagem = verificarAcessoUsuario();
  if (!checagem.autorizado || checagem.perfil === "Patrulheiro") {
    return { erro: "Acesso restrito a Admins e Colaboradores." };
  }

  const props = PropertiesService.getScriptProperties();
  const hoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const LIMITE_DIARIO = 500;

  // Cota de hoje
  const totalHoje = parseInt(props.getProperty("GEMINI_COTA_" + hoje)) || 0;
  const extracaoHoje = parseInt(props.getProperty("GEMINI_COTA_EXTRACAO_" + hoje)) || 0;
  const geoHoje = parseInt(props.getProperty("GEMINI_COTA_GEOCODIFICACAO_" + hoje)) || 0;
  const percentual = Math.round((totalHoje / LIMITE_DIARIO) * 100);

  var status = "normal";
  if (percentual >= 90) status = "critico";
  else if (percentual >= 70) status = "alerta";

  // Histórico dos últimos 7 dias
  var historico = [];
  for (var d = 6; d >= 0; d--) {
    var data = new Date();
    data.setDate(data.getDate() - d);
    var dataStr = Utilities.formatDate(data, Session.getScriptTimeZone(), "yyyy-MM-dd");
    var diaLabel = Utilities.formatDate(data, Session.getScriptTimeZone(), "dd/MM");

    historico.push({
      data: diaLabel,
      total: parseInt(props.getProperty("GEMINI_COTA_" + dataStr)) || 0,
      extracao: parseInt(props.getProperty("GEMINI_COTA_EXTRACAO_" + dataStr)) || 0,
      geocodificacao: parseInt(props.getProperty("GEMINI_COTA_GEOCODIFICACAO_" + dataStr)) || 0
    });
  }

  return {
    hoje: {
      total: totalHoje,
      extracao: extracaoHoje,
      geocodificacao: geoHoje,
      limite: LIMITE_DIARIO,
      percentual: percentual,
      status: status
    },
    historico: historico
  };
}
