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
      const perfil = dados[i][3];
      const status = dados[i][4].toString().trim().toUpperCase();
      const fotoUrl = dados[i][5] ? dados[i][5].toString().trim() : "";
      const pontos = parseInt(dados[i][6]) || 0;
      const badges = dados[i][7] ? dados[i][7].toString().trim() : "";

      if (status === "ATIVO") {
        return {
          autorizado: true,
          perfil: perfil,
          nome: nome,
          unidade: unidade,
          email: emailAtivo,
          fotoUrl: fotoUrl,
          pontos: pontos,
          badges: badges
        };
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

    abaUsuarios.appendRow([
      emailAtivo,         // A - Email
      nomeGuerra,         // B - Nome
      unidade,            // C - Unidade
      "Patrulheiro",      // D - Perfil (padrão)
      "Pendente",         // E - Status (padrão)
      "",                 // F - Foto URL
      "0",                // G - Pontos
      ""                  // H - Badges
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

  const dados = abaUsuarios.getDataRange().getValues();
  const todos = [];

  for (let i = 1; i < dados.length; i++) {
    todos.push({
      email: dados[i][0],
      nome: dados[i][1],
      unidade: dados[i][2],
      perfil: dados[i][3],
      status: dados[i][4].toString().trim().toUpperCase()
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
        cont++;
      }
    }
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
        abaUsuarios.getRange(i + 1, 5).setValue(novoStatus);
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
 * Atualiza a foto do perfil do usuário logado.
 * Salva a foto na pasta Sentinela_Fotos_Usuarios.
 */
function atualizarFotoUsuario(fotoBase64) {
  const emailAtivo = Session.getActiveUser().getEmail();
  if (!emailAtivo) return { sucesso: false, mensagem: "E-mail não identificado." };

  try {
    let urlFoto = "";

    if (fotoBase64 && fotoBase64.startsWith("data:image")) {
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
      return { sucesso: true, mensagem: "Nome atualizado com sucesso!" };
    }
  }
  return { sucesso: false, mensagem: "Usuário não encontrado." };
}
