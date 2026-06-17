<script>
/* ================================================================
     MÓDULO 1: VARIÁVEIS GLOBAIS E ESTADO
     ================================================================ */
  let mapaPrincipal;
  let dadosMandados = [];
  let listaMarcadores = [];
  let infoWindowAtiva = null;
  let usuarioLogado = null;
  let perfilUsuario = '';
  let dadosTemporariosCadastro = null;
  let pinosPreview = [];
  let fotoBase64Temporaria = null;
  let posicaoPatrulheiro = null;
  let watchPositionId = null;
  let configsSistema = {};
  let secundariosVisiveis = {}; // { mandado: { markers: [], polylines: [] } }
  let lastPendingCount = 0; // Para notificar novos cadastros

  /* ================================================================
     MÓDULO 2: AUTENTICAÇÃO E INICIALIZAÇÃO
     ================================================================ */
  function iniciarMapa() {
    const lblStatus = document.getElementById("txt-status-carregamento");
    if (lblStatus) lblStatus.innerText = "VERIFICANDO CREDENCIAIS...";

    // Carregar configs primeiro, depois verificar acesso
    google.script.run.withSuccessHandler(function(configs) {
      configsSistema = configs;
      aplicarConfigsDinamicas(configs);
      if (lblStatus) lblStatus.innerText = "CREDENCIAIS VERIFICADAS. ENTRANDO NO SISTEMA...";
      google.script.run.withSuccessHandler(avaliarAcessoSistema).verificarAcessoUsuario();
    }).obterConfiguracoesSimples();
  }

  function avaliarAcessoSistema(resposta) {
    const boxCarregando = document.getElementById("seg-estado-carregando");
    const telaSeguranca = document.getElementById("tela-seguranca");
    const lblStatus = document.getElementById("txt-status-carregamento");

    if (resposta.autorizado) {
      if (lblStatus) lblStatus.innerText = "CARREGANDO OS MANDADOS E O MAPA...";
      usuarioLogado = resposta;
      perfilUsuario = resposta.perfil || '';
      // Nota: telaSeguranca será escondida apenas após o carregamento dos mandados no JS_Map.html

      // Atualizar sidebar
      document.getElementById('sidebarNomeEl').textContent = resposta.nome;
      document.getElementById('sidebarPerfilEl').textContent = resposta.perfil;
      if (resposta.fotoUrl) {
        document.getElementById('sidebarAvatarContainer').innerHTML =
          '<img src="' + resposta.fotoUrl + '" style="width:100%;height:100%;object-fit:cover;border-radius:50%;" />';
      }

      if (resposta.perfil === "Admin") {
        const btnNovo = document.getElementById("btnAdminNovo");
        const btnMenu = document.getElementById("btnMenuAdmin");
        const btnAjustes = document.getElementById("btnMenuAjustes");
        const btnConferencia = document.getElementById("btnMenuConferencia");
        if (btnNovo) btnNovo.style.display = "flex";
        if (btnMenu) btnMenu.style.display = "block";
        if (btnAjustes) btnAjustes.style.display = "block";
        if (btnConferencia) btnConferencia.style.display = "block";
        
        // Polling para novos cadastros (a cada 2 min)
        setInterval(function() {
          google.script.run.withSuccessHandler(function(usuarios) {
            if (!usuarios) return;
            const pendentesAtual = usuarios.filter(function(u) { return u.status === 'PENDENTE'; }).length;
            if (pendentesAtual > lastPendingCount) {
              const diferenca = pendentesAtual - lastPendingCount;
              chamarAlertaTatico("NOVO CADASTRO", diferenca + " novo(s) operador(es) pendente(s) aguardando aprovação.", "OK", function() {
                mostrarView('view-usuarios');
                carregarListaUsuarios();
              });
            }
            lastPendingCount = pendentesAtual;
            const badge = document.getElementById('badgeUsuariosPendentes');
            if (badge) {
              if (pendentesAtual > 0) { badge.textContent = pendentesAtual; badge.style.display = 'inline-block'; }
              else { badge.style.display = 'none'; }
            }
          }).obterTodosUsuarios();
        }, 120000);
      }
      
      if (typeof mapaPrincipal === 'undefined' || !mapaPrincipal) {
        configurarEstruturaMapa();
      } else {
        // Mapa já existe, apenas sincronizar os dados
        const iconeSinc = document.querySelector('button[title="Sincronizar"] span');
        if (iconeSinc) iconeSinc.style.animation = 'spin 1s linear infinite';
        google.script.run.withSuccessHandler(function(dados) {
          processarDados(dados);
          if (iconeSinc) iconeSinc.style.animation = 'none';
        }).obterDados();
      }
    }
    else if (resposta.status === "NAO_CADASTRADO") {
      if (boxCarregando) boxCarregando.style.display = "none";
      document.getElementById("lbl-email-google").innerText = resposta.email;
      document.getElementById("seg-estado-cadastro").style.display = "block";
    }
    else {
      if (boxCarregando) boxCarregando.style.display = "none";
      document.getElementById("msg-seg-titulo").innerText = resposta.status === "PENDENTE" ? "⏳ AGUARDANDO APROVAÇÃO" : "❌ ACESSO BLOQUEADO";
      document.getElementById("msg-seg-texto").innerText = resposta.motivo;
      document.getElementById("seg-estado-mensagem").style.display = "block";
    }
  }

  function reiniciarVerificacaoSeguranca() {
    document.getElementById("seg-estado-mensagem").style.display = "none";
    document.getElementById("seg-estado-cadastro").style.display = "none";
    document.getElementById("seg-estado-carregando").style.display = "block";
    document.getElementById("tela-seguranca").style.display = "flex";
    google.script.run.withSuccessHandler(avaliarAcessoSistema)
      .withFailureHandler(function(erro) {
        chamarAlertaTatico("ERRO", "Não foi possível contactar o servidor: " + erro.message, "OK");
        document.getElementById("seg-estado-carregando").style.display = "none";
        document.getElementById("seg-estado-mensagem").style.display = "block";
      }).verificarAcessoUsuario();
  }

  /* ================================================================
     MÓDULO 7: OPERAÇÕES DE REDE E CADASTRO
     ================================================================ */
  function chamarAlertaTatico(titulo, mensagem, tipo, callback) {
    const modal = document.getElementById("modalAlertaTatico");
    document.getElementById("alerta-titulo").innerText = titulo;
    document.getElementById("alerta-mensagem").innerText = mensagem;
    const containerBotoes = document.getElementById("alerta-botoes");
    containerBotoes.innerHTML = "";
    const alertBox = modal.querySelector(".alert-modal-box");

    if (tipo === "CONFIRMAR") {
      if (alertBox) alertBox.style.borderTopColor = "var(--warning)";
      const btnSim = document.createElement("button"); btnSim.className = "btn-alerta-sim";
      btnSim.innerText = "SIM, CONFIRMAR";
      btnSim.onclick = function() { modal.style.display = "none"; if(callback) callback(true); };
      const btnNao = document.createElement("button"); btnNao.className = "btn-alerta-nao";
      btnNao.innerText = "NÃO, CANCELAR";
      btnNao.onclick = function() { modal.style.display = "none"; if(callback) callback(false); };
      containerBotoes.appendChild(btnSim); containerBotoes.appendChild(btnNao);
    } else {
      if (alertBox) alertBox.style.borderTopColor = "var(--accent)";
      const btnOk = document.createElement("button"); btnOk.className = "btn-alerta-ok";
      btnOk.innerText = "CIENTE";
      btnOk.onclick = function() { modal.style.display = "none"; if(callback) callback(); };
      containerBotoes.appendChild(btnOk);
    }
    modal.style.display = "flex";
  }

  function enviarAtualizacao(mandado, btnElement) {
    btnElement.innerText = "Sincronizando..."; btnElement.disabled = true;
    const radios = document.getElementsByName("status_" + mandado);
    let novoStatus = "Procurado";
    for(let i=0;i<radios.length;i++) { if(radios[i].checked) { novoStatus = radios[i].value; break; } }
    const novaObs = document.getElementById("obs_" + mandado).value;

    google.script.run.withSuccessHandler(function(res) {
      chamarAlertaTatico("ATUALIZAÇÃO", res.mensagem, "OK");
      infoWindowAtiva.close();
      const alvoLocal = dadosMandados.find(function(d) { return d.mandado === mandado; });
      if(alvoLocal) { alvoLocal.status = novoStatus; alvoLocal.observacoes = novaObs; }
      filtrarNoMapa();
    }).atualizarRegistro(mandado, novoStatus, novaObs);
  }

  function conferirMandadoClick(mandado, btn) {
    btn.innerText = "⏳ Conferindo..."; btn.disabled = true;
    google.script.run.withSuccessHandler(function(res) {
      chamarAlertaTatico("CONFERÊNCIA", res.mensagem, "OK");
      btn.innerText = "✅ CONFERIDO"; btn.disabled = true;
      const alvo = dadosMandados.find(function(d) { return d.mandado === mandado; });
      if (alvo) {
        const hoje = new Date();
        alvo.dataConferencia = hoje.getDate().toString().padStart(2,'0') + '/' + (hoje.getMonth()+1).toString().padStart(2,'0') + '/' + hoje.getFullYear();
      }
    }).conferirMandado(mandado);
  }

  function filtrarNoMapa() {
    const termoNome = document.getElementById('filtroNome').value.toLowerCase();
    const filtroCPI = document.getElementById('filtroCPI').value;
    const filtroBPM = document.getElementById('filtroBatalhao').value;
    const filtroCIA = document.getElementById('filtroCIA').value;
    
    // Filtros avançados (drawer)
    const elCrime = document.getElementById('filtroCrime');
    const filtroCrime = elCrime ? elCrime.value : 'TODOS';
    const elSexo = document.getElementById('filtroSexo');
    const filtroSexo = elSexo ? elSexo.value : 'TODOS';
    const elTipoMandado = document.getElementById('filtroTipoMandado');
    const filtroTipo = elTipoMandado ? elTipoMandado.value : 'TODOS';
    
    const filtrados = dadosMandados.filter(function(alvo) {
      const nomeAlvo = (alvo.nome || "").toLowerCase();
      const mandadoAlvo = (alvo.mandado || "").toLowerCase();
      const matchNome = (nomeAlvo.includes(termoNome) || mandadoAlvo.includes(termoNome));
      
      // Filtros de Área (Principal ou Secundários)
      let matchPrincipal = false;
      let okCPI = (filtroCPI === "TODOS" || alvo.cpiArea === filtroCPI);
      let okBPM = (filtroBPM === "TODOS" || alvo.bpmArea === filtroBPM || alvo.batalhao === filtroBPM);
      let okCIA = (filtroCIA === "TODOS" || alvo.ciaArea === filtroCIA);
      if (okCPI && okBPM && okCIA) matchPrincipal = true;

      let matchSecundario = false;
      if (filtroCPI !== "TODOS" || filtroBPM !== "TODOS" || filtroCIA !== "TODOS") {
        try {
          const listaSec = JSON.parse(alvo.geodataSecundarios || "[]");
          for (let i = 0; i < listaSec.length; i++) {
            let sec = listaSec[i];
            let secOkCPI = (filtroCPI === "TODOS" || sec.cpi === filtroCPI);
            let secOkBPM = (filtroBPM === "TODOS" || sec.batalhao === filtroBPM);
            let secOkCIA = (filtroCIA === "TODOS" || sec.cia === filtroCIA);
            if (secOkCPI && secOkBPM && secOkCIA) {
              matchSecundario = true;
              break;
            }
          }
        } catch(e) {}
      }

      const matchArea = (matchPrincipal || matchSecundario);
      
      // Filtro Crime
      var matchCrime = true;
      if (filtroCrime !== "TODOS") {
        var crimeAlvo = classificarCrime(alvo.infoProcessuais || '');
        matchCrime = (crimeAlvo.nome === filtroCrime);
      }

      // Filtro Sexo
      var matchSexo = true;
      if (filtroSexo !== "TODOS") {
        matchSexo = ((alvo.sexo || "").toUpperCase() === filtroSexo);
      }

      // Filtro Tipo Mandado
      var matchTipo = true;
      if (filtroTipo !== "TODOS") {
        var tituloMandado = "";
        try {
          var extras = alvo.dadosExtrasJSON || "{}";
          var parsed = JSON.parse(extras);
          tituloMandado = (parsed.titulo || "").toUpperCase();
        } catch(e) {}
        matchTipo = (tituloMandado === filtroTipo.toUpperCase());
      }
      
      return matchNome && matchArea && matchCrime && matchSexo && matchTipo;
    });
    
    renderizarPinos(filtrados);
    if (typeof renderizarTabelaMandados === 'function') renderizarTabelaMandados(filtrados);
    if (typeof aplicarFiltrosDePoligono === 'function') aplicarFiltrosDePoligono(filtroCPI, filtroBPM, filtroCIA);
  }

  function abrirGpsNativo(lat, lng) {
    window.open('https://www.google.com/maps/dir/?api=1&destination=' + lat + ',' + lng + '&travelmode=driving', '_blank');
  }

  function abrirStreetView(lat, lng) {
    if (typeof lat !== 'number' || typeof lng !== 'number' || isNaN(lat) || isNaN(lng)) return;
    
    // Fechar o card/InfoWindow antes de abrir o Street View
    if (typeof infoWindowAtiva !== 'undefined' && infoWindowAtiva) infoWindowAtiva.close();

    // Garantir que o mapa está visível
    var mapDiv = document.getElementById('mapaPrincipal');
    if (mapDiv) mapDiv.style.display = 'block';

    var panorama = mapaPrincipal.getStreetView();
    panorama.setPosition({ lat: lat, lng: lng });
    panorama.setPov({ heading: 0, pitch: 0 });
    panorama.setVisible(true);
  }

  function abrirModalSuperBloco() { document.getElementById('modalSuperBloco').style.display = 'flex'; }

  function fecharModalSuperBloco() {
    document.getElementById('modalSuperBloco').style.display = 'none';
    document.getElementById('painelPreview').style.display = 'none';
    document.getElementById('textoPdfBruto').value = '';
    fotoBase64Temporaria = null;
    const inputFoto = document.getElementById('input-foto-admin'); if (inputFoto) inputFoto.value = "";
    const telaRecorte = document.getElementById('tela-de-recorte'); if(telaRecorte) telaRecorte.style.display = 'none';
    const wrapperFicha = document.getElementById('wrapper-ficha-normal'); if(wrapperFicha) wrapperFicha.style.display = 'flex';
    const btnProcessar = document.getElementById('btnProcessarTexto'); if(btnProcessar) btnProcessar.style.display = 'block';
    const textoUpload = document.getElementById('foto-texto-upload'); if (textoUpload) textoUpload.style.display = 'flex';
    const preview = document.getElementById('preview-foto-admin'); if(preview) { preview.style.display = 'none'; preview.className = ""; }
  }

  function copiarDadoTatico(texto, btnElement) {
    if (!texto || texto === "N/A") return;
    navigator.clipboard.writeText(texto).then(function() {
      btnElement.innerHTML = '<span class="material-symbols-outlined" style="color:var(--success)!important;font-size:14px;">check_circle</span>';
      setTimeout(function() { btnElement.innerHTML = '<span class="material-symbols-outlined" style="font-size:14px;">content_copy</span>'; }, 1500);
    }).catch(function() {
      const t = document.createElement("textarea"); t.value = texto;
      document.body.appendChild(t); t.select(); document.execCommand("copy"); document.body.removeChild(t);
      btnElement.innerHTML = '<span class="material-symbols-outlined" style="color:var(--success)!important;font-size:14px;">check_circle</span>';
      setTimeout(function() { btnElement.innerHTML = '<span class="material-symbols-outlined" style="font-size:14px;">content_copy</span>'; }, 1500);
    });
  }

  function geocodeEndereco(geocoder, endereco) {
    return new Promise(function(resolve) {
      geocoder.geocode({ 'address': endereco }, function(results, status) {
        if (status === 'OK') {
          var uf = "";
          var comps = results[0].address_components;
          for (var i = 0; i < comps.length; i++) {
            if (comps[i].types.includes("administrative_area_level_1")) {
              uf = comps[i].short_name;
              break;
            }
          }
          resolve({ 
            lat: results[0].geometry.location.lat(), 
            lng: results[0].geometry.location.lng(),
            uf: uf
          });
        }
        else resolve(null);
      });
    });
  }

  /* ================================================================
     MÓDULO 9: CADASTRO DE OPERADOR + WHATSAPP OPCIONAL
     ================================================================ */
  function executarCadastroTatico() {
    const nomeGuerra = document.getElementById("txt-nome-guerra").value.trim();
    const batalhao = document.getElementById("txt-batalhao-cad").value.trim();
    const email = document.getElementById("lbl-email-google").innerText;
    if (!nomeGuerra || !batalhao) return chamarAlertaTatico("DADOS INCOMPLETOS", "Preencha o Nome de Guerra e a Unidade.", "OK");

    document.getElementById("seg-estado-cadastro").style.display = "none";
    document.getElementById("seg-estado-carregando").style.display = "block";

    google.script.run.withSuccessHandler(function(resultado) {
      if (resultado.sucesso) {
        // WhatsApp opcional
        const whatsAtivo = configsSistema.whatsapp_ativo === 'true';
        const numero = configsSistema.whatsapp_numero || '5519992693763';

        if (whatsAtivo) {
          chamarAlertaTatico("CADASTRO ENVIADO", "Deseja notificar o administrador via WhatsApp para agilizar a liberação?", "CONFIRMAR", function(confirma) {
            if (confirma) {
              const textoMsg = "Solicito liberação de acesso ao Sentinela.\n\n• Operador: " + nomeGuerra + "\n• Unidade: " + batalhao + "\n• Login: " + email;
              window.open("https://api.whatsapp.com/send?phone=" + numero + "&text=" + encodeURIComponent(textoMsg), '_blank');
            }
            mostrarMensagemPendente();
          });
        } else {
          mostrarMensagemPendente();
        }
      } else {
        chamarAlertaTatico("FALHA", resultado.mensagem, "OK");
        document.getElementById("seg-estado-carregando").style.display = "none";
        document.getElementById("seg-estado-cadastro").style.display = "block";
      }
    }).registrarNovoUsuario(nomeGuerra, batalhao);
  }

  function mostrarMensagemPendente() {
    document.getElementById("seg-estado-carregando").style.display = "none";
    document.getElementById("msg-seg-titulo").innerText = "⏳ SOLICITAÇÃO ENVIADA";
    document.getElementById("msg-seg-texto").innerText = "Seu e-mail foi cadastrado como PENDENTE. Aguarde a validação do administrador.";
    document.getElementById("seg-estado-mensagem").style.display = "block";
  }

  /* ================================================================
     MÓDULO 11: GESTÃO DE USUÁRIOS (Admin) — v2.5
     ================================================================ */
  function carregarListaUsuarios() { carregarUsuariosGrade(); }

  function carregarUsuariosGrade() {
    const pendentesContainer = document.getElementById('containerUsuariosPendentes');
    const ativosContainer = document.getElementById('containerUsuariosAtivos');
    const bloqueadosContainer = document.getElementById('containerUsuariosBloqueados');
    if (pendentesContainer) pendentesContainer.innerHTML = '<p style="color:var(--text-muted);font-size:12px;">Carregando...</p>';
    if (ativosContainer) ativosContainer.innerHTML = '<p style="color:var(--text-muted);font-size:12px;">Carregando...</p>';
    if (bloqueadosContainer) bloqueadosContainer.innerHTML = '';

    google.script.run.withSuccessHandler(function(usuarios) {
      if (!usuarios) return;
      
      const pendentes = usuarios.filter(function(u) { return u.status === 'PENDENTE'; });
      const ativos = usuarios.filter(function(u) { return u.status === 'ATIVO'; });
      const bloqueados = usuarios.filter(function(u) { return u.status === 'BLOQUEADO' || u.status === 'INATIVO'; });

      // Stats
      var elTotal = document.getElementById('statUsuTotal');
      var elAtivos = document.getElementById('statUsuAtivos');
      var elPendentes = document.getElementById('statUsuPendentes');
      var elBloqueados = document.getElementById('statUsuBloqueados');
      if (elTotal) elTotal.textContent = usuarios.length;
      if (elAtivos) elAtivos.textContent = ativos.length;
      if (elPendentes) elPendentes.textContent = pendentes.length;
      if (elBloqueados) elBloqueados.textContent = bloqueados.length;

      // Badge na sidebar
      var badge = document.getElementById('badgeUsuariosPendentes');
      if (badge) {
        if (pendentes.length > 0) { badge.textContent = pendentes.length; badge.style.display = 'inline-block'; }
        else { badge.style.display = 'none'; }
      }

      // Renderizar Pendentes
      if (pendentes.length === 0) {
        pendentesContainer.innerHTML = '<div style="padding:16px;background:var(--bg-deep);border-radius:8px;color:var(--text-muted);font-size:12px;text-align:center;">Nenhuma solicitação pendente.</div>';
      } else {
        pendentesContainer.innerHTML = pendentes.map(function(op) {
          var emailSeg = op.email.replace(/'/g, "\\'");
          return '<div class="card-operador" style="background:var(--bg-deep);padding:14px;border-radius:10px;box-shadow:var(--neu-raised-sm);border-left:3px solid var(--warning);">' +
            '<div style="display:flex;justify-content:space-between;align-items:flex-start;">' +
              '<div>' +
                '<div style="font-weight:800;font-size:14px;color:var(--text-primary);">' + op.nome + '</div>' +
                '<div style="font-size:11px;color:var(--text-secondary);margin-top:2px;">' + op.unidade + '</div>' +
                '<div style="font-size:10px;color:var(--text-muted);margin-top:2px;font-family:var(--font-mono);">' + op.email + '</div>' +
              '</div>' +
              '<span style="font-size:9px;font-weight:800;background:var(--warning)20;color:var(--warning);padding:3px 8px;border-radius:99px;">PENDENTE</span>' +
            '</div>' +
            '<div style="display:flex;gap:8px;margin-top:12px;">' +
              '<button class="neu-btn neu-btn-accent" style="flex:1;padding:10px;font-size:11px;" onclick="processarDecisaoAdmin(\'' + emailSeg + '\', \'ATIVO\', this)"><span class="material-symbols-outlined" style="font-size:14px;">check</span> APROVAR</button>' +
              '<button class="neu-btn" style="flex:1;padding:10px;font-size:11px;color:var(--danger);" onclick="processarDecisaoAdmin(\'' + emailSeg + '\', \'BLOQUEADO\', this)"><span class="material-symbols-outlined" style="font-size:14px;">block</span> REJEITAR</button>' +
            '</div></div>';
        }).join('');
      }

      // Renderizar Ativos
      if (ativos.length === 0) {
        ativosContainer.innerHTML = '<div style="padding:16px;background:var(--bg-deep);border-radius:8px;color:var(--text-muted);font-size:12px;text-align:center;">Nenhum operador ativo.</div>';
      } else {
        ativosContainer.innerHTML = ativos.map(function(op) {
          var emailSeg = op.email.replace(/'/g, "\\'");
          return '<div class="card-operador" style="background:var(--bg-deep);padding:14px;border-radius:10px;box-shadow:var(--neu-pressed-sm);border-left:3px solid var(--success);">' +
            '<div style="display:flex;justify-content:space-between;align-items:center;">' +
              '<div style="flex:1;min-width:0;">' +
                '<div style="font-weight:700;font-size:13px;color:var(--text-primary);display:flex;align-items:center;gap:6px;">' + op.nome +
                  ' <span style="font-size:9px;background:var(--accent)15;color:var(--accent);padding:2px 6px;border-radius:4px;font-weight:800;">' + (op.perfil || 'Operador') + '</span>' +
                '</div>' +
                '<div style="font-size:11px;color:var(--text-secondary);margin-top:2px;">' + op.unidade + '</div>' +
                '<div style="font-size:10px;color:var(--text-muted);margin-top:2px;font-family:var(--font-mono);overflow:hidden;text-overflow:ellipsis;">' + op.email + '</div>' +
              '</div>' +
              '<button class="neu-btn" style="padding:8px;color:var(--danger);flex-shrink:0;" onclick="processarDecisaoAdmin(\'' + emailSeg + '\', \'BLOQUEADO\', this)" title="Bloquear"><span class="material-symbols-outlined" style="font-size:16px;">block</span></button>' +
            '</div>' +
          '</div>';
        }).join('');
      }

      // Renderizar Bloqueados
      if (bloqueadosContainer) {
        if (bloqueados.length === 0) {
          bloqueadosContainer.innerHTML = '<div style="padding:16px;background:var(--bg-deep);border-radius:8px;color:var(--text-muted);font-size:12px;text-align:center;">Nenhum operador bloqueado.</div>';
        } else {
          bloqueadosContainer.innerHTML = bloqueados.map(function(op) {
            var emailSeg = op.email.replace(/'/g, "\\'");
            return '<div class="card-operador" style="background:var(--bg-deep);padding:14px;border-radius:10px;border-left:3px solid var(--danger);opacity:0.7;">' +
              '<div style="display:flex;justify-content:space-between;align-items:center;">' +
                '<div style="flex:1;min-width:0;">' +
                  '<div style="font-weight:700;font-size:13px;color:var(--text-primary);">' + op.nome + '</div>' +
                  '<div style="font-size:10px;color:var(--text-muted);margin-top:2px;">' + op.email + '</div>' +
                '</div>' +
                '<button class="neu-btn" style="padding:8px;color:var(--success);flex-shrink:0;" onclick="processarDecisaoAdmin(\'' + emailSeg + '\', \'ATIVO\', this)" title="Reativar"><span class="material-symbols-outlined" style="font-size:16px;">check_circle</span></button>' +
              '</div>' +
            '</div>';
          }).join('');
        }
      }
    }).obterTodosUsuarios();
  }

  function aprovarTodosPendentes() {
    chamarAlertaTatico("APROVAR TODOS", "Deseja realmente aprovar TODOS os operadores pendentes?", "CONFIRMAR", function(confirma) {
      if (confirma) {
        google.script.run.withSuccessHandler(function(res) {
          if (res.sucesso) {
            chamarAlertaTatico("SUCESSO", res.mensagem, "OK");
            carregarUsuariosGrade();
          } else {
            chamarAlertaTatico("ERRO", res.mensagem, "OK");
          }
        }).aprovarTodosUsuariosPendentes();
      }
    });
  }

  function processarDecisaoAdmin(email, decisao, botao) {
    var card = botao.closest('.card-operador'); 
    card.style.opacity = "0.3"; 
    botao.disabled = true;
    google.script.run.withSuccessHandler(function(res) {
      if (res.sucesso) { 
        chamarAlertaTatico("OK", "Operador: " + decisao, "OK"); 
        carregarUsuariosGrade();
      } else { 
        chamarAlertaTatico("FALHA", res.mensagem, "OK"); 
        card.style.opacity = "1"; 
        botao.disabled = false;
      }
    }).alterarStatusOperador(email, decisao);
  }
</script>
