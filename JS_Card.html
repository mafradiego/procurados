<script>
/* ================================================================
     MÓDULO 4: CARD DE INTELIGÊNCIA (NEUMÓRFICO)
     ================================================================ */
  function montarCardInteligencia(alvo) {
    // Função helper para evitar XSS (Sanitização)
    const escapeHTML = function(str) {
      if (!str) return "";
      return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
    };

    const btnCopiar = function(txt) {
      if (!txt || txt === "N/A") return "";
      const txtSeguro = escapeHTML(txt);
      return '<span class="icone-copiar" onclick="copiarDadoTatico(\'' + txtSeguro + '\', this)" title="Copiar"><span class="material-symbols-outlined" style="font-size:14px;">content_copy</span></span>';
    };

    let extras = {};
    try { extras = JSON.parse(alvo.dadosExtrasJSON || "{}"); } catch(e) {}

    const isValido = function(val) {
      if (!val) return false;
      const v = val.trim().toLowerCase();
      return !["", "n/a", "-", "não", "nao", "não informado", "nao informado", "não informada", "nao informada", "não consta", "nao consta"].includes(v);
    };

    // Classificar crime
    const crimeInfo = classificarCrime(alvo.infoProcessuais || '');

    // Dias restantes para validade
    const diasRestantes = calcularDiasRestantes(alvo.validade);
    let validadeClass = '';
    let validadeTexto = '';
    // Helper: converte dias para "X ano(s), Y mês(es) e Z dia(s)"
    const formatarDiasRestantes = function(dias) {
      if (dias <= 0) return Math.abs(dias) + ' dias';
      const anos = Math.floor(dias / 365);
      const meses = Math.floor((dias % 365) / 30);
      const d = dias % 30;
      let parts = [];
      if (anos > 0) parts.push(anos + (anos === 1 ? ' ano' : ' anos'));
      if (meses > 0) parts.push(meses + (meses === 1 ? ' mês' : ' meses'));
      if (d > 0 || parts.length === 0) parts.push(d + (d === 1 ? ' dia' : ' dias'));
      return parts.join(', ');
    };
    if (diasRestantes !== null) {
      const limVerde = parseInt(configsSistema.validade_verde_dias) || 91;
      const limAmarelo = parseInt(configsSistema.validade_amarelo_dias) || 90;
      const limLaranja = parseInt(configsSistema.validade_laranja_dias) || 30;
      const limVermelho = parseInt(configsSistema.validade_vermelho_dias) || 15;
      const textoFormatado = formatarDiasRestantes(diasRestantes);

      if (diasRestantes <= 0) { validadeClass = 'validade-vencido'; validadeTexto = '<span class="material-symbols-outlined" style="font-size:12px;vertical-align:middle;">warning</span> VENCIDO há ' + textoFormatado; }
      else if (diasRestantes <= limVermelho) { validadeClass = 'validade-vermelho'; validadeTexto = '<span class="material-symbols-outlined" style="font-size:12px;vertical-align:middle;">error</span> Faltam ' + textoFormatado; }
      else if (diasRestantes <= limLaranja) { validadeClass = 'validade-laranja'; validadeTexto = '<span class="material-symbols-outlined" style="font-size:12px;vertical-align:middle;">schedule</span> Faltam ' + textoFormatado; }
      else if (diasRestantes <= limAmarelo) { validadeClass = 'validade-amarelo'; validadeTexto = '<span class="material-symbols-outlined" style="font-size:12px;vertical-align:middle;">timer</span> Faltam ' + textoFormatado; }
      else { validadeClass = 'validade-verde'; validadeTexto = '<span class="material-symbols-outlined" style="font-size:12px;vertical-align:middle;">check_circle</span> ' + textoFormatado; }
    }

    // Tooltip helper
    const tooltip = function(texto) {
      const t = configsSistema['tooltip_' + texto] || texto;
      return '<span class="tooltip-trigger">?<span class="tooltip-balloon">' + t + '</span></span>';
    };

    // Distância KM (Haversine)
    let htmlDistancia = '';
    if (posicaoPatrulheiro) {
      const km = calcularHaversine(posicaoPatrulheiro.lat, posicaoPatrulheiro.lng, alvo.lat, alvo.lng);
      htmlDistancia = '<div class="km-badge"><span class="material-symbols-outlined" style="font-size:16px;">near_me</span> ' + km.toFixed(1) + ' km' + tooltip('distancia') + '</div>';
    }

    // Secundários com numeração
    let htmlSecundarios = "";
    // Construir htmlArea para cada endereço secundário (reutilizável)
    const montarAreaBadge = function(areaData) {
      if (!areaData) return '';
      if (areaData.cpi === "FORA DO ESTADO") {
        return '<div style="display:flex;margin-top:4px;"><span style="font-size:9px;font-weight:800;color:var(--danger);padding:3px 8px;background:rgba(255,59,48,0.15);border-radius:3px;border:1px solid rgba(255,59,48,0.3);">FORA DO ESTADO</span></div>';
      }
      let h = '<div style="display:flex;flex-wrap:wrap;gap:4px;margin-top:4px;">';
      if (areaData.cpi) h += '<span style="font-size:9px;font-weight:800;color:var(--accent);padding:3px 6px;background:var(--accent)15;border-radius:3px;border:1px solid var(--accent);">COMANDO: ' + areaData.cpi + '</span>';
      if (areaData.batalhao) h += '<span style="font-size:9px;font-weight:700;color:var(--text-primary);padding:3px 6px;background:rgba(255,255,255,0.08);border-radius:3px;border:1px solid rgba(255,255,255,0.2);">BTL: ' + areaData.batalhao + '</span>';
      if (areaData.cia) h += '<span style="font-size:9px;font-weight:700;color:var(--text-primary);padding:3px 6px;background:rgba(255,255,255,0.08);border-radius:3px;border:1px solid rgba(255,255,255,0.2);">CIA: ' + areaData.cia + '</span>';
      if (areaData.delegacia) h += '<span style="font-size:9px;font-weight:700;color:var(--text-secondary);padding:3px 6px;background:rgba(255,255,255,0.04);border-radius:3px;border:1px solid rgba(255,255,255,0.1);">DP: ' + areaData.delegacia + '</span>';
      h += '</div>';
      return h;
    };
    try {
      const listaSec = JSON.parse(alvo.geodataSecundarios || "[]");
      if (listaSec.length > 0) {
        listaSec.forEach(function(sec, idx) {
          let kmSec = '';
          if (posicaoPatrulheiro && typeof sec.lat === 'number' && typeof sec.lng === 'number') {
            const d = calcularHaversine(posicaoPatrulheiro.lat, posicaoPatrulheiro.lng, sec.lat, sec.lng);
            kmSec = '<div class="km-badge" style="font-size:10px;padding:4px 8px;margin-bottom:6px;"><span class="material-symbols-outlined" style="font-size:14px;">near_me</span> ' + d.toFixed(1) + ' km</div>';
          }
          // Área do endereço secundário
          let areaSec = montarAreaBadge(sec);
          htmlSecundarios += '<div class="endereco-sec-item">' + kmSec +
            '<div class="endereco-texto"><strong>Outro Endereço ' + (idx + 1) + ':</strong>' + sec.endereco + btnCopiar(sec.endereco) + '</div>' +
            areaSec +
            '<div style="display:flex;gap:6px;">' +
              '<button class="btn-rota-sec" onclick="abrirGpsNativo(' + sec.lat + ', ' + sec.lng + ')"><span class="material-symbols-outlined" style="font-size:16px;">explore</span> ROTA</button>' +
              '<button class="btn-rota-sec" style="width:auto;padding:8px;background:#FBBC04;" onclick="abrirStreetView(' + sec.lat + ', ' + sec.lng + ')" title="Street View"><svg width="16" height="16" viewBox="0 0 24 24" fill="#1a1a1a"><path d="M12.56 14.33c-.34.27-.56.7-.56 1.17V21h7c1.1 0 2-.9 2-2v-5.98c-.94-.33-1.95-.52-3-.52-2.03 0-3.93.7-5.44 1.83zM12 6c1.93 0 3.5-1.57 3.5-3.5S13.93-1 12-1 8.5.57 8.5 2.5 10.07 6 12 6zm-1.44 8.33C9.07 12.93 7.14 12.23 5 12.23c-1.05 0-2.06.19-3 .52V19c0 1.1.9 2 2 2h7v-5.5c0-.47-.22-.9-.56-1.17l-.88-.67zM12 8c-1.66 0-3 1.34-3 3 0 .35.07.69.18 1.01L12 14.34l2.82-2.33c.11-.32.18-.66.18-1.01 0-1.66-1.34-3-3-3z"/></svg></button>' +
            '</div></div>';
        });
      } else {
        htmlSecundarios = '<div style="color:var(--text-muted);font-size:11px;text-align:center;padding:10px;">Nenhum endereço adicional.</div>';
      }
    } catch (e) { htmlSecundarios = "Erro ao ler endereços."; }

    // Validade formatada
    let dataPura = alvo.validade.match(/(\d{2}[\/\.]\d{2}[\/\.]\d{4})/);
    dataPura = dataPura ? dataPura[1].replace(/\./g, '/') : alvo.validade;

    // Perfil
    let htmlPerfil = '';
    const addP = function(lbl, val) { if (isValido(val)) htmlPerfil += '<div style="margin:0;padding:2px 0;"><strong>' + lbl + ':</strong> ' + val + '</div>'; };
    addP("Nome Social", extras.nomeSocial); addP("RJI", extras.rji); addP("Alcunha", extras.alcunha);
    addP("Nasc", alvo.nascimento); addP("Sexo", alvo.sexo); addP("Cor", alvo.cor);
    addP("Natural", alvo.naturalidade); addP("Filiação", alvo.filiacao); addP("Marcas", extras.marcas);

    // Info Processual
    const infoRaw = alvo.infoProcessuais || "";
    const extP = function(txt, lbl) { const r = new RegExp(lbl + '\\s*([^|\\n]+)', 'i'); const m = txt.match(r); return m ? m[1].trim() : "N/A"; };
    const nProcesso = extP(infoRaw, 'Nº do processo:');
    const orgao = extP(infoRaw, 'Órgão Judicial:');
    let especie = extP(infoRaw, 'Espécie(?: de| da) (?:prisão|Internação):');
    if (especie === "N/A") especie = extP(infoRaw, 'Espécie de prisão:');
    let lei = extP(infoRaw, 'Lei:'); let artigo = extP(infoRaw, 'Artigo:'); let paragrafo = extP(infoRaw, 'Parágrafo:');
    if (lei !== "N/A" && lei.toLowerCase().includes("art")) {
      let mP = lei.match(/§\s*(.+)$/); if (mP && paragrafo === "N/A") paragrafo = mP[1].trim();
      let mA = lei.match(/art\.?\s*([^,§|]+)/i); if (mA && artigo === "N/A") artigo = mA[1].trim();
      lei = lei.split(/(?:,|\s)\s*art\.?/i)[0].trim();
    }
    if (lei === "N/A") lei = "-"; if (artigo === "N/A") artigo = "-"; if (paragrafo === "N/A") paragrafo = "-";
    let pena = extP(infoRaw, 'Pena restante:');
    if (!isValido(pena)) pena = extP(infoRaw, 'Prazo Mínimo da Internação:');
    if (!isValido(pena)) pena = "-";
    let regime = extP(infoRaw, 'Regime Prisional:');

    let htmlSintese = '';
    if (isValido(extras.sintese)) {
      htmlSintese = '<div style="margin-top:12px;padding:12px;background:var(--bg-deep);box-shadow:var(--neu-pressed-sm);border-radius:8px;font-size:11px;color:var(--text-secondary);line-height:1.5;"><strong style="color:var(--danger);display:flex;align-items:center;gap:4px;margin-bottom:4px;"><span class="material-symbols-outlined" style="font-size:14px;">description</span> SÍNTESE</strong>' + escapeHTML(extras.sintese) + '</div>';
    }


    // Bloco de Área (CPI/BPM/CIA/DP) para o endereço principal
    let htmlArea = '';
    if (alvo.cpi || alvo.bpmArea || alvo.ciaArea || alvo.dpArea) {
      if (alvo.cpi === "FORA DO ESTADO") {
        htmlArea = '<div style="display:flex;margin-top:8px;"><span style="font-size:10px;font-weight:800;color:var(--danger);padding:4px 10px;background:rgba(255,59,48,0.15);border-radius:4px;border:1px solid rgba(255,59,48,0.3);">FORA DO ESTADO</span></div>';
      } else {
        htmlArea = '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-top:8px;padding:8px 10px;background:var(--bg-depressed);border-radius:8px;border-left:3px solid var(--accent);">';
        if (alvo.cpi) htmlArea += '<span style="font-size:10px;font-weight:800;color:var(--accent);padding:4px 8px;background:var(--accent)15;border-radius:4px;border:1px solid var(--accent);">COMANDO: ' + alvo.cpi + '</span>';
        if (alvo.bpmArea) htmlArea += '<span style="font-size:10px;font-weight:700;color:var(--text-primary);padding:4px 8px;background:rgba(255,255,255,0.08);border-radius:4px;border:1px solid rgba(255,255,255,0.2);">BTL: ' + alvo.bpmArea + '</span>';
        if (alvo.ciaArea) htmlArea += '<span style="font-size:10px;font-weight:700;color:var(--text-primary);padding:4px 8px;background:rgba(255,255,255,0.08);border-radius:4px;border:1px solid rgba(255,255,255,0.2);">CIA: ' + alvo.ciaArea + '</span>';
        if (alvo.dpArea) htmlArea += '<span style="font-size:10px;font-weight:700;color:var(--text-secondary);padding:4px 8px;background:rgba(255,255,255,0.04);border-radius:4px;border:1px solid rgba(255,255,255,0.1);">DP: ' + alvo.dpArea + '</span>';
        if (alvo.cidade) htmlArea += '<span style="font-size:10px;font-weight:700;color:var(--text-muted);padding:4px 8px;background:rgba(255,255,255,0.04);border-radius:4px;border:1px solid rgba(255,255,255,0.1);"><span class="material-symbols-outlined" style="font-size:12px;vertical-align:middle;">location_city</span> ' + alvo.cidade + '</span>';
        htmlArea += '</div>';
      }
    }

    let tituloDisplay = extras.titulo || "MANDADO DE PRISÃO";

    // Conferência (Admin only)
    let htmlConferir = '';
    if (usuarioLogado && usuarioLogado.perfil === "Admin") {
      htmlConferir = '<button class="neu-btn" style="width:100%;margin-top:8px;padding:10px;font-size:11px;" onclick="conferirMandadoClick(\'' + alvo.mandado.replace(/'/g, "\\'") + '\', this)"><span class="material-symbols-outlined" style="font-size:16px;">verified</span> MARCAR COMO CONFERIDO</button>';
    }

    let chkProcurado = (!alvo.status || alvo.status.toUpperCase() !== 'CAPTURADO') ? 'checked' : '';
    let chkCapturado = (alvo.status && alvo.status.toUpperCase() === 'CAPTURADO') ? 'checked' : '';

    return '<div class="card-container">' +
      '<div class="card-header-full">' +

        // Linha 0: Data do documento BNMP (topo discreto)
        (extras.emissao ? '<div style="font-size:9px;color:var(--text-muted);text-align:center;font-family:var(--font-mono);letter-spacing:0.5px;opacity:0.7;margin-bottom:8px;">Inserido no BNMP em: ' + (function(){ var m = (extras.emissao || '').match(/(\d{2}[\/\.]\d{2}[\/\.]\d{4})/); return m ? m[1].replace(/\./g,'/') : extras.emissao; })() + '</div>' : '') +

        // Linha 1: Tipo do mandado + Crime (mesma faixa)
        '<div style="display:flex;align-items:center;justify-content:center;gap:8px;flex-wrap:wrap;">' +
          '<div style="font-size:11px;font-weight:900;letter-spacing:1.2px;text-transform:uppercase;color:var(--text-primary);padding:5px 14px;background:rgba(255,255,255,0.06);border-radius:99px;border:1px solid rgba(255,255,255,0.1);">' + tituloDisplay + '</div>' +
          '<div class="crime-badge" style="background:' + crimeInfo.cor + ';color:white;margin:0;">' +
            '<svg width="18" height="18" viewBox="0 0 16 16" fill="none" style="flex-shrink:0;"><g fill="white">' + crimeInfo.svgPath + '</g></svg> ' + crimeInfo.nome +
          '</div>' +
        '</div>' +

        // Linha 2: Número do mandado (centralizado, destaque mono, maior)
        '<div style="display:flex;align-items:center;justify-content:center;gap:6px;margin:10px 0 8px;overflow:hidden;">' +
          '<span class="material-symbols-outlined" style="font-size:15px;color:var(--accent);opacity:0.7;flex-shrink:0;">gavel</span>' +
          '<span style="font-family:var(--font-mono);font-size:13px;font-weight:700;color:var(--accent);letter-spacing:0.3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">' + alvo.mandado + '</span>' + btnCopiar(alvo.mandado) +
        '</div>' +

        // Linha 3: Validade com destaque visual
        '<div class="data-badge ' + validadeClass + '" style="margin-bottom:8px;text-align:center;">Documento válido até: <span class="data-valor" style="font-weight:900;">' + dataPura + '</span>' +
          (validadeTexto ? '<span class="validade-restante" style="margin-left:8px;">' + validadeTexto + '</span>' : '') +
        '</div>' +

        // Linha 4: Datas secundárias (lançamento + conferência, discretas)
        '<div style="display:flex;flex-wrap:wrap;gap:4px;justify-content:center;margin-top:6px;">' +
          '<div style="font-size:9px;color:var(--text-muted);padding:2px 8px;background:rgba(255,255,255,0.03);border-radius:4px;">Lançamento: <strong>' + (alvo.dataLancamento || '—') + '</strong></div>' +
          '<div style="font-size:9px;color:var(--text-muted);padding:2px 8px;background:rgba(255,255,255,0.03);border-radius:4px;">Conferido: <strong>' + (alvo.dataConferencia || '—') + '</strong></div>' +
        '</div>' +

      '</div>' +

      '<div class="grid-wrapper">' +
        '<div class="coluna-tatica-1">' +
          '<div class="bloco-perfil">' +
            '<div class="identificacao-top-row">' +
              '<span class="cpf-alvo">CPF: ' + (alvo.cpf || 'N/A') + btnCopiar(alvo.cpf) + '</span>' +
              '<span class="rg-alvo">RG: ' + (alvo.rg || 'N/A') + btnCopiar(alvo.rg) + '</span>' +
            '</div>' +
            '<div class="perfil-row">' +
              '<div class="nome-alvo-destaque">' + alvo.nome + btnCopiar(alvo.nome) + '</div>' +
              '<div class="perfil-inner-row">' +
                (function() {
                  let link = alvo.fotoUrl;
                  if (link && link.includes("uc?id=")) link = link.replace("uc?id=", "thumbnail?id=") + "&sz=w500";
                  if (link && link !== "N/A" && link.startsWith("http")) return '<img src="' + link + '" class="foto-perfil-card" />';
                  return '<div class="foto-placeholder"><span class="material-symbols-outlined" style="font-size:28px;">photo_camera</span>Sem Foto</div>';
                })() +
                '<div style="flex:1;"><div class="dados-colunas">' + htmlPerfil + '</div></div>' +
              '</div>' +
            '</div>' +
          '</div>' +

          '<div class="bloco-enderecos">' +
            '<div class="enderecos-wrapper-box">' +
              '<div class="enderecos-box-title"><span class="material-symbols-outlined" style="color:var(--danger);">location_on</span> ENDEREÇOS MAPEADOS ' + tooltip('enderecos') + '</div>' +
              htmlDistancia +
              '<div class="endereco-main-item">' +
              '<div class="endereco-texto"><strong>Endereço Principal:</strong>' + alvo.enderecoPrincipal + btnCopiar(alvo.enderecoPrincipal) + '</div>' +
                htmlArea +
                '<div style="display:flex;gap:6px;">' +
                  '<button class="btn-rota" onclick="abrirGpsNativo(' + alvo.lat + ', ' + alvo.lng + ')"><span class="material-symbols-outlined" style="font-size:16px;">directions_car</span> ROTA</button>' +
                  '<button class="btn-rota" style="width:auto;padding:8px 12px;background:#FBBC04;color:#1a1a1a;display:inline-flex;align-items:center;gap:4px;font-weight:700;" onclick="abrirStreetView(' + alvo.lat + ', ' + alvo.lng + ')" title="Street View"><svg width="16" height="16" viewBox="0 0 24 24" fill="#1a1a1a"><path d="M12.56 14.33c-.34.27-.56.7-.56 1.17V21h7c1.1 0 2-.9 2-2v-5.98c-.94-.33-1.95-.52-3-.52-2.03 0-3.93.7-5.44 1.83zM12 6c1.93 0 3.5-1.57 3.5-3.5S13.93-1 12-1 8.5.57 8.5 2.5 10.07 6 12 6zm-1.44 8.33C9.07 12.93 7.14 12.23 5 12.23c-1.05 0-2.06.19-3 .52V19c0 1.1.9 2 2 2h7v-5.5c0-.47-.22-.9-.56-1.17l-.88-.67zM12 8c-1.66 0-3 1.34-3 3 0 .35.07.69.18 1.01L12 14.34l2.82-2.33c.11-.32.18-.66.18-1.01 0-1.66-1.34-3-3-3z"/></svg> STREET</button>' +
                '</div>' +
              '</div>' +
              '<div class="outros-enderecos">' + htmlSecundarios + '</div>' +
            '</div>' +
          '</div>' +
        '</div>' +

        '<div class="coluna-tatica-2">' +
          '<div class="bloco-processo">' +
            '<div class="info-processual-box">' +
              '<div class="proc-titulo-centro"><span class="material-symbols-outlined" style="font-size:16px;vertical-align:middle;">gavel</span> DADOS DO PROCESSO</div>' +
              '<div class="proc-item-centro"><strong>Nº do processo:</strong><br><span class="proc-numero-destaque">' + nProcesso + btnCopiar(nProcesso) + '</span></div>' +
              '<div class="proc-item-comum"><strong>Órgão:</strong> ' + orgao + '</div>' +
              '<div class="proc-item-comum"><strong>Espécie:</strong> ' + especie + '</div>' +
              '<div class="proc-tipificacao-box" style="border-left-color:' + crimeInfo.cor + ';"><strong><svg width="14" height="14" viewBox="0 0 16 16" style="vertical-align:middle;margin-right:4px;"><g fill="' + crimeInfo.cor + '">' + crimeInfo.svgPath + '</g></svg> TIPIFICAÇÃO PENAL</strong><br>Lei: ' + lei + ' | Art: ' + artigo + ' | Par: ' + paragrafo + '</div>' +
              '<div class="proc-pena-regime-box"><strong><span class="material-symbols-outlined" style="font-size:14px;vertical-align:middle;">balance</span> SENTENÇA / EXECUÇÃO</strong><div class="proc-pena-linha"><span class="proc-pena-label">Tempo:</span><span class="proc-pena-valor">' + pena + '</span></div><div class="proc-pena-linha"><span class="proc-pena-label">Regime:</span><span class="proc-pena-valor">' + regime + '</span></div></div>' +
              htmlSintese +
            '</div>' +
          '</div>' +

          '<div class="bloco-form">' +
            '<div class="form-atualizacao">' +
              '<div class="label-instrucao"><span class="material-symbols-outlined" style="font-size:18px;">edit_note</span> ÁREA INTERATIVA ' + tooltip('area_interativa') + '</div>' +
              '<div style="font-size:11px;color:var(--text-muted);margin:6px 0 12px;text-align:center;">Mude o status se o indivíduo foi capturado:</div>' +
              '<div class="status-toggle">' +
                '<input type="radio" id="st_proc_' + alvo.mandado + '" name="status_' + alvo.mandado + '" value="Procurado" ' + chkProcurado + '>' +
                '<label for="st_proc_' + alvo.mandado + '">PROCURADO</label>' +
                '<input type="radio" id="st_cap_' + alvo.mandado + '" name="status_' + alvo.mandado + '" value="Capturado" ' + chkCapturado + '>' +
                '<label for="st_cap_' + alvo.mandado + '">CAPTURADO</label>' +
              '</div>' +
              '<textarea id="obs_' + alvo.mandado + '" placeholder="Informações de inteligência...">' + escapeHTML(alvo.observacoes) + '</textarea>' +
              htmlConferir +
            '</div>' +
          '</div>' +
        '</div>' +
      '</div>' +

      '<div class="sticky-footer">' +
        '<button class="btn-salvar" onclick="enviarAtualizacao(\'' + alvo.mandado.replace(/'/g, "\\'") + '\', this)">SALVAR LOG DA ABORDAGEM</button>' +
      '</div>' +
    '</div>';
  }

  /* ================================================================
     MÓDULO 6: CLASSIFICAÇÃO DE CRIMES (SVG ICONS)
     ================================================================ */
  function classificarCrime(infoProc) {
    const texto = (infoProc || '').toUpperCase();

    // Gerar SVG de texto para dentro do pino (viewBox 16x16)
    function gerarSvgTexto(label) {
      var tam = 9;
      if (label.length <= 2) tam = 12;
      else if (label.length <= 3) tam = 10;
      else if (label.length <= 4) tam = 8.5;
      else if (label.length <= 5) tam = 7.5;
      else tam = 6.5;
      return '<text x="8" y="11.5" text-anchor="middle" font-size="' + tam + '" font-weight="900" font-family="Impact,Arial Black,sans-serif" fill="white" letter-spacing="0.3">' + label + '</text>';
    }

    // 1. Tentar match pela tabela de leis do banco
    if (window._tabelaLeis && window._tabelaLeis.length > 0) {
      for (var i = 0; i < window._tabelaLeis.length; i++) {
        var lei = window._tabelaLeis[i];
        if (!lei.palavrasChave) continue;
        var keywords = lei.palavrasChave.toUpperCase().split('|');
        for (var k = 0; k < keywords.length; k++) {
          var kw = keywords[k].trim();
          if (kw && texto.includes(kw)) {
            var pt = lei.pinoTexto || lei.categoria.substring(0, 5).toUpperCase();
            return {
              nome: lei.categoria,
              cor: lei.cor || '#6b7280',
              pinoTexto: pt,
              svgPath: gerarSvgTexto(pt),
              naoClassificado: false
            };
          }
        }
      }
    }

    // 2. Fallback hardcoded caso _tabelaLeis ainda não tenha sido carregada
    var fallback = [
      { kw: ['HOMIC','121','LATROCÍNIO','LATROCINIO'], nome: 'Homicídio', cor: '#8b0000', pt: 'HOMIC' },
      { kw: ['ESTUPRO','213','217'], nome: 'Estupro', cor: '#991b1b', pt: 'JACK' },
      { kw: ['MARIA DA PENHA','VIOLÊNCIA DOMÉSTICA','VIOLENCIA DOMESTICA','11340'], nome: 'V. Doméstica', cor: '#e11d48', pt: 'PENHA' },
      { kw: ['ROUBO','157'], nome: 'Roubo', cor: '#dc2626', pt: 'ROUBO' },
      { kw: ['FURTO','155'], nome: 'Furto', cor: '#ea580c', pt: 'FURTO' },
      { kw: ['TRÁFICO','TRAFICO','DROGAS','ENTORPECENTE'], nome: 'Tráfico', cor: '#7c3aed', pt: 'TRÁF' },
      { kw: ['ESTELIONATO','171'], nome: 'Estelionato', cor: '#ca8a04', pt: '171' },
      { kw: ['RECEPTAÇÃO','RECEPTACAO','180'], nome: 'Receptação', cor: '#b45309', pt: '180' },
      { kw: ['CIVIL','ALIMENTO','PENSÃO'], nome: 'Alimentícia', cor: '#2563eb', pt: 'CIVIL' },
      { kw: ['ARMA','PORTE','POSSE','244','10826'], nome: 'Porte de Arma', cor: '#475569', pt: 'ARMA' },
      { kw: ['TRÂNSITO','TRANSITO','CTB'], nome: 'Trânsito', cor: '#0891b2', pt: 'CTB' },
      { kw: ['SEQUESTRO','CÁRCERE','148'], nome: 'Sequestro', cor: '#7f1d1d', pt: '148' },
      { kw: ['LESÃO CORPORAL','LESAO CORPORAL','129'], nome: 'Lesão Corporal', cor: '#b91c1c', pt: 'LESÃO' },
      { kw: ['ASSOCIAÇÃO','ASSOCIACAO','288'], nome: 'Organização Crim.', cor: '#4338ca', pt: '288' },
      { kw: ['AMEAÇA','AMEACA','147'], nome: 'Ameaça', cor: '#9f1239', pt: '147' },
      { kw: ['FISCAL','TRIBUTÁRIO'], nome: 'Fiscal', cor: '#059669', pt: 'FISCAL' }
    ];

    for (var f = 0; f < fallback.length; f++) {
      for (var j = 0; j < fallback[f].kw.length; j++) {
        if (texto.includes(fallback[f].kw[j])) {
          return { nome: fallback[f].nome, cor: fallback[f].cor, pinoTexto: fallback[f].pt, svgPath: gerarSvgTexto(fallback[f].pt), naoClassificado: false };
        }
      }
    }

    // 3. Se não bateu com nada → NÃO CLASSIFICADO (alerta amarelo)
    if (texto.trim().length > 0) {
      return { nome: 'Não Classificado', cor: '#f59e0b', pinoTexto: '???', svgPath: gerarSvgTexto('???'), naoClassificado: true };
    }

    // 4. Sem informação processual
    return { nome: 'Outros', cor: '#6b7280', pinoTexto: 'CRIME', svgPath: gerarSvgTexto('CRIME'), naoClassificado: false };
  }
</script>
