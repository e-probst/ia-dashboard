// ============================================================
//  Cronograma MensalMailer.gs  —  Google Apps Script backend
//  Cronograma Mensal · Mabu Hospitalidade
//
//  COMO IMPLANTAR:
//  1. Acesse script.google.com → Novo projeto
//  2. Cole todo este código
//  3. Implantar → Nova implantação → Aplicativo da Web
//     - Executar como: Eu (minha conta Google)
//     - Quem pode acessar: Qualquer pessoa
//  4. Copie a URL gerada e cole no painel Cronograma Mensal (ícone ⚙)
//  5. Execute createDailyTrigger() UMA VEZ para ativar envio às 8h
// ============================================================

// ── CONFIGURAÇÃO — dados sensíveis ficam no PropertiesService, não no código ──
// Execute setupConfig() UMA VEZ no editor do GAS para configurar.
// Isso evita expor credenciais no GitHub.
var SPREADSHEET_ID  = '';
var EMAIL_FROM_NAME = 'Cronograma Mensal · Mabu Hospitalidade';
var ADMIN_EMAIL     = '';
var GAS_WEB_APP_URL = '';
var API_SECRET      = ''; // Chave exigida em todas as ações POST destrutivas

// Carrega configuração do PropertiesService a cada execução (nunca hardcoded)
(function _loadConfig() {
  try {
    var p = PropertiesService.getScriptProperties();
    SPREADSHEET_ID  = p.getProperty('SPREADSHEET_ID')  || '';
    ADMIN_EMAIL     = p.getProperty('ADMIN_EMAIL')      || '';
    GAS_WEB_APP_URL = p.getProperty('GAS_WEB_APP_URL')  || '';
    API_SECRET      = p.getProperty('API_SECRET')       || '';
  } catch(e) { Logger.log('_loadConfig error: ' + e.message); }
})();

// ── SETUP INICIAL — execute UMA VEZ no editor após implantar ─────────────────
// Preencha os valores abaixo e clique em ▶ Executar esta função:
function setupConfig() {
  var p = PropertiesService.getScriptProperties();
  p.setProperties({
    'SPREADSHEET_ID':  '1FMoWYDqersAk8zXy_a_ZDShUDU-s339eOC9f5P2mKfY',
    'ADMIN_EMAIL':     'e.probst@mymabu.com.br',
    'GAS_WEB_APP_URL': 'https://script.google.com/macros/s/AKfycbwK8FMgwRXkF-Z6krN12yHKgMJmDNxsgVBZoka4PJgSlNYx4f29wxs3XTKtW35B27Tc/exec',
    'API_SECRET':      p.getProperty('API_SECRET') || Utilities.getUuid().replace(/-/g,''),
  });
  var secret = p.getProperty('API_SECRET');
  Logger.log('✅ Config salva! Chave de API: ' + secret);
}

function checkConfig() {
  var p = PropertiesService.getScriptProperties();
  Logger.log('SPREADSHEET_ID:  ' + p.getProperty('SPREADSHEET_ID'));
  Logger.log('ADMIN_EMAIL:     ' + p.getProperty('ADMIN_EMAIL'));
  Logger.log('GAS_WEB_APP_URL: ' + p.getProperty('GAS_WEB_APP_URL'));
  Logger.log('API_SECRET:      ' + p.getProperty('API_SECRET'));
}

// ── ROTEADOR PRINCIPAL ───────────────────────────────────────

function doGet(e) {
  var params   = (e && e.parameter) || {};
  var action   = params.action || '';
  var callback = params.callback;

  // Confirmação de entrega via link do e-mail
  if (action === 'confirm') {
    return handleConfirmDelivery(params);
  }

  // Visualização de status sem confirmar (link "Ver status" no e-mail)
  if (action === 'status') {
    return handleViewStatus(params);
  }

  // Busca confirmações pendentes (chamado pelo dashboard no init)
  if (action === 'get_confirmations') {
    return handleGetConfirmations(callback);
  }

  // Lê todas as tarefas da planilha (banco de dados → dashboard)
  if (action === 'get_tasks') {
    return handleGetTasks(callback);
  }

  // Ping / teste de conexão (JSONP)
  var payload = JSON.stringify({ ok: true, message: 'Cronograma Mensal Mailer online ✅', ts: new Date().toISOString() });
  var content = callback ? callback + '(' + payload + ')' : payload;
  return ContentService.createTextOutput(content)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

// ── HELPER JSONP/JSON ─────────────────────────────────────────
function jsonpResponse(obj, callback) {
  var payload = JSON.stringify(obj);
  var content = callback ? callback + '(' + payload + ')' : payload;
  return ContentService.createTextOutput(content)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

// ── CONFIRMAÇÃO DE ENTREGA ────────────────────────────────────

function handleConfirmDelivery(params) {
  var id    = params.id    || '';
  var name  = params.name  || 'Tarefa';
  var prazo = params.prazo || '';
  var resp  = params.resp  || '';
  var month = params.month || '';

  if (!id) {
    return HtmlService.createHtmlOutput('<h2>Link inválido.</h2>');
  }

  // Verifica se já confirmado
  var props = PropertiesService.getScriptProperties();
  var key   = 'confirm_' + id;
  var already = props.getProperty(key);

  if (!already) {
    var confirmedAt   = new Date().toISOString();
    // confirmedDate é a data em horário de Brasília (BRT); confirmedAt (UTC ISO) pode divergir
    // por 1 dia perto da meia-noite → dashboard usa confirmedDate para exibir a data correta
    var confirmedDate = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
    var data = JSON.stringify({ id: id, name: name, prazo: prazo, confirmedAt: confirmedAt, confirmedDate: confirmedDate });
    props.setProperty(key, data);
    Logger.log('Entrega confirmada via e-mail: id=' + id + ' | ' + name);
    // Atualiza imediatamente a planilha (evita lag até o próximo polling do dashboard)
    if (SPREADSHEET_ID) {
      try {
        var today = confirmedDate; // reutiliza a data BRT já calculada
        confirmDeliveryInSheet(id, today, prazo, month);
      } catch(e) {
        Logger.log('handleConfirmDelivery: erro ao atualizar planilha: ' + e.message);
      }
    }
  }

  var allRespTasks = getRespTasksForPanel(resp, id, props);

  return HtmlService.createHtmlOutput(buildConfirmationPage(name, prazo, !!already, resp, allRespTasks))
    .setTitle('Entrega Confirmada · Cronograma Mensal');
}

// Abre o painel de entregas sem confirmar nada (link "Ver status" no e-mail)
function handleViewStatus(params) {
  var id   = params.id   || '';
  var resp = params.resp || '';
  var name = params.name || '';

  var props = PropertiesService.getScriptProperties();
  var already = id ? !!props.getProperty('confirm_' + id) : false;

  var allRespTasks = getRespTasksForPanel(resp, id, props);

  return HtmlService.createHtmlOutput(buildConfirmationPage(name, params.prazo||'', already, resp, allRespTasks, true))
    .setTitle('Minhas Entregas · Cronograma Mensal');
}

// Busca tarefas do responsável para "Minhas Entregas".
// Prioridade: usa lista armazenada do último disparo de e-mail (resp_sent_*).
// Fallback: janela de 30d no Sheets (se nunca houve disparo).
function getRespTasksForPanel(resp, currentId, props) {
  var result = [];
  if (!resp || !SPREADSHEET_ID) return result;

  var today = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
  var tParts = today.split('/');
  var todayDate = new Date(+tParts[2], +tParts[1]-1, +tParts[0]);

  // ── Cache: evita reler planilha em cliques repetidos (TTL 2 min) ──
  var cache    = CacheService.getScriptCache();
  var respKeyNorm = resp.trim().toLowerCase()
    .replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,'').slice(0,60);
  var cacheKey = 'panel_' + respKeyNorm + '_' + today;
  // Invalida cache se currentId presente (confirmação recém feita — dados mudaram)
  if (!currentId) {
    var cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch(e) {}
    }
  }

  // ── Recupera lista armazenada do último e-mail disparado para este resp ──
  var storedRaw = props.getProperty('resp_sent_' + respKeyNorm);
  var storedIds = null; // {id → {name, prazo}} se existir

  if (storedRaw) {
    try {
      var stored = JSON.parse(storedRaw);
      if (stored.tasks && stored.tasks.length) {
        storedIds = {};
        stored.tasks.forEach(function(t) {
          if (t.id) storedIds[String(t.id)] = {name: t.name||'', prazo: t.prazo||''};
        });
      }
    } catch(e) { Logger.log('resp_sent parse error: ' + e.message); }
  }

  // ── Pré-carrega TODAS as confirmações de uma vez (evita N chamadas getProperty no loop) ──
  var allProps = props.getProperties();

  var OVERDUE_WINDOW = 30;

  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Todos') || ss.getSheets()[0];
    var rows  = sheet.getDataRange().getValues();
    if (rows.length < 2) return result;

    var headers = rows[0].map(function(h){ return String(h).trim().toLowerCase(); });
    function hIdx(c){ for(var i=0;i<c.length;i++){var x=headers.indexOf(c[i].toLowerCase());if(x>=0)return x;} return -1; }
    var colId   = hIdx(['id']);
    var colNome = hIdx(['nome','name']);
    var colResp = hIdx(['responsavel','responsável','resp']);
    var colPraz = hIdx(['prazo']);
    var colEnt  = hIdx(['entrega']);
    var colMes  = hIdx(['mes','mês','month']);

    var respNorm = resp.trim().toLowerCase();

    for (var r = 1; r < rows.length; r++) {
      var row     = rows[r];
      var rowResp = colResp >= 0 ? String(row[colResp]||'').trim() : '';
      if (!rowResp || rowResp.toLowerCase() !== respNorm) continue;

      var rowId      = colId   >= 0 ? String(row[colId]  ||'') : '';
      var rowName    = colNome >= 0 ? String(row[colNome] ||'') : '';
      var rowPraz    = colPraz >= 0 ? fmtDateBR(row[colPraz]) : '—';
      var rowEntrega = colEnt  >= 0 ? fmtDateBR(row[colEnt])  : '—';
      var rowMes     = colMes  >= 0 ? String(row[colMes]  ||'') : '';

      // ── Filtro: lista armazenada do e-mail (preferencial) OU janela de 30d (fallback) ──
      var include = false;
      if (storedIds !== null) {
        include = !!(rowId && storedIds[rowId]); // apenas IDs que foram enviados
      } else {
        // fallback: janela de 30d
        if (rowPraz === today) {
          include = true;
        } else if (rowPraz && rowPraz !== '—') {
          var pp = rowPraz.split('/');
          if (pp.length === 3) {
            var dPrazo = new Date(+pp[2], +pp[1]-1, +pp[0]);
            if (!isNaN(dPrazo) && dPrazo < todayDate) {
              var diff = Math.round((todayDate - dPrazo) / 86400000);
              if (diff <= OVERDUE_WINDOW) include = true;
            }
          }
        }
      }

      // ── Status de entrega: Sheets tem precedência; fallback para PropertiesService ──
      // Usa allProps (pré-carregado) em vez de getProperty() individual — evita N round-trips de I/O
      var sheetsDelivered = rowEntrega && rowEntrega !== '—';
      var confRaw = rowId ? (allProps['confirm_' + rowId] || null) : null;
      var confirmed   = sheetsDelivered || !!confRaw;
      var confirmedAt = '';
      if (sheetsDelivered) {
        confirmedAt = rowEntrega;
      } else if (confRaw) {
        try {
          var confObj = JSON.parse(confRaw);
          // Usa confirmedDate (BRT) se disponível; fallback para parse do UTC ISO
          if (confObj.confirmedDate) {
            confirmedAt = confObj.confirmedDate;
          } else {
            var iso = (confObj.confirmedAt||'').slice(0,10).split('-');
            if (iso.length === 3) confirmedAt = iso[2]+'/'+iso[1]+'/'+iso[0];
          }
        } catch(e2) {}
      }

      // Exclui tarefas já entregues EXCETO a tarefa atual (que acabou de ser confirmada)
      // currentId vazio (ex: handleViewStatus sem id) → nenhuma tarefa é "atual"
      var isCurrent = !!(currentId && rowId && rowId === String(currentId));
      if (confirmed && !isCurrent) continue;
      if (!include && !isCurrent) continue;

      result.push({
        id: rowId, name: rowName, prazo: rowPraz, mes: rowMes,
        confirmed: confirmed, confirmedAt: confirmedAt,
        isCurrent: isCurrent
      });
    }

    // Ordena: pendentes primeiro (por prazo asc), depois entregues
    function prazoCmp(p){ var s=(p||'').split('/'); return s.length===3?s[2]+s[1]+s[0]:'99999999'; }
    result.sort(function(a, b) {
      if (!a.confirmed &&  b.confirmed) return -1;
      if ( a.confirmed && !b.confirmed) return  1;
      return prazoCmp(a.prazo) < prazoCmp(b.prazo) ? -1 : prazoCmp(a.prazo) > prazoCmp(b.prazo) ? 1 : 0;
    });

    // Grava no cache (só quando não é confirmação — currentId indica dados recém mudados)
    if (!currentId) {
      try { cache.put(cacheKey, JSON.stringify(result), 120); } catch(e2) {}
    }
  } catch(e) {
    Logger.log('getRespTasksForPanel erro: ' + e.message);
  }
  return result;
}

function handleGetConfirmations(callback) {
  var props  = PropertiesService.getScriptProperties();
  var all    = props.getProperties();
  var result = [];
  var lastSend = null;
  Object.keys(all).forEach(function(k) {
    if (k.indexOf('confirm_') === 0) {
      try { result.push(JSON.parse(all[k])); } catch(e) {}
    }
    if (k === 'last_send_result') {
      try { lastSend = JSON.parse(all[k]); } catch(e) {}
    }
  });
  return jsonpResponse({ ok: true, confirmations: result, lastSend: lastSend }, callback);
}

// Calcula status de entrega dado prazo (DD/MM/YYYY) e data de entrega (DD/MM/YYYY)
function calcDeliveryStatus(prazo, entrega) {
  var pParts = (prazo||'').split('/');
  var eParts = (entrega||'').split('/');
  if (pParts.length !== 3 || eParts.length !== 3) return 'ENTREGUE';
  var dp = new Date(+pParts[2], +pParts[1]-1, +pParts[0]);
  var de = new Date(+eParts[2], +eParts[1]-1, +eParts[0]);
  if (isNaN(dp) || isNaN(de)) return 'ENTREGUE';
  if (de < dp)  return 'ENTREGA ANTECIPADA';
  if (de.getTime() === dp.getTime()) return 'ENTREGUE';
  return 'ENTREGUE COM ATRASO';
}

// Atualiza Entrega + Status + "Atualizado em" em todas as abas (Todos + mês) pelo ID
function confirmDeliveryInSheet(id, today, prazo, month) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var now = new Date().toISOString();
  var status = prazo && prazo !== '—' ? calcDeliveryStatus(prazo, today) : 'ENTREGUE';

  function updateSheet(sheet) {
    if (!sheet) return;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(function(h){ return String(h).trim().toLowerCase(); });
    var iId  = headers.indexOf('id');
    var iEnt = headers.indexOf('entrega');
    var iSt  = headers.indexOf('status');
    var iAtu = headers.indexOf('atualizado em');
    if (iId < 0 || iEnt < 0) return;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    var ids = sheet.getRange(2, iId+1, lastRow-1, 1).getValues();
    for (var r = 0; r < ids.length; r++) {
      if (String(ids[r][0]) === String(id)) {
        var row = r + 2;
        sheet.getRange(row, iEnt+1).setNumberFormat('@').setValue(today);
        if (iSt  >= 0) sheet.getRange(row, iSt+1).setValue(status);
        if (iAtu >= 0) sheet.getRange(row, iAtu+1).setValue(now);
        break;
      }
    }
  }

  updateSheet(ss.getSheetByName('Todos') || ss.getSheets()[0]);
  if (month && MONTH_SHEET[month]) {
    updateSheet(ss.getSheetByName(MONTH_SHEET[month]));
  }
}

// Formata valor de célula de data para DD/MM/YYYY (suporta Date object ou string)
function fmtDateBR(val) {
  if (!val || val === '-' || val === '—') return '—';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'America/Sao_Paulo', 'dd/MM/yyyy');
  }
  var s = String(val).trim();
  return (s === '' || s === '-') ? '—' : s;
}

// ── GET_TASKS — Lê planilha e retorna tarefas ao dashboard ────
var GET_TASKS_CACHE_KEY = 'get_tasks_v1';
function handleGetTasks(callback) {
  if (!SPREADSHEET_ID) {
    return jsonpResponse({ ok: false, error: 'SPREADSHEET_ID nao configurado.' }, callback);
  }

  // Cache de 20s: evita reler a planilha a cada polling do dashboard (30s)
  var cache = CacheService.getScriptCache();
  var cached = cache.get(GET_TASKS_CACHE_KEY);
  if (cached) {
    var obj = JSON.parse(cached);
    obj.ts = new Date().toISOString(); // atualiza timestamp sem reler planilha
    return jsonpResponse(obj, callback);
  }

  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Todos') || ss.getSheets()[0];
    var data  = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return jsonpResponse({ ok: true, tasks: [] }, callback);
    }

    // Mapeamento flexível: aceita nomes com e sem acento
    var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
    function col(candidates) {
      for (var i = 0; i < candidates.length; i++) {
        var idx = headers.indexOf(candidates[i].toLowerCase());
        if (idx >= 0) return idx;
      }
      return -1;
    }
    var iID      = col(['ID','id']);
    var iNome    = col(['Nome','name']);
    var iNota    = col(['Nota','note']);
    var iResp    = col(['Responsavel','Responsável','resp']);
    var iDest    = col(['Destinatario','Destinatário','dest']);
    var iEmail   = col(['Email','E-mail','email']);
    var iPrazo   = col(['Prazo','prazo']);
    var iEntrega = col(['Entrega','entrega']);
    var iMes     = col(['Mes','Mês','month']);
    // Status NÃO é lido do Sheets — o dashboard sempre calcula localmente

    var tasks = [];
    for (var r = 1; r < data.length; r++) {
      var row  = data[r];
      var nome = iNome >= 0 ? String(row[iNome] || '').trim() : '';
      if (!nome) continue;
      // Linhas sem ID válido são ignoradas — evita colisão com IDs do dashboard
      var rowId = iID >= 0 ? Number(row[iID]) : NaN;
      if (isNaN(rowId) || rowId <= 0) continue;
      tasks.push({
        id:      rowId,
        name:    nome,
        note:    iNota    >= 0 ? String(row[iNota]    || '') : '',
        resp:    iResp    >= 0 ? String(row[iResp]    || '—') : '—',
        dest:    iDest    >= 0 ? String(row[iDest]    || '—') : '—',
        email:   iEmail   >= 0 ? String(row[iEmail]   || '') : '',
        prazo:   iPrazo   >= 0 ? fmtDateBR(row[iPrazo])   : '—',
        entrega: iEntrega >= 0 ? fmtDateBR(row[iEntrega]) : '—',
        month:   iMes     >= 0 ? String(row[iMes]     || '') : '',
      });
    }
    var result = { ok: true, tasks: tasks, count: tasks.length, ts: new Date().toISOString() };
    try { cache.put(GET_TASKS_CACHE_KEY, JSON.stringify(result), 20); } catch(e2) {}
    return jsonpResponse(result, callback);
  } catch (err) {
    Logger.log('handleGetTasks error: ' + err.message + ' | stack: ' + err.stack);
    return jsonpResponse({ ok: false, error: String(err.message) }, callback);
  }
}

function invalidateTasksCache() {
  try { CacheService.getScriptCache().remove(GET_TASKS_CACHE_KEY); } catch(e) {}
}

// viewOnly=true → página de status sem ação de confirmação (link "Ver minhas entregas")
function buildConfirmationPage(name, prazo, alreadyDone, resp, allRespTasks, viewOnly) {
  var now  = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');

  // ── Bloco de confirmação (modo confirmar) ──
  var confirmBlock = '';
  if (!viewOnly) {
    var confirmBg  = alreadyDone ? '#fff8ec' : '#e8f5ee';
    var confirmClr = alreadyDone ? '#b45300' : '#1a7a4a';
    var confirmIcon = alreadyDone ? '⚠️' : '✅';
    var confirmMsg  = alreadyDone
      ? 'Esta entrega já havia sido registrada anteriormente.'
      : 'Confirmação registrada em <strong>' + now + '</strong>.<br>O dashboard será atualizado automaticamente.';
    confirmBlock =
      '<div style="background:' + confirmBg + ';border-radius:10px;padding:16px 20px;margin-bottom:20px">'
      + '<div style="display:flex;align-items:flex-start;gap:12px">'
      +   '<div style="font-size:28px;line-height:1">' + confirmIcon + '</div>'
      +   '<div style="flex:1">'
      +     '<div style="font-size:14px;font-weight:700;color:' + confirmClr + ';margin-bottom:4px">'
      +       (alreadyDone ? 'Já confirmado' : 'Entrega confirmada!')
      +     '</div>'
      +     '<div style="font-size:13px;font-weight:700;color:#0a1e45;margin-bottom:6px">' + esc(name) + '</div>'
      +     (prazo ? '<div style="font-size:12px;color:#8096b8">📅 Prazo: ' + esc(prazo) + '</div>' : '')
      +     '<div style="font-size:12px;color:' + confirmClr + ';margin-top:6px">' + confirmMsg + '</div>'
      +   '</div>'
      + '</div>'
      + '</div>';
  }

  // ── Tabela de tarefas do responsável ──
  var totalConf = 0, totalPend = 0;
  var tableRows = '';
  if (allRespTasks && allRespTasks.length > 0) {
    totalConf = allRespTasks.filter(function(t){ return t.confirmed; }).length;
    totalPend = allRespTasks.length - totalConf;

    tableRows = allRespTasks.map(function(t, idx) {
      var isCurrent  = t.isCurrent;
      var conf       = t.confirmed;
      var rowBg      = isCurrent ? '#f0fbf4' : (idx % 2 === 0 ? '#ffffff' : '#f7faff');
      var rowBorder  = isCurrent ? '2px solid #1a7a4a' : 'none';
      var statusBg   = conf ? '#e6f5ee' : '#fff8ec';
      var statusClr  = conf ? '#1a7a4a' : '#b45300';
      var statusIcon = conf ? '✅' : '⏳';
      var statusTxt  = conf
        ? ('Entregue' + (t.confirmedAt ? '<br><span style="font-size:10px;font-weight:400">' + t.confirmedAt + '</span>' : ''))
        : 'Pendente';
      var newBadge = (isCurrent && !viewOnly && !alreadyDone)
        ? '&nbsp;<span style="background:#1a7a4a;color:#fff;font-size:9px;padding:1px 7px;border-radius:99px;vertical-align:middle;white-space:nowrap">agora</span>'
        : '';
      return '<tr style="background:' + rowBg + ';outline:' + rowBorder + '">'
        + '<td style="padding:10px 14px;font-size:13px;font-weight:' + (isCurrent ? '700' : '500') + ';color:#0a1e45;word-break:break-word">'
        +   esc(t.name) + newBadge
        + '</td>'
        + '<td style="padding:10px 14px;font-size:12px;color:#3a5080;white-space:nowrap;text-align:center">'
        +   (t.prazo && t.prazo !== '—' ? t.prazo : '—')
        + '</td>'
        + '<td style="padding:10px 14px;text-align:center">'
        +   '<span style="display:inline-flex;align-items:center;gap:4px;background:' + statusBg + ';color:' + statusClr + ';border-radius:99px;padding:4px 12px;font-size:11px;font-weight:700;white-space:nowrap">'
        +     statusIcon + '&nbsp;' + statusTxt
        +   '</span>'
        + '</td>'
        + '</tr>';
    }).join('');
  }

  var hasPanel = tableRows !== '';
  var summaryHtml = '';
  if (hasPanel) {
    summaryHtml =
      '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px">'
      + '<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.9px;color:#8096b8">'
      +   '📋 Suas tarefas de hoje e em atraso — <strong style="color:#0a1e45">' + esc(resp) + '</strong>'
      + '</div>'
      + '<div style="font-size:11px;color:#8096b8">'
      +   '<span style="color:#1a7a4a;font-weight:700">' + totalConf + ' entregue' + (totalConf !== 1 ? 's' : '') + '</span>'
      +   (totalPend > 0 ? ' &nbsp;·&nbsp; <span style="color:#b45300;font-weight:700">' + totalPend + ' pendente' + (totalPend !== 1 ? 's' : '') + '</span>' : '')
      + '</div>'
      + '</div>'
      + '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid #dce8f5;border-radius:10px;overflow:hidden">'
      + '<thead>'
      + '<tr style="background:#1352b8">'
      + '<th style="padding:9px 14px;text-align:left;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.5px">Tarefa</th>'
      + '<th style="padding:9px 14px;text-align:center;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.5px;white-space:nowrap;width:100px">Prazo</th>'
      + '<th style="padding:9px 14px;text-align:center;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.5px;width:130px">Status</th>'
      + '</tr>'
      + '</thead>'
      + '<tbody>' + tableRows + '</tbody>'
      + '</table>';
  } else if (!viewOnly) {
    // nenhuma tarefa no painel mas estamos em modo confirmar — não mostra tabela vazia
    summaryHtml = '';
  } else {
    summaryHtml = '<div style="text-align:center;padding:28px;color:#8096b8;font-size:13px">Nenhuma tarefa com prazo hoje ou em atraso recente.</div>';
  }

  var pageTitle = viewOnly ? 'Minhas Entregas' : 'Entrega Confirmada!';
  var pageSubtitle = viewOnly ? 'Acompanhamento de Entregas' : 'Confirmação de Entrega';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<style>'
    + '*{box-sizing:border-box;margin:0;padding:0}'
    + 'body{font-family:Arial,sans-serif;background:#eef4fb;display:flex;justify-content:center;min-height:100vh;padding:24px 16px}'
    + '.wrap{max-width:960px;width:100%;align-self:flex-start}'
    + '.card{background:#fff;border-radius:14px;box-shadow:0 4px 24px rgba(13,45,110,.13);overflow:hidden;margin-bottom:16px}'
    + '.hdr{background:linear-gradient(135deg,#0d2d6e,#1352b8);padding:22px 28px;display:flex;align-items:center;gap:16px}'
    + '.hdr-text h1{color:#fff;font-size:17px;font-weight:700;margin-bottom:2px}'
    + '.hdr-text p{color:rgba(255,255,255,.6);font-size:12px}'
    + '.hdr-icon{font-size:34px;line-height:1;flex-shrink:0}'
    + '.body{padding:24px 28px}'
    + '.footer{padding:12px 28px;background:#f4f8fd;border-top:1px solid #dce8f5;text-align:center;font-size:11px;color:#8096b8}'
    + 'table{border-collapse:collapse}'
    + 'tr:last-child td{border-bottom:none!important}'
    + 'td{border-bottom:1px solid #eaf2fc}'
    + '@media(max-width:600px){.hdr{padding:16px 18px}.body{padding:16px 18px}td,th{padding:8px 10px!important}}'
    + '</style></head><body>'
    + '<div class="wrap">'
    + '<div class="card">'
    +   '<div class="hdr"><div class="hdr-icon">' + (viewOnly ? '📋' : '✅') + '</div>'
    +   '<div class="hdr-text"><h1>' + pageTitle + '</h1><p>Cronograma Mensal · Mabu Hospitalidade &amp; Entretenimento · ' + pageSubtitle + '</p></div></div>'
    +   '<div class="body">'
    +     confirmBlock
    +     summaryHtml
    +   '</div>'
    +   '<div class="footer">Mensagem automática · Cronograma Mensal · Mabu Hospitalidade</div>'
    + '</div>'
    + '</div></body></html>';
}

function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action || '';
    var result;

    // Valida segredo da API em todas as ações destrutivas/sensíveis
    // O segredo é configurado via setupConfig() e salvo no PropertiesService
    var PROTECTED_ACTIONS = ['sync','update_task','delete_task','send_all','send_now','send_summary','send_test'];
    if (API_SECRET && PROTECTED_ACTIONS.indexOf(action) >= 0 && body.secret !== API_SECRET) {
      Logger.log('doPost: acesso negado — segredo inválido para ação=' + action);
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Não autorizado' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if      (action === 'send_test')    result = handleSendTest(body);
    else if (action === 'send_all')     result = handleSendAll(body);
    else if (action === 'send_now')     result = handleSendNow(body);
    else if (action === 'send_summary') result = handleSendSummary(body);
    else if (action === 'sync')         result = handleSync(body);
    else if (action === 'update_task')  result = handleUpdateTask(body);
    else if (action === 'delete_task')  result = handleDeleteTask(body);
    else                                result = { ok: false, error: 'Ação desconhecida: ' + action };

    return jsonResponse(result);

  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return jsonResponse({ ok: false, error: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── SEND_TEST ────────────────────────────────────────────────
// Disparado pelo botão "Enviar teste" no painel de configuração

function handleSendTest(body) {
  var to   = ADMIN_EMAIL; // Sempre para o admin — ignora body.to para evitar spam abuse
  var task = body.task || {};

  var subject = '[Cronograma Mensal] Teste de disparo automático — ' + Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  var html    = buildEmailHtml([task], 'Teste de Envio Automático', false);

  MailApp.sendEmail({
    to:       to,
    name:     EMAIL_FROM_NAME,
    subject:  subject,
    htmlBody: html,
  });

  Logger.log('send_test → ' + to);
  return { ok: true, message: 'E-mail de teste enviado para ' + to };
}

// ── SEND_NOW ─────────────────────────────────────────────────
// Disparado pelo botão "Disparar e-mails" no modal de notificações

// Envia geral + individuais agrupados por destinatário
// Aceita formato novo: { groups: [{email, tasks:[...]}, ...] }
// Aceita também formato legado: { tasks: [...] } (converte para grupos)
function handleSendAll(body) {
  var groups = body.groups || [];
  var today  = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');

  // Compatibilidade com formato legado (tasks flat)
  if (!groups.length && (body.tasks || []).length) {
    var grouped = {};
    (body.tasks || []).forEach(function(t) {
      var key = t.email || ADMIN_EMAIL;
      if (!grouped[key]) grouped[key] = [];
      grouped[key].push(t);
    });
    Object.keys(grouped).forEach(function(email) {
      groups.push({ email: email, tasks: grouped[email] });
    });
  }

  if (!groups.length) return { ok: true, skipped: true };

  // Todas as tarefas de todos os grupos para o e-mail geral
  var allTasks = [];
  groups.forEach(function(g) { allTasks = allTasks.concat(g.tasks || []); });

  var sent = 0, errs = [];

  // 1. E-MAIL GERAL — admin recebe resumo com TODAS as tarefas (hoje + atrasadas, sem botão confirmar)
  try {
    var subjGeral = '[Cronograma Mensal] Resumo geral — ' + today +
      ' (' + allTasks.length + ' tarefa' + (allTasks.length !== 1 ? 's' : '') + ')';
    MailApp.sendEmail({
      to:       ADMIN_EMAIL,
      name:     EMAIL_FROM_NAME,
      subject:  subjGeral,
      htmlBody: buildSummaryEmail(allTasks, today),
    });
    Logger.log('send_all geral OK → ' + allTasks.length + ' tarefas');
  } catch(err) {
    errs.push('geral: ' + err.message);
    Logger.log('send_all geral ERRO → ' + err.message);
  }

  // 2. E-MAILS INDIVIDUAIS — 1 e-mail por destinatário com TODAS as suas tarefas + botão confirmar
  groups.forEach(function(g) {
    if (!g.email || !(g.tasks || []).length) return;
    try {
      var atrasadas     = g.tasks.filter(function(t){ return t.status === 'ATRASADO'; });
      var totalCount    = g.tasks.length;
      var atrasadasInfo = atrasadas.length ? ' · ' + atrasadas.length + ' em atraso' : '';
      var subject = '[Cronograma Mensal] ' + today + ' — ' +
        totalCount + ' tarefa' + (totalCount !== 1 ? 's' : '') + atrasadasInfo;
      var html = buildEmailHtml(g.tasks, 'Suas Tarefas — Hoje e Em Atraso', true);
      MailApp.sendEmail({
        to:      g.email,
        bcc:     g.email !== ADMIN_EMAIL ? ADMIN_EMAIL : '',
        name:    EMAIL_FROM_NAME,
        subject: subject,
        htmlBody: html,
      });
      sent++;
      Logger.log('send_all individual OK → ' + g.email + ' | ' + totalCount + ' tarefas');
    } catch(err) {
      errs.push(g.email + ': ' + err.message);
      Logger.log('send_all individual ERRO → ' + g.email + ' | ' + err.message);
    }
  });

  // Armazena lista de tarefas por responsável para "Minhas Entregas"
  var props2 = PropertiesService.getScriptProperties();
  var respMap = {};
  groups.forEach(function(g) {
    (g.tasks || []).forEach(function(t) {
      if (!t.id) return; // ignora tarefas sem ID válido
      var r = (t.resp || '').trim();
      if (!r || r === '—') return;
      if (!respMap[r]) respMap[r] = [];
      respMap[r].push({id: String(t.id), name: t.name||'', prazo: t.prazo||''});
    });
  });
  Object.keys(respMap).forEach(function(r) {
    var rKey = 'resp_sent_' + r.trim().toLowerCase()
      .replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,'').slice(0,60);
    var data = {resp: r, tasks: respMap[r], sentAt: new Date().toISOString()};
    try { props2.setProperty(rKey, JSON.stringify(data)); } catch(e) {
      Logger.log('resp_sent store error: ' + e.message);
    }
  });
  // Grava resultado para que o dashboard possa verificar via JSONP (no-cors não lê a resposta)
  try {
    props2.setProperty('last_send_result', JSON.stringify({
      ok: errs.length === 0, sent: sent, errors: errs, ts: new Date().toISOString()
    }));
  } catch(e) {}

  return { ok: errs.length === 0, sent: sent, errors: errs };
}

// Template simplificado do e-mail geral (evita falha com múltiplas tarefas)
function buildSummaryEmail(tasks, today) {
  return buildEmailHtml(tasks, 'Resumo Geral · Tarefas com Prazo Hoje', false, true);
}

// E-mail individual para o responsável (com botão de confirmação)
function handleSendNow(body) {
  var tasks = body.tasks || [];
  var sent  = 0;
  var errs  = [];

  tasks.forEach(function(t) {
    if (!t.email) return;
    try {
      var subject = '[Cronograma Mensal] Sua tarefa vence hoje: ' + t.name;
      var html    = buildEmailHtml([t], 'Sua Tarefa com Prazo Hoje', true);
      MailApp.sendEmail({
        to:       t.email,
        bcc:      t.email !== ADMIN_EMAIL ? ADMIN_EMAIL : '',
        name:     EMAIL_FROM_NAME,
        subject:  subject,
        htmlBody: html,
      });
      sent++;
      Logger.log('send_now individual → ' + t.email + ' | ' + t.name);
    } catch (err) {
      errs.push(t.email + ': ' + err.message);
      Logger.log('send_now ERROR → ' + t.email + ' | ' + err.message);
    }
  });

  // Armazena lista de tarefas por responsável para "Minhas Entregas"
  // (mesmo tratamento de handleSendAll — send_now também gera link pessoal)
  var props2 = PropertiesService.getScriptProperties();
  var respMap = {};
  tasks.forEach(function(t) {
    if (!t.id) return;
    var r = (t.resp || '').trim();
    if (!r || r === '—') return;
    if (!respMap[r]) respMap[r] = [];
    respMap[r].push({id: String(t.id), name: t.name||'', prazo: t.prazo||''});
  });
  Object.keys(respMap).forEach(function(r) {
    var rKey = 'resp_sent_' + r.trim().toLowerCase()
      .replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,'').slice(0,60);
    // Merge com tarefas já existentes — send_now não deve apagar envios anteriores
    var existingTasks = [];
    try {
      var prev = props2.getProperty(rKey);
      if (prev) existingTasks = JSON.parse(prev).tasks || [];
    } catch(e) {}
    var existingIds = existingTasks.map(function(t){ return String(t.id); });
    respMap[r].forEach(function(nt) {
      if (existingIds.indexOf(String(nt.id)) === -1) existingTasks.push(nt);
    });
    var data = {resp: r, tasks: existingTasks, sentAt: new Date().toISOString()};
    try { props2.setProperty(rKey, JSON.stringify(data)); } catch(e) {
      Logger.log('resp_sent store error (send_now): ' + e.message);
    }
  });
  // Grava resultado para verificação via JSONP pelo dashboard
  try {
    props2.setProperty('last_send_result', JSON.stringify({
      ok: errs.length === 0, sent: sent, errors: errs, ts: new Date().toISOString()
    }));
  } catch(e) {}

  return { ok: errs.length === 0, sent: sent, errors: errs };
}

// E-mail geral para o admin com todas as tarefas (sem botão)
function handleSendSummary(body) {
  var tasks = body.tasks || [];
  if (!tasks.length) return { ok: true, skipped: true };

  try {
    var today   = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
    var subject = '[Cronograma Mensal] Resumo das tarefas do dia — ' + today;
    var html    = buildEmailHtml(tasks, 'Resumo Geral · Tarefas com Prazo Hoje', false, true);
    MailApp.sendEmail({
      to:       ADMIN_EMAIL,
      name:     EMAIL_FROM_NAME,
      subject:  subject,
      htmlBody: html,
    });
    Logger.log('send_summary → ' + ADMIN_EMAIL + ' | ' + tasks.length + ' tarefas');
    return { ok: true, sent: tasks.length };
  } catch (err) {
    Logger.log('send_summary ERROR → ' + err.message);
    return { ok: false, error: err.message };
  }
}

// ── SYNC ─────────────────────────────────────────────────────
// Salva lote de tarefas editadas na planilha de log (opcional)

function handleSync(body) {
  var tasks = body.tasks || [];
  if (!SPREADSHEET_ID || !tasks.length) return { ok: true, skipped: true };

  try {
    logTasksToSheet(tasks);
    invalidateTasksCache();
    return { ok: true, saved: tasks.length };
  } catch (err) {
    Logger.log('sync error: ' + err.message);
    return { ok: false, error: err.message };
  }
}

// ── DELETE_TASK ──────────────────────────────────────────────
// Remove uma linha da planilha pelo ID

function handleDeleteTask(body) {
  var id    = body.id;
  var month = body.month || '';
  // Valida que id é um inteiro positivo (evita enumeração e abuso)
  var numId = Number(id);
  if (!id || isNaN(numId) || numId <= 0 || numId !== Math.floor(numId) || numId > 2147483647) {
    Logger.log('delete_task: ID inválido recebido: ' + JSON.stringify(id));
    return { ok: false, error: 'ID inválido' };
  }
  if (!SPREADSHEET_ID || id === undefined || id === null) return { ok: true, skipped: true };
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // Remove da aba Todos
    var sheetTodos = ss.getSheetByName('Todos') || ss.getSheets()[0];
    deleteRowById(sheetTodos, id);

    // Remove da aba do mês
    var mesLabel = month ? (MONTH_SHEET[month] || '') : '';
    if (mesLabel) {
      var sheetMes = ss.getSheetByName(mesLabel);
      if (sheetMes) deleteRowById(sheetMes, id);
    } else if (month) {
      Logger.log('WARNING: Mês inválido "' + month + '" para tarefa id=' + id);
    }

    Logger.log('delete_task OK: id=' + id + ' mes=' + month);
    invalidateTasksCache();
    return { ok: true };
  } catch (err) {
    Logger.log('delete_task error: ' + err.message);
    return { ok: false, error: err.message };
  }
}

// ── UPDATE_TASK ──────────────────────────────────────────────
// Atualiza uma única tarefa na planilha de log

function handleUpdateTask(body) {
  var task = body.task;
  if (!SPREADSHEET_ID || !task) return { ok: true, skipped: true };

  try {
    logTasksToSheet([task]);
    invalidateTasksCache();
    return { ok: true };
  } catch (err) {
    Logger.log('update_task error: ' + err.message);
    return { ok: false, error: err.message };
  }
}

// ── DISPARO DIÁRIO AUTOMÁTICO (Trigger) ──────────────────────
// Execute createDailyTrigger() UMA VEZ no editor para agendar

function createDailyTrigger() {
  // Remove triggers anteriores para evitar duplicatas
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyEmailJob') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('dailyEmailJob')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone('America/Sao_Paulo')
    .create();
  Logger.log('Trigger diário criado: dailyEmailJob às 8h (Brasília)');
}

// ── JOB DIÁRIO ───────────────────────────────────────────────
// Roda às 8h automaticamente; pode ser testado manualmente

function dailyEmailJob() {
  if (!SPREADSHEET_ID) {
    Logger.log('dailyEmailJob: SPREADSHEET_ID não configurado');
    return;
  }
  try {
    var props = PropertiesService.getScriptProperties();
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Todos') || ss.getSheets()[0];
    var data  = sheet.getDataRange().getValues();
    if (data.length < 2) { Logger.log('dailyEmailJob: planilha vazia'); return; }

    var headers = data[0].map(function(h){ return String(h).trim().toLowerCase(); });
    function col(c){ for(var i=0;i<c.length;i++){var x=headers.indexOf(c[i].toLowerCase());if(x>=0)return x;} return -1; }
    var iID=col(['id']), iNome=col(['nome','name']), iNota=col(['nota','note']),
        iResp=col(['responsavel','responsável','resp']), iDest=col(['destinatario','destinatário','dest']),
        iEmail=col(['email','e-mail']), iPrazo=col(['prazo']), iEntrega=col(['entrega']), iMes=col(['mes','mês','month']);

    var today = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
    var tParts = today.split('/');
    var todayDate = new Date(+tParts[2], +tParts[1]-1, +tParts[0]);
    var OVERDUE_WINDOW = 30; // mesmo critério do dashboard

    var tasks = [];
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var nome = iNome >= 0 ? String(row[iNome]||'').trim() : '';
      if (!nome) continue;
      var rowId = iID >= 0 ? Number(row[iID]) : NaN;
      if (isNaN(rowId) || rowId <= 0) continue;
      var prazo   = iPrazo   >= 0 ? fmtDateBR(row[iPrazo])   : '—';
      var entrega = iEntrega >= 0 ? fmtDateBR(row[iEntrega]) : '—';

      // Calcula status localmente (não lê do Sheets — pode estar desatualizado)
      var status;
      if (entrega && entrega !== '—') {
        status = calcDeliveryStatus(prazo, entrega);
      } else if (prazo && prazo !== '—') {
        var pp = prazo.split('/');
        if (pp.length === 3) {
          var dPrazo = new Date(+pp[2], +pp[1]-1, +pp[0]);
          status = isNaN(dPrazo) ? 'NO PRAZO' : (dPrazo < todayDate ? 'ATRASADO' : 'NO PRAZO');
        } else { status = 'NO PRAZO'; }
      } else { status = 'NO PRAZO'; }

      // Filtra: só tarefas com prazo hoje ou atrasadas dentro da janela
      var delivered = (status==='ENTREGUE'||status==='ENTREGA ANTECIPADA'||status==='ENTREGUE COM ATRASO');
      // Verifica também confirmações via link de e-mail (PropertiesService) que ainda não
      // chegaram ao Sheets (ex: confirmDeliveryInSheet falhou ou sync ainda pendente)
      if (!delivered && rowId > 0) {
        if (props.getProperty('confirm_' + rowId)) delivered = true;
      }
      if (delivered) continue;
      var include = false;
      if (prazo === today) { include = true; }
      else if (status === 'ATRASADO') {
        var pp2 = prazo.split('/');
        if (pp2.length === 3) {
          var dP2 = new Date(+pp2[2], +pp2[1]-1, +pp2[0]);
          if (!isNaN(dP2) && Math.round((todayDate - dP2) / 86400000) <= OVERDUE_WINDOW) include = true;
        }
      }
      if (!include) continue;

      tasks.push({
        id: rowId, name: nome,
        note:  iNota  >= 0 ? String(row[iNota] ||'')  : '',
        resp:  iResp  >= 0 ? String(row[iResp] ||'—') : '—',
        dest:  iDest  >= 0 ? String(row[iDest] ||'—') : '—',
        email: iEmail >= 0 ? String(row[iEmail]||'')  : '',
        prazo: prazo, entrega: entrega, status: status,
        month: iMes   >= 0 ? String(row[iMes]  ||'')  : '',
      });
    }

    if (!tasks.length) {
      Logger.log('dailyEmailJob: nenhuma tarefa com prazo hoje ou atrasada');
      return;
    }

    // Monta grupos por e-mail (mesmo critério do dashboard)
    var grouped = {};
    tasks.forEach(function(t) {
      var email = (t.email && t.email.trim()) ? t.email : ADMIN_EMAIL;
      if (!grouped[email]) grouped[email] = [];
      grouped[email].push(t);
    });
    var groups = Object.keys(grouped).map(function(email) {
      return { email: email, tasks: grouped[email] };
    });

    var result = handleSendAll({ groups: groups });
    Logger.log('dailyEmailJob: ' + tasks.length + ' tarefas em ' + groups.length + ' grupos | result: ' + JSON.stringify(result));
  } catch(e) {
    Logger.log('dailyEmailJob erro: ' + e.message + ' | stack: ' + (e.stack||''));
    MailApp.sendEmail({ to: ADMIN_EMAIL, name: EMAIL_FROM_NAME,
      subject: '[Cronograma Mensal] Erro no disparo diário — ' + new Date().toISOString(),
      htmlBody: '<pre>' + e.message + '\n' + (e.stack||'') + '</pre>' });
  }
}

// ── TESTE MANUAL ─────────────────────────────────────────────
// Execute esta função diretamente no editor para testar

function sendTestEmail() {
  handleSendTest({
    to:   ADMIN_EMAIL,
    task: {
      name:    'TESTE — Cronograma Mensal Mabu: verificação de disparo automático',
      note:    'Este é um e-mail de teste do sistema Cronograma Mensal.',
      resp:    'Elias Luan Probst Schlender',
      dest:    'Contabilidade',
      email:   ADMIN_EMAIL,
      prazo:   Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy'),
      entrega: '—',
      status:  'NO PRAZO',
      month:   'mar',
    },
  });
  Logger.log('sendTestEmail() executado com sucesso.');
}

// ── LOG EM PLANILHA (opcional) ────────────────────────────────

// Colunas de todas as abas
var SHEET_COLS = ['ID','Nome','Nota','Responsavel','Destinatario','Email','Prazo','Entrega','Status','Mes','Atualizado em'];

// Mapa de chave de mês → nome da aba
var MONTH_SHEET = {
  'jan':'Janeiro','fev':'Fevereiro','mar':'Março','abr':'Abril',
  'mai':'Maio','jun':'Junho','jul':'Julho','ago':'Agosto',
  'set':'Setembro','out':'Outubro','nov':'Novembro','dez':'Dezembro'
};

// Upsert de uma linha em uma aba pelo ID (col 0)
// Aplica formato @text nas colunas de data (7=Prazo, 8=Entrega) ANTES de escrever
// para impedir que o Sheets converta automaticamente DD/MM/YYYY para Date object
function upsertRow(sheet, id, row) {
  var data  = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) === String(id)) {
      sheet.getRange(r + 1, 7, 1, 2).setNumberFormat('@');
      sheet.getRange(r + 1, 1, 1, SHEET_COLS.length).setValues([row]);
      return;
    }
  }
  // Nova linha: pré-formata antes de appendRow
  var newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 7, 1, 2).setNumberFormat('@');
  sheet.getRange(newRow, 1, 1, SHEET_COLS.length).setValues([row]);
}

// Remove linha de uma aba pelo ID (col 0)
function deleteRowById(sheet, id) {
  var data = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) === String(id)) {
      sheet.deleteRow(r + 1);
      return;
    }
  }
}

// Helper: garante que a aba existe, tem cabeçalho e que as colunas de data
// estão sempre no formato @text (impede auto-conversão de DD/MM/YYYY pelo Sheets)
function ensureSheet(ss, name) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(SHEET_COLS);
    sh.getRange(1, 1, 1, SHEET_COLS.length).setFontWeight('bold');
  }
  // Aplica @text em toda a coluna de Prazo (7) e Entrega (8) — novo e existente
  var maxRows = sh.getMaxRows();
  if (maxRows > 1) sh.getRange(2, 7, maxRows - 1, 2).setNumberFormat('@');
  return sh;
}

function logTasksToSheet(tasks) {
  if (!SPREADSHEET_ID) return;
  var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  var now = new Date().toISOString();

  // Prepara abas necessárias UMA vez antes de iterar
  var sheetTodos = ensureSheet(ss, 'Todos');
  var mesSheets  = {};

  tasks.forEach(function(t) {
    var mesLabel = t.month ? (MONTH_SHEET[t.month] || '') : '';
    if (mesLabel && !mesSheets[mesLabel]) {
      mesSheets[mesLabel] = ensureSheet(ss, mesLabel);
    }

    var row = [
      t.id, t.name||'', t.note||'', t.resp||'—', t.dest||'—',
      t.email||'', t.prazo||'—', t.entrega||'—', t.status||'NO PRAZO', t.month||'', now
    ];

    upsertRow(sheetTodos, t.id, row);
    if (mesLabel && mesSheets[mesLabel]) {
      upsertRow(mesSheets[mesLabel], t.id, row);
    }
  });
}

// ── TEMPLATE HTML DO E-MAIL ──────────────────────────────────

function buildEmailHtml(tasks, titulo, showConfirm, compact) {
  compact = compact === true;
  var today  = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
  var gasUrl = GAS_WEB_APP_URL;

  var statusMap = {
    'ATRASADO':            { color:'#c0182c', bg:'#fde8eb', label:'Atrasado' },
    'ENTREGUE COM ATRASO': { color:'#c05000', bg:'#fff4e0', label:'Entregue c/ Atraso' },
    'NO PRAZO':            { color:'#1352b8', bg:'#e8f2fc', label:'No Prazo' },
    'ENTREGUE':            { color:'#1a7a4a', bg:'#e6f5ee', label:'Entregue' },
    'ENTREGA ANTECIPADA':  { color:'#0550ae', bg:'#e0f0ff', label:'Antecipada' },
  };

  // Link único "Ver minhas entregas" — usa a primeira tarefa com responsável válido
  var firstTaskWithResp = null;
  for (var fi = 0; fi < tasks.length; fi++) {
    if (tasks[fi].resp && tasks[fi].resp !== '—') { firstTaskWithResp = tasks[fi]; break; }
  }
  var singleStatusLink = (showConfirm && gasUrl && firstTaskWithResp)
    ? gasUrl + '?action=status&id=' + encodeURIComponent(firstTaskWithResp.id) + '&resp=' + encodeURIComponent(firstTaskWithResp.resp || '') + '&name=' + encodeURIComponent(firstTaskWithResp.name || '') + '&prazo=' + encodeURIComponent(firstTaskWithResp.prazo || '') + '&month=' + encodeURIComponent(firstTaskWithResp.month || '')
    : '';

  var rows = tasks.map(function(t, idx) {
    var st          = statusMap[t.status] || { color:'#3a5080', bg:'#f0f6ff', label: t.status || '—' };
    var isDelivered = (t.status === 'ENTREGUE' || t.status === 'ENTREGA ANTECIPADA' || t.status === 'ENTREGUE COM ATRASO');
    var baseParams = '&id=' + encodeURIComponent(t.id) + '&prazo=' + encodeURIComponent(t.prazo || '') + '&name=' + encodeURIComponent(t.name || '') + '&resp=' + encodeURIComponent(t.resp || '') + '&month=' + encodeURIComponent(t.month || '');
    var confirmLink = (showConfirm && gasUrl && !isDelivered)
      ? gasUrl + '?action=confirm' + baseParams
      : '';
    var rowBg = idx % 2 === 0 ? '#ffffff' : '#f7faff';

    var pad = compact ? '7px 10px' : '11px 12px';
    return '<tr style="background:' + rowBg + '">'
      + '<td style="padding:' + pad + ';border-bottom:1px solid #eaf2fc;vertical-align:middle;word-break:break-word;overflow-wrap:break-word">'
      +   '<div style="font-size:' + (compact ? '12' : '13') + 'px;font-weight:700;color:#0a1e45;line-height:1.4">' + esc(t.name) + '</div>'
      +   (!compact && t.note ? '<div style="font-size:11px;color:#8096b8;margin-top:2px;font-style:italic;line-height:1.3">' + esc(t.note) + '</div>' : '')
      + '</td>'
      + '<td style="padding:' + pad + ';border-bottom:1px solid #eaf2fc;font-size:12px;color:#3a5080;vertical-align:middle;word-break:break-word">' + esc(t.resp || '—') + '</td>'
      + '<td style="padding:' + pad + ';border-bottom:1px solid #eaf2fc;font-size:12px;color:#3a5080;vertical-align:middle;word-break:break-word">' + esc(t.dest || '—') + '</td>'
      + '<td style="padding:' + pad + ';border-bottom:1px solid #eaf2fc;font-size:12px;color:#3a5080;vertical-align:middle;text-align:center;white-space:nowrap">' + esc(t.prazo || '—') + '</td>'
      + '<td style="padding:' + pad + ';border-bottom:1px solid #eaf2fc;vertical-align:middle;text-align:center">'
      +   '<span style="display:inline-block;background:' + st.bg + ';color:' + st.color + ';border-radius:99px;padding:3px 8px;font-size:11px;font-weight:700;white-space:nowrap">' + st.label + '</span>'
      + '</td>'
      + (showConfirm
        ? '<td style="padding:' + pad + ';border-bottom:1px solid #eaf2fc;vertical-align:middle;text-align:center">'
          + (confirmLink
            ? '<a href="' + confirmLink + '" style="display:inline-block;background:#1a7a4a;color:#ffffff;text-decoration:none;font-size:11px;font-weight:700;padding:7px 12px;border-radius:6px;white-space:nowrap">✅ Confirmar</a>'
            : (isDelivered ? '<span style="color:#1a7a4a;font-size:12px;font-weight:700">✅</span>' : ''))
          + '</td>'
        : '')
      + '</tr>';
  }).join('');


  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>'
    + '<body style="margin:0;padding:0;background:#eef4fb;font-family:Arial,sans-serif">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#eef4fb;padding:28px 0">'
    + '<tr><td align="center">'
    + '<table width="980" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:14px;overflow:hidden;box-shadow:0 4px 24px rgba(13,45,110,.12)">'

    // Cabeçalho — título à esquerda, botão "Ver minhas entregas" à direita
    + '<tr><td style="background:linear-gradient(135deg,#0d2d6e 0%,#1a52b8 100%);padding:24px 28px">'
    + '<table width="100%" cellpadding="0" cellspacing="0"><tr>'
    +   '<td style="vertical-align:middle">'
    +     '<div style="font-size:11px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:2px;margin-bottom:5px">Cronograma Mensal · Mabu Hospitalidade</div>'
    +     '<div style="font-size:20px;font-weight:700;color:#ffffff">' + titulo + '</div>'
    +     '<div style="font-size:12px;color:rgba(255,255,255,.6);margin-top:4px">📅 ' + today + ' &nbsp;·&nbsp; ' + tasks.length + ' tarefa' + (tasks.length !== 1 ? 's' : '') + '</div>'
    +   '</td>'
    +   (singleStatusLink
      ? '<td style="vertical-align:middle;text-align:right;white-space:nowrap;padding-left:20px">'
        + '<a href="' + singleStatusLink + '" style="display:inline-block;background:rgba(255,255,255,0.15);border:1.5px solid rgba(255,255,255,0.5);color:#ffffff;text-decoration:none;font-size:12px;font-weight:700;padding:9px 18px;border-radius:7px;white-space:nowrap;letter-spacing:.2px">📋 Ver minhas entregas</a>'
        + '</td>'
      : '<td></td>')
    + '</tr></table>'
    + '</td></tr>'

    // Instrução
    + (showConfirm ? '<tr><td style="padding:16px 28px 0">'
      + '<div style="background:#e8f5ee;border-radius:8px;padding:10px 16px;font-size:13px;color:#1a7a4a;font-weight:600">'
      + '👆 Clique em <strong>✅ Confirmar</strong> assim que concluir a tarefa. O dashboard será atualizado automaticamente.'
      + '</div></td></tr>' : '')

    // Tabela
    + '<tr><td style="padding:20px 28px 24px">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border-radius:10px;overflow:hidden;border:1px solid #dce8f5;table-layout:fixed">'
    + '<thead>'
    + '<tr style="background:#1352b8">'
    + (compact
      ? '<th style="padding:8px 10px;text-align:left;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px;width:38%">Tarefa</th>'
        + '<th style="padding:8px 10px;text-align:left;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px;width:15%">Responsável</th>'
        + '<th style="padding:8px 10px;text-align:left;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px;width:15%">Destinatário</th>'
        + '<th style="padding:8px 10px;text-align:center;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px;width:12%">Prazo</th>'
        + '<th style="padding:8px 10px;text-align:center;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px;width:20%">Status</th>'
      : '<th style="padding:10px 12px;text-align:left;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px;width:35%">Tarefa</th>'
        + '<th style="padding:10px 12px;text-align:left;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px;width:14%">Responsável</th>'
        + '<th style="padding:10px 12px;text-align:left;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px;width:14%">Destinatário</th>'
        + '<th style="padding:10px 12px;text-align:center;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px;width:10%">Prazo</th>'
        + '<th style="padding:10px 12px;text-align:center;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px;width:13%">Status</th>'
        + (showConfirm ? '<th style="padding:10px 12px;text-align:center;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px;width:14%">Ação</th>' : '')
    )
    + '</tr>'
    + '</thead>'
    + '<tbody>' + rows + '</tbody>'
    + '</table>'
    + '</td></tr>'

    // Rodapé
    + '<tr><td style="background:#f4f8fd;padding:14px 28px;border-top:1px solid #dce8f5">'
    + '<div style="font-size:11px;color:#8096b8">Mensagem automática · <strong style="color:#1352b8">Cronograma Mensal</strong> · Mabu Hospitalidade &amp; Entretenimento</div>'
    + '</td></tr>'

    + '</table>'
    + '</td></tr></table>'
    + '</body></html>';
}

function esc(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ── TESTE DE ACESSO À PLANILHA ────────────────────────────────
// Execute esta função UMA VEZ no editor para autorizar o acesso ao Sheets
function testSheetAccess() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Todos') || ss.getSheets()[0];
  var rows  = sheet.getLastRow();
  Logger.log('Acesso OK! Planilha: ' + ss.getName() + ' | Aba: ' + sheet.getName() + ' | Linhas: ' + rows);
}
