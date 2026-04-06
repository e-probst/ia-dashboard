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
//  5. Execute createDailyTrigger() UMA VEZ para ativar envio às 7h
// ============================================================

// ⚠️  PREENCHA COM O ID DA SUA PLANILHA GOOGLE SHEETS
//     URL da planilha: https://docs.google.com/spreadsheets/d/<<ID_AQUI>>/edit
//     Cole apenas o trecho entre /d/ e /edit
var SPREADSHEET_ID = '1FMoWYDqersAk8zXy_a_ZDShUDU-s339eOC9f5P2mKfY';
var EMAIL_FROM_NAME = 'Cronograma Mensal · Mabu Hospitalidade';
var ADMIN_EMAIL     = 'e.probst@mymabu.com.br';
var GAS_WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwK8FMgwRXkF-Z6krN12yHKgMJmDNxsgVBZoka4PJgSlNYx4f29wxs3XTKtW35B27Tc/exec';

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
    var confirmedAt = new Date().toISOString();
    var data = JSON.stringify({ id: id, name: name, prazo: prazo, confirmedAt: confirmedAt });
    props.setProperty(key, data);
    Logger.log('Entrega confirmada via e-mail: id=' + id + ' | ' + name);
    // Atualiza imediatamente a planilha (evita lag até o próximo polling do dashboard)
    if (SPREADSHEET_ID) {
      try {
        var today = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
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

// Busca tarefas do responsável no Sheets filtrando: prazo hoje + atrasadas ≤30d
// Mesma janela usada pelo dashboard e pelo dailyEmailJob
function getRespTasksForPanel(resp, currentId, props) {
  var result = [];
  if (!resp || !SPREADSHEET_ID) return result;

  var OVERDUE_WINDOW = 30;
  var today = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
  var tParts = today.split('/');
  var todayDate = new Date(+tParts[2], +tParts[1]-1, +tParts[0]);

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

      // ── Filtro: só prazo hoje OU atrasada dentro da janela de 30d ──
      var include = false;
      if (rowPraz === today) {
        include = true; // vence hoje
      } else if (rowPraz && rowPraz !== '—') {
        var pp = rowPraz.split('/');
        if (pp.length === 3) {
          var dPrazo = new Date(+pp[2], +pp[1]-1, +pp[0]);
          if (!isNaN(dPrazo) && dPrazo < todayDate) {
            var diff = Math.round((todayDate - dPrazo) / 86400000);
            if (diff <= OVERDUE_WINDOW) include = true; // atrasada ≤30d
          }
        }
      }
      if (!include) continue;

      // ── Status de entrega: Sheets tem precedência; fallback para PropertiesService ──
      var sheetsDelivered = rowEntrega && rowEntrega !== '—';
      var confRaw = rowId ? props.getProperty('confirm_' + rowId) : null;
      var confirmed   = sheetsDelivered || !!confRaw;
      var confirmedAt = '';
      if (sheetsDelivered) {
        confirmedAt = rowEntrega; // data de entrega já registrada na planilha
      } else if (confRaw) {
        try {
          var iso = (JSON.parse(confRaw).confirmedAt||'').slice(0,10).split('-');
          if (iso.length === 3) confirmedAt = iso[2]+'/'+iso[1]+'/'+iso[0];
        } catch(e2) {}
      }

      result.push({
        id: rowId, name: rowName, prazo: rowPraz, mes: rowMes,
        confirmed: confirmed, confirmedAt: confirmedAt,
        isCurrent: (rowId === currentId)
      });
    }

    // Ordena: pendentes primeiro (por prazo asc), depois entregues
    function prazoCmp(p){ var s=(p||'').split('/'); return s.length===3?s[2]+s[1]+s[0]:'99999999'; }
    result.sort(function(a, b) {
      if (!a.confirmed &&  b.confirmed) return -1;
      if ( a.confirmed && !b.confirmed) return  1;
      return prazoCmp(a.prazo) < prazoCmp(b.prazo) ? -1 : prazoCmp(a.prazo) > prazoCmp(b.prazo) ? 1 : 0;
    });
  } catch(e) {
    Logger.log('getRespTasksForPanel erro: ' + e.message);
  }
  return result;
}

function handleGetConfirmations(callback) {
  var props  = PropertiesService.getScriptProperties();
  var all    = props.getProperties();
  var result = [];
  Object.keys(all).forEach(function(k) {
    if (k.indexOf('confirm_') === 0) {
      try { result.push(JSON.parse(all[k])); } catch(e) {}
    }
  });
  return jsonpResponse({ ok: true, confirmations: result }, callback);
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
function handleGetTasks(callback) {
  if (!SPREADSHEET_ID) {
    return jsonpResponse({ ok: false, error: 'SPREADSHEET_ID nao configurado.' }, callback);
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
    return jsonpResponse({ ok: true, tasks: tasks, count: tasks.length, ts: new Date().toISOString() }, callback);
  } catch (err) {
    Logger.log('handleGetTasks error: ' + err.message + ' | stack: ' + err.stack);
    return jsonpResponse({ ok: false, error: String(err.message) }, callback);
  }
}

// viewOnly=true → página de status sem ação de confirmação (link "Ver minhas entregas")
function buildConfirmationPage(name, prazo, alreadyDone, resp, allRespTasks, viewOnly) {
  var date = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  var msg;
  if (viewOnly) {
    msg = alreadyDone
      ? 'Esta tarefa já foi confirmada.'
      : 'Esta tarefa ainda <strong>não foi confirmada</strong>. Clique em <em>✅ Confirmar</em> no e-mail para registrar a entrega.';
  } else {
    msg = alreadyDone
      ? 'Esta entrega já havia sido confirmada anteriormente.'
      : 'Confirmação registrada em <strong>' + date + '</strong>. O dashboard será atualizado automaticamente.';
  }

  // ── Painel de entregas do responsável ──
  var painelHtml = '';
  if (allRespTasks && allRespTasks.length > 0) {
    var taskItems = allRespTasks.map(function(t) {
      var isCurrent = t.isCurrent;
      var conf      = t.confirmed;
      var statusBg    = conf ? '#e6f5ee' : '#fff8ec';
      var statusColor = conf ? '#1a7a4a' : '#b45300';
      var statusIcon  = conf ? '✅' : '⏳';
      var statusLabel = conf
        ? 'Confirmado' + (t.confirmedAt ? ' em ' + t.confirmedAt : '')
        : 'Pendente';
      return '<div style="display:flex;align-items:flex-start;gap:12px;padding:11px 14px;'
        + 'border-radius:10px;margin-bottom:8px;'
        + 'background:' + (isCurrent ? '#f0fbf4' : '#f8faff') + ';'
        + 'border:' + (isCurrent ? '2px solid #1a7a4a' : '1px solid #dce8f5') + '">'
        + '<div style="font-size:20px;line-height:1;margin-top:1px">' + statusIcon + '</div>'
        + '<div style="flex:1;min-width:0">'
        +   '<div style="font-size:13px;font-weight:' + (isCurrent ? '700' : '500') + ';color:#0a1e45;word-break:break-word">'
        +     esc(t.name)
        +     (isCurrent && !viewOnly && !alreadyDone ? '&nbsp;<span style="background:#1a7a4a;color:#fff;font-size:10px;padding:1px 8px;border-radius:99px;vertical-align:middle">confirmada agora</span>' : '')
        +   '</div>'
        +   (t.prazo && t.prazo !== '—' ? '<div style="font-size:11px;color:#8096b8;margin-top:3px">📅 Prazo: ' + esc(t.prazo) + '</div>' : '')
        + '</div>'
        + '<div style="flex-shrink:0;text-align:right">'
        +   '<span style="display:inline-block;background:' + statusBg + ';color:' + statusColor + ';border-radius:99px;padding:3px 10px;font-size:11px;font-weight:700;white-space:nowrap">'
        +     statusLabel
        +   '</span>'
        + '</div>'
        + '</div>';
    }).join('');

    var totalConf = allRespTasks.filter(function(t){ return t.confirmed; }).length;
    var totalPend = allRespTasks.length - totalConf;

    painelHtml = '<div style="margin-top:28px;text-align:left">'
      + '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px">'
      +   '<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.9px;color:#8096b8">📋 Suas Entregas — ' + esc(resp) + '</div>'
      +   '<div style="font-size:11px;color:#8096b8">'
      +     '<span style="color:#1a7a4a;font-weight:700">' + totalConf + ' confirmada' + (totalConf !== 1 ? 's' : '') + '</span>'
      +     (totalPend > 0 ? ' &nbsp;·&nbsp; <span style="color:#b45300;font-weight:700">' + totalPend + ' pendente' + (totalPend !== 1 ? 's' : '') + '</span>' : '')
      +   '</div>'
      + '</div>'
      + taskItems
      + '</div>';
  }

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<style>'
    + '*{box-sizing:border-box;margin:0;padding:0}'
    + 'body{font-family:Arial,sans-serif;background:#f0f6ff;display:flex;justify-content:center;min-height:100vh;padding:20px}'
    + '.card{background:#fff;border-radius:16px;box-shadow:0 4px 24px rgba(13,45,110,.13);max-width:560px;width:100%;overflow:hidden;align-self:flex-start;margin-top:20px}'
    + '.header{background:linear-gradient(135deg,#0d2d6e,#1352b8);padding:24px 28px;text-align:center}'
    + '.header h1{color:#fff;font-size:17px;font-weight:700;margin-bottom:3px}'
    + '.header p{color:rgba(255,255,255,.65);font-size:12px}'
    + '.body{padding:28px}'
    + '.icon{font-size:52px;text-align:center;margin-bottom:12px}'
    + '.title{font-size:19px;font-weight:700;color:#1a7a4a;text-align:center;margin-bottom:6px}'
    + '.task-box{font-size:14px;color:#0a1e45;font-weight:600;background:#f0f6ff;border-radius:8px;padding:10px 16px;margin:10px 0;text-align:center}'
    + '.info{font-size:13px;color:#3a5080;text-align:center;margin-top:10px;line-height:1.6}'
    + '.prazo{font-size:12px;color:#8096b8;text-align:center;margin-top:5px}'
    + '.footer{padding:14px 28px;background:#f7f9fc;border-top:1px solid #d6e3f5;text-align:center;font-size:11px;color:#8096b8}'
    + '</style></head><body>'
    + '<div class="card">'
    + '<div class="header"><h1>Cronograma Mensal · Mabu Hospitalidade</h1><p>' + (viewOnly ? 'Acompanhamento de Entregas' : 'Confirmação de Entrega') + '</p></div>'
    + '<div class="body">'
    + '<div class="icon">' + (viewOnly ? '📋' : '✅') + '</div>'
    + '<div class="title">' + (viewOnly ? 'Minhas Entregas' : 'Entrega Confirmada!') + '</div>'
    + (name ? '<div class="task-box">' + esc(name) + '</div>' : '')
    + (prazo ? '<div class="prazo">📅 Prazo: ' + prazo + '</div>' : '')
    + '<div class="info">' + msg + '</div>'
    + painelHtml
    + '</div>'
    + '<div class="footer">Cronograma Mensal · Mabu Hospitalidade &amp; Entretenimento</div>'
    + '</div></body></html>';
}

function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action || '';
    var result;

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
  var to   = body.to   || ADMIN_EMAIL;
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
      var todayTasks    = g.tasks.filter(function(t){ return t.status !== 'ATRASADO'; });
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

  return { ok: errs.length === 0, sent: sent, errors: errs };
}

// Template simplificado do e-mail geral (evita falha com múltiplas tarefas)
function buildSummaryEmail(tasks, today) {
  var statusMap = {
    'ATRASADO':            { color:'#c0182c', bg:'#fde8eb', label:'Atrasado' },
    'ENTREGUE COM ATRASO': { color:'#c05000', bg:'#fff4e0', label:'Entregue c/ Atraso' },
    'NO PRAZO':            { color:'#1352b8', bg:'#e8f2fc', label:'No Prazo' },
    'ENTREGUE':            { color:'#1a7a4a', bg:'#e6f5ee', label:'Entregue' },
    'ENTREGA ANTECIPADA':  { color:'#0550ae', bg:'#e0f0ff', label:'Antecipada' },
  };

  var rows = tasks.map(function(t, idx) {
    var st     = statusMap[t.status] || { color:'#3a5080', bg:'#f0f6ff', label: t.status || '—' };
    var rowBg  = idx % 2 === 0 ? '#ffffff' : '#f7faff';
    return '<tr style="background:' + rowBg + '">'
      + '<td width="38%" style="padding:11px 14px;border-bottom:1px solid #eaf2fc;font-size:13px;font-weight:700;color:#0a1e45">' + esc(t.name) + (t.note ? '<br><span style="font-size:11px;color:#8096b8;font-weight:400;font-style:italic">📝 ' + esc(t.note) + '</span>' : '') + '</td>'
      + '<td width="18%" style="padding:11px 14px;border-bottom:1px solid #eaf2fc;font-size:12px;color:#3a5080">' + esc(t.resp || '—') + '</td>'
      + '<td width="18%" style="padding:11px 14px;border-bottom:1px solid #eaf2fc;font-size:12px;color:#3a5080">' + esc(t.dest || '—') + '</td>'
      + '<td width="12%" style="padding:11px 14px;border-bottom:1px solid #eaf2fc;font-size:12px;color:#3a5080;text-align:center">' + esc(t.prazo || '—') + '</td>'
      + '<td width="14%" style="padding:11px 14px;border-bottom:1px solid #eaf2fc;text-align:center"><span style="background:' + st.bg + ';color:' + st.color + ';border-radius:99px;padding:3px 10px;font-size:11px;font-weight:700">' + st.label + '</span></td>'
      + '</tr>';
  }).join('');

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>'
    + '<body style="margin:0;padding:0;background:#eef4fb;font-family:Arial,sans-serif">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#eef4fb;padding:28px 0"><tr><td align="center">'
    + '<table width="680" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:14px;overflow:hidden;box-shadow:0 4px 24px rgba(13,45,110,.12)">'
    + '<tr><td style="background:linear-gradient(135deg,#0d2d6e,#1a52b8);padding:24px 28px">'
    + '<div style="font-size:11px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:2px;margin-bottom:5px">Cronograma Mensal · Mabu Hospitalidade</div>'
    + '<div style="font-size:20px;font-weight:700;color:#fff">Resumo Geral · Tarefas com Prazo Hoje</div>'
    + '<div style="font-size:12px;color:rgba(255,255,255,.6);margin-top:4px">📅 ' + today + ' &nbsp;·&nbsp; ' + tasks.length + ' tarefa' + (tasks.length !== 1 ? 's' : '') + '</div>'
    + '</td></tr>'
    + '<tr><td style="padding:20px 28px 24px">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid #dce8f5;border-radius:10px;overflow:hidden">'
    + '<thead><tr style="background:#1352b8">'
    + '<th width="38%" style="padding:10px 14px;text-align:left;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px">Tarefa</th>'
    + '<th width="18%" style="padding:10px 14px;text-align:left;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px">Responsável</th>'
    + '<th width="18%" style="padding:10px 14px;text-align:left;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px">Destinatário</th>'
    + '<th width="12%" style="padding:10px 14px;text-align:center;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px">Prazo</th>'
    + '<th width="14%" style="padding:10px 14px;text-align:center;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px">Status</th>'
    + '</tr></thead>'
    + '<tbody>' + rows + '</tbody>'
    + '</table></td></tr>'
    + '<tr><td style="background:#f4f8fd;padding:14px 28px;border-top:1px solid #dce8f5">'
    + '<div style="font-size:11px;color:#8096b8">Mensagem automática · <strong style="color:#1352b8">Cronograma Mensal</strong> · Mabu Hospitalidade &amp; Entretenimento</div>'
    + '</td></tr></table>'
    + '</td></tr></table></body></html>';
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

  return { ok: errs.length === 0, sent: sent, errors: errs };
}

// E-mail geral para o admin com todas as tarefas (sem botão)
function handleSendSummary(body) {
  var tasks = body.tasks || [];
  if (!tasks.length) return { ok: true, skipped: true };

  try {
    var today   = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
    var subject = '[Cronograma Mensal] Resumo das tarefas do dia — ' + today;
    var html    = buildEmailHtml(tasks, 'Resumo Geral · Tarefas com Prazo Hoje', false);
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
    .atHour(7)
    .everyDays(1)
    .inTimezone('America/Sao_Paulo')
    .create();
  Logger.log('Trigger diário criado: dailyEmailJob às 7h (Brasília)');
}

// ── JOB DIÁRIO ───────────────────────────────────────────────
// Roda às 7h automaticamente; pode ser testado manualmente

function dailyEmailJob() {
  if (!SPREADSHEET_ID) {
    Logger.log('dailyEmailJob: SPREADSHEET_ID não configurado');
    return;
  }
  try {
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

function buildEmailHtml(tasks, titulo, showConfirm) {
  var today  = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
  var gasUrl = GAS_WEB_APP_URL;

  var statusMap = {
    'ATRASADO':            { color:'#c0182c', bg:'#fde8eb', label:'Atrasado' },
    'ENTREGUE COM ATRASO': { color:'#c05000', bg:'#fff4e0', label:'Entregue c/ Atraso' },
    'NO PRAZO':            { color:'#1352b8', bg:'#e8f2fc', label:'No Prazo' },
    'ENTREGUE':            { color:'#1a7a4a', bg:'#e6f5ee', label:'Entregue' },
    'ENTREGA ANTECIPADA':  { color:'#0550ae', bg:'#e0f0ff', label:'Antecipada' },
  };

  var rows = tasks.map(function(t, idx) {
    var st          = statusMap[t.status] || { color:'#3a5080', bg:'#f0f6ff', label: t.status || '—' };
    var isDelivered = (t.status === 'ENTREGUE' || t.status === 'ENTREGA ANTECIPADA' || t.status === 'ENTREGUE COM ATRASO');
    var baseParams = '&id=' + encodeURIComponent(t.id) + '&prazo=' + encodeURIComponent(t.prazo || '') + '&name=' + encodeURIComponent(t.name || '') + '&resp=' + encodeURIComponent(t.resp || '') + '&month=' + encodeURIComponent(t.month || '');
    var confirmLink = (showConfirm && gasUrl && !isDelivered)
      ? gasUrl + '?action=confirm' + baseParams
      : '';
    var statusLink = (showConfirm && gasUrl && t.resp && t.resp !== '—')
      ? gasUrl + '?action=status' + baseParams
      : '';
    var rowBg = idx % 2 === 0 ? '#ffffff' : '#f7faff';

    return '<tr style="background:' + rowBg + '">'
      + '<td style="padding:12px 14px;border-bottom:1px solid #eaf2fc;vertical-align:middle">'
      +   '<div style="font-size:13px;font-weight:700;color:#0a1e45">' + esc(t.name) + '</div>'
      +   (t.note ? '<div style="font-size:11px;color:#8096b8;margin-top:2px;font-style:italic">📝 ' + esc(t.note) + '</div>' : '')
      + '</td>'
      + '<td style="padding:12px 14px;border-bottom:1px solid #eaf2fc;font-size:12px;color:#3a5080;vertical-align:middle;white-space:nowrap">' + esc(t.resp || '—') + '</td>'
      + '<td style="padding:12px 14px;border-bottom:1px solid #eaf2fc;font-size:12px;color:#3a5080;vertical-align:middle;white-space:nowrap">' + esc(t.dest || '—') + '</td>'
      + '<td style="padding:12px 14px;border-bottom:1px solid #eaf2fc;font-size:12px;color:#3a5080;vertical-align:middle;white-space:nowrap;text-align:center">' + esc(t.prazo || '—') + '</td>'
      + '<td style="padding:12px 14px;border-bottom:1px solid #eaf2fc;vertical-align:middle;text-align:center">'
      +   '<span style="display:inline-block;background:' + st.bg + ';color:' + st.color + ';border-radius:99px;padding:3px 10px;font-size:11px;font-weight:700;white-space:nowrap">' + st.label + '</span>'
      + '</td>'
      + '<td style="padding:12px 14px;border-bottom:1px solid #eaf2fc;vertical-align:middle;text-align:center">'
      +   (confirmLink
          ? '<a href="' + confirmLink + '" style="display:inline-block;background:#1a7a4a;color:#ffffff;text-decoration:none;font-size:11px;font-weight:700;padding:6px 14px;border-radius:6px;white-space:nowrap">✅ Confirmar</a>'
            + (statusLink ? '<br><a href="' + statusLink + '" style="display:inline-block;margin-top:5px;font-size:10px;color:#8096b8;text-decoration:none;">📋 Ver minhas entregas</a>' : '')
          : (isDelivered
              ? '<span style="color:#1a7a4a;font-size:12px;font-weight:700">✅ Entregue</span>'
                + (statusLink ? '<br><a href="' + statusLink + '" style="display:inline-block;margin-top:5px;font-size:10px;color:#8096b8;text-decoration:none;">📋 Ver minhas entregas</a>' : '')
              : ''))
      + '</td>'
      + '</tr>';
  }).join('');


  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>'
    + '<body style="margin:0;padding:0;background:#eef4fb;font-family:Arial,sans-serif">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#eef4fb;padding:28px 0">'
    + '<tr><td align="center">'
    + '<table width="680" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:14px;overflow:hidden;box-shadow:0 4px 24px rgba(13,45,110,.12)">'

    // Cabeçalho
    + '<tr><td style="background:linear-gradient(135deg,#0d2d6e 0%,#1a52b8 100%);padding:24px 28px">'
    + '<div style="font-size:11px;color:rgba(255,255,255,.5);text-transform:uppercase;letter-spacing:2px;margin-bottom:5px">Cronograma Mensal · Mabu Hospitalidade</div>'
    + '<div style="font-size:20px;font-weight:700;color:#ffffff">' + titulo + '</div>'
    + '<div style="font-size:12px;color:rgba(255,255,255,.6);margin-top:4px">📅 ' + today + ' &nbsp;·&nbsp; ' + tasks.length + ' tarefa' + (tasks.length !== 1 ? 's' : '') + '</div>'
    + '</td></tr>'

    // Instrução
    + (showConfirm ? '<tr><td style="padding:16px 28px 0">'
      + '<div style="background:#e8f5ee;border-radius:8px;padding:10px 16px;font-size:13px;color:#1a7a4a;font-weight:600">'
      + '👆 Clique em <strong>✅ Confirmar</strong> assim que concluir a tarefa. O dashboard será atualizado automaticamente.'
      + '</div></td></tr>' : '')

    // Tabela
    + '<tr><td style="padding:20px 28px 24px">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border-radius:10px;overflow:hidden;border:1px solid #dce8f5">'
    + '<thead>'
    + '<tr style="background:#1352b8">'
    + '<th style="padding:10px 14px;text-align:left;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px;width:35%">Tarefa</th>'
    + '<th style="padding:10px 14px;text-align:left;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px;width:18%;white-space:nowrap">Responsável</th>'
    + '<th style="padding:10px 14px;text-align:left;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px;width:18%;white-space:nowrap">Destinatário</th>'
    + '<th style="padding:10px 14px;text-align:center;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px;width:11%;white-space:nowrap">Prazo</th>'
    + '<th style="padding:10px 14px;text-align:center;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px;width:12%">Status</th>'
    + (showConfirm ? '<th style="padding:10px 14px;text-align:center;font-size:11px;color:rgba(255,255,255,.9);font-weight:700;text-transform:uppercase;letter-spacing:.6px;width:6%">Ação</th>' : '<th style="width:0"></th>')
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
