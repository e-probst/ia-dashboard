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

  if (!id) {
    return HtmlService.createHtmlOutput('<h2>Link inválido.</h2>');
  }

  // Verifica se já confirmado
  var props   = PropertiesService.getScriptProperties();
  var key     = 'confirm_' + id;
  var already = props.getProperty(key);

  if (!already) {
    var data = JSON.stringify({ id: id, name: name, prazo: prazo, confirmedAt: new Date().toISOString() });
    props.setProperty(key, data);
    Logger.log('Entrega confirmada via e-mail: id=' + id + ' | ' + name);
  }

  return HtmlService.createHtmlOutput(buildConfirmationPage(name, prazo, !!already))
    .setTitle('Entrega Confirmada · Cronograma Mensal');
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
    var iSt      = col(['Status','status']);
    var iMes     = col(['Mes','Mês','month']);

    var tasks = [];
    for (var r = 1; r < data.length; r++) {
      var row  = data[r];
      var nome = iNome >= 0 ? String(row[iNome] || '').trim() : '';
      if (!nome) continue;
      tasks.push({
        id:      (iID >= 0 && row[iID] !== '') ? Number(row[iID]) : r,
        name:    nome,
        note:    iNota    >= 0 ? String(row[iNota]    || '') : '',
        resp:    iResp    >= 0 ? String(row[iResp]    || '-') : '-',
        dest:    iDest    >= 0 ? String(row[iDest]    || '-') : '-',
        email:   iEmail   >= 0 ? String(row[iEmail]   || '') : '',
        prazo:   iPrazo   >= 0 ? String(row[iPrazo]   || '-') : '-',
        entrega: iEntrega >= 0 ? String(row[iEntrega] || '-') : '-',
        status:  iSt      >= 0 ? String(row[iSt]      || 'NO PRAZO') : 'NO PRAZO',
        month:   iMes     >= 0 ? String(row[iMes]     || '') : '',
      });
    }
    return jsonpResponse({ ok: true, tasks: tasks, count: tasks.length, ts: new Date().toISOString() }, callback);
  } catch (err) {
    Logger.log('handleGetTasks error: ' + err.message + ' | stack: ' + err.stack);
    return jsonpResponse({ ok: false, error: String(err.message) }, callback);
  }
}

function buildConfirmationPage(name, prazo, alreadyDone) {
  var date = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm');
  var msg  = alreadyDone
    ? 'Esta entrega já havia sido confirmada anteriormente.'
    : 'Confirmação registrada em <strong>' + date + '</strong>. O dashboard será atualizado automaticamente.';
  return '<!DOCTYPE html><html><head><meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Arial,sans-serif;background:#f0f6ff;display:flex;align-items:center;justify-content:center;min-height:100vh;padding:20px}'
    + '.card{background:#fff;border-radius:16px;box-shadow:0 4px 24px rgba(13,45,110,.13);max-width:480px;width:100%;overflow:hidden}'
    + '.header{background:linear-gradient(135deg,#0d2d6e,#1352b8);padding:28px 32px;text-align:center}'
    + '.header h1{color:#fff;font-size:18px;font-weight:700;margin-bottom:4px}'
    + '.header p{color:rgba(255,255,255,.65);font-size:12px}'
    + '.body{padding:32px;text-align:center}'
    + '.icon{font-size:56px;margin-bottom:16px}'
    + '.title{font-size:20px;font-weight:700;color:#1a7a4a;margin-bottom:8px}'
    + '.task{font-size:14px;color:#0a1e45;font-weight:600;background:#f0f6ff;border-radius:8px;padding:10px 16px;margin:12px 0}'
    + '.info{font-size:13px;color:#3a5080;margin-top:12px;line-height:1.6}'
    + '.prazo{font-size:12px;color:#8096b8;margin-top:6px}'
    + '.footer{padding:16px 32px;background:#f7f9fc;border-top:1px solid #d6e3f5;text-align:center;font-size:11px;color:#8096b8}'
    + '</style></head><body>'
    + '<div class="card">'
    + '<div class="header"><h1>Cronograma Mensal · Mabu Hospitalidade</h1><p>Confirmação de Entrega</p></div>'
    + '<div class="body">'
    + '<div class="icon">✅</div>'
    + '<div class="title">Entrega Confirmada!</div>'
    + '<div class="task">' + name + '</div>'
    + (prazo ? '<div class="prazo">📅 Prazo: ' + prazo + '</div>' : '')
    + '<div class="info">' + msg + '</div>'
    + '</div>'
    + '<div class="footer">Cronograma Mensal · Mabu Hospitalidade &amp; Entretenimento<br>Você pode fechar esta janela.</div>'
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

// Envia geral + individuais em uma única chamada
function handleSendAll(body) {
  var tasks = body.tasks || [];
  if (!tasks.length) return { ok: true, skipped: true };

  var sent = 0, errs = [];

  // 1. GERAL primeiro — admin recebe resumo com todas as tarefas (sem botão)
  try {
    var today      = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
    var subjGeral  = '[Cronograma Mensal] Resumo das tarefas do dia — ' + today;
    var nomes      = tasks.map(function(t){ return t.name; }).join(' | ');
    var htmlGeral  = buildSummaryEmail(tasks, today);
    MailApp.sendEmail({
      to:       ADMIN_EMAIL,
      name:     EMAIL_FROM_NAME,
      subject:  subjGeral,
      htmlBody: htmlGeral,
    });
    Logger.log('send_all geral OK → ' + tasks.length + ' tarefas: ' + nomes);
  } catch(err) {
    errs.push('geral: ' + err.message);
    Logger.log('send_all geral ERRO → ' + err.message);
  }

  // 2. INDIVIDUAIS — um e-mail por responsável (com botão de confirmação)
  tasks.forEach(function(t) {
    if (!t.email) return;
    try {
      var subject = '[Cronograma Mensal] Sua tarefa vence hoje: ' + t.name;
      var html    = buildEmailHtml([t], 'Sua Tarefa com Prazo Hoje', true);
      MailApp.sendEmail({
        to:   t.email,
        bcc:  t.email !== ADMIN_EMAIL ? ADMIN_EMAIL : '',
        name: EMAIL_FROM_NAME,
        subject: subject,
        htmlBody: html,
      });
      sent++;
      Logger.log('send_all individual OK → ' + t.email + ' | ' + t.name);
    } catch(err) {
      errs.push(t.name + ': ' + err.message);
      Logger.log('send_all individual ERRO → ' + t.email + ' | ' + err.message);
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
    var mesLabel = MONTH_SHEET[month] || '';
    if (mesLabel) {
      var sheetMes = ss.getSheetByName(mesLabel);
      if (sheetMes) deleteRowById(sheetMes, id);
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
  // Esta função não lê do dashboard (sem acesso ao localStorage).
  // Ela serve para enviar um resumo diário ao admin.
  // Para uso completo, integre com uma planilha Google Sheets
  // onde o dashboard sincroniza as tarefas via action 'sync'.

  var today  = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
  var subject = '[Cronograma Mensal] Relatório diário — ' + today;
  var html = '<div style="font-family:Arial,sans-serif;color:#0a1e45;padding:20px">'
    + '<h2 style="color:#1352b8">Cronograma Mensal · Mabu Hospitalidade</h2>'
    + '<p>O envio diário automático está ativo. ✅</p>'
    + '<p>Para exibir as tarefas do dia neste e-mail, conecte o dashboard a uma planilha Google Sheets '
    + 'e leia os dados na função <strong>dailyEmailJob()</strong>.</p>'
    + '<hr style="border:none;border-top:1px solid #d6e3f5;margin:16px 0">'
    + '<p style="font-size:11px;color:#8096b8">Mensagem automática gerada em ' + new Date().toISOString() + '</p>'
    + '</div>';

  MailApp.sendEmail({ to: ADMIN_EMAIL, name: EMAIL_FROM_NAME, subject: subject, htmlBody: html });
  Logger.log('dailyEmailJob executado: ' + today);
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
  'jan':'Janeiro','fev':'Fevereiro','mar':'Marco','abr':'Abril',
  'mai':'Maio','jun':'Junho','jul':'Julho','ago':'Agosto',
  'set':'Setembro','out':'Outubro','nov':'Novembro','dez':'Dezembro'
};

// Upsert de uma linha em uma aba pelo ID (col 0)
function upsertRow(sheet, id, row) {
  var data  = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]) === String(id)) {
      sheet.getRange(r + 1, 1, 1, SHEET_COLS.length).setValues([row]);
      return;
    }
  }
  sheet.appendRow(row);
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

function logTasksToSheet(tasks) {
  if (!SPREADSHEET_ID) return;
  var ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  var now = new Date().toISOString();

  tasks.forEach(function(t) {
    var row = [
      t.id, t.name||'', t.note||'', t.resp||'-', t.dest||'-',
      t.email||'', t.prazo||'-', t.entrega||'-', t.status||'NO PRAZO', t.month||'', now
    ];

    // Aba Todos
    var sheetTodos = ss.getSheetByName('Todos') || ss.insertSheet('Todos');
    if (sheetTodos.getLastRow() === 0) {
      sheetTodos.appendRow(SHEET_COLS);
      sheetTodos.getRange(1,1,1,SHEET_COLS.length).setFontWeight('bold');
    }
    upsertRow(sheetTodos, t.id, row);

    // Aba do mês correspondente
    var mesLabel = MONTH_SHEET[t.month] || '';
    if (mesLabel) {
      var sheetMes = ss.getSheetByName(mesLabel);
      if (!sheetMes) {
        sheetMes = ss.insertSheet(mesLabel);
        sheetMes.appendRow(SHEET_COLS);
        sheetMes.getRange(1,1,1,SHEET_COLS.length).setFontWeight('bold');
      }
      upsertRow(sheetMes, t.id, row);
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
    var isDelivered = (t.status === 'ENTREGUE' || t.status === 'ENTREGA ANTECIPADA');
    var confirmLink = (showConfirm && gasUrl && !isDelivered)
      ? gasUrl + '?action=confirm&id=' + encodeURIComponent(t.id) + '&prazo=' + encodeURIComponent(t.prazo || '') + '&name=' + encodeURIComponent(t.name || '')
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
          : (isDelivered ? '<span style="color:#1a7a4a;font-size:12px;font-weight:700">✅ Entregue</span>' : ''))
      + '</td>'
      + '</tr>';
  }).join('');

  var instrucao = showConfirm
    ? '<tr><td style="padding:0 28px 16px" colspan="1">'
      + '<div style="background:#e8f5ee;border-radius:8px;padding:10px 16px;font-size:13px;color:#1a7a4a;font-weight:600">'
      + '👆 Clique em <strong>✅ Confirmar</strong> assim que concluir a tarefa. O dashboard será atualizado automaticamente.'
      + '</div></td></tr>'
    : '';

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
    .replace(/"/g, '&quot;');
}

// ── TESTE DE ACESSO À PLANILHA ────────────────────────────────
// Execute esta função UMA VEZ no editor para autorizar o acesso ao Sheets
function testSheetAccess() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('Todos') || ss.getSheets()[0];
  var rows  = sheet.getLastRow();
  Logger.log('Acesso OK! Planilha: ' + ss.getName() + ' | Aba: ' + sheet.getName() + ' | Linhas: ' + rows);
}
