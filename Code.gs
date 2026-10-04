/**
 * Internal Interview Form -> Existing Sheet -> Standard Google Calendar Invitation
 * (Google Meet + CV attachment)
 */

// ====== SETTINGS ======
// Secret/account-specific value gulo (Sheet ID, API keys, group IDs, default emails)
// code-e likha nai -- Apps Script-er "Script Properties"-e rakha ache.
// Setup: Apps Script -> gear icon (Project Settings) -> Script Properties -> Add property.
// Dorkari key gulo README.md-e deya ache.
const CFG = PropertiesService.getScriptProperties();
function cfgStr_(key, fallback) { return CFG.getProperty(key) || fallback || ''; }
function cfgJson_(key, fallback) {
  try { const v = CFG.getProperty(key); return v ? JSON.parse(v) : fallback; }
  catch (e) { return fallback; }
}

const SPREADSHEET_ID = cfgStr_('SPREADSHEET_ID');
const TARGET_SHEET_NAME = 'Interviews';   // Tomar existing tab-er nam
const TIMEZONE = 'Asia/Dhaka';
const COMPANY_NAME = cfgStr_('COMPANY_NAME', 'Your Company');  // Event title-e boshbe
const CV_FOLDER_NAME = 'Interview CVs';   // Drive-e ei nam-e folder auto toiri hobe

// Form field  ->  Tomar sheet-er column header
const COLUMNS = {
  name:          'Employee Name',
  email:         'Email',
  mobile:        'Mobile',
  position:      'Position',
  interviewers:  'Interviewers',
  date:          'Interview Date',   // <-- Tomar Sheet-er header eta
  time:          'Time',           // Sheet-e "03:00-03:15 PM" format-e boshbe
  interviewType: 'Interview Type',
  status:        'Status',
  area:          'Area',
  notes:         'Notes',
  cvLink:        'CV Link'         // auto toiri hobe
};
const INVITE_STATUS_HEADER = 'Invite Status';   // auto toiri hobe

// Ei row-ta "template" hishebe use hobe: Interview Type / Status column-er
// color/dropdown format ei row theke notun row-e copy hobe.
// Sheet-er 1st data row (mane row 2)-e Interview Type = Online, Status = Will Attend
// diye already color/dropdown set kora thakle default 2 rekhe dao.
const FORMAT_TEMPLATE_ROW = 2;

// Interviewer Emails box-e default hisebe thakbe (form-e giye add/remove kora jabe)
// Script Properties key: DEFAULT_INTERVIEWERS -> JSON array, e.g. ["a@x.com","b@y.com"]
const DEFAULT_INTERVIEWERS = cfgJson_('DEFAULT_INTERVIEWERS', []);

// ====== WhatsApp (Green-API) ======
const WHATSAPP_ENABLED = true;
// Script Properties: GREEN_API_ID_INSTANCE, GREEN_API_TOKEN
const GREEN_API_ID_INSTANCE = cfgStr_('GREEN_API_ID_INSTANCE');
const GREEN_API_TOKEN = cfgStr_('GREEN_API_TOKEN');
// getChats theke paoa SHOTHIK Group ID ekhane boshao (format: xxxxxxxxxxxxxxxxxx@g.us)
// Script Properties key: WHATSAPP_GROUPS -> JSON array
// e.g. [{"id":"hr","name":"HR Internal Team","groupId":"xxxxxxxxxxxxxxxxxx@g.us"}]
const WHATSAPP_GROUPS = cfgJson_('WHATSAPP_GROUPS', []);
// ======================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Interview')
    .addItem('New Interview Invite', 'showForm')
    .addToUi();
}

function showForm() {
  const html = HtmlService.createHtmlOutput(renderForm_())
    .setWidth(760).setHeight(840);
  SpreadsheetApp.getUi().showModalDialog(html, 'Interview Scheduling');
}

// Web app link hishebe use korle
function doGet(e) {
  // GitHub Pages (ba onno external page) theke config (default interviewers,
  // WhatsApp groups) fetch korar jonno: <webAppUrl>?action=config
  if (e && e.parameter && e.parameter.action === 'config') {
    return jsonResponse_({
      defaultInterviewers: DEFAULT_INTERVIEWERS,
      whatsappGroups: WHATSAPP_GROUPS.map(g => ({ id: g.id, name: g.name }))
    });
  }
  return HtmlService.createHtmlOutput(renderForm_())
    .setTitle('Interview Scheduling')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// External page (GitHub Pages) theke fetch() diye form submit korar entry point
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const message = submitInterview(data);
    return jsonResponse_({ ok: true, message: message });
  } catch (err) {
    return jsonResponse_({ ok: false, message: String(err && err.message || err) });
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Default interviewer emails FORM_HTML-er bhitor boshiye dey
function renderForm_() {
  return FORM_HTML
    .replace('__DEFAULT_INTERVIEWERS__', JSON.stringify(DEFAULT_INTERVIEWERS))
    .replace('__WHATSAPP_GROUPS__', JSON.stringify(WHATSAPP_GROUPS.map(g => ({ id: g.id, name: g.name }))));
}

function getSS_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID)
                        : SpreadsheetApp.getActiveSpreadsheet();
}

// Form theke call hoy
function submitInterview(d) {
  const sheet = getSS_().getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) throw new Error('Tab paoa jay nai: ' + TARGET_SHEET_NAME);
  if (typeof Calendar === 'undefined') {
    throw new Error('Google Calendar API enable kora nai. Apps Script -> Services (+) -> Google Calendar API -> Add koren, tarpor New version deploy koren.');
  }

  // Start / End time
  const start = Utilities.parseDate(d.date + ' ' + d.startTime, TIMEZONE, 'yyyy-MM-dd HH:mm');
  const end   = Utilities.parseDate(d.date + ' ' + d.endTime,   TIMEZONE, 'yyyy-MM-dd HH:mm');
  if (end <= start) throw new Error('End Time, Start Time-er pore hote hobe.');

  // Guests = candidate + interviewers
  const guestList = [];
  const badEmails = [];
  [d.email].concat(String(d.interviewers || '').split(/[\s,;]+/)).forEach(s => {
    s = String(s).trim().toLowerCase();
    if (!s) return;
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s)) { badEmails.push(s); return; }
    if (guestList.indexOf(s) === -1) guestList.push(s);
  });
  if (badEmails.length) throw new Error('Email thik noy: ' + badEmails.join(', '));

  // 0) CV Drive-e save
  let cvFile = null;
  if (d.cv && d.cv.data) {
    cvFile = saveCv_(d.cv, d.name);
    try { cvFile.addViewers(guestList); } catch (e) { /* share na hole-o cholbe */ }
    d.cvLink = cvFile.getUrl();
  }

  // 1) Sheet-e row add (column header onujayi)
  const statusCol = ensureColumn_(sheet, INVITE_STATUS_HEADER);
  ensureColumn_(sheet, COLUMNS.cvLink);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(h => String(h).trim());
  const rowData = {};
  Object.keys(COLUMNS).forEach(k => rowData[COLUMNS[k]] = d[k]);
  rowData[COLUMNS.date] = Utilities.formatDate(start, TIMEZONE, 'dd-MMM-yyyy');
  rowData[COLUMNS.time] = formatRange_(start, end);
  rowData[COLUMNS.interviewers] = guestList.slice(1).join(', ');
  const newRow = headers.map(h => rowData[h] === undefined ? '' : rowData[h]);
  sheet.appendRow(newRow);
  const rowNum = sheet.getLastRow();
  applyTemplateFormat_(sheet, headers, rowNum, [COLUMNS.interviewType, COLUMNS.status]);

  // 2) Standard Google Calendar invitation
  try {
    const r = createInvite_(d, start, end, guestList, cvFile);
    if (d.interviewType === 'Online' && !r.meetOk) {
      sheet.getRange(rowNum, statusCol).setValue('Invite Sent ✅ (Meet link toiri hoyni)');
      return 'Invite pathano hoyeche, kintu Google Meet link toiri hoyni. Calendar-e giye manually Meet add koren.';
    }
    if (d.interviewType === 'Online' && !r.openOk) {
      sheet.getRange(rowNum, statusCol).setValue('Invite Sent ✅ (Meet Open hoyni: ' + r.openError + ')');
      return 'Invite pathano hoyeche ✅ kintu Meet access "Open" set kora jay nai (' + r.openError + '). Setup step ta abar check koren.';
    }
    sheet.getRange(rowNum, statusCol).setValue('Invite Sent ✅');
    sendWhatsAppInvite_(d, start, end, r.meetLink);
    return 'Invite pathano hoyeche ✅\nGuests: ' + guestList.join(', ');
  } catch (err) {
    sheet.getRange(rowNum, statusCol).setValue('Error: ' + err.message);
    throw new Error('Data save hoyeche kintu invite pathano jay nai: ' + err.message);
  }
}

// Event title: Invitation to Interview at WeGro | TFO - Chuadanga | MD. RABBI HOSSAIN
function buildTitle_(d) {
  const middle = [d.position, d.area].filter(Boolean).join(' - ');
  return ['Invitation to Interview at ' + COMPANY_NAME, middle, d.name]
    .filter(Boolean).join(' | ');
}

// Mail-er bhitorer likha: shudhu Employee Name, Position, Area
// "Tuesday, 29 September · 3:45 – 4:00pm" -- direct Calendar-er format-er moto
function buildScheduleLine_(start, end) {
  const dateStr = Utilities.formatDate(start, TIMEZONE, 'EEEE, d MMMM');
  const sTime = Utilities.formatDate(start, TIMEZONE, 'h:mm');
  const sMer  = Utilities.formatDate(start, TIMEZONE, 'a').toLowerCase();
  const eTime = Utilities.formatDate(end, TIMEZONE, 'h:mm');
  const eMer  = Utilities.formatDate(end, TIMEZONE, 'a').toLowerCase();
  const timeStr = (sMer === eMer)
    ? sTime + ' \u2013 ' + eTime + eMer
    : sTime + sMer + ' \u2013 ' + eTime + eMer;
  return dateStr + ' \u00b7 ' + timeStr;
}

function buildWhatsAppMessage_(d, start, end, meetLink) {
  const lines = [
    buildTitle_(d),
    buildScheduleLine_(start, end),
    'Time zone: ' + TIMEZONE
  ];
  if (d.interviewType === 'Online' && meetLink) {
    lines.push('', 'Google Meet joining info', 'Video call link: ' + meetLink);
  } else if (d.interviewType !== 'Online' && d.area) {
    lines.push('', 'Location: ' + d.area);
  }
  return lines.join('\n');
}

function sendWhatsAppInvite_(d, start, end, meetLink) {
  if (!WHATSAPP_ENABLED) return;
  const message = buildWhatsAppMessage_(d, start, end, meetLink);

  // Groups (checkbox diye select kora)
  const selectedGroups = String(d.whatsappGroups || '').split(',').map(s => s.trim()).filter(Boolean);
  WHATSAPP_GROUPS
    .filter(g => selectedGroups.indexOf(g.id) !== -1)
    .forEach(g => sendGreenApiMessage_(g.groupId, message));

  // Personal numbers (chip box diye add kora)
  String(d.personalWhatsapp || '').split(',').map(s => s.trim()).filter(Boolean)
    .forEach(num => sendGreenApiMessage_(num + '@c.us', message));
}

function sendGreenApiMessage_(chatId, message) {
  try {
    const url = 'https://api.green-api.com/waInstance' + GREEN_API_ID_INSTANCE +
      '/sendMessage/' + GREEN_API_TOKEN;
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chatId: chatId, message: message }),
      muteHttpExceptions: true
    });
  } catch (e) { /* WhatsApp fail korle-o invite already jawa thik ache, script thambe na */ }
}

function buildDescription_(d) {
  const lines = [];
  if (d.name)     lines.push('Employee Name: ' + d.name);
  if (d.position) lines.push('Position: ' + d.position);
  if (d.area)     lines.push('Area: ' + d.area);
  if (d.notes)    lines.push('', d.notes);
  return lines.join('\n');
}

// Return: true = Meet link ready (ba Online noy)
function createInvite_(d, start, end, guestList, cvFile) {
  const isOnline = d.interviewType === 'Online';

  const ev = {
    summary: buildTitle_(d),
    description: buildDescription_(d),
    start: { dateTime: iso_(start), timeZone: TIMEZONE },
    end:   { dateTime: iso_(end),   timeZone: TIMEZONE },
    reminders: { useDefault: true }
  };

  // Offline interview hole Area-ta event-er Location hishebe jabe
  if (!isOnline && d.area) ev.location = d.area;

  // Online -> Google Meet (access type OPEN: kono admin approval / knock lagbe na)
  let openOk = false, openError = '';
  if (isOnline) {
    const res = createOpenMeetSpace_();
    if (res.space) {
      const sp = res.space;
      ev.hangoutLink = sp.meetingUri;
      ev.conferenceData = {
        conferenceId: sp.meetingCode,
        entryPoints: [{
          entryPointType: 'video',
          uri: sp.meetingUri,
          label: 'meet.google.com/' + sp.meetingCode
        }],
        conferenceSolution: { key: { type: 'hangoutsMeet' } }
      };
      openOk = true;
    } else {
      // Fallback: normal Meet (default access). Invite tobu jabe
      openError = res.error;
      ev.conferenceData = {
        createRequest: {
          requestId: Utilities.getUuid(),
          conferenceSolutionKey: { type: 'hangoutsMeet' }
        }
      };
    }
  }

  // CV attachment
  if (cvFile) {
    ev.attachments = [{
      fileUrl: cvFile.getUrl(),
      title: cvFile.getName(),
      mimeType: cvFile.getMimeType()
    }];
  }

  // Step 1: event toiri (guest chhara, kono mail jabe na)
  let created = withRetry_(() => Calendar.Events.insert(ev, 'primary', {
    sendUpdates: 'none',
    supportsAttachments: true,
    conferenceDataVersion: 1
  }));

  // Step 2: Meet link ready hoyeche ki na check (mail-e jeno link thake)
  let meetOk = true;
  if (isOnline) {
    for (let i = 0; i < 6 && !hasMeet_(created); i++) {
      Utilities.sleep(1000);
      created = withRetry_(() => Calendar.Events.get('primary', created.id));
    }
    meetOk = hasMeet_(created);
  }

  // Step 3: guest add -> tokhon standard invitation mail jay (Meet + CV shoho)
  withRetry_(() => Calendar.Events.patch(
    { attendees: guestList.map(email => ({ email: email })) },
    'primary',
    created.id,
    { sendUpdates: 'all', supportsAttachments: true, conferenceDataVersion: 1 }
  ));
  const meetLink = created.hangoutLink ||
    (created.conferenceData && created.conferenceData.entryPoints &&
     created.conferenceData.entryPoints[0] && created.conferenceData.entryPoints[0].uri) || '';
  return { meetOk: meetOk, openOk: openOk, openError: openError, meetLink: meetLink };
}

function hasMeet_(ev) {
  return !!(ev.hangoutLink ||
    (ev.conferenceData && ev.conferenceData.entryPoints && ev.conferenceData.entryPoints.length));
}

// Google Calendar API rate-limit / temporary error hole 3 bar retry kore (wait diye)
function withRetry_(fn) {
  const delays = [2000, 4000, 8000];
  for (let i = 0; i <= delays.length; i++) {
    try {
      return fn();
    } catch (err) {
      const msg = String(err && err.message || err);
      const retryable = /rate limit|quota|backend|internal error|503|500/i.test(msg);
      if (!retryable || i === delays.length) throw err;
      Utilities.sleep(delays[i]);
    }
  }
}

// Meet REST API diye "OPEN" access-er meeting space toiri
function createOpenMeetSpace_() {
  try {
    const resp = UrlFetchApp.fetch('https://meet.googleapis.com/v2/spaces', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      payload: JSON.stringify({ config: { accessType: 'OPEN' } }),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    if (code !== 200) {
      let msg = 'HTTP ' + code;
      try { msg += ' ' + JSON.parse(resp.getContentText()).error.message; } catch (e) {}
      return { space: null, error: msg.slice(0, 160) };
    }
    const sp = JSON.parse(resp.getContentText());
    return sp.meetingUri ? { space: sp, error: '' } : { space: null, error: 'meetingUri paoa jay nai' };
  } catch (err) {
    return { space: null, error: String(err.message || err).slice(0, 160) };
  }
}

function saveCv_(cv, name) {
  const folders = DriveApp.getFoldersByName(CV_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(CV_FOLDER_NAME);
  const safeName = String(name).replace(/[\\\/:*?"<>|]/g, '').trim();
  const blob = Utilities.newBlob(
    Utilities.base64Decode(cv.data),
    cv.mimeType || 'application/pdf',
    safeName + ' - CV - ' + cv.name
  );
  return folder.createFile(blob);
}

function iso_(date) {
  return Utilities.formatDate(date, TIMEZONE, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

// 03:00-03:15 PM  (AM/PM alada hole: 11:45 AM-12:15 PM)
function formatRange_(start, end) {
  const s = Utilities.formatDate(start, TIMEZONE, 'hh:mm a');
  const e = Utilities.formatDate(end, TIMEZONE, 'hh:mm a');
  return (s.slice(-2) === e.slice(-2)) ? s.slice(0, 5) + '-' + e : s + '-' + e;
}

// Template row (FORMAT_TEMPLATE_ROW)-er format/dropdown color notun row-e copy kore,
// value ta ager moto thik thake, shudhu format/validation ashe.
function applyTemplateFormat_(sheet, headers, rowNum, colNames) {
  if (rowNum === FORMAT_TEMPLATE_ROW) return; // template row nijei, kichu korar dorkar nai
  colNames.forEach(name => {
    const idx = headers.indexOf(name);
    if (idx === -1) return;
    const col = idx + 1;
    const src = sheet.getRange(FORMAT_TEMPLATE_ROW, col);
    const dest = sheet.getRange(rowNum, col);
    try {
      src.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      src.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    } catch (e) { /* format copy fail hole-o value thik thakbe */ }
  });
}

function ensureColumn_(sheet, name) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(h => String(h).trim());
  let idx = headers.indexOf(name);
  if (idx === -1) {
    idx = headers.length;
    sheet.getRange(1, idx + 1).setValue(name);
  }
  return idx + 1;
}

// Prothombar ekbar Run korle Calendar / Sheet / Drive permission chaibe
function authorize() {
  CalendarApp.getDefaultCalendar();
  SpreadsheetApp.openById(SPREADSHEET_ID);
  DriveApp.getRootFolder();
  UrlFetchApp.fetch('https://meet.googleapis.com/$discovery/rest?version=v2', { muteHttpExceptions: true });
}

// ====== FORM HTML (Form.html file lagbe na) ======
const FORM_HTML = `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    :root { --navy:#0f2a4a; --accent:#1b6ef3; --bg:#f3f5f9; --border:#d8dee8; --muted:#5f6b7a; --ink:#1c2533; }
    * { box-sizing: border-box; }
    body { margin:0; background:var(--bg); color:var(--ink); font-family:"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif; font-size:14px; }
    .wrap { max-width:760px; margin:0 auto; padding:16px; }
    .header { background:linear-gradient(135deg,#0f2a4a,#1b4b8a); color:#fff; padding:22px 26px; border-radius:14px 14px 0 0; }
    .header h1 { margin:0; font-size:22px; font-weight:600; letter-spacing:.2px; }
    .header p { margin:6px 0 0; font-size:13px; opacity:.85; }
    .card { background:#fff; padding:22px 26px 26px; border-radius:0 0 14px 14px; box-shadow:0 4px 18px rgba(15,42,74,.10); }
    .section { margin-top:22px; }
    .section:first-child { margin-top:0; }
    .section h2 { margin:0 0 12px; padding-bottom:8px; border-bottom:2px solid #eaeef5; font-size:12px; font-weight:700; text-transform:uppercase; letter-spacing:1px; color:var(--navy); }
    .grid { display:grid; grid-template-columns:1fr 1fr; gap:14px 16px; }
    .full { grid-column:1 / -1; }
    label { display:block; margin-bottom:5px; font-size:12.5px; font-weight:600; color:#344054; }
    label .req { color:#d92d20; }
    input, select, textarea { width:100%; padding:10px 12px; border:1px solid var(--border); border-radius:8px; font-size:14px; font-family:inherit; color:var(--ink); background:#fff; transition:border .15s, box-shadow .15s; }
    input, select { height:42px; }
    textarea { resize:vertical; }
    input:focus, select:focus, textarea:focus { outline:none; border-color:var(--accent); box-shadow:0 0 0 3px rgba(27,110,243,.15); }
    .hint { margin-top:4px; font-size:11.5px; color:var(--muted); }
    .upload { display:flex; align-items:center; gap:12px; padding:14px 16px; border:2px dashed #b9c4d6; border-radius:10px; background:#f8fafd; cursor:pointer; }
    .upload:hover { border-color:var(--accent); background:#f1f6ff; }
    .upload .icon { font-size:26px; }
    .upload .txt b { display:block; font-size:13.5px; color:var(--navy); }
    .upload .txt span { font-size:12px; color:var(--muted); }
    #cv { display:none; }
    button { margin-top:24px; width:100%; height:46px; background:var(--accent); color:#fff; border:0; border-radius:10px; font-size:15px; font-weight:600; cursor:pointer; transition:background .15s; }
    button:hover { background:#1558c9; }
    button:disabled { background:#9bbcf0; cursor:not-allowed; }
    #msg { margin-top:14px; padding:0; font-size:13.5px; border-radius:8px; }
    #msg.ok { padding:12px 14px; background:#e7f6ec; color:#136c34; border:1px solid #b7e4c7; }
    #msg.err { padding:12px 14px; background:#fdecea; color:#b42318; border:1px solid #f5c2bd; }
    .chips { display:flex; flex-wrap:wrap; align-items:center; gap:6px; min-height:42px; padding:5px 8px; border:1px solid var(--border); border-radius:8px; background:#fff; cursor:text; }
    .chips:focus-within { border-color:var(--accent); box-shadow:0 0 0 3px rgba(27,110,243,.15); }
    #chipList { display:contents; }
    .chip { display:inline-flex; align-items:center; gap:6px; background:#e8f0fe; color:#174ea6; border-radius:16px; padding:3px 8px 3px 11px; font-size:13px; }
    .chip .x { cursor:pointer; font-size:15px; line-height:1; color:#5f6b7a; }
    .chip .x:hover { color:#d92d20; }
    .chips input { flex:1; min-width:170px; height:28px; padding:0 4px; border:0; box-shadow:none; }
    .chips input:focus { border:0; box-shadow:none; }
    .toggle-row { display:flex; align-items:center; justify-content:space-between; padding:10px 14px; background:#f8fafd; border:1px solid var(--border); border-radius:10px; margin-bottom:8px; }
    .toggle-row:last-child { margin-bottom:0; }
    .toggle-row .lbl b { display:block; font-size:13.5px; color:var(--navy); }
    .toggle-row .lbl span { font-size:11.5px; color:var(--muted); }
    .switch { position:relative; width:42px; height:24px; flex-shrink:0; }
    .switch input { opacity:0; width:0; height:0; }
    .slider { position:absolute; inset:0; background:#c6ccd6; border-radius:24px; cursor:pointer; transition:.15s; }
    .slider:before { content:""; position:absolute; height:18px; width:18px; left:3px; top:3px; background:#fff; border-radius:50%; transition:.15s; }
    .switch input:checked + .slider { background:var(--accent); }
    .switch input:checked + .slider:before { transform:translateX(18px); }
    @media (max-width:600px) { .grid { grid-template-columns:1fr; } .header, .card { padding-left:18px; padding-right:18px; } }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="header">
      <h1>Interview Scheduling</h1>
      <p>Candidate-er details din, calendar invitation auto pathiye deoa hobe.</p>
    </div>

    <div class="card">

      <div class="section">
        <h2>Candidate Details</h2>
        <div class="grid">
          <div class="full">
            <label>Employee Name <span class="req">*</span></label>
            <input id="name" placeholder="e.g. MD. RABBI HOSSAIN">
          </div>
          <div>
            <label>Email <span class="req">*</span></label>
            <input id="email" type="email" placeholder="candidate@email.com">
          </div>
          <div>
            <label>Mobile</label>
            <input id="mobile" type="tel" placeholder="01XXXXXXXXX">
          </div>
          <div>
            <label>Position <span class="req">*</span></label>
            <select id="position">
              <option value="" selected>Select position</option>
              <option value="Field Officer">Field Officer</option>
              <option value="Area Manager">Area Manager</option>
              <option value="TSO">TSO</option>
              <option value="Junior TSO">Junior TSO</option>
              <option value="Regional Manager">Regional Manager</option>
            </select>
          </div>
          <div>
            <label>Area</label>
            <input id="area" placeholder="e.g. Chuadanga">
          </div>
          <div class="full">
            <label>Candidate CV</label>
            <label class="upload" for="cv">
              <span class="icon">📎</span>
              <span class="txt"><b id="cvName">Click kore CV select koren</b><span id="cvSub">PDF, DOC ba DOCX (max 10 MB)</span></span>
            </label>
            <input id="cv" type="file" accept=".pdf,.doc,.docx">
          </div>
        </div>
      </div>

      <div class="section">
        <h2>Interview Schedule</h2>
        <div class="grid">
          <div class="full">
            <label>Date <span class="req">*</span></label>
            <input id="date" type="date">
          </div>
          <div>
            <label>Start Time <span class="req">*</span></label>
            <input id="startTime" type="time">
          </div>
          <div>
            <label>End Time <span class="req">*</span></label>
            <input id="endTime" type="time">
          </div>
          <div>
            <label>Interview Type</label>
            <select id="interviewType">
              <option value="Online" selected>Online</option>
              <option value="Offline">Offline</option>
            </select>
          </div>
          <div>
            <label>Status</label>
            <select id="status">
              <option value="Will Attend" selected>Will Attend</option>
              <option value="Will Not Attend">Will Not Attend</option>
              <option value="Pending">Pending</option>
            </select>
          </div>
        </div>
      </div>

      <div class="section">
        <h2>Interview Panel &amp; Notes</h2>
        <div class="grid">
          <div class="full">
            <label>Interviewer Emails</label>
            <div class="chips" id="chipBox" onclick="$('interviewerInput').focus()">
              <span id="chipList"></span>
              <input id="interviewerInput" placeholder="Email likhe Enter chapun">
            </div>
            <div class="hint">Enter, comma ba space chaple email add hobe. Ei email gulo-o candidate-er moto invitation pabe.</div>
          </div>
          <div class="full">
            <label>Notes</label>
            <textarea id="notes" rows="3" placeholder="Internal note (optional)"></textarea>
            <div class="hint">Notes shudhu Sheet-e save hobe, invitation mail-e jabe na.</div>
          </div>
        </div>
      </div>

      <div class="section" id="waSection">
        <h2>Share on WhatsApp (optional)</h2>
        <div id="waGroups"></div>

        <label style="margin-top:14px;">Personal WhatsApp Numbers</label>
        <div class="chips" id="waNumBox" onclick="$('waNumInput').focus()">
          <span id="waNumList"></span>
          <input id="waNumInput" placeholder="01XXXXXXXXX likhe Enter chapun">
        </div>
        <div class="hint">Ekadhik number add kora jay. Country code na dile Bangladesh (+880) dhore neya hobe.</div>

        <div class="toggle-row" style="margin-top:10px;">
          <div class="lbl"><b>Send to Personal Numbers</b><span>On thakle upor-er number gulote message jabe</span></div>
          <label class="switch">
            <input type="checkbox" id="sendPersonalWa">
            <span class="slider"></span>
          </label>
        </div>
      </div>

      <button id="btn" onclick="submitForm()">Save &amp; Send Invitation</button>
      <div id="msg"></div>
    </div>
  </div>

  <script>
    var fields = ['name','email','mobile','position','date','startTime','endTime','interviewType','status','area','notes'];
    var defaults = { interviewType: 'Online', status: 'Will Attend', position: '' };
    var interviewers = __DEFAULT_INTERVIEWERS__;
    var waGroupsList = __WHATSAPP_GROUPS__;
    var personalNumbers = [];
    var MAX_MB = 10;

    function $(id) { return document.getElementById(id); }

    $('cv').addEventListener('change', function () {
      var f = this.files[0];
      if (!f) { resetCv(); return; }
      if (f.size > MAX_MB * 1024 * 1024) {
        showMsg('err', 'CV file ' + MAX_MB + ' MB-er cheye boro.');
        this.value = ''; resetCv(); return;
      }
      $('cvName').textContent = f.name;
      $('cvSub').textContent = (f.size / 1024 / 1024).toFixed(2) + ' MB - change korte abar click koren';
      showMsg('', '');
    });

    // ---- Interviewer email chips (Google Calendar guest-er moto) ----
    renderChips();
    renderWaGroups();
    renderWaNumbers();

    function renderWaGroups() {
      var box = $('waGroups');
      waGroupsList.forEach(function (g) {
        var row = document.createElement('div');
        row.className = 'toggle-row';
        row.innerHTML =
          '<div class="lbl"><b>' + g.name + '</b></div>' +
          '<label class="switch">' +
            '<input type="checkbox" class="waChk" data-id="' + g.id + '">' +
            '<span class="slider"></span>' +
          '</label>';
        box.appendChild(row);
      });
    }

    // ---- Personal WhatsApp number chips ----
    function cleanPhone(raw) {
      var digits = raw.replace(/[^0-9]/g, '');
      if (!digits) return '';
      if (digits.length === 11 && digits.charAt(0) === '0') digits = '880' + digits.slice(1);
      else if (digits.length === 10) digits = '880' + digits;
      return digits;
    }
    function isPhone(raw) {
      var d = cleanPhone(raw);
      return d.length >= 10 && d.length <= 15;
    }
    function splitPhoneTokens(str) {
      return str.split(/[\s,;]+/).map(function (s) { return s.trim(); }).filter(Boolean);
    }
    function addNumbers(text) {
      var bad = [];
      splitPhoneTokens(text).forEach(function (t) {
        if (!isPhone(t)) { bad.push(t); return; }
        var d = cleanPhone(t);
        if (personalNumbers.indexOf(d) === -1) personalNumbers.push(d);
      });
      renderWaNumbers();
      return bad;
    }
    function commitPhoneInput() {
      var inp = $('waNumInput');
      if (!inp.value.trim()) return true;
      var bad = addNumbers(inp.value);
      inp.value = bad.join(' ');
      if (bad.length) {
        showMsg('err', 'Number thik noy: ' + bad.join(', '));
        return false;
      }
      showMsg('', '');
      return true;
    }
    function renderWaNumbers() {
      var list = $('waNumList');
      while (list.firstChild) list.removeChild(list.firstChild);
      personalNumbers.forEach(function (num, idx) {
        var chip = document.createElement('span');
        chip.className = 'chip';
        var t = document.createElement('span');
        t.textContent = '+' + num;
        var x = document.createElement('span');
        x.className = 'x';
        x.textContent = '\u00d7';
        x.onclick = function (ev) { ev.stopPropagation(); personalNumbers.splice(idx, 1); renderWaNumbers(); };
        chip.appendChild(t); chip.appendChild(x);
        list.appendChild(chip);
      });
    }
    $('waNumInput').addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ',' || e.key === ' ' || e.key === ';') {
        e.preventDefault();
        commitPhoneInput();
      } else if (e.key === 'Backspace' && !this.value && personalNumbers.length) {
        personalNumbers.pop();
        renderWaNumbers();
      }
    });
    $('waNumInput').addEventListener('input', function () {
      if (/[ ,;]/.test(this.value)) commitPhoneInput();
    });
    $('waNumInput').addEventListener('blur', function () { commitPhoneInput(); });

    function isEmail(s) { return /^[^ @]+@[^ @]+[.][^ @]+$/.test(s); }

    function splitTokens(str) {
      var out = [], cur = '';
      for (var i = 0; i < str.length; i++) {
        var ch = str.charAt(i);
        if (' ,;'.indexOf(ch) >= 0 || str.charCodeAt(i) < 32) {
          if (cur) out.push(cur);
          cur = '';
        } else { cur += ch; }
      }
      if (cur) out.push(cur);
      return out;
    }

    function addEmails(text) {
      var bad = [];
      splitTokens(text).forEach(function (t) {
        t = t.toLowerCase();
        if (!isEmail(t)) { bad.push(t); return; }
        if (interviewers.indexOf(t) === -1) interviewers.push(t);
      });
      renderChips();
      return bad;
    }

    function commitInput() {
      var inp = $('interviewerInput');
      if (!inp.value.trim()) return true;
      var bad = addEmails(inp.value);
      inp.value = bad.join(' ');
      if (bad.length) {
        showMsg('err', 'Email thik noy: ' + bad.join(', '));
        return false;
      }
      showMsg('', '');
      return true;
    }

    function renderChips() {
      var list = $('chipList');
      while (list.firstChild) list.removeChild(list.firstChild);
      interviewers.forEach(function (em, idx) {
        var chip = document.createElement('span');
        chip.className = 'chip';
        var t = document.createElement('span');
        t.textContent = em;
        var x = document.createElement('span');
        x.className = 'x';
        x.textContent = '\u00d7';
        x.onclick = function (ev) { ev.stopPropagation(); interviewers.splice(idx, 1); renderChips(); };
        chip.appendChild(t); chip.appendChild(x);
        list.appendChild(chip);
      });
    }

    $('interviewerInput').addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ',' || e.key === ' ' || e.key === ';') {
        e.preventDefault();
        commitInput();
      } else if (e.key === 'Backspace' && !this.value && interviewers.length) {
        interviewers.pop();
        renderChips();
      }
    });
    $('interviewerInput').addEventListener('input', function () {
      if (/[ ,;]/.test(this.value)) commitInput();
    });
    $('interviewerInput').addEventListener('blur', function () { commitInput(); });

    function resetCv() {
      $('cv').value = '';
      $('cvName').textContent = 'Click kore CV select koren';
      $('cvSub').textContent = 'PDF, DOC ba DOCX (max 10 MB)';
    }

    function showMsg(type, text) {
      var m = $('msg');
      m.className = type;
      m.textContent = text;
    }

    function setBusy(busy) {
      var b = $('btn');
      b.disabled = busy;
      b.textContent = busy ? 'Sending... ekto opekkha koren' : 'Save & Send Invitation';
    }

    function submitForm() {
      if (!commitInput()) return;
      if (!commitPhoneInput()) return;
      var d = {};
      fields.forEach(function (f) { d[f] = $(f).value.trim(); });
      d.interviewers = interviewers.join(',');
      d.whatsappGroups = Array.prototype.slice.call(document.querySelectorAll('.waChk:checked'))
        .map(function (el) { return el.getAttribute('data-id'); }).join(',');
      d.personalWhatsapp = $('sendPersonalWa').checked ? personalNumbers.join(',') : '';

      if (!d.name || !d.email || !d.position || !d.date || !d.startTime || !d.endTime) {
        showMsg('err', 'Employee Name, Email, Position, Date, Start Time ar End Time dite hobe.');
        return;
      }
      if (d.endTime <= d.startTime) {
        showMsg('err', 'End Time, Start Time-er pore hote hobe.');
        return;
      }
      showMsg('', '');
      setBusy(true);

      var file = $('cv').files[0];
      if (!file) { send(d); return; }

      var reader = new FileReader();
      reader.onload = function (e) {
        d.cv = { name: file.name, mimeType: file.type, data: String(e.target.result).split(',')[1] };
        send(d);
      };
      reader.onerror = function () {
        setBusy(false);
        showMsg('err', 'CV file porte problem hoyeche.');
      };
      reader.readAsDataURL(file);
    }

    function send(d) {
      google.script.run
        .withSuccessHandler(function (res) {
          showMsg('ok', res);
          setBusy(false);
          fields.forEach(function (f) { $(f).value = defaults[f] !== undefined ? defaults[f] : ''; });
          interviewers = __DEFAULT_INTERVIEWERS__.slice(); renderChips(); $('interviewerInput').value = '';
          document.querySelectorAll('.waChk').forEach(function (el) { el.checked = false; });
          personalNumbers = []; renderWaNumbers(); $('waNumInput').value = ''; $('sendPersonalWa').checked = false;
          resetCv();
        })
        .withFailureHandler(function (err) {
          showMsg('err', err.message);
          setBusy(false);
        })
        .submitInterview(d);
    }
  </script>
</body>
</html>
`;
