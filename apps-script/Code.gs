/**********************************************************************
 * Heino Doessing — Ontario Liberal, Hamilton East–Stoney Creek
 * Campaign Forms backend (single Apps Script web app)
 *
 * Serves 4 embeddable forms, writes each submission to a Google Sheet,
 * and emails hello@hamiltonheino.ca on every submission.
 **********************************************************************/

/*** 1) CONFIG — edit these ***/
const SHEET_ID       = '1KOnyhxwe01g9HOE3cLiAoXIgW1_wgEpMU0ZNUsq1NZU';   // "Hamilton Heino — Form Submissions"
const NOTIFY_EMAIL   = 'hello@hamiltonheino.ca';
const CANDIDATE      = 'Heino Doessing';
const CAMPAIGN       = 'Heino Doessing · Ontario Liberal for Hamilton East–Stoney Creek';  // shown in form + email
const WEBSITE_URL    = 'https://voteheino.ca';
const LEARN_MORE_URL = 'https://ontarioliberal.ca/riding/hamilton-east-stoney-creek/';  // where supporters go for more info
const ACCENT         = '#d71920';                                // Ontario Liberal red
const DONATE_URL     = 'https://ontarioliberal.ca/donate/?pla=37';   // official Elections-compliant donation page (Hamilton East–Stoney Creek PLA 37)

/*** 2) FORM DEFINITIONS — add/remove fields freely ***/
const CONSENT_LABEL = 'I agree to be contacted by the campaign by email, phone, or text. I can unsubscribe at any time.';

const FORMS = {
  vote: {
    tab: 'CommitToVote', title: 'Commit to Vote',
    blurb: 'Add your name and let us know we can count on your support for Heino Doessing.',
    submit: 'I Commit to Vote',
    fields: [
      { name:'first',  label:'First name', type:'text',  required:true },
      { name:'last',   label:'Last name',  type:'text',  required:true },
      { name:'email',  label:'Email',      type:'email', required:true },
      { name:'phone',  label:'Phone',      type:'tel',   required:false },
      { name:'postal', label:'Postal code',type:'text',  required:true }
    ]
  },
  sign: {
    tab: 'SignRequests', title: 'Request a Lawn Sign',
    blurb: 'Show your support across Hamilton East–Stoney Creek — request a sign for your lawn.',
    submit: 'Request My Sign',
    fields: [
      { name:'first',   label:'First name',    type:'text',  required:true },
      { name:'last',    label:'Last name',     type:'text',  required:true },
      { name:'email',   label:'Email',         type:'email', required:true },
      { name:'phone',   label:'Phone',         type:'tel',   required:true },
      { name:'address', label:'Street address',type:'text',  required:true },
      { name:'unit',    label:'Unit / Apt (optional)', type:'text', required:false },
      { name:'city',    label:'City / Town',   type:'text',  required:true },
      { name:'postal',  label:'Postal code',   type:'text',  required:true },
      { name:'size',    label:'Sign size',     type:'select',required:true, options:['Small (lawn)','Large'] },
      { name:'notes',   label:'Notes (placement, etc.)', type:'textarea', required:false }
    ]
  },
  donate: {
    tab: 'Donations', title: 'Commit to Donate',
    blurb: 'Pledge your support. We will follow up with a secure link to complete your contribution.',
    submit: 'Pledge My Support',
    fields: [
      { name:'first',    label:'First name', type:'text',  required:true },
      { name:'last',     label:'Last name',  type:'text',  required:true },
      { name:'email',    label:'Email',      type:'email', required:true },
      { name:'phone',    label:'Phone',      type:'tel',   required:false },
      { name:'postal',   label:'Postal code',type:'text',  required:true },
      { name:'amount',   label:'Amount you would like to pledge (CAD)', type:'select', required:true,
        options:['$25','$50','$100','$250','$500','$1,000','Other'] },
      { name:'frequency',label:'Frequency',  type:'select',required:true, options:['One-time','Monthly'] }
    ],
    note: 'This is a pledge only — no payment is collected here. Contributions must be made on the official, Elections-compliant donation page, which we will send you.'
  },
  volunteer: {
    tab: 'Volunteers', title: 'Volunteer Sign-Up',
    blurb: 'Join the team to elect Heino Doessing. Tell us how you would like to help.',
    submit: 'Sign Me Up',
    fields: [
      { name:'first',       label:'First name', type:'text',  required:true },
      { name:'last',        label:'Last name',  type:'text',  required:true },
      { name:'email',       label:'Email',      type:'email', required:true },
      { name:'phone',       label:'Phone',      type:'tel',   required:true },
      { name:'postal',      label:'Postal code',type:'text',  required:true },
      { name:'interests',   label:'I would like to help with', type:'multicheck', required:false,
        options:['Canvassing (door-knocking)','Phone banking','Sign installation','Data entry','Events','Social media'] },
      { name:'availability',label:'When are you available?', type:'text', required:false }
    ]
  }
};

/*** 3) SERVE THE FORM (GET) ***/
function doGet(e) {
  const p = (e && e.parameter) || {};

  // Staff shift tracker — the page itself is public, but every data call inside
  // it requires a login (see Auth.gs). Serving the shell without a session is fine.
  if (p.view === 'week') {
    return HtmlService.createTemplateFromFile('Week').evaluate()
      .setTitle('Shift tracker — ' + CAMPAIGN)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // Staff phone bank — same story: public shell, every data call requires a login.
  if (p.view === 'phonebank') {
    return HtmlService.createTemplateFromFile('Phonebank').evaluate()
      .setTitle('Phone bank — ' + CAMPAIGN)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // Public canvass shift signup
  if (p.form === 'shift') {
    const t = HtmlService.createTemplateFromFile('Shift');
    t.campaign = CAMPAIGN;
    t.accent = ACCENT;
    t.consentLabel = CONSENT_LABEL;
    return t.evaluate()
      .setTitle('Sign up to canvass — ' + CAMPAIGN)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const key = p.form || 'vote';
  const def = FORMS[key] || FORMS.vote;
  const t = HtmlService.createTemplateFromFile('Form');
  t.form = key;
  t.def = def;
  t.campaign = CAMPAIGN;
  t.accent = ACCENT;
  t.consentLabel = CONSENT_LABEL;
  t.donateUrl = DONATE_URL;
  t.learnMoreUrl = LEARN_MORE_URL;
  return t.evaluate()
    .setTitle(def.title + ' — ' + CAMPAIGN)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // allow embedding on any site
}

/*** 4) HANDLE A SUBMISSION (called by the form via google.script.run) ***/
function submitForm(formKey, data) {
  const def = FORMS[formKey];
  if (!def) throw new Error('Unknown form: ' + formKey);
  if (data && data._hp) return { ok: true };            // honeypot filled = bot; silently drop

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ensureSheet_(ss, def);
  const headers = headerRow_(def);

  const row = [new Date()];
  def.fields.forEach(function (f) {
    let v = data[f.name];
    if (Array.isArray(v)) v = v.join(', ');
    row.push(v == null ? '' : String(v).trim());
  });
  row.push(data.consent ? 'Yes' : 'No');  // Consent
  row.push(WEBSITE_URL);                   // Source
  row.push(false);                         // In Liberalist? (checkbox)
  row.push('');                            // Internal Notes

  sheet.appendRow(row);
  const r = sheet.getLastRow();
  const libCol = headers.indexOf('In Liberalist?') + 1;
  sheet.getRange(r, libCol).insertCheckboxes();
  sheet.getRange(r, libCol).setValue(false);

  notify_(def, data, r);
  return { ok: true };
}

/*** 5) HELPERS ***/
function headerRow_(def) {
  return ['Timestamp']
    .concat(def.fields.map(function (f) { return f.label; }))
    .concat(['Consent', 'Source', 'In Liberalist?', 'Internal Notes']);
}

function ensureSheet_(ss, def) {
  let sheet = ss.getSheetByName(def.tab);
  const headers = headerRow_(def);
  if (!sheet) {
    sheet = ss.insertSheet(def.tab);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function notify_(def, data, rowNum) {
  const lines = def.fields.map(function (f) {
    let v = data[f.name];
    if (Array.isArray(v)) v = v.join(', ');
    return f.label + ': ' + (v || '—');
  });
  lines.push('Consent to contact: ' + (data.consent ? 'Yes' : 'No'));

  const subject = '[' + def.title + '] ' + (data.first || '') + ' ' + (data.last || '');
  const body =
    'New "' + def.title + '" submission for ' + CAMPAIGN + ':\n\n' +
    lines.join('\n') +
    '\n\n— Added to the "' + def.tab + '" tab of the campaign Google Sheet (row ' + rowNum +
    '), ready for Liberalist entry.\n' +
    'Sheet: https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit\n';

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: subject,
    body: body,
    replyTo: data.email || NOTIFY_EMAIL,
    name: CAMPAIGN + ' Forms'
  });
}
