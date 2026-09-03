/**
 * MAGIC MANAGER — LEADS APP (server)
 * Bound to your leads Google Sheet.
 *
 * Serves the app UI, reads/writes lead rows, and keeps a dedicated
 * follow-ups Google Calendar in sync so your phone
 * gets time-based notifications for upcoming follow-ups
 * (future dates in the current calendar year only).
 */

const SHEET_NAME = 'Sheet1';
const TASKS_SHEET = 'Tasks';
const EXPENSES_SHEET = 'Expenses';
const LOG_SHEET = 'Contact Log';
/* ---------- Business configuration (per-install) ----------
 * Everything specific to a given business lives in a "Settings" tab of the
 * spreadsheet (setting key in column A, value in column B), so each copy of
 * this app reflects its own owner without touching code. Any key the owner
 * hasn't set falls back to a sensible generic default. Cached per execution.
 */
var CONFIG_DEFAULTS_ = {
  BUSINESS_NAME: 'Your Business',
  BUSINESS_LEGAL_NAME: 'Your Business LLC',
  OWNER_NAME: 'Owner Name',
  EMAIL: '',
  PHONE: '',
  WEBSITE: '',
  TAGLINE: 'Professional Services',
  ADDRESS: '',
  EVENT_PREP_NOTE: '',
  SERVICE_MESSAGES: '',
  VENMO_USERNAME: '',
  CASHAPP_CASHTAG: '',
  ACH_BANK: '', ACH_ACCOUNT: '', ACH_ROUTING: '',
  WIRE_BANK: '', WIRE_ACCOUNT: '', WIRE_ROUTING: '', WIRE_BANK_ADDRESS: '',
  CHECK_PAYEE: '', CHECK_ADDRESS: '',
  LOGO_URL: '',
  GOVERNING_LAW: '',
  CONTRACTS_FOLDER: '',
  RECEIPTS_FOLDER: '',
  CAL_NAME: '',
  CONTACT_GROUP: '',
  CALL_LINK: '',
  TEXT_LINK: '',
  SERVICES: 'Strolling\nStage\nStage & Strolling',
  EVENT_TYPES: 'Corporate\nBirthday\nWedding\nBar Mitzvah\nFundraiser\nSchool\nBanquet\nCocktail Party\nOther',
  LEAD_SOURCES: 'Bark\nGigsalad\nWebsite\nCurrent Client\nReferral\nOther',
  AD_SOURCES: 'Bark\nGigsalad\nWebsite',
  LOST_REASONS: 'Budget\nNon-responsive\nPostponed\nBooked elsewhere'
};

function getConfig_() {
  if (getConfig_._cache) return getConfig_._cache;
  var cfg = {};
  for (var k in CONFIG_DEFAULTS_) cfg[k] = CONFIG_DEFAULTS_[k];
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName('Settings');
    if (sh && sh.getLastRow() > 0) {
      var rows = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
      for (var i = 0; i < rows.length; i++) {
        var key = String(rows[i][0] || '').trim();
        if (!key || !(key in CONFIG_DEFAULTS_)) continue;
        var val = rows[i][1];
        if (val === '' || val === null || typeof val === 'undefined') continue;
        cfg[key] = (typeof val === 'string') ? val.trim() : val;
      }
    }
  } catch (err) { /* no Settings tab yet — use defaults */ }
  if (!cfg.CONTRACTS_FOLDER) cfg.CONTRACTS_FOLDER = cfg.BUSINESS_NAME + ' Contracts';
  if (!cfg.RECEIPTS_FOLDER) cfg.RECEIPTS_FOLDER = cfg.BUSINESS_NAME + ' Receipts';
  if (!cfg.CAL_NAME) cfg.CAL_NAME = cfg.BUSINESS_NAME + ' Follow-ups';
  if (!cfg.CONTACT_GROUP) cfg.CONTACT_GROUP = cfg.BUSINESS_NAME + ' Leads';
  getConfig_._cache = cfg;
  return cfg;
}

/* ---------- Settings (per-install onboarding) ----------
 * CONFIG_FIELDS_ drives the in-app Settings screen (labels + help text) and the
 * "Settings" sheet. getSettings()/saveSettings() are called from the UI so a new
 * owner can enter their own info without touching code or the raw sheet.
 */
// Setup screen is grouped into collapsible sections (SETTINGS_SECTIONS_ below);
// each field declares which `section` it belongs to. Order here is the order
// shown within a section.
var CONFIG_FIELDS_ = [
  // Your business
  { key: 'BUSINESS_NAME', label: 'Business name', help: 'Shown throughout the app and to clients.', section: 'business' },
  { key: 'BUSINESS_LEGAL_NAME', label: 'Legal business name', help: 'Full legal entity name used on contracts.', section: 'business' },
  { key: 'OWNER_NAME', label: 'Your name', help: 'Owner / service provider name on contracts and the sign page.', section: 'business' },
  { key: 'TAGLINE', label: 'Tagline', help: 'Short line under your name on the contract.', section: 'business' },
  { key: 'LOGO_URL', label: 'Logo', help: 'Upload a PNG or JPG — square works best (it gets embedded into your app). Avoid SVG and iPhone HEIC files. Pasting a link works only if it is a direct, public image URL (Drive/Dropbox share links will not).', section: 'business' },
  // How clients reach you
  { key: 'EMAIL', label: 'Contact email', help: 'Shown to clients on the sign page and contract.', section: 'contact' },
  { key: 'PHONE', label: 'Contact phone', help: 'Shown to clients (optional).', section: 'contact' },
  { key: 'WEBSITE', label: 'Website', help: 'Shown to clients (optional), e.g. yourbusiness.com.', section: 'contact' },
  { key: 'ADDRESS', label: 'Mailing address', help: 'Used on the contract and for check/wire instructions.', section: 'contact' },
  // Getting paid
  { key: 'VENMO_USERNAME', label: 'Venmo username', help: 'Enables the Venmo option. Blank hides it.', section: 'payments' },
  { key: 'CASHAPP_CASHTAG', label: 'Cash App $cashtag', help: 'Enables Cash App. Blank hides it.', section: 'payments' },
  { key: 'ACH_BANK', label: 'ACH bank name', help: 'Shown with ACH instructions.', section: 'payments' },
  { key: 'ACH_ACCOUNT', label: 'ACH account number', help: 'ACH option appears only if account + routing are set.', section: 'payments' },
  { key: 'ACH_ROUTING', label: 'ACH routing number', help: '', section: 'payments' },
  { key: 'WIRE_BANK', label: 'Wire bank name', help: 'Shown with wire instructions.', section: 'payments' },
  { key: 'WIRE_ACCOUNT', label: 'Wire account number', help: 'Wire option appears only if account + routing are set.', section: 'payments' },
  { key: 'WIRE_ROUTING', label: 'Wire routing number', help: '', section: 'payments' },
  { key: 'WIRE_BANK_ADDRESS', label: 'Wire bank address', help: 'Bank address for wire instructions.', section: 'payments' },
  { key: 'CHECK_PAYEE', label: 'Check payable to', help: 'Enables "Mail a check". Blank uses your legal name.', section: 'payments' },
  { key: 'CHECK_ADDRESS', label: 'Check mailing address', help: 'Where checks are mailed. Blank uses your mailing address.', section: 'payments' },
  // Services & client page
  { key: 'SERVICES', label: 'Services offered', help: 'One per line — the choices in the Service dropdown.', multiline: true, section: 'services' },
  { key: 'EVENT_TYPES', label: 'Event types', help: 'One per line — the choices in the Event Type dropdown.', multiline: true, section: 'services' },
  { key: 'LEAD_SOURCES', label: 'Lead sources', help: 'One per line — how a client first found you.', multiline: true, section: 'services' },
  { key: 'AD_SOURCES', label: 'Advertising sources', help: 'One per line — paid channels tracked in Insights ROI (e.g. Bark, Gigsalad).', multiline: true, section: 'services' },
  { key: 'LOST_REASONS', label: 'Lost reasons', help: 'One per line — your own choices in the "Why was this lead lost?" picker (clients never see these; they group your Insights). An "Other…" free-text option is always there too.', multiline: true, section: 'services' },
  { key: 'EVENT_PREP_NOTE', label: 'Event prep note', help: 'Shown to clients after they pay their deposit (e.g. setup needs, what to prepare). Leave blank to hide.', section: 'services' },
  { key: 'SERVICE_MESSAGES', label: 'Event messages by service', help: 'Optional. A message shown to a booked client on their event hub after they sign, matched to their booking\'s service (e.g. one message for a stage show, another for strolling). A box appears for each of your Services above; leave any blank.', editor: 'servicemsgs', section: 'services' },
  // Calls & texts
  { key: 'CALL_LINK', label: 'Custom call link', help: 'Optional — blank uses your phone\'s default dialer. To route through another app, paste its dial link with {number} or {digits} where the number goes — Skype works directly: skype:{number}?call. If the app just opens without a number (e.g. Google Voice: googlevoice://), the client\'s number is copied to your clipboard so you can paste it in. Tip: to use Google Voice for everything, it\'s simplest to set it as your phone\'s default app and leave this blank.', section: 'comms' },
  { key: 'TEXT_LINK', label: 'Custom text link', help: 'Optional — blank uses your phone\'s default messaging. Use {number}/{digits} for the number and {body} for the message where the app\'s link supports them. If the app just opens (e.g. Google Voice: googlevoice://), your message is copied to your clipboard so you can paste it in after you pick the contact.', section: 'comms' },
  // Advanced (optional)
  { key: 'CONTRACTS_FOLDER', label: 'Contracts folder', help: 'Optional. Blank = "[Business name] Contracts".', section: 'advanced' },
  { key: 'RECEIPTS_FOLDER', label: 'Receipts folder', help: 'Optional. Blank = "[Business name] Receipts".', section: 'advanced' },
  { key: 'CAL_NAME', label: 'Follow-ups calendar', help: 'Optional. Blank = "[Business name] Follow-ups".', section: 'advanced' },
  { key: 'CONTACT_GROUP', label: 'Contacts group', help: 'Optional. Blank = "[Business name] Leads".', section: 'advanced' }
];

// Setup screen sections, in display order. `open` = expanded by default (the
// two must-fill sections); the rest start collapsed to keep first-run calm.
var SETTINGS_SECTIONS_ = [
  { id: 'business', title: 'Your business', desc: 'The essentials — these appear in your app, on your contracts, and on the page your clients see.', open: true },
  { id: 'contact', title: 'How clients reach you', desc: 'Your contact details, shown on your contract and the client sign & pay page.', open: true },
  { id: 'payments', title: 'Getting paid', desc: 'Switch on the payment methods you accept — leave the rest blank. You can add these anytime.', open: false },
  { id: 'services', title: 'Services & client page', desc: 'The dropdown choices inside your app, plus the note clients see after they pay their deposit.', open: false },
  { id: 'comms', title: 'Calls & texts', desc: 'Tapping Call or Text opens your phone\'s built-in apps by default. To route through another app (like Google Voice), set it up here — most people can leave this alone.', open: false },
  { id: 'advanced', title: 'Advanced (optional)', desc: 'Names for the Drive folders and calendar the app creates. The defaults work great — most people never change these.', open: false }
];

function settingsSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Settings');
  if (!sh) {
    sh = ss.insertSheet('Settings');
    sh.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'What it’s for']]);
    try { sh.setFrozenRows(1); } catch (e) {}
  }
  return sh;
}

function getRawSettings_() {
  var out = {};
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName('Settings');
    if (sh && sh.getLastRow() > 0) {
      var rows = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
      for (var i = 0; i < rows.length; i++) {
        var k = String(rows[i][0] || '').trim();
        if (k && (k in CONFIG_DEFAULTS_)) { var v = rows[i][1]; out[k] = (v == null ? '' : v); }
      }
    }
  } catch (e) {}
  return out;
}

function getSettings() {
  return { config: getConfig_(), raw: getRawSettings_(), fields: CONFIG_FIELDS_, sections: SETTINGS_SECTIONS_ };
}

function saveSettings(values) {
  var sh = settingsSheet_();
  var last = sh.getLastRow();
  var keyRow = {};
  if (last >= 1) {
    var col = sh.getRange(1, 1, last, 1).getValues();
    for (var i = 0; i < col.length; i++) {
      var k = String(col[i][0] || '').trim();
      if (k) keyRow[k] = i + 1;
    }
  }
  var help = {};
  CONFIG_FIELDS_.forEach(function (f) { help[f.key] = f.help; });
  Object.keys(values || {}).forEach(function (k) {
    if (!(k in CONFIG_DEFAULTS_)) return; // ignore unknown keys
    var v = values[k];
    if (v === null || typeof v === 'undefined') v = '';
    if (keyRow[k]) {
      sh.getRange(keyRow[k], 2).setValue(v);
    } else {
      sh.appendRow([k, v, help[k] || '']);
      keyRow[k] = sh.getLastRow();
    }
  });
  // A settings change (logo, business info, governing law) must rebuild the
  // cached contract template so the next contract/receipt reflects it.
  try { PropertiesService.getScriptProperties().deleteProperty('CONTRACT_TEMPLATE_ID_V3'); } catch (e) {}
  getConfig_._cache = null;
  return { ok: true, config: getConfig_() };
}

/* ---------- Licensing (sell + protect) ----------
 * Each sold copy carries a LICENSE_KEY (a Script Property the buyer enters when they
 * activate). The seller bakes their deployed license-hub URL into LICENSE_HUB_URL below
 * (a code constant, so it travels with a "make a copy"; Script Properties do not).
 * On load the app asks the hub whether this key is active for THIS install, caches the
 * answer, and re-checks daily. If the hub is unreachable it trusts the last verified
 * "active" for a grace window rather than locking out a paying customer. If no hub URL
 * is set, licensing is OFF (the app runs unlocked — for development and un-sold copies).
 */
var LICENSE_HUB_URL = 'https://script.google.com/macros/s/AKfycbyXC9FYWXjjadgGyxWqIAJIRhvWn_iZHEkiuc0M8-kEFrR9qFI0wlFMD9NZdRlW-icd/exec'; // SELLER: deployed license-hub web-app URL (licensing is ON).
var LICENSE_RECHECK_MS = 86400000;       // re-verify with the hub at most once a day
var LICENSE_GRACE_MS = 7 * 86400000;     // if the hub is unreachable, trust last-good "active" this long

// SELLER: bump this on every release you ship to the master template. The
// update banner shows when the hub's Meta "latestVersion" is higher than this.
// (Only copies made from a master that already had this checker will notice —
// the check can't be retro-added to code a customer already deployed.)
var APP_VERSION = '1.5.1';

function getInstallId_() {
  try { return ScriptApp.getScriptId(); } catch (e) {}
  try { return SpreadsheetApp.getActive().getId(); } catch (e) {}
  return 'unknown';
}

function licenseHubUrl_() {
  var p = '';
  try { p = (PropertiesService.getScriptProperties().getProperty('LICENSE_HUB_URL') || '').trim(); } catch (e) {}
  return p || LICENSE_HUB_URL;
}

function getLicenseState_() {
  var props = PropertiesService.getScriptProperties();
  var key = '';
  try { key = (props.getProperty('LICENSE_KEY') || '').trim(); } catch (e) {}
  var hub = licenseHubUrl_();
  if (!hub) return { ok: true, status: 'nohub', message: 'Licensing is off (no hub configured).' };
  if (!key) return { ok: false, status: 'unset', message: 'Enter your license key to activate this app.' };

  var now = Date.now();
  var cache = null;
  try { cache = JSON.parse(props.getProperty('LICENSE_CACHE') || 'null'); } catch (e) {}
  if (cache && cache.key === key && cache.status === 'active' && (now - cache.checkedAt) < LICENSE_RECHECK_MS) {
    return { ok: true, status: 'active', message: 'License active.', checkedAt: cache.checkedAt };
  }

  var resp = null;
  try {
    resp = UrlFetchApp.fetch(hub, {
      method: 'post',
      payload: { key: key, install: getInstallId_() },
      muteHttpExceptions: true,
      followRedirects: true
    });
  } catch (e) { resp = null; }

  if (resp) {
    var data = null;
    try { data = JSON.parse(resp.getContentText()); } catch (e) {}
    if (data) {
      // Capture the latest-version info the hub reports, so the update banner
      // has it even when later loads short-circuit on the cached license.
      try {
        if (data.latestVersion != null) props.setProperty('LATEST_VERSION', String(data.latestVersion));
        if (data.releaseNotes != null) props.setProperty('LATEST_NOTES', String(data.releaseNotes));
        if (data.updateUrl != null) props.setProperty('UPDATE_URL', String(data.updateUrl));
      } catch (e) {}
    }
    if (data && (data.status === 'active' || data.status === 'invalid' || data.status === 'revoked')) {
      props.setProperty('LICENSE_CACHE', JSON.stringify({ status: data.status, checkedAt: now, key: key }));
      if (data.status === 'active') return { ok: true, status: 'active', message: 'License active.', checkedAt: now };
      return { ok: false, status: data.status, message: data.message || ('This license is ' + data.status + '.') };
    }
  }

  // Hub unreachable or unexpected response — fail gracefully within the grace window.
  if (cache && cache.key === key && cache.status === 'active' && (now - cache.checkedAt) < LICENSE_GRACE_MS) {
    return { ok: true, status: 'grace', message: 'Could not reach the license server — using your last verified status.', checkedAt: cache.checkedAt };
  }
  return { ok: false, status: 'error', message: 'Could not verify your license. Check your connection and try again.' };
}

function activateLicense(key) {
  var props = PropertiesService.getScriptProperties();
  if (typeof key === 'string') props.setProperty('LICENSE_KEY', key.trim());
  try { props.deleteProperty('LICENSE_CACHE'); } catch (e) {}
  return getLicenseState_();
}

// Compares APP_VERSION against the latest version the hub reported (captured on
// the daily license check) and returns what the "update available" banner needs.
function getUpdateInfo() {
  var props = PropertiesService.getScriptProperties();
  var latest = (props.getProperty('LATEST_VERSION') || '').trim();
  return {
    current: APP_VERSION,
    latest: latest,
    notes: (props.getProperty('LATEST_NOTES') || '').trim(),
    url: (props.getProperty('UPDATE_URL') || '').trim(),
    updateAvailable: !!latest && versionLt_(APP_VERSION, latest)
  };
}

// True if dotted-numeric version a is strictly older than b. Compares parts
// numerically so "1.2.0" < "1.10.0" (a plain string compare would get this
// wrong). Missing trailing parts count as 0, so "1.0" < "1.0.1".
function versionLt_(a, b) {
  var pa = String(a).split('.'), pb = String(b).split('.');
  for (var i = 0; i < Math.max(pa.length, pb.length); i++) {
    var x = parseInt(pa[i], 10) || 0, y = parseInt(pb[i], 10) || 0;
    if (x < y) return true;
    if (x > y) return false;
  }
  return false;
}

// Menu: force an immediate hub check (the app otherwise re-verifies at most once a day)
// and report whether a newer version is out. Ages the cached check instead of deleting
// it, so if the hub is briefly unreachable the grace window still covers the user rather
// than locking them out.
function checkForUpdatesNow() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  try {
    var c = JSON.parse(props.getProperty('LICENSE_CACHE') || 'null');
    if (c) { c.checkedAt = Date.now() - LICENSE_RECHECK_MS - 1; props.setProperty('LICENSE_CACHE', JSON.stringify(c)); }
  } catch (e) {}
  var state = getLicenseState_();
  if (state.status === 'unset') {
    ui.alert('Check for updates', 'Enter your license key first (Magic Manager → Enter license key), then try again.', ui.ButtonSet.OK);
    return;
  }
  if (state.status === 'nohub') {
    ui.alert('Check for updates', 'Updates aren’t set up for this app.', ui.ButtonSet.OK);
    return;
  }
  if (state.status === 'revoked' || state.status === 'invalid') {
    ui.alert('License problem', state.message || 'This license is not active.', ui.ButtonSet.OK);
    return;
  }
  if (state.status === 'grace' || !state.ok) {
    ui.alert('Check for updates', 'Couldn’t reach the update server just now — check your connection and try again in a few minutes.', ui.ButtonSet.OK);
    return;
  }
  var info = getUpdateInfo();
  if (info.updateAvailable) {
    var msg = 'Version ' + info.latest + ' is available — you’re on ' + info.current + '.';
    if (info.notes) msg += '\n\nWhat’s new:\n' + info.notes;
    if (info.url) msg += '\n\nHow to update:\n' + info.url;
    ui.alert('Update available ✦', msg, ui.ButtonSet.OK);
  } else if (info.latest) {
    ui.alert('You’re up to date', 'You’re running the latest version (' + info.current + ').', ui.ButtonSet.OK);
  } else {
    ui.alert('Check for updates', 'Couldn’t read the latest version just now — please try again in a few minutes.', ui.ButtonSet.OK);
  }
}

function escHtml_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function activationPageHtml_(lic) {
  var biz = getConfig_().BUSINESS_NAME || 'this app';
  var msg = (lic && lic.message) ? lic.message : 'Enter your license key to activate this app.';
  return [
    '<!doctype html><html><head><meta charset="utf-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1">',
    '<title>', escHtml_(biz), ' — Activate</title>',
    '<style>',
    'body{margin:0;background:#0B0B10;color:#F4EBD3;font-family:Arial,Helvetica,sans-serif;display:flex;min-height:100vh;align-items:center;justify-content:center;padding:24px}',
    '.box{max-width:380px;width:100%;text-align:center}',
    'h1{color:#E4C179;font-weight:700;font-size:20px;margin:0 0 6px}',
    'p{color:#928A75;font-size:13.5px;line-height:1.6;margin:0 0 16px}',
    'input{width:100%;box-sizing:border-box;font-size:16px;color:#F4EBD3;background:#1C1C27;border:1px solid rgba(201,160,80,.4);border-radius:8px;padding:12px;outline:none;text-align:center;letter-spacing:.04em}',
    'button{width:100%;margin-top:12px;background:#C9A050;color:#141008;font-weight:700;font-size:15px;border:none;border-radius:10px;padding:13px;cursor:pointer}',
    'button:disabled{opacity:.5}',
    '.err{color:#E9927B;font-size:12.5px;margin-top:10px;min-height:16px}',
    '</style></head><body><div class="box">',
    '<h1>Activate ', escHtml_(biz), '</h1>',
    '<p>', escHtml_(msg), '</p>',
    '<input id="k" placeholder="License key" autocomplete="off" autocapitalize="off" spellcheck="false">',
    '<button id="b" onclick="act()">Activate</button>',
    '<div class="err" id="e"></div>',
    '<script>',
    'function act(){var k=document.getElementById("k").value.trim();if(!k)return;',
    'var b=document.getElementById("b");b.disabled=true;b.textContent="Checking…";document.getElementById("e").textContent="";',
    'google.script.run.withSuccessHandler(function(r){if(r&&r.ok){location.reload();}else{b.disabled=false;b.textContent="Activate";document.getElementById("e").textContent=(r&&r.message)||"That key could not be verified.";}})',
    '.withFailureHandler(function(){b.disabled=false;b.textContent="Activate";document.getElementById("e").textContent="Something went wrong — try again.";})',
    '.activateLicense(k);}',
    'document.getElementById("k").addEventListener("keydown",function(e){if(e.key==="Enter")act();});',
    '</script>',
    '</div></body></html>'
  ].join('');
}

/* ---------- One-time buyer setup (menu in the Sheet) ----------
 * Adds a "Magic Manager" menu so a new owner can get their private app link,
 * enter their license key, and add their Stripe key — without hand-editing code
 * or Script Properties.
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Magic Manager')
      .addItem('Get my app link', 'showAppLink')
      .addItem('Enter license key', 'promptLicenseKey')
      .addItem('Enter Stripe secret key', 'promptStripeKey')
      .addSeparator()
      .addItem('Setup status', 'showSetupStatus')
      .addItem('Check for updates now', 'checkForUpdatesNow')
      .addItem('Reset app URL', 'resetAppUrl')
      .addToUi();
  } catch (e) {}
}

function appAccessKey_() {
  var props = PropertiesService.getScriptProperties();
  var k = (props.getProperty('APP_ACCESS_KEY') || '').trim();
  if (!k) {
    k = 'AK-' + Utilities.getUuid().replace(/-/g, '').substring(0, 20);
    props.setProperty('APP_ACCESS_KEY', k);
  }
  return k;
}

function appUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

function showAppLink() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var url = (props.getProperty('WEBAPP_URL') || '').trim();
  if (!url) {
    var r = ui.prompt('Set your app link (one time)',
      'Paste your Web app URL — the link that ends in /exec.\n\nFind it in the Apps Script editor under Deploy → Manage deployments (copy the Web app URL), then paste it here.',
      ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    url = (r.getResponseText() || '').trim().replace(/\?.*$/, '').replace(/\/+$/, '');
    if (!/^https:\/\/script\.google\.com\/macros\/s\/[^\/]+\/exec$/.test(url)) {
      ui.alert('That URL doesn’t look right',
        'Please paste the full Web app URL ending in /exec (from Deploy → Manage deployments).',
        ui.ButtonSet.OK);
      return;
    }
    props.setProperty('WEBAPP_URL', url);
  }
  var full = url + '?key=' + encodeURIComponent(appAccessKey_());
  ui.alert('Your private app link',
    'Open and bookmark this link to use your app:\n\n' + full + '\n\nKeep it private — anyone with this link can open your app.',
    ui.ButtonSet.OK);
}

function resetAppUrl() {
  var ui = SpreadsheetApp.getUi();
  PropertiesService.getScriptProperties().deleteProperty('WEBAPP_URL');
  ui.alert('App URL cleared', 'Run "Get my app link" again to enter your current Web app URL (needed if you created a brand-new deployment).', ui.ButtonSet.OK);
}

function promptLicenseKey() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Enter license key', 'Paste the license key you received when you purchased:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var res = activateLicense(r.getResponseText());
  ui.alert(res.ok ? 'License activated ✦' : 'Not activated', res.message || '', ui.ButtonSet.OK);
}

function promptStripeKey() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Enter Stripe secret key', 'Paste your Stripe SECRET key (starts with "sk_"). Needed only if you want to accept card payments:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var k = (r.getResponseText() || '').trim();
  if (k) PropertiesService.getScriptProperties().setProperty('STRIPE_SECRET_KEY', k);
  ui.alert(k ? 'Stripe key saved ✦' : 'Nothing entered', k ? 'Card payments are ready to use.' : '', ui.ButtonSet.OK);
}

function showSetupStatus() {
  var ui = SpreadsheetApp.getUi();
  var cfg = getConfig_();
  var props = PropertiesService.getScriptProperties();
  var configured = cfg.BUSINESS_NAME && cfg.BUSINESS_NAME !== 'Your Business';
  var lines = [
    'Business info: ' + (configured ? cfg.BUSINESS_NAME + ' ✓' : 'not set yet (open the app → Setup)'),
    'License: ' + getLicenseState_().status,
    'Stripe (card payments): ' + (props.getProperty('STRIPE_SECRET_KEY') ? 'set ✓' : 'not set'),
    'App link: ' + (props.getProperty('WEBAPP_URL') ? 'set ✓' : 'run "Get my app link" to set it')
  ];
  ui.alert('Setup status', lines.join('\n'), ui.ButtonSet.OK);
}
const REMINDER_HOUR = 9; // follow-up reminders fire at 9:00 AM
const TRASH_SHEET = 'Leads Trash';
const TRASH_RETENTION_DAYS = 30;

/* ---------- Web app ---------- */

function doGet(e) {
  if (e && e.parameter && e.parameter.sign) {
    // Hand the token to the page directly rather than relying on the
    // client re-reading it from the URL — Apps Script renders web app
    // content inside a sandboxed iframe with its own URL, so window.location
    // inside that iframe does not reliably reflect the outer request's
    // query string.
    var template = HtmlService.createTemplateFromFile('Sign');
    template.token = e.parameter.sign;
    template.paid = e.parameter.paid || '';
    template.cfg = getConfig_();
    return template.evaluate()
      .setTitle(getConfig_().BUSINESS_NAME + ' — Sign Agreement')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // The deployment itself has to be set to "Anyone" so clients without a
  // Google account can reach the sign page above — but that same setting
  // would otherwise let anyone who has (or guesses) the base URL load the
  // full app, with every lead, contact, and deposit in it, no login at
  // all. This is the actual lock on the main app: it only loads for a
  // request carrying the correct private key, set once in Script
  // Properties (APP_ACCESS_KEY) and matched against ?key= in the URL.
  var appKey = PropertiesService.getScriptProperties().getProperty('APP_ACCESS_KEY');
  var providedKey = (e && e.parameter && e.parameter.key) ? e.parameter.key : '';
  if (!appKey || providedKey !== appKey) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:60px 20px;text-align:center;color:#888">This page isn\u2019t available.</div>'
    ).setTitle('Not Found');
  }

  var lic = getLicenseState_();
  if (!lic.ok) {
    return HtmlService.createHtmlOutput(activationPageHtml_(lic))
      .setTitle(getConfig_().BUSINESS_NAME + ' — Activate')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var indexTpl = HtmlService.createTemplateFromFile('Index');
  indexTpl.cfg = getConfig_();
  return indexTpl.evaluate()
    .setTitle(getConfig_().BUSINESS_NAME + ' — Leads')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- Contract generation ---------- */
// The template is a real Google Doc, auto-created once and remembered by ID
// (same self-healing pattern as the Calendar) — so it's editable directly
// in Docs afterward (logo, exact wording tweaks) without touching code.
// Single source of truth for the base dollar amounts a client owes on a
// booking — quoted price, travel fee, and the deposit/balance split that
// results from adding them together. Card-processing-fee math and
// payment-method handling vary by call site (contract text, the live
// signing page, receipts, Stripe checkout) so those stay local to each
// one, but the underlying total-due math — the part that was actually
// wrong before, silently leaving travel fee out of what a client was
// asked to pay — now has exactly one place it can be computed, instead of
// four separate chances to drift out of sync with each other.
function computeBaseAmounts_(get) {
  const price = Number(String(get('Quoted Price') || '').replace(/[^0-9.]/g, '')) || 0;
  const travelFee = Number(String(get('Travel Fee') || '').replace(/[^0-9.]/g, '')) || 0;
  const totalDue = price + travelFee;
  const depositRaw = Number(String(get('Deposit Amount') || '').replace(/[^0-9.]/g, '')) || 0;
  const deposit = depositRaw || (totalDue / 2);
  const balance = Math.max(0, totalDue - deposit);
  return { price: price, travelFee: travelFee, totalDue: totalDue, deposit: deposit, balance: balance };
}
function contractsFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('CONTRACTS_FOLDER_ID');
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (e) {}
  }
  const it = DriveApp.getFoldersByName(getConfig_().CONTRACTS_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(getConfig_().CONTRACTS_FOLDER);
  props.setProperty('CONTRACTS_FOLDER_ID', folder.getId());
  return folder;
}
function contractTemplate_() {
  const props = PropertiesService.getScriptProperties();
  // Versioned property key — bumped whenever the template's structure
  // changes (like this fee-table rework), so the next contract generated
  // automatically builds a fresh template instead of reusing an old cached
  // copy that doesn't have the new rows/tokens. Nothing to do by hand.
  const savedId = props.getProperty('CONTRACT_TEMPLATE_ID_V3');
  if (savedId) {
    try { return DocumentApp.openById(savedId); } catch (e) {}
  }
  const doc = buildContractTemplate_();
  props.setProperty('CONTRACT_TEMPLATE_ID_V3', doc.getId());
  return doc;
}

// Returns the buyer's configured logo as an image Blob for the contract/receipt,
// or null if none is set or it can't be loaded. Handles an uploaded data URI
// (data:image/...;base64,...) and a pasted, publicly-reachable image URL.
function getLogoBlob_() {
  var url = (getConfig_().LOGO_URL || '').trim();
  if (!url) return null;
  try {
    var m = url.match(/^data:([^;,]*)(;base64)?,(.*)$/i);
    if (m) {
      var mime = m[1] || 'image/png';
      var bytes = m[2] ? Utilities.base64Decode(m[3]) : Utilities.newBlob(decodeURIComponent(m[3])).getBytes();
      return Utilities.newBlob(bytes, mime, 'logo');
    }
    if (/^https?:\/\//i.test(url)) {
      return UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true }).getBlob().setName('logo');
    }
  } catch (e) {}
  return null;
}

// Returns the buyer's logo as a browser-ready data: URI (or ''), for pages served to the
// CLIENT (the Sign page). The raw LOGO_URL can be a link only the SERVER can load — a Google
// Drive / Dropbox share link, or an image behind the buyer's own Google login — which fetches
// fine for the contract but breaks in a client's <img> (the client isn't logged into the
// buyer's account). Embedding the bytes the way the contract does makes the logo show for
// everyone. A data: URI is already browser-ready and passes straight through; an external URL
// is fetched once (via getLogoBlob_) and cached to avoid refetching on every page load.
function getLogoDataUri_() {
  var url = (getConfig_().LOGO_URL || '').trim();
  if (!url) return '';
  // A stored data: URI near the sheet's ~50k-per-cell limit is almost certainly truncated —
  // an incomplete image still renders (partially) in the Doc contract but shows as a broken
  // image in a browser <img>. Don't serve it. (A valid uploaded logo is well under this.)
  if (/^data:/i.test(url) && url.length > 49000) return '';
  var cache = null, ckey = 'logoDataUri_' + url.length + '_' + url.slice(-32);
  try { cache = CacheService.getScriptCache(); var hit = cache.get(ckey); if (hit != null) return hit; } catch (e) {}
  var out = '';
  try {
    var b = getLogoBlob_();
    if (b) {
      var ct = String(b.getContentType() || ''), bytes = b.getBytes();
      // Only embed a real, non-trivial image — never an HTML page (a Drive/Dropbox "view" link
      // returns HTML, not the file) or an empty/tiny blob, both of which break the <img>.
      if (/^image\//i.test(ct) && bytes && bytes.length > 200) {
        out = 'data:' + ct + ';base64,' + Utilities.base64Encode(bytes);
      }
    }
  } catch (e) {}
  if (cache && out) { try { cache.put(ckey, out, 21600); } catch (e) {} }
  return out;
}

// US states (canonical name -> 2-letter abbr). Used to derive the contract's governing-law
// state from the buyer's own business address, so onboarding never has to ask for it.
var US_STATES_ = {
  'Alabama':'AL','Alaska':'AK','Arizona':'AZ','Arkansas':'AR','California':'CA','Colorado':'CO',
  'Connecticut':'CT','Delaware':'DE','District of Columbia':'DC','Florida':'FL','Georgia':'GA',
  'Hawaii':'HI','Idaho':'ID','Illinois':'IL','Indiana':'IN','Iowa':'IA','Kansas':'KS','Kentucky':'KY',
  'Louisiana':'LA','Maine':'ME','Maryland':'MD','Massachusetts':'MA','Michigan':'MI','Minnesota':'MN',
  'Mississippi':'MS','Missouri':'MO','Montana':'MT','Nebraska':'NE','Nevada':'NV','New Hampshire':'NH',
  'New Jersey':'NJ','New Mexico':'NM','New York':'NY','North Carolina':'NC','North Dakota':'ND',
  'Ohio':'OH','Oklahoma':'OK','Oregon':'OR','Pennsylvania':'PA','Rhode Island':'RI','South Carolina':'SC',
  'South Dakota':'SD','Tennessee':'TN','Texas':'TX','Utah':'UT','Vermont':'VT','Virginia':'VA',
  'Washington':'WA','West Virginia':'WV','Wisconsin':'WI','Wyoming':'WY'
};
// Best-effort: pull a US state from a freeform address, anchored at the END (where the state
// sits) so a street like "1600 Pennsylvania Ave, Washington DC" resolves to DC, not PA. Tries
// a trailing full state name first (longest-first, so "West Virginia" beats "Virginia"), then
// a trailing 2-letter abbreviation. Returns the canonical state name, or '' if none found.
function usStateFromAddress_(addr) {
  var s = String(addr == null ? '' : addr).trim();
  if (!s) return '';
  var abbrToName = {}; for (var k in US_STATES_) abbrToName[US_STATES_[k]] = k;
  var tail = s.replace(/[\s,]*\d{5}(?:-\d{4})?\s*$/, '').replace(/[\s,]+$/, ''); // drop a trailing ZIP
  var names = Object.keys(US_STATES_).sort(function (a, b) { return b.length - a.length; });
  for (var i = 0; i < names.length; i++) {
    if (new RegExp('(^|[\\s,])' + names[i].replace(/ /g, '\\s+') + '$', 'i').test(tail)) return names[i];
  }
  var m = tail.match(/[\s,]([A-Za-z]{2})\.?$/) || tail.match(/^([A-Za-z]{2})\.?$/);
  if (m && abbrToName[m[1].toUpperCase()]) return abbrToName[m[1].toUpperCase()];
  return '';
}
// The phrase dropped into the contract's Governing Law clause. An explicit GOVERNING_LAW
// setting still wins (kept for existing installs / non-US buyers who set one); otherwise it's
// derived from the buyer's business-address state; otherwise a generic fallback.
function governingLawPhrase_() {
  var cfg = getConfig_();
  var explicit = (cfg.GOVERNING_LAW || '').trim();
  if (explicit) return explicit;
  var st = usStateFromAddress_(cfg.ADDRESS || '');
  if (st === 'District of Columbia') return 'the District of Columbia';
  if (st) return 'the State of ' + st;
  return 'the state or country in which the Service Provider operates';
}
/**
 * Rebuilt to match a hand-designed layout (two-column header,
 * side-by-side Parties/Event Schedule, a real fee table, two-column Terms,
 * a signature table) rather than the original simple linear one. Every
 * {{token}} stays exactly the same as before, so none of the merge logic
 * in buildMergedContract_ needed to change — only this layout did.
 */
function buildContractTemplate_() {
  const doc = DocumentApp.create(getConfig_().BUSINESS_NAME + ' - Contract Template');
  const body = doc.getBody();
  body.clear();
  body.setMarginTop(26).setMarginBottom(26).setMarginLeft(34).setMarginRight(34);

  const FONT = 'Arial';
  const BODY_SIZE = 9.5;
  const SMALL_SIZE = 8.5;
  const HEAD_SIZE = 11;

  function styleText(text, size) {
    text.setFontFamily(FONT).setFontSize(size || BODY_SIZE);
    return text;
  }
  // Apps Script refuses to insert a paragraph with completely empty text
  // ("Cannot insert an empty text element"), so every paragraph below is
  // always created WITH its real final text already in place, and bold
  // ranges are applied afterward — never built up from an empty starting
  // point. appendTable() seeds each cell with exactly one paragraph; the
  // first line in a cell reuses that seed paragraph via setText() rather
  // than appending a new one.
  function cellLine(cell, isFirst, text) {
    if (isFirst) {
      var p = cell.getChild(0).asParagraph();
      p.setText(text);
      return p;
    }
    return cell.appendParagraph(text);
  }
  function plainLine(cell, value, opts) {
    opts = opts || {};
    var p = cellLine(cell, !!opts.first, value);
    p.setSpacingAfter(opts.spacing != null ? opts.spacing : 2);
    var t = p.editAsText();
    styleText(t, opts.size);
    if (opts.bold) t.setBold(true);
    return p;
  }
  function boldLine(cell, label, value, opts) {
    opts = opts || {};
    var full = label + value;
    var p = cellLine(cell, !!opts.first, full);
    p.setSpacingAfter(2);
    var t = p.editAsText();
    t.setBold(0, label.length - 1, true);
    styleText(t, opts.size);
    return p;
  }
  function headingPara(text, spacingBefore) {
    var p = body.appendParagraph(text);
    var t = p.editAsText();
    t.setBold(true); styleText(t, HEAD_SIZE);
    p.setSpacingAfter(3);
    if (spacingBefore != null) p.setSpacingBefore(spacingBefore);
    return p;
  }
  function bulletTerm(cell, label, text, isFirst) {
    var lead = '\u2022 ' + label + ': ';
    var full = lead + text;
    var p = cellLine(cell, !!isFirst, full);
    p.setSpacingAfter(4);
    var t = p.editAsText();
    t.setBold(0, lead.length - 1, true);
    styleText(t, SMALL_SIZE);
    return p;
  }

  // ---- Header: logo | company info | contract title ----
  var header = body.appendTable([['', '', '']]);
  header.setBorderWidth(0.75);
  var hRow = header.getRow(0);

  var logoCell = hRow.getCell(0);
  var logoPara = logoCell.getChild(0).asParagraph();
  logoPara.setText(' ');
  try {
    var logoBlob = getLogoBlob_();
    if (logoBlob) {
      var img = logoPara.appendInlineImage(logoBlob);
      img.setWidth(50); img.setHeight(50);
    }
  } catch (e) {}
  logoCell.setWidth(58);

  var infoCell = hRow.getCell(1);
  plainLine(infoCell, getConfig_().BUSINESS_LEGAL_NAME, { bold: true, size: 12, spacing: 1, first: true });
  plainLine(infoCell, getConfig_().TAGLINE, { spacing: 1 });
  plainLine(infoCell, getConfig_().OWNER_NAME + ' | Service Provider', { spacing: 0 });
  infoCell.setWidth(230);

  var titleCell = hRow.getCell(2);
  plainLine(titleCell, 'Performance Contract', { bold: true, size: 12, spacing: 2, first: true });
  boldLine(titleCell, 'Agreement Date: ', '{{agreementdate}}');
  boldLine(titleCell, 'Invoice #: ', '{{invoicenumber}}');

  body.appendParagraph(' ').setSpacingAfter(4);

  // ---- 1. Parties to Agreement | 2. Event Schedule (side by side) ----
  var s12 = body.appendTable([['', '']]);
  s12.setBorderWidth(0.75);
  var s12Row = s12.getRow(0);

  var partiesCell = s12Row.getCell(0);
  plainLine(partiesCell, '1. Parties to Agreement', { bold: true, size: HEAD_SIZE, spacing: 3, first: true });
  boldLine(partiesCell, 'Client: ', '{{client}}');
  boldLine(partiesCell, 'Service Provider: ', getConfig_().OWNER_NAME + ' (' + getConfig_().BUSINESS_LEGAL_NAME + ')');
  boldLine(partiesCell, 'Location: ', '{{location}}');

  var scheduleCell = s12Row.getCell(1);
  plainLine(scheduleCell, '2. Event Schedule', { bold: true, size: HEAD_SIZE, spacing: 3, first: true });
  boldLine(scheduleCell, 'Event Date: ', '{{eventdate}}');
  boldLine(scheduleCell, 'Performance Time: ', '{{performancetime}}');
  boldLine(scheduleCell, 'Services Offered: ', '{{services}}');

  body.appendParagraph(' ').setSpacingAfter(4);

  // ---- 3. Investment & Fee Schedule ----
  headingPara('3. Investment & Fee Schedule');
  var feeTable = body.appendTable([
    ['Performance Fee', '{{price}}'],
    ['Travel & Additional Fees', '{{travelfeeline}}'],
    ['Total Investment', '{{totalinvestment}}'],
    ['{{depositlabel}}', '{{deposit}}'],
    ['Balance (Remaining)', '{{balance}}']
  ]);
  feeTable.setBorderWidth(0.75);
  for (var r = 0; r < feeTable.getNumRows(); r++) {
    for (var c = 0; c < 2; c++) {
      var fc = feeTable.getRow(r).getCell(c);
      var ft = fc.editAsText();
      styleText(ft, BODY_SIZE);
      if (c === 0) ft.setBold(true);
    }
    feeTable.getRow(r).getCell(1).setWidth(150);
  }

  var pmLabel = 'Payment Method: ';
  var pmFull = pmLabel + '{{paymentmethod}}{{cardfeenote}}';
  var pmLine = body.appendParagraph(pmFull);
  pmLine.setSpacingBefore(4).setSpacingAfter(6);
  var pmText = pmLine.editAsText();
  pmText.setBold(0, pmLabel.length - 1, true);
  styleText(pmText, BODY_SIZE);

  // ---- 4. Terms & Conditions (two columns of bullets, no visible borders) ----
  headingPara('4. Terms & Conditions');
  var terms = body.appendTable([['', '']]);
  terms.setBorderWidth(0);
  var termsRow = terms.getRow(0);
  var leftTerms = termsRow.getCell(0);
  var rightTerms = termsRow.getCell(1);

  bulletTerm(leftTerms, 'Confidentiality & Conduct', 'Both parties agree to maintain strict professional conduct and confidentiality throughout the term of service.', true);
  bulletTerm(leftTerms, 'Cancellation & Deposit', 'Deposits are non-refundable upon client event cancellation. Client performance cancellations allow deposit application to an alternate date. Unforeseen performer cancellation yields a full refund.');
  bulletTerm(leftTerms, 'Liability', 'Performer assumes no responsibility for damages or injuries during performance except in cases of gross negligence.');

  bulletTerm(rightTerms, 'Media & Photo Release', 'Photographs of the performance may be taken freely by the Client and guests. Video recording or filming of the performance is not permitted without the Service Provider\u2019s prior consent. The Service Provider may photograph and film the performance for promotional use, which the Client agrees to unless an opt-out notification is provided in writing prior to the event.', true);
  var govLoc = governingLawPhrase_();
  bulletTerm(rightTerms, 'Governing Law', 'This Agreement shall be governed and construed in accordance with the laws of ' + govLoc + '.');
  bulletTerm(rightTerms, 'Entire Agreement', 'Supersedes all prior discussions. Any modifications must be executed in writing and signed by both parties.');

  // ---- 5. Acceptance & Authorization ----
  headingPara('5. Acceptance & Authorization', 4);
  var acceptIntro = body.appendParagraph('By signing below, the Client and Service Provider acknowledge, accept, and agree to all terms and conditions set forth in this Agreement.');
  styleText(acceptIntro.editAsText(), BODY_SIZE);
  acceptIntro.setSpacingAfter(6);

  var sigTable = body.appendTable([['', '']]);
  sigTable.setBorderWidth(0.75);
  var sigRow = sigTable.getRow(0);

  var clientSigCell = sigRow.getCell(0);
  boldLine(clientSigCell, 'Client Signature: ', '{{clientsignature}}', { first: true });
  boldLine(clientSigCell, 'Date: ', '{{signdate}}');

  var providerSigCell = sigRow.getCell(1);
  var provLabel = 'Service Provider Signature: ';
  var provFull = provLabel + getConfig_().OWNER_NAME;
  var provLine = cellLine(providerSigCell, true, provFull);
  provLine.setSpacingAfter(2);
  var provText = provLine.editAsText();
  provText.setBold(0, provLabel.length - 1, true);
  styleText(provText, BODY_SIZE);
  // Best-effort script styling on just the name — falls back gracefully
  // to plain text if this particular font isn't available.
  try {
    var nameStart = provLabel.length;
    var nameEnd = provText.getText().length - 1;
    provText.setFontFamily(nameStart, nameEnd, 'Dancing Script');
    provText.setFontSize(nameStart, nameEnd, 13);
  } catch (e) {}
  boldLine(providerSigCell, 'Date: ', '{{providersigndate}}');

  doc.saveAndClose();
  return doc;
}

function ensureContractCols_(sh) {
  const heads = headers_(sh);
  function ensure(name) {
    var col = heads.indexOf(name) + 1;
    if (!col) { col = heads.length + 1; sh.getRange(1, col).setValue(name); heads.push(name); }
    return col;
  }
  return {
    docIdCol: ensure('Contract Doc ID'),
    tokenCol: ensure('Contract Sign Token'),
    signedCol: ensure('Contract Signed'),
    signedNameCol: ensure('Contract Signed Name'),
    signedDateCol: ensure('Contract Signed Date'),
    pdfUrlCol: ensure('Contract PDF URL'),
    docUrlCol: ensure('Contract Doc URL'),
    signUrlCol: ensure('Contract Sign URL'),
    depositReceiptCol: ensure('Deposit Receipt PDF URL'),
    balanceReceiptCol: ensure('Balance Receipt PDF URL'),
    invoiceNumCol: ensure('Invoice Number'),
    balancePmCol: ensure('Balance Payment Method')
  };
}
// Assigned exactly once, the first time a contract is generated for a
// lead, then reused forever after — including through regeneration at
// signing time — so it stays a stable reference number for the whole
// booking. Uses a real lock around the read-increment-write, so two
// contracts generated at nearly the same instant can never be handed the
// same number.
function ensureInvoiceNumber_(sh, rowNum, cols) {
  var existing = sh.getRange(rowNum, cols.invoiceNumCol).getValue();
  if (existing) return String(existing);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var again = sh.getRange(rowNum, cols.invoiceNumCol).getValue();
    if (again) return String(again);
    var props = PropertiesService.getScriptProperties();
    var year = new Date().getFullYear();
    var key = 'INVOICE_SEQ_' + year;
    var seq = (Number(props.getProperty(key)) || 0) + 1;
    props.setProperty(key, String(seq));
    var num = 'CM-' + year + '-' + String(seq).padStart(4, '0');
    sh.getRange(rowNum, cols.invoiceNumCol).setValue(num);
    return num;
  } finally {
    lock.releaseLock();
  }
}
function fmtMoney_(n) {
  return '$' + (Math.round(n * 100) / 100).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
/**
 * Builds a simple one-page payment receipt as a real PDF — deliberately
 * much lighter than the contract template (no persistent editable Doc to
 * maintain), since a receipt never needs manual tweaking the way a
 * contract layout does. Called automatically the moment "Deposit
 * Received" or "Balance Paid" gets flipped to Yes; never blocks a save
 * if anything here fails.
 */
function generateReceiptPdf_(kind, get) {
  const clientName = get('Company or Organization') || get('Customer Name') || 'Client';
  const eventDateVal = get('Date of Event');
  const eventDateFmt = eventDateVal ? Utilities.formatDate(new Date(eventDateVal), tz_(), 'M/d/yyyy') : '';
  const amounts = computeBaseAmounts_(get);
  const price = amounts.price;
  const deposit = amounts.deposit;
  const balance = amounts.balance;
  const depositPaymentMethod = String(get('Payment Method') || '');
  // Deliberately does NOT fall back to the deposit's payment method — the
  // balance is often paid separately, sometimes much later, and may not
  // end up going through the same method as the deposit. Only an
  // explicitly recorded Balance Payment Method counts toward the balance
  // fee; if it's never set, the balance is treated as fee-free rather
  // than assumed from a field that may no longer reflect reality.
  const balancePaymentMethod = String(get('Balance Payment Method') || '');
  const paymentMethod = kind === 'deposit' ? depositPaymentMethod : balancePaymentMethod;
  const isDepositCard = depositPaymentMethod === 'Credit Card' || depositPaymentMethod === 'Card';
  const isBalanceCard = balancePaymentMethod === 'Credit Card' || balancePaymentMethod === 'Card';
  const feeRate = 0.0325;
  const depositFinal = isDepositCard ? deposit * (1 + feeRate) : deposit;
  const balanceFinal = isBalanceCard ? balance * (1 + feeRate) : balance;
  const amount = kind === 'deposit' ? depositFinal : balanceFinal;
  const paidDateRaw = kind === 'deposit' ? get('Deposit Received Date') : get('Balance Paid Date');
  const paidDateFmt = paidDateRaw ? Utilities.formatDate(new Date(paidDateRaw), tz_(), 'M/d/yyyy') : Utilities.formatDate(new Date(), tz_(), 'M/d/yyyy');

  const folder = contractsFolder_();
  const fileName = 'Receipt - ' + (kind === 'deposit' ? 'Deposit' : 'Balance') + ' - ' + (get('Customer Name') || 'Client') + (eventDateFmt ? ' - ' + eventDateFmt : '');

  const doc = DocumentApp.create(fileName);
  const body = doc.getBody();
  body.clear();
  body.setMarginTop(40).setMarginBottom(40).setMarginLeft(50).setMarginRight(50);
  const FONT = 'Arial';
  const BODY_SIZE = 10.5;

  function styleText(text, size) { text.setFontFamily(FONT).setFontSize(size || BODY_SIZE); return text; }
  function plainLine(cell, value, opts) {
    opts = opts || {};
    var p = cell.appendParagraph(value);
    p.setSpacingAfter(opts.spacing != null ? opts.spacing : 2);
    var t = p.editAsText();
    styleText(t, opts.size);
    if (opts.bold) t.setBold(true);
    if (opts.color) t.setForegroundColor(opts.color);
    return p;
  }

  // ---- Header: logo | company info | document title — same layout as the contract ----
  const header = body.appendTable([['', '', '']]);
  header.setBorderWidth(0.75);
  const hRow = header.getRow(0);

  const logoCell = hRow.getCell(0);
  logoCell.clear();
  try {
    const logoBlob = getLogoBlob_();
    if (logoBlob) {
      const img = logoCell.appendImage(logoBlob);
      img.setWidth(50); img.setHeight(50);
    }
  } catch (e) {}
  logoCell.setWidth(58);

  const infoCell = hRow.getCell(1);
  infoCell.clear();
  plainLine(infoCell, getConfig_().BUSINESS_LEGAL_NAME, { bold: true, size: 12, spacing: 1 });
  plainLine(infoCell, getConfig_().ADDRESS, { spacing: 1, size: 9 });
  plainLine(infoCell, 'crawfordmagic@gmail.com \u00b7 678.960.8278', { spacing: 0, size: 9 });
  infoCell.setWidth(230);

  const titleCell = hRow.getCell(2);
  titleCell.clear();
  plainLine(titleCell, 'Payment Receipt', { bold: true, size: 12, spacing: 2 });
  plainLine(titleCell, kind === 'deposit' ? 'Deposit' : 'Final Balance', { bold: true, size: 11, color: '#8a6d1f', spacing: 0 });

  body.appendParagraph('').setSpacingAfter(6);

  // ---- Payment details table ----
  const detailsTable = body.appendTable([
    ['Invoice #', String(get('Invoice Number') || '\u2014')],
    ['Client', clientName],
    ['Event Date', eventDateFmt],
    ['Services', String(get('Service') || '')],
    ['Payment Method', paymentMethod || '\u2014'],
    ['Date Paid', paidDateFmt]
  ]);
  detailsTable.setBorderWidth(0.75);
  for (var r = 0; r < detailsTable.getNumRows(); r++) {
    for (var c = 0; c < 2; c++) {
      var cell = detailsTable.getRow(r).getCell(c);
      var t = cell.editAsText();
      styleText(t);
      if (c === 0) t.setBold(true);
    }
    detailsTable.getRow(r).getCell(1).setWidth(220);
  }

  body.appendParagraph('').setSpacingAfter(8);

  // ---- Highlighted amount block ----
  const amtTable = body.appendTable([['Amount Paid', fmtMoney_(amount)]]);
  amtTable.setBorderWidth(0.75);
  const amtRow = amtTable.getRow(0);
  amtRow.getCell(0).setBackgroundColor('#f4ebd3');
  amtRow.getCell(1).setBackgroundColor('#f4ebd3');
  var amtLabelText = amtRow.getCell(0).editAsText();
  amtLabelText.setBold(true); styleText(amtLabelText, 13); amtLabelText.setForegroundColor('#141008');
  var amtValueText = amtRow.getCell(1).editAsText();
  amtValueText.setBold(true); styleText(amtValueText, 15); amtValueText.setForegroundColor('#141008');
  amtRow.getCell(1).setWidth(220);

  if (kind === 'balance') {
    var depLine = body.appendParagraph('');
    depLine.setSpacingBefore(8).setSpacingAfter(2);
    var depT = depLine.editAsText();
    var depLabel = 'Deposit Previously Paid: ';
    depT.appendText(depLabel); depT.setBold(0, depLabel.length - 1, true);
    depT.appendText(fmtMoney_(depositFinal)); styleText(depT, 9.5);

    var totLine = body.appendParagraph('');
    totLine.setSpacingAfter(4);
    var totT = totLine.editAsText();
    var totLabel = 'Total Paid: ';
    totT.appendText(totLabel); totT.setBold(0, totLabel.length - 1, true);
    totT.appendText(fmtMoney_(depositFinal + balanceFinal)); styleText(totT, 9.5);
  }

  // ---- Signature-style closing ----
  const closing = body.appendParagraph('Looking forward to being with you for your event,');
  closing.editAsText().setItalic(true).setFontFamily(FONT).setFontSize(10.5);
  closing.setSpacingBefore(20).setSpacingAfter(2);

  const signOff = body.appendParagraph(getConfig_().OWNER_NAME);
  const signOffText = signOff.editAsText();
  signOffText.setFontFamily(FONT).setFontSize(13);
  try { signOffText.setFontFamily('Dancing Script'); } catch (e) {}

  doc.saveAndClose();
  const pdfBlob = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
  const pdfFile = folder.createFile(pdfBlob).setName(fileName + '.pdf');
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // Only the PDF is needed going forward — clean up the working Doc.
  try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch (e) {}
  return pdfFile.getUrl();
}
function maybeGenerateReceipt_(sh, rowNum, kind) {
  const heads = headers_(sh);
  const row = sh.getRange(rowNum, 1, 1, heads.length).getValues()[0];
  const get = function (name) { var idx = heads.indexOf(name); return idx > -1 ? row[idx] : ''; };
  const url = generateReceiptPdf_(kind, get);
  const cols = ensureContractCols_(sh);
  sh.getRange(rowNum, kind === 'deposit' ? cols.depositReceiptCol : cols.balanceReceiptCol).setValue(url);
}
// Shared merge logic used both when first generating a preview contract and
// when finalizing it at signing time. Google Docs' replaceText() is
// one-directional — once a {{token}} is substituted it can't be swapped
// again — so a genuinely final version (with the client's actual chosen
// payment method) is built as a fresh copy from the template rather than
// trying to edit an already-merged document.
function buildMergedContract_(get, fileNameSuffix, overridePaymentMethod, signedName, invoiceNumber) {
  const clientName = get('Company or Organization') || get('Customer Name') || 'Client';
  const eventDateVal = get('Date of Event');
  const eventDateFmt = eventDateVal ? Utilities.formatDate(new Date(eventDateVal), tz_(), 'M/d/yyyy') : '';
  const startVal = get('Start Time');
  const startFmt = startVal ? Utilities.formatDate(new Date(startVal), tz_(), 'h:mm a') : '';

  const amounts = computeBaseAmounts_(get);
  const price = amounts.price;
  const travelFee = amounts.travelFee;
  const totalDue = amounts.totalDue;
  const deposit = amounts.deposit;
  const balance = amounts.balance;
  const paymentMethod = overridePaymentMethod || String(get('Payment Method') || '');
  const isCard = paymentMethod === 'Credit Card' || paymentMethod === 'Card'; // 'Card' kept for any leads saved before the rename
  // Deliberately its own separate check, not reused from the deposit's
  // isCard — the deposit's payment method reflects a real, live choice
  // made at the time of that specific payment, but the balance is often
  // paid separately, sometimes much later, and may not go through the
  // same method. Only an explicitly recorded Balance Payment Method
  // counts toward the balance fee; otherwise the balance shown on the
  // contract stays at its plain, fee-free amount rather than assuming.
  const isBalanceCard = String(get('Balance Payment Method') || '') === 'Credit Card' || String(get('Balance Payment Method') || '') === 'Card';
  const feeRate = 0.0325;
  const depositFinal = isCard ? deposit * (1 + feeRate) : deposit;
  const balanceFinal = isBalanceCard ? balance * (1 + feeRate) : balance;
  // Reflects reality for clients being caught up retroactively — someone
  // who already paid shouldn't have their contract still say the deposit
  // is due today to lock the date, since it plainly isn't anymore.
  const depositReceived = String(get('Deposit Received') || '') === 'Yes';
  const depositReceivedDateRaw = get('Deposit Received Date');
  const depositLabel = depositReceived
    ? 'Deposit (Received' + (depositReceivedDateRaw ? ' ' + Utilities.formatDate(new Date(depositReceivedDateRaw), tz_(), 'M/d/yyyy') : '') + ')'
    : 'Deposit Required (Due Today to Lock Date)';

  const template = contractTemplate_();
  const folder = contractsFolder_();
  const fileName = 'Contract - ' + (get('Customer Name') || 'Client') + (eventDateFmt ? ' - ' + eventDateFmt : '') + (fileNameSuffix || '');
  const copy = DriveApp.getFileById(template.getId()).makeCopy(fileName, folder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  body.replaceText('\\{\\{client\\}\\}', clientName);
  body.replaceText('\\{\\{services\\}\\}', String(get('Service') || ''));
  body.replaceText('\\{\\{location\\}\\}', String(get('Event Location') || ''));
  body.replaceText('\\{\\{eventdate\\}\\}', eventDateFmt);
  body.replaceText('\\{\\{performancetime\\}\\}', startFmt);
  body.replaceText('\\{\\{price\\}\\}', fmtMoney_(price));
  body.replaceText('\\{\\{deposit\\}\\}', fmtMoney_(depositFinal));
  body.replaceText('\\{\\{balance\\}\\}', fmtMoney_(balanceFinal));
  body.replaceText('\\{\\{paymentmethod\\}\\}', paymentMethod || 'Not yet selected');
  body.replaceText('\\{\\{clientsignature\\}\\}', signedName ? (signedName + ' (signed electronically)') : '________________________');
  body.replaceText('\\{\\{signdate\\}\\}', signedName ? Utilities.formatDate(new Date(), tz_(), 'M/d/yyyy') : '________________________');
  body.replaceText('\\{\\{providersigndate\\}\\}', Utilities.formatDate(new Date(), tz_(), 'M/d/yyyy'));
  // Always today's date, whether or not the contract has been signed yet —
  // distinct from {{signdate}} on purpose, since that one is meant to stay
  // blank until the client actually signs.
  body.replaceText('\\{\\{agreementdate\\}\\}', Utilities.formatDate(new Date(), tz_(), 'M/d/yyyy'));
  body.replaceText('\\{\\{invoicenumber\\}\\}', invoiceNumber || '');
  body.replaceText('\\{\\{depositlabel\\}\\}', depositLabel);

  // cardfeenote lives inline sharing a line with the payment method, so it
  // always resolves via a plain text replacement — safe, no structural
  // element removal involved.
  //
  // travelfeeline and totalinvestment are different: they're each a whole
  // row's worth of content, and "only show it if there's actually a travel
  // fee" means those rows should disappear entirely rather than show an
  // empty dash — a "Total Investment" row is meaningless (and redundant
  // with Performance Fee) when there's no travel fee to add to it.
  // Table.removeRow() is the correct, safe API for removing a table row.
  // Removed bottom-up so earlier row indices don't shift out from under
  // the loop as later rows are deleted.
  if (travelFee > 0) {
    body.replaceText('\\{\\{travelfeeline\\}\\}', fmtMoney_(travelFee));
    body.replaceText('\\{\\{totalinvestment\\}\\}', fmtMoney_(totalDue));
  } else {
    var tables = body.getTables();
    for (var ti = 0; ti < tables.length; ti++) {
      var tbl = tables[ti];
      if (tbl.getNumRows() > 0 && tbl.getRow(0).getCell(0).getText().indexOf('Performance Fee') > -1) {
        for (var tr = tbl.getNumRows() - 1; tr >= 0; tr--) {
          var rowLabel = tbl.getRow(tr).getCell(0).getText();
          if (rowLabel.indexOf('Travel & Additional Fees') > -1 || rowLabel.indexOf('Total Investment') > -1) {
            tbl.removeRow(tr);
          }
        }
        break;
      }
    }
    // Fallback only \u2014 shouldn't normally trigger (the rows above are
    // already gone by this point), but guarantees no raw {{token}} is ever
    // left visible even if the table structure changes.
    body.replaceText('\\{\\{travelfeeline\\}\\}', '\u2014');
    body.replaceText('\\{\\{totalinvestment\\}\\}', '\u2014');
  }
  if (isCard && isBalanceCard) {
    body.replaceText('\\{\\{cardfeenote\\}\\}', ' \u2014 3.25% card processing fee included in both amounts above.');
  } else if (isCard) {
    body.replaceText('\\{\\{cardfeenote\\}\\}', ' \u2014 3.25% card processing fee included in the deposit above.');
  } else if (isBalanceCard) {
    body.replaceText('\\{\\{cardfeenote\\}\\}', ' \u2014 3.25% card processing fee included in the balance above.');
  } else {
    body.replaceText('\\{\\{cardfeenote\\}\\}', '');
  }

  doc.saveAndClose();

  const pdfBlob = DriveApp.getFileById(copy.getId()).getAs('application/pdf');
  const pdfFile = folder.createFile(pdfBlob).setName(fileName + '.pdf');
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return { copy: copy, doc: doc, pdfFile: pdfFile, deposit: depositFinal, balance: balanceFinal, price: price, paymentMethod: paymentMethod };
}
function generateContract(rowNum) {
  try {
    const sh = sheet_();
    const heads = headers_(sh);
    const row = sh.getRange(rowNum, 1, 1, heads.length).getValues()[0];
    const get = function (name) { var i = heads.indexOf(name); return i > -1 ? row[i] : ''; };

    const cols = ensureContractCols_(sh);
    const invoiceNumber = ensureInvoiceNumber_(sh, rowNum, cols);
    const oldDocId = String(sh.getRange(rowNum, cols.docIdCol).getValue() || '').trim();
    const result = buildMergedContract_(get, '', null, null, invoiceNumber);
    // Trash the previous unsigned copy now that a fresh one exists —
    // otherwise every regenerate leaves an orphaned old file sitting in
    // Drive, still reachable through its own old link even after the
    // sheet has moved on to a newer one. Signing already did this; a
    // plain regenerate should too, for the same reason.
    try { if (oldDocId) DriveApp.getFileById(oldDocId).setTrashed(true); } catch (e) {}

    // Reuse the same sign link across regenerations (e.g. after tweaking a
    // deposit amount) rather than minting a new token every time — otherwise
    // any link already copied or sent quietly breaks the moment you generate
    // again.
    const existingToken = String(sh.getRange(rowNum, cols.tokenCol).getValue() || '').trim();
    const token = existingToken || Utilities.getUuid();
    sh.getRange(rowNum, cols.docIdCol).setValue(result.copy.getId());
    sh.getRange(rowNum, cols.tokenCol).setValue(token);
    sh.getRange(rowNum, cols.signedCol).setValue('No');
    sh.getRange(rowNum, cols.pdfUrlCol).setValue(result.pdfFile.getUrl());
    sh.getRange(rowNum, cols.docUrlCol).setValue(result.copy.getUrl());

    const webAppUrl = ScriptApp.getService().getUrl();
    const signUrl = webAppUrl + '?sign=' + token;
    sh.getRange(rowNum, cols.signUrlCol).setValue(signUrl);

    return JSON.stringify({ ok: true, docUrl: result.copy.getUrl(), pdfUrl: result.pdfFile.getUrl(), signUrl: signUrl, invoiceNumber: invoiceNumber });
  } catch (e) {
    // Surface the real exception rather than a generic failure, so a
    // problem here is actually diagnosable instead of a dead end.
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
}
function getContractForSigning(token) {
  const sh = sheet_();
  const heads = headers_(sh);
  const tokenCol = heads.indexOf('Contract Sign Token') + 1;
  if (!tokenCol) return JSON.stringify({ ok: false });
  const last = sh.getLastRow();
  const values = sh.getRange(2, 1, last - 1, heads.length).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][tokenCol - 1]) === token) {
      const get = function (name) { var idx = heads.indexOf(name); return idx > -1 ? values[i][idx] : ''; };
      const amounts = computeBaseAmounts_(get);
      const price = amounts.price;
      // Base (pre-fee) amounts, sent as raw numbers rather than pre-formatted
      // strings — the signing page needs to recompute these live as the
      // client changes their payment method, since the card fee only
      // applies to whichever method they actually end up choosing.
      const deposit = amounts.deposit;
      const balance = amounts.balance;
      const ed = get('Date of Event');
      const depositPaid = String(get('Deposit Received') || '') === 'Yes';
      const balancePaid = String(get('Balance Paid') || '') === 'Yes';
      const balanceDueDaysRaw = Number(get('Balance Due Days'));
      const balanceDueDays = (balanceDueDaysRaw > 0) ? balanceDueDaysRaw : 0;
      var daysUntilEvent = null;
      var balanceDueDateFmt = '';
      if (ed) {
        var msPerDay = 24 * 60 * 60 * 1000;
        var today0 = new Date(); today0.setHours(0, 0, 0, 0);
        var eventDateObj = new Date(ed); eventDateObj.setHours(0, 0, 0, 0);
        daysUntilEvent = Math.round((eventDateObj - today0) / msPerDay);
        var balanceDueDateObj = new Date(eventDateObj.getTime() - balanceDueDays * msPerDay);
        balanceDueDateFmt = Utilities.formatDate(balanceDueDateObj, tz_(), 'MMMM d, yyyy');
      }
      return JSON.stringify({
        ok: true,
        client: get('Company or Organization') || get('Customer Name') || '',
        eventDate: ed ? Utilities.formatDate(new Date(ed), tz_(), 'M/d/yyyy') : '',
        services: get('Service') || '',
        location: get('Event Location') || '',
        depositBase: deposit,
        balanceBase: balance,
        paymentMethod: get('Payment Method') || '',
        pdfUrl: get('Contract PDF URL') || '',
        alreadySigned: String(get('Contract Signed') || '') === 'Yes',
        signedName: get('Contract Signed Name') || '',
        depositPaid: depositPaid,
        balancePaid: balancePaid,
        depositReceiptUrl: get('Deposit Receipt PDF URL') || '',
        balanceReceiptUrl: get('Balance Receipt PDF URL') || '',
        daysUntilEvent: daysUntilEvent,
        balanceDueDays: balanceDueDays,
        balanceDueDateFmt: balanceDueDateFmt,
        invoiceNumber: get('Invoice Number') || ''
      });
    }
  }
  return JSON.stringify({ ok: false });
}
function submitContractSignature(token, typedName, paymentMethod) {
  typedName = String(typedName || '').trim();
  if (!typedName) return JSON.stringify({ ok: false, error: 'Name required' });
  const sh = sheet_();
  const heads = headers_(sh);
  const tokenCol = heads.indexOf('Contract Sign Token') + 1;
  if (!tokenCol) return JSON.stringify({ ok: false, error: 'Not found' });
  const last = sh.getLastRow();
  const values = sh.getRange(2, tokenCol, last - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === token) {
      const rowNum = i + 2;
      const cols = ensureContractCols_(sh);

      // The client's own choice at the moment of signing is the ground
      // truth of how they'll actually pay — more authoritative than
      // whatever guess was set when the contract was first generated, so
      // it overwrites that guess rather than just living alongside it.
      const finalPaymentMethod = String(paymentMethod || '').trim();
      if (finalPaymentMethod) {
        var pmCol = heads.indexOf('Payment Method') + 1;
        if (!pmCol) { pmCol = heads.length + 1; sh.getRange(1, pmCol).setValue('Payment Method'); heads.push('Payment Method'); }
        sh.getRange(rowNum, pmCol).setValue(finalPaymentMethod);
      }

      sh.getRange(rowNum, cols.signedCol).setValue('Yes');
      sh.getRange(rowNum, cols.signedNameCol).setValue(typedName);
      sh.getRange(rowNum, cols.signedDateCol).setValue(new Date());

      try {
        // Re-fetch the row (not the headers — those haven't changed) now
        // that Payment Method/Signed fields are saved, so the merge picks
        // up the final values rather than whatever was there before this
        // save.
        const row2 = sh.getRange(rowNum, 1, 1, heads.length).getValues()[0];
        const get2 = function (name) { var idx = heads.indexOf(name); return idx > -1 ? row2[idx] : ''; };
        const oldDocId = sh.getRange(rowNum, cols.docIdCol).getValue();
        // Reuses the same invoice number assigned back at generateContract
        // time — this only ever creates a new one in the unlikely case a
        // contract somehow got signed without going through that step first.
        const invoiceNumber = ensureInvoiceNumber_(sh, rowNum, cols);
        const result = buildMergedContract_(get2, ' (Signed)', finalPaymentMethod || null, typedName, invoiceNumber);
        sh.getRange(rowNum, cols.docIdCol).setValue(result.copy.getId());
        sh.getRange(rowNum, cols.pdfUrlCol).setValue(result.pdfFile.getUrl());
        sh.getRange(rowNum, cols.docUrlCol).setValue(result.copy.getUrl());
        // Retire the earlier preview copy now that a final signed version exists.
        try { if (oldDocId) DriveApp.getFileById(String(oldDocId)).setTrashed(true); } catch (e) {}
      } catch (e) {}

      return JSON.stringify({ ok: true });
    }
  }
  return JSON.stringify({ ok: false, error: 'Not found' });
}
// Deliberately separate from the deposit's payment method — the deposit
// one gets locked in at signing time and never changes, but the balance
// can genuinely be paid a different way later. Saved the instant the
// client picks a method on the balance screen (not on a "confirm" click,
// since some payment paths — Venmo, bank transfer — don't have one), so
// whatever's stored always reflects their latest real choice, and both
// the balance receipt and this app's own display can calculate the
// correct fee for each half independently instead of assuming they match.
function saveBalancePaymentMethod(token, method) {
  try {
    const sh = sheet_();
    const heads = headers_(sh);
    const tokenCol = heads.indexOf('Contract Sign Token') + 1;
    if (!tokenCol) return JSON.stringify({ ok: false, error: 'Not found' });
    const last = sh.getLastRow();
    const values = sh.getRange(2, tokenCol, last - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]) === token) {
        const rowNum = i + 2;
        const cols = ensureContractCols_(sh);
        sh.getRange(rowNum, cols.balancePmCol).setValue(String(method || ''));
        return JSON.stringify({ ok: true });
      }
    }
    return JSON.stringify({ ok: false, error: 'Not found' });
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
}

/* ---------- Stripe: real card payment for the deposit ---------- */
// The secret key lives only in Script Properties, never in code or in any
// client-facing file — set it via the Apps Script editor's Project
// Settings -> Script Properties (key: STRIPE_SECRET_KEY). This function
// creates a Stripe-hosted Checkout Session for the deposit and returns its
// URL for the sign page to redirect to.
//
// Deliberately recomputes the charge amount here from the sheet itself,
// rather than trusting any amount passed in from the browser — the client
// only ever sends a token, never a dollar figure, so there's nothing for a
// tampered request to manipulate.
function createStripeCheckoutSession(token, kind) {
  try {
    kind = (kind === 'balance') ? 'balance' : 'deposit';
    const secretKey = String(PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY') || '').trim();
    if (!secretKey) return JSON.stringify({ ok: false, error: 'Card payment isn\u2019t set up yet \u2014 STRIPE_SECRET_KEY is missing in Script Properties.' });

    const sh = sheet_();
    const heads = headers_(sh);
    const tokenCol = heads.indexOf('Contract Sign Token') + 1;
    if (!tokenCol) return JSON.stringify({ ok: false, error: 'Not found' });
    const last = sh.getLastRow();
    const values = sh.getRange(2, 1, last - 1, heads.length).getValues();

    for (var i = 0; i < values.length; i++) {
      if (String(values[i][tokenCol - 1]) === token) {
        const get = function (name) { var idx = heads.indexOf(name); return idx > -1 ? values[i][idx] : ''; };
        const amounts = computeBaseAmounts_(get);
        const deposit = amounts.deposit;
        const balanceAmt = amounts.balance;
        const baseAmt = kind === 'balance' ? balanceAmt : deposit;
        const feeRate = 0.0325;
        const finalAmt = baseAmt * (1 + feeRate); // Card always includes the processing fee
        const cents = Math.round(finalAmt * 100);
        if (cents <= 0) return JSON.stringify({ ok: false, error: kind === 'balance' ? 'No balance is due on this contract.' : 'No deposit amount is set for this contract yet.' });

        const clientName = get('Company or Organization') || get('Customer Name') || 'Client';
        const eventDateVal = get('Date of Event');
        const eventDateFmt = eventDateVal ? Utilities.formatDate(new Date(eventDateVal), tz_(), 'M/d/yyyy') : '';

        const webAppUrl = ScriptApp.getService().getUrl();
        const successUrl = webAppUrl + '?sign=' + encodeURIComponent(token) + '&paid=1';
        const cancelUrl = webAppUrl + '?sign=' + encodeURIComponent(token);

        const payload = {
          'mode': 'payment',
          'success_url': successUrl,
          'cancel_url': cancelUrl,
          'line_items[0][quantity]': '1',
          'line_items[0][price_data][currency]': 'usd',
          'line_items[0][price_data][unit_amount]': String(cents),
          'line_items[0][price_data][product_data][name]': getConfig_().BUSINESS_NAME + ' ' + (kind === 'balance' ? 'Balance' : 'Deposit') + ' \u2014 ' + clientName + (eventDateFmt ? ' (' + eventDateFmt + ')' : ''),
          'metadata[contract_token]': token,
          'metadata[client]': clientName,
          'metadata[payment_kind]': kind
        };

        const options = {
          method: 'post',
          headers: { 'Authorization': 'Bearer ' + secretKey },
          payload: payload,
          muteHttpExceptions: true
        };
        const resp = UrlFetchApp.fetch('https://api.stripe.com/v1/checkout/sessions', options);
        const code = resp.getResponseCode();
        const bodyText = resp.getContentText();
        var json;
        try { json = JSON.parse(bodyText); } catch (parseErr) {
          return JSON.stringify({ ok: false, error: 'Stripe returned an unexpected response (HTTP ' + code + '): ' + bodyText.slice(0, 200) });
        }
        if (code >= 200 && code < 300 && json.url) {
          return JSON.stringify({ ok: true, url: json.url });
        }
        return JSON.stringify({ ok: false, error: (json.error && json.error.message) || ('Stripe error (HTTP ' + code + ') \u2014 please try again.') });
      }
    }
    return JSON.stringify({ ok: false, error: 'Not found' });
  } catch (e) {
    // Surface the real exception instead of leaving the client with a
    // generic failure — this is what actually shows up in the alert.
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
}



var LEAD_HEADERS = ['Customer Name', 'Phone number', 'Company or Organization', 'E-mail', 'Audience Size', 'Service', 'Event Type', 'Notes about interaction', 'Lead Source', 'Event Location', 'Quoted Price', 'Start Time', 'End Time', 'Status', 'Followup', 'Date of Event', 'Referred To', 'Timestamp', 'Lost Reason', 'Audience Age Range', 'Deposit Amount', 'Payment Method', 'Deposit Received', 'Deposit Received Date', 'Balance Paid', 'Balance Paid Date', 'Travel Fee', 'Contract PDF URL', 'Company Address', 'Balance Payment Method', 'Balance Alert Snoozed Until', 'Referred Date', 'Partner Check-In Dismissed', 'Contract Signed', 'Contract Doc URL', 'Contract Sign URL', 'Deposit Receipt PDF URL', 'Balance Receipt PDF URL', 'Balance Due Days'];

// A brand-new copy of this app starts with an empty leads sheet; lay down the
// header row once so adding/reading leads works. (Contract-specific columns are
// added on demand by ensureContractCols_.)
function ensureLeadHeaders_(sh) {
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, LEAD_HEADERS.length).setValues([LEAD_HEADERS]);
    try { sh.setFrozenRows(1); } catch (e) {}
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  ensureLeadHeaders_(sh);
  return sh;
}
// The spreadsheet's own timezone (File -> Settings -> Time zone), not the
// Apps Script project's separate timezone setting. Sheets always writes
// and interprets dates/times using ITS OWN timezone, regardless of what
// the script project is configured to — so anything read back out has to
// be formatted using that same spreadsheet timezone, or it can silently
// drift by however much the two settings happen to disagree.
function tz_() {
  return SpreadsheetApp.getActive().getSpreadsheetTimeZone();
}

function expensesSheet_() {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(EXPENSES_SHEET);
  if (!sh) {
    sh = ss.insertSheet(EXPENSES_SHEET);
    sh.appendRow(['Date', 'Vendor', 'Amount', 'Category', 'Notes', 'Source', 'Logged', 'Receipt URL', 'Receipt FileId', 'Ad Source']);
  } else {
    // Migration for sheets created before receipts/ad-source existed —
    // add any missing trailing columns without touching existing data.
    var header = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 10)).getValues()[0];
    if (header[7] !== 'Receipt URL' || header[8] !== 'Receipt FileId') {
      sh.getRange(1, 8, 1, 2).setValues([['Receipt URL', 'Receipt FileId']]);
    }
    if (header[9] !== 'Ad Source') {
      sh.getRange(1, 10, 1, 1).setValues([['Ad Source']]);
    }
  }
  return sh;
}
// Same self-healing "find by name or create" pattern as contractsFolder_()
// above — a separate folder from Contracts since receipts are a different
// kind of document with their own lifecycle (can be replaced/removed
// independently of anything else).
function receiptsFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('RECEIPTS_FOLDER_ID');
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (e) {}
  }
  const it = DriveApp.getFoldersByName(getConfig_().RECEIPTS_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(getConfig_().RECEIPTS_FOLDER);
  props.setProperty('RECEIPTS_FOLDER_ID', folder.getId());
  return folder;
}
// Receipts are organized into per-year subfolders inside the main
// receipts folder, based on the expense's own Date — not whatever day it
// happens to get uploaded. A December 2025 expense's receipt lands in a
// "2025" subfolder even if it's actually uploaded in January 2026, so
// everything stays grouped by the fiscal year it actually belongs to.
// Assumes a standard calendar-year fiscal year (Jan–Dec); flag it if
// yours runs on a different cycle and this can be adjusted.
function receiptYearFolder_(parentFolder, year) {
  var name = String(year);
  var it = parentFolder.getFoldersByName(name);
  return it.hasNext() ? it.next() : parentFolder.createFolder(name);
}
// Shared by updateExpense (when a date edit moves a receipt into a
// different year) and the one-time backfill below (for receipts uploaded
// before year-folders existed at all). Safe to call repeatedly — a no-op
// once the file is already in the right place.
function ensureFileInYearFolder_(file, targetFolder) {
  var alreadyThere = false;
  var parents = file.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === targetFolder.getId()) { alreadyThere = true; break; }
  }
  if (!alreadyThere) {
    var oldParents = file.getParents();
    var toRemove = [];
    while (oldParents.hasNext()) { toRemove.push(oldParents.next()); }
    toRemove.forEach(function(p){ p.removeFile(file); });
    targetFolder.addFile(file);
  }
}
// One-time backfill for receipts uploaded before per-year folders existed
// — they were sitting flat in the main receipts folder. Sorts every
// existing receipt into its matching year folder based on the expense's
// own Date, same rule new uploads already follow. Runs automatically on
// the next load, guarded so it only actually does the scan once; any
// individual receipt that fails (e.g. the underlying file was since
// deleted by hand) is skipped rather than blocking the rest.
function migrateReceiptsToYearFolders_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('RECEIPTS_YEAR_MIGRATED') === 'true') return;
  try {
    const sh = expensesSheet_();
    const values = sh.getDataRange().getValues();
    const parentFolder = receiptsFolder_();
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var fileId = row[8]; // Receipt FileId column
      var dateVal = row[0]; // Date column
      if (!fileId) continue;
      if (!(dateVal instanceof Date) || isNaN(dateVal.getTime())) continue;
      try {
        var targetFolder = receiptYearFolder_(parentFolder, dateVal.getFullYear());
        var file = DriveApp.getFileById(String(fileId));
        ensureFileInYearFolder_(file, targetFolder);
      } catch (e) { /* skip this one receipt, keep going */ }
    }
  } catch (e) { /* never let this block getLeads() */ }
  props.setProperty('RECEIPTS_YEAR_MIGRATED', 'true');
}

function getExpenses_() {
  const sh = expensesSheet_();
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const values = sh.getDataRange().getValues();
  const expenses = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[1] && !r[2]) continue; // skip fully blank rows
    expenses.push({
      _row: i + 1,
      date: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : String(r[0] || ''),
      vendor: String(r[1] || ''),
      amount: Number(r[2]) || 0,
      category: String(r[3] || ''),
      notes: String(r[4] || ''),
      source: String(r[5] || ''),
      receiptUrl: String(r[7] || ''),
      receiptFileId: String(r[8] || ''),
      adSource: String(r[9] || '')
    });
  }
  return expenses;
}

function getExpensesScoped_() { return JSON.stringify({ expenses: getExpenses_() }); }

// Normalizes a vendor string into a stable lookup key — strips digits
// (invoice/reference numbers that change every transaction, e.g. "GUSTO,
// TAX 514345" vs "GUSTO, TAX 514346") and punctuation, so the same real
// vendor keeps matching across repeat transactions.
function normalizeVendorKey_(vendor) {
  return String(vendor || '').toLowerCase().replace(/[0-9]+/g, '').replace(/[^a-z]+/g, ' ').trim().replace(/\s+/g, ' ');
}
function getVendorCategoryMemory_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('VENDOR_CATEGORY_MEMORY') || '{}'); }
  catch (e) { return {}; }
}
// Called whenever an expense is saved through the Add/Edit form — that's
// a deliberate, reviewed choice, unlike a CSV import row that might never
// get looked at closely. CSV import reads from this memory but never
// writes to it, to keep "what taught the system this" predictable.
function rememberVendorCategory_(vendor, category) {
  var key = normalizeVendorKey_(vendor);
  if (!key || !category) return;
  var props = PropertiesService.getScriptProperties();
  var mem = getVendorCategoryMemory_();
  mem[key] = category;
  props.setProperty('VENDOR_CATEGORY_MEMORY', JSON.stringify(mem));
}

function addExpense(date, vendor, amount, category, notes, adSource) {
  const sh = expensesSheet_();
  sh.appendRow([date ? parseYMD_(date) : new Date(), vendor || '', Number(amount) || 0, category || 'Other', notes || '', 'Manual', new Date(), '', '', adSource || '']);
  rememberVendorCategory_(vendor, category);
  // newRow lets the client immediately attach a receipt to a brand-new
  // expense right after creating it, without a second round-trip just to
  // find out which row it landed on.
  return JSON.stringify({ expenses: getExpenses_(), newRow: sh.getLastRow() });
}

/* ---------- Recurring expenses ----------
 * A rule in the "Recurring Expenses" tab spawns real rows in the Expenses sheet (Source =
 * "Recurring"), backfilled from the start date up to today and then extended each day. De-dup
 * is by the rule's "Last Generated" date, so re-running the generator never double-posts. */
var RECURRING_SHEET = 'Recurring Expenses';
var RECUR_HEADERS = ['ID', 'Active', 'Vendor', 'Amount', 'Category', 'Notes', 'Ad Source', 'Frequency', 'Start Date', 'End Date', 'Last Generated'];

function recurringSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RECURRING_SHEET);
  if (!sh) { sh = ss.insertSheet(RECURRING_SHEET); sh.appendRow(RECUR_HEADERS); try { sh.setFrozenRows(1); } catch (e) {} }
  return sh;
}
function daysInMonth_(y, m) { return new Date(y, m + 1, 0).getDate(); }
// The Nth occurrence (n = 0,1,2,…) counting from `start`, anchored to the start's day-of-month
// so monthly/quarterly/annual rules never drift — a rule on the 31st clamps to each month's last
// day. Weekly / Bi-weekly are simple day steps.
function addPeriods_(start, n, freq) {
  freq = String(freq);
  if (freq === 'Weekly' || freq === 'Bi-weekly') {
    var d = new Date(start.getFullYear(), start.getMonth(), start.getDate());
    d.setDate(d.getDate() + (freq === 'Weekly' ? 7 : 14) * n);
    return d;
  }
  var per = freq === 'Quarterly' ? 3 : freq === 'Annually' ? 12 : 1; // default Monthly
  var tm = start.getMonth() + per * n;
  var y = start.getFullYear() + Math.floor(tm / 12);
  var m = ((tm % 12) + 12) % 12;
  return new Date(y, m, Math.min(start.getDate(), daysInMonth_(y, m)));
}
function toDate0_(v) {
  if (!v) return null;
  var d = v instanceof Date ? new Date(v.getTime()) : new Date(v);
  if (isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0); return d;
}
// Post any occurrences that are now due for every active rule. Safe to run repeatedly.
function generateRecurringExpenses_() {
  var sh = recurringSheet_();
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var rows = sh.getRange(2, 1, last - 1, RECUR_HEADERS.length).getValues();
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var exp = expensesSheet_();
  var toAppend = [], now = new Date(), generated = 0;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!(String(r[1]).toLowerCase() === 'true' || r[1] === true)) continue; // Active only
    var start = toDate0_(r[8]); if (!start) continue;
    var freq = String(r[7]), end = toDate0_(r[9]), lastGen = toDate0_(r[10]);
    var horizon = (end && end < today) ? end : today;
    var newLast = lastGen;
    for (var n = 0; n < 5000; n++) {
      var occ = addPeriods_(start, n, freq);
      if (occ > horizon) break;
      if (lastGen && occ <= lastGen) continue; // already posted
      toAppend.push([occ, r[2] || '', Number(r[3]) || 0, r[4] || 'Other', r[5] || '', 'Recurring', now, '', '', r[6] || '']);
      newLast = occ; generated++;
    }
    if (newLast && (!lastGen || newLast > lastGen)) sh.getRange(i + 2, 11).setValue(newLast);
  }
  if (toAppend.length) exp.getRange(exp.getLastRow() + 1, 1, toAppend.length, 10).setValues(toAppend);
  return generated;
}
function addRecurringExpense(vendor, amount, category, notes, adSource, freq, startYmd, endYmd) {
  var sh = recurringSheet_();
  var id = 'RX-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
  var start = startYmd ? parseYMD_(startYmd) : new Date();
  var end = endYmd ? parseYMD_(endYmd) : '';
  sh.appendRow([id, true, vendor || '', Number(amount) || 0, category || 'Other', notes || '', adSource || '', freq || 'Monthly', start, end, '']);
  rememberVendorCategory_(vendor, category);
  generateRecurringExpenses_(); // backfill now
  return JSON.stringify({ expenses: getExpenses_(), recurring: getRecurring_() });
}
function getRecurring_() {
  var sh = recurringSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var rows = sh.getRange(2, 1, last - 1, RECUR_HEADERS.length).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!(String(r[1]).toLowerCase() === 'true' || r[1] === true)) continue;
    out.push({
      row: i + 2, id: r[0], vendor: r[2], amount: Number(r[3]) || 0, category: r[4], notes: r[5], adSource: r[6],
      frequency: r[7],
      start: r[8] instanceof Date ? fmtCell_(r[8], tz) : String(r[8] || ''),
      end: r[9] instanceof Date ? fmtCell_(r[9], tz) : String(r[9] || '')
    });
  }
  return out;
}
function stopRecurringExpense(row) {
  recurringSheet_().getRange(Number(row), 2).setValue(false); // Active = false
  return JSON.stringify({ expenses: getExpenses_(), recurring: getRecurring_() });
}

// Bulk insert for CSV import — the client already did the parsing,
// category-guessing, and duplicate review before this ever gets called,
// so this just trusts the reviewed batch and writes it in one shot
// rather than one appendRow per transaction. Imported rows never carry
// an ad-source attribution — that's a manual, per-transaction judgment
// call, not something a bank CSV can tell us.
function importExpensesBatch(rowsJson) {
  const sh = expensesSheet_();
  const rows = JSON.parse(rowsJson);
  const now = new Date();
  const out = rows.map(function(r){
    return [r.date ? parseYMD_(r.date) : now, r.vendor || '', Number(r.amount) || 0, r.category || 'Other', r.notes || '', 'CSV Import', now, '', '', r.adSource || ''];
  });
  if (out.length) {
    sh.getRange(sh.getLastRow()+1, 1, out.length, 10).setValues(out);
  }
  return getExpensesScoped_();
}

// Used by the "tag all suggested" bulk fix — writes only the Ad Source
// column, never touching date/vendor/amount/category/notes, so a batch
// backfill can't accidentally clobber anything else on rows it's meant to
// leave alone otherwise.
function bulkSetAdSource(updatesJson) {
  const sh = expensesSheet_();
  const updates = JSON.parse(updatesJson);
  updates.forEach(function(u) {
    if (u.row >= 2 && u.row <= sh.getLastRow() && u.adSource) {
      sh.getRange(u.row, 10).setValue(u.adSource);
    }
  });
  return getExpensesScoped_();
}

function updateExpense(rowNum, date, vendor, amount, category, notes, adSource) {
  const sh = expensesSheet_();
  if (rowNum < 2 || rowNum > sh.getLastRow()) return getExpensesScoped_();
  sh.getRange(rowNum, 1, 1, 5).setValues([[date ? parseYMD_(date) : '', vendor || '', Number(amount) || 0, category || 'Other', notes || '']]);
  sh.getRange(rowNum, 10).setValue(adSource || '');
  rememberVendorCategory_(vendor, category);
  // If a receipt is already attached and this date edit moved the expense
  // into a different year, relocate the actual Drive file to match —
  // otherwise correcting a date later would leave the receipt stranded in
  // whatever year folder it originally landed in.
  try {
    var fileId = sh.getRange(rowNum, 9).getValue();
    var newDateVal = sh.getRange(rowNum, 1).getValue();
    if (fileId && newDateVal instanceof Date && !isNaN(newDateVal.getTime())) {
      var targetFolder = receiptYearFolder_(receiptsFolder_(), newDateVal.getFullYear());
      var file = DriveApp.getFileById(String(fileId));
      ensureFileInYearFolder_(file, targetFolder);
    }
  } catch (e) { /* receipt relocation never blocks the main save */ }
  return getExpensesScoped_();
}

function deleteExpense(rowNum) {
  const sh = expensesSheet_();
  if (rowNum >= 2 && rowNum <= sh.getLastRow()) {
    // Trash any attached receipt too — otherwise deleting the expense
    // leaves an orphaned file sitting in Drive with nothing pointing to it.
    var fileId = sh.getRange(rowNum, 9).getValue();
    if (fileId) { try { DriveApp.getFileById(String(fileId)).setTrashed(true); } catch (e) {} }
    sh.deleteRow(rowNum);
  }
  return getExpensesScoped_();
}

// Receipts are stored as real Drive files (not inline in the sheet — a
// base64 image/PDF would bloat every read of the sheet even when nobody
// is looking at it). The row only holds a URL + file ID pointer.
function uploadExpenseReceipt(rowNum, base64Data, mimeType, fileName) {
  const sh = expensesSheet_();
  if (rowNum < 2 || rowNum > sh.getLastRow()) return getExpensesScoped_();
  // Replacing an existing receipt: trash the old file first so re-uploads
  // don't quietly pile up orphaned files in the folder over time.
  var existing = sh.getRange(rowNum, 9).getValue();
  if (existing) { try { DriveApp.getFileById(String(existing)).setTrashed(true); } catch (e) {} }
  var expenseDate = sh.getRange(rowNum, 1).getValue();
  var year = (expenseDate instanceof Date && !isNaN(expenseDate.getTime())) ? expenseDate.getFullYear() : new Date().getFullYear();
  var folder = receiptYearFolder_(receiptsFolder_(), year);
  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType || 'application/octet-stream', fileName || 'receipt');
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  sh.getRange(rowNum, 8, 1, 2).setValues([[file.getUrl(), file.getId()]]);
  return getExpensesScoped_();
}

function deleteExpenseReceipt(rowNum) {
  const sh = expensesSheet_();
  if (rowNum < 2 || rowNum > sh.getLastRow()) return getExpensesScoped_();
  var fileId = sh.getRange(rowNum, 9).getValue();
  if (fileId) { try { DriveApp.getFileById(String(fileId)).setTrashed(true); } catch (e) {} }
  sh.getRange(rowNum, 8, 1, 2).setValues([['', '']]);
  return getExpensesScoped_();
}

function tasksSheet_() {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(TASKS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(TASKS_SHEET);
    sh.appendRow(['Created', 'Task', 'Due', 'Done', 'Repeat']);
  }
  if (String(sh.getRange(1, 5).getValue() || '') === '') sh.getRange(1, 5).setValue('Repeat');
  return sh;
}

function getTasks_() {
  const sh = tasksSheet_();
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const values = sh.getDataRange().getValues();
  const tasks = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[1]) continue;
    tasks.push({
      _row: i + 1,
      id: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      task: String(r[1]),
      due: r[2] instanceof Date ? fmtCell_(r[2], tz) : (r[2] ? String(r[2]) : ''),
      done: String(r[3] || '') === 'Yes',
      repeat: String(r[4] || '')
    });
  }
  return tasks;
}

// Next occurrence after now, preserving time of day.
// Accepts Daily/Weekly/Monthly/Yearly or custom "Every N days/weeks/months/years".
function nextDue_(d, repeat) {
  const n = new Date(d.getTime());
  const now = new Date();
  var k = 1, unit = '';
  var m = String(repeat).match(/^Every\s+(\d+)\s+(day|week|month|year)s?$/i);
  if (m) { k = Number(m[1]) || 1; unit = m[2].toLowerCase(); }
  else if (repeat === 'Daily') unit = 'day';
  else if (repeat === 'Weekly') unit = 'week';
  else if (repeat === 'Monthly') unit = 'month';
  else if (repeat === 'Yearly') unit = 'year';
  else return null;
  var guard = 0;
  do {
    if (unit === 'day') n.setDate(n.getDate() + k);
    else if (unit === 'week') n.setDate(n.getDate() + 7 * k);
    else if (unit === 'month') n.setMonth(n.getMonth() + k);
    else n.setFullYear(n.getFullYear() + k);
    guard++;
  } while (n <= now && guard < 1000);
  return n;
}

function logSheet_() {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET);
    sh.appendRow(['When', 'Lead Key', 'Customer', 'Type', 'Note', 'Entry ID']);
  }
  // Backward-compat: older sheets don't have an Entry ID column yet. Add it
  // and backfill so every entry has a permanent ID independent of "When"
  // (which the person can now edit without breaking the entry's identity).
  if (String(sh.getRange(1, 6).getValue() || '') !== 'Entry ID') {
    sh.getRange(1, 6).setValue('Entry ID');
  }
  const last = sh.getLastRow();
  if (last >= 2) {
    const rng = sh.getRange(2, 6, last - 1, 1);
    const vals = rng.getValues();
    var changed = false;
    for (var i = 0; i < vals.length; i++) {
      if (!vals[i][0]) { vals[i][0] = Utilities.getUuid(); changed = true; }
    }
    if (changed) rng.setValues(vals);
  }
  return sh;
}

function getLog_() {
  const sh = logSheet_();
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const values = sh.getDataRange().getValues();
  const logs = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[3]) continue;
    var rawKey = r[1] instanceof Date ? String(r[1].getTime()) : String(r[1]);
    logs.push({
      row: Number(r[1]) || 0,   // legacy: old entries logged before the fix, by row number
      key: rawKey,              // current: lead's permanent numeric Timestamp key
      id: String(r[5] || ''),   // permanent Entry ID, independent of the editable When date
      when: r[0] instanceof Date ? fmtCell_(r[0], tz) : String(r[0]),
      type: String(r[3]),
      note: String(r[4] || '')
    });
  }
  return logs;
}

function findLogRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, 6, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === id) return i + 2;
  }
  return 0;
}

function deleteLog(id) {
  const sh = logSheet_();
  const rowNum = findLogRow_(sh, id);
  if (rowNum) sh.deleteRow(rowNum);
  return getLogsScoped_();
}

function updateLog(id, whenStr, note) {
  const sh = logSheet_();
  const rowNum = findLogRow_(sh, id);
  if (!rowNum) return getLogsScoped_();
  if (whenStr) sh.getRange(rowNum, 1).setValue(parseYMD_(whenStr));
  sh.getRange(rowNum, 5).setValue(note || '');
  return getLogsScoped_();
}

/**
 * ONE-TIME FIX: run this once from the editor after installing the update
 * above. Older history entries were tagged by sheet row number, which could
 * drift onto the wrong lead if a row was ever deleted or the sheet resorted.
 * This rewrites each old entry to the Timestamp of whichever lead currently
 * sits at that row number. It's a best-effort repair: if a row already
 * pointed at the wrong lead before this fix, this cannot know that and will
 * carry the same row-based guess forward one last time. Anything that still
 * looks wrong afterward can be removed with the \u2715 in the app's contact
 * history. Safe to run more than once.
 */
function migrateContactLogKeys() {
  const ls = logSheet_();
  const last = ls.getLastRow();
  if (last < 2) return 'Nothing to migrate.';
  const sh = sheet_();
  const heads = headers_(sh);
  const tsCol = heads.indexOf('Timestamp');
  const leadRows = sh.getDataRange().getValues();
  const rowToKey = {};
  for (var i = 1; i < leadRows.length; i++) {
    const ts = leadRows[i][tsCol];
    if (ts) rowToKey[i + 1] = ts instanceof Date ? String(ts.getTime()) : String(ts);
  }
  const rng = ls.getRange(2, 2, last - 1, 1);
  const vals = rng.getValues();
  var fixed = 0, alreadyOk = 0, orphaned = 0;
  for (var j = 0; j < vals.length; j++) {
    const cur = String(vals[j][0]);
    if (/^\d{10,}$/.test(cur)) { alreadyOk++; continue; } // already a numeric Timestamp key
    const rowNum = Number(cur);
    if (rowNum && rowToKey[rowNum]) { vals[j][0] = rowToKey[rowNum]; fixed++; }
    else orphaned++;
  }
  rng.setValues(vals);
  return 'Migrated ' + fixed + ', already fine ' + alreadyOk + ', could not match ' + orphaned + '.';
}

function logContact(leadKey, customer, type, note) {
  logSheet_().appendRow([new Date(), leadKey, customer, type, note || '', Utilities.getUuid()]);
  return getLogsScoped_();
}

/* ---------- Email templates ---------- */
function templatesSheet_() {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Email Templates');
  if (!sh) {
    sh = ss.insertSheet('Email Templates');
    sh.appendRow(['Created', 'Name', 'Subject', 'Body']);
    sh.appendRow([new Date(), 'Following up', 'Following up on your magic show inquiry',
      'Hi {firstname},\n\nJust wanted to follow up on your inquiry about booking magic for your event. Do you have a few minutes to chat about details?\n\nBest,\n' + getConfig_().OWNER_NAME + '\n' + getConfig_().BUSINESS_NAME]);
  }
  return sh;
}
function getTemplates_() {
  const sh = templatesSheet_();
  const values = sh.getDataRange().getValues();
  const t = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[1]) continue;
    var sortOrder = (r[4] === '' || r[4] === null || r[4] === undefined) ? i : Number(r[4]);
    if (isNaN(sortOrder)) sortOrder = i; // never explicitly reordered — keep natural (creation) order
    t.push({
      id: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      name: String(r[1]), subject: String(r[2] || ''), body: String(r[3] || ''),
      _so: sortOrder
    });
  }
  t.sort(function (a, b) { return a._so - b._so; });
  return t.map(function (x) { return { id: x.id, name: x.name, subject: x.subject, body: x.body }; });
}
function moveTemplate(id, direction) {
  const sh = templatesSheet_();
  const values = sh.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[1]) continue;
    var so = (r[4] === '' || r[4] === null || r[4] === undefined) ? i : Number(r[4]);
    if (isNaN(so)) so = i;
    list.push({
      id: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      name: String(r[1]), subject: String(r[2] || ''), body: String(r[3] || ''),
      so: so, rowNum: i + 1
    });
  }
  list.sort(function (a, b) { return a.so - b.so; });
  var idx = -1;
  for (var j = 0; j < list.length; j++) { if (list[j].id === id) { idx = j; break; } }
  var swapWith = direction === 'up' ? idx - 1 : idx + 1;
  if (idx !== -1 && swapWith >= 0 && swapWith < list.length) {
    // Swap the actual stored order values (not just their ranking) so this
    // stays correct even for rows that never had an explicit value before.
    var a = list[idx], b = list[swapWith];
    sh.getRange(a.rowNum, 5).setValue(b.so);
    sh.getRange(b.rowNum, 5).setValue(a.so);
    var tmp = a.so; a.so = b.so; b.so = tmp;
    list.sort(function (x, y) { return x.so - y.so; });
  }
  // Return the already-fresh-in-memory list directly, rather than calling
  // getTemplatesScoped_() and re-reading the whole sheet a second time —
  // that redundant read was most of what made this feel slow.
  return JSON.stringify({ templates: list.map(function (x) { return { id: x.id, name: x.name, subject: x.subject, body: x.body }; }) });
}
function findTemplateRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i][0];
    var iso = v instanceof Date ? v.toISOString() : String(v);
    if (iso === id) return i + 2;
  }
  return 0;
}
function addTemplate(name, subject, body, id) {
  const sh = templatesSheet_();
  var created = id ? new Date(id) : new Date();
  if (isNaN(created.getTime())) created = new Date();
  sh.appendRow([created, name, subject, body]);
  return getTemplatesScoped_();
}
function updateTemplate(id, name, subject, body) {
  const sh = templatesSheet_();
  const rowNum = findTemplateRow_(sh, id);
  if (!rowNum) return getTemplatesScoped_();
  sh.getRange(rowNum, 2, 1, 3).setValues([[name, subject, body]]);
  return getTemplatesScoped_();
}
function deleteTemplate(id) {
  const sh = templatesSheet_();
  const rowNum = findTemplateRow_(sh, id);
  if (rowNum) sh.deleteRow(rowNum);
  return getTemplatesScoped_();
}

/* ---------- Partner email templates (separate from lead templates) ---------- */
function partnerTemplatesSheet_() {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Partner Email Templates');
  if (!sh) {
    sh = ss.insertSheet('Partner Email Templates');
    sh.appendRow(['Created', 'Name', 'Subject', 'Body']);
  }
  return sh;
}
function getPartnerTemplates_() {
  const sh = partnerTemplatesSheet_();
  const values = sh.getDataRange().getValues();
  const t = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[1]) continue;
    t.push({
      id: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      name: String(r[1]), subject: String(r[2] || ''), body: String(r[3] || '')
    });
  }
  return t;
}
function findPartnerTemplateRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i][0];
    var iso = v instanceof Date ? v.toISOString() : String(v);
    if (iso === id) return i + 2;
  }
  return 0;
}
function addPartnerTemplate(name, subject, body, id) {
  const sh = partnerTemplatesSheet_();
  var created = id ? new Date(id) : new Date();
  if (isNaN(created.getTime())) created = new Date();
  sh.appendRow([created, name, subject, body]);
  return getPartnerTemplatesScoped_();
}
function updatePartnerTemplate(id, name, subject, body) {
  const sh = partnerTemplatesSheet_();
  const rowNum = findPartnerTemplateRow_(sh, id);
  if (!rowNum) return getPartnerTemplatesScoped_();
  sh.getRange(rowNum, 2, 1, 3).setValues([[name, subject, body]]);
  return getPartnerTemplatesScoped_();
}
function deletePartnerTemplate(id) {
  const sh = partnerTemplatesSheet_();
  const rowNum = findPartnerTemplateRow_(sh, id);
  if (rowNum) sh.deleteRow(rowNum);
  return getPartnerTemplatesScoped_();
}

/* ---------- Text (SMS) templates ---------- */
function textTplSheet_() {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Text Templates');
  if (!sh) {
    sh = ss.insertSheet('Text Templates');
    sh.appendRow(['Created', 'Name', 'Body']);
    sh.appendRow([new Date(), 'Quick follow-up', 'Hi {firstname}, just following up on your magic show inquiry \u2014 do you have a few minutes to chat details? \u2014 ' + getConfig_().OWNER_NAME + ', ' + getConfig_().BUSINESS_NAME]);
  }
  return sh;
}
function getTextTemplates_() {
  const sh = textTplSheet_();
  const values = sh.getDataRange().getValues();
  const t = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[1]) continue;
    var sortOrder = (r[3] === '' || r[3] === null || r[3] === undefined) ? i : Number(r[3]);
    if (isNaN(sortOrder)) sortOrder = i;
    t.push({
      id: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      name: String(r[1]), body: String(r[2] || ''),
      _so: sortOrder
    });
  }
  t.sort(function (a, b) { return a._so - b._so; });
  return t.map(function (x) { return { id: x.id, name: x.name, body: x.body }; });
}
function moveTextTemplate(id, direction) {
  const sh = textTplSheet_();
  const values = sh.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[1]) continue;
    var so = (r[3] === '' || r[3] === null || r[3] === undefined) ? i : Number(r[3]);
    if (isNaN(so)) so = i;
    list.push({
      id: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      name: String(r[1]), body: String(r[2] || ''),
      so: so, rowNum: i + 1
    });
  }
  list.sort(function (a, b) { return a.so - b.so; });
  var idx = -1;
  for (var j = 0; j < list.length; j++) { if (list[j].id === id) { idx = j; break; } }
  var swapWith = direction === 'up' ? idx - 1 : idx + 1;
  if (idx !== -1 && swapWith >= 0 && swapWith < list.length) {
    var a = list[idx], b = list[swapWith];
    sh.getRange(a.rowNum, 4).setValue(b.so);
    sh.getRange(b.rowNum, 4).setValue(a.so);
    var tmp = a.so; a.so = b.so; b.so = tmp;
    list.sort(function (x, y) { return x.so - y.so; });
  }
  return JSON.stringify({ textTemplates: list.map(function (x) { return { id: x.id, name: x.name, body: x.body }; }) });
}
function findTextTplRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i][0];
    var iso = v instanceof Date ? v.toISOString() : String(v);
    if (iso === id) return i + 2;
  }
  return 0;
}
function addTextTemplate(name, body, id) {
  const sh = textTplSheet_();
  var created = id ? new Date(id) : new Date();
  if (isNaN(created.getTime())) created = new Date();
  sh.appendRow([created, name, body]);
  return getTextTemplatesScoped_();
}
function updateTextTemplate(id, name, body) {
  const sh = textTplSheet_();
  const rowNum = findTextTplRow_(sh, id);
  if (!rowNum) return getTextTemplatesScoped_();
  sh.getRange(rowNum, 2, 1, 2).setValues([[name, body]]);
  return getTextTemplatesScoped_();
}
function deleteTextTemplate(id) {
  const sh = textTplSheet_();
  const rowNum = findTextTplRow_(sh, id);
  if (rowNum) sh.deleteRow(rowNum);
  return getTextTemplatesScoped_();
}

/* ---------- Partner text templates (separate from lead text templates) ---------- */
function partnerTextTplSheet_() {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Partner Text Templates');
  if (!sh) {
    sh = ss.insertSheet('Partner Text Templates');
    sh.appendRow(['Created', 'Name', 'Body']);
  }
  return sh;
}
function getPartnerTextTemplates_() {
  const sh = partnerTextTplSheet_();
  const values = sh.getDataRange().getValues();
  const t = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[1]) continue;
    t.push({
      id: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      name: String(r[1]), body: String(r[2] || '')
    });
  }
  return t;
}
function findPartnerTextTplRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i][0];
    var iso = v instanceof Date ? v.toISOString() : String(v);
    if (iso === id) return i + 2;
  }
  return 0;
}
function addPartnerTextTemplate(name, body, id) {
  const sh = partnerTextTplSheet_();
  var created = id ? new Date(id) : new Date();
  if (isNaN(created.getTime())) created = new Date();
  sh.appendRow([created, name, body]);
  return getPartnerTextTemplatesScoped_();
}
function updatePartnerTextTemplate(id, name, body) {
  const sh = partnerTextTplSheet_();
  const rowNum = findPartnerTextTplRow_(sh, id);
  if (!rowNum) return getPartnerTextTemplatesScoped_();
  sh.getRange(rowNum, 2, 1, 2).setValues([[name, body]]);
  return getPartnerTextTemplatesScoped_();
}
function deletePartnerTextTemplate(id) {
  const sh = partnerTextTplSheet_();
  const rowNum = findPartnerTextTplRow_(sh, id);
  if (rowNum) sh.deleteRow(rowNum);
  return getPartnerTextTemplatesScoped_();
}

/* ---------- Recycle bin (soft delete) ---------- */
// Deleting a lead never removes it outright — the row is moved into a
// "Leads Trash" sheet where it sits for TRASH_RETENTION_DAYS, restorable
// at any point, then permanently purged by a daily trigger. Every path
// that used to hard-delete a lead now goes through here first; there is
// deliberately no "skip the bin" option.

function trashSheet_() {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(TRASH_SHEET);
  if (!sh) {
    sh = ss.insertSheet(TRASH_SHEET);
    sh.appendRow(['TrashId', 'DeletedAt', 'HeadersJSON', 'ValuesJSON']);
  }
  return sh;
}

// Dates have to survive a JSON round-trip without losing their type, or a
// restored lead would come back with "Followup" as a plain string instead
// of a real Date the rest of the app can filter/sort on. Every Date in the
// row gets tagged with its epoch ms so deserializeRowFromTrash_ can turn it
// back into a real Date on the way out, no matter which column it was in.
function serializeRowForTrash_(values) {
  return JSON.stringify(values.map(function (v) {
    return (v instanceof Date) ? { __d: v.getTime() } : v;
  }));
}
function deserializeRowFromTrash_(json) {
  return JSON.parse(json).map(function (v) {
    return (v && typeof v === 'object' && '__d' in v) ? new Date(v.__d) : v;
  });
}

function findTrashRow_(sh, trashId) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === trashId) return i + 2;
  }
  return 0;
}

// Booked leads carry a live event on the primary calendar (see
// syncBookedCalendarEvent_). Trashing a lead removes that event so it
// doesn't sit on the calendar for a booking that's no longer active, and
// blanks the ID in the captured row so a later restore creates a fresh
// event rather than trying to reuse a deleted one.
function removeBookingEventForTrash_(sh, heads, rowValues) {
  const statusIdx = heads.indexOf('Status');
  const idIdx = heads.indexOf('Booking Calendar Event ID');
  if (statusIdx < 0 || idIdx < 0) return;
  if (String(rowValues[statusIdx] || '') !== 'Booked') return;
  const eventId = rowValues[idIdx];
  if (!eventId) return;
  try {
    var ev = CalendarApp.getDefaultCalendar().getEventById(eventId);
    if (ev) ev.deleteEvent();
  } catch (e) {}
  rowValues[idIdx] = '';
}

function deleteLead(rowNum) {
  const sh = sheet_();
  const heads = headers_(sh);
  const rowValues = sh.getRange(rowNum, 1, 1, heads.length).getValues()[0];
  removeBookingEventForTrash_(sh, heads, rowValues);

  const trash = trashSheet_();
  trash.appendRow([
    Utilities.getUuid(),
    new Date(),
    JSON.stringify(heads),
    serializeRowForTrash_(rowValues)
  ]);

  sh.deleteRow(rowNum);
  // Contact history is keyed by each lead's permanent Timestamp, not row
  // position, so deleting a row no longer requires realigning the log.
  try { runFrequentAutomations(); } catch (e) {}
  try { scheduleCalendarRebuild_(); } catch (e) {}
  try { ensureTrashPurgeTrigger_(); } catch (e) {}
  return getLeadsScoped_();
}

// Read-ready trash list for the client — same shape as buildLeadsArray_
// (header -> formatted value) plus trashId/deletedAt/daysRemaining, so the
// Recycle Bin UI can render a lead card without any client-side parsing.
function getTrash() {
  const sh = trashSheet_();
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, 4).getValues();
  const out = [];
  rows.forEach(function (r) {
    var trashId = String(r[0] || '');
    if (!trashId) return;
    var deletedAt = r[1] instanceof Date ? r[1] : new Date(r[1]);
    var heads;
    try { heads = JSON.parse(r[2]); } catch (e) { heads = []; }
    var values;
    try { values = deserializeRowFromTrash_(r[3]); } catch (e) { values = []; }
    const o = { _trashId: trashId, _deletedAt: fmtCell_(deletedAt, tz) };
    var daysLeft = TRASH_RETENTION_DAYS - Math.floor((Date.now() - deletedAt.getTime()) / 86400000);
    o._daysRemaining = Math.max(0, daysLeft);
    heads.forEach(function (h, i) {
      var v = values[i];
      o[h] = (v instanceof Date) ? fmtCell_(v, tz) : v;
    });
    out.push(o);
  });
  // Most recently deleted first.
  out.sort(function (a, b) { return a._deletedAt < b._deletedAt ? 1 : -1; });
  return out;
}
function getTrashScoped_() { return JSON.stringify({ trash: getTrash() }); }

function restoreLead(trashId) {
  const trash = trashSheet_();
  const trashRow = findTrashRow_(trash, trashId);
  if (!trashRow) return getLeadsScoped_();

  const cells = trash.getRange(trashRow, 1, 1, 4).getValues()[0];
  var storedHeads;
  try { storedHeads = JSON.parse(cells[2]); } catch (e) { storedHeads = []; }
  var storedValues;
  try { storedValues = deserializeRowFromTrash_(cells[3]); } catch (e) { storedValues = []; }

  const sh = sheet_();
  const heads = headers_(sh);
  const row = heads.map(function (h) {
    var idx = storedHeads.indexOf(h);
    return idx > -1 ? storedValues[idx] : '';
  });
  sh.appendRow(row);
  const newRow = sh.getLastRow();

  trash.deleteRow(trashRow);

  try { runFrequentAutomations(); } catch (e) {}
  try { scheduleCalendarRebuild_(); } catch (e) {}
  // existingId is blank (cleared when trashed) so this always creates a
  // fresh calendar event rather than reusing the deleted one.
  try { syncBookedCalendarEvent_(sh, newRow, false); } catch (e) {}

  return getLeadsScoped_();
}

function permanentlyDeleteTrash(trashId) {
  const trash = trashSheet_();
  const trashRow = findTrashRow_(trash, trashId);
  if (trashRow) trash.deleteRow(trashRow);
  return getTrashScoped_();
}

const TRASH_PURGE_TRIGGER_FN = 'purgeOldTrashJob';

// Lazily installed the first time anything is trashed — same defensive
// "remove any existing one, then add ours" pattern as scheduleCalendarRebuild_,
// just daily instead of one-shot, so re-running this never stacks up
// duplicate triggers.
function ensureTrashPurgeTrigger_() {
  const already = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === TRASH_PURGE_TRIGGER_FN;
  });
  if (already) return;
  ScriptApp.newTrigger(TRASH_PURGE_TRIGGER_FN).timeBased().everyDays(1).atHour(3).create();
}

function purgeOldTrashJob() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    const sh = trashSheet_();
    const last = sh.getLastRow();
    if (last < 2) return;
    const cutoff = Date.now() - TRASH_RETENTION_DAYS * 86400000;
    const deletedAt = sh.getRange(2, 2, last - 1, 1).getValues();
    // Delete from the bottom up so removing a row never shifts the index
    // of a row still waiting to be checked.
    for (var i = deletedAt.length - 1; i >= 0; i--) {
      var d = deletedAt[i][0] instanceof Date ? deletedAt[i][0] : new Date(deletedAt[i][0]);
      if (d.getTime() <= cutoff) sh.deleteRow(i + 2);
    }
  } finally {
    lock.releaseLock();
  }
}

function headers_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim(); });
}

/* ---------- Google Contacts sync ---------- */
// Requires the "People API" advanced service to be enabled for this project
// (Apps Script editor -> Services (+) -> People API -> Add). One-way sync:
// this app writes into Google Contacts; edits made directly in Contacts do
// not flow back. Re-running the sync updates existing contacts in place.

const CONTACT_COL = 'Google Contact ID';
function contactGroupResourceName_() {
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('CONTACT_GROUP_ID');
  if (saved) {
    try {
      if (People.ContactGroups.get(saved)) return saved;
    } catch (e) {}
  }
  const groups = People.ContactGroups.list({ pageSize: 200 }).contactGroups || [];
  const found = groups.find(function (g) { return g.name === getConfig_().CONTACT_GROUP; });
  if (found) { props.setProperty('CONTACT_GROUP_ID', found.resourceName); return found.resourceName; }
  const created = People.ContactGroups.create({ contactGroup: { name: getConfig_().CONTACT_GROUP } });
  props.setProperty('CONTACT_GROUP_ID', created.resourceName);
  return created.resourceName;
}

function ensureContactColumn_(sh) {
  const heads = headers_(sh);
  var col = heads.indexOf(CONTACT_COL) + 1;
  if (!col) {
    col = heads.length + 1;
    sh.getRange(1, col).setValue(CONTACT_COL);
  }
  return col;
}

/**
 * Run this from the editor (or tap "Sync to Contacts" if wired into the app)
 * to push every lead with a name into Google Contacts under this account.
 */
const SYNC_STATUS_KEY = 'contactSyncStatus';
const SYNC_TRIGGER_FN = 'runContactSyncJob';

/**
 * Kicks off a contact sync in the background via a one-time trigger, rather
 * than running it inline and making the person wait. Guarded with a lock so
 * two taps (or two tabs) can never start overlapping syncs — that race is
 * exactly how the sync could create duplicate contacts, since two runs could
 * each see the same not-yet-synced lead and both try to create it.
 */
function startContactSync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return { alreadyRunning: true };
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty(SYNC_STATUS_KEY);
    if (raw) {
      const cur = JSON.parse(raw);
      if (cur.status === 'running') return { alreadyRunning: true };
    }
    props.setProperty(SYNC_STATUS_KEY, JSON.stringify({
      status: 'running', startedAt: Date.now(), created: 0, updated: 0, skipped: 0
    }));
    cleanUpSyncTriggers_();
    ScriptApp.newTrigger(SYNC_TRIGGER_FN).timeBased().after(1000).create();
    return { started: true };
  } finally {
    lock.releaseLock();
  }
}

function getContactSyncStatus() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(SYNC_STATUS_KEY);
  return raw ? JSON.parse(raw) : { status: 'idle' };
}

function cleanUpSyncTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === SYNC_TRIGGER_FN) ScriptApp.deleteTrigger(t);
  });
}

/**
 * The actual background worker, fired by a trigger rather than called
 * directly — runs entirely independently of whether anyone has the app
 * open. If a single pass runs out of time (see the time budget inside
 * syncContactsToGoogle), it automatically schedules the next pass itself
 * and keeps going until the whole list is done, with zero further taps
 * needed from the person who started it.
 */
function runContactSyncJob() {
  cleanUpSyncTriggers_();
  const props = PropertiesService.getScriptProperties();
  const prevRaw = props.getProperty(SYNC_STATUS_KEY);
  const prev = prevRaw ? JSON.parse(prevRaw) : { createdSoFar: 0, updatedSoFar: 0, startedAt: Date.now() };

  const result = syncContactsToGoogle();

  const totals = {
    status: result.timedOut ? 'running' : 'done',
    startedAt: prev.startedAt || Date.now(),
    createdSoFar: (prev.createdSoFar || 0) + result.created,
    updatedSoFar: (prev.updatedSoFar || 0) + result.updated,
    created: (prev.createdSoFar || 0) + result.created,
    updated: (prev.updatedSoFar || 0) + result.updated,
    skipped: result.skipped,
    errorSamples: result.errorSamples,
    finishedAt: Date.now()
  };
  props.setProperty(SYNC_STATUS_KEY, JSON.stringify(totals));

  if (result.timedOut) {
    ScriptApp.newTrigger(SYNC_TRIGGER_FN).timeBased().after(1000).create();
  }
}

/**
 * Looks up every contact already in the Leads contact group and maps
 * name -> resourceName. This is the safety net against duplicates: even if
 * a row's "Google Contact ID" tracking is ever lost or was never written
 * back for some reason, the sync can still recognize "a contact with this
 * exact name already exists" before deciding to create a new one.
 */
function buildExistingContactNameMap_(groupRN) {
  const map = {};
  try {
    const group = People.ContactGroups.get(groupRN, { maxMembers: 2000 });
    const resourceNames = group.memberResourceNames || [];
    for (var i = 0; i < resourceNames.length; i += 200) {
      const chunk = resourceNames.slice(i, i + 200);
      const batch = People.People.getBatchGet({ resourceNames: chunk, personFields: 'names' });
      (batch.responses || []).forEach(function (r) {
        if (r.person && r.person.names && r.person.names.length && r.person.resourceName) {
          const nm = (r.person.names[0].unstructuredName || r.person.names[0].displayName || '').trim().toLowerCase();
          if (nm && !map[nm]) map[nm] = r.person.resourceName;
        }
      });
    }
  } catch (e) { /* if this lookup fails for any reason, sync still works — just without this extra safety net */ }
  return map;
}

function syncContactsToGoogle() {
  const sh = sheet_();
  const heads = headers_(sh);
  const idCol = ensureContactColumn_(sh);
  const groupRN = contactGroupResourceName_();
  const existingByName = buildExistingContactNameMap_(groupRN);
  const values = sh.getDataRange().getValues();
  const col = function (name) { return heads.indexOf(name); };

  var created = 0, updated = 0, skipped = 0, timedOut = false;
  const errorSamples = []; // a few real error messages, so failures are diagnosable
  const startedAt = Date.now();
  const TIME_BUDGET_MS = 4.5 * 60 * 1000; // stop with time to spare before Apps Script's ~6 min limit

  for (var i = 1; i < values.length; i++) {
    if (Date.now() - startedAt > TIME_BUDGET_MS) { timedOut = true; break; }

    const row = values[i];
    const name = String(row[col('Customer Name')] || '').trim();
    if (!name) { skipped++; continue; }
    const phone = String(row[col('Phone number')] || '').trim();
    const email = String(row[col('E-mail')] || '').trim();
    const status = String(row[col('Status')] || '').trim();
    const eventDate = row[col('Date of Event')];
    const notes = String(row[col('Notes about interaction')] || '').trim();
    const quoted = row[col('Quoted Price')];
    const existingId = String(row[idCol - 1] || '').trim();

    const bioLines = [getConfig_().BUSINESS_NAME + ' lead \u2014 status: ' + (status || 'New')];
    if (eventDate) bioLines.push('Event date: ' + eventDate);
    if (quoted) bioLines.push('Quoted: $' + quoted);
    if (notes) bioLines.push('Notes: ' + notes);

    const person = {
      names: [{ unstructuredName: name }],
      memberships: [{ contactGroupMembership: { contactGroupResourceName: groupRN } }],
      biographies: [{ value: bioLines.join('\n'), contentType: 'TEXT_PLAIN' }],
    };
    if (phone) person.phoneNumbers = [{ value: phone, type: 'mobile' }];
    if (email) person.emailAddresses = [{ value: email, type: 'home' }];

    try {
      if (existingId) {
        const current = People.People.get(existingId, { personFields: 'names' });
        person.etag = current.etag; // required for updates, but goes in the resource body, not as a separate argument
        People.People.updateContact(person, existingId, {
          updatePersonFields: 'names,phoneNumbers,emailAddresses,biographies,memberships',
        });
        updated++;
      } else {
        const matchedId = existingByName[name.trim().toLowerCase()];
        if (matchedId) {
          // A contact with this exact name already exists in the group —
          // this row's tracking was lost somewhere along the way. Reattach
          // to the real contact and update it, instead of creating a
          // second copy of someone who's already there.
          const current2 = People.People.get(matchedId, { personFields: 'names' });
          person.etag = current2.etag;
          People.People.updateContact(person, matchedId, {
            updatePersonFields: 'names,phoneNumbers,emailAddresses,biographies,memberships',
          });
          sh.getRange(i + 1, idCol).setValue(matchedId);
          SpreadsheetApp.flush();
          delete existingByName[name.trim().toLowerCase()]; // claimed — a second same-named row should create its own, not merge in
          updated++;
        } else {
          const res = People.People.createContact(person);
          sh.getRange(i + 1, idCol).setValue(res.resourceName);
          SpreadsheetApp.flush(); // force the ID to actually land before moving on
          created++;
        }
      }
    } catch (e) {
      // Only recreate if the contact genuinely no longer exists (deleted
      // directly in Google Contacts). Anything else — rate limits,
      // permission hiccups, temporary API errors — should NOT trigger a
      // recreate, since that's exactly how duplicates happen. Those rows
      // just get skipped this round and retried on the next sync.
      const msg = String((e && e.message) || e);
      const genuinelyDeleted = /not found/i.test(msg) || msg.indexOf('404') > -1;
      if (genuinelyDeleted) {
        try {
          const res = People.People.createContact(person);
          sh.getRange(i + 1, idCol).setValue(res.resourceName);
          SpreadsheetApp.flush();
          created++;
        } catch (e2) {
          skipped++;
          if (errorSamples.length < 8) errorSamples.push(name + ': ' + String((e2 && e2.message) || e2));
        }
      } else {
        skipped++;
        if (errorSamples.length < 8) errorSamples.push(name + ': ' + msg);
      }
    }
  }
  return { created: created, updated: updated, skipped: skipped, timedOut: timedOut, errorSamples: errorSamples };
}

function parseYMD_(s) {
  // "2026-07-20" or "2026-07-20 14:30" -> local Date (avoids timezone drift)
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?$/);
  if (!m) return s;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]),
                  m[4] ? Number(m[4]) : 0, m[5] ? Number(m[5]) : 0);
}

/* ---------- API called from the app ---------- */

/**
 * Sheets stores a time-only cell (like 8:30 AM) against the year 1899, when
 * timezones ran on odd local-mean-time offsets. Converting that to UTC and back
 * shifts the clock by a few minutes. So we format every date/time cell here,
 * in the spreadsheet's own timezone, and hand the app plain text instead.
 */
function fmtCell_(v, tz) {
  if (v.getFullYear() < 1900) return Utilities.formatDate(v, tz, 'HH:mm');
  const s = Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm');
  return s.slice(11) === '00:00' ? s.slice(0, 10) : s;
}

function buildLeadsArray_() {
  const sh = sheet_();
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const values = sh.getDataRange().getValues();
  const heads = values[0].map(function (h) { return String(h).trim(); });
  const leads = [];
  for (var i = 1; i < values.length; i++) {
    const row = values[i];
    var empty = true;
    for (var j = 0; j < row.length; j++) {
      if (row[j] !== '' && row[j] !== null) { empty = false; break; }
    }
    if (empty) continue;
    const o = { _row: i + 1 };
    heads.forEach(function (h, k) {
      var v = row[k];
      if (h === 'Timestamp') {
        // A numeric-only key: Sheets can't mistake it for a date and silently
        // reformat it, unlike a human-readable timestamp string.
        o._tsKey = v instanceof Date ? String(v.getTime()) : String(v || '');
      }
      if (v instanceof Date) v = fmtCell_(v, tz);
      o[h] = v;
    });
    leads.push(o);
  }
  return leads;
}

function getTaxRate_() {
  return Number(PropertiesService.getScriptProperties().getProperty('TAX_RATE')) || 25;
}
function setTaxRate(rate) {
  var n = Number(rate);
  if (!isNaN(n) && n >= 0 && n <= 100) {
    PropertiesService.getScriptProperties().setProperty('TAX_RATE', String(n));
  }
  return JSON.stringify({ taxRate: getTaxRate_() });
}

function getLeads() {
  migrateAdSpendToExpenses_();
  migrateReceiptsToYearFolders_();
  return JSON.stringify({
    leads: buildLeadsArray_(), tasks: getTasks_(), logs: getLog_(), templates: getTemplates_(),
    partners: getPartners_(), partnerLogs: getPartnerLog_(),
    textTemplates: getTextTemplates_(), partnerTemplates: getPartnerTemplates_(),
    partnerTextTemplates: getPartnerTextTemplates_(), oracleGate: oracleGateCheck_(),
    trash: getTrash(), recurring: getRecurring_(),
    topClientThreshold: Number(PropertiesService.getScriptProperties().getProperty('TOP_CLIENT_THRESHOLD')) || 2000,
    kvfOrder: getKvfOrder_(),
    kvfHidden: getKvfHidden_(),
    expenses: getExpenses_(),
    customExpenseCategories: getCustomExpenseCategories_(),
    taxRate: getTaxRate_(),
    vendorCategoryMemory: getVendorCategoryMemory_()
  });
}
// Same Script Properties JSON pattern as KVF_ORDER/KVF_HIDDEN above — a
// small app-wide list, not per-expense data. Deleting a category here only
// removes it from the dropdown going forward; existing expenses already
// tagged with that category keep the plain text string on their row, so
// nothing is lost or orphaned on delete.
function getCustomExpenseCategories_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('CUSTOM_EXPENSE_CATEGORIES') || '[]'); }
  catch (e) { return []; }
}
function addExpenseCategory(name) {
  name = String(name || '').trim();
  if (!name) return JSON.stringify({ ok: false, error: 'Category name is empty.' });
  var props = PropertiesService.getScriptProperties();
  var list = getCustomExpenseCategories_();
  var exists = list.some(function(c) { return c.toLowerCase() === name.toLowerCase(); });
  if (!exists) {
    list.push(name);
    props.setProperty('CUSTOM_EXPENSE_CATEGORIES', JSON.stringify(list));
  }
  return JSON.stringify({ ok: true, customExpenseCategories: list });
}
function deleteExpenseCategory(name) {
  var props = PropertiesService.getScriptProperties();
  var list = getCustomExpenseCategories_().filter(function(c) { return c !== name; });
  props.setProperty('CUSTOM_EXPENSE_CATEGORIES', JSON.stringify(list));
  return JSON.stringify({ ok: true, customExpenseCategories: list });
}
// Both stored as JSON in Script Properties, same pattern as every other
// app-wide (not per-lead) setting — this is deliberately universal, not
// something that varies contact to contact.
function getKvfOrder_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('KVF_ORDER') || '[]'); }
  catch (e) { return []; }
}
function getKvfHidden_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('KVF_HIDDEN') || '[]'); }
  catch (e) { return []; }
}
function saveFieldSettings(orderJson, hiddenJson) {
  var props = PropertiesService.getScriptProperties();
  try {
    var order = JSON.parse(orderJson || '[]');
    if (Array.isArray(order)) props.setProperty('KVF_ORDER', JSON.stringify(order));
  } catch (e) {}
  try {
    var hidden = JSON.parse(hiddenJson || '[]');
    if (Array.isArray(hidden)) props.setProperty('KVF_HIDDEN', JSON.stringify(hidden));
  } catch (e) {}
  return JSON.stringify({ ok: true });
}

// The Oracle's passive tips used to show every single time Insights was
// opened, which made them feel like static wallpaper instead of a genuine
// heads-up. This gate makes them show at most once every 6-9 days (picked
// randomly each time so it doesn't feel clockwork) — and only actually
// consumes that window if there's something real to say (see
// markOracleTipShown, called client-side only when a tip is displayed).
function oracleGateCheck_(){
  var props = PropertiesService.getScriptProperties();
  var last = Number(props.getProperty('ORACLE_TIP_LAST_SHOWN') || 0);
  var intervalDays = Number(props.getProperty('ORACLE_TIP_NEXT_INTERVAL') || 7);
  return !last || (Date.now() - last) >= intervalDays*24*60*60*1000;
}
function markOracleTipShown(){
  var props = PropertiesService.getScriptProperties();
  var nextInterval = 6 + Math.floor(Math.random()*4); // 6-9 days
  props.setProperty('ORACLE_TIP_LAST_SHOWN', String(Date.now()));
  props.setProperty('ORACLE_TIP_NEXT_INTERVAL', String(nextInterval));
  return JSON.stringify({ok:true});
}

// Lightweight response for lead-only operations (edit/add/delete a lead, log
// contact history). Returns just {leads, logs} instead of rebuilding every
// sheet in the app — the client already holds everything else in memory.
function getLeadsScoped_() {
  return JSON.stringify({ leads: buildLeadsArray_(), logs: getLog_(), trash: getTrash(), customFields: customLeadFields_() });
}

// Custom fields the buyer created during a CSV import (tracked explicitly in a
// Script Property, filtered to those still present as columns). The app adds
// these to the lead editor so imported custom data is visible and editable.
function customLeadFields_() {
  try {
    var tracked = JSON.parse(PropertiesService.getScriptProperties().getProperty('CUSTOM_LEAD_FIELDS') || '[]');
    if (!Array.isArray(tracked) || !tracked.length) return [];
    var present = {}; headers_(sheet_()).forEach(function (h) { present[String(h).trim()] = 1; });
    return tracked.filter(function (c) { return c && present[c]; });
  } catch (e) { return []; }
}

// Even lighter — for contact-history-only operations that never touch the
// Leads sheet at all.
function getLogsScoped_() {
  return JSON.stringify({ logs: getLog_() });
}

// The same lightweight-response pattern, extended to every other data type
// in the app. Each of these lets its own add/update/delete functions send
// back only what actually changed, instead of rebuilding all 10 sheets.
function getTasksScoped_() { return JSON.stringify({ tasks: getTasks_() }); }
function getPartnersScoped_() { return JSON.stringify({ partners: getPartners_() }); }
function getPartnerLogsScoped_() { return JSON.stringify({ partnerLogs: getPartnerLog_() }); }
function getTemplatesScoped_() { return JSON.stringify({ templates: getTemplates_() }); }
function getPartnerTemplatesScoped_() { return JSON.stringify({ partnerTemplates: getPartnerTemplates_() }); }
function getTextTemplatesScoped_() { return JSON.stringify({ textTemplates: getTextTemplates_() }); }
function getPartnerTextTemplatesScoped_() { return JSON.stringify({ partnerTextTemplates: getPartnerTextTemplates_() }); }

/* ---------- Partner contact log ---------- */
const PARTNER_LOG_SHEET = 'Partner Contact Log';

function partnerLogSheet_() {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(PARTNER_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PARTNER_LOG_SHEET);
    sh.appendRow(['When', 'Partner ID', 'Partner Name', 'Type', 'Note', 'Entry ID']);
  }
  return sh;
}

function getPartnerLog_() {
  const sh = partnerLogSheet_();
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const values = sh.getDataRange().getValues();
  const out = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[3]) continue;
    out.push({
      id: String(r[5] || ''),
      partnerId: String(r[1] || ''),
      when: r[0] instanceof Date ? fmtCell_(r[0], tz) : String(r[0] || ''),
      type: String(r[3] || ''),
      note: String(r[4] || '')
    });
  }
  return out;
}

function findPartnerLogRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, 6, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === id) return i + 2;
  }
  return 0;
}

function logPartnerContact(partnerId, partnerName, type, note) {
  partnerLogSheet_().appendRow([new Date(), partnerId, partnerName, type, note || '', Utilities.getUuid()]);
  return getPartnerLogsScoped_();
}

function updatePartnerLog(id, whenStr, note) {
  const sh = partnerLogSheet_();
  const rowNum = findPartnerLogRow_(sh, id);
  if (!rowNum) return getPartnerLogsScoped_();
  if (whenStr) sh.getRange(rowNum, 1).setValue(parseYMD_(whenStr));
  sh.getRange(rowNum, 5).setValue(note || '');
  return getPartnerLogsScoped_();
}

function deletePartnerLog(id) {
  const sh = partnerLogSheet_();
  const rowNum = findPartnerLogRow_(sh, id);
  if (rowNum) sh.deleteRow(rowNum);
  return getPartnerLogsScoped_();
}

/* ---------- Referral partners ---------- */
const PARTNERS_SHEET = 'Referral Partners';

function partnersSheet_() {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(PARTNERS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PARTNERS_SHEET);
    sh.appendRow(['Name', 'Email', 'Phone', 'Notes', 'Entry ID']);
  }
  return sh;
}

function getPartners_() {
  const sh = partnersSheet_();
  const values = sh.getDataRange().getValues();
  const out = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    out.push({
      id: String(r[4] || ''),
      name: String(r[0] || ''),
      email: String(r[1] || ''),
      phone: String(r[2] || ''),
      notes: String(r[3] || '')
    });
  }
  return out;
}

function findPartnerRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, 5, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === id) return i + 2;
  }
  return 0;
}

function addPartner(name, email, phone, notes) {
  const sh = partnersSheet_();
  sh.appendRow([name || '', email || '', phone || '', notes || '', Utilities.getUuid()]);
  return getPartnersScoped_();
}

function updatePartner(id, name, email, phone, notes) {
  const sh = partnersSheet_();
  const rowNum = findPartnerRow_(sh, id);
  if (!rowNum) return getPartnersScoped_();
  sh.getRange(rowNum, 1, 1, 4).setValues([[name || '', email || '', phone || '', notes || '']]);
  return getPartnersScoped_();
}

function deletePartner(id) {
  const sh = partnersSheet_();
  const rowNum = findPartnerRow_(sh, id);
  if (rowNum) sh.deleteRow(rowNum);
  return getPartnersScoped_();
}

/* ---------- Advertising spend (retired — migrated into Expenses) ---------- */
// The old standalone Ad Spend log has been folded into Expenses (category
// "Marketing & Advertising" + an optional Ad Source tag for ROI matching).
// These two functions are kept only so migrateAdSpendToExpenses_() below
// can read the old sheet one time; the add/update/delete/bundling
// functions that powered the old live-editing UI have been removed since
// nothing calls them anymore.
const ADSPEND_SHEET = 'Ad Spend';

function adSpendSheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(ADSPEND_SHEET); // may be null — that's fine, means nothing to migrate
}

function getAdSpend_() {
  const sh = adSpendSheet_();
  if (!sh) return [];
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const values = sh.getDataRange().getValues();
  const out = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[1] && !r[2]) continue;
    out.push({
      id: String(r[4] || ''),
      date: r[0] instanceof Date ? fmtCell_(r[0], tz) : String(r[0] || ''),
      source: String(r[1] || ''),
      amount: Number(r[2]) || 0,
      note: String(r[3] || '')
    });
  }
  return out;
}

// One-time, idempotent: copies every historical Ad Spend entry into
// Expenses (as "Marketing & Advertising", tagged with Ad Source so the
// ROI-by-channel math keeps working), then archives the old sheet by
// renaming it rather than deleting it — the raw history stays intact and
// inspectable, it just stops being read by the app. Runs automatically
// on every getLeads() call but only actually does anything once, guarded
// by a Script Property flag.
function migrateAdSpendToExpenses_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('ADSPEND_MIGRATED') === 'true') return;
  const oldSh = adSpendSheet_();
  if (!oldSh) { props.setProperty('ADSPEND_MIGRATED', 'true'); return; }
  const entries = getAdSpend_();
  if (entries.length) {
    const sh = expensesSheet_();
    const now = new Date();
    const out = entries.map(function(a) {
      var src = a.source || '';
      return [
        a.date ? parseYMD_(a.date) : now,
        src || 'Ad Spend',
        Number(a.amount) || 0,
        'Marketing & Advertising',
        a.note || '',
        'Ad Spend Migration',
        now,
        '', '',
        src
      ];
    });
    sh.getRange(sh.getLastRow() + 1, 1, out.length, 10).setValues(out);
  }
  try { oldSh.setName('Ad Spend (archived ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + ')'); } catch (e) {}
  props.setProperty('ADSPEND_MIGRATED', 'true');
}

function addTask(name, due, repeat, id) {
  const sh = tasksSheet_();
  var created = id ? new Date(id) : new Date();
  if (isNaN(created.getTime())) created = new Date();
  sh.appendRow([created, name, due ? parseYMD_(due) : '', '', repeat || '']);
  try { runFrequentAutomations(); } catch (e) {}
  try { scheduleCalendarRebuild_(); } catch (e) {}
  return getTasksScoped_();
}

function findTaskRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i][0];
    var iso = v instanceof Date ? v.toISOString() : String(v);
    if (iso === id) return i + 2;
  }
  return 0;
}

function toggleTask(id) {
  const sh = tasksSheet_();
  const rowNum = findTaskRow_(sh, id);
  if (!rowNum) return getTasksScoped_();
  const cur = String(sh.getRange(rowNum, 4).getValue() || '');
  sh.getRange(rowNum, 4).setValue(cur === 'Yes' ? '' : 'Yes');
  // Checking off a repeating task schedules the next occurrence.
  if (cur !== 'Yes') {
    const repeat = String(sh.getRange(rowNum, 5).getValue() || '');
    const dueV = sh.getRange(rowNum, 3).getValue();
    if (repeat && dueV) {
      const d = dueV instanceof Date ? dueV : new Date(parseYMD_(String(dueV)));
      if (!isNaN(d.getTime())) {
        const nd = nextDue_(d, repeat);
        if (nd) sh.appendRow([new Date(), String(sh.getRange(rowNum, 2).getValue()), nd, '', repeat]);
      }
    }
  }
  try { runFrequentAutomations(); } catch (e) {}
  try { scheduleCalendarRebuild_(); } catch (e) {}
  return getTasksScoped_();
}

function deleteTask(id) {
  const sh = tasksSheet_();
  const rowNum = findTaskRow_(sh, id);
  if (rowNum) sh.deleteRow(rowNum);
  try { runFrequentAutomations(); } catch (e) {}
  try { scheduleCalendarRebuild_(); } catch (e) {}
  return getTasksScoped_();
}

function updateTask(id, name, due, repeat) {
  const sh = tasksSheet_();
  const rowNum = findTaskRow_(sh, id);
  if (!rowNum) return getTasksScoped_();
  sh.getRange(rowNum, 2).setValue(name);
  sh.getRange(rowNum, 3).setValue(due ? parseYMD_(due) : '');
  sh.getRange(rowNum, 5).setValue(repeat || '');
  try { runFrequentAutomations(); } catch (e) {}
  try { scheduleCalendarRebuild_(); } catch (e) {}
  return getTasksScoped_();
}

// Every new booking — even from a repeat client — gets a genuinely new
// row, never a reused one. Reusing a row would silently overwrite that
// client's last contract, invoice number, and receipts, and would also
// quietly break Top Clients (which counts real bookings by counting rows)
// and December Rebooking (which relies on a past booking's row still
// existing). This only ever copies contact info explicitly — everything
// else is left as a blank cell in the new row, which stays correct even
// as new columns get added to the sheet later, since there's no list of
// "fields to clear" to maintain.
function duplicateLeadAsNewBooking(row) {
  try {
    const sh = sheet_();
    const heads = headers_(sh);
    const srcRow = sh.getRange(row, 1, 1, heads.length).getValues()[0];
    const get = function (name) { var i = heads.indexOf(name); return i > -1 ? srcRow[i] : ''; };

    const newRow = new Array(heads.length).fill('');
    function setCol(name, val) { var i = heads.indexOf(name); if (i > -1) newRow[i] = val; }

    setCol('Timestamp', new Date());
    setCol('Customer Name', get('Customer Name'));
    setCol('Company or Organization', get('Company or Organization'));
    setCol('Company Address', get('Company Address'));
    setCol('Phone number', get('Phone number'));
    setCol('E-mail', get('E-mail'));
    setCol('Status', 'New');
    // Carries the link to the SAME existing Google Contact forward, so the
    // sync job updates that one instead of creating a second contact for
    // someone it has no way of knowing it's already met — the sync's own
    // same-name protection (there to keep two different people who share a
    // name from getting merged into one) would otherwise treat this new
    // row as a stranger.
    setCol('Google Contact ID', get('Google Contact ID'));

    sh.appendRow(newRow);
    const newRowNum = sh.getLastRow();

    return JSON.stringify({ ok: true, newRow: newRowNum, leads: JSON.parse(getLeads()).leads });
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
}

function updateLead(rowNum, updates) {
  const sh = sheet_();
  const heads = headers_(sh);
  const statusColIdx = heads.indexOf('Status');
  const wasBooked = statusColIdx > -1 && String(sh.getRange(rowNum, statusColIdx + 1).getValue() || '') === 'Booked';
  Object.keys(updates).forEach(function (key) {
    var col = heads.indexOf(key) + 1;
    if (!col) {
      // Column doesn't exist yet (e.g. "Referred To") — create it.
      col = heads.length + 1;
      sh.getRange(1, col).setValue(key);
      heads.push(key);
    }
    var v = updates[key];
    if ((key === 'Followup' || key === 'Date of Event' || key === 'Start Time' || key === 'End Time') && v) v = parseYMD_(v);
    sh.getRange(rowNum, col).setValue(v);
    if (key === 'Followup' && v) {
      // A deliberate new follow-up date means re-engaging, even if the
      // review cycle previously finished — clear that marker so the
      // automation doesn't treat this lead as permanently done.
      const doneCol = ensureReviewDoneCol_(sh);
      sh.getRange(rowNum, doneCol).setValue(false);
    }
  });
  try { runFrequentAutomations(); } catch (e) { /* automations never block a save */ }
  try { scheduleCalendarRebuild_(); } catch (e) { /* calendar hiccups never block a save */ }
  try { syncBookedCalendarEvent_(sh, rowNum, wasBooked); } catch (e) { /* booking calendar hiccups never block a save */ }
  try {
    var depositReceivedCol = heads.indexOf('Deposit Received') + 1;
    var balancePaidCol = heads.indexOf('Balance Paid') + 1;
    var depositIsYes = depositReceivedCol > 0 && String(sh.getRange(rowNum, depositReceivedCol).getValue()) === 'Yes';
    var balanceIsYes = balancePaidCol > 0 && String(sh.getRange(rowNum, balancePaidCol).getValue()) === 'Yes';
    // Regenerates on the Yes flip itself, same as before — but now also
    // when someone corrects the date afterward on a deposit/balance
    // that's already marked received, e.g. entering the real historical
    // date for a client booked before this field existed, or fixing an
    // auto-stamped "today" that turned out to be wrong. Otherwise a
    // date correction would silently never reach an already-generated
    // receipt.
    if (updates['Deposit Received'] === 'Yes' || (('Deposit Received Date' in updates) && depositIsYes)) {
      maybeGenerateReceipt_(sh, rowNum, 'deposit');
    }
    if (updates['Balance Paid'] === 'Yes' || (('Balance Paid Date' in updates) && balanceIsYes)) {
      maybeGenerateReceipt_(sh, rowNum, 'balance');
    }
  } catch (e) { /* receipt generation never blocks a save */ }
  return getLeadsScoped_();
}

function addLead(fields) {
  const sh = sheet_();
  const heads = headers_(sh);
  const row = heads.map(function (h) {
    if (h === 'Timestamp') return new Date();
    var v = fields[h];
    if (v === undefined || v === null) return '';
    if ((h === 'Followup' || h === 'Date of Event') && v) return parseYMD_(v);
    return v;
  });
  sh.appendRow(row);
  const newRow = sh.getLastRow();
  try { runFrequentAutomations(); } catch (e) {}
  try { scheduleCalendarRebuild_(); } catch (e) {}
  try { syncBookedCalendarEvent_(sh, newRow, false); } catch (e) {}
  return getLeadsScoped_();
}

// Bulk insert for the "import your existing client list" CSV importer. The
// client has already parsed the file, mapped each spreadsheet column to a lead
// field, let the user pick a Status for the batch, and reviewed/deduped the
// rows — so this trusts the reviewed batch and writes it in one setValues call
// rather than one appendRow per lead. Deliberately does NOT run follow-up
// automations or create calendar events: an imported back-catalog of past
// clients shouldn't fire reminders or spawn calendar events for shows that
// already happened. Each incoming object is keyed by header name, exactly like
// addLead's `fields`.
function importLeadsBatch(rowsJson) {
  const sh = sheet_();
  const rows = JSON.parse(rowsJson);
  var heads = headers_(sh);
  // Any field key that isn't already a column becomes a NEW column — a custom
  // field the buyer chose to create during import. Track the names in a Script
  // Property so the app can surface them in the lead editor.
  var known = {}; heads.forEach(function (h) { known[h] = 1; });
  var newCols = [];
  rows.forEach(function (f) {
    Object.keys(f).forEach(function (k) {
      if (k && k !== '_row' && !known[k] && newCols.indexOf(k) === -1) newCols.push(k);
    });
  });
  if (newCols.length) {
    sh.getRange(1, heads.length + 1, 1, newCols.length).setValues([newCols]);
    heads = headers_(sh); // re-read so the new columns are included below
    try {
      var props = PropertiesService.getScriptProperties();
      var tracked = [];
      try { tracked = JSON.parse(props.getProperty('CUSTOM_LEAD_FIELDS') || '[]'); } catch (e) {}
      if (!Array.isArray(tracked)) tracked = [];
      newCols.forEach(function (c) { if (tracked.indexOf(c) === -1) tracked.push(c); });
      props.setProperty('CUSTOM_LEAD_FIELDS', JSON.stringify(tracked));
    } catch (e) {}
  }
  const now = new Date();
  const out = rows.map(function (fields) {
    return heads.map(function (h) {
      if (h === 'Timestamp') return now;
      var v = fields[h];
      if (v === undefined || v === null) return '';
      if ((h === 'Followup' || h === 'Date of Event') && v) return parseYMD_(v);
      return v;
    });
  });
  if (out.length) {
    sh.getRange(sh.getLastRow() + 1, 1, out.length, heads.length).setValues(out);
  }
  return getLeadsScoped_();
}

/* ---------- Booking calendar (primary calendar) ---------- */
/**
 * Creates a calendar event on the primary Google Calendar the moment a lead
 * transitions INTO Booked status, and keeps that event in sync afterward if
 * the lead's booking details change while it stays Booked. The event's ID
 * is tracked in a dedicated "Booking Calendar Event ID" column so a lead
 * only ever gets one event, never a duplicate.
 *
 * Deliberately does NOT touch leads that were already Booked before this
 * feature existed — those have no event ID yet, and since their status
 * isn't actually transitioning (it was already Booked before this save,
 * too), the guard below leaves them alone rather than creating an event
 * for a show that's already on the books. Only a genuine New/Pending/etc.
 * \u2192 Booked transition creates a fresh event going forward.
 */
function ensureBookingEventCol_(sh) {
  const heads = headers_(sh);
  var col = heads.indexOf('Booking Calendar Event ID') + 1;
  if (!col) {
    col = heads.length + 1;
    sh.getRange(1, col).setValue('Booking Calendar Event ID');
  }
  return col;
}

function combineDateTime_(dateVal, timeVal, fallbackHour) {
  if (!dateVal) return null;
  const d = new Date(dateVal);
  if (isNaN(d.getTime())) return null;
  const out = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  if (timeVal) {
    const t = new Date(timeVal);
    if (!isNaN(t.getTime())) {
      out.setHours(t.getHours(), t.getMinutes(), 0, 0);
      return out;
    }
  }
  if (fallbackHour == null) return null;
  out.setHours(fallbackHour, 0, 0, 0);
  return out;
}

function syncBookedCalendarEvent_(sh, rowNum, wasBooked) {
  const heads = headers_(sh);
  const row = sh.getRange(rowNum, 1, 1, heads.length).getValues()[0];
  const get = function (h) { var i = heads.indexOf(h); return i > -1 ? row[i] : ''; };

  const status = String(get('Status') || '');
  if (status !== 'Booked') return;

  const idCol = ensureBookingEventCol_(sh);
  const existingId = sh.getRange(rowNum, idCol).getValue();

  // Pre-existing booking from before this feature shipped — no ID, and it
  // was already Booked before this save too, so it's not a new transition.
  if (!existingId && wasBooked) return;

  const eventDate = get('Date of Event');
  if (!eventDate) return; // nothing to schedule against yet

  const start = combineDateTime_(eventDate, get('Start Time'), REMINDER_HOUR);
  if (!start) return;
  var end = combineDateTime_(eventDate, get('End Time'), null);
  if (!end || end <= start) end = new Date(start.getTime() + 2 * 60 * 60000);

  const name = String(get('Customer Name') || '').trim();
  const company = String(get('Company or Organization') || '').trim();
  const who = company ? (company + (name ? ', ' + name : '')) : name;
  const title = who ? ('Magic Show for ' + who) : 'Magic Show';

  const lines = [];
  if (name) lines.push('Client: ' + name);
  if (company) lines.push('Company: ' + company);
  if (get('Service')) lines.push('Service: ' + get('Service'));
  if (get('Audience Size')) lines.push('Audience Size: ' + get('Audience Size'));
  if (get('Event Location')) lines.push('Location: ' + get('Event Location'));
  if (get('Phone number')) lines.push('Phone: ' + get('Phone number'));
  if (get('E-mail')) lines.push('Email: ' + get('E-mail'));
  const description = lines.join('\n');
  const location = String(get('Event Location') || '');

  const cal = CalendarApp.getDefaultCalendar();
  var ev = null;
  if (existingId) {
    try { ev = cal.getEventById(existingId); } catch (e) { ev = null; }
  }
  if (ev) {
    ev.setTitle(title);
    ev.setTime(start, end);
    ev.setDescription(description);
    ev.setLocation(location);
  } else {
    ev = cal.createEvent(title, start, end, { description: description, location: location });
    sh.getRange(rowNum, idCol).setValue(ev.getId());
  }
}

/* ---------- Calendar reminders ---------- */

function calendar_() {
  // Remember the calendar by its permanent ID so we never create duplicates.
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('CAL_ID');
  if (savedId) {
    try {
      const c = CalendarApp.getCalendarById(savedId);
      if (c) return c;
    } catch (e) {}
  }
  const found = CalendarApp.getCalendarsByName(getConfig_().CAL_NAME);
  if (found.length) { props.setProperty('CAL_ID', found[0].getId()); return found[0]; }
  const cal = CalendarApp.createCalendar(getConfig_().CAL_NAME);
  props.setProperty('CAL_ID', cal.getId());
  try { cal.setColor(CalendarApp.Color.YELLOW); } catch (e) {}
  return cal;
}

/**
 * Run this ONCE from the editor to remove duplicate
 * the Follow-ups calendars. Keeps one, deletes the
 * rest, then rebuilds all reminders on the keeper.
 */
function cleanupDuplicateCalendars() {
  const all = CalendarApp.getCalendarsByName(getConfig_().CAL_NAME);
  if (all.length) {
    PropertiesService.getScriptProperties().setProperty('CAL_ID', all[0].getId());
    for (var i = 1; i < all.length; i++) {
      try { all[i].deleteCalendar(); } catch (e) {}
    }
  }
  syncFollowUps();
}

/**
 * Rebuilds follow-up reminders. Rules:
 *  - Only follow-up dates from today forward
 *  - Only dates inside the current calendar year
 *  - Skips leads marked Lost or Completed
 * Runs daily via trigger and after every save from the app.
 */
// Follow-up stages for Booked leads, as days offset from the event date.
// The old 30-days-before stage existed to prompt preparing the balance
// invoice by hand — now that the sign link automatically starts asking
// for the balance on its own once the event's within range, there's
// nothing left for that stage to actually remind you to do.
const FOLLOWUP_STAGE_OFFSETS = [-7, 1];
const FOLLOWUP_STAGE_LABELS = ['1 week before event', 'Review request'];

function stageDatesForEvent_(eventDate) {
  const ed = new Date(eventDate); ed.setHours(0, 0, 0, 0);
  return FOLLOWUP_STAGE_OFFSETS.map(function (offset) {
    const d = new Date(ed); d.setDate(d.getDate() + offset); return d;
  });
}

/**
 * For every Booked lead with an event date, keeps the follow-up date stepping
 * through 1-week-before -> review-request automatically.
 * Only ever touches a follow-up date that exactly matches one of those three
 * auto-generated stage dates — a custom date you set by hand is never
 * overwritten. Advancement happens once a stage date has passed.
 */
/**
 * Sets the initial 1-week-before follow-up the moment a lead becomes
 * Booked. After that, progression from stage to stage is driven by the
 * person explicitly confirming each one (see markFollowupStageDone) rather
 * than happening automatically just because a date passed — the trigger
 * dates only control when a stage's card becomes due, not when it advances.
 *
 * Two safety nets keep things from getting permanently stuck if a stage is
 * never manually confirmed:
 *   1. Once the event date itself has passed, any lead still sitting on an
 *      earlier stage snaps straight to Review Request — asking for a
 *      review only makes sense once the show has actually happened.
 *   2. Two weeks past the event, if it's still sitting on Review Request
 *      unconfirmed, the follow-up clears on its own.
 */
function ensureReviewDoneCol_(sh) {
  const heads = headers_(sh);
  var col = heads.indexOf('Review Cycle Done') + 1;
  if (!col) {
    col = heads.length + 1;
    sh.getRange(1, col).setValue('Review Cycle Done');
  }
  return col;
}

function autoManageBookedFollowups_() {
  const sh = sheet_();
  const heads = headers_(sh);
  const statusCol = heads.indexOf('Status');
  const eventCol = heads.indexOf('Date of Event');
  const followupCol = heads.indexOf('Followup');
  const nameCol = heads.indexOf('Customer Name');
  const tsCol = heads.indexOf('Timestamp');
  if (statusCol < 0 || eventCol < 0 || followupCol < 0) return;
  const doneCol = ensureReviewDoneCol_(sh);

  const values = sh.getDataRange().getValues();
  const today = new Date(); today.setHours(0, 0, 0, 0);

  for (var i = 1; i < values.length; i++) {
    const row = values[i];
    const statusNow = String(row[statusCol]);
    if (statusNow !== 'Booked' && statusNow !== 'Completed') continue;
    const eventVal = row[eventCol];
    if (!(eventVal instanceof Date) || isNaN(eventVal.getTime())) continue;

    const stages = stageDatesForEvent_(eventVal);
    const curVal = row[followupCol];
    var curDate = curVal instanceof Date ? new Date(curVal) : null;
    if (curDate) curDate.setHours(0, 0, 0, 0);

    var curStageIdx = -1;
    if (curDate) {
      for (var s = 0; s < stages.length; s++) {
        if (stages[s].getTime() === curDate.getTime()) { curStageIdx = s; break; }
      }
    }

    var targetStageIdx = -1, clearIt = false;
    const eventDay = new Date(eventVal); eventDay.setHours(0, 0, 0, 0);
    const reviewIdx = stages.length - 1;
    const cycleAlreadyDone = !!row[doneCol - 1];

    if (!curVal) {
      // Blank follow-up means one of two very different things: never
      // assigned yet, or the cycle already finished and was cleared. The
      // marker column is what tells them apart — without it, a completed
      // Review Request for a past event would get reassigned right back.
      if (cycleAlreadyDone) continue;
      for (var s2 = 0; s2 < stages.length; s2++) {
        if (stages[s2].getTime() >= today.getTime()) { targetStageIdx = s2; break; }
      }
      if (targetStageIdx === -1) targetStageIdx = reviewIdx; // event's already past
    } else if (curStageIdx > -1 && curStageIdx < reviewIdx && eventDay.getTime() < today.getTime()) {
      // Safety net 1: event has passed but we're still sitting on an earlier
      // stage that was never confirmed — jump straight to Review Request.
      targetStageIdx = reviewIdx;
    } else if (curStageIdx === reviewIdx && stages[reviewIdx].getTime() < today.getTime()) {
      // Safety net 2: Review Request has sat unconfirmed for 2+ weeks past
      // the event — clear it rather than leave it stale forever.
      const clearAfter = new Date(eventVal); clearAfter.setHours(0, 0, 0, 0);
      clearAfter.setDate(clearAfter.getDate() + 14);
      if (today.getTime() >= clearAfter.getTime()) clearIt = true;
    }
    if (!clearIt && (targetStageIdx === -1 || targetStageIdx === curStageIdx)) continue;

    if (clearIt) {
      sh.getRange(i + 1, followupCol + 1).setValue('');
      sh.getRange(i + 1, doneCol).setValue(true);
      try {
        const key0 = tsCol > -1 ? (row[tsCol] instanceof Date ? String(row[tsCol].getTime()) : String(row[tsCol])) : '';
        logSheet_().appendRow([new Date(), key0, String(row[nameCol] || ''), 'Follow-up',
          'Follow-up cleared \u2014 2 weeks past event, review cycle complete', Utilities.getUuid()]);
      } catch (e) {}
      continue;
    }

    sh.getRange(i + 1, followupCol + 1).setValue(stages[targetStageIdx]);
    try {
      const key = tsCol > -1 ? (row[tsCol] instanceof Date ? String(row[tsCol].getTime()) : String(row[tsCol])) : '';
      logSheet_().appendRow([new Date(), key, String(row[nameCol] || ''), 'Follow-up',
        FOLLOWUP_STAGE_LABELS[targetStageIdx] + ' (auto-scheduled)', Utilities.getUuid()]);
    } catch (e) {}
  }
}

/**
 * Explicitly confirms the current follow-up stage is done and advances the
 * lead to the next one's trigger date. If it was already on Review Request,
 * this clears the follow-up entirely — the sequence is complete.
 */
function markFollowupStageDone(rowNum, note) {
  const sh = sheet_();
  const heads = headers_(sh);
  const eventCol = heads.indexOf('Date of Event');
  const followupCol = heads.indexOf('Followup');
  const nameCol = heads.indexOf('Customer Name');
  const tsCol = heads.indexOf('Timestamp');
  if (eventCol < 0 || followupCol < 0) return getLeadsScoped_();

  const row = sh.getRange(rowNum, 1, 1, sh.getLastColumn()).getValues()[0];
  const eventVal = row[eventCol];
  if (!(eventVal instanceof Date) || isNaN(eventVal.getTime())) return getLeadsScoped_();

  const stages = stageDatesForEvent_(eventVal);
  const curVal = row[followupCol];
  var curDate = curVal instanceof Date ? new Date(curVal) : null;
  if (curDate) curDate.setHours(0, 0, 0, 0);

  var curStageIdx = -1;
  if (curDate) {
    for (var s = 0; s < stages.length; s++) {
      if (stages[s].getTime() === curDate.getTime()) { curStageIdx = s; break; }
    }
  }
  if (curStageIdx === -1) return getLeadsScoped_(); // not on a recognized stage — nothing to confirm

  const key = tsCol > -1 ? (row[tsCol] instanceof Date ? String(row[tsCol].getTime()) : String(row[tsCol])) : '';
  const label = FOLLOWUP_STAGE_LABELS[curStageIdx];
  const doneCol = ensureReviewDoneCol_(sh);

  if (curStageIdx + 1 < stages.length) {
    sh.getRange(rowNum, followupCol + 1).setValue(stages[curStageIdx + 1]);
  } else {
    sh.getRange(rowNum, followupCol + 1).setValue('');
    sh.getRange(rowNum, doneCol).setValue(true);
  }
  try {
    logSheet_().appendRow([new Date(), key, String(row[nameCol] || ''), 'Follow-up',
      label + ' \u2014 marked done' + (note ? ': ' + note : ''), Utilities.getUuid()]);
  } catch (e) {}
  return getLeadsScoped_();
}

/**
 * Once a Booked lead's event has genuinely passed — either the event date
 * itself is in the past, or today IS the event date and the show's End Time
 * has already come and gone — flips its status to Completed automatically.
 * If today is the event date but no End Time is set, it's left alone rather
 * than guessed at; it'll auto-complete the next day once the date itself
 * has passed, same as any other past event.
 *
 * Also runs the reverse direction: a lead sitting at Completed whose event
 * date has since been moved back into the future gets reverted to Booked.
 * That's a real scenario, not just a hypothetical — it's exactly what
 * happens when correcting a date after a show was marked Completed (rightly
 * or, before this fix, by mistake), and a "Completed" show dated in the
 * future is never a state worth leaving stuck.
 */
function autoCompleteBookedShows_() {
  const sh = sheet_();
  const heads = headers_(sh);
  const statusCol = heads.indexOf('Status');
  const eventCol = heads.indexOf('Date of Event');
  const endCol = heads.indexOf('End Time');
  const nameCol = heads.indexOf('Customer Name');
  const tsCol = heads.indexOf('Timestamp');
  if (statusCol < 0 || eventCol < 0) return;

  const values = sh.getDataRange().getValues();
  const now = new Date();
  const today = new Date(); today.setHours(0, 0, 0, 0);

  for (var i = 1; i < values.length; i++) {
    const row = values[i];
    const statusNow = String(row[statusCol]);
    const eventVal = row[eventCol];
    const hasValidEventDate = eventVal instanceof Date && !isNaN(eventVal.getTime());

    if (statusNow === 'Completed') {
      if (!hasValidEventDate) continue;
      const eventDayC = new Date(eventVal); eventDayC.setHours(0, 0, 0, 0);
      if (eventDayC.getTime() > today.getTime()) {
        sh.getRange(i + 1, statusCol + 1).setValue('Booked');
        try {
          const key = tsCol > -1 ? (row[tsCol] instanceof Date ? String(row[tsCol].getTime()) : String(row[tsCol])) : '';
          logSheet_().appendRow([new Date(), key, String(row[nameCol] || ''), 'Note',
            'Status auto-reverted to Booked \u2014 the event date was moved back to the future', Utilities.getUuid()]);
        } catch (e) {}
      }
      continue;
    }

    if (statusNow !== 'Booked') continue;
    if (!hasValidEventDate) continue;
    const eventDay = new Date(eventVal); eventDay.setHours(0, 0, 0, 0);
    if (eventDay.getTime() > today.getTime()) continue; // still upcoming

    var shouldComplete = false;
    if (eventDay.getTime() < today.getTime()) {
      shouldComplete = true; // the whole event day has already passed
    } else {
      // Event is today — only complete once the actual End Time has passed.
      // End Time is stored as a time-only cell (a real time-of-day, but
      // attached to a meaningless placeholder date the way Sheets stores
      // time-formatted cells) — it has to be combined with today's real
      // date before comparing against "now". Comparing the raw cell value
      // straight against "now" would always read as already-past, since
      // the placeholder date is always far in Sheets' past relative to any
      // real "now" — which is exactly what caused shows to complete the
      // instant their date became today, regardless of actual end time.
      const endVal = endCol > -1 ? row[endCol] : null;
      const realEnd = endVal ? combineDateTime_(eventVal, endVal, null) : null;
      if (realEnd && now.getTime() >= realEnd.getTime()) {
        shouldComplete = true;
      }
    }
    if (!shouldComplete) continue;

    sh.getRange(i + 1, statusCol + 1).setValue('Completed');
    try {
      const key = tsCol > -1 ? (row[tsCol] instanceof Date ? String(row[tsCol].getTime()) : String(row[tsCol])) : '';
      logSheet_().appendRow([new Date(), key, String(row[nameCol] || ''), 'Note',
        'Status auto-changed to Completed \u2014 the event has passed', Utilities.getUuid()]);
    } catch (e) {}
  }
}

/**
 * Leads can now arrive from outside the app entirely — e.g. a Zapier
 * automation feeding the website's contact form straight into this sheet.
 * The Timestamp column is critical (it's the permanent key contact history
 * and follow-ups are tracked against), so this backfills it automatically
 * for any row missing one, rather than relying on every external source
 * getting that field exactly right.
 */
function backfillMissingTimestamps_() {
  const sh = sheet_();
  const heads = headers_(sh);
  const tsCol = heads.indexOf('Timestamp');
  const nameCol = heads.indexOf('Customer Name');
  if (tsCol < 0 || nameCol < 0) return;
  const values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[nameCol]) continue; // skip truly empty rows
    if (!row[tsCol]) sh.getRange(i + 1, tsCol + 1).setValue(new Date());
  }
}

/**
 * A brand-new lead (status New) with no follow-up set yet — most commonly
 * one that just arrived via the website's contact form through Zapier,
 * bypassing the app's own "add a lead" flow entirely — gets its follow-up
 * defaulted to right now, so it shows up immediately as needing attention
 * instead of sitting invisible with nothing prompting a first outreach.
 * Deliberately scoped to New only: a Lost or Referral lead with no
 * follow-up shouldn't get one assigned automatically.
 */
function defaultFollowupForNewLeads_() {
  const sh = sheet_();
  const heads = headers_(sh);
  const statusCol = heads.indexOf('Status');
  const followupCol = heads.indexOf('Followup');
  if (statusCol < 0 || followupCol < 0) return;
  const values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[statusCol]) !== 'New') continue;
    if (row[followupCol]) continue; // already has one — leave it alone
    sh.getRange(i + 1, followupCol + 1).setValue(new Date());
  }
}

/**
 * Leads arriving through the Zapier email-parser automation can sometimes
 * pick up extra junk in the Notes or Location fields — the parser grabbing
 * everything from a field's value through to the end of the notification
 * email, including the next field's label and Squarespace's own footer
 * text. This trims anything at or after those known junk markers, so a
 * parser boundary issue can't leave messy data sitting in the app.
 */
function cleanupImportedJunk_() {
  const sh = sheet_();
  const heads = headers_(sh);
  const notesCol = heads.indexOf('Notes about interaction');
  const locCol = heads.indexOf('Event Location');
  if (notesCol < 0 && locCol < 0) return;

  const junkMarkers = ['Sent via form submission', 'Additional Details:', 'squarespace.info', 'Manage Submissions', 'Create Invoice'];
  function trimmed(text) {
    var s = String(text || '');
    var cutIdx = -1;
    junkMarkers.forEach(function (marker) {
      var idx = s.indexOf(marker);
      if (idx > -1 && (cutIdx === -1 || idx < cutIdx)) cutIdx = idx;
    });
    return cutIdx > -1 ? s.slice(0, cutIdx).trim() : s;
  }

  const values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    const row = values[i];
    if (locCol > -1 && row[locCol]) {
      const c = trimmed(row[locCol]);
      if (c !== row[locCol]) sh.getRange(i + 1, locCol + 1).setValue(c);
    }
    if (notesCol > -1 && row[notesCol]) {
      const c2 = trimmed(row[notesCol]);
      if (c2 !== row[notesCol]) sh.getRange(i + 1, notesCol + 1).setValue(c2);
    }
  }
}

/**
 * Lightweight lead data-hygiene automations only — deliberately no Calendar
 * calls in here. Safe to run frequently (every 5 minutes) so leads arriving
 * through Zapier or other outside tools get their defaults applied quickly.
 * Kept separate from the Calendar rebuild in syncFollowUps() on purpose:
 * Google's daily Calendar quota can't handle a full delete-and-recreate of
 * every event every 5 minutes — that rebuild only ever needed to run once a
 * day, plus right after an actual save from the app.
 */
/**
 * Zapier's Google Sheets action can write "Date of your event" through as
 * plain text (e.g. "August 24, 2026") rather than a real date value, since
 * writing a row via an API doesn't go through the same smart-entry parsing
 * Sheets applies when a value is typed directly into a cell. A plain-text
 * date silently fails every check in this codebase that expects a real
 * Date object — sorting, the day-of-show completion automation, Insights,
 * anything using `instanceof Date`. This converts anything parseable in
 * the Date of Event column into a real date in place, so it works the
 * same regardless of how the text arrived. Confirmed safe from the classic
 * UTC-shift bug: this month-name format parses as local midnight, not UTC
 * midnight, so it can't land on the wrong calendar day.
 */
function normalizeEventDateColumn_() {
  const sh = sheet_();
  const heads = headers_(sh);
  const eventCol = heads.indexOf('Date of Event');
  if (eventCol < 0) return;
  const values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    const v = values[i][eventCol];
    if (!v || v instanceof Date) continue; // blank, or already a real date — nothing to do
    const parsed = new Date(v);
    if (!isNaN(parsed.getTime())) {
      sh.getRange(i + 1, eventCol + 1).setValue(parsed);
    }
  }
}

function runFrequentAutomations() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    backfillMissingTimestamps_();
    normalizeEventDateColumn_();
    defaultFollowupForNewLeads_();
    cleanupImportedJunk_();
    autoCompleteBookedShows_();
    autoManageBookedFollowups_();
  } finally {
    lock.releaseLock();
  }
}

const CAL_REBUILD_TRIGGER_FN = 'runCalendarRebuildJob';

/**
 * Call this from any save path instead of the full syncFollowUps(). It
 * schedules the slow Calendar delete-and-recreate to run in the background
 * a couple seconds later, so the save itself returns to the app instantly
 * instead of waiting on Calendar API calls. If several saves happen in
 * quick succession, this debounces down to a single rebuild afterward
 * rather than stacking one per save.
 */
function scheduleCalendarRebuild_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === CAL_REBUILD_TRIGGER_FN) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(CAL_REBUILD_TRIGGER_FN).timeBased().after(3000).create();
}

function runCalendarRebuildJob() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === CAL_REBUILD_TRIGGER_FN) ScriptApp.deleteTrigger(t);
  });
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    syncFollowUps_();
  } finally {
    lock.releaseLock();
  }
}

function syncFollowUps() {
  // Only one sync at a time — prevents duplicate calendars and events.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    backfillMissingTimestamps_();
    defaultFollowupForNewLeads_();
    cleanupImportedJunk_();
    autoCompleteBookedShows_();
    autoManageBookedFollowups_();
    syncFollowUps_();
  } finally {
    lock.releaseLock();
  }
}

function syncFollowUps_() {
  try { generateRecurringExpenses_(); } catch (e) {}
  const cal = calendar_();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const yearEnd = new Date(today.getFullYear() + 1, 0, 1);

  // The calendar is dedicated to this app, so clear and rebuild.
  cal.getEvents(today, new Date(today.getFullYear() + 2, 0, 1))
    .forEach(function (ev) { try { ev.deleteEvent(); } catch (e) {} });

  const data = JSON.parse(getLeads());
  data.leads.forEach(function (L) {
    const raw = L['Followup'];
    if (!raw) return;
    const status = String(L['Status'] || '').toLowerCase();
    if (status === 'lost' || status === 'completed') return;
    const d = new Date(raw);
    if (isNaN(d.getTime())) return;
    // If a time of day was chosen, remind at that exact time; otherwise 9 AM.
    const hasTime = d.getHours() !== 0 || d.getMinutes() !== 0;
    if (!hasTime) d.setHours(REMINDER_HOUR, 0, 0, 0);
    const d0 = new Date(d); d0.setHours(0, 0, 0, 0);
    if (d0 < today || d0 >= yearEnd) return;

    const name = L['Customer Name'] || 'Lead';
    const bits = [];
    if (L['Phone number']) bits.push('Phone: ' + L['Phone number']);
    if (L['E-mail']) bits.push('Email: ' + L['E-mail']);
    if (L['Event Type']) bits.push('Event: ' + L['Event Type']);
    if (L['Date of Event']) {
      const ed = new Date(L['Date of Event']);
      if (!isNaN(ed.getTime())) bits.push('Event date: ' + Utilities.formatDate(ed, tz_(), 'M/d/yyyy'));
    }
    if (L['Quoted Price']) bits.push('Quoted: ' + L['Quoted Price']);
    if (L['Notes about interaction']) bits.push('Notes: ' + L['Notes about interaction']);

    const ev = cal.createEvent(
      '\u2726 Follow up: ' + name,
      d,
      new Date(d.getTime() + 15 * 60000),
      { description: bits.join('\n') }
    );
    ev.addPopupReminder(0);
  });

  // Task reminders — same rules: today forward, current year, exact time or 9 AM.
  getTasks_().forEach(function (T) {
    if (T.done || !T.due) return;
    const d = new Date(T.due);
    if (isNaN(d.getTime())) return;
    const hasTime = d.getHours() !== 0 || d.getMinutes() !== 0;
    if (!hasTime) d.setHours(REMINDER_HOUR, 0, 0, 0);
    const d0 = new Date(d); d0.setHours(0, 0, 0, 0);
    if (d0 < today || d0 >= yearEnd) return;
    const ev = cal.createEvent(
      '\u2726 Task: ' + T.task,
      d,
      new Date(d.getTime() + 5 * 60000),
      { description: getConfig_().BUSINESS_NAME + ' task' }
    );
    ev.addPopupReminder(0);
  });
}

/**
 * Fires within seconds of the sheet actually changing — including a new
 * row appended via an outside API call like Zapier's "Create Spreadsheet
 * Row" action. Deliberately using onChange rather than onEdit for this:
 * onEdit only fires for edits made directly by a person in the Sheets UI,
 * not for rows written by a script or an external API, which is exactly
 * how new leads arrive from the website. Kept intentionally narrow — just
 * the date fix, not the full automation suite — so it stays cheap enough
 * to run on every change without also touching Calendar every time.
 */
function onSheetChangeNormalizeDates_(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    normalizeEventDateColumn_();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Run this ONCE from the Apps Script editor after pasting the code.
 * Creates the daily 6 AM full sync (data automations + Calendar rebuild),
 * a frequent 5-minute check running only the lightweight lead-cleanup
 * automations (so leads arriving through Zapier or other outside tools get
 * their default follow-up applied quickly, without touching Calendar), an
 * onChange trigger that fixes a plain-text event date within seconds of a
 * new row arriving rather than waiting on the 5-minute cycle, and does a
 * first full sync immediately.
 */
function setup() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === 'syncFollowUps' || fn === 'runFrequentAutomations' || fn === 'onSheetChangeNormalizeDates_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncFollowUps').timeBased().atHour(6).everyDays(1).create();
  ScriptApp.newTrigger('runFrequentAutomations').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('onSheetChangeNormalizeDates_').forSpreadsheet(SpreadsheetApp.getActive()).onChange().create();
  syncFollowUps();
}

