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
  SERVICE_MESSAGES: '',
  PORTAL_THEME: '',
  VENMO_USERNAME: '',
  CASHAPP_CASHTAG: '',
  PAYPAL_USERNAME: '',
  ZELLE_HANDLE: '',
  ACH_BANK: '', ACH_ACCOUNT: '', ACH_ROUTING: '',
  WIRE_BANK: '', WIRE_ACCOUNT: '', WIRE_ROUTING: '', WIRE_BANK_ADDRESS: '',
  CHECK_PAYEE: '', CHECK_ADDRESS: '',
  CASH_INSTRUCTIONS: '',
  DEPOSIT_PERCENT: '50',
  LOGO_URL: '',
  LOGO_BACKING: 'yes',
  GOVERNING_LAW: '',
  CANCELLATION_POLICY: '',
  ADDITIONAL_TERMS: '',
  CONTRACTS_FOLDER: '',
  RECEIPTS_FOLDER: '',
  CAL_NAME: '',
  CONTACT_GROUP: '',
  CALL_LINK: '',
  TEXT_LINK: '',
  SERVICES: 'Strolling\nStage\nStage & Strolling',
  EVENT_TYPES: 'Corporate\nBirthday\nWedding\nBar Mitzvah\nFundraiser\nSchool\nBanquet\nCocktail Party\nPrivate Event\nOther',
  LEAD_SOURCES: 'Bark\nGigsalad\nWebsite\nCurrent Client\nReferral\nOther',
  AD_SOURCES: 'Bark\nGigsalad\nWebsite',
  LOST_REASONS: 'Budget\nNon-responsive\nPostponed\nBooked elsewhere',
  AUDIENCES: '',
  CARD_FEE_ENABLED: 'yes',
  PAYMENT_REPORT_EMAIL: 'yes',
  SIGN_NOTIFY_EMAIL: 'yes',
  STORE_ENABLED: '',
  STORE_PAYMENT_METHODS: '',
  W9_URL: '', W9_FILE_ID: '',
  INSURANCE_URL: '', INSURANCE_FILE_ID: '',
  GOOGLE_REVIEW_URL: ''
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

// The owner's legal business name ONLY if they've actually set one — otherwise ''.
// A blank field keeps the generic placeholder default (getConfig_ skips empty
// values), so that placeholder counts as "not set" and never leaks into a document.
function legalBusinessName_() {
  var v = String(getConfig_().BUSINESS_LEGAL_NAME || '').trim();
  return (v && v !== CONFIG_DEFAULTS_.BUSINESS_LEGAL_NAME) ? v : '';
}

// The company name to print on a document letterhead: the legal business name if
// set, otherwise the owner's display business name (so the header is never blank
// and never shows the placeholder).
function letterheadName_() {
  return legalBusinessName_() || String(getConfig_().BUSINESS_NAME || '').trim();
}

// The credit-card surcharge rate. 3.25% by default; 0 when the owner turns the
// card fee off in Settings (CARD_FEE_ENABLED === 'no'). Enabled unless explicitly
// 'no', so a blank/unset value keeps the fee on. Single source of truth for every
// server-side fee calc (sign-page deposit/balance, receipts, contract, store).
function cardFeeRate_() {
  return String(getConfig_().CARD_FEE_ENABLED || '').trim().toLowerCase() === 'no' ? 0 : 0.0325;
}

/* ---------- Settings (per-install onboarding) ----------
 * CONFIG_FIELDS_ drives the in-app Settings screen (labels + help text) and the
 * "Settings" sheet. getSettings_()/saveSettings_() are called from the UI so a new
 * owner can enter their own info without touching code or the raw sheet.
 */
// Setup screen is grouped into collapsible sections (SETTINGS_SECTIONS_ below);
// each field declares which `section` it belongs to. Order here is the order
// shown within a section.
var CONFIG_FIELDS_ = [
  // Your business
  { key: 'BUSINESS_NAME', label: 'Business name', help: 'Shown throughout the app and to clients.', section: 'business' },
  { key: 'BUSINESS_LEGAL_NAME', label: 'Legal business name', help: 'Full legal entity name used on contracts.', section: 'business' },
  { key: 'OWNER_NAME', label: 'Your name', help: 'Owner / service provider name on contracts and the event portal.', section: 'business' },
  { key: 'TAGLINE', label: 'Tagline', help: 'Short line under your name on the contract.', section: 'business' },
  { key: 'LOGO_URL', label: 'Logo image link', help: 'Paste a direct, public link to your logo image (it should end in .png or .jpg). Easiest way: open the logo on your own website, right-click it, and choose "Copy image address," then paste it here. A Google Drive or Dropbox share link will not work — it has to be a direct image link. A preview appears below once the link is valid.', section: 'business' },
  { key: 'BUSINESS_DOCS', label: 'Tax form & insurance', help: 'Optional. Upload your W-9 and proof of insurance (PDF or image). Clients can view them on their own booking page — handy when a venue or a client\'s accounting team asks for one.', editor: 'businessdocs', section: 'business' },
  // How clients reach you
  { key: 'EMAIL', label: 'Contact email', help: 'Shown to clients on the event portal and contract.', section: 'contact' },
  { key: 'PHONE', label: 'Contact phone', help: 'Shown to clients (optional).', section: 'contact' },
  { key: 'WEBSITE', label: 'Website', help: 'Shown to clients (optional), e.g. yourbusiness.com.', section: 'contact' },
  { key: 'ADDRESS', label: 'Mailing address', help: 'Used on the contract and for check/wire instructions.', section: 'contact' },
  // Getting paid
  { key: 'DEPOSIT_PERCENT', label: 'Deposit', help: 'How much you collect up front, as a percent of the booking total. The rest becomes the balance, due before the event. Set it to 0 to skip deposits entirely — clients then pay the full amount in one payment, and their booking page goes straight to the balance. You can still set a specific deposit on an individual lead; this is only the default.', editor: 'depositpct', section: 'payments' },
  { key: 'CARD_FEE_ENABLED', label: 'Card processing fee', help: 'When on, a 3.25% fee is added to credit-card payments to cover processing costs (deposits, balances, and store card sales). Turn it off to absorb the fee yourself — clients then pay the plain amount by card, with no surcharge. On by default.', editor: 'cardfee', section: 'payments' },
  { key: 'PAYMENT_REPORT_EMAIL', label: 'Payment-report emails', help: 'When on, you get an email whenever a client taps "I’ve sent my payment" so you know to verify and confirm it. Turn it off to rely on the "Payments to confirm" card in the app, which always shows regardless. On by default.', editor: 'toggle', toggleLabel: 'Email me when a client reports a payment', section: 'payments' },
  { key: 'SIGN_NOTIFY_EMAIL', label: 'Contract-signed emails', help: 'When on, you get an email the moment a client signs their contract. On by default.', editor: 'toggle', toggleLabel: 'Email me when a client signs', section: 'payments' },
  { key: 'VENMO_USERNAME', label: 'Venmo username', help: 'Enables the Venmo option. Blank hides it.', section: 'methods' },
  { key: 'CASHAPP_CASHTAG', label: 'Cash App $cashtag', help: 'Enables Cash App. Blank hides it.', section: 'methods' },
  { key: 'PAYPAL_USERNAME', label: 'PayPal.Me username', help: 'Enables PayPal. Your PayPal.Me handle — the part after paypal.me/ (you can set one up free at paypal.me). Blank hides it.', section: 'methods' },
  { key: 'ZELLE_HANDLE', label: 'Zelle email or phone', help: 'Enables Zelle. The email or US mobile number enrolled with Zelle; clients send to it from their own banking app. Blank hides it.', section: 'methods' },
  { key: 'ACH_BANK', label: 'ACH bank name', help: 'Shown with ACH instructions.', section: 'methods' },
  { key: 'ACH_ACCOUNT', label: 'ACH account number', help: 'ACH option appears only if account + routing are set.', section: 'methods' },
  { key: 'ACH_ROUTING', label: 'ACH routing number', help: '', section: 'methods' },
  { key: 'WIRE_BANK', label: 'Wire bank name', help: 'Shown with wire instructions.', section: 'methods' },
  { key: 'WIRE_ACCOUNT', label: 'Wire account number', help: 'Wire option appears only if account + routing are set.', section: 'methods' },
  { key: 'WIRE_ROUTING', label: 'Wire routing number', help: '', section: 'methods' },
  { key: 'WIRE_BANK_ADDRESS', label: 'Wire bank address', help: 'Bank address for wire instructions.', section: 'methods' },
  { key: 'CHECK_PAYEE', label: 'Check payable to', help: 'Enables "Mail a check". Blank uses your legal name.', section: 'methods' },
  { key: 'CHECK_ADDRESS', label: 'Check mailing address', help: 'Where checks are mailed. Blank uses your mailing address.', section: 'methods' },
  { key: 'CASH_INSTRUCTIONS', label: 'Cash instructions', help: 'Enables Cash on your client page. What clients see when they choose it — e.g. "Pay in cash in person before the event." Blank hides Cash from the client page (you can still record a cash payment yourself on a lead).', section: 'methods' },
  // Services & client page
  { key: 'AUDIENCES', label: 'Who you perform for', help: 'Tick the audiences and events you take — this tailors the app to your act. If you do children\'s or family shows, it stops flagging kids leads as "refer out." Leave everything unticked and the app assumes nothing.', editor: 'audiences', section: 'services' },
  { key: 'SERVICES', label: 'Services offered', help: 'One per line — the choices in the Service dropdown.', multiline: true, section: 'services' },
  { key: 'EVENT_TYPES', label: 'Event types', help: 'One per line — the choices in the Event Type dropdown.', multiline: true, section: 'services' },
  { key: 'PORTAL_THEME', label: 'Client portal look', help: 'Optional. Match your brand on the page your clients see — accent color, background, and fonts. Leave everything on Default to keep the classic dark-and-gold look.', editor: 'portaltheme', section: 'services' },
  { key: 'LEAD_SOURCES', label: 'Lead sources', help: 'One per line — how a client first found you.', multiline: true, section: 'services' },
  { key: 'AD_SOURCES', label: 'Advertising sources', help: 'One per line — paid channels tracked in Insights ROI (e.g. Bark, Gigsalad).', multiline: true, section: 'services' },
  { key: 'LOST_REASONS', label: 'Lost reasons', help: 'One per line — your own choices in the "Why was this lead lost?" picker (clients never see these; they group your Insights). An "Other…" free-text option is always there too.', multiline: true, section: 'services' },
  { key: 'SERVICE_MESSAGES', label: 'Event messages by service', help: 'Optional. A message shown to a booked client on their event hub after they sign, matched to their booking\'s service (e.g. one message for a stage show, another for strolling). A box appears for each of your Services above; leave any blank.', editor: 'servicemsgs', section: 'services' },
  { key: 'GOOGLE_REVIEW_URL', label: 'Google review link', help: 'Optional. Paste your Google "write a review" link — the short one that opens the review box directly (it looks like https://g.page/r/…/review). Get it from your Google Business Profile → "Ask for reviews" / "Get more reviews" and copy the link. After a client pays their balance and the event has passed, their portal asks for a rating + comment; a 4- or 5-star rating then offers to share it on Google — and their comment is copied to their clipboard so they just tap the stars and paste. Lower ratings stay private with you. Leave blank to keep ALL feedback private (no Google prompt).', section: 'services' },
  // Contract terms
  { key: 'CANCELLATION_POLICY', label: 'Cancellation policy', help: 'Your cancellation and refund terms, in your own words. This replaces the standard cancellation wording in the Terms section of the contract. Leave blank to keep the standard wording. (It\'s your agreement — review the wording yourself, or with an advisor.)', multiline: true, section: 'contract' },
  { key: 'ADDITIONAL_TERMS', label: 'Additional terms', help: 'Optional extra clauses to add to the Terms section of the contract — one per line (e.g. an outdoor/weather backup requirement, travel, setup space, rescheduling). Each line becomes its own bullet. Leave blank to add none.', multiline: true, section: 'contract' },
  // Store (optional)
  { key: 'STORE_ENABLED', label: 'Store page', help: 'Optional. Turn on a simple public checkout page for selling products — merch, gift cards, downloads — separate from your event bookings. Add products below; each one gets its own shareable link. Off by default.', editor: 'storefront', section: 'storefront' },
  { key: 'STORE_PAYMENT_METHODS', label: 'Store payment methods', help: 'Optional. By default your store offers the same payment methods your clients see on the event portal. Tick specific methods here to show only those in the store — leave them all unticked to offer everything. A method only appears once you\'ve set it up under "Getting paid" above.', editor: 'storemethods', section: 'storefront' },
  // Calls & texts
  { key: 'CALL_LINK', label: 'Custom call link', help: 'Optional — blank uses your phone\'s default dialer. To route through another app, paste its dial link with {number} or {digits} where the number goes — Skype works directly: skype:{number}?call. If the app just opens without a number (e.g. Google Voice: googlevoice://), the client\'s number is copied to your clipboard so you can paste it in. Tip: to use Google Voice for everything, it\'s simplest to set it as your phone\'s default app and leave this blank.', section: 'comms' },
  { key: 'TEXT_LINK', label: 'Custom text link', help: 'Optional — blank uses your phone\'s default messaging. Use {number}/{digits} for the number and {body} for the message where the app\'s link supports them. If the app just opens (e.g. Google Voice: googlevoice://), your message is copied to your clipboard so you can paste it in after you pick the contact.', section: 'comms' },
  // Advanced (optional)
  { key: 'LOGO_BACKING', label: 'Logo backing', help: 'If your logo has a transparent background, Magic Manager puts a solid black square behind it on your contracts and receipts so it doesn’t look odd on white paper. Logos that aren’t transparent are never touched. Turn this off to always show your logo exactly as-is. On by default.', editor: 'toggle', toggleLabel: 'Put a black square behind a transparent logo', section: 'advanced' },
  { key: 'CONTRACTS_FOLDER', label: 'Contracts folder', help: 'Optional. Blank = "[Business name] Contracts".', section: 'advanced' },
  { key: 'RECEIPTS_FOLDER', label: 'Receipts folder', help: 'Optional. Blank = "[Business name] Receipts".', section: 'advanced' },
  { key: 'CAL_NAME', label: 'Follow-ups calendar', help: 'Optional. Blank = "[Business name] Follow-ups".', section: 'advanced' },
  { key: 'CONTACT_GROUP', label: 'Contacts group', help: 'Optional. Blank = "[Business name] Leads".', section: 'advanced' }
];

// Setup screen sections, in display order. `open` = expanded by default (the
// two must-fill sections); the rest start collapsed to keep first-run calm.
var SETTINGS_SECTIONS_ = [
  { id: 'business', title: 'Your business', desc: 'The essentials — these appear in your app, on your contracts, and on the page your clients see.', open: false },
  { id: 'contact', title: 'How clients reach you', desc: 'Your contact details, shown on your contract and the client event portal.', open: false },
  { id: 'payments', title: 'Getting paid', desc: 'Your deposit, the card processing fee, and the emails you get when a client reports a payment or signs their contract.', open: false },
  { id: 'methods', title: 'Payment methods', desc: 'Switch on the payment methods you accept — leave the rest blank. You can add these anytime.', open: false },
  { id: 'services', title: 'Services & client page', desc: 'The dropdown choices inside your app, plus the note clients see after they pay their deposit.', open: false },
  { id: 'contract', title: 'Contract terms', desc: 'Your cancellation policy and any extra clauses — these appear in the Terms section of the agreement your clients sign. Leave blank to use the standard wording.', open: false },
  { id: 'storefront', title: 'Store (optional)', desc: 'Sell products with a simple public checkout page — merch, gift cards, anything that isn\'t an event booking. Off unless you turn it on.', open: false },
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

function getSettings_() {
  return { config: getConfig_(), raw: getRawSettings_(), fields: CONFIG_FIELDS_, sections: SETTINGS_SECTIONS_ };
}

function saveSettings_(values) {
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
  try { PropertiesService.getScriptProperties().deleteProperty('CONTRACT_TEMPLATE_ID_V5'); } catch (e) {}
  getConfig_._cache = null;
  return { ok: true, config: getConfig_() };
}

/* ---------- Client portal look (Settings -> PORTAL_THEME) ----------
 * Optional branding for the client-facing event portal (Sign.html): accent color,
 * background (light / custom color) and a curated font pair. Stored as JSON
 * {accent,bg,bgColor,font}. Blank/absent = the original dark-and-gold look, so
 * new installs are unchanged until the owner chooses otherwise. ONE function
 * computes the CSS variables; the live preview in Settings calls the same code
 * (previewPortalTheme) so the preview can never drift from the real page. */
var PORTAL_FONTS_ = {
  elegant: { label: 'Elegant — Cormorant Garamond + Lato', h: 'Cormorant Garamond', b: 'Lato', hw: '600;700', bw: '400;700' },
  modern:  { label: 'Modern — Poppins + Inter',            h: 'Poppins',            b: 'Inter', hw: '600;700', bw: '400;500;600;700' },
  opensans: { label: 'Clean — Open Sans',                  h: 'Open Sans',          b: 'Open Sans', hw: '600;700', bw: '400;600;700', g: 'sans-serif' },
  classic: { label: 'Classic — Libre Baskerville + Source Sans 3', h: 'Libre Baskerville', b: 'Source Sans 3', hw: '700', bw: '400;600;700' },
  bold:    { label: 'Bold — Oswald + Open Sans',           h: 'Oswald',             b: 'Open Sans', hw: '500;700', bw: '400;600;700' },
  playful: { label: 'Playful — Fredoka + Nunito',          h: 'Fredoka',            b: 'Nunito', hw: '500;600', bw: '400;600;700' }
};
function hexOk_(h) { return /^#[0-9a-fA-F]{6}$/.test(String(h || '')); }
function hexToRgb_(h) { return [parseInt(h.substr(1, 2), 16), parseInt(h.substr(3, 2), 16), parseInt(h.substr(5, 2), 16)]; }
function rgbToHex_(c) { return '#' + c.map(function (v) { v = Math.max(0, Math.min(255, Math.round(v))); return (v < 16 ? '0' : '') + v.toString(16); }).join(''); }
function mixHex_(a, b, t) { var x = hexToRgb_(a), y = hexToRgb_(b); return rgbToHex_([0, 1, 2].map(function (i) { return x[i] + (y[i] - x[i]) * t; })); }
function lumHex_(h) {
  var c = hexToRgb_(h).map(function (v) { v /= 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); });
  return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2];
}
function contrast_(a, b) { var l1 = lumHex_(a), l2 = lumHex_(b); return (Math.max(l1, l2) + 0.05) / (Math.min(l1, l2) + 0.05); }
// Returns null for the default look, else {vars, fonts:{h,b,href}, extra, warn}.
function portalTheme_(themeStr) {
  var t = {};
  try { t = themeStr ? (JSON.parse(themeStr) || {}) : {}; } catch (e) { t = {}; }
  var accent = hexOk_(t.accent) ? t.accent : '';
  var bg = (t.bg === 'light' || t.bg === 'custom') ? t.bg : 'dark';
  var bgColor = bg === 'light' ? '#F7F4EE' : (bg === 'custom' && hexOk_(t.bgColor) ? t.bgColor : '');
  if (bg === 'custom' && !bgColor) bg = 'dark';
  var font = PORTAL_FONTS_[t.font] ? t.font : '';
  if (!accent && bg === 'dark' && !font) return null;

  var vars = {}, extra = '', warn = '';
  var acc = accent || '#C9A050';
  var light = bgColor && lumHex_(bgColor) > 0.4;
  if (accent || bgColor) {
    vars['--gold'] = acc;
    vars['--gold-bright'] = light ? mixHex_(acc, '#000000', 0.25) : mixHex_(acc, '#FFFFFF', 0.28);
    vars['--gold-deep'] = mixHex_(acc, '#000000', 0.45);
    var rgb = hexToRgb_(acc).join(',');
    vars['--line'] = 'rgba(' + rgb + ',' + (light ? '.38' : '.28') + ')';
    vars['--tint'] = 'rgba(' + rgb + ',.14)';
    vars['--on-accent'] = contrast_(acc, '#141008') >= contrast_(acc, '#FFFFFF') ? '#141008' : '#FFFFFF';
  }
  if (bgColor) {
    vars['--ink'] = bgColor;
    if (light) {
      vars['--panel'] = mixHex_(bgColor, '#FFFFFF', 0.65);
      vars['--panel2'] = mixHex_(bgColor, '#000000', 0.05);
      vars['--cream'] = '#1E1B16'; vars['--muted'] = '#6B6558';
      extra += 'select{color-scheme:light}';
    } else {
      vars['--panel'] = mixHex_(bgColor, '#FFFFFF', 0.06);
      vars['--panel2'] = mixHex_(bgColor, '#FFFFFF', 0.11);
      vars['--cream'] = '#F4EBD3'; vars['--muted'] = '#A39C88';
    }
    if (contrast_(vars['--gold-bright'], bgColor) < 3) warn = 'That accent color is hard to read on this background — try a lighter or darker shade.';
  } else if (accent && contrast_(vars['--gold-bright'], '#0B0B10') < 3) {
    warn = 'That accent color is hard to read on a dark background — try a lighter shade.';
  }
  var fonts = null;
  if (font) {
    var F = PORTAL_FONTS_[font];
    var fam;
    if (F.h === F.b) {
      // One typeface for both roles: request it ONCE with the merged weight set
      // (asking Google Fonts for the same family twice is malformed).
      var ws = F.hw.split(';').concat(F.bw.split(';')).filter(function (w, i, a) { return a.indexOf(w) === i; })
        .sort(function (a, b) { return a - b; });
      fam = 'family=' + F.h.replace(/ /g, '+') + ':wght@' + ws.join(';');
    } else {
      fam = 'family=' + F.h.replace(/ /g, '+') + ':wght@' + F.hw
        + '&family=' + F.b.replace(/ /g, '+') + ':wght@' + F.bw;
    }
    // g = generic fallback for headings if the web font can't load (serif unless the pair is all-sans).
    fonts = { h: F.h, b: F.b, g: F.g || 'serif', href: 'https://fonts.googleapis.com/css2?' + fam + '&display=swap' };
  }
  return { vars: vars, fonts: fonts, extra: extra, warn: warn };
}
// <link>+<style> for the head of Sign.html; '' for the default look.
function portalThemeHead_(themeStr) {
  var th = portalTheme_(themeStr);
  if (!th) return '';
  var css = '', k;
  var v = '';
  for (k in th.vars) v += k + ':' + th.vars[k] + ';';
  if (v) css += ':root{' + v + '}';
  var html = '';
  if (th.fonts) {
    html += '<link href="' + th.fonts.href + '" rel="stylesheet">';
    css += "h1,.bookingid .name{font-family:'" + th.fonts.h + "'," + (th.fonts.g || 'serif') + "}"
      + "html,body,input[type=\"text\"],select,button{font-family:'" + th.fonts.b + "',sans-serif}";
  }
  css += th.extra;
  return html + '<style>' + css + '</style>';
}
// Settings live preview: same computation as the real page.
function previewPortalTheme_(themeStr) {
  var th = portalTheme_(themeStr);
  var d = { vars: {}, fonts: null, warn: '' };
  return th ? th : d;
}

/* ---------- Business documents (W-9, proof of insurance) ----------
 * Onboarding upload of the owner's W-9 and proof-of-insurance, stored as real
 * Drive files (far too big for a sheet cell) and shared view-only by link so a
 * client can open them from their booking page. The view URL + file id live in
 * config (W9_URL/W9_FILE_ID, INSURANCE_URL/INSURANCE_FILE_ID); the client sign
 * page reads the URLs straight off the config it already receives. */
var BUSINESS_DOC_KINDS_ = {
  w9:        { urlKey: 'W9_URL',        idKey: 'W9_FILE_ID',        label: 'W-9' },
  insurance: { urlKey: 'INSURANCE_URL', idKey: 'INSURANCE_FILE_ID', label: 'Proof of Insurance' }
};

function businessDocsFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('BUSINESS_DOCS_FOLDER_ID');
  if (savedId) { try { return DriveApp.getFolderById(savedId); } catch (e) {} }
  const name = getConfig_().BUSINESS_NAME + ' Business Documents';
  const it = DriveApp.getFoldersByName(name);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(name);
  props.setProperty('BUSINESS_DOCS_FOLDER_ID', folder.getId());
  return folder;
}

function businessDocsResult_() {
  var cfg = getConfig_();
  return JSON.stringify({ ok: true, W9_URL: cfg.W9_URL || '', INSURANCE_URL: cfg.INSURANCE_URL || '' });
}

// Upload (or replace) the W-9 / proof of insurance. Stores it in Drive, shares it
// view-only by link, and records the URL + file id in config so the client page
// can link to it. Returns the two current doc URLs.
function uploadBusinessDoc_(kind, base64Data, mimeType, fileName) {
  try {
    var spec = BUSINESS_DOC_KINDS_[String(kind)];
    if (!spec) return JSON.stringify({ ok: false, error: 'Unknown document type.' });
    if (!base64Data) return JSON.stringify({ ok: false, error: 'No file was received — please try again.' });
    // Replacing an existing doc: trash the old file so re-uploads don't pile up orphans.
    var oldId = getConfig_()[spec.idKey];
    if (oldId) { try { DriveApp.getFileById(String(oldId)).setTrashed(true); } catch (e) {} }
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType || 'application/octet-stream', fileName || spec.label);
    var file = businessDocsFolder_().createFile(blob);
    shareAnyoneWithLink_(file);
    var vals = {}; vals[spec.urlKey] = file.getUrl(); vals[spec.idKey] = file.getId();
    saveSettings_(vals); // persists to the Settings sheet + clears the config cache
    return businessDocsResult_();
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
}

// Remove a stored W-9 / proof of insurance: trash the Drive file and clear config.
function removeBusinessDoc_(kind) {
  try {
    var spec = BUSINESS_DOC_KINDS_[String(kind)];
    if (!spec) return JSON.stringify({ ok: false, error: 'Unknown document type.' });
    var oldId = getConfig_()[spec.idKey];
    if (oldId) { try { DriveApp.getFileById(String(oldId)).setTrashed(true); } catch (e) {} }
    var vals = {}; vals[spec.urlKey] = ''; vals[spec.idKey] = '';
    saveSettings_(vals);
    return businessDocsResult_();
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
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
var APP_VERSION = '1.5.43';

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
    // A trial can't outlive its end date even inside the daily recheck window —
    // enforce it locally so a cached "active" never runs past expiry.
    if (!(cache.trialEndsAt && now > cache.trialEndsAt)) {
      return { ok: true, status: 'active', message: 'License active.', checkedAt: cache.checkedAt, trialEndsAt: cache.trialEndsAt || 0 };
    }
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
        if (data.storeUrl != null) props.setProperty('STORE_URL', String(data.storeUrl));
        // Remember a trial's end date for the in-app warnings + email reminders;
        // a full (non-trial) active license clears it so a converted buyer sees no trial UI.
        if (data.trialEndsAt) props.setProperty('TRIAL_ENDS_AT', String(data.trialEndsAt));
        else if (data.status === 'active') props.deleteProperty('TRIAL_ENDS_AT');
      } catch (e) {}
    }
    if (data && (data.status === 'active' || data.status === 'invalid' || data.status === 'revoked' || data.status === 'expired')) {
      var trialEndsAt = Number(data.trialEndsAt) || 0;
      props.setProperty('LICENSE_CACHE', JSON.stringify({ status: data.status, checkedAt: now, key: key, trialEndsAt: trialEndsAt }));
      if (data.status === 'active') return { ok: true, status: 'active', message: 'License active.', checkedAt: now, trialEndsAt: trialEndsAt };
      return { ok: false, status: data.status, message: data.message || ('This license is ' + data.status + '.'), storeUrl: String(data.storeUrl || props.getProperty('STORE_URL') || '') };
    }
  }

  // Hub unreachable or unexpected response — fail gracefully within the grace window.
  if (cache && cache.key === key && cache.status === 'active' && (now - cache.checkedAt) < LICENSE_GRACE_MS) {
    return { ok: true, status: 'grace', message: 'Could not reach the license server — using your last verified status.', checkedAt: cache.checkedAt };
  }
  return { ok: false, status: 'error', message: 'Could not verify your license. Check your connection and try again.' };
}

function activateLicense_(key) {
  var props = PropertiesService.getScriptProperties();
  if (typeof key === 'string') props.setProperty('LICENSE_KEY', key.trim());
  try { props.deleteProperty('LICENSE_CACHE'); } catch (e) {}
  return getLicenseState_();
}

// Self-service free trial. Asks the hub for a trial tied to this email (the hub
// gives one per email), then activates the returned key. Called from the
// activation page's "Start free trial" form.
function startTrial_(email, name) {
  var props = PropertiesService.getScriptProperties();
  var hub = licenseHubUrl_();
  if (!hub) return { ok: false, message: 'Free trials aren’t available for this app.' };
  // Never let starting a trial clobber an already-active PAID license. The trial
  // option isn't shown once the app is activated, but this guarantees a paid key
  // can't be overwritten even if the trial call were somehow reached.
  var cur = getLicenseState_();
  if (cur && cur.ok && cur.status === 'active' && !cur.trialEndsAt) {
    return { ok: false, message: 'You already have an active license — no trial needed.' };
  }
  email = String(email || '').trim();
  if (!email || email.indexOf('@') < 1 || email.indexOf('.') < 0) {
    return { ok: false, message: 'Please enter a valid email address.' };
  }
  name = String(name || '').trim();
  if (!name) {
    return { ok: false, message: 'Please enter your name.' };
  }
  var resp = null;
  try {
    resp = UrlFetchApp.fetch(hub, {
      method: 'post',
      payload: { action: 'trial', email: email, name: name, install: getInstallId_() },
      muteHttpExceptions: true,
      followRedirects: true
    });
  } catch (e) { resp = null; }
  if (!resp) return { ok: false, message: 'Could not reach the trial server — check your connection and try again.' };
  var data = null;
  try { data = JSON.parse(resp.getContentText()); } catch (e) {}
  if (!data) return { ok: false, message: 'Unexpected response from the server — please try again.' };
  if (!data.ok) {
    if (data.storeUrl) { try { props.setProperty('STORE_URL', String(data.storeUrl)); } catch (e) {} }
    return { ok: false, reason: data.reason || '', message: data.message || 'Could not start your trial.', storeUrl: String(data.storeUrl || '') };
  }
  // Success — store the key + trial state, then activate.
  props.setProperty('LICENSE_KEY', String(data.key));
  props.setProperty('TRIAL_EMAIL', email);
  props.setProperty('TRIAL_NAME', name);
  if (data.trialEndsAt != null) props.setProperty('TRIAL_ENDS_AT', String(data.trialEndsAt));
  if (data.storeUrl != null) props.setProperty('STORE_URL', String(data.storeUrl));
  try { props.deleteProperty('LICENSE_CACHE'); } catch (e) {}
  // Fresh trial — clear any prior warning/reminder flags.
  try {
    props.deleteProperty('TRIAL_WARN_3'); props.deleteProperty('TRIAL_WARN_1');
    props.deleteProperty('TRIAL_MAIL_3'); props.deleteProperty('TRIAL_MAIL_1');
  } catch (e) {}
  var state = getLicenseState_();
  state.startedTrial = true;
  return state;
}

// What the in-app trial banner needs: whether this is a trial, days remaining,
// the purchase link, and which warnings have already been shown (so the 3-day and
// 1-day notices each appear once). Non-trial installs get {trial:false}.
function getTrialInfo_() {
  var props = PropertiesService.getScriptProperties();
  var ends = Number(props.getProperty('TRIAL_ENDS_AT')) || 0;
  if (!ends) return { trial: false };
  var msLeft = ends - Date.now();
  return {
    trial: true,
    endsAt: ends,
    daysLeft: Math.ceil(msLeft / (24 * 60 * 60 * 1000)),
    expired: msLeft <= 0,
    storeUrl: (props.getProperty('STORE_URL') || '').trim(),
    warn3Shown: props.getProperty('TRIAL_WARN_3') === '1',
    warn1Shown: props.getProperty('TRIAL_WARN_1') === '1'
  };
}

// Remember that a trial warning (the 3-day or 1-day one) has been shown, so it
// doesn't reappear on every load.
function markTrialWarned_(which) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(String(which) === '1' ? 'TRIAL_WARN_1' : 'TRIAL_WARN_3', '1');
  return JSON.stringify({ ok: true });
}

// Trial reminder EMAILS are sent by the license hub (from the brand address), not
// here — so they come from Magic Manager rather than the buyer's own account. The
// in-app 3-day / 1-day banner (getTrialInfo / markTrialWarned) still lives in the app.

// Compares APP_VERSION against the latest version the hub reported (captured on
// the daily license check) and returns what the "update available" banner needs.
function getUpdateInfo_() {
  var props = PropertiesService.getScriptProperties();
  var latest = (props.getProperty('LATEST_VERSION') || '').trim();
  return {
    current: APP_VERSION,
    latest: latest,
    notes: (props.getProperty('LATEST_NOTES') || '').trim(),
    url: (props.getProperty('UPDATE_URL') || '').trim(),
    updateAvailable: !!latest && versionLt_(APP_VERSION, latest),
    // Banner-snooze state (remembered server-side so it holds across devices).
    snoozedUntil: Number(props.getProperty('UPDATE_SNOOZE_UNTIL')) || 0,
    dismissedVersion: (props.getProperty('UPDATE_DISMISSED_VERSION') || '').trim()
  };
}

// The update banner is snoozable so it doesn't nag on every open. snoozeUpdate
// hides it for a number of days (stored as a timestamp in Script Properties, so
// the choice is remembered across the user's devices for this install).
// dismissUpdateVersion hides it until a version newer than the skipped one ships.
function snoozeUpdate_(days) {
  var d = Number(days) || 0;
  if (d < 1) d = 7;
  var until = Date.now() + d * 24 * 60 * 60 * 1000;
  PropertiesService.getScriptProperties().setProperty('UPDATE_SNOOZE_UNTIL', String(until));
  return JSON.stringify({ ok: true, snoozedUntil: until });
}
function dismissUpdateVersion_() {
  var props = PropertiesService.getScriptProperties();
  var latest = (props.getProperty('LATEST_VERSION') || '').trim();
  if (latest) props.setProperty('UPDATE_DISMISSED_VERSION', latest);
  return JSON.stringify({ ok: true, dismissedVersion: latest });
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
  var info = getUpdateInfo_();
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

function activationPageHtml_(lic, appKey) {
  var biz = getConfig_().BUSINESS_NAME || 'this app';
  var status = (lic && lic.status) || '';
  var expired = (status === 'expired');
  var showTrial = (status === 'unset'); // offer a trial only to someone who hasn't had one
  var msg = (lic && lic.message) ? lic.message : 'Enter your license key to activate this app.';
  var store = '';
  try { store = (lic && lic.storeUrl) ? lic.storeUrl : (PropertiesService.getScriptProperties().getProperty('STORE_URL') || ''); } catch (e) {}
  var storeOk = /^https?:\/\//i.test(store);

  var p = [];
  p.push('<!doctype html><html><head><meta charset="utf-8">');
  p.push('<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">');
  p.push('<title>', escHtml_(biz), ' — Activate</title>');
  p.push('<style>');
  p.push('body{margin:0;background:#0B0B10;color:#F4EBD3;font-family:Arial,Helvetica,sans-serif;display:flex;min-height:100vh;align-items:center;justify-content:center;padding:24px}');
  p.push('.box{max-width:400px;width:100%;text-align:center}');
  p.push('h1{color:#E4C179;font-weight:700;font-size:20px;margin:0 0 6px}');
  p.push('p{color:#928A75;font-size:13.5px;line-height:1.6;margin:0 0 16px}');
  p.push('input{width:100%;box-sizing:border-box;font-size:16px;color:#F4EBD3;background:#1C1C27;border:1px solid rgba(201,160,80,.4);border-radius:8px;padding:12px;outline:none;text-align:center}');
  p.push('input:focus{border-color:#C9A050}');
  p.push('button{width:100%;margin-top:12px;background:#C9A050;color:#141008;font-weight:700;font-size:15px;border:none;border-radius:10px;padding:13px;cursor:pointer}');
  p.push('button:disabled{opacity:.5}button.ghost{background:transparent;border:1px solid rgba(201,160,80,.45);color:#E4C179}');
  p.push('a.buy{display:block;text-decoration:none;margin-top:6px;background:#C9A050;color:#141008;font-weight:700;font-size:15px;border-radius:10px;padding:13px}');
  p.push('.divider{display:flex;align-items:center;gap:10px;color:#6b6250;font-size:11px;text-transform:uppercase;letter-spacing:.1em;margin:22px 0 14px}');
  p.push('.divider::before,.divider::after{content:"";flex:1;height:1px;background:rgba(201,160,80,.18)}');
  p.push('.lbl{color:#E4C179;font-size:13px;font-weight:700;margin:0 0 8px}');
  p.push('.err{color:#E9927B;font-size:12.5px;margin-top:10px;min-height:16px}');
  p.push('</style></head><body><div class="box">');

  if (expired) {
    p.push('<h1>Your free trial has ended</h1>');
    p.push('<p>Your data is safe and untouched. Get the full version to pick up right where you left off — nothing is deleted.</p>');
    if (storeOk) p.push('<a class="buy" href="', escHtml_(store), '" target="_top" rel="noopener">Get the full version</a>');
    p.push('<div class="divider">already purchased?</div>');
    p.push('<div class="lbl">Enter your license key</div>');
  } else {
    p.push('<h1>Activate ', escHtml_(biz), '</h1>');
    p.push('<p>', escHtml_(msg), '</p>');
  }

  p.push('<input id="k" placeholder="License key" autocomplete="off" autocapitalize="off" spellcheck="false">');
  p.push('<button id="b" onclick="act()">Activate</button>');
  p.push('<div class="err" id="e"></div>');

  if (showTrial) {
    p.push('<div class="divider">or try it free</div>');
    p.push('<div class="lbl">Start your 14-day free trial</div>');
    p.push('<input id="tn" placeholder="Your name" autocomplete="name" autocapitalize="words" spellcheck="false">');
    p.push('<input id="te" type="email" placeholder="you@email.com" autocomplete="email" autocapitalize="off" spellcheck="false">');
    p.push('<button id="tb" class="ghost" onclick="startTrialUI()">Start free trial</button>');
    p.push('<div class="err" id="tee"></div>');
    p.push('<p style="margin-top:10px;font-size:12px">No credit card. Full access for 14 days.</p>');
  }

  p.push('<script>');
  p.push('var K=' + JSON.stringify(String(appKey || '')).replace(/</g, '\\u003c') + ';');
  p.push('function act(){var k=document.getElementById("k").value.trim();if(!k)return;');
  p.push('var b=document.getElementById("b");b.disabled=true;b.textContent="Checking…";document.getElementById("e").textContent="";');
  p.push('google.script.run.withSuccessHandler(function(r){if(r&&r.ok){location.reload();}else{b.disabled=false;b.textContent="Activate";document.getElementById("e").textContent=(r&&r.message)||"That key could not be verified.";}})');
  p.push('.withFailureHandler(function(){b.disabled=false;b.textContent="Activate";document.getElementById("e").textContent="Something went wrong — try again.";})');
  p.push('.api(K,"activateLicense",[k]);}');
  p.push('document.getElementById("k").addEventListener("keydown",function(e){if(e.key==="Enter")act();});');
  if (showTrial) {
    p.push('function startTrialUI(){var nm=document.getElementById("tn").value.trim();var em=document.getElementById("te").value.trim();var tb=document.getElementById("tb");var ee=document.getElementById("tee");ee.textContent="";if(!nm){ee.textContent="Enter your name.";return;}if(!em){ee.textContent="Enter your email.";return;}');
    p.push('tb.disabled=true;tb.textContent="Starting…";');
    p.push('google.script.run.withSuccessHandler(function(r){if(r&&r.ok){location.reload();}else{tb.disabled=false;tb.textContent="Start free trial";ee.textContent=(r&&r.message)||"Could not start your trial.";}})');
    p.push('.withFailureHandler(function(){tb.disabled=false;tb.textContent="Start free trial";ee.textContent="Something went wrong — try again.";})');
    p.push('.api(K,"startTrial",[em,nm]);}');
    p.push('document.getElementById("tn").addEventListener("keydown",function(e){if(e.key==="Enter")startTrialUI();});');
    p.push('document.getElementById("te").addEventListener("keydown",function(e){if(e.key==="Enter")startTrialUI();});');
  }
  p.push('</script>');
  p.push('</div></body></html>');
  return p.join('');
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
  var res = activateLicense_(r.getResponseText());
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
    // A card checkout returns here with &paid=1&pk=deposit|balance&cs=<session>.
    // Re-fetch the session straight from Stripe (with the secret key) to verify
    // it's really paid — if so, auto-mark it RECEIVED (card is the one method we
    // can truly verify). If we can't verify (no session id / not paid / API
    // hiccup), fall back to the same "reported → owner confirms" path as the
    // pay-outside methods, so nothing is ever marked received unverified.
    if (e.parameter.paid && e.parameter.pk) {
      try {
        var _info = e.parameter.cs ? stripeSessionInfo_(e.parameter.cs) : null;
        var _expected = expectedCardCents_(e.parameter.sign, e.parameter.pk);
        // Auto-confirm only when Stripe reports it PAID and collected the EXPECTED
        // amount. Anything else (unverifiable, not paid, or an amount mismatch)
        // falls back to "reported" so you verify and confirm it yourself.
        if (_info && _info.paid && _expected > 0 && _info.amount === _expected) {
          autoConfirmCardPayment_(e.parameter.sign, e.parameter.pk);
        } else {
          reportPayment(e.parameter.sign, e.parameter.pk);
        }
      } catch (e2) {}
    }
    // Shallow-copy the config so we can add a derived CARD_ENABLED flag without
    // polluting the cached getConfig_ object. The portal hides Credit Card unless
    // Stripe is actually set up (the secret key lives in Script Properties, not
    // in Settings, so the portal can't tell otherwise).
    var signCfg = getConfig_(), cfgOut = {};
    for (var ck in signCfg) cfgOut[ck] = signCfg[ck];
    cfgOut.CARD_ENABLED = !!String(PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY') || '').trim();
    template.cfg = cfgOut;
    template.themeHtml = portalThemeHead_(signCfg.PORTAL_THEME);
    return template.evaluate()
      .setTitle(getConfig_().BUSINESS_NAME + ' — Event Portal')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ?store=<productId> serves the optional storefront product page (its own
  // public link, separate from the app and the sign page). The page itself
  // checks the store toggle + product status via getStoreProduct.
  if (e && e.parameter && e.parameter.store) {
    var st = HtmlService.createTemplateFromFile('Store');
    st.productId = String(e.parameter.store);
    st.paid = e.parameter.paid || '';
    // Same as the sign page: only offer Credit Card when Stripe is set up.
    var stCfg = getConfig_(), stCfgOut = {};
    for (var sk in stCfg) stCfgOut[sk] = stCfg[sk];
    stCfgOut.CARD_ENABLED = !!String(PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY') || '').trim();
    st.cfg = stCfgOut;
    return st.evaluate()
      .setTitle(getConfig_().BUSINESS_NAME + ' — Store')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ?rc=<inbox code>&ref=<lead> — a lead referred by another Magic Manager user. Public by
  // design (the sender has no access to this app), so it is gated by the inbox code, only
  // ever writes to the Referral Inbox, and never exposes the app or its access key.
  if (e && e.parameter && e.parameter.ref) return referralReceivePage_(e.parameter);

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
    return HtmlService.createHtmlOutput(activationPageHtml_(lic, providedKey))
      .setTitle(getConfig_().BUSINESS_NAME + ' — Activate')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var indexTpl = HtmlService.createTemplateFromFile('Index');
  // Code.gs and Index.html are updated together by pasting files. If Index.html is still the
  // old one (it calls server functions that are now private), say so plainly instead of
  // showing an app that silently can't load anything.
  if (indexTpl.getRawContent().indexOf('var PAGE_API = ' + PAGE_API_) < 0) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;max-width:480px;margin:0 auto;padding:60px 20px;text-align:center;color:#ddd;background:#0B0B10;min-height:100vh">' +
      '<h2 style="color:#C9A050">Update not finished</h2>' +
      '<p>The new <b>Code.gs</b> is in place, but <b>Index.html</b> is still the old version. In the Apps Script editor, replace <b>Index.html</b> with the new file (and Sign.html, Store.html and appsscript.json if you haven\u2019t), save, then choose Deploy \u2192 Manage deployments \u2192 New version.</p>' +
      '<p style="color:#888;font-size:13px">Your data is safe and untouched.</p></div>'
    ).setTitle('Update not finished').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  indexTpl.cfg = getConfig_();
  indexTpl.appKey = providedKey;
  indexTpl.serverApi = PAGE_API_;
  return indexTpl.evaluate()
    .setTitle(getConfig_().BUSINESS_NAME + ' — Leads')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- Private API gate ----------
 * The web app is deployed to run as the owner and open to "Anyone" (so clients can
 * reach the sign page), which means EVERY top-level function whose name does not end in _
 * can be called from the browser by anyone who loads any public page. The ?key= check in
 * doGet only guards the page load, not those calls. So the owner-only functions are all
 * named with a trailing _ (Apps Script hides those from the browser) and the main app reaches
 * them ONLY through api(), which checks the private access key first.
 * To add a new owner-only endpoint: define it as foo_() and add foo: foo_ to API_ below.
 * The few functions clients and the store page must call directly (sign, pay, store checkout)
 * stay public on purpose. */
// Bump when the page<->server call contract changes; Index.html carries the same number.
var PAGE_API_ = 2;

var API_ = {
  getSettings: getSettings_, saveSettings: saveSettings_, previewPortalTheme: previewPortalTheme_,
  uploadBusinessDoc: uploadBusinessDoc_, removeBusinessDoc: removeBusinessDoc_, activateLicense: activateLicense_,
  startTrial: startTrial_, getTrialInfo: getTrialInfo_, markTrialWarned: markTrialWarned_, getUpdateInfo: getUpdateInfo_,
  snoozeUpdate: snoozeUpdate_, dismissUpdateVersion: dismissUpdateVersion_, generateContract: generateContract_,
  confirmReportedPayment: confirmReportedPayment_, dismissReportedPayment: dismissReportedPayment_,
  recordManualStoreSale: recordManualStoreSale_, setStoreSaleStatus: setStoreSaleStatus_, deleteStoreSale: deleteStoreSale_,
  updateStoreSale: updateStoreSale_, storeListProducts: storeListProducts_, storeSaveProduct: storeSaveProduct_,
  storeDeleteProduct: storeDeleteProduct_, addExpense: addExpense_, addRecurringExpense: addRecurringExpense_,
  stopRecurringExpense: stopRecurringExpense_, editRecurringExpense: editRecurringExpense_, importExpensesBatch: importExpensesBatch_,
  bulkSetAdSource: bulkSetAdSource_, updateExpense: updateExpense_, deleteExpense: deleteExpense_,
  uploadExpenseReceipt: uploadExpenseReceipt_, deleteExpenseReceipt: deleteExpenseReceipt_, deleteLog: deleteLog_,
  updateLog: updateLog_, logContact: logContact_, addStarterTemplates: addStarterTemplates_, moveTemplate: moveTemplate_,
  addTemplate: addTemplate_, updateTemplate: updateTemplate_, deleteTemplate: deleteTemplate_,
  addPartnerTemplate: addPartnerTemplate_, updatePartnerTemplate: updatePartnerTemplate_, deletePartnerTemplate: deletePartnerTemplate_,
  moveTextTemplate: moveTextTemplate_, addTextTemplate: addTextTemplate_, updateTextTemplate: updateTextTemplate_,
  deleteTextTemplate: deleteTextTemplate_, addPartnerTextTemplate: addPartnerTextTemplate_, updatePartnerTextTemplate: updatePartnerTextTemplate_,
  deletePartnerTextTemplate: deletePartnerTextTemplate_, deleteLead: deleteLead_, getTrash: getTrash_,
  restoreLead: restoreLead_, permanentlyDeleteTrash: permanentlyDeleteTrash_, permanentlyDeleteTrashBatch: permanentlyDeleteTrashBatch_,
  startContactSync: startContactSync_, getContactSyncStatus: getContactSyncStatus_, syncContactsToGoogle: syncContactsToGoogle_,
  setTaxRate: setTaxRate_, getLeads: getLeads_, addExpenseCategory: addExpenseCategory_, deleteExpenseCategory: deleteExpenseCategory_,
  saveFieldSettings: saveFieldSettings_, markOracleTipShown: markOracleTipShown_, logPartnerContact: logPartnerContact_,
  updatePartnerLog: updatePartnerLog_, deletePartnerLog: deletePartnerLog_, addPartner: addPartner_,
  updatePartner: updatePartner_, deletePartner: deletePartner_, getReferralInboxLink: getReferralInboxLink_,
  resetReferralInboxCode: resetReferralInboxCode_, acceptReferral: acceptReferral_, dismissReferral: dismissReferral_,
  addTask: addTask_, toggleTask: toggleTask_, deleteTask: deleteTask_, updateTask: updateTask_,
  duplicateLeadAsNewBooking: duplicateLeadAsNewBooking_, updateLead: updateLead_, clearStaleFollowups: clearStaleFollowups_,
  addLead: addLead_, importLeadsBatch: importLeadsBatch_, markFollowupStageDone: markFollowupStageDone_,
};

function api(key, name, args) {
  var k = String(PropertiesService.getScriptProperties().getProperty('APP_ACCESS_KEY') || '');
  if (!k || String(key || '') !== k) throw new Error('Not authorized');
  if (!Object.prototype.hasOwnProperty.call(API_, name)) throw new Error('Unknown method: ' + name);
  return API_[name].apply(null, args || []);
}

/* Editor-only guard for the manual maintenance scripts, which have to keep public names (a
 * name ending in _ is hidden from the editor's Run menu). Run from the editor, the active
 * user is the owner; an anonymous web caller has no email. */
function ownerOnly_() {
  var a = Session.getActiveUser().getEmail();
  if (!a || a !== Session.getEffectiveUser().getEmail()) throw new Error('Owner only');
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
  // Deposit = the configured percent of the total (default 50%), unless a specific
  // Deposit Amount is set on the lead (that always wins). A percent of 0 means no
  // deposit: the balance becomes the whole total and the client's booking page
  // skips the deposit step. Kept as a config default so each business sets its
  // own deposit policy — or waives deposits — during onboarding.
  // Empty / non-numeric -> fall back to 50% (a garbage value must never silently
  // waive deposits). An explicit "0" is a real value and still means "no deposit".
  var rawPct = String(getConfig_().DEPOSIT_PERCENT).replace(/[^0-9.]/g, '');
  var pct = (rawPct === '') ? 50 : Number(rawPct);
  if (isNaN(pct)) pct = 50;
  pct = Math.max(0, Math.min(100, pct));
  const deposit = depositRaw || (Math.round(totalDue * pct) / 100);
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
  const savedId = props.getProperty('CONTRACT_TEMPLATE_ID_V5');
  if (savedId) {
    try { return DocumentApp.openById(savedId); } catch (e) {}
  }
  const doc = buildContractTemplate_();
  props.setProperty('CONTRACT_TEMPLATE_ID_V5', doc.getId());
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

/* ---------- Black backing for transparent logos ----------
 * A logo with a see-through background can look wrong on a white contract/receipt (a light logo
 * vanishes, and some PDF renderers fill transparent areas black on their own). When — and ONLY
 * when — the logo genuinely has transparency, its header cell gets a solid black backing so it
 * reads as one deliberate square. Opaque logos are left exactly as they were.
 * A header flag isn't enough to decide: many PNGs carry an alpha channel that is fully opaque, so
 * the PNG is actually decoded and its pixels counted (bounded to the top rows so a huge image
 * can't slow contract generation). The verdict is cached per logo, so this runs once per logo.
 * Anything we can't analyze falls back to "not transparent" — i.e. today's behavior. */
var LOGO_TRANSPARENT_MIN_FRACTION_ = 0.01; // >= 1% of pixels not fully opaque => a transparent logo
var LOGO_SCAN_MAX_ROWS_ = 400;             // decode at most this many rows of the image

var INFLATE_LBASE_ = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
var INFLATE_LEXT_ = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
var INFLATE_DBASE_ = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
var INFLATE_DEXT_ = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
var INFLATE_ORDER_ = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

function inflateHuff_(lengths, n) {
  var count = [], i;
  for (i = 0; i <= 15; i++) count[i] = 0;
  for (i = 0; i < n; i++) count[lengths[i]]++;
  var offs = [0, 0];
  for (i = 1; i < 15; i++) offs[i + 1] = offs[i] + count[i];
  var symbol = [];
  for (i = 0; i < n; i++) if (lengths[i]) symbol[offs[lengths[i]]++] = i;
  return { count: count, symbol: symbol };
}
// Raw DEFLATE (RFC 1951) decoder. Stops once maxOut bytes are produced. Returns a Uint8Array.
function inflateRaw_(src, start, maxOut) {
  var out = new Uint8Array(maxOut), op = 0, pos = start, bitBuf = 0, bitCnt = 0, i;
  function bits(n) {
    while (bitCnt < n) {
      if (pos >= src.length) throw new Error('eof');
      bitBuf |= src[pos++] << bitCnt; bitCnt += 8;
    }
    var v = bitBuf & ((1 << n) - 1);
    bitBuf >>>= n; bitCnt -= n;
    return v;
  }
  function decode(h) {
    var code = 0, first = 0, index = 0;
    for (var len = 1; len <= 15; len++) {
      code |= bits(1);
      var c = h.count[len];
      if (code - c < first) return h.symbol[index + (code - first)];
      index += c; first += c; first <<= 1; code <<= 1;
    }
    throw new Error('bad code');
  }
  var fixedLit = null, fixedDist = null;
  var last;
  do {
    last = bits(1);
    var type = bits(2);
    if (type === 0) {
      bitBuf = 0; bitCnt = 0; // discard the partial byte
      if (pos + 4 > src.length) throw new Error('eof');
      var slen = src[pos] | (src[pos + 1] << 8); pos += 4;
      for (i = 0; i < slen; i++) {
        if (op >= maxOut) return out;
        if (pos >= src.length) throw new Error('eof');
        out[op++] = src[pos++];
      }
    } else if (type === 1 || type === 2) {
      var lit, dist;
      if (type === 1) {
        if (!fixedLit) {
          var fl = [];
          for (i = 0; i < 144; i++) fl[i] = 8;
          for (; i < 256; i++) fl[i] = 9;
          for (; i < 280; i++) fl[i] = 7;
          for (; i < 288; i++) fl[i] = 8;
          fixedLit = inflateHuff_(fl, 288);
          var fd = []; for (i = 0; i < 30; i++) fd[i] = 5;
          fixedDist = inflateHuff_(fd, 30);
        }
        lit = fixedLit; dist = fixedDist;
      } else {
        var nlen = bits(5) + 257, ndist = bits(5) + 1, ncode = bits(4) + 4;
        var cl = []; for (i = 0; i < 19; i++) cl[i] = 0;
        for (i = 0; i < ncode; i++) cl[INFLATE_ORDER_[i]] = bits(3);
        var lencode = inflateHuff_(cl, 19);
        var all = [], idx = 0, total = nlen + ndist;
        while (idx < total) {
          var s = decode(lencode);
          if (s < 16) { all[idx++] = s; }
          else {
            var prev = 0, rep;
            if (s === 16) { if (idx === 0) throw new Error('bad repeat'); prev = all[idx - 1]; rep = 3 + bits(2); }
            else if (s === 17) { rep = 3 + bits(3); }
            else { rep = 11 + bits(7); }
            if (idx + rep > total) throw new Error('bad lengths');
            while (rep--) all[idx++] = prev;
          }
        }
        lit = inflateHuff_(all.slice(0, nlen), nlen);
        dist = inflateHuff_(all.slice(nlen), ndist);
      }
      for (;;) {
        var sym = decode(lit);
        if (sym < 256) {
          if (op >= maxOut) return out;
          out[op++] = sym;
        } else if (sym === 256) {
          break;
        } else {
          sym -= 257;
          if (sym >= 29) throw new Error('bad length symbol');
          var mlen = INFLATE_LBASE_[sym] + bits(INFLATE_LEXT_[sym]);
          var ds = decode(dist);
          if (ds >= 30) throw new Error('bad distance symbol');
          var d = INFLATE_DBASE_[ds] + bits(INFLATE_DEXT_[ds]);
          if (d > op) throw new Error('bad distance');
          while (mlen--) {
            if (op >= maxOut) return out;
            out[op] = out[op - d]; op++;
          }
        }
      }
    } else {
      throw new Error('bad block type');
    }
  } while (!last);
  return out.subarray(0, op);
}
function pngPaeth_(a, b, c) {
  var p = a + b - c, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c);
  return (pa <= pb && pa <= pc) ? a : (pb <= pc ? b : c);
}
// Fraction (0..1) of the PNG's pixels that are not fully opaque, or null if it can't be analyzed.
function pngTransparentFraction_(u) {
  try {
    if (u.length < 33 || u[0] !== 137 || u[1] !== 80 || u[2] !== 78 || u[3] !== 71) return null;
    var u32 = function (p) { return ((u[p] << 24) | (u[p + 1] << 16) | (u[p + 2] << 8) | u[p + 3]) >>> 0; };
    var pos = 8, w = 0, h = 0, depth = 0, ctype = -1, interlace = 0, trns = null, chunks = [], zlen = 0;
    while (pos + 8 <= u.length) {
      var len = u32(pos), type = String.fromCharCode(u[pos + 4], u[pos + 5], u[pos + 6], u[pos + 7]), d = pos + 8;
      if (d + len > u.length) break;
      if (type === 'IHDR') { w = u32(d); h = u32(d + 4); depth = u[d + 8]; ctype = u[d + 9]; interlace = u[d + 12]; }
      else if (type === 'tRNS') { trns = u.subarray(d, d + len); }
      else if (type === 'IDAT') { chunks.push(u.subarray(d, d + len)); zlen += len; }
      else if (type === 'IEND') { break; }
      pos = d + len + 4;
    }
    var channels = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[ctype];
    if (!channels || !w || !h || interlace !== 0) return null;
    // Only these can hold transparency; everything else is opaque without decoding a byte.
    var canHaveAlpha = (ctype === 4 || ctype === 6) || (ctype === 3 && trns) || ((ctype === 0 || ctype === 2) && trns && depth === 8);
    if (!canHaveAlpha) return 0;
    if ((ctype === 4 || ctype === 6) && depth !== 8 && depth !== 16) return null;
    if (!zlen) return null;
    var z = new Uint8Array(zlen), zo = 0;
    for (var ci = 0; ci < chunks.length; ci++) { z.set(chunks[ci], zo); zo += chunks[ci].length; }
    if ((z[0] & 15) !== 8 || (z[1] & 32)) return null; // not plain zlib deflate (or preset dictionary)
    var bitsPP = channels * depth, rowBytes = Math.ceil(w * bitsPP / 8), bpp = Math.max(1, bitsPP >> 3);
    var rows = Math.min(h, LOGO_SCAN_MAX_ROWS_);
    var raw = inflateRaw_(z, 2, (rowBytes + 1) * rows);
    var prev = new Uint8Array(rowBytes), cur = new Uint8Array(rowBytes), total = 0, trans = 0;
    for (var y = 0; y < rows; y++) {
      var off = y * (rowBytes + 1);
      if (off + rowBytes + 1 > raw.length) break; // ran out of data: judge on the complete rows we have
      var ft = raw[off], x, v, a, b, c, r;
      for (x = 0; x < rowBytes; x++) {
        v = raw[off + 1 + x];
        a = x >= bpp ? cur[x - bpp] : 0; b = prev[x]; c = x >= bpp ? prev[x - bpp] : 0;
        if (ft === 0) r = v; else if (ft === 1) r = v + a; else if (ft === 2) r = v + b;
        else if (ft === 3) r = v + ((a + b) >> 1); else if (ft === 4) r = v + pngPaeth_(a, b, c); else return null;
        cur[x] = r & 255;
      }
      for (var px = 0; px < w; px++) {
        var alpha = 255;
        if (ctype === 6) alpha = depth === 8 ? cur[px * 4 + 3] : cur[px * 8 + 6];
        else if (ctype === 4) alpha = depth === 8 ? cur[px * 2 + 1] : cur[px * 4 + 2];
        else if (ctype === 3) {
          var pi;
          if (depth === 8) pi = cur[px];
          else { var bit = px * depth; pi = (cur[bit >> 3] >> (8 - depth - (bit & 7))) & ((1 << depth) - 1); }
          alpha = (trns && pi < trns.length) ? trns[pi] : 255;
        } else if (ctype === 0) { if (cur[px] === trns[1]) alpha = 0; }
        else if (ctype === 2) { if (cur[px * 3] === trns[1] && cur[px * 3 + 1] === trns[3] && cur[px * 3 + 2] === trns[5]) alpha = 0; }
        total++; if (alpha < 250) trans++;
      }
      var tmp = prev; prev = cur; cur = tmp;
    }
    return total ? trans / total : null;
  } catch (e) { return null; }
}
function gifHasTransparency_(u) {
  var p = 13;
  if (u[10] & 0x80) p += 3 * (1 << ((u[10] & 7) + 1));
  while (p < u.length) {
    if (u[p] === 0x21) { // extension block
      if (u[p + 1] === 0xF9 && (u[p + 3] & 1)) return true; // graphic control ext. flags a transparent color
      p += 2;
      while (p < u.length && u[p] !== 0) p += u[p] + 1;
      p++;
    } else { return false; } // reached the image itself (or the trailer)
  }
  return false;
}
function webpHasAlpha_(u) {
  if (u[12] === 0x56 && u[13] === 0x50 && u[14] === 0x38) {
    if (u[15] === 0x58) return (u[20] & 0x10) !== 0;                                   // VP8X: alpha flag
    if (u[15] === 0x4C && u[20] === 0x2F) return ((u[24] >> 4) & 1) === 1;             // VP8L: alpha_is_used
  }
  return false;
}
function blobHasTransparency_(blob) {
  try {
    var raw = blob.getBytes(), n = raw.length;
    if (n < 24) return false;
    var u = new Uint8Array(n);
    for (var i = 0; i < n; i++) u[i] = raw[i] & 255;
    if (u[0] === 137 && u[1] === 80 && u[2] === 78 && u[3] === 71) {
      var f = pngTransparentFraction_(u);
      return f !== null && f >= LOGO_TRANSPARENT_MIN_FRACTION_;
    }
    if (u[0] === 71 && u[1] === 73 && u[2] === 70) return gifHasTransparency_(u);
    if (u[0] === 82 && u[1] === 73 && u[2] === 70 && u[3] === 70 && u[8] === 87 && u[9] === 69 && u[10] === 66 && u[11] === 80) return webpHasAlpha_(u);
  } catch (e) {}
  return false; // JPEG and anything unrecognized: no transparency
}
// Should this logo get the black backing? Off if the owner turned it off; otherwise only when the
// logo really is transparent. Cached per logo (by content hash) so the decode runs once.
function logoNeedsBacking_(blob) {
  if (!blob) return false;
  if (String(getConfig_().LOGO_BACKING || '').trim().toLowerCase() === 'no') return false;
  var sig = '';
  try { sig = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, blob.getBytes())); } catch (e) {}
  var props = PropertiesService.getScriptProperties();
  if (sig) {
    try { var c = JSON.parse(props.getProperty('LOGO_TRANSPARENCY') || 'null'); if (c && c.sig === sig) return !!c.t; } catch (e) {}
  }
  var t = blobHasTransparency_(blob);
  if (sig) { try { props.setProperty('LOGO_TRANSPARENCY', JSON.stringify({ sig: sig, t: t })); } catch (e) {} }
  return t;
}
// Fill the logo's header cell solid black when (and only when) the logo is transparent.
function applyLogoBacking_(cell, blob) {
  try { if (logoNeedsBacking_(blob)) cell.setBackgroundColor('#000000'); } catch (e) {}
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
      applyLogoBacking_(logoCell, logoBlob); // black square only if the logo is transparent
    }
  } catch (e) {}
  logoCell.setWidth(58);

  var infoCell = hRow.getCell(1);
  plainLine(infoCell, letterheadName_(), { bold: true, size: 12, spacing: 1, first: true });
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
  // Only show the "(Legal Business Name)" parenthetical when the owner has
  // actually set one (legalBusinessName_ returns '' for a blank/placeholder field).
  var svcLegal = legalBusinessName_();
  boldLine(partiesCell, 'Service Provider: ', getConfig_().OWNER_NAME + (svcLegal ? (' (' + svcLegal + ')') : ''));
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
  // Cancellation wording is buyer-configurable (Settings -> Contract terms); the
  // standard text below is the fallback when no custom policy is set.
  var cancelText = String(getConfig_().CANCELLATION_POLICY || '').trim()
    || 'Deposits are non-refundable upon client event cancellation. Client performance cancellations allow deposit application to an alternate date. Unforeseen performer cancellation yields a full refund.';
  bulletTerm(leftTerms, 'Cancellation & Deposit', cancelText);
  bulletTerm(leftTerms, 'Liability', 'Performer assumes no responsibility for damages or injuries during performance except in cases of gross negligence.');

  bulletTerm(rightTerms, 'Media & Photo Release', 'Photographs of the performance may be taken freely by the Client and guests. Video recording or filming of the performance is not permitted without the Service Provider\u2019s prior consent. The Service Provider may photograph and film the performance for promotional use, which the Client agrees to unless an opt-out notification is provided in writing prior to the event.', true);
  var govLoc = governingLawPhrase_();
  bulletTerm(rightTerms, 'Governing Law', 'This Agreement shall be governed and construed in accordance with the laws of ' + govLoc + '.');
  bulletTerm(rightTerms, 'Entire Agreement', 'Supersedes all prior discussions. Any modifications must be executed in writing and signed by both parties.');

  // Buyer-defined extra clauses (Settings -> Contract terms), one per line,
  // appended full-width under the standard two-column terms.
  var addlTerms = String(getConfig_().ADDITIONAL_TERMS || '').split(/\r?\n/).map(function (s) { return s.trim(); }).filter(function (s) { return s; });
  addlTerms.forEach(function (line) {
    var ap = body.appendParagraph('• ' + line);
    ap.setSpacingBefore(2).setSpacingAfter(4);
    styleText(ap.editAsText(), SMALL_SIZE);
  });

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
  const feeRate = cardFeeRate_();
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
      applyLogoBacking_(logoCell, logoBlob); // black square only if the logo is transparent
    }
  } catch (e) {}
  logoCell.setWidth(58);

  const infoCell = hRow.getCell(1);
  infoCell.clear();
  plainLine(infoCell, letterheadName_(), { bold: true, size: 12, spacing: 1 });
  plainLine(infoCell, getConfig_().ADDRESS, { spacing: 1, size: 9 });
  plainLine(infoCell, [getConfig_().EMAIL, getConfig_().PHONE].filter(function (x) { return x; }).join(' \u00b7 '), { spacing: 0, size: 9 });
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
  shareAnyoneWithLink_(pdfFile);
  // Only the PDF is needed going forward — clean up the working Doc.
  try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch (e) {}
  return pdfFile.getUrl();
}
// Share a Drive file as "anyone with the link can view", then VERIFY it actually
// became public. On Google Workspace accounts that block external sharing, the
// request gets capped to the org (or blocked), so the owner can open the file but
// their client can't — the "unable to open the file" error. We record that in a
// Script Property so the app can warn the owner instead of the client hitting a
// dead link. Returns true if the file is genuinely public.
function shareAnyoneWithLink_(file) {
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  var ok = false;
  try { var acc = file.getSharingAccess(); ok = (acc === DriveApp.Access.ANYONE_WITH_LINK || acc === DriveApp.Access.ANYONE); } catch (e) {}
  try {
    var props = PropertiesService.getScriptProperties();
    if (ok) props.deleteProperty('SHARING_BLOCKED'); else props.setProperty('SHARING_BLOCKED', '1');
  } catch (e) {}
  return ok;
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
  const feeRate = cardFeeRate_();
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
  if (feeRate > 0 && isCard && isBalanceCard) {
    body.replaceText('\\{\\{cardfeenote\\}\\}', ' \u2014 3.25% card processing fee included in both amounts above.');
  } else if (feeRate > 0 && isCard) {
    body.replaceText('\\{\\{cardfeenote\\}\\}', ' \u2014 3.25% card processing fee included in the deposit above.');
  } else if (feeRate > 0 && isBalanceCard) {
    body.replaceText('\\{\\{cardfeenote\\}\\}', ' \u2014 3.25% card processing fee included in the balance above.');
  } else {
    body.replaceText('\\{\\{cardfeenote\\}\\}', '');
  }

  doc.saveAndClose();

  const pdfBlob = DriveApp.getFileById(copy.getId()).getAs('application/pdf');
  const pdfFile = folder.createFile(pdfBlob).setName(fileName + '.pdf');
  shareAnyoneWithLink_(pdfFile);

  return { copy: copy, doc: doc, pdfFile: pdfFile, deposit: depositFinal, balance: balanceFinal, price: price, paymentMethod: paymentMethod };
}
function generateContract_(rowNum) {
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
        invoiceNumber: get('Invoice Number') || '',
        // Post-event feedback: only once the balance is paid AND the event date
        // has arrived (so an early-payer isn't asked to review a show they
        // haven't seen), and only if they haven't already rated it.
        rating: Number(get('Event Rating')) || 0,
        showFeedback: balancePaid && (daysUntilEvent != null && daysUntilEvent <= 0) && !(Number(get('Event Rating')) > 0),
        googleReviewUrl: String(getConfig_().GOOGLE_REVIEW_URL || '')
      });
    }
  }
  return JSON.stringify({ ok: false });
}
// Client-submitted post-event feedback from their portal: a 1-5 star rating and
// an optional comment, written to the lead so the owner sees it. The portal
// decides (from GOOGLE_REVIEW_URL + the rating) whether to then show the Google
// review prompt; anything it captures here stays private to the owner.
function submitEventFeedback(token, rating, review, referenceOk) {
  try {
    rating = Math.round(Number(rating) || 0);
    if (rating < 1 || rating > 5) return JSON.stringify({ ok: false, error: 'Please choose a star rating.' });
    review = String(review || '').trim().slice(0, 4000);
    // The client can opt in to being a reference. Only ever SET the flag here —
    // never clear it — so an owner's manual flag isn't wiped by a later review.
    var wantsReference = (referenceOk === true || String(referenceOk).trim().toLowerCase() === 'yes' || String(referenceOk).trim().toLowerCase() === 'true');
    const sh = sheet_();
    var heads = headers_(sh);
    const tokenCol = heads.indexOf('Contract Sign Token') + 1;
    if (!tokenCol) return JSON.stringify({ ok: false, error: 'Not found' });
    const last = sh.getLastRow();
    const tokens = sh.getRange(2, tokenCol, last - 1, 1).getValues();
    for (var i = 0; i < tokens.length; i++) {
      if (String(tokens[i][0]) === token) {
        const rowNum = i + 2;
        var vals = { 'Event Rating': rating, 'Event Review': review, 'Reviewed Date': new Date() };
        if (wantsReference) vals['Reference OK'] = 'Yes';
        Object.keys(vals).forEach(function (name) {
          var col = heads.indexOf(name) + 1;
          if (!col) { col = heads.length + 1; sh.getRange(1, col).setValue(name); heads.push(name); }
          sh.getRange(rowNum, col).setValue(vals[name]);
        });
        return JSON.stringify({ ok: true });
      }
    }
    return JSON.stringify({ ok: false, error: 'Not found' });
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
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

      try { notifyOwnerContractSigned_(heads, sh.getRange(rowNum, 1, 1, heads.length).getValues()[0], typedName, finalPaymentMethod); } catch (e) {}
      return JSON.stringify({ ok: true });
    }
  }
  return JSON.stringify({ ok: false, error: 'Not found' });
}
// Emails the owner the moment a client signs (opt-out via SIGN_NOTIFY_EMAIL).
// Best-effort: a mail failure must never block or undo the signature.
function notifyOwnerContractSigned_(heads, rowValues, typedName, paymentMethod) {
  var cfg = getConfig_();
  if (String(cfg.SIGN_NOTIFY_EMAIL || '').trim().toLowerCase() === 'no') return;
  var ownerEmail = String(cfg.EMAIL || '').trim();
  if (!ownerEmail) return;
  var get = function (name) { var idx = heads.indexOf(name); return idx > -1 ? rowValues[idx] : ''; };
  var client = String(get('Company or Organization') || get('Customer Name') || typedName || 'A client');
  var evd = get('Date of Event');
  var amounts = computeBaseAmounts_(get);
  var subj = client + ' signed their contract';
  var body = client + ' just signed their contract (signed as "' + typedName + '").\n\n'
    + (evd ? ('Event date: ' + Utilities.formatDate(new Date(evd), tz_(), 'EEEE, M/d/yyyy') + '\n') : '')
    + (amounts && amounts.deposit ? ('Deposit due: $' + Number(amounts.deposit).toFixed(2) + '\n') : '')
    + (paymentMethod ? ('Payment method chosen: ' + paymentMethod + '\n') : '')
    + '\nOpen ' + (cfg.BUSINESS_NAME || 'your app') + ' to see the lead. The signed contract is saved to the lead.';
  MailApp.sendEmail(ownerEmail, subj, body);
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

// Set a cell by its header name, creating the column if it doesn't exist yet.
function setCellByHeader_(sh, heads, rowNum, name, val) {
  var col = heads.indexOf(name) + 1;
  if (!col) { col = heads.length + 1; sh.getRange(1, col).setValue(name); heads.push(name); }
  sh.getRange(rowNum, col).setValue(val);
}

// Client tapped "I've sent my payment" on the portal (or returned from a card
// checkout). Records that they REPORTED paying the deposit/balance — a claim, not
// a confirmation. Never marks the money received; that stays the owner's one-tap
// job (confirmReportedPayment). Also emails the owner so they know to verify.
function reportPayment(token, kind) {
  try {
    kind = (kind === 'balance') ? 'balance' : 'deposit';
    const sh = sheet_();
    const heads = headers_(sh);
    const tokenCol = heads.indexOf('Contract Sign Token') + 1;
    if (!tokenCol) return JSON.stringify({ ok: false, error: 'Not found' });
    const last = sh.getLastRow();
    const values = sh.getRange(2, 1, last - 1, heads.length).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][tokenCol - 1]) === token) {
        const rowNum = i + 2;
        var flagName = kind === 'balance' ? 'Balance Reported' : 'Deposit Reported';
        var flagIdx = heads.indexOf(flagName);
        // Skip if already reported — so a page reload (e.g. the card success URL
        // being refreshed) doesn't re-notify the owner.
        var already = flagIdx > -1 && String(values[i][flagIdx]) === 'Yes';
        setCellByHeader_(sh, heads, rowNum, flagName, 'Yes');
        setCellByHeader_(sh, heads, rowNum, kind === 'balance' ? 'Balance Reported Date' : 'Deposit Reported Date', new Date());
        if (!already) { try { notifyOwnerReportedPayment_(heads, values[i], kind); } catch (e) {} }
        return JSON.stringify({ ok: true });
      }
    }
    return JSON.stringify({ ok: false, error: 'Not found' });
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
}

// Emails the owner that a client reported a payment — explicitly framed as "verify
// before confirming," never as money in hand.
function notifyOwnerReportedPayment_(heads, rowValues, kind) {
  var get = function (name) { var idx = heads.indexOf(name); return idx > -1 ? rowValues[idx] : ''; };
  var cfg = getConfig_();
  if (String(cfg.PAYMENT_REPORT_EMAIL || '').trim().toLowerCase() === 'no') return; // owner opted out
  var ownerEmail = String(cfg.EMAIL || '').trim();
  if (!ownerEmail) return;
  var amounts = computeBaseAmounts_(get);
  var method = kind === 'balance' ? String(get('Balance Payment Method') || get('Payment Method') || '') : String(get('Payment Method') || '');
  var isCard = /credit card|^card$/i.test(method);
  var base = kind === 'balance' ? amounts.balance : amounts.deposit;
  var amt = isCard ? base * (1 + cardFeeRate_()) : base;
  var client = String(get('Company or Organization') || get('Customer Name') || 'A client');
  var kindLabel = kind === 'balance' ? 'balance' : 'deposit';
  var evd = get('Date of Event');
  var subj = client + ' reports paying their ' + kindLabel;
  var body = client + ' just indicated they’ve paid their ' + kindLabel + ' of $' + Number(amt).toFixed(2)
    + (method ? ' via ' + method : '') + '.\n\n'
    + 'This is their report — the payment is NOT marked received yet. Open ' + (cfg.BUSINESS_NAME || 'your app')
    + ' and tap Confirm once you’ve verified it actually arrived.\n\n'
    + (evd ? ('Event date: ' + Utilities.formatDate(new Date(evd), tz_(), 'M/d/yyyy')) : '');
  MailApp.sendEmail(ownerEmail, subj, body);
}

// Owner confirmed (in the app) that a reported payment really arrived: marks it
// received + dated and clears the report flag. Reuses updateLead so the receipt
// generates exactly as it would from a manual "received" flip.
function confirmReportedPayment_(rowNum, kind) {
  kind = (kind === 'balance') ? 'balance' : 'deposit';
  rowNum = Number(rowNum);
  // Use the date the client REPORTED paying (≈ when the money actually moved) as
  // the received date when we have it, so the receipt reflects the real payment
  // date rather than whenever you happened to confirm. Falls back to today (which
  // is also what a verified card gets, since it's confirmed the same day it's paid).
  var sh = sheet_(), heads = headers_(sh);
  var rIdx = heads.indexOf(kind === 'balance' ? 'Balance Reported Date' : 'Deposit Reported Date');
  var reported = (rIdx > -1) ? sh.getRange(rowNum, rIdx + 1).getValue() : '';
  var dateVal = (reported instanceof Date && !isNaN(reported.getTime())) ? reported : new Date();
  var dateYmd = Utilities.formatDate(dateVal, tz_(), 'yyyy-MM-dd');
  var updates = kind === 'balance'
    ? { 'Balance Paid': 'Yes', 'Balance Paid Date': dateYmd, 'Balance Reported': '' }
    : { 'Deposit Received': 'Yes', 'Deposit Received Date': dateYmd, 'Deposit Reported': '' };
  updateLead_(rowNum, updates);
  return JSON.stringify({ ok: true });
}

// Owner dismissed a reported payment (mistaken tap, or they'll handle it manually)
// — just clears the report flag; nothing is marked received.
function dismissReportedPayment_(rowNum, kind) {
  kind = (kind === 'balance') ? 'balance' : 'deposit';
  updateLead_(Number(rowNum), kind === 'balance' ? { 'Balance Reported': '' } : { 'Deposit Reported': '' });
  return JSON.stringify({ ok: true });
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
        const feeRate = cardFeeRate_();
        const finalAmt = baseAmt * (1 + feeRate); // Card always includes the processing fee
        const cents = Math.round(finalAmt * 100);
        if (cents <= 0) return JSON.stringify({ ok: false, error: kind === 'balance' ? 'No balance is due on this contract.' : 'No deposit amount is set for this contract yet.' });

        const clientName = get('Company or Organization') || get('Customer Name') || 'Client';
        const eventDateVal = get('Date of Event');
        const eventDateFmt = eventDateVal ? Utilities.formatDate(new Date(eventDateVal), tz_(), 'M/d/yyyy') : '';

        const webAppUrl = ScriptApp.getService().getUrl();
        // {CHECKOUT_SESSION_ID} is a literal template Stripe swaps for the real
        // session id on redirect — so the return can re-fetch it and verify payment.
        const successUrl = webAppUrl + '?sign=' + encodeURIComponent(token) + '&paid=1&pk=' + kind + '&cs={CHECKOUT_SESSION_ID}';
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

// Re-fetches a Checkout Session from Stripe and returns true only if Stripe itself
// reports it paid. This is the verification: the answer comes from Stripe (via the
// secret key), never the client's browser, so a spoofed return URL can't pass.
function stripeSessionInfo_(sessionId) {
  try {
    var secretKey = String(PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY') || '').trim();
    if (!secretKey || !sessionId) return null;
    var resp = UrlFetchApp.fetch('https://api.stripe.com/v1/checkout/sessions/' + encodeURIComponent(sessionId), {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + secretKey },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) return null;
    var json = JSON.parse(resp.getContentText());
    return { paid: String(json && json.payment_status) === 'paid', amount: Number(json && json.amount_total) || 0 };
  } catch (e) { return null; }
}

// The exact cents a card charge for this lead's deposit/balance should be — the
// same math createStripeCheckoutSession used — so the return can confirm Stripe
// collected the RIGHT amount, not just that some payment succeeded.
function expectedCardCents_(token, kind) {
  try {
    var sh = sheet_(), heads = headers_(sh);
    var tokenCol = heads.indexOf('Contract Sign Token') + 1;
    if (!tokenCol) return 0;
    var last = sh.getLastRow();
    var vals = sh.getRange(2, 1, last - 1, heads.length).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][tokenCol - 1]) === token) {
        var get = function (name) { var idx = heads.indexOf(name); return idx > -1 ? vals[i][idx] : ''; };
        var a = computeBaseAmounts_(get);
        var base = kind === 'balance' ? a.balance : a.deposit;
        return Math.round(base * (1 + cardFeeRate_()) * 100);
      }
    }
    return 0;
  } catch (e) { return 0; }
}

// A card payment Stripe verified as paid — mark it received (once), firing the
// receipt exactly like a manual/one-tap confirm. Idempotent: a page reload of the
// success URL won't re-fire it.
function autoConfirmCardPayment_(token, kind) {
  kind = (kind === 'balance') ? 'balance' : 'deposit';
  var sh = sheet_();
  var heads = headers_(sh);
  var tokenCol = heads.indexOf('Contract Sign Token') + 1;
  if (!tokenCol) return;
  var last = sh.getLastRow();
  var vals = sh.getRange(2, tokenCol, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === token) {
      var rowNum = i + 2;
      var recCol = heads.indexOf(kind === 'balance' ? 'Balance Paid' : 'Deposit Received') + 1;
      var already = recCol > 0 && String(sh.getRange(rowNum, recCol).getValue()) === 'Yes';
      if (already) return; // already received — don't re-run / re-generate the receipt
      confirmReportedPayment_(rowNum, kind);
      return;
    }
  }
}

/* ===== Storefront (optional product / merch sales) =====
 * A separate public product page (?store=<id>), off unless STORE_ENABLED is set.
 * Products live in a "Store Products" sheet; orders log to "Store Sales".
 * Reuses the same Stripe flow as the sign page. Phase 1: card payment. */
var STORE_PRODUCTS_SHEET = 'Store Products';
var STORE_PRODUCTS_HEADERS = ['Product ID', 'Name', 'Price', 'Description', 'Active'];
var STORE_SALES_SHEET = 'Store Sales';
var STORE_SALES_HEADERS = ['Date', 'Product', 'Buyer Name', 'Buyer Email', 'Amount', 'Method', 'Status'];

function storeEnabled_() {
  return String(getConfig_().STORE_ENABLED || '').trim().toLowerCase() === 'yes';
}

function newStoreId_() {
  var chars = 'abcdefghijkmnpqrstuvwxyz23456789';
  var s = ''; for (var i = 0; i < 8; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return 'p_' + s;
}

function storeProductsSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(STORE_PRODUCTS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(STORE_PRODUCTS_SHEET);
    sh.getRange(1, 1, 1, STORE_PRODUCTS_HEADERS.length).setValues([STORE_PRODUCTS_HEADERS]);
    try { sh.setFrozenRows(1); } catch (e) {}
    // Example row so there's something to try — Active is "No" until you turn it on.
    sh.appendRow([newStoreId_(), 'Example product', 25, 'A short description your buyer sees on the page.', 'No']);
  }
  return sh;
}

function storeSalesSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(STORE_SALES_SHEET);
  if (!sh) {
    sh = ss.insertSheet(STORE_SALES_SHEET);
    sh.getRange(1, 1, 1, STORE_SALES_HEADERS.length).setValues([STORE_SALES_HEADERS]);
    try { sh.setFrozenRows(1); } catch (e) {}
  }
  return sh;
}

function storeGetProduct_(productId) {
  var sh = storeProductsSheet_();
  var last = sh.getLastRow();
  if (last < 2) return null;
  var rows = sh.getRange(2, 1, last - 1, STORE_PRODUCTS_HEADERS.length).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(productId).trim()) {
      return {
        id: String(rows[i][0]).trim(),
        name: String(rows[i][1] || '').trim(),
        price: Number(String(rows[i][2] || '').replace(/[^0-9.]/g, '')) || 0,
        description: String(rows[i][3] || '').trim(),
        active: String(rows[i][4] || '').trim().toLowerCase() === 'yes'
      };
    }
  }
  return null;
}

// Client-facing: the product data the store page renders.
function getStoreProduct(productId) {
  try {
    if (!storeEnabled_()) return JSON.stringify({ ok: false, error: 'This store isn’t available.' });
    var p = storeGetProduct_(productId);
    if (!p || !p.active) return JSON.stringify({ ok: false, error: 'This product isn’t available.' });
    var cfg = getConfig_();
    return JSON.stringify({ ok: true, id: p.id, name: p.name, price: p.price, description: p.description,
      businessName: cfg.BUSINESS_NAME, logo: cfg.LOGO_URL, email: cfg.EMAIL });
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
}

function storeLogOrder_(productName, buyerName, buyerEmail, amount, method, status) {
  try {
    storeSalesSheet_().appendRow([new Date(), productName || '', buyerName || '', buyerEmail || '', amount || 0, method || '', status || 'Pending']);
  } catch (e) {}
}

// Owner-facing: read the Store Sales sheet into an array the app can display.
// Each row carries its sheet row number so status/delete can target it.
function getStoreSales_() {
  try {
    var sh = storeSalesSheet_();
    var last = sh.getLastRow();
    if (last < 2) return [];
    var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    var rows = sh.getRange(2, 1, last - 1, STORE_SALES_HEADERS.length).getValues();
    var out = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!String(r[2] || '').trim() && !String(r[3] || '').trim() && !(Number(r[4]) > 0)) continue; // skip blank rows
      out.push({
        row: i + 2,
        date: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : String(r[0] || ''),
        product: String(r[1] || ''),
        name: String(r[2] || ''),
        email: String(r[3] || ''),
        amount: Number(r[4]) || 0,
        method: String(r[5] || ''),
        status: String(r[6] || 'Pending')
      });
    }
    return out;
  } catch (e) { return []; }
}

// Products for the store-sale "What they bought" dropdown: id + name + active.
// Gated on the store being enabled, and read-only (never creates the sheet) so
// installs that don't use the store are left untouched.
function getStoreProducts_() {
  try {
    if (!storeEnabled_()) return [];
    var sh = SpreadsheetApp.getActive().getSheetByName(STORE_PRODUCTS_SHEET);
    if (!sh) return [];
    var last = sh.getLastRow();
    if (last < 2) return [];
    var rows = sh.getRange(2, 1, last - 1, STORE_PRODUCTS_HEADERS.length).getValues();
    var out = [];
    for (var i = 0; i < rows.length; i++) {
      var id = String(rows[i][0] || '').trim();
      var name = String(rows[i][1] || '').trim();
      if (!id && !name) continue;
      out.push({ id: id, name: name, active: String(rows[i][4] || '').trim().toLowerCase() === 'yes' });
    }
    return out;
  } catch (e) { return []; }
}

// Owner logs a sale by hand — e.g. a Magic Manager sale via the Stripe Payment
// Link, which never passes through the in-app store. Returns the fresh list.
function recordManualStoreSale_(name, email, product, amount, method, dateYmd, received) {
  try {
    var nm = String(name || '').trim();
    var amt = Number(String(amount == null ? '' : amount).replace(/[^0-9.]/g, '')) || 0;
    if (!nm && !amt) return JSON.stringify({ ok: false, error: 'Add at least a name or an amount.' });
    var d = dateYmd ? parseYMD_(dateYmd) : new Date();
    if (!(d instanceof Date) || isNaN(d.getTime())) d = new Date();
    var recd = (received === true || /^(yes|received|true)$/i.test(String(received).trim()));
    storeSalesSheet_().appendRow([d, String(product || '').trim(), nm, String(email || '').trim(), amt, String(method || '').trim(), recd ? 'Received' : 'Pending']);
    return JSON.stringify({ ok: true, storeSales: getStoreSales_() });
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
}

// Mark a store sale Received / Pending (Status column). Returns the fresh list.
function setStoreSaleStatus_(row, status) {
  try {
    row = Number(row) || 0;
    var sh = storeSalesSheet_();
    if (row < 2 || row > sh.getLastRow()) return JSON.stringify({ ok: false, error: 'That sale no longer exists.' });
    var s = String(status || '').trim().toLowerCase() === 'received' ? 'Received' : 'Pending';
    sh.getRange(row, 7).setValue(s);
    return JSON.stringify({ ok: true, storeSales: getStoreSales_() });
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
}

function deleteStoreSale_(row) {
  try {
    row = Number(row) || 0;
    var sh = storeSalesSheet_();
    if (row < 2 || row > sh.getLastRow()) return JSON.stringify({ ok: false, error: 'That sale no longer exists.' });
    sh.deleteRow(row);
    return JSON.stringify({ ok: true, storeSales: getStoreSales_() });
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
}

// Edit a store sale in place — e.g. adjust the amount after a discount. Rewrites
// the whole row (keeps its position). Returns the fresh list.
function updateStoreSale_(row, name, email, product, amount, method, dateYmd, received) {
  try {
    row = Number(row) || 0;
    var sh = storeSalesSheet_();
    if (row < 2 || row > sh.getLastRow()) return JSON.stringify({ ok: false, error: 'That sale no longer exists.' });
    var nm = String(name || '').trim();
    var amt = Number(String(amount == null ? '' : amount).replace(/[^0-9.]/g, '')) || 0;
    if (!nm && !amt) return JSON.stringify({ ok: false, error: 'Add at least a name or an amount.' });
    var d = dateYmd ? parseYMD_(dateYmd) : new Date();
    if (!(d instanceof Date) || isNaN(d.getTime())) d = new Date();
    var recd = (received === true || /^(yes|received|true)$/i.test(String(received).trim()));
    sh.getRange(row, 1, 1, STORE_SALES_HEADERS.length).setValues([[d, String(product || '').trim(), nm, String(email || '').trim(), amt, String(method || '').trim(), recd ? 'Received' : 'Pending']]);
    return JSON.stringify({ ok: true, storeSales: getStoreSales_() });
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
}

// Card checkout for a product — mirrors createStripeCheckoutSession.
function createStoreCheckoutSession(productId, buyerName, buyerEmail) {
  try {
    if (!storeEnabled_()) return JSON.stringify({ ok: false, error: 'This store isn’t available.' });
    var secretKey = String(PropertiesService.getScriptProperties().getProperty('STRIPE_SECRET_KEY') || '').trim();
    if (!secretKey) return JSON.stringify({ ok: false, error: 'Card payment isn’t set up yet.' });
    var p = storeGetProduct_(productId);
    if (!p || !p.active) return JSON.stringify({ ok: false, error: 'This product isn’t available.' });
    var cents = Math.round(p.price * (1 + cardFeeRate_()) * 100); // card includes the processing fee (unless turned off in Settings), same as the sign page
    if (cents <= 0) return JSON.stringify({ ok: false, error: 'This product has no price set yet.' });
    var cfg = getConfig_();
    var webAppUrl = ScriptApp.getService().getUrl();
    var payload = {
      'mode': 'payment',
      'success_url': webAppUrl + '?store=' + encodeURIComponent(productId) + '&paid=1',
      'cancel_url': webAppUrl + '?store=' + encodeURIComponent(productId),
      'line_items[0][quantity]': '1',
      'line_items[0][price_data][currency]': 'usd',
      'line_items[0][price_data][unit_amount]': String(cents),
      'line_items[0][price_data][product_data][name]': cfg.BUSINESS_NAME + ' — ' + p.name,
      'metadata[product_id]': productId,
      'metadata[buyer]': buyerName || '',
      'metadata[buyer_email]': buyerEmail || ''
    };
    if (buyerEmail) payload['customer_email'] = buyerEmail;
    var resp = UrlFetchApp.fetch('https://api.stripe.com/v1/checkout/sessions', {
      method: 'post', headers: { 'Authorization': 'Bearer ' + secretKey }, payload: payload, muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    var json; try { json = JSON.parse(resp.getContentText()); } catch (e2) { return JSON.stringify({ ok: false, error: 'Stripe returned an unexpected response (HTTP ' + code + ').' }); }
    if (code >= 200 && code < 300 && json.url) {
      storeLogOrder_(p.name, buyerName, buyerEmail, p.price, 'Card', 'Pending');
      return JSON.stringify({ ok: true, url: json.url });
    }
    return JSON.stringify({ ok: false, error: (json.error && json.error.message) || ('Stripe error (HTTP ' + code + ').') });
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
}

// Client-facing: a buyer chose a pay-outside method (Venmo / Cash App / ACH /
// Check / Wire) and tapped "I've sent my payment." Card never reaches here — it
// goes through createStoreCheckoutSession/Stripe. We log a Pending order at the
// face price (no card fee, same as those methods on the sign page) for the
// owner to confirm against their own account.
function recordStoreOrder(productId, buyerName, buyerEmail, method) {
  try {
    if (!storeEnabled_()) return JSON.stringify({ ok: false, error: 'This store isn’t available.' });
    var name = String(buyerName || '').trim();
    var email = String(buyerEmail || '').trim();
    if (!name) return JSON.stringify({ ok: false, error: 'Please enter your name.' });
    if (email.indexOf('@') < 1) return JSON.stringify({ ok: false, error: 'Please enter a valid email.' });
    var m = String(method || '').trim();
    if (!m || /card/i.test(m)) return JSON.stringify({ ok: false, error: 'Please choose a payment method.' });
    var p = storeGetProduct_(productId);
    if (!p || !p.active) return JSON.stringify({ ok: false, error: 'This product isn’t available.' });
    storeLogOrder_(p.name, name, email, p.price, m, 'Pending');
    return JSON.stringify({ ok: true });
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
}

/* ---- In-app product manager (Settings) ----
 * The store owner adds/edits/deletes products from inside the app — no sheet
 * editing. Each save/delete returns the fresh list so the client re-renders
 * from one source of truth. baseUrl lets the client build each product's link. */

function storeListProducts_() {
  try {
    var sh = storeProductsSheet_();
    var last = sh.getLastRow();
    var out = [];
    if (last >= 2) {
      var rows = sh.getRange(2, 1, last - 1, STORE_PRODUCTS_HEADERS.length).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (!String(rows[i][0]).trim()) continue;
        out.push({
          id: String(rows[i][0]).trim(),
          name: String(rows[i][1] || '').trim(),
          price: Number(String(rows[i][2] || '').replace(/[^0-9.]/g, '')) || 0,
          description: String(rows[i][3] || '').trim(),
          active: String(rows[i][4] || '').trim().toLowerCase() === 'yes'
        });
      }
    }
    return JSON.stringify({ ok: true, products: out, baseUrl: ScriptApp.getService().getUrl() });
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
}

// Add (no id) or update (matching id) one product, then return the fresh list.
function storeSaveProduct_(product) {
  try {
    product = product || {};
    var name = String(product.name || '').trim();
    if (!name) return JSON.stringify({ ok: false, error: 'Please give the product a name.' });
    var price = Number(String(product.price == null ? '' : product.price).replace(/[^0-9.]/g, '')) || 0;
    var description = String(product.description || '').trim();
    var active = (product.active === true || String(product.active).trim().toLowerCase() === 'yes') ? 'Yes' : 'No';
    var id = String(product.id || '').trim();
    var sh = storeProductsSheet_();
    var last = sh.getLastRow();
    var rowNum = 0;
    if (id && last >= 2) {
      var ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === id) { rowNum = i + 2; break; }
      }
    }
    if (rowNum) {
      sh.getRange(rowNum, 1, 1, STORE_PRODUCTS_HEADERS.length).setValues([[id, name, price, description, active]]);
    } else {
      id = newStoreId_();
      sh.appendRow([id, name, price, description, active]);
    }
    return storeListProducts_();
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
}

// Delete one product by id, then return the fresh list.
function storeDeleteProduct_(id) {
  try {
    id = String(id || '').trim();
    if (!id) return JSON.stringify({ ok: false, error: 'No product specified.' });
    var sh = storeProductsSheet_();
    var last = sh.getLastRow();
    if (last >= 2) {
      var ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === id) { sh.deleteRow(i + 2); break; }
      }
    }
    return storeListProducts_();
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
}



var LEAD_HEADERS = ['Customer Name', 'Phone number', 'Company or Organization', 'E-mail', 'Audience Size', 'Service', 'Event Type', 'Notes about interaction', 'Lead Source', 'Event Location', 'Quoted Price', 'Start Time', 'End Time', 'Status', 'Followup', 'Date of Event', 'Referred To', 'Timestamp', 'Lost Reason', 'Audience Age Range', 'Deposit Amount', 'Payment Method', 'Deposit Received', 'Deposit Received Date', 'Balance Paid', 'Balance Paid Date', 'Travel Fee', 'Contract PDF URL', 'Company Address', 'Balance Payment Method', 'Balance Alert Snoozed Until', 'Referred Date', 'Partner Check-In Dismissed', 'Contract Signed', 'Contract Doc URL', 'Contract Sign URL', 'Deposit Receipt PDF URL', 'Balance Receipt PDF URL', 'Balance Due Days', 'Event Rating', 'Event Review', 'Reviewed Date', 'Reference OK'];

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
// Shared by updateExpense_ (when a date edit moves a receipt into a
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
  } catch (e) { /* never let this block getLeads_() */ }
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

function addExpense_(date, vendor, amount, category, notes, adSource) {
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
function addRecurringExpense_(vendor, amount, category, notes, adSource, freq, startYmd, endYmd, lastGenYmd) {
  var sh = recurringSheet_();
  var id = 'RX-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
  var start = startYmd ? parseYMD_(startYmd) : new Date();
  var end = endYmd ? parseYMD_(endYmd) : '';
  // lastGenYmd is set when converting an already-logged expense into a series:
  // it's that expense's date, so backfill starts the period AFTER it and never
  // duplicates the entry that's already sitting in the list.
  var lastGen = lastGenYmd ? parseYMD_(lastGenYmd) : '';
  sh.appendRow([id, true, vendor || '', Number(amount) || 0, category || 'Other', notes || '', adSource || '', freq || 'Monthly', start, end, lastGen]);
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
      end: r[9] instanceof Date ? fmtCell_(r[9], tz) : String(r[9] || ''),
      startYmd: r[8] instanceof Date ? Utilities.formatDate(r[8], tz, 'yyyy-MM-dd') : '',
      endYmd: r[9] instanceof Date ? Utilities.formatDate(r[9], tz, 'yyyy-MM-dd') : ''
    });
  }
  return out;
}
function stopRecurringExpense_(row) {
  recurringSheet_().getRange(Number(row), 2).setValue(false); // Active = false
  return JSON.stringify({ expenses: getExpenses_(), recurring: getRecurring_() });
}

// Update an existing rule's fields in place. Keeps ID (col 1), Active (2) and
// Last Generated (11) untouched, so already-posted occurrences are never
// altered or re-posted — a changed amount/frequency/end takes effect from the
// next post forward. Re-runs the generator in case the end date was pushed out
// or the frequency changed and something is now due.
function editRecurringExpense_(row, vendor, amount, category, notes, adSource, freq, startYmd, endYmd) {
  var sh = recurringSheet_();
  var r = Number(row);
  var start = startYmd ? parseYMD_(startYmd) : new Date();
  var end = endYmd ? parseYMD_(endYmd) : '';
  sh.getRange(r, 3, 1, 8).setValues([[vendor || '', Number(amount) || 0, category || 'Other', notes || '', adSource || '', freq || 'Monthly', start, end]]);
  rememberVendorCategory_(vendor, category);
  generateRecurringExpenses_();
  return JSON.stringify({ expenses: getExpenses_(), recurring: getRecurring_() });
}

// Bulk insert for CSV import — the client already did the parsing,
// category-guessing, and duplicate review before this ever gets called,
// so this just trusts the reviewed batch and writes it in one shot
// rather than one appendRow per transaction. Imported rows never carry
// an ad-source attribution — that's a manual, per-transaction judgment
// call, not something a bank CSV can tell us.
function importExpensesBatch_(rowsJson) {
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
function bulkSetAdSource_(updatesJson) {
  const sh = expensesSheet_();
  const updates = JSON.parse(updatesJson);
  updates.forEach(function(u) {
    if (u.row >= 2 && u.row <= sh.getLastRow() && u.adSource) {
      sh.getRange(u.row, 10).setValue(u.adSource);
    }
  });
  return getExpensesScoped_();
}

function updateExpense_(rowNum, date, vendor, amount, category, notes, adSource) {
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

function deleteExpense_(rowNum) {
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
function uploadExpenseReceipt_(rowNum, base64Data, mimeType, fileName) {
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

function deleteExpenseReceipt_(rowNum) {
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

function deleteLog_(id) {
  const sh = logSheet_();
  const rowNum = findLogRow_(sh, id);
  if (rowNum) sh.deleteRow(rowNum);
  return getLogsScoped_();
}

function updateLog_(id, whenStr, note) {
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
  ownerOnly_();
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

function logContact_(leadKey, customer, type, note) {
  logSheet_().appendRow([new Date(), leadKey, customer, type, note || '', Utilities.getUuid()]);
  return getLogsScoped_();
}

/* ---------- Email templates ---------- */
/* ---------- Starter message templates ----------
 * A ready-made set covering the life of a booking, so the Email / Text pickers look like a
 * menu the first time someone opens them. New installs get these when each templates sheet is
 * first created; existing users get them only when they tap "Add starter templates"
 * (Settings -> Message templates). Signatures use {owner}/{business} merge fields (filled in at
 * send time), not baked-in names. Entries are [name, subject, body] (email) or [name, body] (text). */
var STARTER_TEMPLATES_ = {
  email: [
    ['Thanks for your inquiry', 'Thanks for reaching out!',
      "Hi {firstname},\n\nThank you for reaching out. I'd love to be part of your event! I'm checking my calendar now and will confirm availability shortly.\n\nIn the meantime, it helps to know a little more: the date and location, roughly how many guests, and the atmosphere you're hoping for. Feel free to reply here, or tell me a good time for a quick call.\n\nBest,\n{owner}\n{business}"],
    ['Quote & availability', 'Your {service} quote for {eventdate}',
      "Hi {firstname},\n\nGood news: I'm available on {eventdate} and would be delighted to perform at {location}.\n\nHere's what I have for you:\n• Service: {service}\n• Date: {eventdate}\n• Location: {location}\n• Investment: {quoted}\n\nIf everything looks good, just reply and I'll send over the agreement so we can lock in your date. I'm happy to adjust anything to fit your event.\n\nBest,\n{owner}\n{business}"],
    ['Agreement ready', 'Your agreement is ready',
      "Hi {firstname},\n\nYour agreement for {eventdate} is ready. You can review it, sign, and take care of your deposit all in one place:\n\n{signlink}\n\nOnce the deposit is in, your date is officially locked. Let me know if you have any questions!\n\nBest,\n{owner}\n{business}"],
    ['Getting close', 'Looking forward to {eventdate}!',
      "Hi {firstname},\n\nJust a quick note as your event gets close! I have you down for {eventdate} at {location}, starting at {starttime}.\n\nTo help me make it great: who's the best contact person on the day, where should I park and check in, and are there any last-minute changes to the schedule or guest count?\n\nCan't wait!\n{owner}\n{business}"],
    ['Thank you & review', 'Thank you, it was a pleasure!',
      "Hi {firstname},\n\nThank you for having me at your event. It was a true pleasure, and I hope you and your guests had a wonderful time!\n\nIf you have a moment, I'd be grateful if you'd share a quick review through your event page: {signlink}\n\nReviews mean the world to a small business like mine, and if you ever need entertainment again, I'd love to be part of it.\n\nWith thanks,\n{owner}\n{business}"]
  ],
  text: [
    ['Got your inquiry', "Hi {firstname}, thanks for reaching out about your event! I'm checking my calendar and will get back to you shortly. — {owner}, {business}"],
    ['Available on your date', "Hi {firstname}, great news: I'm available on {eventdate}! I'll email your details and quote shortly. Any questions, just text me. — {owner}"],
    ['Deposit reminder', "Hi {firstname}, a quick reminder: your deposit is what officially locks in {eventdate}. You can take care of it here: {signlink} — {owner}"],
    ['Getting close', "Hi {firstname}! Looking forward to performing for you on {eventdate} at {location}. Let me know if anything changes. See you soon! — {owner}"],
    ['Thank you', "Thank you, {firstname}! It was a pleasure being part of your event. If you have a moment, a quick review would mean a lot: {signlink} — {owner}, {business}"]
  ],
  pemail: [
    ['Referring a lead to you', 'A lead for you: {lead_name}',
      "Hi {firstname},\n\nI have a date conflict and can't take this one, and I thought of you. Here are the details:\n\n{lead_details}\n\nThey're expecting to hear from someone, so please reach out to them directly if you're available. Thanks so much!\n\n— {owner}\n{business}"]
  ],
  ptext: [
    ['Referring a lead to you', "Hi {firstname}, I can't take a gig on my date and thought of you. I just emailed the details. Please reach out to {lead_firstname} if you're free! — {owner}"]
  ]
};
// Append any starters whose NAME isn't already on the sheet (case-insensitive). Never edits,
// reorders, or deletes existing rows. Each new row gets its own timestamp — a template's id IS its
// Created time, so they must all differ. Returns how many were added.
function appendMissingStarters_(sh, starters, hasSubject) {
  var have = {}, last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 2, last - 1, 1).getValues().forEach(function (r) { have[String(r[0]).trim().toLowerCase()] = 1; });
  var base = Date.now(), rows = [];
  starters.forEach(function (s) {
    if (have[String(s[0]).toLowerCase()]) return;
    var created = new Date(base + (rows.length + 1) * 1000);
    rows.push(hasSubject ? [created, s[0], s[1], s[2]] : [created, s[0], s[1]]);
  });
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return rows.length;
}
// Lower-cased template names already on a sheet ({} if the sheet doesn't exist yet).
function templateNamesOn_(sheetName) {
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName), have = {};
  if (sh && sh.getLastRow() >= 2) sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { have[String(r[0]).trim().toLowerCase()] = 1; });
  return have;
}
// The opt-in "Add starter templates" button (Settings -> Message templates).
function addStarterTemplates_() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {}
  try {
    var added = {};
    [['email', 'Email Templates', templatesSheet_, true], ['text', 'Text Templates', textTplSheet_, false],
     ['pemail', 'Partner Email Templates', partnerTemplatesSheet_, true], ['ptext', 'Partner Text Templates', partnerTextTplSheet_, false]
    ].forEach(function (s) {
      var before = templateNamesOn_(s[1]);
      appendMissingStarters_(s[2](), STARTER_TEMPLATES_[s[0]], s[3]); // (a brand-new sheet seeds its own starters too)
      // Report what the user actually gained: starters whose names weren't there before.
      added[s[0]] = STARTER_TEMPLATES_[s[0]].filter(function (t) { return !before[String(t[0]).toLowerCase()]; }).length;
    });
    return JSON.stringify({
      ok: true, added: added,
      templates: getTemplates_(), textTemplates: getTextTemplates_(),
      partnerTemplates: getPartnerTemplates_(), partnerTextTemplates: getPartnerTextTemplates_()
    });
  } finally { try { lock.releaseLock(); } catch (e) {} }
}
function templatesSheet_() {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Email Templates');
  if (!sh) {
    sh = ss.insertSheet('Email Templates');
    sh.appendRow(['Created', 'Name', 'Subject', 'Body']);
    sh.appendRow([new Date(), 'Following up', 'Following up on your magic show inquiry',
      'Hi {firstname},\n\nJust wanted to follow up on your inquiry about booking magic for your event. Do you have a few minutes to chat about details?\n\nBest,\n' + getConfig_().OWNER_NAME + '\n' + getConfig_().BUSINESS_NAME]);
    appendMissingStarters_(sh, STARTER_TEMPLATES_.email, true); // new installs start with the full starter set
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
function moveTemplate_(id, direction) {
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
function addTemplate_(name, subject, body, id) {
  const sh = templatesSheet_();
  var created = id ? new Date(id) : new Date();
  if (isNaN(created.getTime())) created = new Date();
  sh.appendRow([created, name, subject, body]);
  return getTemplatesScoped_();
}
function updateTemplate_(id, name, subject, body) {
  const sh = templatesSheet_();
  const rowNum = findTemplateRow_(sh, id);
  if (!rowNum) return getTemplatesScoped_();
  sh.getRange(rowNum, 2, 1, 3).setValues([[name, subject, body]]);
  return getTemplatesScoped_();
}
function deleteTemplate_(id) {
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
    appendMissingStarters_(sh, STARTER_TEMPLATES_.pemail, true);
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
function addPartnerTemplate_(name, subject, body, id) {
  const sh = partnerTemplatesSheet_();
  var created = id ? new Date(id) : new Date();
  if (isNaN(created.getTime())) created = new Date();
  sh.appendRow([created, name, subject, body]);
  return getPartnerTemplatesScoped_();
}
function updatePartnerTemplate_(id, name, subject, body) {
  const sh = partnerTemplatesSheet_();
  const rowNum = findPartnerTemplateRow_(sh, id);
  if (!rowNum) return getPartnerTemplatesScoped_();
  sh.getRange(rowNum, 2, 1, 3).setValues([[name, subject, body]]);
  return getPartnerTemplatesScoped_();
}
function deletePartnerTemplate_(id) {
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
    appendMissingStarters_(sh, STARTER_TEMPLATES_.text, false); // new installs start with the full starter set
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
function moveTextTemplate_(id, direction) {
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
function addTextTemplate_(name, body, id) {
  const sh = textTplSheet_();
  var created = id ? new Date(id) : new Date();
  if (isNaN(created.getTime())) created = new Date();
  sh.appendRow([created, name, body]);
  return getTextTemplatesScoped_();
}
function updateTextTemplate_(id, name, body) {
  const sh = textTplSheet_();
  const rowNum = findTextTplRow_(sh, id);
  if (!rowNum) return getTextTemplatesScoped_();
  sh.getRange(rowNum, 2, 1, 2).setValues([[name, body]]);
  return getTextTemplatesScoped_();
}
function deleteTextTemplate_(id) {
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
    appendMissingStarters_(sh, STARTER_TEMPLATES_.ptext, false);
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
function addPartnerTextTemplate_(name, body, id) {
  const sh = partnerTextTplSheet_();
  var created = id ? new Date(id) : new Date();
  if (isNaN(created.getTime())) created = new Date();
  sh.appendRow([created, name, body]);
  return getPartnerTextTemplatesScoped_();
}
function updatePartnerTextTemplate_(id, name, body) {
  const sh = partnerTextTplSheet_();
  const rowNum = findPartnerTextTplRow_(sh, id);
  if (!rowNum) return getPartnerTextTemplatesScoped_();
  sh.getRange(rowNum, 2, 1, 2).setValues([[name, body]]);
  return getPartnerTextTemplatesScoped_();
}
function deletePartnerTextTemplate_(id) {
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

function deleteLead_(rowNum) {
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
function getTrash_() {
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
    // Month the lead was deleted, for the Recycle Bin's per-month grouping +
    // "empty this month" action. Key sorts; label displays.
    o._deletedMonthKey = Utilities.formatDate(deletedAt, tz, 'yyyy-MM');
    o._deletedMonthLabel = Utilities.formatDate(deletedAt, tz, 'MMMM yyyy');
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
function getTrashScoped_() { return JSON.stringify({ trash: getTrash_() }); }

function restoreLead_(trashId) {
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

function permanentlyDeleteTrash_(trashId) {
  const trash = trashSheet_();
  const trashRow = findTrashRow_(trash, trashId);
  if (trashRow) trash.deleteRow(trashRow);
  return getTrashScoped_();
}

// Empty a whole batch at once (used by the Recycle Bin's per-month "Empty").
// Resolve every id to its current row first, then delete bottom-up so the
// shifting row numbers never invalidate a pending deletion.
function permanentlyDeleteTrashBatch_(trashIds) {
  const trash = trashSheet_();
  const want = {};
  (trashIds || []).forEach(function (id) { want[String(id)] = true; });
  const last = trash.getLastRow();
  if (last < 2) return getTrashScoped_();
  const ids = trash.getRange(2, 1, last - 1, 1).getValues();
  const rows = [];
  for (var i = 0; i < ids.length; i++) {
    if (want[String(ids[i][0])]) rows.push(i + 2);
  }
  rows.sort(function (a, b) { return b - a; });
  rows.forEach(function (r) { trash.deleteRow(r); });
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
function startContactSync_() {
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

function getContactSyncStatus_() {
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

  const result = syncContactsToGoogle_();

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

function syncContactsToGoogle_() {
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
function setTaxRate_(rate) {
  var n = Number(rate);
  if (!isNaN(n) && n >= 0 && n <= 100) {
    PropertiesService.getScriptProperties().setProperty('TAX_RATE', String(n));
  }
  return JSON.stringify({ taxRate: getTaxRate_() });
}

function getLeads_() {
  migrateAdSpendToExpenses_();
  migrateReceiptsToYearFolders_();
  return JSON.stringify({
    leads: buildLeadsArray_(), tasks: getTasks_(), logs: getLog_(), templates: getTemplates_(),
    partners: getPartners_(), partnerLogs: getPartnerLog_(),
    textTemplates: getTextTemplates_(), partnerTemplates: getPartnerTemplates_(),
    partnerTextTemplates: getPartnerTextTemplates_(), oracleGate: oracleGateCheck_(),
    trash: getTrash_(), recurring: getRecurring_(),
    topClientThreshold: Number(PropertiesService.getScriptProperties().getProperty('TOP_CLIENT_THRESHOLD')) || 2000,
    kvfOrder: getKvfOrder_(),
    kvfHidden: getKvfHidden_(),
    expenses: getExpenses_(),
    customExpenseCategories: getCustomExpenseCategories_(),
    taxRate: getTaxRate_(),
    vendorCategoryMemory: getVendorCategoryMemory_(),
    storeSales: getStoreSales_(),
    storeProducts: getStoreProducts_(),
    sharingBlocked: (PropertiesService.getScriptProperties().getProperty('SHARING_BLOCKED') === '1'),
    referralInbox: getReferralInbox_()
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
function addExpenseCategory_(name) {
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
function deleteExpenseCategory_(name) {
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
function saveFieldSettings_(orderJson, hiddenJson) {
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
function markOracleTipShown_(){
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
  return JSON.stringify({ leads: buildLeadsArray_(), logs: getLog_(), trash: getTrash_(), customFields: customLeadFields_() });
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

function logPartnerContact_(partnerId, partnerName, type, note) {
  partnerLogSheet_().appendRow([new Date(), partnerId, partnerName, type, note || '', Utilities.getUuid()]);
  return getPartnerLogsScoped_();
}

function updatePartnerLog_(id, whenStr, note) {
  const sh = partnerLogSheet_();
  const rowNum = findPartnerLogRow_(sh, id);
  if (!rowNum) return getPartnerLogsScoped_();
  if (whenStr) sh.getRange(rowNum, 1).setValue(parseYMD_(whenStr));
  sh.getRange(rowNum, 5).setValue(note || '');
  return getPartnerLogsScoped_();
}

function deletePartnerLog_(id) {
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
    sh.appendRow(['Name', 'Email', 'Phone', 'Notes', 'Entry ID', 'Referral Link']);
  } else if (!String(sh.getRange(1, 6).getValue() || '').trim()) {
    // Older sheets predate the optional partner referral link column.
    sh.getRange(1, 6).setValue('Referral Link');
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
      notes: String(r[3] || ''),
      appLink: String(r[5] || '')
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

function addPartner_(name, email, phone, notes) {
  const sh = partnersSheet_();
  sh.appendRow([name || '', email || '', phone || '', notes || '', Utilities.getUuid()]);
  return getPartnersScoped_();
}

function updatePartner_(id, name, email, phone, notes, appLink) {
  const sh = partnersSheet_();
  const rowNum = findPartnerRow_(sh, id);
  if (!rowNum) return getPartnersScoped_();
  sh.getRange(rowNum, 1, 1, 4).setValues([[name || '', email || '', phone || '', notes || '']]);
  // Only touch the referral link when the caller sent one (older clients don't).
  if (appLink !== undefined) sh.getRange(rowNum, 6).setValue(normalizeReferralLink_(appLink));
  return getPartnersScoped_();
}

function deletePartner_(id) {
  const sh = partnersSheet_();
  const rowNum = findPartnerRow_(sh, id);
  if (rowNum) sh.deleteRow(rowNum);
  return getPartnersScoped_();
}

/* ---------- Referral link + inbox (one-tap lead hand-off between Magic Manager users) ----------
 * The main app is locked behind a private access key, so a referral link can't open the
 * recipient's app directly (and must never carry that key). Instead each owner has a
 * separate "referral link" — their app URL + a rotatable inbox code (?rc=…), NOT the
 * app key. A partner saves it on the referring owner's partner record; the referring
 * owner's email then includes  <partner link>&ref=<lead>. Opening it drops the lead into
 * the recipient's Referral Inbox (this sheet), and the recipient reviews it inside their
 * own app and taps Add. Nothing becomes a lead without that tap; the code gates who can
 * post; and the code can be reset. */
const REFERRAL_INBOX_SHEET = 'Referral Inbox';
var REFERRAL_INBOX_MAX_ = 25;
// Compact payload key -> lead column header.
var REFERRAL_KEYS_ = { n: 'Customer Name', p: 'Phone number', e: 'E-mail', c: 'Company or Organization', d: 'Date of Event', s: 'Start Time', en: 'End Time', t: 'Event Type', sv: 'Service', l: 'Event Location', a: 'Audience Size', ag: 'Audience Age Range', q: 'Quoted Price', nt: 'Notes about interaction' };
var REFERRAL_MAXLEN_ = { n: 120, p: 40, e: 160, c: 160, d: 10, s: 16, en: 16, t: 80, sv: 120, l: 240, a: 40, ag: 60, q: 20, nt: 1500 };

// Keep only a well-formed https Apps Script web-app URL that carries an inbox code (rc).
// Strips everything else — critically any ?key= (the private app access key) — so a partner
// who pastes their full app address by mistake can't leak it into this sheet or into emails.
function normalizeReferralLink_(link) {
  link = String(link || '').trim();
  var m = link.match(/^(https:\/\/script\.google\.com\/(?:macros|a\/macros\/[A-Za-z0-9.\-]+)\/s\/[A-Za-z0-9_\-]+\/exec)(?:\?([^#\s]*))?/);
  if (!m) return '';
  var rc = ((m[2] || '').match(/(?:^|&)rc=([A-Za-z0-9]+)/) || [])[1];
  return rc ? (m[1] + '?rc=' + rc) : '';
}
function referralInboxCode_(create) {
  var props = PropertiesService.getScriptProperties();
  var code = props.getProperty('REFERRAL_INBOX_CODE') || '';
  if (!code && create) { code = Utilities.getUuid().replace(/-/g, '').slice(0, 14); props.setProperty('REFERRAL_INBOX_CODE', code); }
  return code;
}
// This owner's shareable referral link (created on first use). Called from the app.
function getReferralInboxLink_() {
  var base = '';
  try { base = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  if (!base) return { ok: false, message: 'Could not read your app address. Open the app from its normal link and try again.' };
  return { ok: true, link: base + '?rc=' + referralInboxCode_(true) };
}
// Rotate the code — links partners already saved stop working (use if it was shared by mistake).
function resetReferralInboxCode_() {
  PropertiesService.getScriptProperties().setProperty('REFERRAL_INBOX_CODE', Utilities.getUuid().replace(/-/g, '').slice(0, 14));
  return getReferralInboxLink_();
}

// A leading = + - @ in a cell written through the API is evaluated as a formula. Anything a
// remote sender supplies is untrusted, so defuse it (a phone number's leading + is allowed).
function neutralizeCell_(v) {
  v = String(v == null ? '' : v);
  if (/^[=@\t\r]/.test(v) || (/^[+\-]/.test(v) && !/^\+[\d\s().\-]+$/.test(v))) return ' ' + v;
  return v;
}
function sanitizeReferral_(raw) {
  if (!raw || typeof raw !== 'object') return null;
  var lead = {};
  Object.keys(REFERRAL_KEYS_).forEach(function (k) {
    if (raw[k] == null) return;
    var v = String(raw[k]).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim();
    if (!v) return;
    lead[k] = neutralizeCell_(v.slice(0, REFERRAL_MAXLEN_[k]));
  });
  if (!lead.n) return null;
  // Real calendar dates / clock times only — new Date() would quietly roll 2026-13-99 into 2027.
  var validYmd = function (s) {
    var m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!m) return false;
    var y = +m[1], mo = +m[2], d = +m[3], dt = new Date(y, mo - 1, d);
    return y >= 2000 && y <= 2100 && dt.getFullYear() === y && dt.getMonth() === mo - 1 && dt.getDate() === d;
  };
  if (lead.d && !(/^\d{4}-\d{2}-\d{2}$/.test(lead.d) && validYmd(lead.d))) delete lead.d;
  ['s', 'en'].forEach(function (k) {
    if (!lead[k]) return;
    var tm = lead[k].match(/^(\d{4}-\d{2}-\d{2}) (\d{2}):(\d{2})$/);
    // 1970-01-01 is the app's own "time with no event date" placeholder.
    if (!tm || +tm[2] > 23 || +tm[3] > 59 || !(tm[1] === '1970-01-01' || validYmd(tm[1]))) delete lead[k];
  });
  if (lead.q) { lead.q = lead.q.replace(/[^\d.]/g, ''); if (!lead.q) delete lead.q; }
  return { lead: lead, from: neutralizeCell_(String(raw.f == null ? '' : raw.f).replace(/[\u0000-\u001F]/g, ' ').trim().slice(0, 120)) };
}
function referralInboxSheet_(create) {
  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(REFERRAL_INBOX_SHEET);
  if (!sh && create) {
    sh = ss.insertSheet(REFERRAL_INBOX_SHEET);
    sh.appendRow(['Received', 'From', 'Payload', 'Entry ID']);
  }
  return sh;
}
function referralDupKey_(lead) {
  return [lead.n, lead.e, lead.p, lead.d].map(function (x) { return String(x || '').trim().toLowerCase(); }).join('|');
}
function addReferralToInbox_(clean) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {}
  try {
    var sh = referralInboxSheet_(true);
    var last = sh.getLastRow(), pending = 0, key = referralDupKey_(clean.lead);
    if (last >= 2) {
      var vals = sh.getRange(2, 3, last - 1, 1).getValues();
      for (var i = 0; i < vals.length; i++) {
        pending++;
        try { if (referralDupKey_(JSON.parse(vals[i][0])) === key) return { ok: true, dup: true }; } catch (e) {}
      }
    }
    if (pending >= REFERRAL_INBOX_MAX_) return { ok: false, full: true };
    sh.appendRow([new Date(), clean.from, JSON.stringify(clean.lead), Utilities.getUuid()]);
    return { ok: true };
  } finally { try { lock.releaseLock(); } catch (e) {} }
}
function getReferralInbox_() {
  var sh = referralInboxSheet_(false);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues(), out = [];
  for (var i = 0; i < vals.length; i++) {
    var lead = null;
    try { lead = JSON.parse(vals[i][2]); } catch (e) {}
    if (!lead || !lead.n) continue;
    out.push({
      id: String(vals[i][3] || ''),
      received: (vals[i][0] instanceof Date) ? Utilities.formatDate(vals[i][0], tz_(), 'yyyy-MM-dd') : '',
      from: String(vals[i][1] || '').trim(),
      lead: lead
    });
  }
  out.reverse(); // newest first
  return out;
}
function findReferralRow_(sh, id) {
  if (!sh || sh.getLastRow() < 2) return 0;
  var vals = sh.getRange(2, 4, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) if (String(vals[i][0]) === String(id)) return i + 2;
  return 0;
}
// The owner tapped "Add lead" on a received referral: create the lead, then clear it from the inbox.
function acceptReferral_(id) {
  var sh = referralInboxSheet_(false);
  var row = findReferralRow_(sh, id);
  if (!row) return withReferralInbox_(getLeadsScoped_());
  var r = sh.getRange(row, 1, 1, 4).getValues()[0];
  var parsed = null;
  try { var stored = JSON.parse(r[2]); stored.f = r[1]; parsed = sanitizeReferral_(stored); } catch (e) {}
  if (!parsed) { sh.deleteRow(row); return withReferralInbox_(getLeadsScoped_()); }
  var fields = {};
  Object.keys(REFERRAL_KEYS_).forEach(function (k) { if (parsed.lead[k]) fields[REFERRAL_KEYS_[k]] = parsed.lead[k]; });
  ['Start Time', 'End Time'].forEach(function (h) { if (fields[h]) fields[h] = parseYMD_(fields[h]); });
  var note = 'Referred by ' + (parsed.from || 'another Magic Manager user') + ' on ' + Utilities.formatDate(new Date(), tz_(), 'M/d/yyyy');
  fields['Notes about interaction'] = note + (fields['Notes about interaction'] ? '\n\n' + fields['Notes about interaction'] : '');
  fields['Status'] = 'New';
  fields['Lead Source'] = 'Referral';
  fields['Followup'] = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); // a warm lead — surface it for a same-day reply
  sh.deleteRow(row);
  return withReferralInbox_(addLead_(fields));
}
function dismissReferral_(id) {
  var sh = referralInboxSheet_(false);
  var row = findReferralRow_(sh, id);
  if (row) sh.deleteRow(row);
  return JSON.stringify({ referralInbox: getReferralInbox_() });
}
function withReferralInbox_(json) {
  var o = {}; try { o = JSON.parse(json); } catch (e) {}
  o.referralInbox = getReferralInbox_();
  return JSON.stringify(o);
}

// Public landing page for a referral link (?rc=<code>&ref=<payload>). Reached by the
// recipient tapping the link in an email; it only ever writes to the Referral Inbox.
function referralReceivePage_(p) {
  var esc = function (s) { return escHtml_(String(s == null ? '' : s)); };
  var page = function (icon, title, msg) {
    return HtmlService.createHtmlOutput(
      '<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Magic Manager</title></head>'
      + '<body style="margin:0;background:#0B0B10;color:#F4EBD3;font-family:Arial,Helvetica,sans-serif;display:flex;min-height:100vh;align-items:center;justify-content:center;padding:24px">'
      + '<div style="max-width:420px;text-align:center"><div style="font-size:30px;color:#E4C179">' + icon + '</div>'
      + '<h1 style="color:#E4C179;font-size:21px;margin:8px 0 10px">' + esc(title) + '</h1>'
      + '<p style="color:#B9AE94;font-size:14.5px;line-height:1.6;margin:0">' + msg + '</p></div></body></html>'
    ).setTitle('Magic Manager').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  };
  var code = referralInboxCode_(false);
  if (!code || String(p.rc || '') !== code) {
    return page('&#9888;', 'This referral link isn’t valid', 'It may have been reset. Ask whoever sent it to check that your current referral link is saved on their partner record.');
  }
  var raw = String(p.ref || '').replace(/\s+/g, '');
  if (!raw || raw.length > 8000) return page('&#9888;', 'Couldn’t read this referral', 'The link looks incomplete. Ask the sender to send it again.');
  var parsed = null;
  try {
    while (raw.length % 4) raw += '=';
    parsed = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(raw)).getDataAsString('UTF-8'));
  } catch (e) {}
  var clean = sanitizeReferral_(parsed);
  if (!clean) return page('&#9888;', 'Couldn’t read this referral', 'The link looks incomplete. Ask the sender to send it again.');
  var res = addReferralToInbox_(clean);
  if (!res.ok) return page('&#9888;', 'Your referral inbox is full', 'Open Magic Manager and add or dismiss the referrals waiting there, then tap this link again.');
  var who = clean.from ? ' from <b>' + esc(clean.from) + '</b>' : '';
  return page('&#10038;', res.dup ? 'Already in your inbox' : 'Referral received',
    '<b>' + esc(clean.lead.n) + '</b>' + who + (res.dup ? ' is already waiting in your Magic Manager.' : ' is now waiting in your Magic Manager.')
    + '<br><br>Open your app and look for <b>Referral inbox</b> at the top of Leads to review it and add it as a lead.');
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
// on every getLeads_() call but only actually does anything once, guarded
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

function addTask_(name, due, repeat, id) {
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

function toggleTask_(id) {
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

function deleteTask_(id) {
  const sh = tasksSheet_();
  const rowNum = findTaskRow_(sh, id);
  if (rowNum) sh.deleteRow(rowNum);
  try { runFrequentAutomations(); } catch (e) {}
  try { scheduleCalendarRebuild_(); } catch (e) {}
  return getTasksScoped_();
}

function updateTask_(id, name, due, repeat) {
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
function duplicateLeadAsNewBooking_(row) {
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

    return JSON.stringify({ ok: true, newRow: newRowNum, leads: JSON.parse(getLeads_()).leads });
  } catch (e) {
    return JSON.stringify({ ok: false, error: 'Server error: ' + (e && e.message ? e.message : String(e)) });
  }
}

function updateLead_(rowNum, updates) {
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
    // Also regenerate when the payment METHOD is corrected on an already-received
    // payment, so the receipt reflects how it was actually paid. maybeGenerateReceipt_
    // only reads the stored Received Date — it never rewrites it — so this keeps the
    // real payment date intact.
    if (updates['Deposit Received'] === 'Yes' || (('Deposit Received Date' in updates) && depositIsYes) || (('Payment Method' in updates) && depositIsYes)) {
      maybeGenerateReceipt_(sh, rowNum, 'deposit');
    }
    if (updates['Balance Paid'] === 'Yes' || (('Balance Paid Date' in updates) && balanceIsYes) || (('Balance Payment Method' in updates) && balanceIsYes)) {
      maybeGenerateReceipt_(sh, rowNum, 'balance');
    }
  } catch (e) { /* receipt generation never blocks a save */ }
  return getLeadsScoped_();
}

// Clears the Followup date on Completed/Lost leads whose date has passed
// (re-checked here so a stale client list can never wipe a live follow-up).
function clearStaleFollowups_(rowsJson) {
  const sh = sheet_();
  const heads = headers_(sh);
  const fc = heads.indexOf('Followup') + 1, sc = heads.indexOf('Status') + 1;
  if (!fc || !sc) return getLeadsScoped_();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  var rows = [];
  try { rows = JSON.parse(rowsJson) || []; } catch (e) {}
  rows.slice(0, 2000).forEach(function (r) {
    r = Number(r);
    if (!(r > 1) || r > sh.getLastRow()) return;
    var st = String(sh.getRange(r, sc).getValue() || '');
    var f = sh.getRange(r, fc).getValue();
    if ((st === 'Completed' || st === 'Lost') && f instanceof Date && f < today) sh.getRange(r, fc).setValue('');
  });
  return getLeadsScoped_();
}

function addLead_(fields) {
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
function importLeadsBatch_(rowsJson) {
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
      if ((h === 'Followup' || h === 'Date of Event' || h === 'Deposit Received Date' || h === 'Balance Paid Date') && v) return parseYMD_(v);
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
  ownerOnly_();
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
function markFollowupStageDone_(rowNum, note) {
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

  const data = JSON.parse(getLeads_());
  data.leads.forEach(function (L) {
    const raw = L['Followup'];
    if (!raw) return;
    const status = String(L['Status'] || '').toLowerCase();
    if (status === 'lost' || status === 'completed') return;
    // Followup arrives as a spreadsheet-tz string ("2026-09-20" or
    // "2026-09-20 14:30"). parseYMD_ reads it as a LOCAL date; new Date() alone
    // would treat the date-only form as UTC and land it the night before at the
    // wrong hour (the "reminder on the wrong day at 8:xx" bug).
    const d = new Date(parseYMD_(raw));
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
      const ed = new Date(parseYMD_(L['Date of Event']));
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
    const d = new Date(parseYMD_(T.due)); // same tz-safe parse as the follow-ups above
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
  ownerOnly_();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === 'syncFollowUps' || fn === 'runFrequentAutomations' || fn === 'onSheetChangeNormalizeDates_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncFollowUps').timeBased().atHour(6).everyDays(1).create();
  ScriptApp.newTrigger('runFrequentAutomations').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('onSheetChangeNormalizeDates_').forSpreadsheet(SpreadsheetApp.getActive()).onChange().create();
  syncFollowUps();
}

