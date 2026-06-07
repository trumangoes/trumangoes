/*****************************************************************************
 * TRU MANGOES — OPS COMMAND CENTER  ·  Google Apps Script Web App backend
 * -------------------------------------------------------------------------
 * Binds to your existing "TRU_Mangoes_Order_System" sheet.
 * Live data (no browser cache) · Google-account login · email allowlist · audit
 * Workflow rule: PAY FIRST, THEN PICKUP (enforced on the server).
 *
 * SETUP (one time):
 *   1. Set CONFIG.ALLOWLIST to your authorized Google account(s).
 *   2. Run setup() once (Run ▸ setup) and approve the permission prompt.
 *      (Detects your columns, creates an AuditLog tab. Never edits your data.)
 *   3. Deploy ▸ New deployment ▸ Web app  (Execute as: Me · Access: see README)
 *****************************************************************************/

var CONFIG = {
  SPREADSHEET_ID: '',                 // '' = bound to this sheet (Extensions ▸ Apps Script)
  ORDERS_TAB: '',                     // '' = auto-detect the tab with the order headers
  ALLOWLIST: [],                      // ONLY these Google accounts may use the app — set per deployment via setConfig({ALLOWLIST:[...]})
  REQUIRE_ALLOWLIST: true,
  TIMEZONE: 'America/Chicago',
  ZELLE_HANDLE: '9726549231',                   // your Zelle email or US mobile # that RECEIVES payments
  BUSINESS_NAME: 'TRU Mangoes',       // shown in the payment-request message to customers
  PAYPAL_HANDLE: '',                  // optional PayPal.Me username (paypal.me/<this>) — adds a one-tap card/PayPal link
  VENMO_HANDLE: '',                   // optional Venmo username (no @) — adds a prefilled Venmo link
  // Payment-request/reminder message template. Placeholders: {name} {business} {order} {amount} {zelle}.
  // Change live anytime with setConfig({ PAYMENT_MSG: '...' }) — no redeploy. Use \n for line breaks.
  PAYMENT_MSG: 'Hi {name}! Your {business} order {order} is {amount}.\nPlease pay by Zelle to: {zelle}\nIMPORTANT: put your Order ID "{order}" in the Zelle memo/note so we can match your payment. Your order will be cancelled if unpaid. \nThank you!',
  REQUIRE_PICKUP_OTP: false,          // true = customer reads back a one-time code at pickup
  OTP_TTL_MIN: 10,                    // minutes a pickup code stays valid (OTP mode only)
  OTP_CHANNEL: 'auto',                // 'sms' | 'email' | 'manual' | 'auto' (OTP mode only)
  TWILIO: { sid: '', token: '', from: '' },  // optional: fill in to send real SMS (OTP mode)
  ORDER_ID_MODE: 'random6',           // 'random6' = random unique 6-digit IDs (customer quotes it at pickup); 'sequential' = 1001,1002,...
  EMAIL: {                            // automatic customer emails (uses Gmail/MailApp — free; ~100/day Gmail, ~1500/day Workspace)
    ENABLED: false,                   // master switch for all app-sent email
    ORDER_CONFIRMATION: false,        // email the customer when their order is created
    PAYMENT_RECEIPT: false,           // email a receipt when you mark the order Paid
    PICKUP_CONFIRMATION: false,       // email a confirmation when you mark it Picked Up
    REPLY_TO: ''                     // optional reply-to address (blank = the account running the script)
  },
  UI: {                               // front-end behavior (sent to the browser at startup)
    REFRESH_SECONDS: 30,             // background auto-sync cadence (>=15 recommended)
    RECONNECT_SECONDS: 15,           // retry interval after a dropped/expired session
    DEFAULT_TAB: 'dashboard',        // landing tab: dashboard | orders | pickups | payments | locations | reports | settings
    DEFAULT_LOCATION: 'all',         // default location filter: 'all' or an exact pickup location
    ROLES: {                         // named roles -> which tabs that role may see (others are hidden AND their data is withheld)
      admin:   ['dashboard','orders','pickups','payments','locations','reports','settings'],
      pickup:  ['dashboard','orders','pickups','payments','locations','reports'],
      cashier: ['payments','orders'],
      office:  ['dashboard','orders','reports']
    },
    USER_ROLES: {                    // map a Google account -> role. Anyone allowlisted but not listed gets DEFAULT_ROLE.
      // 'you@gmail.com':   'admin',
      // 'desk@gmail.com':  'pickup'
    },
    DEFAULT_ROLE: 'admin'            // role for allowlisted users not in USER_ROLES (set to 'office'/'pickup' to lock down by default)
  },
  CANCEL_ENABLED: false,              // master ON/OFF for the entire cancel-no-show feature (set per deployment via setConfig)
  CANCEL_ADMINS: [],                  // ONLY these accounts may cancel no-show orders — set per deployment via setConfig({CANCEL_ADMINS:[...]})
  INVENTORY: {                        // boxes you BOUGHT this batch, per variety id (oversell warning; 0 = not tracked) — set via setConfig per batch
    banganapalli: 0, kesar: 0, rasalu: 0, himayat: 0
  },
  OPERATORS: {                        // standalone username/password login (for non-Google users on a public URL)
    ENABLED: true,                   // master switch for the whole operator-login path
    SESSION_HOURS: 12,               // how long a logged-in operator session lasts
    MAX_FAILS: 5,                    // failed attempts before a cooldown
    COOLDOWN_MIN: 10,                // cooldown length after too many fails
    ACCOUNTS: {                      // set per deployment via setConfig({OPERATORS:{ACCOUNTS:{...}}}); then setOperatorPassword(user,pw).
      // shape: 'username': { role:'admin'|'pickup'|'cashier'|'office', canCancel:bool, canEdit:bool, location:'all'|'<sheet pickup location>' }
      //   e.g. 'operator': { role:'admin',  canCancel:true,  canEdit:true,  location:'all'    }
      //        'frisco':   { role:'pickup', canCancel:false, canEdit:false, location:'Frisco' }
    }
  }
  // Coordination: payment/pickup emails are gated by your sheet's "Payment Email Sent" /
  // "Pickup Email Sent" columns — the app only sends if the box is empty, then ticks it,
  // so it never double-sends with any existing sheet automation (and vice-versa).
};

/* ===== PORTABLE CONFIG: in-file values above are DEFAULTS; Script Properties override them =====
   The SAME code file deploys to any account/domain unchanged. Per deployment, configure via:
       setConfig({ SPREADSHEET_ID:'...', BUSINESS_NAME:'...', ALLOWLIST:['you@gmail.com'], ... })
   and set operator passwords via setOperatorPassword(...). Everything per-deployment lives in
   Script Properties (key 'config_overrides'), never edited in the code. ----------------------- */
function isPlainObj_(x) { return x && typeof x === 'object' && !Array.isArray(x); }
function deepMerge_(target, src) {                 // objects merge recursively; arrays/scalars replace
  Object.keys(src || {}).forEach(function (k) {
    if (isPlainObj_(src[k]) && isPlainObj_(target[k])) deepMerge_(target[k], src[k]);
    else target[k] = src[k];
  });
  return target;
}
function applyConfigOverrides_() {                 // overlay stored overrides onto the in-file defaults
  var raw;
  try { raw = PropertiesService.getScriptProperties().getProperty('config_overrides'); } catch (e) { return; }
  if (!raw) return;
  try { deepMerge_(CONFIG, JSON.parse(raw)); } catch (e) {}
}
(function () { try { applyConfigOverrides_(); } catch (e) {} })();   // runs every execution, before any handler

// ----- MANUAL editor functions (not exposed to the client) -----
// Merge a partial config into this deployment's overrides. Run in the editor, e.g.:
//   setConfig({ SPREADSHEET_ID:'1abc...', BUSINESS_NAME:'TRU Mangoes', ALLOWLIST:['you@gmail.com'],
//               UI:{ USER_ROLES:{'you@gmail.com':'admin'}, DEFAULT_ROLE:'admin' },
//               OPERATORS:{ ACCOUNTS:{ super:{role:'admin',canCancel:true,canEdit:true,location:'all'} } } });
function setConfig(partial) {
  if (!isPlainObj_(partial)) throw new Error('Usage: setConfig({ key: value, ... })');
  var p = PropertiesService.getScriptProperties(), cur = {};
  var raw = p.getProperty('config_overrides');
  if (raw) { try { cur = JSON.parse(raw); } catch (e) {} }
  deepMerge_(cur, partial);
  p.setProperty('config_overrides', JSON.stringify(cur));
  applyConfigOverrides_();
  return 'Saved. Effective overrides:\n' + JSON.stringify(cur, null, 2);
}
function getConfig() { applyConfigOverrides_(); return JSON.stringify(CONFIG, null, 2); }   // view effective config
function clearConfig() { PropertiesService.getScriptProperties().deleteProperty('config_overrides'); return 'Overrides cleared — using in-file defaults.'; }

// Admin-only runtime switch for the cancel feature (persists in Script Properties; no redeploy needed).
// Exposed via api(); the SERVER cancel gate (assertCancelAdmin_) still enforces authority on every cancel.
function setCancelEnabled(on) {
  assertAuthorized_();
  if (pRole_() !== 'admin') throw new Error('Not authorized to change this setting.');
  setConfig({ CANCEL_ENABLED: !!on });
  return { ok: true, cancelEnabled: !!CONFIG.CANCEL_ENABLED, admin: pCanCancel_() };
}

// Admin-only runtime inventory setter (boxes bought this batch). Persists in Script Properties; no redeploy.
function setInventory(inv) {
  assertAuthorized_();
  if (pRole_() !== 'admin') throw new Error('Not authorized to change inventory.');
  inv = inv || {};
  var clean = {
    banganapalli: Math.max(0, num_(inv.banganapalli)),
    kesar:        Math.max(0, num_(inv.kesar)),
    rasalu:       Math.max(0, num_(inv.rasalu)),
    himayat:      Math.max(0, num_(inv.himayat))
  };
  setConfig({ INVENTORY: clean });
  return { ok: true, inventory: CONFIG.INVENTORY };
}

var AUDIT_TAB  = 'AuditLog';
var PRICES_TAB = 'Prices';     // configurable per-box prices (edit here or in Settings)

// Box varieties + DEFAULT prices (used to seed the Prices tab / as fallback).
var VARIETIES_DEF = [
  { id: 'banganapalli', name: 'Banganapalli', color: '#e8a92e', price: 47 },
  { id: 'kesar',        name: 'Kesar',        color: '#f0b84a', price: 49 },
  { id: 'rasalu',       name: 'Rasalu',       color: '#d9962a', price: 52 },
  { id: 'himayat',      name: 'Himayat',      color: '#c8780a', price: 58 }
];

/* ============================ WEB APP ENTRY ============================ */

function doGet(e) {
  // Serve the app shell to everyone (incl. anonymous). No order data is in the page —
  // all data is authorized per-call in api(); anonymous users get the operator login screen.
  var t = HtmlService.createTemplateFromFile('Index');
  t.viewerEmail = getViewerEmail_() || '';
  return t.evaluate()
    .setTitle('TRU Mangoes — Ops Command Center')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

/* ============================== SECURITY ============================== */

function getViewerEmail_() {
  // ONLY the actual signed-in visitor. Under "Execute as: Me", getEffectiveUser() is always the
  // deployer — using it as a fallback would mis-identify every anonymous visitor as the owner.
  var a = '';
  try { a = Session.getActiveUser().getEmail() || ''; } catch (err) { a = ''; }
  return a.toLowerCase();
}
function isAuthorized_(email) {
  if (!CONFIG.REQUIRE_ALLOWLIST) return true;
  return CONFIG.ALLOWLIST.map(function (x) { return String(x).toLowerCase(); }).indexOf(email) !== -1;
}
/* ----- current principal: Google viewer OR a standalone operator (set per-request by api()) ----- */
var _PRINCIPAL = null;   // null => resolve identity from the Google session (allowlist path)
function principal_() {
  if (_PRINCIPAL) return _PRINCIPAL;
  var email = getViewerEmail_();
  return { type: 'google', id: email || 'authorized user', email: email,
           authorized: isAuthorized_(email), role: roleFor_(email),
           canCancel: isCancelAdmin_(email), canEdit: isEditAdmin_(email), location: 'all' };
}
function assertAuthorized_() {
  var p = principal_();
  if (!p.authorized) throw new Error('Not authorized.');
  return p.id;
}
function pRole_()      { return principal_().role; }
function pLocation_()  { return principal_().location || 'all'; }
function pId_()        { return principal_().id; }
function pCanCancel_() { return !!CONFIG.CANCEL_ENABLED && !!principal_().canCancel; }
function pCanEdit_()   { return !!principal_().canEdit; }
function scopeOrders_(orders) {                 // location operators see only their location's orders
  var loc = pLocation_();
  return (loc && loc !== 'all') ? orders.filter(function (o) { return String(o.pickupLoc || '') === loc; }) : orders;
}
function assertOrderInScope_(order) {           // and may only act on their location's orders
  var loc = pLocation_();
  if (loc && loc !== 'all' && order && String(order.pickupLoc || '') !== loc)
    throw new Error('That order is at ' + (order.pickupLoc || 'another location') + ' — outside your pickup location (' + loc + ').');
}
/* ----- role-based tab visibility + airtight data scoping ----- */
var ALL_TABS_ = ['dashboard','orders','pickups','payments','locations','reports','settings'];
function roleFor_(email) {
  var ui = CONFIG.UI || {}, map = ui.USER_ROLES || {};
  var key = String(email || '').toLowerCase(), found = '';
  Object.keys(map).forEach(function (k) { if (String(k).toLowerCase() === key) found = map[k]; });
  return found || ui.DEFAULT_ROLE || 'admin';
}
function tabsForRole_(role) {
  var roles = (CONFIG.UI && CONFIG.UI.ROLES) || {};
  var t = roles[role];
  if (!t || !t.length) return ALL_TABS_.slice();
  return t.filter(function (x) { return ALL_TABS_.indexOf(x) >= 0; });
}
function scopesForTabs_(tabs) {
  return {
    financial: tabs.some(function (t) { return ['dashboard','orders','payments','reports'].indexOf(t) >= 0; }),
    contact:   tabs.some(function (t) { return ['orders','payments','pickups'].indexOf(t) >= 0; })
  };
}
// Return a copy of the orders with fields the role may not see removed (airtight — never sent to the browser).
function redactOrders_(orders, scopes) {
  var FIN = ['total','paymentMethod','paymentRef','paymentDate','zelleRef'];
  var CON = ['phone','whatsapp','email'];
  return orders.map(function (o) {
    var c = {}; for (var k in o) if (o.hasOwnProperty(k)) c[k] = o[k];
    if (!scopes.financial) FIN.forEach(function (f) { delete c[f]; });
    if (!scopes.contact)   CON.forEach(function (f) { delete c[f]; });
    return c;
  });
}
function isCancelAdmin_(email) {
  if (!CONFIG.CANCEL_ENABLED) return false;   // master switch off -> nobody is a cancel admin
  return (CONFIG.CANCEL_ADMINS || []).map(function (x) { return String(x).toLowerCase(); }).indexOf(String(email || '').toLowerCase()) !== -1;
}
function assertCancelAdmin_() {
  if (!CONFIG.CANCEL_ENABLED) throw new Error('Cancellation is disabled.');
  var p = principal_();
  if (!p.authorized) throw new Error('Not authorized.');
  if (!p.canCancel) throw new Error('Not authorized to cancel orders.');
  return p.id;
}
function isEditAdmin_(email) { return roleFor_(email) === 'admin'; }   // editing orders = admin role
function assertEditAdmin_() {
  var p = principal_();
  if (!p.authorized) throw new Error('Not authorized.');
  if (!p.canEdit) throw new Error('Not authorized to edit orders.');
  return p.id;
}
// Authoritative total from the Prices tab (never trust a client-sent total on edit).
function calcTotal_(boxes) {
  var price = {}; readVarieties_().forEach(function (v) { price[v.id] = num_(v.price); });
  return ['banganapalli', 'kesar', 'rasalu', 'himayat'].reduce(function (s, k) {
    return s + (num_(boxes[k]) * (price[k] || 0));
  }, 0);
}
function deniedHtml_(email) {
  return '<html><head><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>body{font-family:system-ui;background:#15100a;color:#f3e9d6;display:flex;min-height:100vh;' +
    'align-items:center;justify-content:center;margin:0}.c{text-align:center;max-width:420px;padding:40px}' +
    'h1{color:#e0654f}</style></head><body><div class="c"><h1>Access Denied</h1><p>The account <b>' +
    (email || 'unknown') + '</b> is not authorized.</p><p style="color:#8a785c;font-size:13px">' +
    'Ask the owner to add you to the allowlist in Code.gs.</p></div></body></html>';
}

/* ============================ SPREADSHEET ============================= */

function ss_() {
  return CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
                               : SpreadsheetApp.getActiveSpreadsheet();
}
function nowIso_() { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss"); }
function today_()  { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd'); }
function fmtDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]')
    return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  return String(v);
}
function normHeader_(h) { return String(h == null ? '' : h).replace(/[\u2705\u2714]/g, '').replace(/\s+/g, ' ').trim().toLowerCase(); }
function num_(v) { var n = Number(String(v).replace(/[^0-9.\-]/g, '')); return isNaN(n) ? 0 : n; }
function bool_(v) { return v === true || String(v).toLowerCase() === 'true'; }

function detectHeaderRow_(sheet) {
  var rows = Math.min(6, sheet.getLastRow() || 1);
  var cols = Math.min(40, sheet.getLastColumn() || 1);
  var grid = sheet.getRange(1, 1, rows, cols).getValues();
  for (var r = 0; r < grid.length; r++) {
    var norm = grid[r].map(normHeader_);
    if (norm.indexOf('order id') >= 0 && norm.indexOf('customer name') >= 0) return r + 1;
  }
  return 0;
}
function ordersSheet_() {
  var ss = ss_();
  if (CONFIG.ORDERS_TAB) { var s = ss.getSheetByName(CONFIG.ORDERS_TAB); if (s) return s; }
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) if (detectHeaderRow_(sheets[i]) > 0) return sheets[i];
  return sheets[0];
}
function ordersMap_() {
  var sheet = ordersSheet_();
  var hr = detectHeaderRow_(sheet) || 2;
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(hr, 1, 1, lastCol).getValues()[0].map(normHeader_);
  function col(cands) { for (var i = 0; i < cands.length; i++) { var idx = headers.indexOf(cands[i]); if (idx >= 0) return idx + 1; } return 0; }
  var map = {
    orderId: col(['order id']), orderDate: col(['order date']), name: col(['customer name', 'name']),
    phone: col(['phone']), whatsapp: col(['whatsapp']), email: col(['email']), city: col(['city']),
    banganapalli: col(['banganapalli boxes', 'banganapalli']), kesar: col(['kesar boxes', 'kesar']),
    rasalu: col(['rasalu boxes', 'rasalu']), himayat: col(['himayat boxes', 'himayat']),
    totalBoxes: col(['total boxes']), total: col(['total $', 'total', 'total amount', 'total$']),
    paymentStatus: col(['payment status']), paymentMethod: col(['payment method']),
    paymentRef: col(['payment ref', 'zelle ref', 'ref']), paymentDate: col(['payment date']),
    pickupStatus: col(['pickup status']), pickupLocation: col(['pickup location']),
    notes: col(['notes']), paymentReceived: col(['payment received']), pickedUp: col(['picked up']),
    paymentEmailSent: col(['payment email sent']), pickupEmailSent: col(['pickup email sent'])
  };
  return { sheet: sheet, headerRow: hr, map: map };
}
function cell_(row, m, key) { return m[key] ? row[m[key] - 1] : ''; }

/* ===================== CUSTOMER EMAIL (optional) =====================
   App-sent transactional email via Gmail/MailApp. Never blocks a write:
   any send failure is caught and logged, the core action still succeeds.
   Payment/pickup sends are gated by the sheet's "* Email Sent" columns so
   they coordinate with any existing automation and never double-send.   */
function emailOn_(kind) { var E = CONFIG.EMAIL || {}; return !!E.ENABLED && !!E[kind]; }
function mailSafe_(to, subject, htmlBody, orderId, tag) {
  try {
    if (!to || String(to).indexOf('@') < 0) return false;
    var opts = { htmlBody: htmlBody, name: CONFIG.BUSINESS_NAME || 'Orders' };
    if (CONFIG.EMAIL && CONFIG.EMAIL.REPLY_TO) opts.replyTo = CONFIG.EMAIL.REPLY_TO;
    MailApp.sendEmail(to, subject, htmlBody.replace(/<[^>]+>/g, ''), opts);
    return true;
  } catch (err) {
    try { audit_('system', 'EMAIL_FAIL', orderId || '', (tag || '') + ': ' + (err && err.message || err)); } catch (e) {}
    return false;
  }
}
function emailShell_(title, lines) {
  var biz = CONFIG.BUSINESS_NAME || 'TRU Mangoes';
  var body = lines.map(function (l) { return '<p style="margin:0 0 12px;color:#3a2d1d;font-size:15px;line-height:1.6">' + l + '</p>'; }).join('');
  return '<div style="font-family:Arial,Helvetica,sans-serif;max-width:520px;margin:0 auto;padding:24px;background:#fffaf2;border:1px solid #eddbb8;border-radius:12px">'
    + '<div style="font-size:20px;font-weight:bold;color:#c8780a;margin-bottom:4px">' + biz + '</div>'
    + '<div style="height:3px;width:48px;background:#e8a92e;border-radius:3px;margin-bottom:18px"></div>'
    + '<h2 style="font-size:18px;color:#2a1a10;margin:0 0 14px">' + title + '</h2>'
    + body
    + '<div style="margin-top:18px;padding-top:14px;border-top:1px solid #eddbb8;color:#8a785c;font-size:12px">Thank you for choosing ' + biz + '. 🥭</div></div>';
}
function moneyStr_(n) { return '$' + (Math.round((+n || 0) * 100) / 100); }
// Order confirmation — sent once at creation.
function sendOrderConfirmation_(order) {
  if (!emailOn_('ORDER_CONFIRMATION') || !order || !order.email) return;
  var lines = ['Hi ' + (order.name || 'there') + ',',
    'We\'ve received your order <b>#' + order.id + '</b>' + (order.totalBoxes ? ' (' + order.totalBoxes + ' box' + (order.totalBoxes > 1 ? 'es' : '') + ')' : '') + ', total <b>' + moneyStr_(order.total) + '</b>.',
    'Pickup location: <b>' + (order.pickupLoc || 'TBD') + '</b>.'];
  if (!order.paid && CONFIG.ZELLE_HANDLE)
    lines.push('To pay by Zelle, send <b>' + moneyStr_(order.total) + '</b> to <b>' + CONFIG.ZELLE_HANDLE + '</b> and include <b>' + order.id + '</b> in the memo.');
  lines.push('Please keep your order number <b>' + order.id + '</b> — you\'ll share it at pickup.');
  mailSafe_(order.email, (CONFIG.BUSINESS_NAME || 'TRU Mangoes') + ' — order #' + order.id + ' received', emailShell_('Order received', lines), order.id, 'confirm');
}
// Payment receipt — gated by the "Payment Email Sent" column (send-once).
function maybeSendPaymentEmail_(o, row, m, order) {
  if (!emailOn_('PAYMENT_RECEIPT') || !order || !order.email) return;
  if (m.paymentEmailSent && bool_(getByCol_(o.sheet, row, m.paymentEmailSent))) return;  // already sent (by app or existing automation)
  var lines = ['Hi ' + (order.name || 'there') + ',',
    'We\'ve received your payment of <b>' + moneyStr_(order.total) + '</b> for order <b>#' + order.id + '</b>. Thank you!',
    (order.zelleRef ? 'Reference: <b>' + order.zelleRef + '</b>.' : ''),
    'We\'ll have it ready at <b>' + (order.pickupLoc || 'your pickup location') + '</b>. Bring your order number <b>' + order.id + '</b> for pickup.'];
  var ok = mailSafe_(order.email, (CONFIG.BUSINESS_NAME || 'TRU Mangoes') + ' — payment received for #' + order.id, emailShell_('Payment received', lines.filter(String)), order.id, 'receipt');
  if (ok && m.paymentEmailSent) setByCol_(o.sheet, row, m.paymentEmailSent, true);
}
// Pickup confirmation — gated by the "Pickup Email Sent" column (send-once).
function maybeSendPickupEmail_(o, row, m, order) {
  if (!emailOn_('PICKUP_CONFIRMATION') || !order || !order.email) return;
  if (m.pickupEmailSent && bool_(getByCol_(o.sheet, row, m.pickupEmailSent))) return;
  var lines = ['Hi ' + (order.name || 'there') + ',',
    'Your order <b>#' + order.id + '</b> has been picked up at <b>' + (order.pickupLoc || 'our pickup location') + '</b>.',
    'We hope you enjoy your mangoes! If something isn\'t right, just reply to this email.'];
  var ok = mailSafe_(order.email, (CONFIG.BUSINESS_NAME || 'TRU Mangoes') + ' — order #' + order.id + ' picked up', emailShell_('Picked up — enjoy!', lines), order.id, 'pickup');
  if (ok && m.pickupEmailSent) setByCol_(o.sheet, row, m.pickupEmailSent, true);
}
function getByCol_(sheet, row, col) { return col > 0 ? sheet.getRange(row, col).getValue() : ''; }
function setByCol_(sheet, row, col, val) { if (col > 0) sheet.getRange(row, col).setValue(val); }
function findOrderRow_(o, orderId) {
  var col = o.map.orderId, last = o.sheet.getLastRow();
  if (last <= o.headerRow) return -1;
  var vals = o.sheet.getRange(o.headerRow + 1, col, last - o.headerRow, 1).getValues();
  for (var i = 0; i < vals.length; i++) if (String(vals[i][0]) === String(orderId)) return i + o.headerRow + 1;
  return -1;
}

/* ============================== AUDIT ================================ */
function audit_(user, action, orderId, details) {
  try {
    var sh = ss_().getSheetByName(AUDIT_TAB);
    if (!sh) { sh = ss_().insertSheet(AUDIT_TAB); sh.appendRow(['Timestamp', 'User', 'Action', 'OrderID', 'Details']); sh.setFrozenRows(1); }
    sh.appendRow([nowIso_(), user, action, orderId || '', details || '']);
  } catch (e) {}
}

/* ========================= READ (live pulls) ========================= */

// Lightweight heartbeat. Its SUCCESS means the web session is still alive.
// (No sheet read, so it's cheap to call frequently.)
function ping() {
  return { ok: true, time: nowIso_(), user: getViewerEmail_() };
}

function getBootstrap() {
  assertAuthorized_();
  var p = principal_();
  var role = p.role, tabs = tabsForRole_(role), scopes = scopesForTabs_(tabs);
  var orders = scopeOrders_(readOrders_());
  var locations = locationsFrom_(orders);
  return { user: p.id, today: today_(), orders: redactOrders_(orders, scopes), locations: locations, varieties: readVarieties_(),
           zelle: { handle: CONFIG.ZELLE_HANDLE || '', name: CONFIG.BUSINESS_NAME || 'TRU Mangoes',
                    paypal: CONFIG.PAYPAL_HANDLE || '', venmo: CONFIG.VENMO_HANDLE || '' },
           paymentMsg: CONFIG.PAYMENT_MSG || '',
           otp: { required: !!CONFIG.REQUIRE_PICKUP_OTP, channel: CONFIG.OTP_CHANNEL || 'auto', ttl: CONFIG.OTP_TTL_MIN || 10 },
           ui: { refreshSeconds: (CONFIG.UI && CONFIG.UI.REFRESH_SECONDS) || 30,
                 reconnectSeconds: (CONFIG.UI && CONFIG.UI.RECONNECT_SECONDS) || 15,
                 defaultTab: (CONFIG.UI && CONFIG.UI.DEFAULT_TAB) || 'dashboard',
                 defaultLocation: (CONFIG.UI && CONFIG.UI.DEFAULT_LOCATION) || 'all' },
           role: role, tabs: tabs, admin: pCanCancel_(), canEdit: pCanEdit_(), inventory: CONFIG.INVENTORY || {},
           cancelEnabled: !!CONFIG.CANCEL_ENABLED,
           isOperator: (p.type === 'operator'), lockedLocation: (p.location && p.location !== 'all') ? p.location : '' };
}
function getOrders() {
  assertAuthorized_();
  var scopes = scopesForTabs_(tabsForRole_(pRole_()));
  var orders = scopeOrders_(readOrders_());
  return { orders: redactOrders_(orders, scopes), today: today_(), locations: locationsFrom_(orders), varieties: readVarieties_() };
}

// Live prices from the Prices tab, falling back to the code defaults.
function readVarieties_() {
  var byId = {};
  VARIETIES_DEF.forEach(function (v) { byId[v.id] = { id: v.id, name: v.name, color: v.color, price: v.price, active: true }; });
  var sh = ss_().getSheetByName(PRICES_TAB);
  if (sh && sh.getLastRow() > 1) {
    var values = sh.getDataRange().getValues(), head = values.shift().map(normHeader_);
    var ic = { id: head.indexOf('varietyid'), name: head.indexOf('name'), price: head.indexOf('price'), active: head.indexOf('active') };
    values.forEach(function (r) {
      var id = String((ic.id >= 0 ? r[ic.id] : '') || '').trim().toLowerCase();
      if (!id) return;
      if (!byId[id]) byId[id] = { id: id, name: id, color: '#e8a92e', price: 0, active: true };
      if (ic.price >= 0) byId[id].price = num_(r[ic.price]);
      if (ic.name >= 0 && r[ic.name]) byId[id].name = String(r[ic.name]);
      if (ic.active >= 0) byId[id].active = bool_(r[ic.active]);
    });
  }
  var out = [];
  VARIETIES_DEF.forEach(function (v) { if (byId[v.id]) { out.push(byId[v.id]); delete byId[v.id]; } });
  Object.keys(byId).forEach(function (k) { out.push(byId[k]); });
  return out.filter(function (v) { return v.active !== false; });
}

// Save configurable prices back to the Prices tab.
function savePrices(prices) {
  var user = assertAuthorized_();
  if (pRole_() !== 'admin') throw new Error('Not authorized to change prices.');
  var sh = ss_().getSheetByName(PRICES_TAB);
  if (!sh) { sh = ss_().insertSheet(PRICES_TAB); sh.appendRow(['VarietyID', 'Name', 'Price', 'Active']); sh.setFrozenRows(1); }
  var last = sh.getLastRow();
  var existing = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
  var rowById = {};
  existing.forEach(function (r, i) { rowById[String(r[0]).toLowerCase()] = i + 2; });
  VARIETIES_DEF.forEach(function (v) {
    if (prices[v.id] == null) return;
    var p = Number(prices[v.id]);
    if (rowById[v.id]) sh.getRange(rowById[v.id], 3).setValue(p);
    else sh.appendRow([v.id, v.name, p, true]);
  });
  audit_(user, 'PRICES', '', JSON.stringify(prices));
  return { ok: true, varieties: readVarieties_() };
}
function locationsFrom_(orders) {
  var seen = {}, out = [];
  orders.forEach(function (o) { if (o.pickupLoc && !seen[o.pickupLoc]) { seen[o.pickupLoc] = 1; out.push(o.pickupLoc); } });
  return out;
}
function readOrders_() {
  var o = ordersMap_(), m = o.map, sheet = o.sheet, last = sheet.getLastRow();
  if (last <= o.headerRow) return [];
  var width = sheet.getLastColumn();
  var data = sheet.getRange(o.headerRow + 1, 1, last - o.headerRow, width).getValues();
  var out = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i], idVal = m.orderId ? row[m.orderId - 1] : '';
    if (idVal === '' || idVal == null) continue;
    out.push(rowToOrder_(row, m, i + o.headerRow + 1));
  }
  return out;
}
function rowToOrder_(row, m, rowNum) {
  var boxes = [
    { v: 'banganapalli', name: 'Banganapalli', q: num_(cell_(row, m, 'banganapalli')) },
    { v: 'kesar',        name: 'Kesar',        q: num_(cell_(row, m, 'kesar')) },
    { v: 'rasalu',       name: 'Rasalu',       q: num_(cell_(row, m, 'rasalu')) },
    { v: 'himayat',      name: 'Himayat',      q: num_(cell_(row, m, 'himayat')) }
  ].filter(function (b) { return b.q > 0; });
  var psRaw = String(cell_(row, m, 'paymentStatus') || '').trim();
  var received = bool_(cell_(row, m, 'paymentReceived'));
  var cancelled = /cancel/i.test(psRaw);
  var paid = (/paid/i.test(psRaw) || received) && !cancelled;
  var puRaw = String(cell_(row, m, 'pickupStatus') || '').trim();
  var pickedUp = /picked\s*up/i.test(puRaw) || bool_(cell_(row, m, 'pickedUp'));
  var pickupState = pickedUp ? 'picked_up' : (/await/i.test(puRaw) ? 'awaiting' : (/not\s*picked/i.test(puRaw) ? 'not_picked' : 'pending'));
  var total = num_(cell_(row, m, 'total'));
  var zelleRef = String(cell_(row, m, 'paymentRef') || '');
  return {
    id: String(cell_(row, m, 'orderId')), rowNum: rowNum,
    name: String(cell_(row, m, 'name') || ''), phone: String(cell_(row, m, 'phone') || ''),
    whatsapp: String(cell_(row, m, 'whatsapp') || ''), email: String(cell_(row, m, 'email') || ''),
    city: String(cell_(row, m, 'city') || ''), items: boxes,
    totalBoxes: num_(cell_(row, m, 'totalBoxes')), total: total,
    paid: paid, cancelled: cancelled,
    paymentStatus: paid ? 'paid' : (cancelled ? 'cancelled' : 'unpaid'),
    paymentMethod: String(cell_(row, m, 'paymentMethod') || ''),
    zelleRef: zelleRef,
    paymentDate: fmtDate_(cell_(row, m, 'paymentDate')),
    pickedUp: pickedUp, pickupStatus: pickupState, pickupStatusRaw: puRaw,
    pickupLoc: String(cell_(row, m, 'pickupLocation') || cell_(row, m, 'city') || ''),
    notes: String(cell_(row, m, 'notes') || ''), orderDate: fmtDate_(cell_(row, m, 'orderDate')),
    // fingerprint of decision-critical fields — used to detect concurrent changes
    stamp: [paid ? 'P' : cancelled ? 'X' : 'U', pickupState, total, zelleRef].join('|')
  };
}

/* ========================= WRITE operations ========================== */

// Force the write to commit, re-read the row from the sheet, return persisted state.
function confirmRow_(o, row) {
  SpreadsheetApp.flush();                                   // commit pending writes
  var width = o.sheet.getLastColumn();
  var vals = o.sheet.getRange(row, 1, 1, width).getValues()[0];
  return rowToOrder_(vals, o.map, row);
}

// Fresh single-row read — call when opening an order so the UI reflects the
// current sheet state, not a possibly-stale list snapshot.
function getOrder(orderId) {
  assertAuthorized_();
  var o = ordersMap_(), row = findOrderRow_(o, orderId);
  if (row < 0) throw new Error('Order not found: ' + orderId);
  var ord = confirmRow_(o, row);
  assertOrderInScope_(ord);
  var scopes = scopesForTabs_(tabsForRole_(pRole_()));
  return { order: redactOrders_([ord], scopes)[0], savedAt: nowIso_() };
}

/* ---- Admin: edit boxes on an UNPAID order (add/reduce/swap). Total recomputed from Prices. ---- */
function updateOrderItems(orderId, boxes, expect) {
  var user = assertEditAdmin_();
  boxes = boxes || {};
  var o = ordersMap_(), m = o.map, row = findOrderRow_(o, orderId);
  if (row < 0) throw new Error('Order not found: ' + orderId);
  var before = confirmRow_(o, row);
  assertOrderInScope_(before);
  if (expect && before.stamp !== expect) return { ok: false, stale: true, order: before, savedAt: nowIso_() };
  if (before.paid) throw new Error('Cannot edit a paid order. Use "Add another order" for additions.');
  if (before.pickedUp) throw new Error('Cannot edit a picked-up order.');
  if (before.cancelled) throw new Error('Cannot edit a cancelled order.');
  var b = {
    banganapalli: Math.max(0, num_(boxes.banganapalli)), kesar: Math.max(0, num_(boxes.kesar)),
    rasalu: Math.max(0, num_(boxes.rasalu)), himayat: Math.max(0, num_(boxes.himayat))
  };
  var totalBoxes = b.banganapalli + b.kesar + b.rasalu + b.himayat;
  if (totalBoxes <= 0) throw new Error('Order would have no boxes — use Cancel instead.');
  var total = calcTotal_(b);
  setByCol_(o.sheet, row, m.banganapalli, b.banganapalli);
  setByCol_(o.sheet, row, m.kesar, b.kesar);
  setByCol_(o.sheet, row, m.rasalu, b.rasalu);
  setByCol_(o.sheet, row, m.himayat, b.himayat);
  setByCol_(o.sheet, row, m.totalBoxes, totalBoxes);
  setByCol_(o.sheet, row, m.total, total);
  SpreadsheetApp.flush();
  var confirmed = confirmRow_(o, row);
  audit_(user, 'EDIT_ITEMS', orderId, 'boxes ' + before.totalBoxes + '->' + totalBoxes + ', total ' + before.total + '->' + total);
  return { ok: true, order: confirmed, savedAt: nowIso_() };
}

/* ---- Admin: cancel no-show orders (unpaid & not picked up). Reversible — paying reactivates. ---- */
function cancelRow_(o, m, row, user) {
  var before = confirmRow_(o, row);
  if (before.paid)     return { ok: false, skipped: true, reason: 'paid', order: before };
  if (before.pickedUp) return { ok: false, skipped: true, reason: 'picked_up', order: before };
  setByCol_(o.sheet, row, m.paymentStatus, 'Cancelled');
  if (m.paymentReceived) o.sheet.getRange(row, m.paymentReceived).setValue(false);
  var confirmed = confirmRow_(o, row);
  return { ok: true, order: confirmed };
}
function cancelOrder(orderId, expect) {
  var user = assertCancelAdmin_();
  var o = ordersMap_(), m = o.map, row = findOrderRow_(o, orderId);
  if (row < 0) throw new Error('Order not found: ' + orderId);
  var before = confirmRow_(o, row);
  assertOrderInScope_(before);
  if (expect && before.stamp !== expect) return { ok: false, stale: true, order: before, savedAt: nowIso_() };
  if (before.paid) throw new Error('Cannot cancel a paid order.');
  if (before.pickedUp) throw new Error('Cannot cancel a picked-up order.');
  var r = cancelRow_(o, m, row, user);
  SpreadsheetApp.flush();
  audit_(user, 'CANCEL', orderId, 'Cancelled (no-show)');
  if (!r.order.cancelled) throw new Error('Cancel not confirmed in sheet — please retry.');
  return { ok: true, order: r.order, savedAt: nowIso_() };
}
function cancelOrders(ids) {
  var user = assertCancelAdmin_();
  if (!ids || !ids.length) return { ok: true, cancelled: 0, skipped: [], orders: [] };
  var o = ordersMap_(), m = o.map, out = [], skipped = [];
  ids.forEach(function (id) {
    var row = findOrderRow_(o, id);
    if (row < 0) { skipped.push({ id: id, reason: 'not_found' }); return; }
    var loc = pLocation_();
    if (loc !== 'all' && String(confirmRow_(o, row).pickupLoc || '') !== loc) { skipped.push({ id: id, reason: 'out_of_scope' }); return; }
    var r = cancelRow_(o, m, row, user);
    if (r.ok) out.push(r.order); else skipped.push({ id: id, reason: r.reason });
  });
  SpreadsheetApp.flush();
  audit_(user, 'CANCEL_BULK', '', 'Cancelled ' + out.length + ' of ' + ids.length + ' (skipped ' + skipped.length + ')');
  return { ok: true, cancelled: out.length, skipped: skipped, orders: out, savedAt: nowIso_() };
}

// PAY NOW — marks the order paid (Zelle). Enables pickup.
// `expect` = the row fingerprint the client last saw; mismatch => stale, refuse.
function recordPayment(orderId, zelleRef, expect) {
  var user = assertAuthorized_();
  var o = ordersMap_(), m = o.map, row = findOrderRow_(o, orderId);
  if (row < 0) throw new Error('Order not found: ' + orderId);
  var before = confirmRow_(o, row);                          // current LIVE row
  assertOrderInScope_(before);
  if (expect && before.stamp !== expect) return { ok: false, stale: true, order: before, savedAt: nowIso_() };
  // NOTE: a Cancelled order CAN be paid — doing so overrides the cancellation and reactivates it as Paid.
  setByCol_(o.sheet, row, m.paymentStatus, 'Paid');
  if (m.paymentReceived) o.sheet.getRange(row, m.paymentReceived).setValue(true);
  if (m.paymentMethod && !getByCol_(o.sheet, row, m.paymentMethod)) setByCol_(o.sheet, row, m.paymentMethod, 'Zelle');
  if (m.paymentRef && zelleRef) setByCol_(o.sheet, row, m.paymentRef, zelleRef);
  if (m.paymentDate) setByCol_(o.sheet, row, m.paymentDate, today_());
  var confirmed = confirmRow_(o, row);                       // read-back from sheet
  audit_(user, 'PAYMENT', orderId, (before.cancelled ? 'Reactivated+Paid' : 'Paid') + ' ref=' + (zelleRef || ''));
  if (!confirmed.paid) throw new Error('Write not confirmed in sheet — please retry.');
  maybeSendPaymentEmail_(o, row, m, confirmed);              // receipt (gated, never blocks)
  return { ok: true, order: confirmed, savedAt: nowIso_() };
}

// MARK AS PICKED UP — gated: order MUST be paid first (checked against LIVE row).
function markPickedUp(orderId, expect) {
  var user = assertAuthorized_();
  var o = ordersMap_(), m = o.map, row = findOrderRow_(o, orderId);
  if (row < 0) throw new Error('Order not found: ' + orderId);
  var before = confirmRow_(o, row);
  assertOrderInScope_(before);
  if (expect && before.stamp !== expect) return { ok: false, stale: true, order: before, savedAt: nowIso_() };
  if (!before.paid) throw new Error('Payment required before pickup. Collect payment first.');
  setByCol_(o.sheet, row, m.pickupStatus, 'Picked Up');
  if (m.pickedUp) o.sheet.getRange(row, m.pickedUp).setValue(true);
  var confirmed = confirmRow_(o, row);
  audit_(user, 'PICKUP', orderId, 'Picked Up');
  if (!confirmed.pickedUp) throw new Error('Write not confirmed in sheet — please retry.');
  maybeSendPickupEmail_(o, row, m, confirmed);               // confirmation (gated, never blocks)
  return { ok: true, order: confirmed, savedAt: nowIso_() };
}

/* ===================== PICKUP OTP (optional) =========================
   Verifies the right person is collecting. SMS needs Twilio creds (CONFIG.TWILIO);
   otherwise falls back to email (free) or manual (staff reads the code aloud).
   Codes are stored hashed + salted in Script Properties with a short TTL — never
   in the sheet, never returned to the browser (except in 'manual' mode).        */
var OTP_MAX_TRIES = 5;
function otpProps_() { return PropertiesService.getScriptProperties(); }
function otpKey_(id) { return 'otp_' + id; }
function otpSalt_() {
  var p = otpProps_(), s = p.getProperty('otp_salt');
  if (!s) { s = Utilities.getUuid(); p.setProperty('otp_salt', s); }
  return s;
}
function otpHash_(code, id) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(code) + '|' + id + '|' + otpSalt_());
  return bytes.map(function (b) { b = (b < 0 ? b + 256 : b); return ('0' + b.toString(16)).slice(-2); }).join('');
}
function otpChannel_(order) {
  var ch = CONFIG.OTP_CHANNEL || 'auto';
  if (ch !== 'auto') return ch;
  var t = CONFIG.TWILIO || {};
  if (t.sid && t.token && t.from && order.phone) return 'sms';
  if (order.email) return 'email';
  return 'manual';
}
function maskPhone_(p) { var d = String(p || '').replace(/\D/g, ''); return d.length >= 4 ? '•••-' + d.slice(-4) : (p || ''); }
function maskEmail_(e) { e = String(e || ''); var at = e.indexOf('@'); return at < 1 ? e : e.charAt(0) + '***' + e.slice(at); }
function sendSms_(to, body) {
  var t = CONFIG.TWILIO;
  var resp = UrlFetchApp.fetch('https://api.twilio.com/2010-04-01/Accounts/' + t.sid + '/Messages.json', {
    method: 'post',
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(t.sid + ':' + t.token) },
    payload: { To: to, From: t.from, Body: body }, muteHttpExceptions: true
  });
  if (resp.getResponseCode() >= 300) throw new Error('SMS send failed (' + resp.getResponseCode() + ').');
}

// Step 1: generate + deliver a pickup code. Requires the order to be paid & not yet collected.
function generatePickupOtp(orderId) {
  var user = assertAuthorized_();
  var o = ordersMap_(), row = findOrderRow_(o, orderId);
  if (row < 0) throw new Error('Order not found: ' + orderId);
  var order = confirmRow_(o, row);
  assertOrderInScope_(order);
  if (order.cancelled) throw new Error('Order is cancelled.');
  if (!order.paid) throw new Error('Payment required before pickup. Collect payment first.');
  if (order.pickedUp) throw new Error('Order is already picked up.');
  var code = '' + Math.floor(100000 + Math.random() * 900000);
  var ttl = (CONFIG.OTP_TTL_MIN || 10) * 60 * 1000;
  otpProps_().setProperty(otpKey_(orderId), JSON.stringify({ h: otpHash_(code, orderId), exp: Date.now() + ttl, tries: 0 }));
  var channel = otpChannel_(order);
  var biz = CONFIG.BUSINESS_NAME || 'TRU Mangoes';
  var body = biz + ': your pickup code for order ' + orderId + ' is ' + code + ' (valid ' + (CONFIG.OTP_TTL_MIN || 10) + ' min).';
  var sentTo = '';
  if (channel === 'sms') { if (!order.phone) throw new Error('No phone number on file.'); sendSms_(order.phone, body); sentTo = maskPhone_(order.phone); }
  else if (channel === 'email') { if (!order.email) throw new Error('No email on file.'); MailApp.sendEmail(order.email, biz + ' — pickup code for order ' + orderId, body); sentTo = maskEmail_(order.email); }
  audit_(user, 'OTP_SENT', orderId, channel + ' ' + sentTo);
  var out = { ok: true, channel: channel, sentTo: sentTo, ttl: CONFIG.OTP_TTL_MIN || 10 };
  if (channel === 'manual') out.code = code;   // staff conveys it (no SMS/email configured)
  return out;
}

// Step 2: verify the code, then perform the (gated, concurrency-checked) pickup.
function verifyPickupOtp(orderId, code, expect) {
  var user = assertAuthorized_();
  var key = otpKey_(orderId), raw = otpProps_().getProperty(key);
  if (!raw) throw new Error('No active code — send a new one.');
  var rec = JSON.parse(raw);
  if (Date.now() > rec.exp) { otpProps_().deleteProperty(key); throw new Error('Code expired — send a new one.'); }
  if (rec.tries >= OTP_MAX_TRIES) { otpProps_().deleteProperty(key); throw new Error('Too many attempts — send a new code.'); }
  if (otpHash_(String(code || ''), orderId) !== rec.h) {
    rec.tries++; otpProps_().setProperty(key, JSON.stringify(rec));
    throw new Error('Incorrect code (' + (OTP_MAX_TRIES - rec.tries) + ' attempt' + (OTP_MAX_TRIES - rec.tries === 1 ? '' : 's') + ' left).');
  }
  otpProps_().deleteProperty(key);
  audit_(user, 'OTP_VERIFIED', orderId, '');
  return markPickedUp(orderId, expect);   // reuses the pay gate + concurrency guard + read-back
}

// Revert / set other pickup states (no gate needed for non-pickup states).
function updatePickupStatus(orderId, status, expect) {
  if (status === 'picked_up') return markPickedUp(orderId, expect);
  var user = assertAuthorized_();
  var o = ordersMap_(), m = o.map, row = findOrderRow_(o, orderId);
  if (row < 0) throw new Error('Order not found: ' + orderId);
  var before = confirmRow_(o, row);
  assertOrderInScope_(before);
  if (expect && before.stamp !== expect) return { ok: false, stale: true, order: before, savedAt: nowIso_() };
  var label = status === 'awaiting' ? 'Awaiting' : 'Not Picked';
  setByCol_(o.sheet, row, m.pickupStatus, label);
  if (m.pickedUp) o.sheet.getRange(row, m.pickedUp).setValue(false);
  var confirmed = confirmRow_(o, row);
  audit_(user, 'PICKUP', orderId, label);
  return { ok: true, order: confirmed, savedAt: nowIso_() };
}

function updateNotes(orderId, notes) {
  var user = assertAuthorized_();
  var o = ordersMap_(), m = o.map, row = findOrderRow_(o, orderId);
  if (row < 0) throw new Error('Order not found: ' + orderId);
  assertOrderInScope_(confirmRow_(o, row));
  setByCol_(o.sheet, row, m.notes, notes || '');
  var confirmed = confirmRow_(o, row);
  audit_(user, 'NOTES', orderId, '');
  return { ok: true, order: confirmed, savedAt: nowIso_() };
}

function createOrder(p) {
  var user = assertAuthorized_();
  p = p || {};
  var locLock = pLocation_();
  if (locLock !== 'all') p.pickupLoc = locLock;   // location operators can only create at their own location
  var o = ordersMap_(), m = o.map, id = p.id || nextOrderId_();
  var width = o.sheet.getLastColumn(), row = [];
  for (var i = 0; i < width; i++) row.push('');
  function put(col, val) { if (col > 0) row[col - 1] = val; }
  var b = p.boxes || {};
  put(m.orderId, id); put(m.orderDate, p.orderDate || today_()); put(m.name, p.name || '');
  put(m.phone, p.phone || ''); put(m.whatsapp, p.whatsapp || p.phone || ''); put(m.email, p.email || '');
  put(m.city, p.city || p.pickupLoc || '');
  put(m.banganapalli, b.banganapalli || 0); put(m.kesar, b.kesar || 0);
  put(m.rasalu, b.rasalu || 0); put(m.himayat, b.himayat || 0);
  put(m.totalBoxes, (b.banganapalli || 0) + (b.kesar || 0) + (b.rasalu || 0) + (b.himayat || 0));
  put(m.total, p.total || 0);
  put(m.paymentStatus, 'Pending'); put(m.paymentMethod, 'Zelle');
  put(m.pickupStatus, 'Not Picked'); put(m.pickupLocation, p.pickupLoc || '');
  put(m.notes, p.notes || '');
  o.sheet.appendRow(row);
  SpreadsheetApp.flush();
  var newRow = findOrderRow_(ordersMap_(), id);
  var confirmed = newRow > 0 ? confirmRow_(o, newRow) : null;
  audit_(user, 'CREATE_ORDER', id, 'total=' + (p.total || 0));
  if (!confirmed) throw new Error('Order write not confirmed in sheet — please retry.');
  sendOrderConfirmation_(confirmed);                          // confirmation (never blocks)
  return { ok: true, id: id, order: confirmed, savedAt: nowIso_() };
}
function nextOrderId_() {
  if ((CONFIG.ORDER_ID_MODE || 'sequential') === 'random6') return randomOrderId_();
  var orders = readOrders_(), max = 1000;
  orders.forEach(function (r) { var n = Number(String(r.id).match(/\d+/)); if (n && n > max) max = n; });
  return String(max + 1);
}
// Random, unique 6-digit Order ID (100000–999999). Customer quotes this at pickup.
function randomOrderId_() {
  var used = {};
  readOrders_().forEach(function (r) { used[String(r.id)] = true; });
  for (var i = 0; i < 50; i++) {
    var n = String(Math.floor(100000 + Math.random() * 900000));
    if (!used[n]) return n;
  }
  throw new Error('Could not generate a unique Order ID — please try again.');
}

/* ============================== EXPORT =============================== */
function exportCsvToDrive(name, csv) {
  assertAuthorized_();
  var file = DriveApp.createFile(name + '_' + today_() + '.csv', csv, MimeType.PLAIN_TEXT);
  return { url: file.getUrl() };
}

/* ============== STANDALONE OPERATOR AUTH (username/password) ============== */
function opProps_() { return PropertiesService.getScriptProperties(); }
function sha256_(s) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s))); }
function opSecret_() {
  var p = opProps_(), s = p.getProperty('op_secret');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('op_secret', s); }
  return s;
}
function opSign_(s) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(String(s), opSecret_())); }
// Set/rotate an operator password — RUN MANUALLY in the Apps Script editor, e.g. setOperatorPassword('super','MyStrongPass!')
function setOperatorPassword(user, password) {
  user = String(user || '').toLowerCase().trim();
  if (!user || !password) throw new Error('Usage: setOperatorPassword("username","password")');
  if (!CONFIG.OPERATORS || !CONFIG.OPERATORS.ACCOUNTS[user]) throw new Error('No account "' + user + '" in CONFIG.OPERATORS.ACCOUNTS');
  var salt = Utilities.getUuid();
  opProps_().setProperty('op_pw_' + user, salt + ':' + sha256_(salt + password));
  return 'Password set for operator "' + user + '".';
}
function opFailKey_(u) { return 'op_fail_' + u; }
function opLockedOut_(u) {
  var raw = opProps_().getProperty(opFailKey_(u)); if (!raw) return false;
  try { var r = JSON.parse(raw); return r.until && Date.now() < r.until; } catch (e) { return false; }
}
function opRecordFail_(u) {
  var raw = opProps_().getProperty(opFailKey_(u)), r = { n: 0, until: 0 };
  if (raw) { try { r = JSON.parse(raw); } catch (e) {} }
  r.n = (r.n || 0) + 1;
  if (r.n >= (CONFIG.OPERATORS.MAX_FAILS || 5)) { r.until = Date.now() + (CONFIG.OPERATORS.COOLDOWN_MIN || 10) * 60000; r.n = 0; }
  opProps_().setProperty(opFailKey_(u), JSON.stringify(r));
}
function opClearFails_(u) { opProps_().deleteProperty(opFailKey_(u)); }
function issueOpToken_(user) {
  var exp = Date.now() + (CONFIG.OPERATORS.SESSION_HOURS || 12) * 3600000;
  var payload = Utilities.base64EncodeWebSafe(JSON.stringify({ u: user, exp: exp }));
  return payload + '.' + opSign_(payload);
}
function operatorFromToken_(token) {
  try {
    var parts = String(token || '').split('.'); if (parts.length !== 2) return null;
    if (opSign_(parts[0]) !== parts[1]) return null;                       // tamper check
    var p = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());
    if (!p.exp || Date.now() > p.exp) return null;                         // expired
    var acct = CONFIG.OPERATORS && CONFIG.OPERATORS.ACCOUNTS[String(p.u).toLowerCase()];
    if (!acct) return null;
    return { type: 'operator', id: p.u, user: p.u, authorized: true,
             role: acct.role || 'pickup', canCancel: !!acct.canCancel, canEdit: !!acct.canEdit, location: acct.location || 'all' };
  } catch (e) { return null; }
}
// Public login (callable on the public URL). Returns a session token on success.
function operatorLogin(user, password) {
  if (!CONFIG.OPERATORS || !CONFIG.OPERATORS.ENABLED) throw new Error('Operator login is disabled.');
  user = String(user || '').toLowerCase().trim();
  var acct = CONFIG.OPERATORS.ACCOUNTS[user];
  if (opLockedOut_(user)) throw new Error('Too many attempts. Please wait a few minutes and try again.');
  var stored = acct ? opProps_().getProperty('op_pw_' + user) : null;
  var ok = false;
  if (stored) { var sp = stored.split(':'); ok = (sha256_(sp[0] + String(password || '')) === sp[1]); }
  if (!acct || !stored || !ok) { opRecordFail_(user); throw new Error('Invalid username or password.'); }
  opClearFails_(user);
  return { ok: true, token: issueOpToken_(user), user: user };
}

/* ===== Single dispatcher for the client: resolves principal (operator token or Google), then calls the action ===== */
var API_MAP = {
  getBootstrap: getBootstrap, getOrders: getOrders, getOrder: getOrder, ping: ping,
  recordPayment: recordPayment, markPickedUp: markPickedUp, updatePickupStatus: updatePickupStatus,
  updateNotes: updateNotes, createOrder: createOrder, cancelOrder: cancelOrder, cancelOrders: cancelOrders,
  updateOrderItems: updateOrderItems, generatePickupOtp: generatePickupOtp, verifyPickupOtp: verifyPickupOtp,
  savePrices: savePrices, exportCsvToDrive: exportCsvToDrive, setCancelEnabled: setCancelEnabled, setInventory: setInventory
};
function api(token, fn, args) {
  var f = API_MAP[fn];
  if (!f) throw new Error('Unknown action: ' + fn);
  _PRINCIPAL = token ? operatorFromToken_(token) : null;   // operator if valid token; else Google fallback
  if (token && !_PRINCIPAL) { _PRINCIPAL = null; throw new Error('Session expired — please sign in again.'); }
  try { return f.apply(null, args || []); }
  finally { _PRINCIPAL = null; }                            // never leak principal between requests
}

/* ===================== SETUP / VERIFY (run once) ===================== */
function setup() {
  var ss = ss_();
  if (!ss.getSheetByName(AUDIT_TAB)) {
    var sh = ss.insertSheet(AUDIT_TAB);
    sh.appendRow(['Timestamp', 'User', 'Action', 'OrderID', 'Details']);
    sh.setFrozenRows(1); sh.getRange(1, 1, 1, 5).setFontWeight('bold');
  }
  var ps = ss.getSheetByName(PRICES_TAB);
  if (!ps) ps = ss.insertSheet(PRICES_TAB);
  if (ps.getLastRow() === 0) {
    ps.appendRow(['VarietyID', 'Name', 'Price', 'Active']);
    ps.setFrozenRows(1); ps.getRange(1, 1, 1, 4).setFontWeight('bold');
    VARIETIES_DEF.forEach(function (v) { ps.appendRow([v.id, v.name, v.price, true]); });
    ps.getRange(2, 3, VARIETIES_DEF.length, 1).setNumberFormat('$#,##0.00');
  }
  var v = verifyMapping();
  return 'Setup OK. Orders tab: "' + v.tab + '" (header row ' + v.headerRow + '). ' +
         (v.missing.length ? 'Missing columns: ' + v.missing.join(', ') : 'All key columns found.') +
         ' Prices tab ready. Deploy the web app next.';
}
function verifyMapping() {
  var o = ordersMap_(), m = o.map, missing = [];
  ['orderId', 'name', 'total', 'paymentStatus', 'pickupStatus'].forEach(function (k) { if (!m[k]) missing.push(k); });
  return { tab: o.sheet.getName(), headerRow: o.headerRow, map: m, missing: missing };
}

function setMyPaymentMsg() {
  setConfig({ PAYMENT_MSG: 'Hi {name}! Your {business} order {order} is {amount}.\nPlease pay by Zelle to: {zelle}\nIMPORTANT: put your Order ID "{order}" in the Zelle memo/note so we can match your payment. Your order will be cancelled if unpaid. \nThank you!' });
}
function showConfig() {
  Logger.log(getConfig());        // prints the full effective config to the log
}
