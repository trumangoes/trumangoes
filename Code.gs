/**
 * ═══════════════════════════════════════════════════════════════════
 * TRU MANGOES — ORDER AUTOMATION (v2 — pre-launch May 23, 2026)
 * ═══════════════════════════════════════════════════════════════════
 *
 * What changed from v1:
 *   • Sheet now has separate columns per variety (Banganapalli/Kesar/Rasalu/Himayat)
 *     instead of one "Variety" string + meaningless averaged unit price.
 *   • Removed: Customer Type, Batch, Box Size, Unit Price, Variety, Scheduled Pickup Date.
 *   • EmailJS sends the owner notification + initial customer confirmation from the browser.
 *     Apps Script no longer sends the owner notification (no more duplicates).
 *   • Payment Confirmed and Pickup Thank You emails now show each variety with its
 *     count × unit price × line total, and BCC orders@trumangoes.shop.
 *   • Order ID is shown prominently in every owner-visible email.
 *
 * Setup (one time, on Google Apps Script):
 *   1. Extensions → Apps Script in your Google Sheet.
 *   2. Replace ALL existing code with this file.
 *   3. Save (Ctrl/Cmd + S).
 *   4. Run setup() once. Grant permissions.
 *   5. Deploy → New deployment → Web app:
 *      • Execute as: Me
 *      • Who has access: Anyone
 *   6. Copy the new Web App URL into index.html line 1039 (the `URL` constant).
 *      (If you're keeping the same script, the URL doesn't change — only re-deploy.)
 *
 * Last updated: May 20, 2026
 * ═══════════════════════════════════════════════════════════════════
 */


// ═══════════════════════════════════════════════════════════════════
// CONFIGURATION — Business settings, prices, sheet schema
// ═══════════════════════════════════════════════════════════════════

const CONFIG = {
  ORDERS_SHEET_NAME: '📋 ORDERS',

  // Email addresses
  OWNER_EMAIL: 'orders@trumangoes.shop',
  REPLY_TO: 'orders@trumangoes.shop',
  BCC_ALL_CUSTOMER_EMAILS: 'orders@trumangoes.shop', // every customer email also goes here

  // Admin email allowlist — these accounts can OVERRIDE locked cells
  // (e.g. uncheck a Payment Received / Picked Up box that was accidentally checked).
  // Anyone else with sheet access can SEE locked cells but cannot edit them.
  // IMPORTANT: Use the Google account email that owns/edits this spreadsheet.
  ADMIN_EMAILS: [
    'sm_mandala@yahoo.com',
    // Add additional admin Google accounts here:
    // 'partner@gmail.com',
  ],

  // ===================================================================
  // 🔐 ADMIN GUI USERS (used by trumangoes.shop/admin)
  // ===================================================================
  // Passwords are stored as SHA-256 hashes (salt + password), NOT plaintext.
  // To add/change a user: run generatePasswordHash() from the Apps Script
  // editor, enter the new password, copy the hash it shows, and paste it
  // as the passwordHash below.
  //
  // ROLES:
  //   'admin'    — full control, can override/undo any locked action
  //   'operator' — can mark payments/pickups done, but CANNOT undo
  //
  // Credentials are stored as SHA-256 hashes only — no plaintext here.
  // To change a password: run generatePasswordHash() from the Apps Script
  // editor, enter the new password, copy the hash, and replace the
  // passwordHash value below. Keep your plaintext passwords in a private
  // password manager, never in this file.
  // ===================================================================
  PASSWORD_SALT: 'tru-mango-salt-v1-9x7k2zQ',
  USERS: [
    { username: 'admin',    passwordHash: '21cbe6e3933acdc6308b968cf3b91dc1bcb3d1d7023d1df34659ea288bdbc378', role: 'admin' },
    { username: 'operator', passwordHash: '080e77cb32046b590fc58235a67c05e9cb0d2ec6e65126b02ca2f3a0be6d6acb', role: 'operator' },
  ],

  // Business details
  BUSINESS_NAME: 'TRU Mangoes',
  WHATSAPP_LINK: 'https://chat.whatsapp.com/CUPFPvN3Du04kKiVKm0dVH',
  ZELLE_PHONE: '972-654-9231',
  ZELLE_NAME: 'VSVV International',
  WEBSITE: 'https://trumangoes.shop/',
  PHONE: '972-654-9231',

  // Unit prices per variety per 3kg box — single source of truth
  PRICES: {
    banganapalli: 47,
    kesar: 55,
    rasalu: 52,
    himayat: 55
  },

  // First Order ID for this season. Bumped to 101 for the May 23, 2026 launch
  // so order numbers feel established rather than starting at #1.
  // Once real orders exist with IDs >= this value, the system continues from max+1.
  STARTING_ORDER_ID: 1001,

  // ─────────────────────────────────────────────────────────────────
  // 📋 ORDERS sheet column layout (v2)
  // Headers row = 2. Data starts row 3.
  // ─────────────────────────────────────────────────────────────────
  COL: {
    ORDER_ID: 'A',
    ORDER_DATE: 'B',
    CUSTOMER_NAME: 'C',
    PHONE: 'D',
    WHATSAPP: 'E',
    EMAIL: 'F',
    CITY: 'G',
    BANGAN_BOXES: 'H',
    KESAR_BOXES: 'I',
    RASALU_BOXES: 'J',
    HIMAYAT_BOXES: 'K',
    TOTAL_BOXES: 'L',
    TOTAL_AMOUNT: 'M',
    PAYMENT_STATUS: 'N',
    PAYMENT_METHOD: 'O',
    PAYMENT_REF: 'P',
    PAYMENT_DATE: 'Q',
    PICKUP_STATUS: 'R',
    PICKUP_LOCATION: 'S',
    NOTES: 'T',
    // Automation columns (added by setup())
    PAYMENT_RECEIVED: 'U',
    PAYMENT_EMAIL_SENT: 'V',
    PICKED_UP: 'W',
    PICKUP_EMAIL_SENT: 'X'
  },

  HEADER_ROW: 2,
  DATA_START_ROW: 3
};


// ═══════════════════════════════════════════════════════════════════
// &#127821; PART 1 — SHEET AUTOMATION (checkbox triggers)
// ═══════════════════════════════════════════════════════════════════

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.ORDERS_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Could not find sheet: ' + CONFIG.ORDERS_SHEET_NAME +
      '\n\nMake sure the ORDERS tab is named exactly: ' + CONFIG.ORDERS_SHEET_NAME);
    return;
  }
  verifyHeaders_(sheet);
  addAutomationColumns_(sheet);
  createAuditLogSheet_(ss);
  installEditTrigger_();
  SpreadsheetApp.getUi().alert(
    'Setup complete!\n\n' +
    '✓ Headers verified\n' +
    '✓ Automation columns added (U–X)\n' +
    '✓ Email log tab ready\n' +
    '✓ Edit trigger installed\n\n' +
    'Test with a tick on row 3, column U (Payment Received).'
  );
}

function verifyHeaders_(sheet) {
  const expected = [
    'Order ID','Order Date','Customer Name','Phone','WhatsApp','Email','City',
    'Banganapalli Boxes','Kesar Boxes','Rasalu Boxes','Himayat Boxes','Total Boxes',
    'Total $','Payment Status','Payment Method','Payment Ref','Payment Date',
    'Pickup Status','Pickup Location','Notes'
  ];
  const actual = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, expected.length).getValues()[0];
  for (let i = 0; i < expected.length; i++) {
    if ((actual[i] || '').toString().trim() !== expected[i]) {
      sheet.getRange(CONFIG.HEADER_ROW, i + 1).setValue(expected[i]).setFontWeight('bold').setBackground('#2D5240').setFontColor('#FFFFFF').setHorizontalAlignment('center');
    }
  }
}

function addAutomationColumns_(sheet) {
  const newHeaders = ['Payment Received ✅', 'Payment Email Sent', 'Picked Up ✅', 'Pickup Email Sent'];
  const headerRow = CONFIG.HEADER_ROW;
  const startCol = letterToColumnNumber_(CONFIG.COL.PAYMENT_RECEIVED); // U
  newHeaders.forEach((header, idx) => {
    const col = startCol + idx;
    const cell = sheet.getRange(headerRow, col);
    cell.setValue(header)
      .setFontWeight('bold').setBackground('#2D5240').setFontColor('#FFFFFF').setHorizontalAlignment('center');
    if (header.indexOf('✅') > -1) {
      const range = sheet.getRange(CONFIG.DATA_START_ROW, col, 500, 1);
      range.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    }
    sheet.setColumnWidth(col, 130);
  });
}

function createAuditLogSheet_(ss) {
  let logSheet = ss.getSheetByName('📜 EMAIL LOG');
  if (!logSheet) {
    logSheet = ss.insertSheet('📜 EMAIL LOG');
    logSheet.getRange(1, 1, 1, 6).setValues([['Timestamp', 'Order ID', 'Customer', 'Email Type', 'Recipient', 'Status']]);
    logSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#2D5240').setFontColor('#FFFFFF');
    logSheet.setFrozenRows(1);
    [160, 80, 180, 180, 220, 200].forEach((w, i) => logSheet.setColumnWidth(i + 1, w));
  }
}

function installEditTrigger_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onEditHandler') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditHandler').forSpreadsheet(ss).onEdit().create();
}


// ─────────────────────────────────────────────────────────────────
// EDIT HANDLER — fires when any cell is edited
// ─────────────────────────────────────────────────────────────────

function onEditHandler(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.ORDERS_SHEET_NAME) return;
    if (e.range.getRow() < CONFIG.DATA_START_ROW) return;
    const editedColLetter = columnNumberToLetter_(e.range.getColumn());
    const row = e.range.getRow();
    const newValue = e.range.getValue();
    if (editedColLetter === CONFIG.COL.PAYMENT_RECEIVED && newValue === true) {
      handlePaymentReceived_(sheet, row);
    }
    if (editedColLetter === CONFIG.COL.PICKED_UP && newValue === true) {
      handlePickedUp_(sheet, row);
    }
  } catch (err) {
    console.error('onEditHandler error:', err);
  }
}

function handlePaymentReceived_(sheet, row) {
  const orderData = getOrderData_(sheet, row);
  // ALWAYS update the status columns first — independent of email state.
  // (Previously this ran AFTER the "already sent" check, so re-ticking the
  //  box left Payment Status stuck on "Pending".)
  sheet.getRange(row, letterToColumnNumber_(CONFIG.COL.PAYMENT_STATUS)).setValue('Paid');
  const paymentDateCell = sheet.getRange(row, letterToColumnNumber_(CONFIG.COL.PAYMENT_DATE));
  if (!paymentDateCell.getValue()) paymentDateCell.setValue(new Date());

  if (!orderData.customerName || !orderData.email) {
    logEmail_(orderData.orderId, orderData.customerName, 'Payment Confirmed', orderData.email, 'SKIPPED - missing data');
    return;
  }
  const alreadySent = sheet.getRange(row, letterToColumnNumber_(CONFIG.COL.PAYMENT_EMAIL_SENT)).getValue();
  if (alreadySent) {
    logEmail_(orderData.orderId, orderData.customerName, 'Payment Confirmed', orderData.email, 'SKIPPED - already sent');
    return;
  }
  try {
    sendPaymentConfirmedEmail_(orderData);
    sheet.getRange(row, letterToColumnNumber_(CONFIG.COL.PAYMENT_EMAIL_SENT)).setValue(formatTimestamp_(new Date()));
    logEmail_(orderData.orderId, orderData.customerName, 'Payment Confirmed', orderData.email, 'SENT');
  } catch (err) {
    logEmail_(orderData.orderId, orderData.customerName, 'Payment Confirmed', orderData.email, 'FAILED: ' + err.toString());
  }
}

function handlePickedUp_(sheet, row) {
  const orderData = getOrderData_(sheet, row);
  // ALWAYS update pickup status first, independent of email state.
  sheet.getRange(row, letterToColumnNumber_(CONFIG.COL.PICKUP_STATUS)).setValue('Picked Up');

  if (!orderData.customerName || !orderData.email) {
    logEmail_(orderData.orderId, orderData.customerName, 'Pickup Thank You', orderData.email, 'SKIPPED - missing data');
    return;
  }
  const alreadySent = sheet.getRange(row, letterToColumnNumber_(CONFIG.COL.PICKUP_EMAIL_SENT)).getValue();
  if (alreadySent) {
    logEmail_(orderData.orderId, orderData.customerName, 'Pickup Thank You', orderData.email, 'SKIPPED - already sent');
    return;
  }
  try {
    sendThankYouEmail_(orderData);
    sheet.getRange(row, letterToColumnNumber_(CONFIG.COL.PICKUP_EMAIL_SENT)).setValue(formatTimestamp_(new Date()));
    logEmail_(orderData.orderId, orderData.customerName, 'Pickup Thank You', orderData.email, 'SENT');
  } catch (err) {
    logEmail_(orderData.orderId, orderData.customerName, 'Pickup Thank You', orderData.email, 'FAILED: ' + err.toString());
  }
}


// ─────────────────────────────────────────────────────────────────
// Read all data from one order row
// ─────────────────────────────────────────────────────────────────

function getOrderData_(sheet, row) {
  const get = (colLetter) => sheet.getRange(row, letterToColumnNumber_(colLetter)).getValue();
  const customerName = (get(CONFIG.COL.CUSTOMER_NAME) || '').toString().trim();
  const firstName = customerName.split(' ')[0] || customerName;
  return {
    row: row,
    orderId: get(CONFIG.COL.ORDER_ID),
    orderDate: formatDate_(get(CONFIG.COL.ORDER_DATE)),
    customerName: customerName,
    firstName: firstName,
    phone: get(CONFIG.COL.PHONE),
    email: (get(CONFIG.COL.EMAIL) || '').toString().trim(),
    city: get(CONFIG.COL.CITY),
    banganBoxes: Number(get(CONFIG.COL.BANGAN_BOXES)) || 0,
    kesarBoxes: Number(get(CONFIG.COL.KESAR_BOXES)) || 0,
    rasaluBoxes: Number(get(CONFIG.COL.RASALU_BOXES)) || 0,
    himayatBoxes: Number(get(CONFIG.COL.HIMAYAT_BOXES)) || 0,
    totalBoxes: Number(get(CONFIG.COL.TOTAL_BOXES)) || 0,
    total: Number(get(CONFIG.COL.TOTAL_AMOUNT)) || 0,
    pickupLocation: get(CONFIG.COL.PICKUP_LOCATION)
  };
}


// ─────────────────────────────────────────────────────────────────
// HELPER: build the per-variety HTML rows for an order
// Shows only varieties with boxes > 0, each with count × unit × line total.
// ─────────────────────────────────────────────────────────────────

function buildVarietyRows_(data) {
  const rows = [
    { name: 'Banganapalli', boxes: data.banganBoxes, price: CONFIG.PRICES.banganapalli },
    { name: 'Pedda Rasalu', boxes: data.rasaluBoxes, price: CONFIG.PRICES.rasalu },
    { name: 'Himayat',      boxes: data.himayatBoxes, price: CONFIG.PRICES.himayat },
    { name: 'Kesar',        boxes: data.kesarBoxes, price: CONFIG.PRICES.kesar }
  ];
  let html = '';
  let alt = false;
  rows.forEach(r => {
    if (r.boxes <= 0) return;
    const lineTotal = r.boxes * r.price;
    const bg = alt ? 'background:#FFF8EC;' : '';
    html += '<tr style="' + bg + '">' +
      '<td style="padding:8px 12px;">' + r.name + '</td>' +
      '<td style="padding:8px 12px;text-align:center;"><strong>' + r.boxes + '</strong> × $' + r.price + '</td>' +
      '<td style="padding:8px 12px;text-align:right;"><strong>$' + lineTotal + '</strong></td>' +
      '</tr>';
    alt = !alt;
  });
  return html;
}


// ─────────────────────────────────────────────────────────────────
// EMAIL: Payment Confirmed (to customer, BCC owner)
// ─────────────────────────────────────────────────────────────────

function sendPaymentConfirmedEmail_(data) {
  const subject = 'Payment Confirmed — Order #' + data.orderId + ' — TRU Mangoes';
  const varietyRows = buildVarietyRows_(data);

  const htmlBody =
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#FFFBF1;color:#2A1500;">' +
    // Header
    '<div style="background:#059669;color:#fff;padding:24px 20px;text-align:center;">' +
      '<h1 style="margin:0;font-size:24px;">Payment Confirmed!</h1>' +
      '<p style="margin:6px 0 0;font-size:13px;opacity:.92;">TRU MANGOES · DALLAS</p>' +
    '</div>' +
    // Order # + paid amount bar
    '<div style="background:#C8973A;color:#2A1500;padding:14px 20px;text-align:center;">' +
      '<table style="width:100%;"><tr>' +
        '<td style="text-align:left;">' +
          '<div style="font-size:11px;font-weight:700;">ORDER</div>' +
          '<div style="font-size:22px;font-weight:700;">#' + data.orderId + '</div>' +
        '</td>' +
        '<td style="text-align:right;">' +
          '<div style="font-size:11px;font-weight:700;">PAID</div>' +
          '<div style="font-size:22px;font-weight:700;">$' + data.total + '</div>' +
        '</td>' +
      '</tr></table>' +
    '</div>' +
    // Body
    '<div style="padding:24px 20px;">' +
      '<p style="font-size:16px;margin-top:0;">Hi ' + data.firstName + ',</p>' +
      '<p style="font-size:14px;line-height:1.7;">Your Zelle payment of <strong>$' + data.total + '</strong> has been received and your order is now <strong>CONFIRMED</strong>. Thank you for trusting TRU Mangoes!</p>' +

      // Order breakdown table
      '<h3 style="color:#2D6B4D;margin:18px 0 10px;">Your Order — #' + data.orderId + '</h3>' +
      '<table style="width:100%;font-size:14px;background:#fff;border-radius:8px;border:1px solid rgba(0,0,0,.08);border-collapse:separate;border-spacing:0;overflow:hidden;">' +
        '<thead><tr style="background:#2D6B4D;color:#fff;">' +
          '<th style="padding:8px 12px;text-align:left;">Variety</th>' +
          '<th style="padding:8px 12px;text-align:center;">Boxes × Price</th>' +
          '<th style="padding:8px 12px;text-align:right;">Line Total</th>' +
        '</tr></thead>' +
        '<tbody>' + varietyRows + '</tbody>' +
        '<tfoot>' +
          '<tr style="background:#F0FBF5;border-top:2px solid #2D6B4D;">' +
            '<td style="padding:10px 12px;font-weight:700;color:#2D6B4D;">TOTAL</td>' +
            '<td style="padding:10px 12px;text-align:center;font-weight:700;">' + data.totalBoxes + ' boxes</td>' +
            '<td style="padding:10px 12px;text-align:right;font-weight:700;font-size:16px;color:#2D6B4D;">$' + data.total + '</td>' +
          '</tr>' +
          '<tr><td style="padding:6px 12px;color:#666;">Pickup Area</td><td colspan="2" style="padding:6px 12px;text-align:right;"><strong>' + (data.pickupLocation || data.city || 'TBD') + '</strong></td></tr>' +
        '</tfoot>' +
      '</table>' +

      // Pickup info callout
      '<div style="background:#FFF8EC;border:2px solid #C8973A;border-radius:10px;padding:18px;margin:20px 0;">' +
        '<h3 style="margin:0 0 10px;color:#5A2800;">What about pickup?</h3>' +
        '<p style="font-size:14px;line-height:1.7;margin:0;">Pickup location, date, and time window will be shared in our WhatsApp group <strong>1–2 days before pickup</strong> — once the batch has arrived in Dallas and is sorted for distribution.</p>' +
      '</div>' +

      // WhatsApp CTA
      '<div style="background:#fff;border:1px solid rgba(0,0,0,.1);border-radius:10px;padding:18px;margin:18px 0;text-align:center;">' +
        '<h3 style="margin:0 0 8px;color:#2D6B4D;">Join our WhatsApp Group</h3>' +
        '<p style="font-size:13px;color:#555;margin:0 0 10px;">Pickup details, batch updates, and direct contact with us.</p>' +
        '<a href="' + CONFIG.WHATSAPP_LINK + '" style="display:inline-block;background:#25D366;color:#fff;padding:11px 24px;border-radius:6px;text-decoration:none;font-weight:600;">&#128172; Join WhatsApp Group</a>' +
      '</div>' +

      '<p style="font-size:14px;margin-top:20px;">Thank you again for your order!<br/><strong>— The TRU Mangoes Team</strong></p>' +
    '</div>' +
    // Footer
    '<div style="background:#2A1500;color:#FFFBF1;padding:16px;text-align:center;font-size:11px;">trumangoes.shop · orders@trumangoes.shop · 972-654-9231<br/>© 2026 TRU Mangoes · VSVV International, LLC</div>' +
    '</div>';

  GmailApp.sendEmail(data.email, subject, '', {
    htmlBody: htmlBody,
    name: CONFIG.BUSINESS_NAME,
    replyTo: CONFIG.REPLY_TO,
    bcc: CONFIG.BCC_ALL_CUSTOMER_EMAILS
  });
}


// ─────────────────────────────────────────────────────────────────
// EMAIL: Pickup Thank You (to customer, BCC owner)
// ─────────────────────────────────────────────────────────────────

function sendThankYouEmail_(data) {
  const subject = 'Thanks for picking up — Order #' + data.orderId + ' — TRU Mangoes';
  const varietyRows = buildVarietyRows_(data);

  const htmlBody =
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#FFFBF1;color:#2A1500;">' +
    '<div style="background:#C8973A;color:#fff;padding:24px 20px;text-align:center;">' +
      '<h1 style="margin:0;font-size:24px;">Thank You! &#127821;</h1>' +
      '<p style="margin:6px 0 0;font-size:13px;opacity:.92;">YOUR MANGOES ARE HOME — ORDER #' + data.orderId + '</p>' +
    '</div>' +
    '<div style="padding:24px 20px;">' +
      '<p style="font-size:16px;margin-top:0;">Hi ' + data.firstName + ',</p>' +
      '<p style="font-size:14px;line-height:1.7;">Thank you for picking up your TRU Mangoes order today! We hope every box brings you and your family joy and delicious memories.</p>' +

      // Order summary
      '<h3 style="color:#2D6B4D;margin:18px 0 10px;">Order #' + data.orderId + ' Summary</h3>' +
      '<table style="width:100%;font-size:14px;background:#fff;border-radius:8px;border:1px solid rgba(0,0,0,.08);border-collapse:separate;border-spacing:0;overflow:hidden;">' +
        '<thead><tr style="background:#2D6B4D;color:#fff;">' +
          '<th style="padding:8px 12px;text-align:left;">Variety</th>' +
          '<th style="padding:8px 12px;text-align:center;">Boxes × Price</th>' +
          '<th style="padding:8px 12px;text-align:right;">Line Total</th>' +
        '</tr></thead>' +
        '<tbody>' + varietyRows + '</tbody>' +
        '<tfoot><tr style="background:#F0FBF5;border-top:2px solid #2D6B4D;">' +
          '<td style="padding:10px 12px;font-weight:700;color:#2D6B4D;">TOTAL</td>' +
          '<td style="padding:10px 12px;text-align:center;font-weight:700;">' + data.totalBoxes + ' boxes</td>' +
          '<td style="padding:10px 12px;text-align:right;font-weight:700;font-size:16px;color:#2D6B4D;">$' + data.total + '</td>' +
        '</tr></tfoot>' +
      '</table>' +

      // Mango care tips
      '<div style="background:#FFF8EC;border-radius:10px;padding:18px;margin:20px 0;border-left:4px solid #C8973A;">' +
        '<h3 style="margin:0 0 10px;color:#5A2800;">A few tips to enjoy your mangoes</h3>' +
        '<ul style="font-size:13px;line-height:1.7;padding-left:20px;margin:0;">' +
          '<li><strong>Naturally ripened</strong> — give green ones 2–3 days at room temperature</li>' +
          '<li><strong>Store ripe mangoes</strong> in the fridge to slow ripening</li>' +
          '<li><strong>Best at room temp</strong> for eating — cold dulls the flavor</li>' +
        '</ul>' +
      '</div>' +

      // Word of mouth
      '<div style="background:#F0FBF5;border:2px solid #2D6B4D;border-radius:10px;padding:18px;margin:18px 0;text-align:center;">' +
        '<h3 style="margin:0 0 8px;color:#2D6B4D;">Loved your mangoes?</h3>' +
        '<p style="font-size:13px;margin:0 0 10px;">A quick word from happy customers helps us bring more mangoes next season.</p>' +
        '<a href="' + CONFIG.WHATSAPP_LINK + '" style="display:inline-block;background:#25D366;color:#fff;padding:10px 22px;border-radius:6px;text-decoration:none;font-weight:600;margin:4px;">Share on WhatsApp</a>' +
      '</div>' +

      '<p style="font-size:14px;margin-top:20px;">Thank you again for being part of the TRU Mangoes family.<br/><strong>— The TRU Mangoes Team</strong></p>' +
    '</div>' +
    '<div style="background:#2A1500;color:#FFFBF1;padding:16px;text-align:center;font-size:11px;">trumangoes.shop · orders@trumangoes.shop · 972-654-9231<br/>© 2026 TRU Mangoes · VSVV International, LLC</div>' +
    '</div>';

  GmailApp.sendEmail(data.email, subject, '', {
    htmlBody: htmlBody,
    name: CONFIG.BUSINESS_NAME,
    replyTo: CONFIG.REPLY_TO,
    bcc: CONFIG.BCC_ALL_CUSTOMER_EMAILS
  });
}


// ═══════════════════════════════════════════════════════════════════
// 🌐 PART 2 — WEBSITE FORM WEBHOOK (doPost)
// Receives form submissions and creates a row in 📋 ORDERS.
// EmailJS sends both the customer confirmation and owner notification.
// This function just writes the row and assigns an Order ID.
// ═══════════════════════════════════════════════════════════════════

function doPost(e) {
  // ── Concurrency-safe order intake ──
  // A script lock serializes ONLY the Order-ID assignment + row write, so two
  // simultaneous orders can never grab the same Order ID. The lock is released
  // before the (slower) email send, so orders don't queue behind each other's
  // emails. This keeps intake fast even under a burst of submissions.
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    lock.waitLock(30000); // wait up to 30s
    locked = true;
  } catch (lockErr) {
    return jsonResponse_({status: 'error', message: 'Server busy — please tap submit again.'});
  }

  var orderId, customerName, data, bangan, kesar, rasalu, himayat, totalBoxes, totalAmount, serviceArea;
  try {
    data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.ORDERS_SHEET_NAME);
    if (!sheet) { lock.releaseLock(); return jsonResponse_({status: 'error', message: 'Orders sheet not found'}); }

    var nextRow = findNextEmptyRow_(sheet);
    orderId = getNextOrderId_(sheet);
    customerName = ((data.firstName || '') + ' ' + (data.lastName || '')).trim();
    serviceArea = mapServiceArea_(data.pickupLocation);

    bangan = Number(data.banganBoxes) || 0;
    kesar = Number(data.kesarBoxes) || 0;
    rasalu = Number(data.rasaluBoxes) || 0;
    himayat = Number(data.himayatBoxes) || 0;
    totalBoxes = Number(data.totalBoxes) || (bangan + kesar + rasalu + himayat);
    totalAmount = Number(data.grandTotal) || Number(data.boxTotal) ||
      (bangan * CONFIG.PRICES.banganapalli +
       kesar * CONFIG.PRICES.kesar +
       rasalu * CONFIG.PRICES.rasalu +
       himayat * CONFIG.PRICES.himayat);

    var notesParts = [];
    if (data.source) notesParts.push('Source: ' + data.source);
    if (data.notes) notesParts.push('Customer note: ' + data.notes);
    if (data.cityState) notesParts.push('City/State: ' + data.cityState);

    var rowData = [
      orderId,                         // A  Order ID
      new Date(),                      // B  Order Date
      customerName,                    // C  Customer Name
      data.phone || '',                // D  Phone
      data.phone || '',                // E  WhatsApp
      data.email || '',                // F  Email
      serviceArea,                     // G  City
      bangan,                          // H
      kesar,                           // I
      rasalu,                          // J
      himayat,                         // K
      totalBoxes,                      // L
      totalAmount,                     // M
      'Pending',                       // N  Payment Status
      '',                              // O  Payment Method
      '',                              // P  Payment Ref
      '',                              // Q  Payment Date
      'Not Picked Up',                 // R  Pickup Status
      serviceArea + ' Hub',            // S  Pickup Location
      notesParts.join(' | ')           // T  Notes
    ];
    sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
  } catch (err) {
    if (locked) { try { lock.releaseLock(); } catch (x) {} }
    return jsonResponse_({status: 'error', message: err.toString()});
  }

  // Critical section done — release the lock so other orders proceed.
  try { lock.releaseLock(); } catch (x) {}

  // ── Send ONE email: customer confirmation, BCC'd to the owner. ──
  // (The separate owner-notification email was redundant — the owner already
  //  receives every order via the BCC — so it was removed to halve email
  //  volume and speed up each submission.)
  try {
    sendOrderConfirmationEmail_({
      orderId: orderId, customerName: customerName, email: data.email, phone: data.phone,
      city: data.city || data.cityState || '',
      pickupLocation: data.pickupLocation || (serviceArea + ' Hub'),
      bangan: bangan, kesar: kesar, rasalu: rasalu, himayat: himayat,
      totalBoxes: totalBoxes, totalAmount: totalAmount,
      notes: data.notes || '', howHeard: data.source || ''
    });
    logEmail_(orderId, customerName, 'Order Confirmation', data.email, 'SENT');
  } catch (err) {
    // Email failure must NOT lose the order — the row is already saved.
    logEmail_(orderId, customerName, 'Order Confirmation', data.email, 'FAILED: ' + err.toString());
  }

  return jsonResponse_({status: 'success', orderId: orderId, message: 'Order received'});
}


// ═══════════════════════════════════════════════════════════════════
// 📧 ORDER CONFIRMATION EMAILS (sent from Apps Script via Gmail)
// These bypass EmailJS entirely. Called from doPost() on every new order.
// ═══════════════════════════════════════════════════════════════════

function buildItemsRows_(d) {
  // Returns HTML <tr> rows for the variety table (only varieties with qty > 0)
  var rows = '';
  var items = [
    { name: 'Banganapalli', qty: d.bangan,  price: CONFIG.PRICES.banganapalli },
    { name: 'Rasalu',       qty: d.rasalu,  price: CONFIG.PRICES.rasalu },
    { name: 'Himayat',      qty: d.himayat, price: CONFIG.PRICES.himayat },
    { name: 'Kesar',        qty: d.kesar,   price: CONFIG.PRICES.kesar }
  ];
  items.forEach(function(i){
    if (i.qty > 0) {
      var lt = i.qty * i.price;
      rows += '<tr>' +
        '<td style="padding:10px;border-bottom:1px solid #eee;">' + i.name + '</td>' +
        '<td style="padding:10px;border-bottom:1px solid #eee;text-align:center;">' + i.qty + ' box' + (i.qty > 1 ? 'es' : '') + '</td>' +
        '<td style="padding:10px;border-bottom:1px solid #eee;text-align:right;">$' + i.price + '</td>' +
        '<td style="padding:10px;border-bottom:1px solid #eee;text-align:right;font-weight:600;">$' + lt + '</td>' +
        '</tr>';
    }
  });
  return rows;
}

function sendOrderConfirmationEmail_(d) {
  var subject = 'Order #' + d.orderId + ' Received — TRU Mangoes Pre-Order';
  var itemsHtml = buildItemsRows_(d);

  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif;background:#f7f4ec;padding:20px 10px;color:#2a1500;">' +
      '<div style="max-width:600px;margin:0 auto;background:#ffffff;border-radius:10px;overflow:hidden;border:1px solid #e6dfd0;">' +

        // Header with order ID
        '<div style="background:#1F4A2C;padding:30px 20px;text-align:center;color:#ffffff;">' +
          '<div style="font-size:42px;line-height:1;margin-bottom:10px;">&#127821;</div>' +
          '<div style="font-family:Georgia,serif;font-size:26px;color:#F5C842;margin-bottom:14px;">Thank You, ' + d.customerName + '!</div>' +
          '<div style="background:#F5C842;color:#1F4A2C;display:inline-block;padding:10px 24px;border-radius:30px;font-size:20px;font-weight:bold;letter-spacing:.05em;">ORDER #' + d.orderId + '</div>' +
        '</div>' +

        '<div style="padding:24px;">' +
          '<p style="font-size:15px;line-height:1.6;margin:0 0 16px;">Your TRU Mangoes pre-order has been received for our <strong>May 29th batch</strong>. Please save your order number — you will need it when you pay.</p>' +

          // Order summary table
          '<h3 style="color:#1F4A2C;border-bottom:2px solid #C8973A;padding-bottom:6px;margin:24px 0 12px;font-size:16px;">Your Reservation</h3>' +
          '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="font-size:14px;border-collapse:collapse;">' +
            '<tr style="background:#f7f4ec;">' +
              '<th align="left" style="padding:10px;font-size:12px;color:#6b5b40;letter-spacing:.06em;">VARIETY</th>' +
              '<th align="center" style="padding:10px;font-size:12px;color:#6b5b40;letter-spacing:.06em;">QUANTITY</th>' +
              '<th align="right" style="padding:10px;font-size:12px;color:#6b5b40;letter-spacing:.06em;">PRICE/BOX</th>' +
              '<th align="right" style="padding:10px;font-size:12px;color:#6b5b40;letter-spacing:.06em;">SUBTOTAL</th>' +
            '</tr>' +
            itemsHtml +
            '<tr style="background:#f7f4ec;">' +
              '<td colspan="3" align="right" style="padding:14px 10px;font-weight:bold;">Total (' + d.totalBoxes + ' boxes)</td>' +
              '<td align="right" style="padding:14px 10px;font-weight:bold;color:#1F4A2C;font-size:18px;">$' + d.totalAmount + '</td>' +
            '</tr>' +
          '</table>' +

          // Payment
          '<h3 style="color:#1F4A2C;border-bottom:2px solid #C8973A;padding-bottom:6px;margin:24px 0 12px;font-size:16px;">Complete Your Payment (Zelle)</h3>' +
          '<div style="background:#FBF6ED;border:1.5px solid #C8973A;border-radius:8px;padding:16px;font-size:15px;line-height:1.8;">' +
            '<div><strong>Send to:</strong> 972-654-9231</div>' +
            '<div><strong>Recipient name:</strong> VSVV International</div>' +
            '<div><strong>Amount:</strong> $' + d.totalAmount + '</div>' +
          '</div>' +

          // BIG warning box
          '<div style="background:#FFF4D6;border:3px solid #C8973A;border-radius:8px;padding:20px;margin:18px 0;">' +
            '<div style="font-size:17px;font-weight:bold;color:#a02020;margin-bottom:14px;text-align:center;">⚠️ IMPORTANT — Read Before Paying</div>' +
            '<div style="font-size:14px;line-height:1.7;margin-bottom:12px;color:#2a1500;">In the Zelle <strong>memo / note field</strong>, please write exactly:</div>' +
            '<div style="background:#ffffff;border:2px dashed #a02020;padding:14px;border-radius:6px;font-family:Courier New,monospace;font-size:16px;text-align:center;font-weight:bold;color:#5A2800;margin-bottom:14px;">Order #' + d.orderId + ' — ' + d.customerName + '</div>' +
            '<div style="font-size:13px;line-height:1.7;color:#5A2800;">' +
              '<strong>This is REQUIRED so we can match your payment to your order.</strong><br><br>' +
              '<strong>Especially important if someone else (spouse, family member, friend) is sending the payment on your behalf</strong> — without the order number in the memo, we have no way to know who the payment is for and your reservation may be delayed or missed.' +
            '</div>' +
          '</div>' +

          '<p style="font-size:13px;line-height:1.6;margin:16px 0 8px;color:#666;font-style:italic;">Boxes are reserved on a first-paid basis. Please complete payment as soon as possible. Other payment options available — reply to this email or message us on WhatsApp.</p>' +

          // Pickup
          '<h3 style="color:#1F4A2C;border-bottom:2px solid #C8973A;padding-bottom:6px;margin:24px 0 12px;font-size:16px;">Pickup Details</h3>' +
          '<table cellpadding="4" cellspacing="0" border="0" style="font-size:14px;line-height:1.6;">' +
            '<tr><td width="100"><strong>Location:</strong></td><td>' + d.pickupLocation + '</td></tr>' +
            '<tr><td><strong>Date:</strong></td><td>May 29th, 2026</td></tr>' +
            '<tr><td valign="top"><strong>Note:</strong></td><td>Pickup only — no home delivery for this batch</td></tr>' +
          '</table>' +
          '<p style="font-size:13px;color:#666;margin:12px 0 0;">We will confirm the exact pickup time via WhatsApp or email once payment is received.</p>' +

          // WhatsApp
          '<div style="text-align:center;margin:28px 0;">' +
            '<p style="font-size:14px;margin:0 0 12px;color:#2a1500;">Join our WhatsApp group for batch updates and pickup notifications:</p>' +
            '<a href="https://chat.whatsapp.com/CUPFPvN3Du04kKiVKm0dVH" style="display:inline-block;background:#25D366;color:#ffffff;padding:14px 28px;border-radius:30px;text-decoration:none;font-weight:bold;font-size:15px;">&#128172; Join WhatsApp Group</a>' +
          '</div>' +

          // Order meta
          '<div style="background:#f7f4ec;padding:14px;border-radius:6px;font-size:12px;color:#6b5b40;line-height:1.7;margin:20px 0 0;">' +
            '<strong style="color:#1F4A2C;">Order #' + d.orderId + '</strong><br>' +
            'Received: ' + new Date().toLocaleString('en-US', {timeZone:'America/Chicago'}) + ' (Central Time)<br>' +
            'City: ' + d.city + '<br>' +
            'Notes: ' + (d.notes || 'None') +
          '</div>' +
        '</div>' +

        // Footer
        '<div style="background:#1F4A2C;padding:20px;text-align:center;color:#F5C842;font-size:13px;line-height:1.6;">' +
          '<strong>TRU Mangoes</strong> — A VSVV International, LLC brand<br>' +
          '<span style="color:#ffffff;opacity:.85;">Frisco, TX</span> · <a href="mailto:orders@trumangoes.shop" style="color:#F5C842;text-decoration:none;">orders@trumangoes.shop</a>' +
        '</div>' +

      '</div>' +
      '<p style="font-size:11px;color:#888;text-align:center;max-width:600px;margin:14px auto 0;">You received this email because you placed a pre-order at trumangoes.shop. Reply to this email with any questions and please include your order number #' + d.orderId + '.</p>' +
    '</div>';

  GmailApp.sendEmail(d.email, subject, '', {
    htmlBody: html,
    name: 'TRU Mangoes',
    replyTo: CONFIG.REPLY_TO,
    bcc: CONFIG.BCC_ALL_CUSTOMER_EMAILS
  });
}

function sendOwnerNotificationEmail_(d) {
  var subject = 'NEW ORDER #' + d.orderId + ' — ' + d.customerName + ' — ' + d.totalBoxes + ' boxes — $' + d.totalAmount;
  var itemsHtml = buildItemsRows_(d);

  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif;background:#f0f0f0;padding:20px 10px;color:#222;">' +
      '<div style="max-width:640px;margin:0 auto;background:#ffffff;border-radius:8px;overflow:hidden;border:1px solid #ddd;">' +

        '<div style="background:#1F4A2C;padding:20px 24px;color:#ffffff;">' +
          '<div style="font-size:22px;font-weight:bold;color:#F5C842;">&#127821; NEW ORDER #' + d.orderId + '</div>' +
          '<div style="font-size:15px;margin-top:8px;">' + d.customerName + ' — ' + d.totalBoxes + ' boxes — $' + d.totalAmount + '</div>' +
          '<div style="font-size:12px;opacity:.85;margin-top:4px;">' + new Date().toLocaleString('en-US', {timeZone:'America/Chicago'}) + ' CT</div>' +
        '</div>' +

        '<div style="padding:20px 24px;">' +

          '<h3 style="margin:0 0 10px;font-size:13px;color:#1F4A2C;text-transform:uppercase;letter-spacing:.06em;">Customer</h3>' +
          '<table cellpadding="4" cellspacing="0" border="0" style="font-size:14px;line-height:1.6;width:100%;">' +
            '<tr><td width="100"><strong>Name:</strong></td><td>' + d.customerName + '</td></tr>' +
            '<tr><td><strong>Email:</strong></td><td><a href="mailto:' + d.email + '" style="color:#1F4A2C;">' + d.email + '</a></td></tr>' +
            '<tr><td><strong>Phone:</strong></td><td><a href="tel:' + d.phone + '" style="color:#1F4A2C;">' + d.phone + '</a></td></tr>' +
            '<tr><td><strong>City:</strong></td><td>' + d.city + '</td></tr>' +
          '</table>' +

          '<h3 style="margin:24px 0 10px;font-size:13px;color:#1F4A2C;text-transform:uppercase;letter-spacing:.06em;">Order #' + d.orderId + '</h3>' +
          '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="font-size:14px;border-collapse:collapse;border:1px solid #eee;">' +
            '<tr style="background:#f7f4ec;">' +
              '<th align="left" style="padding:10px;font-size:12px;color:#555;">VARIETY</th>' +
              '<th align="center" style="padding:10px;font-size:12px;color:#555;">QTY</th>' +
              '<th align="right" style="padding:10px;font-size:12px;color:#555;">PRICE</th>' +
              '<th align="right" style="padding:10px;font-size:12px;color:#555;">SUBTOTAL</th>' +
            '</tr>' +
            itemsHtml +
            '<tr style="background:#f7f4ec;">' +
              '<td colspan="3" align="right" style="padding:12px 10px;font-weight:bold;">Total (' + d.totalBoxes + ' boxes)</td>' +
              '<td align="right" style="padding:12px 10px;font-weight:bold;color:#1F4A2C;font-size:16px;">$' + d.totalAmount + '</td>' +
            '</tr>' +
          '</table>' +

          '<h3 style="margin:24px 0 10px;font-size:13px;color:#1F4A2C;text-transform:uppercase;letter-spacing:.06em;">Pickup &amp; Source</h3>' +
          '<table cellpadding="4" cellspacing="0" border="0" style="font-size:14px;line-height:1.6;">' +
            '<tr><td width="100"><strong>Pickup:</strong></td><td>' + d.pickupLocation + '</td></tr>' +
            '<tr><td><strong>Heard via:</strong></td><td>' + (d.howHeard || 'Not specified') + '</td></tr>' +
            '<tr><td valign="top"><strong>Notes:</strong></td><td>' + (d.notes || 'None') + '</td></tr>' +
          '</table>' +

          '<div style="margin:24px 0 8px;text-align:center;">' +
            '<a href="https://wa.me/1' + d.phone.replace(/[^0-9]/g,'') + '" style="display:inline-block;background:#25D366;color:#ffffff;padding:10px 18px;border-radius:6px;text-decoration:none;font-size:13px;margin:4px;font-weight:bold;">WhatsApp</a>' +
            '<a href="mailto:' + d.email + '" style="display:inline-block;background:#1F4A2C;color:#ffffff;padding:10px 18px;border-radius:6px;text-decoration:none;font-size:13px;margin:4px;font-weight:bold;">Email</a>' +
            '<a href="tel:' + d.phone + '" style="display:inline-block;background:#C8973A;color:#ffffff;padding:10px 18px;border-radius:6px;text-decoration:none;font-size:13px;margin:4px;font-weight:bold;">Call</a>' +
          '</div>' +

          '<div style="background:#FFF4D6;border:2px solid #C8973A;border-radius:6px;padding:14px;margin-top:16px;font-size:13px;color:#5A2800;line-height:1.6;">' +
            '<strong>⚠️ Watch for Zelle payment with memo:</strong><br>' +
            '<span style="font-family:Courier New,monospace;font-size:14px;font-weight:bold;">Order #' + d.orderId + ' — ' + d.customerName + '</span>' +
          '</div>' +

        '</div>' +

        '<div style="background:#222;padding:12px;text-align:center;color:#aaa;font-size:11px;">' +
          'TRU Mangoes Order Notification · orders@trumangoes.shop' +
        '</div>' +

      '</div>' +
    '</div>';

  GmailApp.sendEmail(CONFIG.OWNER_EMAIL, subject, '', {
    htmlBody: html,
    name: 'TRU Mangoes Orders',
    replyTo: CONFIG.REPLY_TO
  });
}

function findNextEmptyRow_(sheet) {
  const data = sheet.getRange('C3:C506').getValues();
  for (let i = 0; i < data.length; i++) {
    if (!data[i][0] || data[i][0].toString().trim() === '') return i + 3;
  }
  return sheet.getLastRow() + 1;
}

function getNextOrderId_(sheet) {
  const data = sheet.getRange('A3:A506').getValues();
  let maxId = 0;
  for (let i = 0; i < data.length; i++) {
    const val = data[i][0];
    if (val && !isNaN(val)) {
      const num = parseInt(val, 10);
      if (num > maxId) maxId = num;
    }
  }
  // Floor the next ID at CONFIG.STARTING_ORDER_ID (e.g., 1001 for May 23 launch).
  // Once orders exist at or above that floor, continue from max+1 naturally.
  const next = maxId + 1;
  return next < CONFIG.STARTING_ORDER_ID ? CONFIG.STARTING_ORDER_ID : next;
}

function mapServiceArea_(formValue) {
  if (!formValue) return 'Out of Service Area';
  const v = formValue.toString().trim();
  if (v === 'Frisco') return 'Frisco';
  if (v === 'Plano') return 'Plano';
  if (v === 'Irving/Coppell') return 'Irving/Coppell';
  if (v === 'Celina/Prosper' || v === 'Prosper/Celina') return 'Prosper/Celina';
  return 'Out of Service Area';
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}


// ═══════════════════════════════════════════════════════════════════
// 🛠️ UTILITY FUNCTIONS
// ═══════════════════════════════════════════════════════════════════

function columnNumberToLetter_(num) {
  let letter = '';
  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }
  return letter;
}

function letterToColumnNumber_(letter) {
  letter = letter.toUpperCase();
  let num = 0;
  for (let i = 0; i < letter.length; i++) {
    num = num * 26 + (letter.charCodeAt(i) - 64);
  }
  return num;
}

function formatDate_(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  try { return Utilities.formatDate(date, Session.getScriptTimeZone(), 'M/d/yyyy'); }
  catch (e) { return date.toString(); }
}

function formatTimestamp_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'M/d/yyyy h:mm a');
}

function logEmail_(orderId, customer, emailType, recipient, status) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('📜 EMAIL LOG');
    if (!logSheet) {
      createAuditLogSheet_(ss);
      logSheet = ss.getSheetByName('📜 EMAIL LOG');
    }
    logSheet.appendRow([new Date(), orderId || '', customer || '', emailType || '', recipient || '', status || '']);
  } catch (e) {
    console.error('Failed to log email:', e);
  }
}


// ═══════════════════════════════════════════════════════════════════
// 🧪 TEST FUNCTIONS — Run manually before going live
// ═══════════════════════════════════════════════════════════════════

// Send sample Payment Confirmed + Thank You emails to yourself, with mixed varieties.
function sendTestEmail() {
  const testData = {
    orderId: 'TEST-001',
    customerName: 'Venkata Mandala',
    firstName: 'Venkata',
    email: CONFIG.OWNER_EMAIL,
    banganBoxes: 2,
    kesarBoxes: 0,
    rasaluBoxes: 1,
    himayatBoxes: 1,
    totalBoxes: 4,
    total: 2 * 47 + 1 * 52 + 1 * 55, // = 201
    city: 'Frisco',
    pickupLocation: 'Frisco Hub'
  };
  sendPaymentConfirmedEmail_(testData);
  sendThankYouEmail_(testData);
  SpreadsheetApp.getUi().alert('Test emails sent to ' + CONFIG.OWNER_EMAIL + ' (and BCC).');
}

// Simulate a website form submission. Creates a real test row — delete it after.
function testWebhookLocal() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        firstName: 'Test', lastName: 'WebhookOrder',
        email: CONFIG.OWNER_EMAIL, phone: '9999999999',
        city: 'Frisco', state: 'TX', cityState: 'Frisco, TX',
        pickupLocation: 'Frisco',
        banganBoxes: 2, kesarBoxes: 0, rasaluBoxes: 1, himayatBoxes: 0,
        totalBoxes: 3,
        boxTotal: 2 * 47 + 1 * 52, // 146
        grandTotal: 146,
        source: 'Test', notes: 'This is a test webhook submission — delete after.'
      })
    }
  };
  const result = doPost(fakeEvent);
  Logger.log('Webhook test result: ' + result.getContent());
  SpreadsheetApp.getUi().alert('Test submission complete. Check the ORDERS tab for the new test row, then delete it before launch.');
}


// ═══════════════════════════════════════════════════════════════════
// 🚀 LAUNCH PREP — One-click sheet cleanup
// Run this once after pasting Code.gs and running setup(). It wipes all
// test data and rebuilds the helper tabs to use the new v2 schema.
// ═══════════════════════════════════════════════════════════════════

function prepareForLaunch() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Prepare TRU Mangoes Sheet for Launch',
    'This will:\n\n' +
    '1. WIPE all data rows in 📋 ORDERS (header stays).\n' +
    '2. CLEAR all entries in 📜 EMAIL LOG (header stays).\n' +
    '3. REBUILD these helper tabs with new schema formulas:\n' +
    '       🔍 SEARCH\n' +
    '       📦 TODAY\'S PICKUPS\n' +
    '       💰 PAYMENTS PENDING\n' +
    '       👥 CUSTOMERS\n' +
    '       &#127821; BATCHES\n' +
    '       📊 DASHBOARD\n\n' +
    '📋 INSTRUCTIONS tab is NOT touched.\n' +
    'Apps Script triggers and Web App deployment are preserved.\n\n' +
    'First real order will be #' + CONFIG.STARTING_ORDER_ID + '.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (resp !== ui.Button.YES) {
    ui.alert('Cancelled. No changes made.');
    return;
  }

  const log = [];
  try {
    log.push(cleanOrdersData_());
    log.push(cleanEmailLog_());
    log.push(rebuildSearchTab_());
    log.push(rebuildTodaysPickupsTab_());
    log.push(rebuildPaymentsPendingTab_());
    log.push(rebuildCustomersTab_());
    log.push(rebuildBatchesTab_());
    log.push(rebuildDashboardTab_());

    ui.alert(
      '&#127821; Ready for Launch!',
      log.join('\n') + '\n\n' +
      'First real order will be Order #' + CONFIG.STARTING_ORDER_ID + '.\n\n' +
      'Next steps:\n' +
      '  • Run sendTestEmail() to verify email delivery\n' +
      '  • Run testWebhookLocal() to verify sheet write\n' +
      '  • Delete the test row before Saturday\n' +
      '  • Saturday morning: flip SOLD_OUT flags in index.html',
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('Error during launch prep', err.toString(), ui.ButtonSet.OK);
  }
}

function cleanOrdersData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.ORDERS_SHEET_NAME);
  if (!sheet) return '⚠️  ORDERS tab not found';

  const lastRow = sheet.getLastRow();
  if (lastRow >= 3) {
    sheet.getRange(3, 1, lastRow - 2, sheet.getMaxColumns()).clearContent();
    return '✓ Wiped ' + (lastRow - 2) + ' data row(s) from 📋 ORDERS';
  }
  return '✓ 📋 ORDERS already clean';
}

function cleanEmailLog_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📜 EMAIL LOG');
  if (!sheet) return '⚠️  EMAIL LOG tab not found';

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getMaxColumns()).clearContent();
    return '✓ Cleared ' + (lastRow - 1) + ' entries from 📜 EMAIL LOG';
  }
  return '✓ 📜 EMAIL LOG already clean';
}

function getOrInsertSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clear();
  return sheet;
}

function rebuildSearchTab_() {
  const sheet = getOrInsertSheet_('🔍 SEARCH');
  sheet.getRange('A1').setValue('Search:').setFontWeight('bold').setFontSize(14);
  sheet.getRange('B1').setBackground('#FFF8EC').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A2').setValue('Type a name, phone, email, or city in cell B1 to find matching orders.')
    .setFontStyle('italic').setFontColor('#666');
  sheet.getRange('A4').setFormula(
    '=IFERROR(QUERY(\'📋 ORDERS\'!A2:X, ' +
    '"SELECT * WHERE LOWER(C) CONTAINS LOWER(\'"&B1&"\') ' +
    'OR LOWER(D) CONTAINS LOWER(\'"&B1&"\') ' +
    'OR LOWER(F) CONTAINS LOWER(\'"&B1&"\') ' +
    'OR LOWER(G) CONTAINS LOWER(\'"&B1&"\')", 1), ' +
    '"No matches.")'
  );
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 250);
  return '✓ Rebuilt 🔍 SEARCH';
}

function rebuildTodaysPickupsTab_() {
  const sheet = getOrInsertSheet_('📦 TODAY\'S PICKUPS');
  sheet.getRange('A1').setValue('Paid orders scheduled or ready for pickup').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A3').setFormula(
    '=IFERROR(QUERY(\'📋 ORDERS\'!A2:X, ' +
    '"SELECT A,C,D,E,L,M,R,S ' +
    'WHERE (R=\'Scheduled\' OR R=\'Ready for Pickup\') AND N=\'Paid\' ' +
    'ORDER BY S", 1), ' +
    '"No pickups scheduled.")'
  );
  return '✓ Rebuilt 📦 TODAY\'S PICKUPS';
}

function rebuildPaymentsPendingTab_() {
  const sheet = getOrInsertSheet_('💰 PAYMENTS PENDING');
  sheet.getRange('A1').setValue('Orders awaiting payment').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A3').setFormula(
    '=IFERROR(QUERY(\'📋 ORDERS\'!A2:X, ' +
    '"SELECT A,B,C,D,F,L,M,N ' +
    'WHERE A IS NOT NULL AND N!=\'Paid\' ' +
    'ORDER BY B DESC", 1), ' +
    '"No pending payments.")'
  );
  return '✓ Rebuilt 💰 PAYMENTS PENDING';
}

function rebuildCustomersTab_() {
  const sheet = getOrInsertSheet_('👥 CUSTOMERS');
  sheet.getRange('A1').setValue('Customer history (sorted by $ spent)').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A3').setFormula(
    '=IFERROR(QUERY(\'📋 ORDERS\'!A3:X, ' +
    '"SELECT C, D, F, G, COUNT(A), SUM(L), SUM(M) ' +
    'WHERE A IS NOT NULL ' +
    'GROUP BY C, D, F, G ' +
    'ORDER BY SUM(M) DESC ' +
    'LABEL C \'Customer\', D \'Phone\', F \'Email\', G \'City\', ' +
    'COUNT(A) \'Orders\', SUM(L) \'Total Boxes\', SUM(M) \'$ Spent\'"), ' +
    '"No customers yet.")'
  );
  return '✓ Rebuilt 👥 CUSTOMERS';
}

function rebuildBatchesTab_() {
  const sheet = getOrInsertSheet_('&#127821; BATCHES');

  const headers = ['Variety', 'Unit $', 'Boxes Ordered', 'Boxes Paid', 'Boxes Picked Up', 'Revenue'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#2D6B4D').setFontColor('#FFFFFF');

  const rows = [
    ['Banganapalli',  CONFIG.PRICES.banganapalli,
      "=SUM('📋 ORDERS'!H3:H)",
      "=SUMIFS('📋 ORDERS'!H3:H, '📋 ORDERS'!N3:N, \"Paid\")",
      "=SUMIFS('📋 ORDERS'!H3:H, '📋 ORDERS'!R3:R, \"Picked Up\")",
      "=C2*B2"],
    ['Kesar',  CONFIG.PRICES.kesar,
      "=SUM('📋 ORDERS'!I3:I)",
      "=SUMIFS('📋 ORDERS'!I3:I, '📋 ORDERS'!N3:N, \"Paid\")",
      "=SUMIFS('📋 ORDERS'!I3:I, '📋 ORDERS'!R3:R, \"Picked Up\")",
      "=C3*B3"],
    ['Pedda Rasalu',  CONFIG.PRICES.rasalu,
      "=SUM('📋 ORDERS'!J3:J)",
      "=SUMIFS('📋 ORDERS'!J3:J, '📋 ORDERS'!N3:N, \"Paid\")",
      "=SUMIFS('📋 ORDERS'!J3:J, '📋 ORDERS'!R3:R, \"Picked Up\")",
      "=C4*B4"],
    ['Himayat',  CONFIG.PRICES.himayat,
      "=SUM('📋 ORDERS'!K3:K)",
      "=SUMIFS('📋 ORDERS'!K3:K, '📋 ORDERS'!N3:N, \"Paid\")",
      "=SUMIFS('📋 ORDERS'!K3:K, '📋 ORDERS'!R3:R, \"Picked Up\")",
      "=C5*B5"]
  ];
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  const totalRow = 2 + rows.length + 1; // row 7
  sheet.getRange(totalRow, 1, 1, headers.length).setValues([[
    'TOTAL', '', '=SUM(C2:C5)', '=SUM(D2:D5)', '=SUM(E2:E5)', '=SUM(F2:F5)'
  ]]).setFontWeight('bold').setBackground('#FFF8EC');

  sheet.getRange('B2:B5').setNumberFormat('$#,##0');
  sheet.getRange('F2:F7').setNumberFormat('$#,##0');
  sheet.setColumnWidth(1, 140);
  return '✓ Rebuilt &#127821; BATCHES';
}

function rebuildDashboardTab_() {
  const sheet = getOrInsertSheet_('📊 DASHBOARD');

  const rows = [
    ['Metric', 'Value'],

    // Order totals
    ['📦 Total Orders',         "=COUNTA('📋 ORDERS'!A3:A)"],
    ['📦 Total Boxes (All Varieties)', "=SUM('📋 ORDERS'!L3:L)"],
    ['💰 Total Revenue',        "=SUM('📋 ORDERS'!M3:M)"],
    ['', ''],

    // Per-variety boxes sold
    ['🥭 Banganapalli Boxes Sold', "=SUM('📋 ORDERS'!H3:H)"],
    ['🥭 Kesar Boxes Sold',        "=SUM('📋 ORDERS'!I3:I)"],
    ['🥭 Rasalu Boxes Sold',       "=SUM('📋 ORDERS'!J3:J)"],
    ['🥭 Himayat Boxes Sold',      "=SUM('📋 ORDERS'!K3:K)"],
    ['🥭 Total Boxes (Cross-Check)', "=H+I+J+K"],  // placeholder, will fix below
    ['', ''],

    // Per-variety revenue
    ['💵 Banganapalli Revenue', "=SUM('📋 ORDERS'!H3:H) * " + CONFIG.PRICES.banganapalli],
    ['💵 Kesar Revenue',        "=SUM('📋 ORDERS'!I3:I) * " + CONFIG.PRICES.kesar],
    ['💵 Rasalu Revenue',       "=SUM('📋 ORDERS'!J3:J) * " + CONFIG.PRICES.rasalu],
    ['💵 Himayat Revenue',      "=SUM('📋 ORDERS'!K3:K) * " + CONFIG.PRICES.himayat],
    ['', ''],

    // Payment status
    ['✅ Payments Paid',        "=COUNTIF('📋 ORDERS'!N3:N, \"Paid\")"],
    ['⏳ Payments Pending',     "=COUNTA('📋 ORDERS'!A3:A) - COUNTIF('📋 ORDERS'!N3:N, \"Paid\")"],
    ['💰 Revenue Collected',    "=SUMIFS('📋 ORDERS'!M3:M, '📋 ORDERS'!N3:N, \"Paid\")"],
    ['💰 Revenue Outstanding',  "=SUM('📋 ORDERS'!M3:M) - SUMIFS('📋 ORDERS'!M3:M, '📋 ORDERS'!N3:N, \"Paid\")"],
    ['', ''],

    // Pickup status
    ['🚚 Picked Up',            "=COUNTIF('📋 ORDERS'!R3:R, \"Picked Up\")"],
    ['📍 Awaiting Pickup',      "=COUNTIF('📋 ORDERS'!N3:N, \"Paid\") - COUNTIF('📋 ORDERS'!R3:R, \"Picked Up\")"],
    ['', ''],

    // Latest activity
    ['🆔 Latest Order ID',      "=IFERROR(MAX('📋 ORDERS'!A3:A), \"—\")"],
    ['📅 Latest Order Date',    "=IFERROR(MAX('📋 ORDERS'!B3:B), \"—\")"]
  ];

  // Fix the cross-check formula (couldn't reference cells inside the array)
  rows[9][1] = "=SUM('📋 ORDERS'!H3:H) + SUM('📋 ORDERS'!I3:I) + SUM('📋 ORDERS'!J3:J) + SUM('📋 ORDERS'!K3:K)";

  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);

  // Header styling
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#2D6B4D').setFontColor('#FFFFFF').setFontSize(12);

  // Section header styling — rows starting each section
  [2, 6, 12, 17, 22, 25].forEach(function(r){
    if (r <= rows.length) {
      sheet.getRange(r, 1, 1, 2).setFontWeight('bold');
    }
  });

  // Number formatting
  // Revenue rows (B4, B12-B15, B19, B20) as currency
  sheet.getRange('B4').setNumberFormat('$#,##0');
  sheet.getRange('B12:B15').setNumberFormat('$#,##0');
  sheet.getRange('B19:B20').setNumberFormat('$#,##0');
  // Latest order date
  sheet.getRange('B' + rows.length).setNumberFormat('M/d/yyyy h:mm a');

  // Light alternating background for readability
  for (var i = 2; i <= rows.length; i++) {
    if (rows[i-1][0] !== '') {
      sheet.getRange(i, 1, 1, 2).setBorder(false, false, true, false, false, false, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 200);
  return '✓ Rebuilt 📊 DASHBOARD with per-variety breakdown';
}


// ═══════════════════════════════════════════════════════════════════
// 🔒 CELL LOCKING — Auto-protect Payment Received & Picked Up after check
//
// Once someone ticks the Payment Received (col U) or Picked Up (col W)
// checkbox, this trigger automatically protects that cell so only the
// admin emails in CONFIG.ADMIN_EMAILS can untick it.
//
// Run installCellLockTrigger() ONCE to set up the trigger.
// Visual indicator: locked cells get a dark grey background.
// ═══════════════════════════════════════════════════════════════════

function onEditInstallable(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== CONFIG.ORDERS_SHEET_NAME) return;

    var col = e.range.getColumn();
    var row = e.range.getRow();
    if (row < 3) return; // skip header rows

    // Column U = 21 (Payment Received), Column W = 23 (Picked Up)
    if (col !== 21 && col !== 23) return;

    // Only lock when the cell is newly checked TRUE
    // (don't lock if someone is unchecking)
    var val = e.range.getValue();
    if (val !== true) return;

    // Protect this single cell
    var protection = e.range.protect();
    var actionType = (col === 21 ? 'Payment Received' : 'Picked Up');
    protection.setDescription('Auto-locked: ' + actionType + ' for row ' + row + ' — only admins can override');

    // Remove all existing editors except the sheet owner
    var editors = protection.getEditors();
    var ownerEmail = Session.getEffectiveUser().getEmail();
    editors.forEach(function(ed) {
      try {
        var em = ed.getEmail();
        if (em && em !== ownerEmail) {
          protection.removeEditor(em);
        }
      } catch (innerErr) { /* ignore individual editor removal failures */ }
    });
    if (protection.canDomainEdit()) protection.setDomainEdit(false);

    // Add admin emails back as editors (they can override)
    if (CONFIG.ADMIN_EMAILS && CONFIG.ADMIN_EMAILS.length > 0) {
      var admins = CONFIG.ADMIN_EMAILS.filter(function(em){
        return em && typeof em === 'string' && em.indexOf('@') > 0;
      });
      if (admins.length > 0) {
        try { protection.addEditors(admins); } catch (e2) { /* invalid emails ignored */ }
      }
    }

    // Visual: grey out the cell so it looks locked
    e.range.setBackground('#cfd8dc');
    e.range.setFontColor('#37474f');

  } catch (err) {
    console.error('onEditInstallable error:', err.toString());
  }
}

function installCellLockTrigger() {
  // Remove any existing installable onEdit trigger for this function
  var existing = ScriptApp.getProjectTriggers();
  var removed = 0;
  existing.forEach(function(t) {
    if (t.getHandlerFunction() === 'onEditInstallable') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });

  // Create a fresh trigger
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Cell-Lock Trigger Installed\n\n' +
    'Removed ' + removed + ' old trigger(s) and created a fresh one.\n\n' +
    'Behavior:\n' +
    '• Tick Payment Received (col U) → cell becomes grey and locked\n' +
    '• Tick Picked Up (col W) → cell becomes grey and locked\n' +
    '• Only emails in CONFIG.ADMIN_EMAILS can override locked cells\n\n' +
    'Current admin emails:\n' +
    (CONFIG.ADMIN_EMAILS && CONFIG.ADMIN_EMAILS.length > 0
      ? CONFIG.ADMIN_EMAILS.join('\n')
      : '⚠️ NONE configured — add yours to CONFIG.ADMIN_EMAILS in Code.gs')
  );
}

function uninstallCellLockTrigger() {
  // For emergency use if behavior misfires
  var existing = ScriptApp.getProjectTriggers();
  var removed = 0;
  existing.forEach(function(t) {
    if (t.getHandlerFunction() === 'onEditInstallable') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  SpreadsheetApp.getUi().alert('Removed ' + removed + ' cell-lock trigger(s). New checks will NOT auto-lock until you reinstall.');
}

function unlockAllProtectedCells() {
  // Admin utility: clears all auto-locks on the ORDERS sheet
  // (use if you want to bulk-edit historical data)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.ORDERS_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Orders sheet not found');
    return;
  }
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  var unlocked = 0;
  protections.forEach(function(p) {
    var desc = p.getDescription() || '';
    if (desc.indexOf('Auto-locked') === 0) {
      // Also reset the visual background
      var range = p.getRange();
      range.setBackground(null);
      range.setFontColor(null);
      p.remove();
      unlocked++;
    }
  });
  SpreadsheetApp.getUi().alert('Unlocked ' + unlocked + ' auto-locked cell(s) on the ORDERS sheet.');
}


// ═══════════════════════════════════════════════════════════════════
// 🌐 ADMIN GUI API — serves data to trumangoes.shop/admin
// ═══════════════════════════════════════════════════════════════════
// All requests come as GET with ?action=X&pw=Y parameters.
// Password identifies role: admin (full) vs partner (limited).
// Partners can MARK things done but cannot UNDO them.
// Admins can override anything.

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = params.action || '';
  var username = params.user || '';
  var pw = params.pw || '';
  var orderId = params.orderId || '';

  // Identify role by verifying username + password against hashed credentials
  var role = verifyUser_(username, pw);

  if (!role) {
    return jsonResponse_({status: 'error', code: 'unauthorized', message: 'Invalid username or password'});
  }

  // Determine if this role can override/undo locked actions (admin only)
  var canOverride = (role === 'admin');

  try {
    switch (action) {
      case 'verify':
        return jsonResponse_({status: 'ok', role: role, canOverride: canOverride, username: username});

      case 'dashboard':
        return jsonResponse_({status: 'ok', data: getDashboardData_()});

      case 'orders':
        return jsonResponse_({status: 'ok', data: getOrdersData_('all')});

      case 'pendingPayments':
        return jsonResponse_({status: 'ok', data: getOrdersData_('pending')});

      case 'paidPayments':
        return jsonResponse_({status: 'ok', data: getOrdersData_('paid')});

      case 'pickups':
        return jsonResponse_({status: 'ok', data: getOrdersData_('pickups')});

      case 'pickedUp':
        return jsonResponse_({status: 'ok', data: getOrdersData_('pickedup')});

      case 'markPaymentReceived':
        return jsonResponse_(markPaymentReceived_(orderId, role, params.method || ''));

      case 'markPickedUp':
        return jsonResponse_(markPickedUp_(orderId, role));

      case 'unmarkPaymentReceived':
        if (!canOverride) return jsonResponse_({status: 'error', code: 'admin_only', message: 'Admin override required'});
        return jsonResponse_(unmarkAction_(orderId, 'payment'));

      case 'unmarkPickedUp':
        if (!canOverride) return jsonResponse_({status: 'error', code: 'admin_only', message: 'Admin override required'});
        return jsonResponse_(unmarkAction_(orderId, 'pickup'));

      default:
        return jsonResponse_({status: 'error', message: 'Unknown action: ' + action});
    }
  } catch (err) {
    return jsonResponse_({status: 'error', message: err.toString()});
  }
}

// ───────────────────────────────────────────────────────────────────
// 🔐 USER AUTHENTICATION (hashed passwords)
// ───────────────────────────────────────────────────────────────────

function hashPassword_(password) {
  // SHA-256 of (salt + password), returned as lowercase hex
  var raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    CONFIG.PASSWORD_SALT + password,
    Utilities.Charset.UTF_8
  );
  return raw.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

function verifyUser_(username, password) {
  if (!username || !password) return null;
  var uname = String(username).trim().toLowerCase();
  var hash = hashPassword_(password);
  for (var i = 0; i < CONFIG.USERS.length; i++) {
    var u = CONFIG.USERS[i];
    if (u.username.toLowerCase() === uname && u.passwordHash === hash) {
      return u.role;
    }
  }
  return null;
}

// Run this from the Apps Script editor to generate a hash for a new password.
// It shows a popup with the hash to copy into CONFIG.USERS.
function generatePasswordHash() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    '🔐 Generate Password Hash',
    'Enter the new password you want to hash:',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var pw = resp.getResponseText();
  if (!pw) { ui.alert('No password entered.'); return; }
  var hash = hashPassword_(pw);
  ui.alert(
    'Password Hash Generated',
    'Password: ' + pw + '\n\n' +
    'Hash (copy this into CONFIG.USERS passwordHash):\n\n' + hash + '\n\n' +
    'Example entry:\n' +
    "{ username: 'newuser', passwordHash: '" + hash + "', role: 'partner' },",
    ui.ButtonSet.OK
  );
}

// ───────────────────────────────────────────────────────────────────
// 🥭 SPREADSHEET MENU (appears at the top of the sheet after refresh)
// ───────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🥭 TRU Mangoes')
    .addItem('🔧 Repair sheet layout (clean)', 'repairSheet')
    .addItem('🧹 Remove extra columns', 'cleanupExtraColumns')
    .addItem('💳 Set Payment Method dropdown (Zelle/Cash)', 'setupPaymentMethodDropdown')
    .addSeparator()
    .addItem('🔐 Generate password hash', 'generatePasswordHash')
    .addToUi();
}

// One-click clean repair: removes extra columns, fixes the 4 automation
// headers (U–X), and re-applies the Zelle/Cash dropdown. Safe to run anytime.
function repairSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.ORDERS_SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert('Orders sheet not found.'); return; }
  var msgs = [];

  // 1. Remove any columns past X (24)
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 24) {
    sheet.deleteColumns(25, maxCols - 24);
    msgs.push('Removed ' + (maxCols - 24) + ' extra column(s) past X.');
  }

  // 2. Re-write the 4 automation headers in U–X so they are correct & unique
  var autoHeaders = ['Payment Received \u2705', 'Payment Email Sent', 'Picked Up \u2705', 'Pickup Email Sent'];
  var startCol = letterToColumnNumber_(CONFIG.COL.PAYMENT_RECEIVED); // U = 21
  autoHeaders.forEach(function(h, i) {
    sheet.getRange(CONFIG.HEADER_ROW, startCol + i).setValue(h)
      .setFontWeight('bold').setBackground('#2D5240').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  });
  msgs.push('Verified automation headers in columns U\u2013X.');

  // 3. Re-apply Zelle/Cash dropdown on Payment Method (O)
  try {
    var col = letterToColumnNumber_(CONFIG.COL.PAYMENT_METHOD);
    var lastRow = Math.max(sheet.getMaxRows(), CONFIG.DATA_START_ROW);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Zelle', 'Cash'], true).setAllowInvalid(false).build();
    sheet.getRange(CONFIG.DATA_START_ROW, col, lastRow - CONFIG.DATA_START_ROW + 1, 1).setDataValidation(rule);
    msgs.push('Applied Zelle/Cash dropdown to Payment Method.');
  } catch (e) {}

  SpreadsheetApp.getUi().alert('\u2705 Sheet repaired:\n\n\u2022 ' + msgs.join('\n\u2022 '));
}



// Removes any duplicate/extra columns beyond column X (24). Your sheet has
// a second "Picked Up ✅" and "Pickup Email Sent" in columns Y & Z — this
// deletes everything from column 25 onward so the layout is clean (A–X only).
function cleanupExtraColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.ORDERS_SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert('Orders sheet not found.'); return; }
  var maxCols = sheet.getMaxColumns();
  var keep = 24; // A–X
  if (maxCols > keep) {
    sheet.deleteColumns(keep + 1, maxCols - keep);
    SpreadsheetApp.getUi().alert('✅ Removed ' + (maxCols - keep) + ' extra column(s). Sheet now ends at column X.');
  } else {
    SpreadsheetApp.getUi().alert('Nothing to remove — sheet already has ' + maxCols + ' columns.');
  }
}

// Sets the Payment Method column (O) to a dropdown of Zelle / Cash.
// "Cash" is for walk-ins who pay cash for boxes left out for pickup.
function setupPaymentMethodDropdown() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.ORDERS_SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert('Orders sheet not found.'); return; }
  var col = letterToColumnNumber_(CONFIG.COL.PAYMENT_METHOD); // O = 15
  var lastRow = Math.max(sheet.getMaxRows(), CONFIG.DATA_START_ROW);
  var numRows = lastRow - CONFIG.DATA_START_ROW + 1;
  var range = sheet.getRange(CONFIG.DATA_START_ROW, col, numRows, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Zelle', 'Cash'], true)
    .setAllowInvalid(false)
    .setHelpText('Choose Zelle or Cash (Cash = walk-in / box left out).')
    .build();
  range.setDataValidation(rule);
  SpreadsheetApp.getUi().alert('✅ Payment Method (column O) now has a Zelle / Cash dropdown for all order rows.');
}

function getDashboardData_() {
  // Derive ALL stats from the same getOrdersData_ source the list tabs use,
  // so the dashboard and the lists can never disagree. Paid/picked status
  // comes from the checkbox columns (source of truth).
  var orders = getOrdersData_('all');
  var stats = {
    totalOrders: 0, totalBoxes: 0, totalRevenue: 0,
    paid: 0, pending: 0, revenueCollected: 0,
    pickedUp: 0, awaitingPickup: 0,
    varieties: {banganapalli:0, kesar:0, rasalu:0, himayat:0},
    varietyRevenue: {banganapalli:0, kesar:0, rasalu:0, himayat:0}
  };
  orders.forEach(function(o) {
    stats.totalOrders++;
    stats.varieties.banganapalli += Number(o.bangan)  || 0;
    stats.varieties.kesar        += Number(o.kesar)   || 0;
    stats.varieties.rasalu       += Number(o.rasalu)  || 0;
    stats.varieties.himayat      += Number(o.himayat) || 0;
    stats.totalBoxes   += Number(o.totalBoxes)  || 0;
    stats.totalRevenue += Number(o.totalAmount) || 0;
    if (o.paymentReceived) {
      stats.paid++;
      stats.revenueCollected += Number(o.totalAmount) || 0;
      if (o.pickedUp) stats.pickedUp++;
      else stats.awaitingPickup++;
    } else {
      stats.pending++;
    }
  });
  stats.varietyRevenue.banganapalli = stats.varieties.banganapalli * CONFIG.PRICES.banganapalli;
  stats.varietyRevenue.kesar        = stats.varieties.kesar        * CONFIG.PRICES.kesar;
  stats.varietyRevenue.rasalu       = stats.varieties.rasalu       * CONFIG.PRICES.rasalu;
  stats.varietyRevenue.himayat      = stats.varieties.himayat      * CONFIG.PRICES.himayat;
  return stats;
}

function getOrdersData_(filter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.ORDERS_SHEET_NAME);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  // Defensive: read only as many columns as actually exist (up to 24).
  // This prevents "range exceeds grid" errors if the sheet is narrower,
  // and ignores any duplicate/extra columns beyond X.
  var maxCols = sheet.getLastColumn();
  var numCols = Math.min(24, maxCols);
  if (numCols < 13) return [];
  var numRows = lastRow - 2;
  var data = sheet.getRange(3, 1, numRows, numCols).getValues();
  var rows = [];
  data.forEach(function(r, idx) {
    if (r[0] === '' || r[0] === null || r[0] === undefined) return; // skip empty rows
    // Pad row to 24 cells so index access never goes out of bounds
    while (r.length < 24) r.push('');
    // Source of truth: the CHECKBOX columns (U=21→idx20, W=23→idx22).
    // Fall back to the text status columns (N=14→idx13, R=18→idx17).
    var paid   = (r[20] === true) || (String(r[13]).trim() === 'Paid');
    var picked = (r[22] === true) || (String(r[17]).trim() === 'Picked Up');
    if (filter === 'pending'  && paid) return;
    if (filter === 'paid'     && !paid) return;
    if (filter === 'pickups'  && (!paid || picked)) return;
    if (filter === 'pickedup' && !picked) return;
    rows.push({
      rowNum:         idx + 3,
      orderId:        r[0],
      orderDate:      r[1] ? Utilities.formatDate(new Date(r[1]), 'America/Chicago', 'M/d h:mm a') : '',
      customerName:   r[2],
      phone:          r[3],
      whatsapp:       r[4],
      email:          r[5],
      city:           r[6],
      bangan:         r[7]  || 0,
      kesar:          r[8]  || 0,
      rasalu:         r[9]  || 0,
      himayat:        r[10] || 0,
      totalBoxes:     r[11] || 0,
      totalAmount:    r[12] || 0,
      paymentStatus:  paid ? 'Paid' : (r[13] || 'Pending'),
      paymentMethod:  r[14] || '',
      pickupStatus:   picked ? 'Picked Up' : (r[17] || 'Awaiting'),
      pickupLocation: r[18] || '',
      notes:          r[19] || '',
      paymentReceived: paid,
      pickedUp:        picked
    });
  });
  return rows;
}

function markPaymentReceived_(orderId, role, method) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ORDERS_SHEET_NAME);
  var rowNum = findRowByOrderId_(sheet, orderId);
  if (!rowNum) return {status: 'error', message: 'Order #' + orderId + ' not found'};
  // Column U (21) = Payment Received checkbox
  var cell = sheet.getRange(rowNum, 21);
  if (cell.getValue() === true && role !== 'admin') {
    return {status: 'error', code: 'already_done', message: 'Already marked. Admin override required to undo.'};
  }
  cell.setValue(true);
  cell.setBackground('#cfd8dc');
  // Keep the text status columns in sync so dashboard + lists always agree
  sheet.getRange(rowNum, 14).setValue('Paid');       // N = Payment Status
  if (method) sheet.getRange(rowNum, 15).setValue(method); // O = Payment Method (Zelle/Cash)
  if (!sheet.getRange(rowNum, 17).getValue()) {
    sheet.getRange(rowNum, 17).setValue(new Date());  // Q = Payment Date
  }
  try {
    sendPaymentEmailIfNeeded_(sheet, rowNum);
  } catch (err) { /* email send is best-effort */ }
  return {status: 'ok', message: 'Order #' + orderId + ' marked as paid' + (method ? ' (' + method + ')' : '')};
}

function markPickedUp_(orderId, role) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ORDERS_SHEET_NAME);
  var rowNum = findRowByOrderId_(sheet, orderId);
  if (!rowNum) return {status: 'error', message: 'Order #' + orderId + ' not found'};
  // Column W (23) = Picked Up checkbox
  var cell = sheet.getRange(rowNum, 23);
  if (cell.getValue() === true && role !== 'admin') {
    return {status: 'error', code: 'already_done', message: 'Already marked. Admin override required to undo.'};
  }
  cell.setValue(true);
  cell.setBackground('#cfd8dc');
  // Keep text status column in sync
  sheet.getRange(rowNum, 18).setValue('Picked Up');  // R = Pickup Status
  try {
    sendPickupEmailIfNeeded_(sheet, rowNum);
  } catch (err) { /* best-effort */ }
  return {status: 'ok', message: 'Order #' + orderId + ' marked as picked up'};
}

function unmarkAction_(orderId, type) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ORDERS_SHEET_NAME);
  var rowNum = findRowByOrderId_(sheet, orderId);
  if (!rowNum) return {status: 'error', message: 'Order #' + orderId + ' not found'};
  var col = (type === 'payment') ? 21 : 23;
  var cell = sheet.getRange(rowNum, col);
  cell.setValue(false);
  cell.setBackground(null);
  // Keep text status columns in sync
  if (type === 'payment') {
    sheet.getRange(rowNum, 14).setValue('Pending'); // N = Payment Status
    sheet.getRange(rowNum, 17).setValue('');         // Q = Payment Date (clear)
  } else {
    sheet.getRange(rowNum, 18).setValue('Awaiting'); // R = Pickup Status
  }
  // Remove any auto-lock protection on this cell
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(function(p) {
    var pRange = p.getRange();
    if (pRange.getRow() === rowNum && pRange.getColumn() === col && (p.getDescription() || '').indexOf('Auto-locked') === 0) {
      p.remove();
    }
  });
  return {status: 'ok', message: 'Order #' + orderId + ' override applied — ' + type + ' unmarked'};
}

function findRowByOrderId_(sheet, orderId) {
  if (!orderId) return null;
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return null;
  var ids = sheet.getRange(3, 1, lastRow - 2, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(orderId)) return i + 3;
  }
  return null;
}

// Helpers to send the action emails when an admin marks via the GUI
// (the existing onEdit-based handlers only fire on user sheet edits, not script edits)
function sendPaymentEmailIfNeeded_(sheet, rowNum) {
  var row = sheet.getRange(rowNum, 1, 1, 24).getValues()[0];
  var data = {
    orderId: row[0], customerName: row[2], email: row[5],
    bangan: row[7], kesar: row[8], rasalu: row[9], himayat: row[10],
    totalBoxes: row[11], total: row[12],
    pickupLocation: row[18]
  };
  if (row[21] === true) return; // already sent
  try {
    sendPaymentConfirmedEmail_(data);
    logEmail_(data.orderId, data.customerName, 'Payment Confirmed (via Admin GUI)', data.email, 'SENT');
    sheet.getRange(rowNum, 22).setValue(true); // mark Payment Email Sent column
  } catch (err) {
    logEmail_(data.orderId, data.customerName, 'Payment Confirmed (via Admin GUI)', data.email, 'FAILED: ' + err);
  }
}

function sendPickupEmailIfNeeded_(sheet, rowNum) {
  var row = sheet.getRange(rowNum, 1, 1, 24).getValues()[0];
  var data = {
    orderId: row[0], customerName: row[2], email: row[5],
    bangan: row[7], kesar: row[8], rasalu: row[9], himayat: row[10],
    totalBoxes: row[11], total: row[12],
    pickupLocation: row[18]
  };
  if (row[23] === true) return; // already sent
  try {
    sendPickupThankYouEmail_(data);
    logEmail_(data.orderId, data.customerName, 'Pickup Thank You (via Admin GUI)', data.email, 'SENT');
    sheet.getRange(rowNum, 24).setValue(true); // mark Pickup Email Sent
  } catch (err) {
    logEmail_(data.orderId, data.customerName, 'Pickup Thank You (via Admin GUI)', data.email, 'FAILED: ' + err);
  }
}
