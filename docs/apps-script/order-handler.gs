/* ============================================================
   TuKan Designs — Order Handler (Google Apps Script Web App)
   ------------------------------------------------------------
   Handles three actions POSTed from the website:

     action: 'customOrder'  -> saves the spec + uploaded photos to a
                               Drive folder and appends a row to the
                               "Orders" sheet.
     action: 'newsletter'   -> appends the email to the "Newsletter"
                               sheet (deduped).
     action: 'paymentLog'   -> logs a completed PayPal payment to the
                               "Payments" sheet.

   HOW TO DEPLOY (short version — full steps in SETUP.md):
     1. Create a Google Sheet. Copy its ID into SHEET_ID below.
     2. Create a Drive folder for order photos. Copy its ID into
        DRIVE_FOLDER_ID below.
     3. Extensions -> Apps Script, paste this file.
     4. Deploy -> New deployment -> type "Web app".
        Execute as: Me.  Who has access: Anyone.
     5. Copy the /exec Web App URL into assets/data/site-config.json
        as "orderHandlerUrl".

   The site posts JSON as text/plain to avoid CORS preflight, so we
   parse e.postData.contents ourselves.
   ============================================================ */

// ---- CONFIG: fill these two in ------------------------------------
var SHEET_ID        = 'PASTE_GOOGLE_SHEET_ID_HERE';
var DRIVE_FOLDER_ID = 'PASTE_DRIVE_FOLDER_ID_HERE';
// Where to email new-order notifications (blank = no email sent).
var NOTIFY_EMAIL    = 'Bwstudebaker7@gmail.com';
// -------------------------------------------------------------------

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    var action = body.action || '';

    if (action === 'customOrder')  return _json(handleCustomOrder(body));
    if (action === 'consultation') return _json(handleConsultation(body));
    if (action === 'newsletter')   return _json(handleNewsletter(body));
    if (action === 'paymentLog')   return _json(handlePaymentLog(body));

    return _json({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return _json({ ok: false, error: String(err && err.message || err) });
  }
}

// Simple GET so you can sanity-check the deployment in a browser.
function doGet() {
  return _json({ ok: true, service: 'TuKan Designs order handler', ts: new Date().toISOString() });
}

/* ---------------- Custom order ---------------- */
function handleCustomOrder(body) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = _sheet(ss, 'Orders', [
    'Timestamp', 'Name', 'Email', 'Phone', 'Table Type', 'Wood',
    'Finish', 'Dimensions', 'Budget', 'Notes', 'Photos'
  ]);

  var now = new Date();
  var stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var safeName = (body.name || 'order').replace(/[^\w\- ]+/g, '').trim() || 'order';

  // Save uploaded photos (array of { name, dataUrl }) to a per-order folder.
  var photoLinks = [];
  var photos = body.photos || [];
  if (photos.length) {
    var parent = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var folder = parent.createFolder(stamp + ' — ' + safeName);
    for (var i = 0; i < photos.length; i++) {
      var p = photos[i];
      var dataUrl = p && (p.data || p.dataUrl);
      if (!dataUrl) continue;
      try {
        var blob = _dataUrlToBlob(dataUrl, p.name || ('photo-' + (i + 1)));
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoLinks.push(file.getUrl());
      } catch (imgErr) {
        photoLinks.push('(failed: ' + String(imgErr) + ')');
      }
    }
  }

  sheet.appendRow([
    stamp,
    body.name || '', body.email || '', body.phone || '',
    body.tableType || '', body.wood || '', body.finish || '',
    body.dimensions || '', body.budget || '', body.notes || '',
    photoLinks.join('\n')
  ]);

  if (NOTIFY_EMAIL) {
    try {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: 'New custom order — ' + (body.name || 'Unknown'),
        body: [
          'A new custom table order came in:', '',
          'Name: ' + (body.name || ''),
          'Email: ' + (body.email || ''),
          'Phone: ' + (body.phone || ''),
          'Table type: ' + (body.tableType || ''),
          'Wood: ' + (body.wood || ''),
          'Finish: ' + (body.finish || ''),
          'Dimensions: ' + (body.dimensions || ''),
          'Budget: ' + (body.budget || ''),
          '', 'Notes:', (body.notes || '(none)'),
          '', 'Photos:', (photoLinks.join('\n') || '(none)')
        ].join('\n')
      });
    } catch (mailErr) { /* non-fatal */ }
  }

  return { ok: true, saved: true, photos: photoLinks.length };
}

/* ---------------- Consultation request ---------------- */
function handleConsultation(body) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = _sheet(ss, 'Consultations', [
    'Timestamp', 'Name', 'Email', 'Phone', 'Preferred Contact', 'Availability', 'Notes', 'Status'
  ]);

  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([
    stamp, body.name || '', body.email || '', body.phone || '',
    body.method || '', body.availability || '', body.notes || '', 'New'
  ]);

  if (NOTIFY_EMAIL) {
    try {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        replyTo: body.email || NOTIFY_EMAIL,
        subject: 'New consultation request — ' + (body.name || 'Unknown'),
        body: [
          'Someone requested a free consultation on the website:', '',
          'Name: ' + (body.name || ''),
          'Email: ' + (body.email || ''),
          'Phone: ' + (body.phone || ''),
          'Preferred contact: ' + (body.method || ''),
          'Availability: ' + (body.availability || ''),
          '', 'Notes:', (body.notes || '(none)'),
          '', 'Reply to this email to reach them directly.'
        ].join('\n')
      });
    } catch (mailErr) { /* non-fatal */ }
  }

  return { ok: true, requested: true };
}

/* ---------------- Newsletter ---------------- */
function handleNewsletter(body) {
  var email = (body.email || '').trim().toLowerCase();
  if (!email || email.indexOf('@') === -1) return { ok: false, error: 'Invalid email' };

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = _sheet(ss, 'Newsletter', ['Timestamp', 'Email', 'Source']);

  // Dedupe against existing emails (column B).
  var existing = sheet.getRange(2, 2, Math.max(sheet.getLastRow() - 1, 0), 1).getValues();
  for (var i = 0; i < existing.length; i++) {
    if (String(existing[i][0]).trim().toLowerCase() === email) {
      return { ok: true, already: true };
    }
  }

  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([stamp, email, body.source || 'website']);
  return { ok: true, subscribed: true };
}

/* ---------------- Payment log ---------------- */
function handlePaymentLog(body) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = _sheet(ss, 'Payments', [
    'Timestamp', 'Payer', 'Email', 'Invoice Ref', 'Amount', 'Currency', 'Order ID', 'Status'
  ]);
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([
    stamp, body.payer || '', body.email || '', body.invoiceRef || '',
    body.amount || '', body.currency || 'USD', body.orderId || '', body.status || 'COMPLETED'
  ]);
  return { ok: true, logged: true };
}

/* ---------------- helpers ---------------- */
function _sheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _dataUrlToBlob(dataUrl, filename) {
  var parts = dataUrl.split(',');
  var meta = parts[0] || '';
  var b64 = parts[1] || '';
  var mimeMatch = meta.match(/data:([^;]+);/);
  var mime = mimeMatch ? mimeMatch[1] : 'application/octet-stream';
  var bytes = Utilities.base64Decode(b64);
  // Ensure a sensible extension.
  var ext = (mime.split('/')[1] || 'bin').replace('jpeg', 'jpg');
  var name = /\.\w+$/.test(filename) ? filename : (filename + '.' + ext);
  return Utilities.newBlob(bytes, mime, name);
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
