/**
 * NRKards order intake — Google Apps Script backend
 * -------------------------------------------------
 * Handles BOTH form payloads from nrkards.com:
 *   - type: 'order'  (Order-for-delivery form)  → name, email, address, pack, energy, notes
 *   - type: 'custom' (Custom-build form)         → name, kid, pokemon, contact(email), notes
 *
 * On each submit it (1) appends a row to the "Orders" tab and
 * (2) emails OWNER_EMAIL so nothing sits unseen in the Sheet.
 *
 * SETUP:
 *   1. Open your Google Sheet (signed in as nrkards@gmail.com).
 *   2. Copy its ID from the URL — the chunk between /d/ and /edit:
 *        docs.google.com/spreadsheets/d/THIS_PART/edit
 *      Paste it into SHEET_ID below.
 *   3. Save.
 *   4. In the editor, choose "doGet" in the function dropdown → Run.
 *      Approve BOTH permission prompts (Sheets + Gmail). This authorizes it.
 *   5. Deploy → Manage deployments → edit (pencil) → Version: New version → Deploy.
 *      (If no deployment exists yet: Deploy → New deployment → Web app,
 *       Execute as: Me, Who has access: Anyone. Copy the .../exec URL.)
 *   6. The /exec URL goes in BOTH SHEET_ENDPOINT constants in index.html.
 */

const OWNER_EMAIL = 'nrkards@gmail.com';
const SHEET_NAME  = 'Orders';
const SHEET_ID    = '1hMu8SdgAK-D4h4CFsX9dnhBtSsasa9Br45CyvF98MDE';   // ← the chunk between /d/ and /edit in your Sheet URL

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const email = data.email || data.contact || '';   // order uses `email`, custom uses `contact`
    getSheet().appendRow([
      data.submittedAt || new Date().toISOString(),
      data.type    || '',
      data.name    || '',
      email,
      data.pack    || '',
      data.energy  || '',
      data.address || '',
      data.kid     || '',
      data.pokemon || '',
      data.notes   || ''
    ]);
    notifyOwner(data, email);
    return json({ ok: true });
  } catch (err) {
    // Still try to flag the failure to the owner so a broken submit isn't silent.
    try {
      MailApp.sendEmail(OWNER_EMAIL, 'NRKards form ERROR', String(err) + '\n\n' + (e && e.postData ? e.postData.contents : '(no body)'));
    } catch (_) {}
    return json({ ok: false, error: String(err) });
  }
}

// Health check — open the /exec URL in a browser to confirm it's live.
function doGet() {
  return json({ ok: true, service: 'NRKards order intake', time: new Date().toISOString() });
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp','Type','Name','Email','Pack','Energy','Address','Kid','Pokemon','Notes']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function notifyOwner(data, email) {
  const isOrder = data.type === 'order';
  const subject = isOrder
    ? 'New pack order — ' + (data.name || 'someone')
    : 'New custom request — ' + (data.name || 'someone');
  let body;
  if (isOrder) {
    body =
      'New pack order from the site:\n\n' +
      'Name:    ' + (data.name    || '') + '\n' +
      'Email:   ' + (email        || '') + '\n' +
      'Address: ' + (data.address || '') + '\n' +
      'Pack:    ' + (data.pack    || '') + '\n' +
      'Energy:  ' + (data.energy  || '') + '\n' +
      'Notes:   ' + (data.notes   || '—') + '\n';
  } else {
    body =
      'New custom pack request from the site:\n\n' +
      'Name:    ' + (data.name    || '') + '\n' +
      'Kid:     ' + (data.kid     || '—') + '\n' +
      'Email:   ' + (email        || '') + '\n' +
      'Request: ' + (data.pokemon || '') + '\n' +
      'Notes:   ' + (data.notes   || '—') + '\n';
  }
  body += '\nSubmitted: ' + (data.submittedAt || new Date().toISOString());
  // replyTo = customer's email so you can just hit "Reply" to answer them.
  MailApp.sendEmail({
    to: OWNER_EMAIL,
    subject: subject,
    body: body,
    replyTo: email || OWNER_EMAIL
  });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
