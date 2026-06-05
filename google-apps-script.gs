/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  PRADEEP'S INSTITUTE — LEAD FORM ENDPOINT
 *  Google Apps Script Web App that receives form submissions from
 *  landing-page-v4.html and appends them to a Google Sheet.
 *
 *  ┌─────────────────────────────────────────────────────────────────────┐
 *  │ SETUP — one-time, takes about 5 minutes                             │
 *  ├─────────────────────────────────────────────────────────────────────┤
 *  │ 1. Create a new Google Sheet (e.g. "Pradeep Institute Leads").      │
 *  │ 2. Copy the long ID from its URL:                                   │
 *  │       docs.google.com/spreadsheets/d/<<<THIS_PART>>>/edit            │
 *  │    Paste it into SHEET_ID below.                                    │
 *  │ 3. In the Sheet menu: Extensions → Apps Script.                     │
 *  │ 4. Delete the default Code.gs, paste THIS entire file, save.        │
 *  │ 5. Deploy → New deployment → choose type: Web app.                  │
 *  │       Description : "Pradeep Lead API v1"                           │
 *  │       Execute as  : Me (your Google account)                        │
 *  │       Who has access: Anyone                                        │
 *  │    Click Deploy → Authorize when prompted.                          │
 *  │ 6. Copy the resulting Web App URL                                   │
 *  │    (looks like https://script.google.com/macros/s/AKfy.../exec).    │
 *  │ 7. Open landing-page-v4.html, find the line                         │
 *  │       const GSHEET_ENDPOINT = 'PASTE_YOUR_GOOGLE_APPS_SCRIPT_...    │
 *  │    and replace the placeholder with your Web App URL.               │
 *  │                                                                     │
 *  │ Optional — to update later: Deploy → Manage deployments → edit.     │
 *  └─────────────────────────────────────────────────────────────────────┘
 * ═══════════════════════════════════════════════════════════════════════════
 */

const SHEET_ID  = 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE';
const SHEET_TAB = 'Leads';
const DUPLICATE_WINDOW_HOURS = 24;   // Reject same phone within this window
const HEADERS = [
  'Timestamp',
  'Name',
  'Phone',
  'Age',
  'Course',
  'Source',
  'Referer',
  'User Agent',
  'Client Submitted At',
];

// ── Main entry point ───────────────────────────────────────────────────────
function doPost(e) {
  try {
    const params = (e && e.parameter) || {};

    // ─── Server-side validation ──────────────────────────────────────────
    const name  = String(params.name  || '').trim();
    const phone = String(params.phone || '').replace(/\D/g, '');
    const age   = String(params.age   || '').trim();

    if (!name || name.length < 2) {
      return json({ status: 'error', field: 'name',
                    message: 'Name is required (minimum 2 characters).' });
    }
    if (!/^[a-zA-ZÀ-ɏ\s.'\-]+$/.test(name)) {
      return json({ status: 'error', field: 'name',
                    message: 'Name can only contain letters.' });
    }
    if (!phone || phone.length < 10 || phone.length > 15) {
      return json({ status: 'error', field: 'phone',
                    message: 'A valid phone number (10–15 digits) is required.' });
    }
    if (age) {
      const a = parseInt(age, 10);
      if (isNaN(a) || a < 16 || a > 65) {
        return json({ status: 'error', field: 'age',
                      message: 'Age must be between 16 and 65.' });
      }
    }

    // ─── Open / initialise the target sheet ──────────────────────────────
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SHEET_TAB);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_TAB);
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // ─── Server-side duplicate detection ─────────────────────────────────
    const now    = new Date();
    const cutoff = new Date(now.getTime() - DUPLICATE_WINDOW_HOURS * 3600 * 1000);
    const lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();   // Timestamp, Name, Phone
      for (let i = values.length - 1; i >= 0; i--) {
        const rowDate  = new Date(values[i][0]);
        if (!rowDate || isNaN(rowDate.getTime()) || rowDate < cutoff) break;
        const rowPhone = String(values[i][2] || '').replace(/\D/g, '');
        if (rowPhone === phone) {
          return json({
            status: 'error',
            duplicate: true,
            message: "We've already received your details. Our team will call you shortly.",
          });
        }
      }
    }

    // ─── Append new lead row ─────────────────────────────────────────────
    sheet.appendRow([
      now,
      name,
      "'" + phone,                                  // leading apostrophe → preserved as text
      age,
      String(params.course      || '').trim(),
      String(params.source      || 'website').trim(),
      String(params.referer     || '').trim(),
      String(params.userAgent   || '').trim(),
      String(params.submittedAt || '').trim(),
    ]);

    return json({ status: 'success', message: 'Lead saved.' });
  } catch (err) {
    return json({ status: 'error',
                  message: 'Server error: ' + (err && err.message ? err.message : err) });
  }
}

// ── Health check (GET) ─────────────────────────────────────────────────────
function doGet() {
  return json({ status: 'ok',
                service: 'Pradeep Institute Lead API',
                method: 'POST application/x-www-form-urlencoded or multipart/form-data' });
}

// ── Helper: JSON response with proper MIME ─────────────────────────────────
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
