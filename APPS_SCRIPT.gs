/**
 * Apleona PO — Google Apps Script backend
 *
 * Handles:
 *   GET  ?action=claim      → atomically increment PO and return new number
 *   GET  ?action=peek       → return next PO without incrementing
 *   GET  ?action=set&n=X    → set counter so next claim returns X
 *   GET  ?action=release&n=X→ if X is the last claimed, decrement back (rollback)
 *   GET  ?action=info       → capabilities (translation availability)
 *   POST { texts: [...] }   → translate Hebrew→English via OpenAI (server-side key)
 *
 * SETUP:
 *   1) Create a Google Sheet named "PO Counter" with starting number in Sheet1!A1
 *      (e.g. 2000149  →  first claim returns 2000150)
 *   2) Extensions → Apps Script → paste this code
 *   3) (Optional, for translation) Project Settings → Script Properties:
 *        OPENAI_KEY = sk-proj-...
 *   4) Deploy → New deployment → Web app
 *        Execute as: Me
 *        Who has access: Anyone (with the URL)
 *   5) Copy the /exec URL into the app settings
 */

function doGet(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const action = (e && e.parameter && e.parameter.action) || 'claim';
    if (action === 'claim')   return claim_();
    if (action === 'peek')    return peek_();
    if (action === 'set')     return setNext_(e.parameter.n);
    if (action === 'release') return release_(e.parameter.n);
    if (action === 'info')    return info_();
    return error_('Unknown action: ' + action);
  } catch (err) {
    return error_(err.message || String(err));
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function doPost(e) {
  // Translation proxy — keeps OpenAI key on the server
  const key = PropertiesService.getScriptProperties().getProperty('OPENAI_KEY');
  if (!key) return error_('Translation not configured (set OPENAI_KEY in Script Properties)');
  try {
    const body = JSON.parse(e.postData.contents);
    const texts = body.texts || [];
    if (!Array.isArray(texts) || !texts.length) return json_({ translations: [] });

    const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + key, 'Content-Type': 'application/json' },
      payload: JSON.stringify({
        model: 'gpt-4o-mini',
        temperature: 0,
        messages: [
          { role: 'system', content:
            'You are a professional translator for formal business purchase orders. ' +
            'Translate Hebrew to English using formal business English. ' +
            'Preserve numbers, codes, and brand names verbatim. ' +
            'Return ONLY a valid JSON array of strings — no markdown, no code blocks, no explanation. ' +
            'The output array must have exactly the same number of elements as the input.'
          },
          { role: 'user', content: JSON.stringify(texts) }
        ]
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(res.getContentText());
    if (data.error) return error_(data.error.message || 'OpenAI error');
    let raw = data.choices[0].message.content.trim()
      .replace(/^```json\s*/i, '').replace(/\s*```$/i, '').trim();
    const translations = JSON.parse(raw);
    if (!Array.isArray(translations) || translations.length !== texts.length) {
      return error_('Translation returned wrong number of items');
    }
    return json_({ translations: translations });
  } catch (err) {
    return error_(err.message || String(err));
  }
}

// ── PO counter operations ─────────────────────────────────
function sheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1');
}
function current_() {
  return Number(sheet_().getRange('A1').getValue()) || 2000149;
}
function claim_() {
  const n = current_() + 1;
  sheet_().getRange('A1').setValue(n);
  return json_({ po: n });
}
function peek_() {
  return json_({ po: current_() + 1 });
}
function setNext_(nStr) {
  const n = parseInt(nStr, 10);
  if (!n || n < 1) return error_('Invalid number');
  sheet_().getRange('A1').setValue(n - 1);
  return json_({ ok: true, next: n });
}
function release_(nStr) {
  // Only rollback if the number being released IS the current "last assigned"
  // (i.e. nobody else claimed after it). Safe against accidental decrement.
  const n = parseInt(nStr, 10);
  const cur = current_();
  if (n === cur) {
    sheet_().getRange('A1').setValue(n - 1);
    return json_({ ok: true, released: n });
  }
  return json_({ ok: false, reason: 'newer numbers claimed since', current: cur });
}
function info_() {
  const hasKey = !!PropertiesService.getScriptProperties().getProperty('OPENAI_KEY');
  return json_({
    poNumbering: true,
    translation: hasKey,
    nextPo: current_() + 1,
    version: '2'
  });
}

// ── HTTP helpers ──────────────────────────────────────────
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function error_(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
