/** FOCUS Sheets receiver v1.0.0 — deploy as Google Apps Script Web App.
 * Script properties: SPREADSHEET_ID, TELEGRAM_BOT_TOKEN.
 * Never put bot token in GitHub or browser code.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    if (!e || !e.postData || !e.postData.contents) throw Error('Empty request');
    const body = JSON.parse(e.postData.contents);
    const identity = verifyTelegram_(String(body.initData || ''));
    if (!identity || !identity.user || !identity.user.id) throw Error('Unauthorized');
    if (!lock.tryLock(10000)) throw Error('Busy');
    try {
      const sheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
      if (!sheetId) throw Error('Spreadsheet not configured');
      const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('Недельные отчёты');
      if (!sheet) throw Error('Sheet not found');
      const answers = body.answers;
      if (!Array.isArray(answers) || answers.length !== 11 ||
          answers.some(v => typeof v !== 'string' || !v.trim() || v.length > 500)) throw Error('Invalid answers');
      const comment = body.comment;
      if (typeof comment !== 'string' || !comment.trim() || comment.length > 3000) throw Error('Invalid comment');
      const requestId = String(body.requestId || '');
      if (!/^[a-zA-Z0-9_-]{12,80}$/.test(requestId)) throw Error('Invalid request ID');
      // Idempotency: prevent double submissions. Request IDs are stored in script properties.
      const cache = CacheService.getScriptCache();
      const key = 'focus_' + requestId;
      if (cache.get(key)) return json_({ok:true, duplicate:true});
      const user = identity.user;
      const name = [user.first_name || '', user.last_name || ''].join(' ').trim() || user.username || String(user.id);
      const safe = v => {
        const s = String(v);
        return /^[=+@\-\t\r]/.test(s) ? "'" + s : s;
      };
      sheet.appendRow([new Date(), safe(name + ' (TG ' + user.id + ')')].concat(answers.map(safe), [safe(comment)]));
      cache.put(key, '1', 21600);
      return json_({ok:true});
    } finally { lock.releaseLock(); }
  } catch (err) {
    return json_({ok:false,error:String(err.message || err)});
  }
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function verifyTelegram_(initData) {
  if (!initData || initData.length > 12000) return null;
  const token = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (!token) throw Error('Bot token not configured');
  const pairs = initData.split('&').map(p => {
    const i = p.indexOf('=');
    return [decodeURIComponent((i < 0 ? p : p.slice(0,i)).replace(/\+/g,' ')),
            decodeURIComponent((i < 0 ? '' : p.slice(i+1)).replace(/\+/g,' '))];
  });
  const data = {};
  pairs.forEach(([k,v]) => { if (Object.prototype.hasOwnProperty.call(data,k)) throw Error('Duplicate field'); data[k]=v; });
  const hash = data.hash;
  const authDate = Number(data.auth_date);
  if (!/^[a-f0-9]{64}$/i.test(hash || '') || !Number.isFinite(authDate) ||
      Math.abs(Date.now()/1000 - authDate) > 86400) return null;
  const check = pairs.filter(([k]) => k !== 'hash').sort((a,b) => a[0].localeCompare(b[0]))
    .map(([k,v]) => k + '=' + v).join('\n');
  const secret = Utilities.computeHmacSha256Signature(token, 'WebAppData');
  const signature = Utilities.computeHmacSha256Signature(check, secret)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
  if (signature !== hash.toLowerCase()) return null;
  try { return {user: JSON.parse(data.user || '{}')}; } catch (_) { return null; }
}
