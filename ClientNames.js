/************************************************
 * EDEN BI V4 — ClientNames.js
 *
 * Arabic and English client names as on the papers, for the new Deals DB
 * files only (EDEN Deals DB — Egypt / Georgia & Past).
 *  1) Scan: for every deal with a "Documents Link" folder, Google OCR reads
 *     the ID / passport files in that folder (a temporary Google Doc copy in
 *     the private Deals DB folder, deleted right after).
 *     · Arabic name  ← Egyptian ID card ("بطاقة تحقيق الشخصية").
 *     · English name ← passport machine line (P<EGY…); without a passport
 *       the English name on the deal is kept and marked for review.
 *     Only the names are kept — never the ID number, address or job.
 *  2) Review: results go to the "Name Review" tab of each file.
 *  3) Approve (System › Client names): writes "Client Name (AR)" and
 *     "Client Name (EN)" on the deal row. Nothing is written before that.
 ************************************************/

const CN_TAB_ = 'Name Review';
const CN_HEAD_ = ['Deal ID', 'Project', 'Unit Code', 'Name on the deal', 'Arabic name (found)', 'English name (found)', 'Arabic from', 'English from', 'Confidence', 'Note', 'Status', 'Checked At', 'Approved By'];
const CN_MAX_FILE_ = 12 * 1024 * 1024;   // bigger scans (full contracts) are skipped
const CN_MAX_OCR_ = 4;                    // files read per client folder at most

function cnRequireAdmin_(token) { return targetsRequireAdmin_(token); }

function cnTab_(ss) {
  let sh = ss.getSheetByName(CN_TAB_);
  if (!sh) {
    sh = ss.insertSheet(CN_TAB_);
    sh.getRange(1, 1, 1, CN_HEAD_.length).setValues([CN_HEAD_]).setFontWeight('bold').setBackground('#1f3a33').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, sh.getMaxRows() - 1, CN_HEAD_.length).setNumberFormat('@');
    [110, 110, 110, 260, 260, 260, 220, 220, 90, 260, 90, 140, 160].forEach((w, i) => { try { sh.setColumnWidth(i + 1, w); } catch (e) {} });
  }
  return sh;
}
function cnReviewMap_(sh) {
  const last = sh.getLastRow(), map = {};
  if (last < 2) return map;
  sh.getRange(2, 1, last - 1, CN_HEAD_.length).getValues().forEach((r, i) => { const id = String(r[0]).trim(); if (id) map[id] = { row: i + 2, v: r }; });
  return map;
}
function cnFolderId_(url) { const m = String(url || '').match(/folders\/([A-Za-z0-9_-]{10,})/) || String(url || '').match(/[?&]id=([A-Za-z0-9_-]{10,})/); return m ? m[1] : ''; }

/* ---------- reading the papers ---------- */
function cnOcr_(fileId) {
  const copy = Drive.Files.copy({ name: 'EDEN OCR (temporary)', mimeType: 'application/vnd.google-apps.document', parents: [ddbFolder_()] }, fileId, { ocrLanguage: 'ar', supportsAllDrives: true, fields: 'id' });
  try {
    const r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + copy.id + '/export?mimeType=text/plain', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    return r.getResponseCode() === 200 ? r.getContentText('UTF-8') : '';
  } finally {
    try { (Drive.Files.remove || Drive.Files['delete']).call(Drive.Files, copy.id, { supportsAllDrives: true }); } catch (e) { try { Drive.Files.update({ trashed: true }, copy.id, null, { supportsAllDrives: true }); } catch (e2) {} }
  }
}
function cnNorm_(s) {
  return String(s || '').replace(/[۰-۹]/g, d => String(d.charCodeAt(0) - 0x06F0)).replace(/[٠-٩]/g, d => String(d.charCodeAt(0) - 0x0660))
    .replace(/[ی]/g, 'ي').replace(/[ک]/g, 'ك').replace(/[ـً-ْ]/g, '').replace(/[ \t]+/g, ' ');
}
const CN_STOP_ = /جمهورية|بطاق|تحقيق|الشخصي|مصر العربية|ذكر|انثى|أنثى|مسلم|مسيحي|قبطي|اعزب|أعزب|انسة|آنسة|أنسة|متزوج|مطلق|ارمل|أرمل|سارية|طالب|بكلية|كلية|جامعة|جامعه|مهندس|محاسب|طبيب|دكتور|موظف|معاش|اعمال|أعمال|ربة|منزل|بدون|مدير|محامي|ضابط|مدرس|معلم|شركة|شركه|وزارة|بنك|شارع|محافظ|قسم|مركز|مدين|القاهر|الجيز|الاسكندري|عمارة|عماره|برج|قرية|قريه|عزبة|عزبه|مساكن|الحي|حي |امتياز|ليسانس|بكالوريوس|دبلوم|اخصائي|أخصائي|فني|عامل|سائق|تاجر|صيدلي|رجل|سيدة|الرقم|القومي|تاريخ|الميلاد|الاصدار|الإصدار|جواز|سفر|وزارة الداخلية/;
/* Egyptian ID card → full Arabic name (first name line + the rest). */
function cnArabicFromId_(raw) {
  const t = cnNorm_(raw);
  if (!/بطاق[ةه]\s*تحقيق|تحقيق\s*الشخصي/.test(t)) return '';
  const lines = t.split(/\r?\n/).map(x => x.trim()).filter(Boolean);
  const hi = lines.findIndex(x => /تحقيق\s*الشخصي|بطاق[ةه]/.test(x));
  const ok = x => /^[ء-ي ]+$/.test(x) && !CN_STOP_.test(x) && !/^ش /.test(x);
  let restI = -1;
  for (let i = hi + 1; i < lines.length; i++) { const w = lines[i].split(' ').length; if (ok(lines[i]) && w >= 2 && w <= 7) { restI = i; break; } }
  if (restI < 0) return '';
  // first name: a single Arabic word, nearest to the rest line (before it)
  let first = '';
  for (let i = restI - 1; i >= 0; i--) { if (ok(lines[i]) && lines[i].split(' ').length === 1 && lines[i].length >= 2) { first = lines[i]; break; } }
  return (first ? first + ' ' : '') + lines[restI];
}
/* Passport → { name (English, as printed), country, ar (Arabic printed above it, if any) }. */
const CN_EN_STOP_ = /REPUBLIC|PASSPORT|EGYPT|AUTHORITY|PALESTIN|KINGDOM|EMIRATES|MINISTRY|NATIONALITY|DATE|BIRTH|ISSUE|EXPIRY|PLACE|SEX|MALE|FEMALE|SIGNATURE|BEARER|VALID|COUNTRIES|PROFESSION|ADDRESS|NAME|NOTES|ENGINEER|OWNER|COMPANY|HOLDER|TYPE|CODE|STATE|UNITED|ARAB|SURNAME|GIVEN/;
function cnPassport_(raw) {
  const txt = String(raw || '').replace(/\\/g, '');
  const t = txt.toUpperCase().replace(/[ \t]/g, '').replace(/«/g, '<').replace(/[‹〈]/g, '<');
  const isPass = /PASSPORT|جواز|P<[A-Z]{3}/i.test(txt) || /[A-Z]{3}\d{7}[MF<]\d{7}/.test(t) || (/DATE\s*OF\s*BIRTH/i.test(txt) && /NATIONALITY/i.test(txt));
  if (!isPass) return null;
  const tc = s => s.toLowerCase().replace(/(^|[\s.-])([a-z])/g, (x, a, c) => a + c.toUpperCase());
  let mrz = null;
  const m = t.match(/P[<A-Z0-9]([A-Z]{3})([A-Z]+(?:<[A-Z]+)*)<<([A-Z<]+)/);
  if (m) mrz = { country: m[1], sur: m[2].replace(/</g, ' ').trim(), given: m[3].replace(/<+$/, '').replace(/</g, ' ').trim() };
  // the name printed in capitals (fuller than the machine line: "MOHAMMED J. H. ALMUGHRABI")
  const lines = cnNorm_(txt).split(/\r?\n/).map(x => x.trim()).filter(Boolean);
  const caps = lines.map((x, i) => ({ x: x, i: i })).filter(o => /^[A-Z][A-Z .'-]+$/.test(o.x) && o.x.split(/\s+/).length >= 2 && o.x.split(/\s+/).length <= 7 && !CN_EN_STOP_.test(o.x));
  let printed = null;
  if (mrz) printed = caps.filter(o => o.x.indexOf(mrz.sur.split(' ')[0]) >= 0 || o.x.indexOf(mrz.given.split(' ')[0]) >= 0).sort((a, b) => b.x.length - a.x.length)[0] || null;
  else printed = caps.filter(o => o.x.split(/\s+/).length >= 3)[0] || null;
  let name = printed ? tc(printed.x.replace(/\s+/g, ' ')) : mrz ? tc((mrz.given + ' ' + mrz.sur).replace(/\s+/g, ' ')) : '';
  if (!name || name.split(' ').length < 2) return null;
  // Arabic name printed just above the English one
  let ar = '';
  if (printed) {
    const run = [];
    for (let i = printed.i - 1; i >= 0 && run.length < 4; i--) {
      const x = lines[i];
      if (/^[A-Z]{3}$/.test(x)) continue;
      if (/^[ء-ي ]+$/.test(x) && !CN_STOP_.test(x)) run.unshift(x); else break;
    }
    ar = run.join(' ').replace(/\s+/g, ' ').trim();
    if (ar.split(' ').length < 2) ar = '';
  }
  return { name: name, country: mrz ? mrz.country : '', ar: ar, sure: !!(mrz && printed) };
}
function cnKindOfFile_(f) {
  const n = String(f.name || '').toLowerCase();
  if (/broker|wallet|receipt|ايصال|إيصال|contract|عقد|res ?form|reservation|payment|transfer|تحويل|عائد|plan|\bex\b|^ex\./.test(n)) return 'skip';
  if (/pass|جواز|باسبور/.test(n)) return 'passport';
  if (/\bid\b|^id[\s._-]|[\s_-]id\.|national|بطاق|رقم قومي|identity|\bead\b|card/.test(n)) return 'id';
  return 'other';
}
/* Files of one client folder (+ one level of sub-folders), best first. */
function cnFiles_(folderId) {
  const out = [], seen = {};
  const list = (pid, depth) => {
    let token = null;
    do {
      const r = Drive.Files.list({ q: "'" + pid + "' in parents and trashed = false", fields: 'nextPageToken,files(id,name,mimeType,size)', pageSize: 200, pageToken: token, supportsAllDrives: true, includeItemsFromAllDrives: true });
      (r.files || []).forEach(f => {
        if (f.mimeType === 'application/vnd.google-apps.folder') { if (depth < 2) list(f.id, depth + 1); return; }
        if (seen[f.id]) return; seen[f.id] = 1;
        if (!/^image\/|application\/pdf/.test(f.mimeType) || Number(f.size || 0) > CN_MAX_FILE_) return;
        const k = cnKindOfFile_(f); if (k === 'skip') return;
        out.push({ id: f.id, name: f.name, kind: k, inInfo: depth > 0 });
      });
      token = r.nextPageToken;
    } while (token);
  };
  list(folderId, 0);
  const rank = f => (f.kind === 'id' ? 0 : f.kind === 'passport' ? 1 : 2);
  return out.sort((a, b) => rank(a) - rank(b));
}
/* Read one client folder → list of people found on the papers. */
function cnReadFolder_(folderId, hardStop) {
  const files = cnFiles_(folderId), people = [];
  let read = 0;
  for (const f of files) {
    if (read >= CN_MAX_OCR_ || (hardStop && Date.now() > hardStop && people.length)) break;
    if (f.kind === 'other' && people.length && read >= 3) break;
    let text = '';
    try { text = cnOcr_(f.id); read++; } catch (e) { console.warn('OCR ' + f.name + ': ' + e); continue; }
    const ar = cnArabicFromId_(text), pp = cnPassport_(text);
    if (ar) people.push({ ar: ar, file: f.name, fileId: f.id, type: 'ID' });
    if (pp) people.push({ en: pp.name, country: pp.country, arPass: pp.ar, sure: pp.sure, file: f.name, fileId: f.id, type: 'Passport' });
  }
  return { people: people, files: files.length, read: read };
}
/* Who on the deal is who: by the file name first (e.g. "Mariam ID.pdf"), then in order. */
function cnMatch_(dealName, found) {
  const parts = cleanClientName_(dealName).split(/\s+-\s+/).map(x => x.trim()).filter(Boolean);
  const ids = found.people.filter(p => p.ar), pps = found.people.filter(p => p.en);
  const used = {}, low = s => String(s || '').toLowerCase();
  const pick = (list, person, i) => {
    const fn = low(person.split(' ')[0]).slice(0, 4);
    let hit = list.find(p => !used[p.fileId] && fn.length >= 3 && low(p.file).indexOf(fn) >= 0);
    if (!hit && list.length === parts.length) hit = list.filter(p => !used[p.fileId])[0];
    if (!hit && parts.length === 1) hit = list.find(p => !used[p.fileId]);
    if (hit) used[hit.fileId] = 1;
    return hit || null;
  };
  // passports: match by English words too
  const pickPassport = person => {
    const words = low(person).split(/\s+/);
    let best = null, score = 0;
    pps.forEach(p => { if (used[p.fileId]) return; const s = low(p.en).split(' ').filter(w => words.indexOf(w) >= 0).length; if (s > score) { score = s; best = p; } });
    if (!best && pps.filter(p => !used[p.fileId]).length && parts.length === 1) best = pps.find(p => !used[p.fileId]);
    if (best) used[best.fileId] = 1;
    return best;
  };
  let sure = parts.length > 0;
  const ar = [], en = [], arFrom = [], enFrom = [], notes = [];
  parts.forEach((person, i) => {
    const a = pick(ids, person, i), p = pickPassport(person);
    if (a) { ar.push(a.ar); arFrom.push(a.file); }
    else if (p && p.arPass) { ar.push(p.arPass); arFrom.push(p.file + ' (passport — check)'); sure = false; }
    else { ar.push(''); sure = false; }
    if (p && !p.sure) sure = false;
    if (p) { en.push(p.en); enFrom.push(p.file); } else { en.push(person); enFrom.push('name on the deal'); }
  });
  if (!ids.length && !pps.length) notes.push(found.read ? 'No ID or passport could be read in the folder' : 'No ID / passport file in the folder');
  else if (ar.some(x => !x)) notes.push('Arabic name not found for: ' + parts.filter((_, i) => !ar[i]).join(', '));
  if (ids.length > parts.length) notes.push(ids.length + ' IDs found for ' + parts.length + ' name(s) — check the match');
  if (enFrom.some(x => x === 'name on the deal')) notes.push('English kept from the deal (no passport)');
  const conf = !parts.length ? 'Check' : sure && ids.length === parts.length && !enFrom.some(x => x === 'name on the deal') ? 'High' : sure ? 'Good' : 'Check';
  return { ar: ar.filter(Boolean).join(' - '), en: en.filter(Boolean).join(' - '), arFrom: [...new Set(arFrom)].join(', '), enFrom: [...new Set(enFrom)].join(', '), conf: conf, note: notes.join(' · ') };
}

/* ---------- admin calls ---------- */
/* Read the papers of the next deals (about 4½ minutes per call). opts: { rescan: bool } */
function scanClientNames(authToken, opts) {
  const user = cnRequireAdmin_(authToken);
  opts = opts || {};
  const started = Date.now(), budget = 40000, hardStop = started + 70000;   // the page waits 90 s for an answer
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { busy: true, message: 'A scan is already running — try again in a minute.' };
  let done = 0, left = 0, total = 0, noLink = 0;
  try {
    const at = ddbColIndex_(), N = DDB_COLS_.length, folderCache = {};
    for (const key of ['EG', 'GP']) {
      const ss = ddbOpen_(key); if (!ss) continue;
      const sh = ss.getSheetByName('Deals'), tab = cnTab_(ss), map = cnReviewMap_(tab);
      const last = sh.getLastRow(); if (last < DDB_FIRST_ROW_) continue;
      const V = sh.getRange(DDB_FIRST_ROW_, 1, last - DDB_FIRST_ROW_ + 1, N).getValues();
      const L = sh.getRange(DDB_FIRST_ROW_, at['Documents Link'], last - DDB_FIRST_ROW_ + 1, 1).getRichTextValues();
      const append = [];
      for (let i = 0; i < V.length; i++) {
        const v = V[i], id = String(v[at['Deal ID'] - 1] || '').trim(); if (!id) continue;
        if (String(v[at['Status'] - 1]) === 'Cancel') continue;
        const rt = L[i][0], link = (rt && (rt.getLinkUrl() || (rt.getRuns().find(r => r.getLinkUrl()) || { getLinkUrl: () => '' }).getLinkUrl())) || String(v[at['Documents Link'] - 1] || '');
        const fid = cnFolderId_(link);
        total++;
        if (!fid) { noLink++; continue; }
        const prev = map[id];
        if (prev && !opts.rescan) continue;
        if (prev && opts.rescan && (String(prev.v[10]) === 'Approved' || (opts.since && String(prev.v[11]) >= opts.since))) continue;
        if (Date.now() - started > budget) { left++; continue; }
        let found = folderCache[fid];
        if (!found) { try { found = cnReadFolder_(fid, hardStop); } catch (e) { found = { people: [], files: 0, read: 0, error: String(e && e.message || e) }; } folderCache[fid] = found; }
        const m = cnMatch_(String(v[at['Client Name'] - 1] || ''), found);
        const row = [id, String(v[at['Project'] - 1] || ''), String(v[at['Unit Code'] - 1] || ''), cleanClientName_(String(v[at['Client Name'] - 1] || '')),
          m.ar, m.en, m.arFrom, m.enFrom, found.error ? 'Check' : m.conf, found.error ? 'Folder could not be opened: ' + found.error : m.note, 'Pending', Utilities.formatDate(new Date(), 'Africa/Cairo', 'yyyy-MM-dd HH:mm'), ''];
        if (prev) tab.getRange(prev.row, 1, 1, CN_HEAD_.length).setValues([row]); else append.push(row);
        done++;
      }
      if (append.length) tab.getRange(tab.getLastRow() + 1, 1, append.length, CN_HEAD_.length).setValues(append);
    }
  } finally { lock.releaseLock(); }
  if (done) auditLog_(user, 'Client names read from papers', { details: done + ' deals' + (left ? ' · ' + left + ' left' : '') });
  return { done: done, left: left, total: total, noLink: noLink };
}

function getClientNameReview(authToken) {
  cnRequireAdmin_(authToken);
  const rows = [], counts = { EG: 0, GP: 0 };
  ['EG', 'GP'].forEach(key => {
    const ss = ddbOpen_(key); if (!ss) return;
    const tab = ss.getSheetByName(CN_TAB_); if (!tab || tab.getLastRow() < 2) return;
    tab.getRange(2, 1, tab.getLastRow() - 1, CN_HEAD_.length).getDisplayValues().forEach(r => {
      if (!String(r[0]).trim()) return;
      counts[key]++;
      rows.push({ store: key, dealId: r[0], project: r[1], unit: r[2], name: r[3], ar: r[4], en: r[5], arFrom: r[6], enFrom: r[7], conf: r[8], note: r[9], status: r[10], at: r[11], by: r[12] });
    });
  });
  return { rows: rows, counts: counts };
}

/* items: [{store, dealId, ar, en}] → written on the deal row + marked Approved. */
function approveClientNames(authToken, items) {
  const user = cnRequireAdmin_(authToken);
  items = (items || []).filter(x => x && x.dealId && DDB_STORES_[x.store]);
  if (!items.length) return { saved: 0 };
  const at = ddbColIndex_(), who = user.name || user.username, now = new Date();
  let saved = 0;
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    ['EG', 'GP'].forEach(key => {
      const mine = items.filter(x => x.store === key); if (!mine.length) return;
      const ss = ddbOpen_(key), sh = ss.getSheetByName('Deals'), tab = cnTab_(ss), map = cnReviewMap_(tab);
      const last = sh.getLastRow(); if (last < DDB_FIRST_ROW_) return;
      const ids = sh.getRange(DDB_FIRST_ROW_, at['Deal ID'], last - DDB_FIRST_ROW_ + 1, 1).getValues().map(r => String(r[0]).trim());
      mine.forEach(x => {
        const i = ids.indexOf(String(x.dealId).trim()); if (i < 0) return;
        const r = DDB_FIRST_ROW_ + i, ar = String(x.ar || '').replace(/\s+/g, ' ').trim(), en = String(x.en || '').replace(/\s+/g, ' ').trim();
        sh.getRange(r, at['Client Name (AR)'], 1, 2).setValues([[ar, en]]);
        sh.getRange(r, at['Updated At'], 1, 2).setValues([[now, who + ' (client names)']]);
        const p = map[String(x.dealId).trim()];
        if (p) { tab.getRange(p.row, 5, 1, 2).setValues([[ar, en]]); tab.getRange(p.row, 11, 1, 3).setValues([['Approved', Utilities.formatDate(now, 'Africa/Cairo', 'yyyy-MM-dd HH:mm'), who]]); }
        saved++;
      });
    });
  } finally { lock.releaseLock(); }
  auditLog_(user, 'Client names approved', { details: saved + ' deals' });
  try { if (typeof dashCacheClear_ === 'function') dashCacheClear_(); } catch (e) {}
  return { saved: saved };
}
