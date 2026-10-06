/************************************************
 * EDEN BI V4 — Api.js
 *
 * JSON API so the dashboard page can be hosted outside Google
 * (GitHub Pages) and work on any browser or phone, whatever Google
 * accounts are signed in. The page calls the same server functions it
 * calls through google.script.run; each one still checks the login token.
 *
 * All requests are GET (Google's redirect is only reliable for GET):
 *   /exec?api=ping                          -> {ok:true}
 *   /exec?api=call&q=<base64url {fn,args}>  -> {ok,result} | {ok,gz} | {ok,parts,id} | {ok:false,error}
 *   /exec?api=put&id=X&i=N&d=<text>         -> stores one slice of a long request
 *   /exec?api=call&ref=X&n=N                -> runs a request sent with api=put
 *   /exec?api=chunk&id=X&i=N                -> one slice of a large result
 ************************************************/

const API_SLICE_CHARS_ = 90000;       // under the 100 KB CacheService value limit
const API_DIRECT_CHARS_ = 450000;     // larger results are sent in slices
const API_GZIP_FROM_CHARS_ = 20000;   // results above this are gzip-compressed
const API_CACHE_SECONDS_ = 600;
const API_BLOCKED_ = { doGet: true, doPost: true, include: true, warmDashboardCache: true, setupDashboardWarmup: true, setupUsersSheet: true, setupDealsSheetGuard: true, checkDealsSheetMapping: true, debugDealUnitCheck: true, setupLiveSync: true, onDataSheetEdit: true };

function apiJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function apiDoGet_(e) {
  const p = (e && e.parameter) || {};
  try {
    switch (p.api) {
      case 'ping':  return apiJson_({ ok: true, time: new Date().toISOString() });
      case 'put':   return apiJson_(apiPut_(p));
      case 'chunk': return apiJson_({ ok: true, data: apiReadSlice_(p.id, p.i) });
      case 'call':  return apiJson_(apiDispatch_(apiReadRequest_(p), p.z === '1'));
    }
    return apiJson_({ ok: false, error: 'Unknown API action.' });
  } catch (err) {
    return apiJson_({ ok: false, error: apiErr_(err) });
  }
}

function apiErr_(err) { return (err && err.message) ? err.message : String(err); }

function apiPut_(p) {
  const id = String(p.id || '');
  if (!/^[A-Za-z0-9]{8,64}$/.test(id)) throw new Error('Invalid request id.');
  CacheService.getScriptCache().put('Q' + id + ':' + Number(p.i), String(p.d || ''), API_CACHE_SECONDS_);
  return { ok: true };
}

function apiReadRequest_(p) {
  let q;
  if (p.ref) {
    const n = Number(p.n || 0), cache = CacheService.getScriptCache(), keys = [];
    for (let i = 0; i < n; i++) keys.push('Q' + p.ref + ':' + i);
    const got = cache.getAll(keys);
    q = keys.map(k => { if (got[k] == null) throw new Error('Request expired, please try again.'); return got[k]; }).join('');
  } else {
    q = String(p.q || '');
  }
  q += '===='.slice(0, (4 - q.length % 4) % 4);
  const text = Utilities.newBlob(Utilities.base64DecodeWebSafe(q)).getDataAsString('UTF-8');
  return JSON.parse(text);
}

function apiDispatch_(req, gzipOk) {
  const name = String((req && req.fn) || '');
  // Same rule as google.script.run: no private (trailing "_") functions.
  if (!/^[A-Za-z][A-Za-z0-9]*$/.test(name) || API_BLOCKED_[name]) {
    return { ok: false, error: 'Function not allowed: ' + name };
  }
  const fn = globalThis[name];
  if (typeof fn !== 'function') return { ok: false, error: 'Function not found: ' + name };
  try {
    const result = fn.apply(null, Array.isArray(req.args) ? req.args : []);
    let text = JSON.stringify(result === undefined ? null : result);
    const out = { ok: true };
    if (gzipOk && text.length > API_GZIP_FROM_CHARS_) {
      text = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(text, 'application/json')).getBytes());
      out.gz = true;
    }
    if (text.length <= API_DIRECT_CHARS_) {
      out.data = text;
      return out;
    }
    const id = 'R' + Utilities.getUuid().replace(/-/g, '');
    out.id = id;
    out.parts = apiStoreSlices_(id, text);
    return out;
  } catch (err) {
    return { ok: false, error: apiErr_(err) };
  }
}

function apiStoreSlices_(id, text) {
  const cache = CacheService.getScriptCache();
  let parts = 0, batch = {};
  for (let i = 0; i < text.length; i += API_SLICE_CHARS_) {
    batch[id + ':' + parts++] = text.slice(i, i + API_SLICE_CHARS_);
    if (Object.keys(batch).length === 20) { cache.putAll(batch, API_CACHE_SECONDS_); batch = {}; }
  }
  if (Object.keys(batch).length) cache.putAll(batch, API_CACHE_SECONDS_);
  return parts;
}

function apiReadSlice_(id, i) {
  const v = CacheService.getScriptCache().get(String(id) + ':' + Number(i));
  if (v === null) throw new Error('Expired, please reload.');
  return v;
}
