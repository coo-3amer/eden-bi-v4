/************************************************
 * EDEN BI V4 — Api.js
 *
 * Lets the dashboard run from a custom domain (dashboard.edenalkhalij.com)
 * without an iframe, so it never depends on the visitor's Google sign-in
 * (browsers signed into several Google accounts showed
 * "Sorry, unable to open the file at this time").
 *
 *   GET  /exec?view=ping                  -> {ok:true}
 *   GET  /exec?view=raw                   -> {id, parts, html}: first slice of the page
 *   GET  /exec?view=pagever               -> {id, parts, hash}: current page version
 *   GET  /exec?view=chunk&id=X&i=N        -> {data}: one slice of a large result
 *   POST /exec  {fn, args}                -> {ok, result} | {ok, chunked, parts} | {ok:false, error}
 *
 * Responses are kept small (Google's content server rejects large ones), so
 * the page and large results are sent in slices.
 *
 * Only the functions the dashboard already calls through google.script.run
 * are allowed; every one of them still validates the login token itself.
 ************************************************/

const API_SLICE_CHARS_ = 90000;      // < 100 KB CacheService value limit
const API_CACHE_SECONDS_ = 1200;

const API_ALLOWED_FUNCTIONS_ = {
  loginUser: true,
  logoutUser: true,
  getDashboardData: true,
  getTargetsData: true,
  saveTargetRecord: true,
  deleteTargetRecord: true,
  getUsersList: true,
  createUser: true,
  updateUser: true,
  deleteUser: true,
  getInventoryData: true,
  getInventoryUnit: true,
  getInventoryMovements: true,
  validateInventoryUnitForDeal: true,
  getDealEntryOptions: true,
  getDealInventoryUnits: true,
  checkDealUnitAvailability: true,
  saveNewDeal: true,
  getDashboardConfig: true,
  saveDashboardConfig: true,
  resetDashboardConfig: true,
  logDashboardActivity: true
};

function apiJson_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Splits text into slices and stores them in the script cache; returns the slice count.
function apiStoreSlices_(id, text) {
  const values = {};
  let parts = 0;
  for (let i = 0; i < text.length; i += API_SLICE_CHARS_) {
    values[id + ':' + parts] = text.slice(i, i + API_SLICE_CHARS_);
    parts++;
  }
  const keys = Object.keys(values);
  for (let i = 0; i < keys.length; i += 20) {
    const batch = {};
    keys.slice(i, i + 20).forEach(k => batch[k] = values[k]);
    CacheService.getScriptCache().putAll(batch, API_CACHE_SECONDS_);
  }
  return parts;
}

function apiReadSlice_(id, i) {
  const v = CacheService.getScriptCache().get(String(id) + ':' + Number(i));
  if (v === null) throw new Error('Expired, please reload.');
  return v;
}

// The rendered page is cached for a short time so most visits skip rendering.
const API_PAGE_META_KEY_ = 'PAGE_META_V1';
const API_PAGE_SECONDS_ = 900;

function apiPageMeta_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(API_PAGE_META_KEY_);
  if (cached) {
    const meta = JSON.parse(cached);
    if (cache.get(meta.id + ':0') !== null) return meta;
  }
  getUsersSheet_();
  const html = HtmlService.createTemplateFromFile('Index').evaluate().getContent();
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, html, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
  const id = 'P' + hash;
  const parts = apiStoreSlices_(id, html);
  const meta = { id: id, parts: parts, hash: hash };
  cache.put(API_PAGE_META_KEY_, JSON.stringify(meta), API_PAGE_SECONDS_);
  return meta;
}

function apiRawPage_() {
  const meta = apiPageMeta_();
  return apiJson_({ id: meta.id, parts: meta.parts, hash: meta.hash, html: apiReadSlice_(meta.id, 0) });
}

function apiDoGet_(e) {
  const view = e.parameter.view;
  try {
    if (view === 'ping') return apiJson_({ ok: true, time: new Date().toISOString() });
    if (view === 'raw') return apiRawPage_();
    if (view === 'pagever') { const m = apiPageMeta_(); return apiJson_({ id: m.id, parts: m.parts, hash: m.hash }); }
    if (view === 'chunk') return apiJson_({ data: apiReadSlice_(e.parameter.id, e.parameter.i) });
  } catch (err) {
    return apiJson_({ ok: false, error: (err && err.message) ? err.message : String(err) });
  }
  return apiJson_({ ok: false, error: 'Unknown view.' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return apiJson_({ ok: false, error: 'Invalid request.' });
  }

  const name = String(req.fn || '');
  if (!Object.prototype.hasOwnProperty.call(API_ALLOWED_FUNCTIONS_, name)) {
    return apiJson_({ ok: false, error: 'Function not allowed: ' + name });
  }

  const fn = globalThis[name];
  if (typeof fn !== 'function') {
    return apiJson_({ ok: false, error: 'Function not found: ' + name });
  }

  try {
    const args = Array.isArray(req.args) ? req.args : [];
    const result = fn.apply(null, args);
    const text = JSON.stringify(result === undefined ? null : result);
    if (text.length <= API_SLICE_CHARS_) {
      return apiJson_({ ok: true, result: JSON.parse(text) });
    }
    const id = 'R' + Utilities.getUuid().replace(/-/g, '');
    const parts = apiStoreSlices_(id, text);
    return apiJson_({ ok: true, chunked: id, parts: parts });
  } catch (err) {
    return apiJson_({ ok: false, error: (err && err.message) ? err.message : String(err) });
  }
}
