/************************************************
 * EDEN BI V4 — Api.js
 *
 * Lets the dashboard run from a custom domain (dashboard.edenalkhalij.com)
 * without an iframe, so it never depends on the visitor's Google sign-in
 * (browsers signed into several Google accounts showed
 * "Sorry, unable to open the file at this time").
 *
 *   GET  /exec?view=raw         -> {html}: the rendered dashboard page
 *   POST /exec  {fn, args}      -> {ok, result} | {ok:false, error}
 *
 * Only the functions the dashboard already calls through google.script.run
 * are allowed; every one of them still validates the login token itself.
 ************************************************/

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

function apiRawPage_() {
  getUsersSheet_();
  const html = HtmlService.createTemplateFromFile('Index').evaluate().getContent();
  return apiJson_({ html: html });
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
    return apiJson_({ ok: true, result: result === undefined ? null : result });
  } catch (err) {
    return apiJson_({ ok: false, error: (err && err.message) ? err.message : String(err) });
  }
}
