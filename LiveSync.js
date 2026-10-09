/************************************************
 * EDEN BI V4 — LiveSync.js
 *
 * Sheets -> dashboard in seconds instead of minutes.
 *  1. An installable trigger on each source spreadsheet (Deals, Brokers,
 *     inventories) fires on every edit / row change.
 *  2. It rebuilds only what changed (dashboard data or that inventory)
 *     straight away; edits made close together are batched into one rebuild.
 *  3. Open dashboards ask every 30 s "is there newer data?" (a tiny call)
 *     and swap it in without a reload, keeping filters and page.
 * The 5-minute warm-up stays as a safety net.
 * Run setupLiveSync once from the Apps Script editor.
 ************************************************/

const LIVE_DEBOUNCE_MS_ = 4000;

/* Run ONCE from the editor (and again if spreadsheets change). */
function setupLiveSync() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onDataSheetEdit')
    .forEach(t => ScriptApp.deleteTrigger(t));
  const ids = liveSourceIds_();
  const done = [], failed = [];
  Object.keys(ids).forEach(id => {
    ['onEdit', 'onChange'].forEach(kind => {
      for (let attempt = 1; attempt <= 4; attempt++) {
        try {
          const b = ScriptApp.newTrigger('onDataSheetEdit').forSpreadsheet(id);
          (kind === 'onEdit' ? b.onEdit() : b.onChange()).create();
          done.push(ids[id] + ' ' + kind);
          return;
        } catch (err) {
          if (attempt === 4) failed.push(ids[id] + ' ' + kind + ' (' + err + ')');
          else Utilities.sleep(2000 * attempt);   // Google sometimes refuses once; try again
        }
      }
    });
  });
  Logger.log('LiveSync ON: ' + done.join(', '));
  if (failed.length) Logger.log('NOT set (run setupLiveSync again in a minute): ' + failed.join(' | '));
}

/* spreadsheet id -> what it feeds ('DASH' or 'INV:<key>'). */
function liveSourceIds_() {
  const out = {};
  out[SPREADSHEET_ID] = 'DASH';
  if (typeof BROKER_SPREADSHEET_ID !== 'undefined') out[BROKER_SPREADSHEET_ID] = 'DASH';
  Object.keys(INVENTORY_SOURCES).forEach(k => {
    const id = INVENTORY_SOURCES[k].spreadsheetId;
    if (id && !out[id]) out[id] = 'INV:' + k;
  });
  return out;
}

/* Installable trigger handler (edit or change in a source spreadsheet). */
function onDataSheetEdit(e) {
  let id = '';
  try { id = e && e.source ? e.source.getId() : ''; } catch (err) {}
  const target = liveSourceIds_()[id];
  if (!target) return;
  // In the main file only the Deals tab feeds the dashboard (skip Users, Targets…).
  if (target === 'DASH' && id === SPREADSHEET_ID && e && e.range) {
    try { if (e.range.getSheet().getSheetId() !== DEALS_GID) return; } catch (err) {}
    // A Status typed by hand in Deals also moves the unit in the Inventory.
    if (typeof auditDealsSheetEdit_ === 'function') auditDealsSheetEdit_(e);
    // Status is a formula in Deals: compare it with the last Status seen (alerts + Inventory).
    if (typeof watchDealStatus_ === 'function') watchDealStatus_(e);
  }
  // Hand edits in an Inventory sheet are announced (they never change the deals).
  if (target.indexOf('INV:') === 0 && e && e.range && typeof auditInventorySheetEdit_ === 'function') auditInventorySheetEdit_(e, target.slice(4));
  const cache = CacheService.getScriptCache();
  const dirtyKey = 'LIVE_DIRTY_' + target;
  cache.put(dirtyKey, String(Date.now()), 600);
  if (cache.get('LIVE_RUNNING_' + target)) return;      // the running rebuild will pick it up
  cache.put('LIVE_RUNNING_' + target, '1', 300);
  try {
    for (let round = 0; round < 3; round++) {
      Utilities.sleep(LIVE_DEBOUNCE_MS_);               // let quick follow-up edits land first
      const mark = cache.get(dirtyKey);
      cache.remove(dirtyKey);
      if (!mark && round > 0) break;
      liveRebuild_(target);
      if (!cache.get(dirtyKey)) break;                  // nothing new while rebuilding
    }
  } finally {
    cache.remove('LIVE_RUNNING_' + target);
  }
}

function liveRebuild_(target) {
  if (target === 'DASH') {
    dashCacheWrite_('ALL', buildDashboardData_({ role: 'Admin' }));
    try { dashCacheWrite_('EGV', buildDashboardData_({ role: 'Egypt Viewer' })); } catch (e) { console.warn('LiveSync EGV: ' + e); }
  } else if (target.indexOf('INV:') === 0) {
    inventoryDataCached_(target.slice(4), true, false);
  }
}

/* Tiny call the open dashboards make every 30 s. */
function getDataVersion(authToken) {
  const user = validateAuthToken_(authToken);
  const scope = isEgyptViewerUser_(user) ? 'EGV' : 'ALL';
  return { v: dashCacheVersion_(scope) };
}
