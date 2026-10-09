/************************************************
 * EDEN BI V4 — StatusWatch.js
 *
 * In the Deals sheet the Status column is a FORMULA (it turns to Cancel when
 * a cancel date is entered, to Part of DP / Total DP / Client Sigend… when
 * the Deal Confirmation boxes are ticked). A formula never fires an edit
 * event, so instead every hand edit re-reads the Status of the edited rows
 * and compares it with the last Status seen (kept in the hidden tab
 * "_Status Seen"). When it differs:
 *   - Activity Log + Telegram: "Deal status changed (sheet)"
 *   - the unit's Inventory status follows (InventorySync rules)
 * Called by LiveSync.onDataSheetEdit for edits in the Deals tab.
 ************************************************/

const STATUS_SEEN_SHEET_ = '_Status Seen';

function statusSeenSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(STATUS_SEEN_SHEET_);
  if (sh) return { sh: sh, fresh: false };
  sh = ss.insertSheet(STATUS_SEEN_SHEET_);
  try { sh.hideSheet(); } catch (e) {}
  sh.getRange(1, 1).setValue('Last Status seen for each row of the Deals tab (row number = same row). Used by EDEN BI alerts — do not edit.');
  // First time: remember every row's current Status, so nothing is announced by mistake.
  const deals = getMainDealsSheet_();
  const cols = mainDealsColumns_(deals);
  const sc = mainDealsCol_(cols, ['deal status|status', 'Status']);
  const last = deals.getLastRow();
  if (sc && last >= DATA_START_ROW) {
    const vals = deals.getRange(DATA_START_ROW, sc, last - DATA_START_ROW + 1, 1).getDisplayValues();
    if (sh.getMaxRows() < last) sh.insertRowsAfter(sh.getMaxRows(), last - sh.getMaxRows());
    sh.getRange(DATA_START_ROW, 1, vals.length, 1).setNumberFormat('@').setValues(vals);
  }
  return { sh: sh, fresh: true };
}

function watchDealStatus_(e) {
  try {
    if (!e || !e.range) return;
    const deals = e.range.getSheet();
    if (deals.getSheetId() !== DEALS_GID) return;
    const r0 = e.range.getRow(), nR = e.range.getNumRows();
    const first = Math.max(r0, DATA_START_ROW), last = Math.min(r0 + nR - 1, first + 49);
    if (last < first) return;

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) return;
    try {
      const seen = statusSeenSheet_();
      if (seen.fresh) return;                         // just took the first snapshot
      const cols = mainDealsColumns_(deals);
      const at = names => mainDealsCol_(cols, names);
      const sc = at(['deal status|status', 'Status']);
      if (!sc) return;
      const n = last - first + 1;
      const S = deals.getRange(first, sc, n, 1).getDisplayValues().map(x => String(x[0] || '').trim());
      const sh = seen.sh;
      if (sh.getMaxRows() < last) sh.insertRowsAfter(sh.getMaxRows(), last - sh.getMaxRows());
      const old = sh.getRange(first, 1, n, 1).getDisplayValues().map(x => String(x[0] || '').trim());
      const changed = [];
      for (let i = 0; i < n; i++) if (S[i].toLowerCase() !== old[i].toLowerCase()) changed.push(i);
      if (!changed.length) return;
      sh.getRange(first, 1, n, 1).setNumberFormat('@').setValues(S.map(v => [v]));

      const pc = at(['Project']), uc = at(['Unit Code']), cc = at(['Final Client Name', 'Client Name']);
      let who = '';
      try { who = (e.user && e.user.getEmail && e.user.getEmail()) || ''; } catch (err) {}
      changed.forEach(i => {
        const row = first + i;
        const v = deals.getRange(row, 1, 1, cols.lastCol).getDisplayValues()[0];
        const project = pc ? String(v[pc - 1] || '').trim() : '', unit = uc ? String(v[uc - 1] || '').trim() : '', client = cc ? String(v[cc - 1] || '').trim() : '';
        if (!project && !unit && !client) return;     // empty row
        if (!old[i]) {
          // A new deal typed straight into the sheet.
          if (S[i]) auditLog_(who || 'Sheet user', 'Deal added (sheet)', { project: project, unit: unit, client: client, details: 'Status ' + S[i], ref: 'Deals row ' + row });
        } else {
          auditLog_(who || 'Sheet user', 'Deal status changed (sheet)', { project: project, unit: unit, client: client,
            details: 'Status: ' + old[i] + ' → ' + (S[i] || '—'), ref: 'Deals row ' + row });
        }
        if (project && unit && S[i] && typeof syncInventoryForDeal_ === 'function') {
          syncInventoryForDeal_(project, unit, S[i], { by: who ? 'sheet edit by ' + who : 'sheet edit' });
        }
      });
    } finally { lock.releaseLock(); }
  } catch (err) {
    console.warn('watchDealStatus_: ' + err);
  }
}

/* Deals saved or updated from the system: remember their Status too, so the
   sheet watcher does not announce them a second time. */
function statusSeenRemember_(row, status) {
  try {
    const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(STATUS_SEEN_SHEET_);
    if (!sh || !row) return;
    if (sh.getMaxRows() < row) sh.insertRowsAfter(sh.getMaxRows(), row - sh.getMaxRows());
    sh.getRange(row, 1).setNumberFormat('@').setValue(String(status || '').trim());
  } catch (e) {}
}

/* Re-reads the Status of one Deals row and remembers it (after a save from the system). */
function statusSeenRefresh_(row) {
  try {
    if (!row || row < DATA_START_ROW) return;
    SpreadsheetApp.flush();
    const deals = getMainDealsSheet_();
    const sc = mainDealsCol_(mainDealsColumns_(deals), ['deal status|status', 'Status']);
    if (sc) statusSeenRemember_(row, deals.getRange(row, sc).getDisplayValue());
  } catch (e) {}
}
