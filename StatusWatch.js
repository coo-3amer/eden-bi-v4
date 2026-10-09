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
 *   (The Inventory is not touched by hand edits in the sheet — only deals
 *    added / updated from the system move the unit's status.)
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
        // Inventory is NOT changed from hand edits in the Deals sheet (owner's decision).
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

/* Hand edits in an Inventory sheet → Activity Log + Telegram (called by LiveSync). */
function auditInventorySheetEdit_(e, key) {
  try {
    if (!e || !e.range) return;
    const cfg = INVENTORY_SOURCES[key];
    if (!cfg) return;
    const sheet = e.range.getSheet();
    const tab = sheet.getName();
    if (/^BI /i.test(tab)) return;                                   // the system's own tabs
    const r0 = e.range.getRow(), c0 = e.range.getColumn(), nR = e.range.getNumRows(), nC = e.range.getNumColumns();
    const lastCol = sheet.getLastColumn();
    const top = sheet.getRange(1, 1, Math.min(30, Math.max(1, sheet.getLastRow())), lastCol).getDisplayValues();
    const hr = inventoryFindHeaderRow_(top, cfg);
    if (hr >= 0 && r0 + nR - 1 <= hr + 1) return;                    // header rows
    const heads = hr >= 0 ? top[hr] : [];
    const idx = inventoryIndex_(heads);
    let codeCol = 0;
    for (const a of cfg.headers.unitCode || []) { const p = idx[inventoryHeaderKey_(a)]; if (p !== undefined) { codeCol = p + 1; break; } }
    const colName = c => String(heads[c - 1] || '').replace(/\s+/g, ' ').trim() || ('Column ' + c);
    let who = '';
    try { who = (e.user && e.user.getEmail && e.user.getEmail()) || ''; } catch (err) {}
    if (nR === 1 && nC === 1) {
      const oldV = e.oldValue === undefined ? '' : String(e.oldValue);
      const newV = e.value === undefined ? String(e.range.getDisplayValue() || '') : String(e.value);
      if (oldV === newV) return;
      const unit = codeCol ? String(sheet.getRange(r0, codeCol).getDisplayValue() || '').trim() : '';
      auditLog_(who || 'Sheet user', 'Inventory edit (sheet)', { project: cfg.project, unit: unit,
        details: tab + ' › ' + colName(c0) + ': ' + (oldV || '—') + ' → ' + (newV || '—'), ref: tab + ' row ' + r0 });
    } else {
      let units = [];
      if (codeCol) {
        const first = Math.max(r0, hr + 2), last = Math.min(r0 + nR - 1, first + 29);
        if (last >= first) units = sheet.getRange(first, codeCol, last - first + 1, 1).getDisplayValues().map(x => String(x[0] || '').trim()).filter(Boolean);
      }
      auditLog_(who || 'Sheet user', 'Inventory edit (sheet)', { project: cfg.project,
        unit: units.length === 1 ? units[0] : (units.length ? units.length + ' units' : ''),
        details: tab + ' · ' + (nR * nC) + ' cells changed · ' + colName(c0) + (nC > 1 ? ' … ' + colName(c0 + nC - 1) : '') +
          (units.length > 1 ? ' · ' + units.slice(0, 15).join(', ') + (units.length > 15 ? ' …' : '') : ''),
        ref: tab + ' rows ' + r0 + (nR > 1 ? '–' + (r0 + nR - 1) : '') });
    }
  } catch (err) {
    console.warn('auditInventorySheetEdit_: ' + err);
  }
}
