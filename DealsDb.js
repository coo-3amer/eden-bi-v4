/************************************************
 * EDEN BI V4 — DealsDb.js
 *
 * The new deals sheet ("EDEN Deals DB"), a separate Google Sheets file that
 * will replace the old Deals tab step by step.
 *   - One row per deal, one header row (row 2; row 1 only shows the groups),
 *     a fixed Deal ID (EDN-YY-0001) for every deal.
 *   - Deals are added / changed from the dashboard only: every tab is locked
 *     to the account that runs EDEN BI.
 *   - The old Deals sheet keeps working until it is switched off.
 * Phase 1 (this file): create the file (admins, from the dashboard).
 * Script Property DEALS_DB_ID holds the file id.
 ************************************************/

const DDB_COLS_ = [["Record", "Deal ID", "id", "System", "Fixed ID for every deal, e.g. EDN-26-0001. Never changes, never reused."], ["Record", "Created At", "datetime", "System", "When the deal was added."], ["Record", "Created By", "text", "System", "Who added it (dashboard user)."], ["Record", "Updated At", "datetime", "System", "Last change."], ["Record", "Updated By", "text", "System", "Who made the last change."], ["Client", "Client Name", "text", "Sales", "Full name as on the ID. No unit code after the name."], ["Client", "Mobile 1", "phone", "Sales", "International format, digits only: 201001234567 / 966551234567."], ["Client", "Mobile 2", "phone", "Sales", "Second number, same format."], ["Client", "Email", "email", "Sales", "One email, lower case."], ["Client", "Nationality", "list", "Sales", "From the list."], ["Client", "Client Type", "list", "Sales", "Internal / Overseas / Company …"], ["Client", "Gender", "list", "Sales", "Male / Female."], ["Client", "Birth Date", "date", "Sales", "Real date."], ["Client", "Job Title", "text", "Sales", ""], ["Client", "Job Type", "list", "Sales", ""], ["Client", "ID Type", "list", "Sales", "National ID / Passport."], ["Client", "ID Number", "text", "Sales", "Kept as text (no scientific notation)."], ["Client", "ID Address", "text", "Sales", ""], ["Client", "Residence Address", "text", "Sales", ""], ["Client", "Documents Link", "url", "System", "Drive folder with the client papers."], ["Unit", "Project", "list", "Sales", "From the list (EDEN WALK, CCG, …)."], ["Unit", "Unit Code", "text", "Sales", "Checked against the inventory."], ["Unit", "Unit Category", "list", "System", "Main / Shared / Divided (from the inventory)."], ["Unit", "Block", "text", "System", "From the inventory."], ["Unit", "Floor", "text", "System", "From the inventory."], ["Unit", "View", "text", "System", "From the inventory."], ["Unit", "Unit Type", "list", "System", "Commercial / Admin / Medical …"], ["Unit", "Finishing Type", "list", "System", ""], ["Unit", "Area m²", "number", "System", "Indoor area."], ["Unit", "Outdoor Area m²", "number", "System", ""], ["Price", "Currency", "list", "System", "EGP / USD — by project."], ["Price", "Primary Meter Price", "money", "System", "List price per m² on the day of the deal."], ["Price", "Outdoor Meter Price", "money", "System", ""], ["Price", "Primary Total Price", "money", "System", "List total."], ["Price", "Discount %", "percent", "Sales", "Needs approval above the limit (Discount Approvals)."], ["Price", "Discount Offer", "list", "Sales", "Name of the offer, if any."], ["Price", "Meter Price After Discount", "money", "System", "Calculated."], ["Price", "Final Price", "money", "System", "Calculated: after discount."], ["Price", "Maintenance %", "percent", "Sales", ""], ["Price", "Maintenance Amount", "money", "System", "Calculated."], ["Reservation", "Reservation Date", "date", "Sales", "The deal date used everywhere in the dashboard."], ["Reservation", "Res Code", "text", "Sales", "Receipt / reservation code."], ["Reservation", "Reservation Amount", "money", "Sales", ""], ["Reservation", "Reservation Currency", "list", "Sales", ""], ["Reservation", "Reservation Payment Method", "list", "Sales", "Cash / Transfer / Cheque …"], ["Reservation", "EOI Code", "text", "Sales", "Only for EOI deals."], ["Status", "Status", "list", "System", "Set by the system from the ticks below: Reservation → Part of DP → Total DP → Client Signed → Company Signed → Delivered → Closed, or Cancel."], ["Status", "Additional Status", "list", "Sales", "Renovation / Resale / …"], ["Status", "Deal Type", "list", "Sales", "Direct / Share / …"], ["Status", "Days to Close", "number", "System", "Calculated: reservation → company signed."], ["Down payment", "DP %", "percent", "Sales", ""], ["Down payment", "DP Required", "money", "System", "Calculated from DP % × final price."], ["Down payment", "DP Paid", "money", "Finance", ""], ["Down payment", "Total Paid", "money", "Finance", "All money received on the deal."], ["Down payment", "Remaining DP", "money", "System", "Calculated."], ["Down payment", "DP Payment Method", "list", "Finance", ""], ["Down payment", "DP Currency", "list", "Finance", ""], ["Down payment", "DP Date", "date", "Finance", ""], ["Down payment", "Receipt No.", "text", "Finance", ""], ["Payment plan", "Plan Type", "list", "Sales", ""], ["Payment plan", "Installments Period", "text", "Sales", ""], ["Payment plan", "Installment Plan", "text", "Sales", ""], ["Payment plan", "Period Type", "list", "Sales", "Monthly / Quarterly …"], ["Payment plan", "Payment Exception", "text", "Sales", ""], ["Contract", "Contract Type", "list", "Contracts", ""], ["Contract", "Contract Code", "text", "Contracts", ""], ["Contract", "Contract Date", "date", "Contracts", ""], ["Contract", "Client Sign Date", "date", "Contracts", ""], ["Contract", "Company Sign Date", "date", "Contracts", ""], ["Contract", "Delivery Date", "date", "Contracts", ""], ["Confirmation", "Part of DP ✓", "check", "Finance", "Ticks move the Status."], ["Confirmation", "Total DP ✓", "check", "Finance", ""], ["Confirmation", "Client Signed ✓", "check", "Contracts", "(typo fixed)"], ["Confirmation", "Company Signed ✓", "check", "Contracts", ""], ["Confirmation", "Wallet Checked ✓", "check", "Finance", ""], ["Confirmation", "Delivery Signed ✓", "check", "Contracts", ""], ["Confirmation", "Delivered ✓", "check", "Contracts", ""], ["Confirmation", "COO Confirmed ✓", "check", "COO", ""], ["Sales team", "Branch", "list", "Sales", ""], ["Sales team", "Sales Name", "list", "Sales", "From the users / sales list."], ["Sales team", "Shared With", "list", "Sales", "For shared deals."], ["Sales team", "Sales Manager", "list", "System", "Filled from the salesperson’s team."], ["Sales team", "Head Of Sales", "list", "System", ""], ["Sales team", "CCO", "list", "System", ""], ["Source", "Main Source", "list", "Sales", "Direct / Indirect …"], ["Source", "Source Type", "list", "Sales", "Broker / Campaign / Personal / Previous Client …"], ["Source", "Campaign Name", "list", "Sales", ""], ["Source", "Media Buyer", "list", "Sales", ""], ["Broker", "Brokerage Company", "list", "Sales", "From the broker companies list."], ["Broker", "BC Sales", "text", "Sales", ""], ["Broker", "BC Manager", "text", "Sales", "(typo fixed)"], ["Internal commission", "Comm Sales Paid ✓", "check", "Finance", ""], ["Internal commission", "Comm Managers Paid ✓", "check", "Finance", ""], ["Internal commission", "Comm Staff Paid ✓", "check", "Finance", ""], ["Internal commission", "Comm Note", "text", "Finance", ""], ["Cancellation", "Cancelled ✓", "check", "Admin", "Ticking it (with a date and reason) sets Status = Cancel and frees the unit."], ["Cancellation", "Cancel Date", "date", "Admin", ""], ["Cancellation", "Cancel Reason", "list", "Admin", "From a list (+ free text): changed unit / could not pay / changed mind …"], ["Cancellation", "Cancel Code", "text", "Admin", ""], ["Cancellation", "Refund Amount", "money", "Finance", "Money returned to the client."], ["Other", "Notes", "text", "Anyone", ""], ["Other", "Old Deals Row", "number", "System", "Row in the old Deals sheet (while we move)."], ["Other", "Old Deal Num", "text", "System", ""]];
const DDB_HEADER_ROW_ = 2;
const DDB_FIRST_ROW_ = 3;
const DDB_ROWS_ = 1000;            // rows prepared with formats / checkboxes (more are added when needed)
const DDB_GROUP_COLORS_ = { 'Record': '#5B6B85', 'Client': '#2F6E9E', 'Unit': '#3D8B6D', 'Price': '#B07C1E', 'Reservation': '#8A5CD6', 'Status': '#C0392B',
  'Down payment': '#1F8A70', 'Payment plan': '#4A7A3A', 'Contract': '#6D4C41', 'Confirmation': '#2E7D32', 'Sales team': '#D35400', 'Source': '#1565C0',
  'Broker': '#7B1FA2', 'Internal commission': '#AD8B00', 'Cancellation': '#B71C1C', 'Other': '#616161' };
/* Lists tab: list name → dashboard row field used to fill it from the current data. */
const DDB_LISTS_ = [
  ['Project', 'project'], ['Status', null], ['Additional Status', 'additionalStatus'], ['Deal Type', 'dealStatus'], ['Branch', 'branch'], ['Sales Name', 'salesName'],
  ['Shared With', 'salesName'], ['Sales Manager', 'salesManager'], ['Head Of Sales', 'headOfSales'], ['CCO', 'cco'], ['Main Source', 'mainSource'], ['Source Type', 'sourceType'],
  ['Campaign Name', 'campaignName'], ['Media Buyer', 'mediaBuyer'], ['Brokerage Company', 'brokerageCompany'], ['Nationality', 'nationality'], ['Client Type', 'clientType'],
  ['Unit Type', 'unitType'], ['Finishing Type', 'finishing'], ['Currency', null], ['Reservation Currency', null], ['DP Currency', null], ['Gender', null], ['ID Type', null],
  ['Unit Category', null], ['Cancel Reason', null], ['Plan Type', 'planType'], ['Period Type', null]
];
const DDB_FIXED_LISTS_ = {
  'Status': ['Reservation', 'Part of DP', 'Total DP', 'Client Signed', 'Company Signed', 'Delivered', 'Closed', 'Cancel'],
  'Currency': ['EGP', 'USD'], 'Reservation Currency': ['EGP', 'USD'], 'DP Currency': ['EGP', 'USD'], 'Gender': ['Male', 'Female'],
  'ID Type': ['National ID', 'Passport'], 'Unit Category': ['Main Unit', 'Shared Unit', 'Divided Unit'],
  'Cancel Reason': ['Changed unit', 'Could not pay the down payment', 'Changed mind', 'Financial reasons', 'Personal reasons', 'Payment default', 'Other'],
  'Period Type': ['Monthly', 'Quarterly', 'Semi-annual', 'Annual']
};
const DDB_REMOVED_ = ['Ren', 'SH', 'RS', 'Rent', 'Indirect Deal › Brokerage Company / BC Sales / Contracted / Statue', 'Archive › Rev / Exception / Contract / Statue',
  'Cancel Details › Termination', 'Year and Month (calculated from the reservation date)', 'Age (calculated from the birth date)', 'Empty spacer columns'];

function ddbId_() { return String(PropertiesService.getScriptProperties().getProperty('DEALS_DB_ID') || '').trim(); }
function ddbOpen_() { const id = ddbId_(); return id ? SpreadsheetApp.openById(id) : null; }
function ddbColIndex_() { const o = {}; DDB_COLS_.forEach((c, i) => o[c[1]] = i + 1); return o; }

/* Admins: the new file's status (created or not, link, deals in it). */
function getDealsDbInfo(authToken) {
  targetsRequireAdmin_(authToken);
  const id = ddbId_();
  if (!id) return { created: false, columns: DDB_COLS_.length, groups: [...new Set(DDB_COLS_.map(c => c[0]))], removed: DDB_REMOVED_ };
  try {
    const ss = SpreadsheetApp.openById(id), sh = ss.getSheetByName('Deals');
    const last = sh ? sh.getLastRow() : 0;
    let deals = 0;
    if (last >= DDB_FIRST_ROW_) deals = sh.getRange(DDB_FIRST_ROW_, 1, last - DDB_FIRST_ROW_ + 1, 1).getValues().filter(r => String(r[0]).trim()).length;
    return { created: true, id: id, url: ss.getUrl(), name: ss.getName(), deals: deals, columns: DDB_COLS_.length, removed: DDB_REMOVED_ };
  } catch (e) {
    return { created: true, id: id, error: 'The file cannot be opened: ' + (e && e.message || e) };
  }
}

/* Admins: create the new file (once). */
function setupDealsDb(authToken) {
  const user = targetsRequireAdmin_(authToken);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    if (ddbId_()) return getDealsDbInfo(authToken);
    const ss = SpreadsheetApp.create('EDEN Deals DB');
    try { ss.setSpreadsheetTimeZone('Africa/Cairo'); } catch (e) {}
    const deals = ss.getSheets()[0]; deals.setName('Deals');
    const lists = ss.insertSheet('Lists'), guide = ss.insertSheet('Guide'), log = ss.insertSheet('Log');
    const n = DDB_COLS_.length;
    if (deals.getMaxColumns() < n) deals.insertColumnsAfter(deals.getMaxColumns(), n - deals.getMaxColumns());
    if (deals.getMaxRows() < DDB_ROWS_ + 2) deals.insertRowsAfter(deals.getMaxRows(), DDB_ROWS_ + 2 - deals.getMaxRows());

    // row 1: groups (merged, coloured) · row 2: column names
    deals.getRange(DDB_HEADER_ROW_, 1, 1, n).setValues([DDB_COLS_.map(c => c[1])]).setFontWeight('bold').setFontColor('#ffffff')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    let start = 0;
    for (let i = 1; i <= n; i++) {
      if (i === n || DDB_COLS_[i][0] !== DDB_COLS_[start][0]) {
        const g = DDB_COLS_[start][0], color = DDB_GROUP_COLORS_[g] || '#616161';
        const top = deals.getRange(1, start + 1, 1, i - start);
        if (i - start > 1) top.merge();
        top.getCell(1, 1).setValue(g);
        top.setBackground(color).setFontColor('#ffffff').setFontWeight('bold').setFontSize(9);
        deals.getRange(DDB_HEADER_ROW_, start + 1, 1, i - start).setBackground(color);
        start = i;
      }
    }
    deals.setRowHeight(DDB_HEADER_ROW_, 40);
    deals.setFrozenRows(DDB_HEADER_ROW_); deals.setFrozenColumns(6);
    deals.getRange(1, 1, DDB_ROWS_ + 2, n).setFontFamily('Arial').setFontSize(10);
    deals.getRange(1, 1, 1, n).setFontSize(9);

    // formats per column type
    const body = c => deals.getRange(DDB_FIRST_ROW_, c, DDB_ROWS_, 1);
    DDB_COLS_.forEach((col, i) => {
      const c = i + 1, t = col[2];
      if (t === 'date') body(c).setNumberFormat('dd-mmm-yyyy');
      else if (t === 'datetime') body(c).setNumberFormat('dd-mmm-yyyy hh:mm');
      else if (t === 'money') body(c).setNumberFormat('#,##0');
      else if (t === 'number') body(c).setNumberFormat('#,##0.##');
      else if (t === 'percent') body(c).setNumberFormat('0.0#%');
      else if (t === 'phone' || t === 'id' || t === 'text' || t === 'email' || t === 'url') body(c).setNumberFormat('@');
      else if (t === 'check') body(c).insertCheckboxes().setHorizontalAlignment('center');
      const w = { id: 110, phone: 120, email: 200, date: 100, datetime: 130, money: 110, number: 80, percent: 80, check: 90, url: 120 }[t] || 130;
      deals.setColumnWidth(c, ['Client Name', 'Residence Address', 'ID Address', 'Notes', 'Cancel Reason'].indexOf(col[1]) >= 0 ? 220 : w);
      deals.getRange(DDB_HEADER_ROW_, c).setNote(col[4] + (col[4] ? '\n' : '') + 'Type: ' + t + ' · Filled by: ' + col[3]);
    });
    deals.getRange(DDB_FIRST_ROW_, 1, DDB_ROWS_, n).setBorder(null, null, null, null, true, true, '#e6dece', SpreadsheetApp.BorderStyle.SOLID);

    // Lists: from the deals we already have
    let rows = [];
    try { const d = dashCacheRead_('ALL', 6 * 3600000) || buildDashboardData_({ role: 'Admin' }); rows = (d && d.rows) || []; } catch (e) {}
    const uniq = f => Array.from(new Set(rows.map(r => String(r[f] == null ? '' : r[f]).trim()).filter(v => v && !/^not assigned$/i.test(v)))).sort((a, b) => a.localeCompare(b));
    const listCols = DDB_LISTS_.map(([name, f]) => [name, DDB_FIXED_LISTS_[name] || (f ? uniq(f) : [])]);
    const maxLen = Math.max(1, ...listCols.map(x => x[1].length));
    if (lists.getMaxColumns() < listCols.length) lists.insertColumnsAfter(lists.getMaxColumns(), listCols.length - lists.getMaxColumns());
    if (lists.getMaxRows() < maxLen + 200) lists.insertRowsAfter(lists.getMaxRows(), maxLen + 200 - lists.getMaxRows());
    lists.getRange(1, 1, 1, listCols.length).setValues([listCols.map(x => x[0])]).setFontWeight('bold').setBackground('#1F1B16').setFontColor('#ffffff');
    lists.getRange(2, 1, maxLen, listCols.length).setValues(Array.from({ length: maxLen }, (_, i) => listCols.map(x => x[1][i] || '')));
    lists.setFrozenRows(1); lists.getRange(1, 1, maxLen + 1, listCols.length).setFontFamily('Arial');
    for (let j = 1; j <= listCols.length; j++) lists.setColumnWidth(j, 170);

    // dropdowns on the Deals columns that have a list
    const at = ddbColIndex_();
    listCols.forEach(([name], j) => {
      const c = at[name]; if (!c) return;
      const letter = String.fromCharCode(65 + (j % 26)); const colA1 = (j >= 26 ? String.fromCharCode(64 + Math.floor(j / 26)) : '') + letter;
      const rule = SpreadsheetApp.newDataValidation().requireValueInRange(lists.getRange(colA1 + '2:' + colA1), true)
        .setAllowInvalid(name !== 'Project' && name !== 'Status' && name !== 'Currency').build();
      body(c).setDataValidation(rule);
    });

    // Guide
    const g = [['EDEN Deals DB — what each column is', '', '', '', ''],
      ['One row per deal · row 2 holds the column names · every deal has a fixed Deal ID · deals are added and changed from the EDEN BI dashboard only.', '', '', '', ''],
      ['', '', '', '', ''], ['Group', 'Column', 'Type', 'Filled by', 'What it is']]
      .concat(DDB_COLS_.map(c => [c[0], c[1], c[2], c[3], c[4]]))
      .concat([['', '', '', '', ''], ['Not carried over from the old Deals sheet', '', '', '', '']]).concat(DDB_REMOVED_.map(x => ['', x, '', '', '']));
    guide.getRange(1, 1, g.length, 5).setValues(g).setFontFamily('Arial').setWrap(true).setVerticalAlignment('top');
    guide.getRange(1, 1).setFontSize(14).setFontWeight('bold'); guide.getRange(2, 1).setFontColor('#6b6255');
    guide.getRange(4, 1, 1, 5).setFontWeight('bold').setBackground('#1F1B16').setFontColor('#ffffff');
    DDB_COLS_.forEach((c, i) => guide.getRange(5 + i, 1).setBackground(DDB_GROUP_COLORS_[c[0]] || '#616161').setFontColor('#ffffff').setFontWeight('bold'));
    guide.getRange(6 + DDB_COLS_.length, 1).setFontWeight('bold');
    [170, 220, 80, 100, 520].forEach((w, i) => guide.setColumnWidth(i + 1, w));
    guide.setFrozenRows(4);

    // Log
    log.getRange(1, 1, 1, 7).setValues([['Time', 'User', 'Deal ID', 'Action', 'Field', 'Old value', 'New value']]).setFontWeight('bold').setBackground('#1F1B16').setFontColor('#ffffff');
    log.setFrozenRows(1); log.getRange('A:A').setNumberFormat('dd-mmm-yyyy hh:mm:ss');

    // lock every tab to the account that runs EDEN BI
    ss.getSheets().forEach(sh => {
      try { const p = sh.protect().setDescription('EDEN BI only — change deals from the dashboard'); p.removeEditors(p.getEditors()); if (p.canDomainEdit()) p.setDomainEdit(false); } catch (e) {}
    });

    // keep it next to the old deals file
    try {
      const parents = (Drive.Files.get(SPREADSHEET_ID, { fields: 'parents', supportsAllDrives: true }).parents || []);
      if (parents.length) Drive.Files.update({}, ss.getId(), null, { addParents: parents.join(','), removeParents: 'root', supportsAllDrives: true });
    } catch (e) { console.warn('Deals DB folder: ' + e); }

    PropertiesService.getScriptProperties().setProperty('DEALS_DB_ID', ss.getId());
    auditLog_(user, 'Deals DB created', { details: ss.getName() + ' · ' + n + ' columns · ' + ss.getUrl() });
    return getDealsDbInfo(authToken);
  } finally { lock.releaseLock(); }
}
