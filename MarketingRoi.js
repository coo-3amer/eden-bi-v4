/************************************************
 * EDEN BI V4 — MarketingRoi.js
 *
 * 1) Funnel from EDEN CRM: leads (by CRM source / campaign) → contacted →
 *    site visit → won in CRM → became a deal in the Deals sheet (matched by
 *    phone number). Needs the CRM data spreadsheet (Script Property
 *    CRM_SPREADSHEET_ID, set from the dashboard).
 * 2) Return on spend: deals and value per Main Source / Campaign from
 *    Deals, against the marketing spend entered per month in the hidden
 *    sheet "_Marketing Spend" (Month | Source | Campaign | Amount EGP).
 * Admins only.
 ************************************************/

const MS_SHEET_ = '_Marketing Spend';
const MS_HEAD_ = ['Month', 'Source', 'Campaign', 'Amount EGP', 'By', 'At'];
const CRM_STAGES_ = ['new', 'contacted', 'qualified', 'site visit', 'negotiation', 'won', 'lost'];

function msSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(MS_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(MS_SHEET_);
    sh.getRange(1, 1, 1, MS_HEAD_.length).setValues([MS_HEAD_]).setFontWeight('bold');
    sh.setFrozenRows(1);
    try { sh.hideSheet(); } catch (e) {}
  }
  return sh;
}
function mrPhoneKey_(v) {
  return String(v == null ? '' : v).split(/[\/,;|\n]| - | – |\s{2,}/).map(x => x.replace(/\D/g, '')).filter(x => x.length >= 9).map(x => x.slice(-10));
}
function mrMonthOf_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, 'Africa/Cairo', 'yyyy-MM');
  const s = String(v || ''); const m = s.match(/^(\d{4})-(\d{2})/); if (m) return m[1] + '-' + m[2];
  const d = new Date(s); return isNaN(d) ? '' : Utilities.formatDate(d, 'Africa/Cairo', 'yyyy-MM');
}
function mrCrmId_() {
  const raw = String(PropertiesService.getScriptProperties().getProperty('CRM_SPREADSHEET_ID') || '').trim();
  const m = raw.match(/[-\w]{25,}/); return m ? m[0] : '';
}

function getMarketingRoi(authToken, req) {
  targetsRequireAdmin_(authToken);
  req = req || {};
  const from = String(req.from || '').slice(0, 7), to = String(req.to || '').slice(0, 7);
  const inRange = m => !!m && (!from || m >= from) && (!to || m <= to);
  const rows = dailyRows_().filter(r => !dailyCancelled_(r) && Number(r.amount) > 0);
  const dealMonth = r => (Number(r.year) && Number(r.month)) ? r.year + '-' + String(r.month).padStart(2, '0') : '';

  // ---- 2) return on spend (Deals) ----
  const spendRows = (() => { const sh = msSheet_(), last = sh.getLastRow(); return last < 2 ? [] : sh.getRange(2, 1, last - 1, 4).getValues(); })();
  const spend = spendRows.map(r => ({ month: mrMonthOf_(r[0]) || String(r[0]).slice(0, 7), source: String(r[1] || '').trim(), campaign: String(r[2] || '').trim(), amount: Number(r[3]) || 0 }));
  const roi = {};
  const key = (s, c) => s + '||' + (c || '');
  rows.filter(r => inRange(dealMonth(r))).forEach(r => {
    const s = String(r.mainSource || '—').trim() || '—', egp = dailyCur_(r) === 'EGP';
    [key(s, ''), key(s, String(r.campaignName || '').trim())].forEach((k, i) => {
      if (i === 1 && !String(r.campaignName || '').trim()) return;
      const o = roi[k] || (roi[k] = { source: s, campaign: i ? String(r.campaignName).trim() : '', deals: 0, egp: 0, usd: 0, spend: 0, leads: 0 });
      o.deals++; o[egp ? 'egp' : 'usd'] += Number(r.amount) || 0;
    });
  });
  spend.filter(x => inRange(x.month)).forEach(x => {
    [key(x.source, ''), x.campaign ? key(x.source, x.campaign) : null].filter(Boolean).forEach(k => {
      const o = roi[k] || (roi[k] = { source: x.source, campaign: k.split('||')[1], deals: 0, egp: 0, usd: 0, spend: 0, leads: 0 });
      o.spend += x.amount;
    });
  });

  // ---- 1) CRM funnel ----
  let crm = { linked: false };
  const id = mrCrmId_();
  if (id) {
    try {
      const sh = SpreadsheetApp.openById(id).getSheetByName('Leads');
      if (!sh) throw new Error('No "Leads" tab in the CRM file.');
      const v = sh.getDataRange().getValues(), h = v[0].map(x => String(x).trim());
      const at = n => h.indexOf(n);
      const iC = at('createdAt'), iP = at('phone'), iP2 = at('phone2'), iK = at('phoneKey'), iS = at('source'), iCa = at('campaign'), iSt = at('stage'), iO = at('ownerEmail'), iV = at('dealValue');
      // phone → deals in BI (any time after the lead)
      const byPhone = {};
      rows.forEach(r => mrPhoneKey_(r.mobile).forEach(k => (byPhone[k] = byPhone[k] || []).push(r)));
      const fun = {}, stageRank = s => { const i = CRM_STAGES_.indexOf(String(s || '').trim().toLowerCase()); return i < 0 ? 0 : i; };
      let total = 0;
      v.slice(1).forEach(L => {
        const m = mrMonthOf_(L[iC]);
        if (!inRange(m)) return;
        total++;
        const src = String(L[iS] || '—').trim() || '—', camp = iCa >= 0 ? String(L[iCa] || '').trim() : '';
        const st = String(L[iSt] || '').trim().toLowerCase(), rk = stageRank(st);
        const phones = [].concat(mrPhoneKey_(L[iP]), iP2 >= 0 ? mrPhoneKey_(L[iP2]) : [], iK >= 0 ? mrPhoneKey_(L[iK]) : []);
        const deals = []; phones.forEach(k => (byPhone[k] || []).forEach(d => { if (deals.indexOf(d) < 0) deals.push(d); }));
        [key(src, ''), camp ? key(src, camp) : null].filter(Boolean).forEach(k => {
          const o = fun[k] || (fun[k] = { source: src, campaign: k.split('||')[1], leads: 0, contacted: 0, visit: 0, won: 0, lost: 0, deals: 0, egp: 0, usd: 0 });
          o.leads++;
          if (rk >= 1 && st !== 'lost') o.contacted++; else if (st === 'lost') o.contacted++;
          if (rk >= 3 && st !== 'lost') o.visit++;
          if (st === 'won') o.won++;
          if (st === 'lost') o.lost++;
          if (deals.length) { o.deals++; deals.forEach(d => { o[dailyCur_(d) === 'EGP' ? 'egp' : 'usd'] += Number(d.amount) || 0; }); }
        });
      });
      // leads per source for cost per lead
      Object.keys(fun).forEach(k => { if (roi[k]) roi[k].leads = fun[k].leads; });
      crm = { linked: true, total: total, rows: Object.values(fun) };
    } catch (e) { crm = { linked: false, error: String(e && e.message || e), id: id }; }
  }
  return { from: from, to: to, roi: Object.values(roi), crm: crm, spend: spend.filter(x => inRange(x.month)),
    sources: Array.from(new Set(rows.map(r => String(r.mainSource || '').trim()).filter(Boolean))).sort(),
    campaigns: Array.from(new Set(rows.map(r => String(r.campaignName || '').trim()).filter(Boolean))).sort(), crmId: id };
}

/* Spend for one month: items [{source, campaign, amount}] replace that month's rows. */
function saveMarketingSpend(authToken, req) {
  const user = targetsRequireAdmin_(authToken);
  req = req || {};
  const month = String(req.month || '').slice(0, 7);
  if (!/^\d{4}-\d{2}$/.test(month)) throw new Error('Pick a month.');
  const items = (Array.isArray(req.items) ? req.items : []).map(i => ({ source: String(i.source || '').trim(), campaign: String(i.campaign || '').trim(), amount: Math.max(0, Number(i.amount) || 0) })).filter(i => i.source && i.amount > 0);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = msSheet_(), last = sh.getLastRow();
    const keep = (last < 2 ? [] : sh.getRange(2, 1, last - 1, MS_HEAD_.length).getValues()).filter(r => (mrMonthOf_(r[0]) || String(r[0]).slice(0, 7)) !== month);
    const now = new Date(), who = user.name || user.username;
    const norm = r => [r[0] instanceof Date ? mrMonthOf_(r[0]) : String(r[0]).slice(0, 7), r[1], r[2], r[3], r[4], r[5]];
    const all = keep.map(norm).concat(items.map(i => [month, i.source, i.campaign, i.amount, who, now]));
    sh.getRange('A:A').setNumberFormat('@');
    if (all.length) sh.getRange(2, 1, all.length, MS_HEAD_.length).setValues(all);
    const extra = (last - 1) - all.length; if (extra > 0) sh.getRange(2 + all.length, 1, extra, MS_HEAD_.length).clearContent();
    auditLog_(user, 'Marketing spend saved', { details: month + ' · ' + items.length + ' lines · ' + Math.round(items.reduce((a, i) => a + i.amount, 0)).toLocaleString('en-US') + ' EGP' });
  } finally { lock.releaseLock(); }
  return { ok: true };
}

function setCrmSheet(authToken, urlOrId) {
  const user = targetsRequireAdmin_(authToken);
  const m = String(urlOrId || '').match(/[-\w]{25,}/);
  if (!m) { PropertiesService.getScriptProperties().deleteProperty('CRM_SPREADSHEET_ID'); return { ok: true, cleared: true }; }
  let ss;
  try { ss = SpreadsheetApp.openById(m[0]); } catch (e) { throw new Error('Cannot open that file with the account running EDEN BI. Share the CRM data sheet with it (Viewer is enough).'); }
  if (!ss.getSheetByName('Leads')) throw new Error('That file has no "Leads" tab — it is not the EDEN CRM data sheet.');
  PropertiesService.getScriptProperties().setProperty('CRM_SPREADSHEET_ID', m[0]);
  auditLog_(user, 'CRM linked', { details: ss.getName() });
  return { ok: true, name: ss.getName() };
}
