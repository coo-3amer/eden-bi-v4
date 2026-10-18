/************************************************
 * EDEN BI V4 — ContractGen.js
 *
 * Fills the sale contract (Word) of a deal from the Deals sheet:
 * client, ID, address, unit, area, floor, price, down payment, remaining
 * (numbers + Arabic words) and Annex 2 (payment schedule).
 *
 * The template is a .docx in the owner's Google Drive (not in GitHub):
 *   "EDEN WALK contract template.docx"  — found by name, or set the Script
 *   Property CONTRACT_TEMPLATE_EDEN_WALK = its file id.
 * Markers in the template: {{CLIENT_NAME}} {{PRICE}} …; Annex 1 / Annex 2 images go under their titles.
 * The filled contract is saved in Drive › "EDEN BI Contracts" and also
 * returned to the page for download.
 ************************************************/

const CONTRACT_TEMPLATES_ = { 'EDENWALK': { name: 'EDEN WALK contract template.docx', prop: 'CONTRACT_TEMPLATE_EDEN_WALK' } };
const CONTRACT_FOLDER_ = 'EDEN BI Contracts';

function contractCanMake_(user) {
  const role = String(user && user.role || '').trim().toLowerCase();
  return role === 'admin' || role === 'super admin';
}
function contractTemplate_(project) {
  const key = String(project || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const t = CONTRACT_TEMPLATES_[key];
  if (!t) throw new Error('There is no contract template for ' + project + ' yet.');
  // Drive advanced service + UrlFetch only (already authorised — no new permissions)
  // the newest file with this name wins (so uploading a new version of the template just works);
  // the Script Property is only a fallback / manual override
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(t.prop + '_PIN') || '';
  if (!id) {
    const found = (Drive.Files.list({ q: "name='" + t.name.replace(/'/g, "\\'") + "' and trashed=false", fields: 'files(id,name,modifiedTime)', orderBy: 'modifiedTime desc', pageSize: 5 }).files || [])[0];
    id = found ? found.id : props.getProperty(t.prop);
    if (!id) throw new Error('Contract template not found. Upload "' + t.name + '" to the Google Drive of the account that runs EDEN BI.');
    if (found) props.setProperty(t.prop, id);
  }
  return { id: id, name: t.name };
}
function contractStatus(authToken) {
  const user = validateAuthToken_(authToken);
  if (!contractCanMake_(user)) return { can: false };
  const out = { can: true, templates: {} };
  Object.keys(CONTRACT_TEMPLATES_).forEach(k => { try { contractTemplate_(k); out.templates[k] = true; } catch (e) { out.templates[k] = false; } });
  return out;
}

/* ---------- Arabic number words (tafqeet) ---------- */
function arWords_(num) {
  num = Math.floor(Math.abs(Number(num) || 0));
  if (!num) return 'صفر';
  const ones = ['', 'واحد', 'اثنان', 'ثلاثة', 'أربعة', 'خمسة', 'ستة', 'سبعة', 'ثمانية', 'تسعة'];
  const teens = ['عشرة', 'أحد عشر', 'اثنا عشر', 'ثلاثة عشر', 'أربعة عشر', 'خمسة عشر', 'ستة عشر', 'سبعة عشر', 'ثمانية عشر', 'تسعة عشر'];
  const tens = ['', '', 'عشرون', 'ثلاثون', 'أربعون', 'خمسون', 'ستون', 'سبعون', 'ثمانون', 'تسعون'];
  const hund = ['', 'مائة', 'مائتان', 'ثلاثمائة', 'أربعمائة', 'خمسمائة', 'ستمائة', 'سبعمائة', 'ثمانمائة', 'تسعمائة'];
  const small = n => {   // 1..999
    const h = Math.floor(n / 100), r = n % 100, parts = [];
    if (h) parts.push(hund[h]);
    if (r) {
      if (r < 10) parts.push(ones[r]);
      else if (r < 20) parts.push(teens[r - 10]);
      else { const o = r % 10, t = Math.floor(r / 10); parts.push(o ? ones[o] + ' و' + tens[t] : tens[t]); }
    }
    return parts.join(' و');
  };
  const scales = [['ألف', 'ألفان', 'آلاف', 'ألفاً'], ['مليون', 'مليونان', 'ملايين', 'مليوناً'], ['مليار', 'ملياران', 'مليارات', 'ملياراً']];
  const groups = [];
  let n = num, i = -1;
  while (n > 0) {
    const g = n % 1000;
    if (g) {
      if (i < 0) groups.unshift(small(g));
      else {
        const s = scales[i], r = g % 100;
        let w;
        if (g === 1) w = s[0];
        else if (g === 2) w = s[1];
        else if (r >= 3 && r <= 10) w = small(g) + ' ' + s[2];
        else if (r >= 11) w = small(g) + ' ' + s[3];
        else w = small(g) + ' ' + s[0];
        groups.unshift(w);
      }
    }
    n = Math.floor(n / 1000); i++;
  }
  return groups.join(' و');
}
/* areas can have decimals: 25.5 -> "خمسة وعشرون فاصل خمسة" */
function arArea_(v) {
  const n = Math.round((Number(v) || 0) * 100) / 100, i = Math.floor(n), f = Math.round((n - i) * 100);
  if (!f) return arWords_(i);
  const fs = String(f).padStart(2, '0').replace(/0$/, '');
  return arWords_(i) + ' فاصل ' + arWords_(Number(fs));
}
function areaNum_(v) { const n = Math.round((Number(v) || 0) * 100) / 100; return n.toLocaleString('en-US', { maximumFractionDigits: 2 }); }
function arMoney_(v) { const n = Math.round(Number(v) || 0); return arWords_(n) + ' جنيه مصري فقط لا غير'; }
function enNum_(v) { return Math.round(Number(v) || 0).toLocaleString('en-US'); }
function arDay_(d) { return ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'][d.getDay()]; }
function arFloor_(f) {
  const t = String(f || '').trim(), k = t.toLowerCase();
  const map = [[/ground|^g$|أرضي|ارضي/, 'الأرضي'], [/first|^f1$|^1/, 'الأول'], [/second|^f2$|^2/, 'الثاني'], [/third|^f3$|^3/, 'الثالث'], [/fourth|^f4$|^4/, 'الرابع'],
    [/basement 2|b2/, 'البدروم الثاني'], [/basement|^b1?$/, 'البدروم'], [/roof/, 'السطح'], [/mezz/, 'الميزانين']];
  for (const [re, ar] of map) if (re.test(k)) return ar;
  return t || '..........';
}
function arNationality_(v) {
  const t = String(v || '').trim(), k = t.toLowerCase();
  const map = { egyptian: 'مصري', egypt: 'مصري', saudi: 'سعودي', 'saudi arabian': 'سعودي', emirati: 'إماراتي', kuwaiti: 'كويتي', qatari: 'قطري', jordanian: 'أردني', sudanese: 'سوداني', libyan: 'ليبي', iraqi: 'عراقي', syrian: 'سوري', lebanese: 'لبناني', palestinian: 'فلسطيني', yemeni: 'يمني', omani: 'عماني', bahraini: 'بحريني' };
  return map[k] || t;
}
function xmlEsc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

/* Annex 2: one Word table (right-to-left). rows = [[no, label, due, amount], …] */
function contractScheduleXml_(rows, total) {
  const cell = (txt, w, bold, fill) => '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/>' + (fill ? '<w:shd w:val="clear" w:color="auto" w:fill="' + fill + '"/>' : '') + '<w:vAlign w:val="center"/></w:tcPr>' +
    '<w:p><w:pPr><w:bidi/><w:spacing w:before="40" w:after="40"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>' + (bold ? '<w:b/><w:bCs/>' : '') +
    '<w:sz w:val="22"/><w:szCs w:val="22"/><w:rtl/></w:rPr><w:t xml:space="preserve">' + xmlEsc_(txt) + '</w:t></w:r></w:p></w:tc>';
  const W = [700, 3600, 2200, 2500];
  const tr = (cells, bold, fill) => '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + cells.map((c, i) => cell(c, W[i], bold, fill)).join('') + '</w:tr>';
  const b = '<w:top w:val="single" w:sz="6" w:space="0" w:color="8F6D2C"/><w:left w:val="single" w:sz="6" w:space="0" w:color="8F6D2C"/><w:bottom w:val="single" w:sz="6" w:space="0" w:color="8F6D2C"/><w:right w:val="single" w:sz="6" w:space="0" w:color="8F6D2C"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFA77A"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFA77A"/>';
  return '<w:tbl><w:tblPr><w:bidiVisual/><w:tblW w:w="' + W.reduce((a, x) => a + x, 0) + '" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>' + b + '</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>' +
    '<w:tblGrid>' + W.map(w => '<w:gridCol w:w="' + w + '"/>').join('') + '</w:tblGrid>' +
    tr(['م', 'البيان', 'تاريخ الاستحقاق', 'المبلغ (جنيه مصري)'], true, 'EFE4CC') +
    rows.map(r => tr(r, false, '')).join('') +
    tr(['', 'الإجمالي', '', enNum_(total)], true, 'F6F0E4') + '</w:tbl>';
}

/* Reads one deal row (by header names). */
function contractDeal_(row, unitCode) {
  const sheet = getMainDealsSheet_();
  const cols = mainDealsColumns_(sheet);
  if (!row || row < DATA_START_ROW || row > sheet.getLastRow()) throw new Error('Deal row not found.');
  const disp = sheet.getRange(row, 1, 1, cols.lastCol).getDisplayValues()[0];
  const vals = sheet.getRange(row, 1, 1, cols.lastCol).getValues()[0];
  const at = names => { const c = mainDealsCol_(cols, names); return c ? String(disp[c - 1] || '').trim() : ''; };
  const num = names => { const c = mainDealsCol_(cols, names); if (!c) return 0; const v = vals[c - 1]; return typeof v === 'number' ? v : Number(String(v || '').replace(/[^\d.\-]/g, '')) || 0; };
  const dateOf = names => { const c = mainDealsCol_(cols, names); const v = c ? vals[c - 1] : ''; return v instanceof Date && !isNaN(v) ? v : (dealDate_(String(disp[c - 1] || '')) || null); };
  const d = {
    unit: at(['Unit Code']), project: at(['Project']), client: at(['Final Client Name', 'Client Name']).replace(/\s*-?\s*EW-[A-Z0-9/.\- ]+$/i, '').trim(),
    nationality: at(['client info|nationality', 'Nationality']), idType: at(['ID Type']), idNumber: at(['ID Number']),
    address: at(['Residence Address']) || at(['ID Address']), mobile: at(['client info|contact', 'Contact']), email: at(['Email']),
    floor: at(['unit info|floor', 'Floor']), area: num(['unit info|area', 'Area']), outArea: num(['OutArea', 'Out Area']),
    price: num(['price details|final price', 'Final Price']), dp: num(['Actual DP', 'Acctual DP']),
    contractDate: dateOf(['contract details|date'])
  };
  if (!d.unit || d.unit.toUpperCase().replace(/\.+$/, '') !== String(unitCode || '').trim().toUpperCase().replace(/\.+$/, '')) throw new Error('The Deals sheet changed (rows moved). Refresh and try again.');
  return d;
}

/* English schedule names -> Arabic for Annex 2 (the wizard lets you edit them). */
function contractArLabel_(name) {
  const t = String(name || '').trim();
  const tail = (t.match(/\s*\(.*\)\s*$/) || [''])[0];
  const b = t.replace(/\s*\(.*\)\s*$/, '');
  let m;
  if (/^down payment$/i.test(b)) return 'مقدم حجز وتعاقد' + tail;
  if (/^part of dp$/i.test(b)) return 'جزء من المقدم' + tail;
  if ((m = b.match(/^remain of dp\s*(\d*)$/i))) return 'باقي المقدم' + (m[1] ? ' ' + m[1] : '') + tail;
  if ((m = b.match(/^bulk(?: payment)?\s*(\d*)$/i))) return 'دفعة إضافية' + (m[1] ? ' ' + m[1] : '') + tail;
  if ((m = b.match(/^marketing expenses$/i))) return 'مصاريف تسويق' + tail;
  if ((m = b.match(/^SA\s*(\d+)$/i))) return 'القسط نصف السنوي ' + m[1] + tail;
  if ((m = b.match(/^Q\s*(\d+)$/i))) return 'القسط الربع سنوي ' + m[1] + tail;
  if ((m = b.match(/^A\s*(\d+)$/i))) return 'القسط السنوي ' + m[1] + tail;
  return t;
}
function contractIsDp_(name) { return /^(down payment|part of dp|remain of dp)/i.test(String(name || '').trim()); }

/* Annex 2 rows (maintenance left out): the unit's schedule if one was entered, else DP + the EDEN WALK yearly plan.
   -> {rows: [{label, due:'yyyy-mm-dd'|'', dueText, amount}], dp, source} */
function contractSchedule_(d) {
  const tz = Session.getScriptTimeZone() || 'Africa/Cairo', out = [];
  let dp = 0, source = '';
  try {
    const key = instKey_(d.project, d.unit);
    const rows = instRows_(instSheet_(INST_SHEET_, INST_HEAD_, false)).filter(r => String(r[0]) === key && !/^M/i.test(String(r[5])))
      .sort((a, b) => (a[7] instanceof Date ? a[7].getTime() : 0) - (b[7] instanceof Date ? b[7].getTime() : 0) || (String(a[5]) === '0' ? -1 : String(b[5]) === '0' ? 1 : 0));
    rows.forEach(r => {
      const amt = Number(r[8]) || 0;
      if (contractIsDp_(r[6])) dp += amt;
      out.push({ label: contractArLabel_(r[6]), due: r[7] instanceof Date ? Utilities.formatDate(r[7], tz, 'yyyy-MM-dd') : '', dueText: '', amount: amt });
    });
    if (out.length) source = 'schedule';
  } catch (e) {}
  if (!out.length) {
    dp = d.dp;
    out.push({ label: 'مقدم حجز وتعاقد', due: d.contractDate ? Utilities.formatDate(d.contractDate, tz, 'yyyy-MM-dd') : '', dueText: '', amount: d.dp });
    if (typeof ewRead_ === 'function') {
      try {
        const u = ewRead_().units.find(x => ewKey_(x.unit) === ewKey_(d.unit));
        if (u) Object.keys(u.plan).sort().forEach(y => { const v = Number(u.plan[y][0]) || 0; if (v) out.push({ label: 'أقساط عام ' + y, due: '', dueText: 'خلال عام ' + y, amount: v }); });
        if (out.length > 1) source = 'partners';
      } catch (e) {}
    }
  }
  return { rows: out, dp: dp, source: source };
}

/* Wizard step 0: everything the wizard pre-fills. req = {row, unitCode} */
function contractDraft(authToken, req) {
  const user = validateAuthToken_(authToken);
  if (!contractCanMake_(user)) throw new Error('Only admins can generate contracts.');
  req = req || {};
  const d = contractDeal_(Number(req.row), req.unitCode);
  const s = contractSchedule_(d);
  let template = true;
  try { contractTemplate_(d.project); } catch (e) { template = false; }
  const tz = Session.getScriptTimeZone() || 'Africa/Cairo';
  return {
    template: template, project: d.project, row: Number(req.row),
    client: { name: d.client, idType: /pass/i.test(d.idType) ? 'passport' : 'id', idNumber: d.idNumber, nationality: arNationality_(d.nationality) || 'مصري',
      address: d.address, mobile: d.mobile, email: d.email },
    unit: { code: d.unit, area: d.area, outArea: d.outArea, floor: arFloor_(d.floor) === '..........' ? '' : arFloor_(d.floor) },
    price: d.price, dp: s.dp || d.dp, schedule: s.rows, scheduleSource: s.source,
    date: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')
  };
}

/* ---------- Word XML helpers (pure: tested outside Apps Script) ---------- */
function contractParas_(xml) {
  const out = [], re = /<w:p[ >][\s\S]*?<\/w:p>/g;
  let m;
  while ((m = re.exec(xml))) out.push({ s: m.index, e: m.index + m[0].length, x: m[0], t: (m[0].match(/<w:t[^>]*>[^<]*<\/w:t>/g) || []).map(z => z.replace(/<[^>]+>/g, '')).join('') });
  return out;
}
/* Puts `inner` under the title "ملحق رقم (n)", replacing the blank lines the template keeps there. */
function contractAnnexPut_(xml, n, inner) {
  const ps = contractParas_(xml);
  let i = -1;
  const re = new RegExp('^\\s*ملحق\\s*رقم\\s*\\(\\s*' + n + '\\s*\\)');
  ps.forEach((q, k) => { if (re.test(q.t)) i = k; });   // the annex title itself, not a clause that mentions it
  if (i < 0) return xml;
  let j = i + 1;
  while (j < ps.length && !ps[j].t.trim() && !/sectPr|w:drawing|w:type="page"|\{\{/.test(ps[j].x)) j++;
  const s = ps[i].e, e = j > i + 1 ? ps[j - 1].e : ps[i].e;
  return xml.slice(0, s) + inner + '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' + xml.slice(e);
}
function contractImageXml_(im, id) {
  return '<w:p><w:pPr><w:bidi/><w:spacing w:before="240" w:after="0"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + im.cx + '" cy="' + im.cy + '"/><wp:docPr id="' + id + '" name="EDEN annex ' + id + '"/>' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:nvPicPr><pic:cNvPr id="' + id + '" name="annex' + id + '"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="' + im.rid + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
    '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + im.cx + '" cy="' + im.cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
}
/* vals: {MARKER: text}; opts: {passport, images: {1: {rid,cx,cy}, 2: {…}}, table: {rows, total}} */
function contractFillXml_(xml, vals, opts) {
  opts = opts || {};
  if (opts.passport) {   // "بطاقة رقم:" -> "جواز سفر رقم:" in the buyer's ID line
    const p = contractParas_(xml).find(q => q.x.indexOf('{{ID_NUMBER}}') >= 0);
    if (p) xml = xml.slice(0, p.s) + p.x.replace('>بطاقة<', '>جواز سفر<').replace('>بطاقة رقم', '>جواز سفر رقم') + xml.slice(p.e);
  }
  const im = opts.images || {};
  if (im[1]) xml = contractAnnexPut_(xml, 1, contractImageXml_(im[1], 9101));
  const hasMarker = xml.indexOf('{{SCHEDULE}}') >= 0;
  if (im[2]) xml = contractAnnexPut_(xml, 2, contractImageXml_(im[2], 9102));
  else if (opts.table && !hasMarker) xml = contractAnnexPut_(xml, 2, contractScheduleXml_(opts.table.rows, opts.table.total) + '<w:p/>');
  Object.keys(vals).forEach(k => { xml = xml.split('{{' + k + '}}').join(xmlEsc_(vals[k])); });
  const i = xml.indexOf('{{SCHEDULE}}');   // older template: a marker paragraph under Annex 2
  if (i >= 0) {
    const s = Math.max(xml.lastIndexOf('<w:p ', i), xml.lastIndexOf('<w:p>', i));
    const e = xml.indexOf('</w:p>', i) + 6;
    const inner = im[2] ? contractImageXml_(im[2], 9102) : (opts.table ? contractScheduleXml_(opts.table.rows, opts.table.total) : '');
    xml = xml.slice(0, s) + inner + '<w:p/>' + xml.slice(e);
  }
  return xml;
}
function contractImageFit_(w, h) {   // EMU, keep the aspect ratio
  const MW = 5760000, MH = 6300000;   // 16 x 17.5 cm: leaves room for the annex title, header and footer
  w = Number(w) || 1600; h = Number(h) || 1200;
  const k = Math.min(MW / w, MH / h);
  return { cx: Math.round(w * k), cy: Math.round(h * k) };
}

/* Builds the contract from the wizard. req = {row, unitCode, date, data: {client, unit, price, dp, schedule, drawing}} */
function generateDealContract(authToken, req) {
  const user = validateAuthToken_(authToken);
  if (!contractCanMake_(user)) throw new Error('Only admins can generate contracts.');
  req = req || {};
  const d = contractDeal_(Number(req.row), req.unitCode);
  const draft = req.data || contractDraft(authToken, { row: req.row, unitCode: req.unitCode });
  const c = draft.client || {}, u = draft.unit || {};
  const file = contractTemplate_(d.project);
  const date = (req.date && instDateIn_(req.date)) || d.contractDate || new Date();
  const tz = Session.getScriptTimeZone() || 'Africa/Cairo';
  const price = Math.round(Number(draft.price) || 0), dp = Math.round(Number(draft.dp) || 0), rest = Math.max(0, price - dp);
  const area = Number(u.area) || 0, outArea = Number(u.outArea) || 0;
  const fmtDue = r => { const t = String(r.due || '').trim(); const m = t.match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? m[3] + '/' + m[2] + '/' + m[1] : (r.dueText || t || '..........'); };
  const sched = (draft.schedule || []).filter(r => String(r.label || '').trim() || Number(r.amount));
  const rows = sched.map((r, i) => [String(i + 1), String(r.label || '').trim() || '..........', fmtDue(r), enNum_(r.amount)]);
  const schedTotal = sched.reduce((a, r) => a + (Number(r.amount) || 0), 0);
  const w = draft.words || {};
  const pick = (k, auto) => String(w[k] || '').trim() || auto;
  const vals = {
    DAY_NAME: pick('DAY_NAME', arDay_(date)), DD: Utilities.formatDate(date, tz, 'dd'), MM: Utilities.formatDate(date, tz, 'MM'), YYYY: Utilities.formatDate(date, tz, 'yyyy'),
    CLIENT_NAME: String(c.name || '').trim(), NATIONALITY: String(c.nationality || '').trim(), ID_NUMBER: String(c.idNumber || '').trim(),
    ADDRESS: String(c.address || '').trim(), EMAIL: String(c.email || '').trim(), MOBILE: String(c.mobile || '').trim(),
    UNIT_CODE: String(u.code || d.unit).trim(), AREA: area ? areaNum_(area) : '', AREA_WORDS: area ? pick('AREA_WORDS', arArea_(area)) : '', FLOOR: String(u.floor || '').trim(),
    OUT_AREA: outArea ? areaNum_(outArea) + ' متراً مربعاً' : 'لا يوجد', OUT_AREA_WORDS: outArea ? pick('OUT_AREA_WORDS', arArea_(outArea) + ' متراً مربعاً') : '—',
    PRICE: enNum_(price), PRICE_WORDS: pick('PRICE_WORDS', arMoney_(price)), DP: enNum_(dp), DP_WORDS: pick('DP_WORDS', arMoney_(dp)), REST: enNum_(rest), REST_WORDS: pick('REST_WORDS', arMoney_(rest))
  };
  const missing = ['CLIENT_NAME', 'ID_NUMBER', 'ADDRESS', 'MOBILE', 'AREA', 'FLOOR'].filter(k => !vals[k]);
  Object.keys(vals).forEach(k => { if (vals[k] === '' || vals[k] == null) vals[k] = '..........'; });

  const res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + file.id + '?alt=media&supportsAllDrives=true',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    PropertiesService.getScriptProperties().deleteProperty(CONTRACT_TEMPLATES_[String(d.project).toUpperCase().replace(/[^A-Z0-9]/g, '')].prop);
    throw new Error('Could not read the contract template from Drive (' + res.getResponseCode() + '). Try again.');
  }
  // Annex 1 (drawing) and Annex 2 (payment plan) are images the user uploads
  const images = {}, blobs = [];
  [[1, draft.drawing], [2, draft.planImage]].forEach(([n, dr]) => {
    if (!dr || !dr.b64) return;
    const ext = /jpe?g/i.test(dr.mime || '') ? 'jpeg' : 'png', file = 'eden_annex' + n + '.' + ext;
    blobs.push(Utilities.newBlob(Utilities.base64Decode(String(dr.b64).replace(/^data:[^,]*,/, '')), 'image/' + ext, 'word/media/' + file));
    images[n] = Object.assign({ rid: 'rIdEdenAnnex' + n, ext: ext, file: file }, contractImageFit_(dr.w, dr.h));
  });
  const useTable = !images[2] && draft.useTable && rows.length;
  const parts = Utilities.unzip(res.getBlob().setContentType('application/zip'));
  const out = parts.map(p => {
    const name = p.getName();
    if (name === 'word/document.xml')
      return Utilities.newBlob(contractFillXml_(p.getDataAsString('UTF-8'), vals, { passport: c.idType === 'passport', images: images, table: useTable ? { rows: rows, total: schedTotal } : null }), 'application/xml', name);
    if (blobs.length && name === 'word/_rels/document.xml.rels')
      return Utilities.newBlob(p.getDataAsString('UTF-8').replace('</Relationships>', Object.keys(images).map(n => '<Relationship Id="' + images[n].rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' + images[n].file + '"/>').join('') + '</Relationships>'), 'application/xml', name);
    if (blobs.length && name === '[Content_Types].xml') {
      let t = p.getDataAsString('UTF-8');
      Object.keys(images).forEach(n => { const ext = images[n].ext; if (t.indexOf('Extension="' + ext + '"') < 0) t = t.replace('<Default ', '<Default Extension="' + ext + '" ContentType="image/' + ext + '"/><Default '); });
      return Utilities.newBlob(t, 'application/xml', name);
    }
    return p;
  });
  blobs.forEach(b => out.push(b));
  const fname = 'عقد ' + vals.UNIT_CODE + ' - ' + vals.CLIENT_NAME + ' - ' + Utilities.formatDate(date, tz, 'yyyy-MM-dd') + '.docx';
  const docx = Utilities.zip(out, fname).setContentType('application/vnd.openxmlformats-officedocument.wordprocessingml.document');

  let url = '';
  try {
    const folder = backupFolder_(CONTRACT_FOLDER_, null);
    const f = Drive.Files.create({ name: fname, parents: [folder], mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }, docx, { fields: 'id,webViewLink' });
    url = f.webViewLink || ('https://drive.google.com/file/d/' + f.id + '/view');
  } catch (e) { console.warn('Contract save: ' + e); }
  auditLog_(user, 'Contract generated', { project: d.project, unit: vals.UNIT_CODE, client: vals.CLIENT_NAME, details: 'Price ' + enNum_(price) + ' · DP ' + enNum_(dp) + (images[1] ? ' · Annex 1 image' : '') + (images[2] ? ' · Annex 2 image' : useTable ? ' · Annex 2 table ' + rows.length + ' rows' : '') + (missing.length ? ' · missing: ' + missing.join(', ') : ''), ref: 'Deals row ' + req.row });
  return { name: fname, url: url, b64: Utilities.base64Encode(docx.getBytes()), missing: missing };
}
