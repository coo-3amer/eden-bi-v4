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
 * Markers in the template: {{CLIENT_NAME}} {{PRICE}} … {{SCHEDULE}}.
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
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(t.prop);
  if (!id) {
    const found = (Drive.Files.list({ q: "name='" + t.name.replace(/'/g, "\\'") + "' and trashed=false", fields: 'files(id,name,modifiedTime)', orderBy: 'modifiedTime desc', pageSize: 5 }).files || [])[0];
    if (!found) throw new Error('Contract template not found. Upload "' + t.name + '" to the Google Drive of the account that runs EDEN BI.');
    id = found.id; props.setProperty(t.prop, id);
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
    unit: at(['Unit Code']), project: at(['Project']), client: at(['Final Client Name', 'Client Name']),
    nationality: at(['client info|nationality', 'Nationality']), idNumber: at(['ID Number']),
    address: at(['Residence Address']) || at(['ID Address']), mobile: at(['client info|contact', 'Contact']), email: at(['Email']),
    floor: at(['unit info|floor', 'Floor']), area: num(['unit info|area', 'Area']), outArea: num(['OutArea', 'Out Area']),
    price: num(['price details|final price', 'Final Price']), dp: num(['Actual DP', 'Acctual DP']),
    contractDate: dateOf(['contract details|date'])
  };
  if (!d.unit || d.unit.toUpperCase().replace(/\.+$/, '') !== String(unitCode || '').trim().toUpperCase().replace(/\.+$/, '')) throw new Error('The Deals sheet changed (rows moved). Refresh and try again.');
  return d;
}

/* Installments for Annex 2: the unit's schedule if one was entered, else the EDEN WALK yearly plan. */
function contractSchedule_(d) {
  const tz = Session.getScriptTimeZone() || 'Africa/Cairo', out = [];
  try {
    const key = instKey_(d.project, d.unit);
    const rows = instRows_(instSheet_(INST_SHEET_, INST_HEAD_, false)).filter(r => String(r[0]) === key)
      .sort((a, b) => (a[7] instanceof Date ? a[7].getTime() : 0) - (b[7] instanceof Date ? b[7].getTime() : 0));
    rows.forEach(r => out.push([String(r[6] || ('القسط ' + r[5])), r[7] instanceof Date ? Utilities.formatDate(r[7], tz, 'dd/MM/yyyy') : String(r[7] || ''), Number(r[8]) || 0]));
  } catch (e) {}
  if (!out.length && typeof ewRead_ === 'function') {
    try {
      const u = ewRead_().units.find(x => ewKey_(x.unit) === ewKey_(d.unit));
      if (u) Object.keys(u.plan).sort().forEach(y => { const v = Number(u.plan[y][0]) || 0; if (v) out.push(['أقساط عام ' + y, 'خلال عام ' + y, v]); });
    } catch (e) {}
  }
  return out;
}

/* Builds the contract. req = {row, unitCode, date: 'yyyy-mm-dd' (optional)} */
function generateDealContract(authToken, req) {
  const user = validateAuthToken_(authToken);
  if (!contractCanMake_(user)) throw new Error('Only admins can generate contracts.');
  req = req || {};
  const d = contractDeal_(Number(req.row), req.unitCode);
  const file = contractTemplate_(d.project);
  const date = (req.date && instDateIn_(req.date)) || d.contractDate || new Date();
  const tz = Session.getScriptTimeZone() || 'Africa/Cairo';
  const rest = Math.max(0, d.price - d.dp);
  const sched = contractSchedule_(d);
  const rows = [['1', 'مقدم حجز وتعاقد (عند التوقيع)', Utilities.formatDate(date, tz, 'dd/MM/yyyy'), enNum_(d.dp)]]
    .concat(sched.map((s, i) => [String(i + 2), s[0], s[1], enNum_(s[2])]));
  const schedTotal = d.dp + sched.reduce((a, s) => a + s[2], 0);
  const vals = {
    DAY_NAME: arDay_(date), DD: Utilities.formatDate(date, tz, 'dd'), MM: Utilities.formatDate(date, tz, 'MM'), YYYY: Utilities.formatDate(date, tz, 'yyyy'),
    CLIENT_NAME: d.client, NATIONALITY: arNationality_(d.nationality), ID_NUMBER: d.idNumber, ADDRESS: d.address, EMAIL: d.email, MOBILE: d.mobile,
    UNIT_CODE: d.unit, AREA: enNum_(d.area), AREA_WORDS: arWords_(d.area), FLOOR: arFloor_(d.floor),
    OUT_AREA: d.outArea ? enNum_(d.outArea) + ' متراً مربعاً' : 'لا يوجد', OUT_AREA_WORDS: d.outArea ? arWords_(d.outArea) + ' متراً مربعاً' : '—',
    PRICE: enNum_(d.price), PRICE_WORDS: arMoney_(d.price), DP: enNum_(d.dp), DP_WORDS: arMoney_(d.dp), REST: enNum_(rest), REST_WORDS: arMoney_(rest)
  };
  const missing = ['CLIENT_NAME', 'ID_NUMBER', 'ADDRESS', 'MOBILE'].filter(k => !vals[k]);
  Object.keys(vals).forEach(k => { if (vals[k] === '' || vals[k] == null) vals[k] = '..........'; });

  const res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + file.id + '?alt=media&supportsAllDrives=true',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    PropertiesService.getScriptProperties().deleteProperty(CONTRACT_TEMPLATES_[String(d.project).toUpperCase().replace(/[^A-Z0-9]/g, '')].prop);
    throw new Error('Could not read the contract template from Drive (' + res.getResponseCode() + '). Try again.');
  }
  const parts = Utilities.unzip(res.getBlob().setContentType('application/zip'));
  const out = parts.map(p => {
    const name = p.getName();
    if (name !== 'word/document.xml') return p;
    let xml = p.getDataAsString('UTF-8');
    Object.keys(vals).forEach(k => { xml = xml.split('{{' + k + '}}').join(xmlEsc_(vals[k])); });
    // replace the whole paragraph that holds {{SCHEDULE}} with the table
    const i = xml.indexOf('{{SCHEDULE}}');
    if (i >= 0) {
      const s = xml.lastIndexOf('<w:p ', i) > xml.lastIndexOf('<w:p>', i) ? xml.lastIndexOf('<w:p ', i) : xml.lastIndexOf('<w:p>', i);
      const e = xml.indexOf('</w:p>', i) + 6;
      xml = xml.slice(0, s) + contractScheduleXml_(rows, schedTotal) + '<w:p/>' + xml.slice(e);
    }
    return Utilities.newBlob(xml, 'application/xml', name);
  });
  const fname = 'عقد ' + d.unit + ' - ' + d.client + ' - ' + Utilities.formatDate(date, tz, 'yyyy-MM-dd') + '.docx';
  const docx = Utilities.zip(out, fname).setContentType('application/vnd.openxmlformats-officedocument.wordprocessingml.document');

  let url = '';
  try {
    const folder = backupFolder_(CONTRACT_FOLDER_, null);
    const f = Drive.Files.create({ name: fname, parents: [folder], mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }, docx, { fields: 'id,webViewLink' });
    url = f.webViewLink || ('https://drive.google.com/file/d/' + f.id + '/view');
  } catch (e) { console.warn('Contract save: ' + e); }
  auditLog_(user, 'Contract generated', { project: d.project, unit: d.unit, client: d.client, details: 'Price ' + enNum_(d.price) + ' · DP ' + enNum_(d.dp) + (missing.length ? ' · missing: ' + missing.join(', ') : ''), ref: 'Deals row ' + req.row });
  return { name: fname, url: url, b64: Utilities.base64Encode(docx.getBytes()), missing: missing };
}
