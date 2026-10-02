/************************************************
 * EDEN BI — Daily email report
 *
 * Setup (once, from the Apps Script editor):
 *   1. Edit DAILY_REPORT.recipients below if needed.
 *   2. Run  installDailyReportTrigger  and approve the permissions.
 *   3. Optional: run  sendDailyReportNow  to receive a test email.
 * To stop it: run  removeDailyReportTrigger.
 ************************************************/

const DAILY_REPORT = {
  recipients: ['coo@edenalkhalij.com'], // add more addresses here
  hour: 9,                               // 9:00 AM, project time zone (Africa/Cairo)
  staleReservationDays: 30,
  dashboardUrl: ''                       // optional: paste the /exec link to add an "Open dashboard" button
};

function installDailyReportTrigger() {
  removeDailyReportTrigger();
  ScriptApp.newTrigger('sendDailyReport')
    .timeBased()
    .everyDays(1)
    .atHour(DAILY_REPORT.hour)
    .nearMinute(0)
    .create();
  return 'Daily report scheduled every day around ' + DAILY_REPORT.hour + ':00 (' + Session.getScriptTimeZone() + ').';
}

function removeDailyReportTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendDailyReport')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function sendDailyReportNow() {
  sendDailyReport();
}

function sendDailyReport() {
  const data = buildDashboardData_({ role: 'Super Admin', name: 'Daily Report' });
  const rows = (data.rows || []).filter(r => !cancelledStatus_(r.status));
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const today = dailyStartOfDay_(now);
  const yesterday = new Date(today.getTime() - 86400000);

  const dateOf = r => {
    const y = Number(r.year), m = Number(r.month);
    if (!y || !m) return null;
    return new Date(y, m - 1, Number(r.day) || 1);
  };
  const isUsd = r => String(r.currency || '').toUpperCase() === 'USD';
  const sumBy = (rs, f) => rs.reduce((a, r) => a + (Number(f(r)) || 0), 0);
  const money = n => Number(n || 0).toLocaleString('en-US', { maximumFractionDigits: 0 });

  const yRows = rows.filter(r => { const d = dateOf(r); return d && d.getTime() === yesterday.getTime(); });
  const mRows = rows.filter(r => Number(r.year) === yesterday.getFullYear() && Number(r.month) === yesterday.getMonth() + 1);

  const pm = new Date(yesterday.getFullYear(), yesterday.getMonth() - 1, 1);
  const pmSameDay = Math.min(yesterday.getDate(), new Date(pm.getFullYear(), pm.getMonth() + 1, 0).getDate());
  const pRows = rows.filter(r => Number(r.year) === pm.getFullYear() && Number(r.month) === pm.getMonth() + 1 && (Number(r.day) || 1) <= pmSameDay);

  const summary = rs => ({
    deals: rs.length,
    egp: sumBy(rs.filter(r => !isUsd(r)), r => r.amount),
    usd: sumBy(rs.filter(isUsd), r => r.amount),
    clients: new Set(rs.map(r => r.clientKey || r.clientName).filter(Boolean)).size
  });
  const Y = summary(yRows), M = summary(mRows), P = summary(pRows);
  const pct = (a, b) => (b ? Math.round((a - b) / b * 100) : null);

  const top = {};
  mRows.forEach(r => {
    const n = String(r.salesName || '').trim();
    if (!n || n === 'Not Assigned') return;
    top[n] = top[n] || { deals: 0, egp: 0 };
    top[n].deals++;
    if (!isUsd(r)) top[n].egp += Number(r.amount) || 0;
  });
  const leaders = Object.keys(top).map(n => ({ n, ...top[n] }))
    .sort((a, b) => b.deals - a.deals || b.egp - a.egp).slice(0, 5);

  const ageDays = r => { const d = dateOf(r); return d ? Math.floor((today - d) / 86400000) : 0; };
  const blank = v => { const s = String(v == null ? '' : v).trim().toLowerCase(); return !s || s === '-' || s === 'false' || s === 'n/a'; };
  const stale = rows.filter(r => String(r.status || '').toLowerCase().includes('reserv') && ageDays(r) > DAILY_REPORT.staleReservationDays)
    .sort((a, b) => ageDays(b) - ageDays(a));
  const awaiting = rows.filter(r => !blank(r.clientSignDate) && blank(r.companySignDate));
  const remainRows = rows.filter(r => Number(r.remain) > 0);
  const remainEgp = sumBy(remainRows.filter(r => !isUsd(r)), r => r.remain);
  const remainUsd = sumBy(remainRows.filter(isUsd), r => r.remain);

  const dayLabel = Utilities.formatDate(yesterday, tz, 'EEEE d MMMM yyyy');
  const monthLabel = Utilities.formatDate(yesterday, tz, 'MMMM yyyy');

  const esc = v => String(v == null ? '' : v).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const delta = (a, b) => {
    const p = pct(a, b);
    if (p === null) return '<span style="color:#8a7f70">—</span>';
    return `<span style="color:${p >= 0 ? '#2f9e62' : '#d64545'};font-weight:700">${p >= 0 ? '▲' : '▼'} ${Math.abs(p)}%</span>`;
  };
  const kpi = (label, value, sub) => `
    <td style="width:25%;padding:14px 12px;background:#fffdf9;border:1px solid #e8e0d1;border-radius:10px;vertical-align:top">
      <div style="font-size:11px;color:#83796b;font-weight:700;letter-spacing:.06em;text-transform:uppercase">${label}</div>
      <div style="font-size:22px;font-weight:800;color:#16130f;margin-top:4px">${value}</div>
      <div style="font-size:11px;color:#83796b;margin-top:2px">${sub || ''}</div>
    </td>`;
  const listRows = (rs, cols) => rs.slice(0, 8).map(r => `<tr>${cols.map(c => `<td style="padding:7px 8px;border-bottom:1px solid #eee5d6;font-size:12px;color:#2a2621">${c(r)}</td>`).join('')}</tr>`).join('');
  const th = t => `<th style="text-align:left;padding:7px 8px;font-size:10px;color:#83796b;letter-spacing:.06em;text-transform:uppercase;border-bottom:1px solid #e8e0d1">${t}</th>`;
  const section = (title, body) => `
    <tr><td style="padding:22px 28px 0">
      <div style="font-family:Georgia,serif;font-size:19px;font-weight:700;color:#16130f;margin-bottom:10px">${title}</div>${body}
    </td></tr>`;

  const yesterdayList = yRows.length
    ? `<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse">
        <tr>${th('Client')}${th('Project · Unit')}${th('Sales')}${th('Status')}${th('Value')}</tr>
        ${listRows(yRows, [r => esc(r.clientName), r => esc(r.project) + ' · ' + esc(r.unitCode), r => esc(r.salesName), r => esc(r.status), r => money(r.amount) + ' ' + esc(r.currency)])}
      </table>${yRows.length > 8 ? `<div style="font-size:12px;color:#83796b;margin-top:6px">+ ${yRows.length - 8} more</div>` : ''}`
    : '<div style="font-size:13px;color:#83796b">No new deals were recorded yesterday.</div>';

  const leadersList = leaders.length
    ? `<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse">
        <tr>${th('#')}${th('Salesperson')}${th('Deals')}${th('Value EGP')}</tr>
        ${leaders.map((l, i) => `<tr><td style="padding:7px 8px;border-bottom:1px solid #eee5d6;font-size:12px;color:#b08a42;font-weight:700">${i + 1}</td><td style="padding:7px 8px;border-bottom:1px solid #eee5d6;font-size:12px">${esc(l.n)}</td><td style="padding:7px 8px;border-bottom:1px solid #eee5d6;font-size:12px">${l.deals}</td><td style="padding:7px 8px;border-bottom:1px solid #eee5d6;font-size:12px">${money(l.egp)}</td></tr>`).join('')}
      </table>`
    : '<div style="font-size:13px;color:#83796b">No deals this month yet.</div>';

  const alertBox = (color, title, n, sub) => `
    <td style="width:33%;padding:12px;border:1px solid #e8e0d1;border-radius:10px;background:#fffdf9;vertical-align:top">
      <div style="font-size:20px;font-weight:800;color:${color};white-space:nowrap">${n}</div>
      <div style="font-size:12px;font-weight:700;color:#16130f">${title}</div>
      <div style="font-size:11px;color:#83796b">${sub}</div>
    </td>`;

  const staleList = stale.length
    ? `<table width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;margin-top:12px">
        <tr>${th('Client')}${th('Unit')}${th('Sales')}${th('Age')}</tr>
        ${listRows(stale, [r => esc(r.clientName), r => esc(r.project) + ' · ' + esc(r.unitCode), r => esc(r.salesName), r => ageDays(r) + ' days'])}
      </table>` : '';

  const button = DAILY_REPORT.dashboardUrl
    ? `<tr><td style="padding:24px 28px 0"><a href="${esc(DAILY_REPORT.dashboardUrl)}" style="display:inline-block;background:#b08a42;color:#fff;text-decoration:none;font-weight:700;font-size:13px;padding:11px 18px;border-radius:9px">Open EDEN BI dashboard</a></td></tr>` : '';

  const html = `
  <div style="background:#f4f0e8;padding:24px 0;font-family:Arial,Helvetica,sans-serif">
  <table width="640" align="center" cellspacing="0" cellpadding="0" style="background:#fbf8f2;border:1px solid #e8e0d1;border-radius:14px;overflow:hidden">
    <tr><td style="background:#16130f;padding:22px 28px">
      <div style="color:#d3ae66;font-size:11px;font-weight:700;letter-spacing:.3em">EDEN BUSINESS INTELLIGENCE</div>
      <div style="color:#fff;font-family:Georgia,serif;font-size:24px;margin-top:6px">Daily sales report</div>
      <div style="color:#bfb4a3;font-size:13px;margin-top:2px">${esc(dayLabel)}</div>
    </td></tr>

    ${section('Yesterday', `
      <table width="100%" cellspacing="8" cellpadding="0" style="margin:0 -8px"><tr>
        ${kpi('Deals', Y.deals, '')}${kpi('Value EGP', money(Y.egp), '')}${kpi('Value USD', money(Y.usd), '')}${kpi('Clients', Y.clients, '')}
      </tr></table>
      <div style="margin-top:12px">${yesterdayList}</div>`)}

    ${section(esc(monthLabel) + ' to date', `
      <table width="100%" cellspacing="8" cellpadding="0" style="margin:0 -8px"><tr>
        ${kpi('Deals', M.deals, delta(M.deals, P.deals) + ' vs same days last month')}
        ${kpi('Value EGP', money(M.egp), delta(M.egp, P.egp))}
        ${kpi('Value USD', money(M.usd), delta(M.usd, P.usd))}
        ${kpi('Clients', M.clients, delta(M.clients, P.clients))}
      </tr></table>`)}

    ${section('Top salespeople this month', leadersList)}

    ${section('Needs attention', `
      <table width="100%" cellspacing="8" cellpadding="0" style="margin:0 -8px"><tr>
        ${alertBox('#b08a42', 'Stale reservations', stale.length, 'Reserved over ' + DAILY_REPORT.staleReservationDays + ' days')}
        ${alertBox('#d64545', 'Awaiting company signature', awaiting.length, 'Client signed, company not')}
        ${alertBox('#3b6fd8', 'Remaining balances', money(remainEgp) + ' EGP', money(remainUsd) + ' USD · ' + remainRows.length + ' deals')}
      </tr></table>${staleList}`)}

    ${button}
    <tr><td style="padding:24px 28px;color:#a99f90;font-size:11px">Cancelled deals are excluded. Sent automatically by EDEN BI.</td></tr>
  </table></div>`;

  const subject = `EDEN BI · ${Utilities.formatDate(yesterday, tz, 'd MMM')}: ${Y.deals} deal${Y.deals === 1 ? '' : 's'} · ${monthLabel}: ${M.deals} deals`;
  MailApp.sendEmail({
    to: DAILY_REPORT.recipients.join(','),
    subject: subject,
    htmlBody: html,
    name: 'EDEN BI'
  });
}

function dailyStartOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}
