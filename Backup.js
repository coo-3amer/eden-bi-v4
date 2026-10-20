/************************************************
 * EDEN BI V4 — Backup.js
 *
 * Daily copy of every spreadsheet the system uses (main file, Brokers,
 * inventories) into Google Drive:  "EDEN BI Backups / yyyy-MM-dd / …".
 * Keeps the last 30 days and trashes older days.
 * The daily trigger installs itself (from the 5-minute warm-up), so
 * nothing has to be run by hand. runDailyBackup can also be run any time.
 * Uses the Drive advanced service already enabled in appsscript.json.
 ************************************************/

const BACKUP_ROOT_NAME_ = 'EDEN BI Backups';
const BACKUP_KEEP_DAYS_ = 30;

function backupSources_() {
  const out = [{ id: SPREADSHEET_ID, name: 'EDEN BI - Main (Deals, Users, Targets, Installments, Audit Log)' }];
  if (typeof BROKER_SPREADSHEET_ID !== 'undefined') out.push({ id: BROKER_SPREADSHEET_ID, name: 'EDEN BI - Brokers' });
  Object.keys(INVENTORY_SOURCES).forEach(k => {
    const id = INVENTORY_SOURCES[k].spreadsheetId;
    if (id && !out.some(x => x.id === id)) out.push({ id: id, name: 'Inventory - ' + (INVENTORY_SOURCES[k].title || k) });
  });
  return out;
}

function backupFolder_(name, parentId) {
  const q = "mimeType='application/vnd.google-apps.folder' and trashed=false and name='" + name.replace(/'/g, "\\'") + "'" +
    (parentId ? " and '" + parentId + "' in parents" : " and 'root' in parents");
  const found = Drive.Files.list({ q: q, fields: 'files(id,name)', pageSize: 10 }).files || [];
  if (found.length) return found[0].id;
  const meta = { name: name, mimeType: 'application/vnd.google-apps.folder' };
  if (parentId) meta.parents = [parentId];
  const f = Drive.Files.create(meta, null, { fields: 'id' });
  return f.id;
}

function backupExists_(folder, name) {
  try {
    const q = "'" + folder + "' in parents and trashed=false and name='" + name.replace(/\\/g, '\\\\').replace(/'/g, "\\'") + "'";
    return (Drive.Files.list({ q: q, fields: 'files(id)', pageSize: 1 }).files || []).length > 0;
  } catch (e) { return false; }
}

function backupScheduleRetry_(minutes) {
  try {
    const has = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'runBackupRetry');
    if (!has) ScriptApp.newTrigger('runBackupRetry').timeBased().after(minutes * 60000).create();
  } catch (e) { console.warn('backupScheduleRetry_: ' + e); }
}

/* From the warm-up: today's backup has failures and no retry is waiting → schedule one. */
function backupCatchUp_() {
  try {
    const last = JSON.parse(PropertiesService.getScriptProperties().getProperty('BACKUP_LAST') || '{}') || {};
    const day = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Africa/Cairo', 'yyyy-MM-dd');
    if (last.day !== day || !(last.failed || []).length || (Number(last.tries) || 1) >= 4) return;
    if (Date.now() - new Date(last.at).getTime() < 25 * 60000) return;
    backupScheduleRetry_(1);
  } catch (e) { console.warn('backupCatchUp_: ' + e); }
}

/* One-off retry after a failed backup: copies only what is missing today. */
function runBackupRetry(e) {
  try {
    ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'runBackupRetry').forEach(t => ScriptApp.deleteTrigger(t));
  } catch (err) {}
  return runDailyBackup();
}

/* Copies every source into today's folder. Safe to run more than once a day. */
function runDailyBackup() {
  const tz = Session.getScriptTimeZone() || 'Africa/Cairo';
  const day = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const root = backupFolder_(BACKUP_ROOT_NAME_, null);
  const folder = backupFolder_(day, root);
  const existing = (Drive.Files.list({ q: "'" + folder + "' in parents and trashed=false", fields: 'files(name)', pageSize: 100 }).files || []).map(f => f.name);
  const done = [], failed = [];
  backupSources_().forEach(src => {
    const name = src.name + ' — ' + day;
    if (existing.indexOf(name) >= 0) { done.push(src.name + ' (already there)'); return; }
    // Google's Drive API sometimes answers "Internal Error" once: try up to 3 times.
    let err = null;
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        if (attempt > 1 && backupExists_(folder, name)) { err = null; break; }   // the failed call did copy it
        Drive.Files.copy({ name: name, parents: [folder] }, src.id, { fields: 'id', supportsAllDrives: true });
        err = null;
        break;
      } catch (e) {
        err = e;
        if (attempt < 3) Utilities.sleep(attempt * 8000);
      }
    }
    if (err) failed.push(src.name + ': ' + (err && err.message ? err.message : err));
    else done.push(src.name);
  });

  // Keep the last BACKUP_KEEP_DAYS_ days.
  const cutoff = Utilities.formatDate(new Date(Date.now() - BACKUP_KEEP_DAYS_ * 864e5), tz, 'yyyy-MM-dd');
  try {
    (Drive.Files.list({ q: "'" + root + "' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields: 'files(id,name)', pageSize: 200 }).files || [])
      .filter(f => /^\d{4}-\d{2}-\d{2}$/.test(f.name) && f.name < cutoff)
      .forEach(f => { try { Drive.Files.update({ trashed: true }, f.id); } catch (e) {} });
  } catch (e) { console.warn('Backup cleanup: ' + e); }

  const props = PropertiesService.getScriptProperties();
  let prev = {};
  try { prev = JSON.parse(props.getProperty('BACKUP_LAST') || '{}') || {}; } catch (e) {}
  const tries = prev.day === day ? (Number(prev.tries) || 1) + 1 : 1;
  props.setProperty('BACKUP_LAST', JSON.stringify({ day: day, at: new Date().toISOString(), done: done.length, failed: failed, tries: tries }));
  // Something failed: try again by itself in 30 minutes (up to 4 runs a day).
  if (failed.length && tries < 4) backupScheduleRetry_(30);
  const retryNote = failed.length ? (tries < 4 ? ' · will retry automatically in 30 min' : ' · gave up for today') : '';
  if (typeof auditLog_ === 'function' && (failed.length || prev.day !== day || (prev.failed || []).length))
    auditLog_('System', 'Daily backup', { details: done.length + ' files copied to Drive › ' + BACKUP_ROOT_NAME_ + ' › ' + day +
      (failed.length ? ' · Failed: ' + failed.join(' | ') + retryNote : (tries > 1 ? ' · retry succeeded, all files backed up' : '')) });
  console.log('Backup ' + day + ': ' + done.join(', ') + (failed.length ? ' | FAILED: ' + failed.join(' | ') : ''));
  return { day: day, done: done, failed: failed };
}

/* Called from the warm-up: makes sure the daily 3 AM trigger exists (checked once a day). */
function ensureDailyBackupTrigger_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Africa/Cairo', 'yyyy-MM-dd');
    if (props.getProperty('BACKUP_TRIGGER_CHECKED') === today) return;
    props.setProperty('BACKUP_TRIGGER_CHECKED', today);
    const has = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'runDailyBackup');
    if (!has) ScriptApp.newTrigger('runDailyBackup').timeBased().everyDays(1).atHour(3).create();
  } catch (e) {
    console.warn('ensureDailyBackupTrigger_: ' + e);
  }
}

/* Admins: when the last backup ran (shown in the Activity Log too). */
function getBackupStatus(authToken) {
  const user = validateAuthToken_(authToken);
  const role = String(user && user.role || '').trim().toLowerCase();
  if (role !== 'admin' && role !== 'super admin') throw new Error('Admins only.');
  const raw = PropertiesService.getScriptProperties().getProperty('BACKUP_LAST');
  return raw ? JSON.parse(raw) : null;
}

/* Admins: run the backup now from the Activity Log page. */
function runBackupNow(authToken) {
  const user = validateAuthToken_(authToken);
  const role = String(user && user.role || '').trim().toLowerCase();
  if (role !== 'admin' && role !== 'super admin') throw new Error('Admins only.');
  return runDailyBackup();
}
