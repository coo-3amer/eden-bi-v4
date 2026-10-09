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
    try {
      Drive.Files.copy({ name: name, parents: [folder] }, src.id, { fields: 'id', supportsAllDrives: true });
      done.push(src.name);
    } catch (e) {
      failed.push(src.name + ': ' + (e && e.message ? e.message : e));
    }
  });

  // Keep the last BACKUP_KEEP_DAYS_ days.
  const cutoff = Utilities.formatDate(new Date(Date.now() - BACKUP_KEEP_DAYS_ * 864e5), tz, 'yyyy-MM-dd');
  try {
    (Drive.Files.list({ q: "'" + root + "' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false", fields: 'files(id,name)', pageSize: 200 }).files || [])
      .filter(f => /^\d{4}-\d{2}-\d{2}$/.test(f.name) && f.name < cutoff)
      .forEach(f => { try { Drive.Files.update({ trashed: true }, f.id); } catch (e) {} });
  } catch (e) { console.warn('Backup cleanup: ' + e); }

  PropertiesService.getScriptProperties().setProperty('BACKUP_LAST', JSON.stringify({ day: day, at: new Date().toISOString(), done: done.length, failed: failed }));
  if (typeof auditLog_ === 'function') auditLog_('System', 'Daily backup', { details: done.length + ' files copied to Drive › ' + BACKUP_ROOT_NAME_ + ' › ' + day + (failed.length ? ' · Failed: ' + failed.join(' | ') : '') });
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
