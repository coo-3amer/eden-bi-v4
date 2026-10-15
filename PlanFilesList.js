/************************************************
 * EDEN BI V4 — PlanFilesList.js
 * Lists the payment-plan images the COO uploaded to Drive into the tab
 * "Plan Files" (id, name, type…), refreshed every 30 min from the warm-up,
 * so they can be read and imported into Installments. Read-only on Drive.
 ************************************************/
const PLAN_FOLDER_ID_ = '1FFpiMevxl0ztFk09UT9vAuo3MRHBpOUf';

function planFilesExport_() {
  try {
    const cache = CacheService.getScriptCache();
    if (cache.get('PLAN_FILES_DONE')) return;
    cache.put('PLAN_FILES_DONE', '1', 1800);
    const out = [];
    const walk = (folderId, path, depth) => {
      let tok;
      do {
        const r = Drive.Files.list({ q: "'" + folderId + "' in parents and trashed=false", fields: 'nextPageToken,files(id,name,mimeType,size,createdTime)',
          pageSize: 200, pageToken: tok, supportsAllDrives: true, includeItemsFromAllDrives: true });
        (r.files || []).forEach(f => {
          if (f.mimeType === 'application/vnd.google-apps.folder') { if (depth < 2) walk(f.id, path + f.name + '/', depth + 1); }
          else out.push([f.id, path + f.name, f.mimeType, Number(f.size) || '', f.createdTime || '']);
        });
        tok = r.nextPageToken;
      } while (tok);
    };
    walk(PLAN_FOLDER_ID_, '', 0);
    const sh = instSheet_('Plan Files', ['File ID', 'Name', 'Type', 'Size', 'Created'], true);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 5).clearContent();
    if (out.length) sh.getRange(2, 1, out.length, 5).setValues(out);
  } catch (e) { console.warn('planFilesExport_: ' + e); }
}
