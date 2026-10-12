/* EDEN BI — google.script.run for the page hosted on GitHub Pages.
   Sends each server call to the Apps Script web app (Api.js) over GET and
   hands the result to the same success / failure handlers. */
(function () {
  'use strict';
  var API = window.EDEN_API_URL;
  var CAN_GZIP = typeof DecompressionStream === 'function';
  var NOT_UPDATED = 'The server has not been updated yet. In Apps Script: Deploy > Manage deployments > Edit > New version > Deploy.';
  var OFFLINE = 'Connection problem. Check your internet connection and try again.';

  /* Google sometimes answers with an HTML page instead of JSON (error,
     permission request, quota…). Show what it says instead of guessing. */
  function googleError(html, status) {
    var t = String(html || '');
    if (/Script function not found|doGet/i.test(t) && !/Exception/i.test(t)) return NOT_UPDATED;
    var title = (t.match(/<title[^>]*>([\s\S]*?)<\/title>/i) || [])[1] || '';
    var body = t.replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>/gi, ' ').replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&#39;/g, "'").replace(/&quot;/g, '"').replace(/\s+/g, ' ').trim();
    var msg = (body || title).slice(0, 220);
    if (/authoriz|permission|sign in/i.test(msg)) msg += ' — open Apps Script, run any function once and press Allow.';
    return 'Google server message (' + status + '): ' + (msg || 'empty response') + '';
  }

  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  function b64url(str) {
    var bytes = new TextEncoder().encode(str), bin = '';
    for (var i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }

  async function gunzip(b64) {
    var bin = atob(b64), u = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    var stream = new Blob([u]).stream().pipeThrough(new DecompressionStream('gzip'));
    return await new Response(stream).text();
  }

  async function get(params) {
    var lastMsg = OFFLINE;
    for (var attempt = 0; attempt < 3; attempt++) {
      var url = new URL(API);
      Object.keys(params).forEach(function (k) { url.searchParams.set(k, params[k]); });
      url.searchParams.set('_', Date.now().toString(36) + Math.random().toString(36).slice(2, 6));
      var res = null;
      try {
        var ctrl = typeof AbortController === 'function' ? new AbortController() : null;
        var timer = ctrl ? setTimeout(function () { ctrl.abort(); }, 90000) : null;
        try {
          res = await fetch(url.toString(), { credentials: 'omit', cache: 'no-store', redirect: 'follow', signal: ctrl ? ctrl.signal : undefined });
        } finally { if (timer) clearTimeout(timer); }
      } catch (e) {
        lastMsg = (e && e.name === 'AbortError') ? 'The Google server did not answer within 90 seconds. Open Apps Script > Executions to see what is running.' : OFFLINE;
        if (e && e.name === 'AbortError') throw new Error(lastMsg);
        if (attempt < 2) { await sleep(700 * (attempt + 1)); continue; }
        throw new Error(OFFLINE);
      }
      var text = await res.text(), json = null;
      try { json = JSON.parse(text); } catch (e) { json = null; }
      if (!json) {
        // Google sometimes answers once with an HTML error page; try again.
        lastMsg = googleError(text, res.status);
        if (attempt < 2) { await sleep(900 * (attempt + 1)); continue; }
        throw new Error(lastMsg);
      }
      if (json.ok === false) throw new Error(json.error || 'Server error.');
      return json;
    }
    throw new Error(lastMsg);
  }

  async function pool(items, size, fn) {
    var out = new Array(items.length), next = 0;
    async function worker() { while (next < items.length) { var i = next++; out[i] = await fn(items[i], i); } }
    var workers = [];
    for (var w = 0; w < Math.min(size, items.length); w++) workers.push(worker());
    await Promise.all(workers);
    return out;
  }

  /* Desktop app only (window.edenDesktop from its preload): keep a copy of
     every "get…" answer and use it when there is no internet. */
  var DESK = window.edenDesktop || null;
  async function call(fn, args) {
    if (!DESK) return callNet(fn, args);
    var tok = (typeof AUTH_TOKEN !== 'undefined' && AUTH_TOKEN) || '';
    var who = (typeof AUTH_USER !== 'undefined' && AUTH_USER && (AUTH_USER.username || AUTH_USER.name)) || '';
    var read = /^get/.test(fn), key = '';
    if (read) { try { key = who + '|' + fn + '|' + JSON.stringify((args || []).map(function (a) { return tok && a === tok ? '#token' : a; })); } catch (e) { read = false; } }
    try {
      var r = await callNet(fn, args);
      DESK.offline(false);
      if (read) { try { DESK.cacheSet(key, JSON.stringify(r)); } catch (e) {} }
      return r;
    } catch (err) {
      if (read && err && err.message === OFFLINE) {
        var c = null; try { c = await DESK.cacheGet(key); } catch (e) {}
        if (c) { DESK.offline(true); return JSON.parse(c); }
      }
      if (err && err.message === OFFLINE) DESK.offline(true);
      throw err;
    }
  }
  async function callNet(fn, args) {
    var q = b64url(JSON.stringify({ fn: fn, args: args })), z = CAN_GZIP ? '1' : '0', res;
    if (q.length <= 6000) {
      res = await get({ api: 'call', q: q, z: z });
    } else {
      var id = 'q' + Date.now().toString(36) + Math.random().toString(36).slice(2, 12), parts = [];
      for (var i = 0; i < q.length; i += 5000) parts.push(q.slice(i, i + 5000));
      await pool(parts, 4, function (d, n) { return get({ api: 'put', id: id, i: n, d: d }); });
      res = await get({ api: 'call', ref: id, n: parts.length, z: z });
    }
    var text = res.data;
    if (res.id) {
      var idx = []; for (var p = 0; p < res.parts; p++) idx.push(p);
      text = (await pool(idx, 6, function (n) { return get({ api: 'chunk', id: res.id, i: n }).then(function (r) { return r.data; }); })).join('');
    }
    if (res.gz) text = await gunzip(text);
    return JSON.parse(text);
  }

  function runner(onOk, onFail, userObj) {
    return new Proxy({}, {
      get: function (_, key) {
        if (key === 'withSuccessHandler') return function (h) { return runner(h, onFail, userObj); };
        if (key === 'withFailureHandler') return function (h) { return runner(onOk, h, userObj); };
        if (key === 'withUserObject') return function (o) { return runner(onOk, onFail, o); };
        if (typeof key !== 'string' || key === 'then') return undefined;
        return function () {
          var args = Array.prototype.slice.call(arguments);
          call(key, args).then(function (r) {
            if (onOk) { try { onOk(r, userObj); } catch (e) { console.error(e); } }
          }, function (err) {
            var e = err instanceof Error ? err : new Error(String(err));
            if (onFail) { try { onFail(e, userObj); } catch (x) { console.error(x); } } else console.error(e);
          });
        };
      }
    });
  }

  window.google = window.google || {};
  window.google.script = {
    get run() { return runner(null, null, undefined); },
    host: { close: function () {}, setHeight: function () {}, setWidth: function () {}, editor: { focus: function () {} } }
  };
})();

/* New version check: every few minutes (and when the app comes back to the
   front) ask for version.json; if the site was updated since this page
   loaded, show a small "Update" bar. Clicking it loads the new version. */
(function () {
  'use strict';
  var mine = window.EDEN_BUILD;
  if (!mine || !window.fetch) return;
  try {                                   // clean the ?v= left by the update click
    var u = new URL(location.href);
    if (u.searchParams.has('v')) { u.searchParams.delete('v'); history.replaceState(history.state, '', u.pathname + (u.search || '') + u.hash); }
  } catch (e) {}
  var shown = false, last = 0;
  function bar(v) {
    if (shown || !document.body) return; shown = true;
    var ar = document.documentElement.dir === 'rtl' || /^ar/i.test(document.documentElement.lang || '');
    var d = document.createElement('div');
    d.id = 'edenUpdateBar';
    d.setAttribute('style', 'position:fixed;left:50%;bottom:22px;transform:translateX(-50%);z-index:2147483000;display:flex;align-items:center;gap:12px;' +
      'padding:10px 12px 10px 18px;border-radius:14px;background:#16130f;color:#f3e7cc;box-shadow:0 12px 34px rgba(0,0,0,.35);' +
      'font:600 13.5px Inter,Segoe UI,Arial,sans-serif;max-width:calc(100vw - 32px)');
    d.innerHTML = '<span>' + (ar ? '✨ فيه تحديث جديد للسيستم' : '✨ A new version of EDEN BI is ready') + '</span>' +
      '<button type="button" style="border:0;border-radius:10px;padding:8px 14px;background:#c9a55c;color:#16130f;font:800 13px inherit;cursor:pointer">' +
      (ar ? 'تحديث' : 'Update') + '</button>' +
      '<button type="button" aria-label="Later" style="border:0;background:transparent;color:#8d7f66;font-size:18px;cursor:pointer;padding:0 4px">×</button>';
    var b = d.querySelectorAll('button');
    b[0].onclick = function () { location.replace(location.pathname + '?v=' + encodeURIComponent(v) + location.hash); };
    b[1].onclick = function () { d.remove(); setTimeout(function () { shown = false; }, 30 * 60 * 1000); };
    document.body.appendChild(d);
  }
  function check() {
    var now = Date.now();
    if (now - last < 60 * 1000) return; last = now;
    fetch('version.json?_=' + now.toString(36), { cache: 'no-store' })
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(function (j) { if (j && j.v && j.v !== mine) bar(j.v); })
      .catch(function () {});
  }
  setInterval(check, 5 * 60 * 1000);
  document.addEventListener('visibilitychange', function () { if (document.visibilityState === 'visible') check(); });
  window.addEventListener('focus', check);
  setTimeout(check, 20 * 1000);
})();

