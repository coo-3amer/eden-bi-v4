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

  async function call(fn, args) {
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
