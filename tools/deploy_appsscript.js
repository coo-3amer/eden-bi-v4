/* Updates the live Apps Script web app after `clasp push`:
   1) creates a new version, 2) points the existing deployment at it
   (same /exec URL, so the dashboard keeps working).
   Credentials: the .clasprc.json made by `clasp login` (env CLASPRC_PATH). */
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const scriptId = JSON.parse(fs.readFileSync(path.join(ROOT, '.clasp.json'), 'utf8')).scriptId;
const apiUrl = JSON.parse(fs.readFileSync(path.join(ROOT, 'web', 'config.json'), 'utf8')).apiUrl;
const deploymentId = (process.env.DEPLOYMENT_ID || (apiUrl.match(/\/s\/([^/]+)\/exec/) || [])[1] || '').trim();
const rc = JSON.parse(fs.readFileSync(process.env.CLASPRC_PATH || path.join(require('os').homedir(), '.clasprc.json'), 'utf8'));

function creds() {
  if (rc.tokens) {                                   // clasp 3
    const t = rc.tokens.default || Object.values(rc.tokens)[0];
    return { client_id: t.client_id, client_secret: t.client_secret, refresh_token: t.refresh_token };
  }
  if (rc.token) {                                    // clasp 2
    const s = rc.oauth2ClientSettings || {};
    return { client_id: s.clientId, client_secret: s.clientSecret, refresh_token: rc.token.refresh_token };
  }
  throw new Error('Unknown .clasprc.json format');
}

async function main() {
  if (!deploymentId) throw new Error('Deployment ID not found');
  const c = creds();
  const tok = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ ...c, grant_type: 'refresh_token' })
  }).then(r => r.json());
  if (!tok.access_token) throw new Error('Google sign-in failed: ' + JSON.stringify(tok) + ' — run `npx clasp login` again and update the CLASPRC_JSON secret.');
  const H = { Authorization: 'Bearer ' + tok.access_token, 'Content-Type': 'application/json' };
  const base = 'https://script.googleapis.com/v1/projects/' + scriptId;
  const desc = 'Auto ' + (process.env.GITHUB_SHA || '').slice(0, 7) + ' ' + new Date().toISOString().slice(0, 16);

  const ver = await fetch(base + '/versions', { method: 'POST', headers: H, body: JSON.stringify({ description: desc }) }).then(r => r.json());
  if (!ver.versionNumber) throw new Error('Could not create version: ' + JSON.stringify(ver));

  const cur = await fetch(base + '/deployments/' + deploymentId, { headers: H }).then(r => r.json());
  if (!cur.deploymentConfig) throw new Error('Deployment not found: ' + JSON.stringify(cur));
  const cfg = { ...cur.deploymentConfig, versionNumber: ver.versionNumber, description: desc };
  const upd = await fetch(base + '/deployments/' + deploymentId, { method: 'PUT', headers: H, body: JSON.stringify({ deploymentConfig: cfg }) }).then(r => r.json());
  if (!upd.deploymentId) throw new Error('Could not update deployment: ' + JSON.stringify(upd));
  console.log('Deployed version ' + ver.versionNumber + ' to ' + deploymentId);
}
main().catch(e => { console.error(e.message || e); process.exit(1); });
