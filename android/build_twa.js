/* Builds the EDEN BI Android app (a Trusted Web Activity: the real
   dashboard in Chrome, full screen, with its own icon) — run by
   .github/workflows/android.yml. Generates the Android project with
   Bubblewrap from the live web manifest; Gradle builds it; apksigner signs it. */
const path = require('path');
const { TwaManifest, TwaGenerator, ConsoleLog } = require('@bubblewrap/core');

(async () => {
  const site = process.env.SITE_URL || 'https://dashboard.edenalkhalij.com';
  const out = path.resolve(process.argv[2] || 'twa');
  const log = new ConsoleLog('twa');
  const m = await TwaManifest.fromWebManifest(site + '/manifest.webmanifest');
  m.packageId = 'com.edenalkhalij.dashboard';
  m.name = 'EDEN BI';
  m.launcherName = 'EDEN';
  m.startUrl = '/';
  m.display = 'standalone';
  m.orientation = 'default';
  m.enableNotifications = false;
  m.enableSiteSettingsShortcut = false;
  m.fallbackType = 'customtabs';
  m.appVersionCode = Number(process.env.VERSION_CODE || 1);
  m.appVersionName = process.env.VERSION_NAME || ('1.0.' + m.appVersionCode);
  m.signingKey = { path: path.resolve('android/eden-release.p12'), alias: 'eden' };
  m.generatorApp = 'eden-bi-ci';
  m.minSdkVersion = 24;   // Android 7+ (required by androidbrowserhelper)
  await new TwaGenerator().createTwaProject(out, m, log, () => {});
  console.log('Android project ready in ' + out);
})().catch(e => { console.error(e && e.stack || e); process.exit(1); });
