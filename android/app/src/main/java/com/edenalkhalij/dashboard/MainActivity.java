package com.edenalkhalij.dashboard;

import android.app.Activity;
import android.app.Dialog;
import android.content.ActivityNotFoundException;
import android.content.ContentValues;
import android.content.Context;
import android.content.Intent;
import android.graphics.Color;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.os.Message;
import android.print.PrintAttributes;
import android.print.PrintDocumentAdapter;
import android.print.PrintManager;
import android.provider.MediaStore;
import android.util.Base64;
import android.view.Gravity;
import android.view.View;
import android.view.ViewGroup;
import android.view.Window;
import android.webkit.CookieManager;
import android.webkit.JavascriptInterface;
import android.webkit.ValueCallback;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceRequest;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Button;
import android.widget.FrameLayout;
import android.widget.LinearLayout;
import android.widget.Toast;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

/**
 * EDEN BI — the dashboard in its own built-in browser (WebView), so it works
 * on every Android phone, with or without Google Chrome.
 *  - Stays inside the app for dashboard.edenalkhalij.com; other links
 *    (WhatsApp, phone, e-mail, Google Drive…) open in their own apps.
 *  - CSV / Excel exports are saved to Downloads.
 *  - Statements / reports (window.open + print) open in a full-screen page
 *    with Print / Save as PDF and Close buttons.
 */
public class MainActivity extends Activity {

    static final String HOME = "https://dashboard.edenalkhalij.com/";
    static final String HOST = "dashboard.edenalkhalij.com";
    static final int FILE_PICK = 41;

    WebView web;
    Dialog popup;
    WebView popupWeb;
    ValueCallback<Uri[]> fileCallback;

    /* Runs in every page: exports (blob links) → saved by the app; print → Android printing. */
    static final String BRIDGE_JS =
        "(function(){if(window.__edenApp)return;window.__edenApp=1;" +
        "function save(a){var h=a.href||'',n=a.download||'EDEN-export';" +
        " if(h.indexOf('blob:')===0||h.indexOf('data:')===0){fetch(h).then(function(r){return r.blob();}).then(function(b){" +
        "  var fr=new FileReader();fr.onload=function(){var s=String(fr.result);EDENAndroid.saveFile(n,b.type||'application/octet-stream',s.substring(s.indexOf(',')+1));};fr.readAsDataURL(b);});return true;}" +
        " return false;}" +
        "var oc=HTMLAnchorElement.prototype.click;HTMLAnchorElement.prototype.click=function(){if(this.hasAttribute('download')&&save(this))return;return oc.apply(this,arguments);};" +
        "document.addEventListener('click',function(e){var a=e.target&&e.target.closest&&e.target.closest('a[download]');if(a&&save(a)){e.preventDefault();e.stopPropagation();}},true);" +
        "window.print=function(){EDENAndroid.print();};" +
        "})();";

    @Override
    protected void onCreate(Bundle state) {
        super.onCreate(state);
        web = new WebView(this);
        web.setBackgroundColor(Color.parseColor("#07090F"));
        setContentView(web);
        setup(web, true);
        CookieManager.getInstance().setAcceptCookie(true);
        CookieManager.getInstance().setAcceptThirdPartyCookies(web, true);

        Uri data = getIntent() != null ? getIntent().getData() : null;
        if (state != null) web.restoreState(state);
        else web.loadUrl(data != null && HOST.equals(data.getHost()) ? data.toString() : HOME);
    }

    @Override
    protected void onSaveInstanceState(Bundle out) {
        super.onSaveInstanceState(out);
        web.saveState(out);
    }

    @Override
    protected void onNewIntent(Intent intent) {
        super.onNewIntent(intent);
        Uri data = intent.getData();
        if (data != null && HOST.equals(data.getHost())) web.loadUrl(data.toString());
    }

    void setup(final WebView v, final boolean main) {
        WebSettings s = v.getSettings();
        s.setJavaScriptEnabled(true);
        s.setDomStorageEnabled(true);
        s.setDatabaseEnabled(true);
        s.setJavaScriptCanOpenWindowsAutomatically(true);
        s.setSupportMultipleWindows(true);
        s.setLoadWithOverviewMode(true);
        s.setUseWideViewPort(true);
        s.setMediaPlaybackRequiresUserGesture(false);
        s.setMixedContentMode(WebSettings.MIXED_CONTENT_NEVER_ALLOW);
        s.setUserAgentString(s.getUserAgentString() + " EDENApp/2");
        v.addJavascriptInterface(new Bridge(v), "EDENAndroid");

        v.setWebViewClient(new WebViewClient() {
            @Override
            public boolean shouldOverrideUrlLoading(WebView view, WebResourceRequest req) {
                return openOutside(req.getUrl());
            }
            @Override
            public void onPageFinished(WebView view, String url) {
                view.evaluateJavascript(BRIDGE_JS, null);
            }
        });

        v.setWebChromeClient(new WebChromeClient() {
            @Override
            public boolean onCreateWindow(WebView view, boolean isDialog, boolean userGesture, Message resultMsg) {
                showPopup(resultMsg);
                return true;
            }
            @Override
            public void onCloseWindow(WebView window) {
                closePopup();
            }
            @Override
            public void onProgressChanged(WebView view, int p) {
                if (p > 30) view.evaluateJavascript(BRIDGE_JS, null);
            }
            @Override
            public boolean onShowFileChooser(WebView view, ValueCallback<Uri[]> cb, FileChooserParams params) {
                if (fileCallback != null) fileCallback.onReceiveValue(null);
                fileCallback = cb;
                try {
                    startActivityForResult(params.createIntent(), FILE_PICK);
                } catch (ActivityNotFoundException e) {
                    fileCallback = null;
                    return false;
                }
                return true;
            }
        });
    }

    /** true = handled outside the app. */
    boolean openOutside(Uri u) {
        String scheme = u.getScheme() == null ? "" : u.getScheme();
        String host = u.getHost() == null ? "" : u.getHost();
        if (("https".equals(scheme) || "http".equals(scheme))
                && (host.equals(HOST) || host.endsWith("script.google.com") || host.endsWith("googleusercontent.com"))) {
            return false; // stays in the app
        }
        try {
            Intent i = new Intent(Intent.ACTION_VIEW, u);
            i.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK);
            startActivity(i);
        } catch (Exception e) {
            Toast.makeText(this, "No app can open this link.", Toast.LENGTH_SHORT).show();
        }
        return true;
    }

    /* Reports / statements opened with window.open: a full-screen page with Print + Close. */
    void showPopup(Message resultMsg) {
        closePopup();
        popupWeb = new WebView(this);
        popupWeb.setBackgroundColor(Color.WHITE);
        setup(popupWeb, false);

        LinearLayout box = new LinearLayout(this);
        box.setOrientation(LinearLayout.VERTICAL);
        box.setBackgroundColor(Color.WHITE);
        LinearLayout bar = new LinearLayout(this);
        bar.setOrientation(LinearLayout.HORIZONTAL);
        bar.setGravity(Gravity.CENTER_VERTICAL);
        bar.setBackgroundColor(Color.parseColor("#07090F"));
        int pad = dp(8);
        bar.setPadding(pad, pad, pad, pad);
        Button close = new Button(this);
        close.setText("✕  Close");
        close.setOnClickListener(x -> closePopup());
        Button print = new Button(this);
        print.setText("🖨  Print / Save PDF");
        print.setOnClickListener(x -> printPage(popupWeb));
        View gap = new View(this);
        bar.addView(close);
        bar.addView(gap, new LinearLayout.LayoutParams(0, 1, 1f));
        bar.addView(print);
        box.addView(bar, new LinearLayout.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, ViewGroup.LayoutParams.WRAP_CONTENT));
        box.addView(popupWeb, new LinearLayout.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, 0, 1f));

        popup = new Dialog(this, android.R.style.Theme_Material_Light_NoActionBar_Fullscreen);
        popup.requestWindowFeature(Window.FEATURE_NO_TITLE);
        popup.setContentView(box);
        popup.setOnDismissListener(d -> {
            if (popupWeb != null) { popupWeb.destroy(); popupWeb = null; }
            popup = null;
        });
        popup.show();

        WebView.WebViewTransport t = (WebView.WebViewTransport) resultMsg.obj;
        t.setWebView(popupWeb);
        resultMsg.sendToTarget();
    }

    void closePopup() {
        if (popup != null) popup.dismiss();
    }

    void printPage(WebView v) {
        if (v == null) return;
        try {
            PrintManager pm = (PrintManager) getSystemService(Context.PRINT_SERVICE);
            PrintDocumentAdapter ad = v.createPrintDocumentAdapter("EDEN");
            pm.print("EDEN", ad, new PrintAttributes.Builder().setMediaSize(PrintAttributes.MediaSize.ISO_A4).build());
        } catch (Exception e) {
            Toast.makeText(this, "Printing is not available on this phone.", Toast.LENGTH_LONG).show();
        }
    }

    int dp(int v) { return Math.round(v * getResources().getDisplayMetrics().density); }

    class Bridge {
        final WebView owner;
        Bridge(WebView owner) { this.owner = owner; }

        @JavascriptInterface
        public void print() {
            runOnUiThread(() -> printPage(owner));
        }

        @JavascriptInterface
        public void saveFile(String name, String mime, String base64) {
            try {
                byte[] bytes = Base64.decode(base64, Base64.DEFAULT);
                String safe = name.replaceAll("[\\\\/:*?\"<>|]", "_");
                if (Build.VERSION.SDK_INT >= 29) {
                    ContentValues cv = new ContentValues();
                    cv.put(MediaStore.Downloads.DISPLAY_NAME, safe);
                    cv.put(MediaStore.Downloads.MIME_TYPE, mime);
                    Uri uri = getContentResolver().insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, cv);
                    try (OutputStream os = getContentResolver().openOutputStream(uri)) { os.write(bytes); }
                } else {
                    File dir = getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS);
                    try (FileOutputStream os = new FileOutputStream(new File(dir, safe))) { os.write(bytes); }
                }
                runOnUiThread(() -> Toast.makeText(MainActivity.this, "Saved to Downloads: " + safe, Toast.LENGTH_LONG).show());
            } catch (Exception e) {
                runOnUiThread(() -> Toast.makeText(MainActivity.this, "Could not save the file.", Toast.LENGTH_LONG).show());
            }
        }
    }

    @Override
    protected void onActivityResult(int req, int res, Intent data) {
        if (req == FILE_PICK && fileCallback != null) {
            fileCallback.onReceiveValue(WebChromeClient.FileChooserParams.parseResult(res, data));
            fileCallback = null;
            return;
        }
        super.onActivityResult(req, res, data);
    }

    @Override
    public void onBackPressed() {
        if (popup != null) { closePopup(); return; }
        // Let the dashboard close its own menu / pop-up first (Escape), then go back.
        web.evaluateJavascript(
            "(function(){var o=document.querySelector('#bc,#sc,#da,#pm,#al,.ip,.c360,.fm')||document.body.classList.contains('mNav');" +
            "if(o){document.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}));return 'closed';}return 'none';})()",
            r -> {
                if (r != null && r.contains("closed")) return;
                if (web.canGoBack()) web.goBack();
                else moveTaskToBack(true);
            });
    }
}
