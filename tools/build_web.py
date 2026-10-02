"""Build the GitHub Pages version of the dashboard.

Expands the Apps Script includes in Index.html into one static page and
adds web/shim.js, which sends google.script.run calls to the Apps Script
API (Api.js). Output: <out>/index.html  (default: _site/)

    python tools/build_web.py [out_dir]
"""
import json, pathlib, re, sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
OUT = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / '_site'
INCLUDE = re.compile(r"<\?!=\s*include\('([^']+)'\);?\s*\?>")


def include(name, seen=()):
    if name in seen:
        raise SystemExit(f'include loop: {name}')
    text = (ROOT / f'{name}.html').read_text(encoding='utf-8')
    return INCLUDE.sub(lambda m: include(m.group(1), seen + (name,)), text)


def main():
    cfg = json.loads((ROOT / 'web' / 'config.json').read_text(encoding='utf-8'))
    shim = (ROOT / 'web' / 'shim.js').read_text(encoding='utf-8')
    html = include('Index')
    if '<?' in html:
        raise SystemExit('Index.html still has Apps Script scriptlets the static build cannot run.')
    head = (
        '<meta name="theme-color" content="#07090f">'
        '<meta name="apple-mobile-web-app-capable" content="yes">'
        '<meta name="mobile-web-app-capable" content="yes">'
        f'<script>window.EDEN_API_URL={json.dumps(cfg["apiUrl"])};</script>'
        f'<script>{shim}</script>'
    )
    html = html.replace('<head>', '<head>' + head, 1)
    html = html.replace('<base target="_top">', '', 1)
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / 'index.html').write_text(html, encoding='utf-8')
    (OUT / '.nojekyll').write_text('', encoding='utf-8')
    cname = ROOT / 'web' / 'CNAME'
    if cname.exists():  # custom domain for GitHub Pages
        (OUT / 'CNAME').write_text(cname.read_text(encoding='utf-8').strip() + '\n', encoding='utf-8')
    print(f'wrote {OUT / "index.html"} ({len(html):,} chars)')


if __name__ == '__main__':
    main()
