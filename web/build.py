"""Builds web/plausi-checker.html: one self-contained page from the template, the shared engine and the web code."""
import pathlib, re
root = pathlib.Path(__file__).resolve().parent.parent
libs = "\n".join((root / f).read_text() for f in ["Engine.gs", "Demo.gs", "web/core.js"])
app = (root / "web/app.js").read_text()
for name, src in [("libs", libs), ("app", app)]:
    assert "</script" not in src.lower(), name
html = (root / "web/template.html").read_text().replace("/*@LIBS@*/", libs).replace("/*@APP@*/", app)
out = root / "web/plausi-checker.html"
out.write_text(html)
print(f"wrote {out} ({len(html)//1024} KB)")
# Standalone copy for opening straight from disk: same page inside a full HTML skeleton.
skeleton = ('<!doctype html><html lang="en"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
            '<style>body{margin:0}[hidden]{display:none!important}</style></head><body>')
(root / "web/plausi-checker-standalone.html").write_text(skeleton + html + "</body></html>")
# GitHub Pages serves index.html at the repo root.
(root / "index.html").write_text(skeleton + html + "</body></html>")
