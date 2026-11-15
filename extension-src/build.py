"""Builds the Chrome extension into ../extension/ from the shared Plausi sources."""
import pathlib, shutil, re
from PIL import Image, ImageDraw

root = pathlib.Path(__file__).resolve().parent.parent
src = root / "extension-src"
out = root / "extension"
if out.exists():
    keep = out / "lib" / "xlsx.full.min.js"
    saved = keep.read_bytes() if keep.exists() else None
    shutil.rmtree(out)
else:
    saved = None
(out / "icons").mkdir(parents=True)
(out / "lib").mkdir()
if saved:
    (out / "lib" / "xlsx.full.min.js").write_bytes(saved)

for f in ["manifest.json", "background.js", "bridge.js", "README.md"]:
    shutil.copy(src / f, out / f)
(out / "plausi-core.js").write_text("\n".join((root / f).read_text() for f in ["Engine.gs", "Demo.gs", "web/core.js"]))
shutil.copy(root / "web/app.js", out / "app.js")

page = (root / "web/template.html").read_text()
page = page.replace('<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>',
                    '<script src="lib/xlsx.full.min.js"></script>')
page = page.replace("<script>\n/*@LIBS@*/\n</script>", '<script src="plausi-core.js"></script>')
page = page.replace("<script>\n/*@APP@*/\n</script>", '<script src="app.js"></script>\n<script src="bridge.js"></script>')
assert "/*@" not in page and "<script>" not in page, "inline script left (not allowed in extensions)"
skeleton = ('<!doctype html><html lang="en"><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1">'
            '<style>body{margin:0}[hidden]{display:none!important}</style></head><body>')
(out / "checker.html").write_text(skeleton + page + "</body></html>")
(out / "lib" / "PUT-xlsx.full.min.js-HERE.txt").write_text(
    "Save https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js into this folder\n"
    "as xlsx.full.min.js, then reload the extension. Without it, CSV files and web tables still work.\n")

# Icon: ledger-blue tile with a check mark
for size in (16, 32, 48, 128):
    s = size * 4
    im = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle([0, 0, s - 1, s - 1], radius=s // 5, fill=(31, 95, 153, 255))
    w = max(3, s // 9)
    d.line([(s * 0.24, s * 0.53), (s * 0.43, s * 0.72), (s * 0.77, s * 0.30)], fill=(255, 255, 255, 255), width=w, joint="curve")
    im.resize((size, size), Image.LANCZOS).save(out / "icons" / f"icon{size}.png")
print("built", out, sorted(p.name for p in out.iterdir()))
