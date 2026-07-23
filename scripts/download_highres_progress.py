import json
import shutil
import ssl
import urllib.request
from pathlib import Path

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from PIL import Image, ImageOps


ROOT = Path(__file__).resolve().parents[1]
RESULTS_PATH = ROOT / "tmp_images" / "highres_url_results.json"
SRC_XLSX = ROOT / "secondary_process_duration.xlsx"
OUT_XLSX = ROOT / "secondary_process_duration_with_highres_images_progress.xlsx"
RAW_DIR = ROOT / "tmp_images" / "highres_downloaded"
EMBED_DIR = ROOT / "tmp_images" / "highres_embedded"
MANIFEST = ROOT / "tmp_images" / "highres_downloaded_progress.json"


def ext_from_url(url: str) -> str:
    clean = url.split("?", 1)[0].lower()
    for ext in [".jpg", ".jpeg", ".png", ".webp"]:
        if clean.endswith(ext):
            return ext
    return ".jpg"


def download(url: str, out: Path, ctx: ssl.SSLContext) -> None:
    if out.exists() and out.stat().st_size > 0:
        return
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=90, context=ctx) as resp:
        out.write_bytes(resp.read())


def make_embed(raw: Path, out: Path) -> tuple[int, int, int, int]:
    with Image.open(raw) as im:
        im = ImageOps.exif_transpose(im)
        orig_w, orig_h = im.size
        if im.mode != "RGB":
            im = im.convert("RGB")
        im.thumbnail((1000, 1000), Image.LANCZOS)
        emb_w, emb_h = im.size
        if not out.exists() or out.stat().st_size == 0:
            im.save(out, format="JPEG", quality=88, optimize=True)
    return orig_w, orig_h, emb_w, emb_h


def main() -> None:
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    EMBED_DIR.mkdir(parents=True, exist_ok=True)
    results = json.loads(RESULTS_PATH.read_text(encoding="utf-8"))
    ok_results = [r for r in results if r.get("ok") and r.get("url")]
    ctx = ssl.create_default_context()

    downloaded = []
    for r in ok_results:
        ext = ext_from_url(r["url"])
        raw = RAW_DIR / f"row_{r['row']}_{r['task']}{ext}"
        emb = EMBED_DIR / f"row_{r['row']}_{r['task']}.jpg"
        item = dict(r)
        try:
            download(r["url"], raw, ctx)
            ow, oh, ew, eh = make_embed(raw, emb)
            item.update(
                {
                    "download_ok": True,
                    "raw_file": str(raw),
                    "embed_file": str(emb),
                    "raw_bytes": raw.stat().st_size,
                    "embed_bytes": emb.stat().st_size,
                    "actual_width": ow,
                    "actual_height": oh,
                    "embed_width": ew,
                    "embed_height": eh,
                }
            )
        except Exception as exc:
            item.update({"download_ok": False, "download_error": str(exc)})
        downloaded.append(item)

    MANIFEST.write_text(json.dumps(downloaded, ensure_ascii=False, indent=2), encoding="utf-8")

    shutil.copyfile(SRC_XLSX, OUT_XLSX)
    wb = openpyxl.load_workbook(OUT_XLSX)
    ws = wb.worksheets[0]
    ws.column_dimensions["AO"].width = 20
    inserted = 0
    for item in downloaded:
        if not item.get("download_ok"):
            continue
        emb = Path(item["embed_file"])
        if not emb.exists():
            continue
        row = int(item["row"])
        with Image.open(emb) as im:
            ow, oh = im.size
        max_w, max_h = 130, 95
        scale = min(max_w / ow, max_h / oh)
        disp_w = max(1, int(ow * scale))
        disp_h = max(1, int(oh * scale))
        ws.row_dimensions[row].height = max(72, disp_h * 0.75)
        img = XLImage(str(emb))
        img.width = disp_w
        img.height = disp_h
        ws.add_image(img, f"AO{row}")
        inserted += 1

    wb.save(OUT_XLSX)
    summary = {
        "source_urls": len(ok_results),
        "downloaded_ok": sum(1 for x in downloaded if x.get("download_ok")),
        "download_failed": sum(1 for x in downloaded if not x.get("download_ok")),
        "inserted": inserted,
        "output": str(OUT_XLSX),
        "file_size": OUT_XLSX.stat().st_size,
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
