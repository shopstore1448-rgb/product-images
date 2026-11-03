"""Download product images listed in manifests/*.json into images/, as JPEG (max 1200px).
Manifest format: [["<source url>", "images/<SKU>/01.jpg"], ...]. Existing files are skipped."""
import json, glob, io, os, sys, time, urllib.request
from PIL import Image

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
failed = []
done = 0
for mf in sorted(glob.glob("manifests/*.json")):
    for url, path in json.load(open(mf, encoding="utf-8")):
        if os.path.exists(path):
            continue
        for attempt in range(3):
            try:
                req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "image/*"})
                data = urllib.request.urlopen(req, timeout=40).read()
                im = Image.open(io.BytesIO(data))
                if im.mode in ("RGBA", "LA", "P"):
                    im = im.convert("RGBA"); bg = Image.new("RGB", im.size, "white"); bg.paste(im, mask=im.split()[-1]); im = bg
                else:
                    im = im.convert("RGB")
                im.thumbnail((1200, 1200))
                os.makedirs(os.path.dirname(path), exist_ok=True)
                im.save(path, "JPEG", quality=85, optimize=True, progressive=True)
                done += 1
                break
            except Exception as e:
                if attempt == 2:
                    failed.append([url, path, str(e)])
                time.sleep(2)
print(f"downloaded {done}, failed {len(failed)}")
json.dump(failed, open("manifests/_failed.json", "w"), ensure_ascii=False, indent=1)
