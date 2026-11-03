# -*- coding: utf-8 -*-
"""Build a Salla import file for one brand from Boutiqaat (+ optional official Salla store).

Usage: python tools/pipeline/build_brand.py jobs/<job>.json
Job file: {"brand": "الماجد للعود", "slug": "almajed",
           "boutiqaat_url": "https://www.boutiqaat.com/ar-sa/women/al-majed-oud-1/br/",
           "official_url": "https://sa.almajed4oud.com/ar"}   # official_url optional
Pricing: cost = source price; before discount = cost*1.5+100; after discount = cost*1.25+30 (rounded).
Outputs: output/<slug>/import.xlsx, output/<slug>/review.xlsx, manifests/<slug>.json (images fetched by fetch_images.py)
"""
import json, re, sys, os, math, html, time, urllib.request, urllib.parse, concurrent.futures as cf
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
GH = "https://raw.githubusercontent.com/shopstore1448-rgb/product-images/main/"
T = json.load(open(os.path.join(HERE, "templates.json"), encoding="utf-8"))
LOG = []

def log(*a):
    s = " ".join(str(x) for x in a); print(s, flush=True); LOG.append(s)

def get(url, headers=None, tries=3):
    for i in range(tries):
        try:
            h = {"User-Agent": UA, "Accept-Language": "ar"}; h.update(headers or {})
            return urllib.request.urlopen(urllib.request.Request(url, headers=h), timeout=40).read().decode("utf-8", "replace")
        except Exception as e:
            if i == tries - 1: raise
            time.sleep(2 + i * 3)

def rnd(x): return int(math.floor(x + 0.5))
def strip_html(s): return html.unescape(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", s or ""))).strip()

# ------------------------------------------------------------------ Boutiqaat
def next_data(h):
    m = re.search(r'<script id="__NEXT_DATA__" type="application/json">(.*?)</script>', h, re.S)
    return json.loads(m.group(1)) if m else None

def find_obj(o, eid):
    if isinstance(o, dict):
        if str(o.get("entity_id")) == str(eid) and "description" in o: return o
        for v in o.values():
            r = find_obj(v, eid)
            if r: return r
    elif isinstance(o, list):
        for v in o:
            r = find_obj(v, eid)
            if r: return r
    return None

def boutiqaat(url):
    url = url.split("?")[0]
    prods = {}
    for pg in range(1, 40):
        d = next_data(get(f"{url}?page={pg}"))
        r = d["props"]["pageProps"]["response"]
        ps = (r.get("products") or [[]])[0] or []
        if not ps: break
        new = 0
        for p in ps:
            if p["entity_id"] not in prods: new += 1
            prods[p["entity_id"]] = dict(id=p["entity_id"], sku=p["sku"], name=p["name"].strip(), price=float(str(p["final_price"]).replace(",", "")),
                                          avail=str(p.get("is_saleable")) == "1", url=p.get("product_url") or ("https://www.boutiqaat.com/ar-sa/women/" + p["slug"]))
        total = (r.get("metadata") or {}).get("total_pages") or 99
        if pg >= total or not new: break
    log("boutiqaat products:", len(prods), "available:", sum(p["avail"] for p in prods.values()))

    def detail(p):
        try:
            o = find_obj(next_data(get(p["url"])), p["id"]) or {}
            g = o.get("gallery_images") or []
            if isinstance(g, str):
                try: g = json.loads(g)
                except Exception: g = []
            imgs = [x["image"] for x in g if isinstance(x, dict) and not x.get("isVideo")]
            main = o.get("image_url")
            if main:
                imgs = [main] + [x for x in imgs if x != main]
            p.update(desc=strip_html(o.get("description")), imgs=imgs, cat=o.get("category_name"))
            if o.get("final_price_with_tax"): p["price"] = float(str(o["final_price_with_tax"]).replace(",", ""))
            if o.get("is_saleable") is not None: p["avail"] = str(o["is_saleable"]) == "1"
        except Exception as e:
            p.update(desc="", imgs=[], err=str(e))
        return p
    with cf.ThreadPoolExecutor(6) as ex: list(ex.map(detail, prods.values()))
    log("boutiqaat details errors:", sum(1 for p in prods.values() if p.get("err")))
    return prods

# ------------------------------------------------------------------ Salla official store
def salla(url):
    h = get(url)
    m = re.search(r'"store"\s*:\s*\{[^{}]*?"id"\s*:\s*(\d+)', h) or re.search(r'store_id["\']?\s*[:=]\s*["\']?(\d+)', h) or re.search(r'"id"\s*:\s*(\d{6,})\s*,\s*"is_merchant"', h)
    if not m or "salla" not in h:
        log("official site: not a Salla store or store id not found -> skipped"); return {}
    sid = m.group(1); hdr = {"Store-Identifier": sid, "Accept": "application/json", "Currency": "SAR"}
    out, nxt = {}, "https://api.salla.dev/store/v1/products?per_page=100"
    while nxt:
        j = json.loads(get(nxt, hdr))
        for p in j.get("data", []):
            out[str(p["id"])] = dict(id=str(p["id"]), name=p["name"].strip(), price=float(p["price"]), avail=bool(p.get("is_available")), url=p["url"], cur=p.get("currency"))
        nxt = (j.get("cursor") or {}).get("next")
    log("official products:", len(out), "available:", sum(p["avail"] for p in out.values()), "store", sid)

    def detail(p):
        try:
            j = json.loads(get(f"https://api.salla.dev/store/v1/products/{p['id']}/details", hdr))
            p["desc"] = strip_html(j["data"].get("description"))
            ph = get(p["url"])
            imgs = re.findall(r'(?:data-src|src)="(https://cdn\.salla\.sa/[^"]+?\.(?:jpg|jpeg|png|webp))"', ph)
            seen, res = set(), []
            for u in imgs:
                if "cdn-cgi" in u or "form-builder" in u or u in seen: continue
                seen.add(u); res.append(u)
            p["imgs"] = res
        except Exception as e:
            p.update(desc="", imgs=[], err=str(e))
        return p
    with cf.ThreadPoolExecutor(6) as ex: list(ex.map(detail, out.values()))
    # drop store-wide images (logos/banners) that appear on many products
    from collections import Counter
    c = Counter(u for p in out.values() for u in p.get("imgs", []))
    common = {u for u, n in c.items() if n > 2}
    for p in out.values(): p["imgs"] = [u for u in p.get("imgs", []) if u not in common][:8]
    return out

# ------------------------------------------------------------------ names, types, categories
SYN = [("أ", "ا"), ("إ", "ا"), ("آ", "ا"), ("ى", "ي"), ("ة", "ه"), ("جراي", "غراي"), ("بريستيج", "برستيج"), ("سيكرت", "سكريت"), ("سيكريت", "سكريت"),
       ("جلو", "غلو"), ("كوير", "كواير"), ("بوازيه", "بويزيه"), ("بوازية", "بويزيه"), ("ربع توله", "3مل"), ("نصف توله", "6مل"),
       ("الإرث", "الارث"), ("ليدى", "ليدي")]
STOP = set("دخون بخور انسنس دهن خمريه الجسم للجسم الملابس شاور جل غسول واليدين لوشن بودي مرطب زيت موزع عطري جهاز تعطير فواحه اعواد مع معطر الجو جو رذاذ والجسم الشعر للشعر بخاخ سبراي اول اوفر افر للسياره سيارات سياره عبوه اعاده تعبئه هديه هدايا بوكس عطر المركز من مجموعه او دو بارفيوم تواليت للعود مل غ - حبه حبات قطع و كلاسيك ليدي برستيج الارث ليغاسي تشكيله عالي التركيز المخفف مستخلص اكستريت للجسم والملابس".split())

def norm(s):
    for a, b in SYN: s = s.replace(a, b)
    s = re.sub(r"(\d+)\s*(جرام|جم|غم|غرام|غ)(?![\u0600-\u06FF])", r" \1غ ", s)
    return re.sub(r"(\d+)\s*(مل)", r" \1\2 ", s)
def size(s):
    m = re.findall(r"(\d+)(مل|غ)", norm(s)); return m[-1] if m else None
def toks(s, brand):
    s = norm(s.replace(brand, " ")); s = re.sub(r"\d+(مل|غ)", "", s); s = re.sub(r"[()\-–,*×:]", " ", s)
    return set(w for w in s.split() if w not in STOP and not w.isdigit() and len(w) > 1)

def kind(n):
    if "زعفران" in n: return "saffron"
    if "توزيعات" in n: return "eid"
    if "زيت موزع" in n or "زيت جهاز" in n or "زيت فواحة" in n: return "diffoil"
    if "سيار" in n: return "car"
    if any(w in n for w in ["بوكس", "صندوق"]): return "set"
    if "اكسير" in n or "إكسير" in n: return "poil"
    if any(w in n for w in ["بوكس", "صندوق", "مجموعة عطر", "مجموعة هدايا", "طقم", "× ", " x "]) or re.match(r"^مجموعة\s", n): return "set"
    if "سيار" in n: return "car"
    if "جهاز" in n: return "device"
    if re.search(r"زيت عطري .* - 500", n): return "diffoil"
    if any(w in n for w in ["فواحة", "موزع روائح", "موزع عطري", "أعواد عطرية"]): return "reed"
    if "شمع" in n: return "candle"
    if any(w in n for w in ["معطر جو", "معطر الجو", "معطر المنزل", "رشوش", "عطر جوك", "معطر مفارش"]): return "home"
    if any(w in n for w in ["غسول", "شاور", "صابون"]): return "wash"
    if any(w in n for w in ["لوشن", "مرطب", "كريم الجسم", "زبدة"]): return "lotion"
    if "كريم" in n: return "cream"
    if "مبخر" in n: return "burner"
    if "بودرة" in n or "بودر" in n.split()[0:1]: return "powder"
    if any(w in n for w in ["رذاذ", "للشعر والجسم", "الشعر و الجسم", "للجسم -", "بخاخ الجسم"]): return "mist"
    if any(w in n for w in ["معطر الشعر", "عطر الشعر", "معطر للشعر"]): return "hair"
    if any(w in n for w in ["دهن", "خمرية"]): return "dehn"
    if "زيت" in n and "جسم" in n: return "bodyoil"
    if "زيت" in n: return "poil"
    if "مسك" in n and re.search(r"\d+\s*غ", n) and "عطر" not in n: return "musk"
    if any(w in n for w in ["دخون", "بخور", "مبسوس", "معمول", "نفحة", "كبسولات", "لبان", "مخمرية"]): return "incense"
    if n.startswith("عود") or n.startswith("دهن عود"): return "oud"
    return "edp"

def category(n, t):
    if t == "edp":
        if any(w in n for w in ["ليدي", "ليدى", "نسائي", "للنساء"]): return ("العطور", "عطور نسائية")
        if any(w in n for w in ["رجالي", "للرجال"]): return ("العطور", "عطور رجالية")
        return ("العطور", "عطور للجنسين")
    if t == "set":
        if "عطر" not in n and any(w in n for w in ["عود", "دخون", "بخور", "معمول"]): return ("البخور والعود", "مجموعات البخور والعود")
        return ("العطور", "المجموعات العطرية")
    return {"hair": ("العطور", "معطرات الجسم والشعر"), "mist": ("العطور", "معطرات الجسم والشعر"), "poil": ("العطور", "معطرات الجسم والشعر"),
            "bodyoil": ("العطور", "معطرات الجسم والشعر"), "powder": ("العطور", "معطرات الجسم والشعر"), "musk": ("العطور", "معطرات الجسم والشعر"),
            "dehn": ("البخور والعود", "دهن العود"), "incense": ("البخور والعود", "البخور والدخون"), "oud": ("البخور والعود", "العود"),
            "burner": ("البخور والعود", "المباخر"), "home": ("معطرات المنزل", "بخاخات المنزل"), "diffoil": ("معطرات المنزل", "فواحات وأعواد عطرية"),
            "reed": ("معطرات المنزل", "فواحات وأعواد عطرية"), "device": ("معطرات المنزل", "فواحات وأعواد عطرية"), "car": ("معطرات المنزل", "فواحات وأعواد عطرية"),
            "candle": ("معطرات المنزل", "شموع معطرة"), "wash": ("العناية بالجسم والاستحمام", "الاستحمام والصابون"), "lotion": ("العناية بالجسم والاستحمام", "مرطبات الجسم"),
            "cream": ("العناية بالجسم والاستحمام", "العناية باليدين والقدمين"), "saffron": ("المنزل والمطبخ", "الضيافة والتقديم"), "eid": ("الهدايا والمجموعات", "الهدايا")}[t]

def clean_name(n, brand):
    n = n.replace(" من " + brand, "").replace(brand + " - ", "").replace(brand, "").strip(" -")
    n = n.replace("عالي التركيز", "المركز").replace("سيارات", "سيارات ").replace("مبخره", "مبخرة").replace("من تشكيلة", "من مجموعة").replace("الارث", "الإرث")
    n = re.sub(r"\s*(او|أو) دو (بارفيوم|تواليت)|\s*اكستريت دي? بارفيوم|\s*اكستريت دو بارفيوم", "", n)
    n = n.replace("ليدى", "ليدي").replace("عطرى", "عطري").replace("بودى", "بودي")
    m = re.search(r"\s(\d+)\s?(مل|جم|جرام|غم|غ)$", n)
    if m and " - " not in n:
        n = n[:m.start()] + f" - {m.group(1)} {'مل' if m.group(2) == 'مل' else 'غ'}"
    n = re.sub(r"(\d+)\s?(مل|غ)\b", r"\1 \2", n).replace(" جم", " غ").replace(" جرام", " غ").replace(" غم", " غ")
    return re.sub(r"\s+", " ", n).strip(" -")

# ------------------------------------------------------------------ content
def parse_notes(d):
    out = []
    pats = [("مقدمة العطر", r"(?:مقدمة العطر|القمة|الافتتاحية)\s*:?\s*"), ("قلب العطر", r"(?:قلب العطر|القلب)\s*:?\s*"), ("قاعدة العطر", r"(?:قاعدة العطر|قعدة العطر|القاعدة)\s*:?\s*")]
    for label, p in pats:
        m = re.search(p + r"(.+?)(?=\s*-?\s*(?:مقدمة العطر|قلب العطر|قاعدة العطر|قعدة العطر|القمة|القلب|القاعدة|العائلة|طريقة|التصنيف|الحجم)|$)", d)
        if m:
            v = m.group(1).strip(" -:،.")
            if 2 < len(v) < 160: out.append((label, v))
    m = re.search(r"العائلة العطرية\s*:?\s*([^-\n]+?)(?=\s*(?:التصنيف|الروائح|مقدمة|القمة|الحجم|-)|$)", d)
    if m and len(m.group(1)) < 60: out.insert(0, ("العائلة العطرية", m.group(1).strip(" :،.")))
    return out

def first_sentences(d, maxlen=260):
    d = re.split(r"(?:الروائح الأساسية|الروائح الاساسية|العائلة العطرية|مقدمة العطر|طريقة الاستخدام|القمة|التصنيف|المكونات)", d)[0].strip(" -:")
    if len(d) < 25: return ""
    out = ""
    for x in re.split(r"(?<=[.!؟])\s+", d):
        if len(out) + len(x) > maxlen: break
        out += x + " "
    return out.strip() or d[:maxlen].rsplit(" ", 1)[0] + "."

def shortname(n):
    s = re.sub(r"\s*-\s*[\d.]+\s*(مل|غ|حبات|حبة|حبتين|قطع).*$", "", n)
    return re.sub(r"\s*-\s*(حبتين|\d+ حبات?)$", "", s).strip()

def build_desc(name, t, srcdesc, i, b):
    tk = t if t in T["WHY"] else "edp"
    h1s = T["H1"].get(tk) or T["H1"]["edp"]
    h = f"<h1>{h1s[i % len(h1s)].format(s=shortname(name))}</h1><h2>عن المنتج</h2>"
    intro = (T["INTRO"].get(tk) or T["INTRO"]["edp"]).format(n=name, b=b)
    extra = first_sentences(srcdesc or "")
    h += f"<p>{intro}{(' ' + extra) if extra else ''}</p><h3>المميزات والتفاصيل</h3>"
    notes = parse_notes(srcdesc or "")
    if notes: h += "<ul>" + "".join(f"<li><strong>{a}:</strong> {v}</li>" for a, v in notes) + "</ul>"
    W, U = T["WHY"][tk], T["USAGE"].get(tk, T["USAGE"]["edp"])
    h += f"<p>{W[i % 2].format(b=b)}</p><p>{U[i % 2].format(b=b)}</p><p>{W[(i + 1) % 2].format(b=b)}</p>"
    h += f"<h3>اختيارك اليومي</h3><p>{T['CLOSING'][i % 2].format(b=b)}</p>"
    return h, intro + ((" " + extra) if extra else "")

def slug(s):
    s = re.sub(r"[^\w\s؀-ۿ-]", " ", s).replace("-", " ")
    return re.sub(r"\s+", "-", s.strip())

# ------------------------------------------------------------------ main
def main(job_path, outdir="."):
    job = json.load(open(job_path, encoding="utf-8"))
    os.chdir(outdir)
    b, sl = job["brand"], job["slug"]
    # Prefer data collected in the browser (Boutiqaat blocks servers); fall back to direct fetching.
    if job.get("bq_json"):
        bq = {str(p["id"]): p for p in json.load(open(job["bq_json"], encoding="utf-8"))}
        log("boutiqaat (browser data):", len(bq), "available:", sum(p["avail"] for p in bq.values()))
    else:
        bq = boutiqaat(job["boutiqaat_url"]) if job.get("boutiqaat_url") else {}
    if job.get("off_json"):
        off = {str(p["id"]): p for p in json.load(open(job["off_json"], encoding="utf-8"))}
        from collections import Counter
        c = Counter(u for p in off.values() for u in p.get("imgs", []))
        for p in off.values(): p["imgs"] = [u for u in p.get("imgs", []) if c[u] <= 2 and "form-builder" not in u][:8]
        log("official (browser data):", len(off), "available:", sum(p["avail"] for p in off.values()))
    else:
        off = salla(job["official_url"]) if job.get("official_url") else {}

    # match each official product to a Boutiqaat product (same type + same size + name tokens)
    B = {k: (kind(clean_name(p["name"], b)), size(p["name"]), toks(p["name"], b)) for k, p in bq.items()}
    matches = {}
    for oid, o in off.items():
        t, sz, tk = kind(clean_name(o["name"], b)), size(o["name"]), toks(o["name"], b)
        best = (0, None)
        for k, (bt, bs, btk) in B.items():
            if bt != t or (bs != sz and t != "set") or not tk or not btk: continue
            j = len(tk & btk) / len(tk | btk)
            QUAL = {"الممتاز", "الملكي", "ملكي", "الخاص", "خاص", "الفاخر", "فاخر", "فاخره", "نادر", "عتيق", "سوبر", "ممتاز", "مبسوس", "معمول"}
            if t != "set" and (tk <= btk or btk <= tk) and not ((tk ^ btk) & QUAL): j = max(j, 0.6)
            if t == "set" and j < 0.75: continue
            if (tk ^ btk) & QUAL and j < 0.75: continue
            if j > best[0]: best = (j, k)
        if best[1] and (best[0] >= 0.5 or (best[0] >= 0.34 and len(tk & B[best[1]][2]) >= 1 and len(tk) <= 2)):
            matches[oid] = best
    matched_bq = {}
    for oid, (sc, k) in matches.items():
        if k not in matched_bq or sc > matched_bq[k][0]: matched_bq[k] = (sc, oid)

    items = []
    for k, p in bq.items():
        if p["avail"]:
            items.append(dict(src="بوتيكات", sku=p["sku"], name=clean_name(p["name"], b), cost=p["price"], desc=p.get("desc", ""), srcimgs=p.get("imgs", []), match=""))
        elif k in matched_bq and off[matched_bq[k][1]]["avail"]:
            o = off[matched_bq[k][1]]
            items.append(dict(src="الموقع الرسمي (غير متوفر في بوتيكات)", sku=p["sku"], name=clean_name(p["name"], b), cost=o["price"], desc=p.get("desc") or o.get("desc", ""),
                              srcimgs=p.get("imgs") or o.get("imgs", []), match=f"{o['name']} ({matched_bq[k][0]:.2f})"))
    avail_bq = {k: B[k] for k, p in bq.items() if p["avail"]}
    for oid, o in off.items():
        if not o["avail"] or oid in matches: continue
        t, tk = kind(clean_name(o["name"], b)), toks(o["name"], b)
        near = max(((len(tk & v[2]) / max(1, len(tk | v[2])), k) for k, v in avail_bq.items() if v[0] == t), default=(0, None))
        note = f"قد يكون مكررًا مع: {bq[near[1]]['name']}" if near[0] >= 0.5 else ""
        items.append(dict(src="الموقع الرسمي فقط", sku=f"OFF-{sl}-{oid}"[:40], name=clean_name(o["name"], b), cost=o["price"], desc=o.get("desc", ""), srcimgs=o.get("imgs", []), match="", note=note))
    items = [it for it in items if it["srcimgs"]] + [dict(it, skipped=True) for it in items if not it["srcimgs"]]

    wb = openpyxl.load_workbook(os.path.join(HERE, "salla_template.xlsx")); ws = wb.active
    rv = openpyxl.Workbook(); w2 = rv.active; w2.title = "المراجعة"; w2.sheet_view.rightToLeft = True
    w2.append(["SKU", "الاسم", "المصدر", "التكلفة", "قبل الخصم", "بعد الخصم", "التصنيف", "عدد الصور", "مطابق مع (الموقع الرسمي)", "ملاحظة"])
    manifest, slugs, names, row, words = [], set(), {}, 2, []
    for i, it in enumerate(items):
        if it.get("skipped"):
            w2.append([it["sku"], it["name"], it["src"], "", "", "", "", 0, it["match"], "بدون صور - لم يُضف"]); continue
        nm = it["name"]
        if nm in names:  # same name twice -> make distinct
            nm = f"{nm} ({it['sku']})"
        names[nm] = 1
        t = kind(nm); cat = category(nm, t); title = f"{nm} - {b}"
        s = slug(nm); base, n = s, 2
        while s in slugs: s, n = f"{base}-{n}", n + 1
        slugs.add(s)
        d, plain = build_desc(nm, t, it["desc"], i, b); words.append(len(re.sub("<[^>]+>", " ", d).split()))
        seo = plain if len(plain) <= 160 else plain[:157].rsplit(" ", 1)[0] + "..."
        paths = [f"images/{it['sku']}/{k:02d}.jpg" for k in range(1, len(it["srcimgs"]) + 1)]
        urls = [GH + p for p in paths]
        c = it["cost"]
        vals = [urls[0], ",".join(urls[1:]) or None, None, title, rnd(c * 1.5 + 100), rnd(c), ",".join(cat), b, d, "نعم", 1, it["sku"], "نعم", title, s, seo,
                rnd(c * 1.25 + 30), None, "لا", "الخدمات المالية", "لا"]
        for j, v in enumerate(vals): ws.cell(row=row, column=j + 1, value=v)
        row += 1
        note = it.get("note", "") or ("اسم ناقص - راجعه" if not toks(nm, b) else "")
        w2.append([it["sku"], title, it["src"], rnd(c), rnd(c * 1.5 + 100), rnd(c * 1.25 + 30), ",".join(cat), len(urls), it["match"], note])
        manifest += [[u, p] for u, p in zip(it["srcimgs"], paths)]
    w3 = rv.create_sheet("مطابقات الموقع الرسمي"); w3.sheet_view.rightToLeft = True
    w3.append(["منتج الموقع الرسمي", "مطابق مع بوتيكات", "درجة التطابق", "متوفر في بوتيكات"])
    for oid, (sc, k) in sorted(matches.items(), key=lambda x: x[1][0]):
        w3.append([off[oid]["name"], bq[k]["name"], round(sc, 2), "نعم" if bq[k]["avail"] else "لا"])
    os.makedirs(f"output/{sl}", exist_ok=True)
    wb.save(f"output/{sl}/import.xlsx"); rv.save(f"output/{sl}/review.xlsx")
    json.dump(manifest, open(f"manifests/{sl}.json", "w", encoding="utf-8"), ensure_ascii=False, indent=0)
    from collections import Counter
    summary = dict(brand=b, products=row - 2, sources=Counter(x["src"] for x in items if not x.get("skipped")), skipped_no_images=sum(1 for x in items if x.get("skipped")),
                   images=len(manifest), words_min=min(words or [0]), words_avg=(sum(words) // len(words)) if words else 0,
                   official_matched=len(matches), official_total=len(off), log=LOG)
    json.dump(summary, open(f"output/{sl}/summary.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=dict)
    log(json.dumps({k: v for k, v in summary.items() if k != "log"}, ensure_ascii=False, default=dict))

if __name__ == "__main__":
    main(os.path.abspath(sys.argv[1]), sys.argv[2] if len(sys.argv) > 2 else ".")
