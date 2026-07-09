# -*- coding: utf-8 -*-
"""
Scraper for PakWheels (Engine Oil + Gear Oil) and Metro Online (car-care search).
PakWheels: server-rendered HTML -- div.well.search-list cards.
Metro Online: JS-rendered -- JSON-LD / DOM fallback.
Output: output/reports/competitor_prices_extended.csv

Run:  uv run python scrape_new_platforms.py
"""
import re, csv, time, sys, io
from dataclasses import dataclass
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
OUT_CSV = Path("output/reports/competitor_prices_extended.csv")
OUT_CSV.parent.mkdir(parents=True, exist_ok=True)

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
      "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36")

# ── Normalisation ─────────────────────────────────────────────────────────────

_GRADE_RE = re.compile(
    r"\b(0W-?20|5W-?20|5W-?30|5W-?40|10W-?30|10W-?40|15W-?40|"
    r"20W-?50|25W-?50|20W-?40)\b", re.I)
_ML_RE = re.compile(r"\b(\d+)\s*ML\b", re.I)
_L_RE  = re.compile(r"\b(\d+(?:\.\d+)?)\s*(?:litre|liter|ltr|lt|L)\b", re.I)
_HDEO_RE = re.compile(r"\b(rimula|delo|rubia|tir|diesel|fleet|ci-4|ch-4|cj-4|isosyn)\b", re.I)
_MCO_RE  = re.compile(r"\b(4t|4-t|motorcycle|moto|scooter|super\s*4t|havoline\s*4t)\b", re.I)

_BRANDS = [
    (re.compile(r"\bshell\s+helix\s+ultra\b", re.I),    "Shell Helix Ultra"),
    (re.compile(r"\bshell\s+helix\s+hx\s*8\b", re.I),   "Shell Helix HX8"),
    (re.compile(r"\bshell\s+helix\s+hx\s*7\b", re.I),   "Shell Helix HX7"),
    (re.compile(r"\bshell\s+helix\s+hx\s*5\b", re.I),   "Shell Helix HX5"),
    (re.compile(r"\bshell\s+helix\s+hx\s*3\b", re.I),   "Shell Helix HX3"),
    (re.compile(r"\bshell\s+helix\b", re.I),              "Shell Helix"),
    (re.compile(r"\bshell\s+rimula\s+r4\b", re.I),       "Shell Rimula R4"),
    (re.compile(r"\bshell\s+rimula\s+r3\b", re.I),       "Shell Rimula R3"),
    (re.compile(r"\bshell\s+rimula\b", re.I),             "Shell Rimula"),
    (re.compile(r"\bshell\b", re.I),                       "Shell"),
    (re.compile(r"\bhavoline\s+pro.?ds\b", re.I),         "Havoline Pro-DS"),
    (re.compile(r"\bhavoline\s+ultra\b", re.I),           "Havoline Ultra"),
    (re.compile(r"\bhavoline\s+super\s+4t\b", re.I),     "Havoline Super 4T"),
    (re.compile(r"\bhavoline\s+4t\b", re.I),              "Havoline 4T"),
    (re.compile(r"\bhavoline\b", re.I),                    "Havoline"),
    (re.compile(r"\bcaltex\s+havoline\b", re.I),          "Caltex Havoline"),
    (re.compile(r"\bcaltex\s+delo\b", re.I),              "Caltex Delo"),
    (re.compile(r"\bcaltex\b", re.I),                      "Caltex"),
    (re.compile(r"\bdelo\s+gold\b", re.I),                "Delo Gold"),
    (re.compile(r"\bdelo\s+silver\b", re.I),              "Delo Silver"),
    (re.compile(r"\bdelo\b", re.I),                        "Delo"),
    (re.compile(r"\btotal\s+quartz\s+9000\b", re.I),     "Total Quartz 9000"),
    (re.compile(r"\btotal\s+quartz\s+7000\b", re.I),     "Total Quartz 7000"),
    (re.compile(r"\btotal\s+quartz\s+5000\b", re.I),     "Total Quartz 5000"),
    (re.compile(r"\btotal\s+quartz\b", re.I),             "Total Quartz"),
    (re.compile(r"\btotal\s+rubia\b", re.I),              "Total Rubia"),
    (re.compile(r"\btotalenergies\b", re.I),               "Total"),
    (re.compile(r"\btotal\b", re.I),                       "Total"),
    (re.compile(r"\bzic\s+x9\b", re.I),                   "ZIC X9"),
    (re.compile(r"\bzic\s+x7\b", re.I),                   "ZIC X7"),
    (re.compile(r"\bzic\s+x5\b", re.I),                   "ZIC X5"),
    (re.compile(r"\bzic\b", re.I),                         "ZIC"),
    (re.compile(r"\bkixx\s+pao\b", re.I),                 "Kixx PAO"),
    (re.compile(r"\bkixx\s+g1\b", re.I),                  "Kixx G1"),
    (re.compile(r"\bkixx\s+d1\b", re.I),                  "Kixx D1"),
    (re.compile(r"\bkixx\b", re.I),                        "Kixx"),
    (re.compile(r"\btotachi\b", re.I),                     "Totachi"),
    (re.compile(r"\bmobil\s+1\b", re.I),                  "Mobil 1"),
    (re.compile(r"\bmobil\s+super\b", re.I),              "Mobil Super"),
    (re.compile(r"\bmobil\b", re.I),                       "Mobil"),
    (re.compile(r"\bcastrol\s+edge\b", re.I),             "Castrol EDGE"),
    (re.compile(r"\bcastrol\s+magnatec\b", re.I),         "Castrol Magnatec"),
    (re.compile(r"\bcastrol\s+gtx\b", re.I),              "Castrol GTX"),
    (re.compile(r"\bcastrol\b", re.I),                     "Castrol"),
    (re.compile(r"\bhonda\s+genuine\b", re.I),            "Honda Genuine"),
    (re.compile(r"\bhonda\b", re.I),                       "Honda"),
    (re.compile(r"\bsuzuki\b", re.I),                      "Suzuki"),
    (re.compile(r"\byamaha\b", re.I),                      "Yamaha"),
    (re.compile(r"\bvalvoline\b", re.I),                   "Valvoline"),
    (re.compile(r"\bpennzoil\b", re.I),                    "Pennzoil"),
    (re.compile(r"\baramco\b", re.I),                      "Aramco"),
    (re.compile(r"\bequimoli\b", re.I),                    "Equimoli"),
    (re.compile(r"\bcarient\s+ultra\b", re.I),            "PSO Carient Ultra"),
    (re.compile(r"\bcarient\s+fs\b", re.I),               "PSO Carient FS"),
    (re.compile(r"\bcarient\s+plus\b", re.I),             "PSO Carient Plus"),
    (re.compile(r"\bcarient\s+spro\b", re.I),             "PSO Carient SPRO"),
    (re.compile(r"\bcarient\b", re.I),                     "PSO Carient"),
    (re.compile(r"\bpso\b", re.I),                         "PSO"),
    (re.compile(r"\bwurth\b|wuerth\b", re.I),             "Wurth"),
    (re.compile(r"\bsinopec\b", re.I),                     "Sinopec"),
    (re.compile(r"\bautol\b", re.I),                       "Autol"),
]


@dataclass
class Product:
    platform:       str = ""
    title:          str = ""
    brand:          str = ""
    grade:          str = ""
    pack_l:         float = 0.0
    oil_type:       str = ""
    price_pkr:      float = 0.0
    price_per_l:    float = 0.0
    original_price: float = 0.0
    discount_pct:   float = 0.0
    url:            str = ""


def _px(t):
    try: return float(re.sub(r"[^\d.]", "", str(t).replace(",", ""))) or 0.0
    except: return 0.0

def _grade(t):
    m = _GRADE_RE.search(t)
    if not m: return ""
    raw = m.group(1).upper()
    return raw if "W-" in raw else raw.replace("W", "W-")

def _pack(t):
    m = _ML_RE.search(t)
    if m: return round(int(m.group(1)) / 1000, 3)
    m = _L_RE.search(t)
    if m: return round(float(m.group(1)), 3)
    return 0.0

def _brand(t):
    for pat, name in _BRANDS:
        if pat.search(t): return name
    return ""

def _type(t):
    if _MCO_RE.search(t):  return "MCO"
    if _HDEO_RE.search(t): return "HDEO"
    return "PCMO"

def make(platform, title, price, url="", orig=0.0):
    p = Product(platform=platform, title=title.strip(),
                price_pkr=price, url=url, original_price=orig)
    p.brand   = _brand(title)
    p.grade   = _grade(title)
    p.pack_l  = _pack(title)
    p.oil_type= _type(title)
    if p.pack_l > 0:
        p.price_per_l = round(price / p.pack_l, 0)
    if orig > 0 and price < orig:
        p.discount_pct = round((orig - price) / orig * 100, 1)
    return p


# ── PakWheels scraper ─────────────────────────────────────────────────────────
# Card:  div.well.search-list  (class contains "well" and "search-list")
# Title: a[title] inside the card
# Price: strong.fs18 (innerText = "PKR 1,290")
# Orig:  .discount-strike (innerText = "PKR 1,440")

PW_BASE = "https://www.pakwheels.com"
PW_SECTIONS = [
    ("engine-oil", PW_BASE + "/accessories-spare-parts/search/-/buynow_1/ctg_oils-lubricants/sctg_engine-oil/"),
    ("gear-oil",   PW_BASE + "/accessories-spare-parts/search/-/buynow_1/ctg_oils-lubricants/sctg_gear-oil/"),
    ("atf-oil",    PW_BASE + "/accessories-spare-parts/search/-/buynow_1/ctg_oils-lubricants/sctg_car-atf-mtf-oil/"),
]
PW_MAX_PAGES = 40


def _scrape_pw_page(page) -> list[Product]:
    data = page.evaluate("""() => {
        const out = [];
        // Walk from price elements up to find the card anchor
        const priceEls = [...document.querySelectorAll("strong.fs18")];
        for (const pe of priceEls) {
            let node = pe;
            for (let i = 0; i < 10; i++) {
                node = node.parentElement;
                if (!node) break;
                const link = node.querySelector('a[title]');
                if (link) {
                    const title    = link.getAttribute('title') || link.innerText.trim();
                    const origEl   = node.querySelector('.discount-strike');
                    const orig     = origEl ? origEl.innerText.trim() : '';
                    const href     = link.href;
                    if (title.length > 3)
                        out.push({title, price: pe.innerText.trim(), orig, href});
                    break;
                }
            }
        }
        return out;
    }""")
    products = []
    for item in (data or []):
        price = _px(item["price"])
        orig  = _px(item["orig"])
        if price > 50:
            products.append(make("pakwheels", item["title"], price, item["href"], orig))
    return products


def scrape_pakwheels(page) -> list[Product]:
    all_products: list[Product] = []
    for section_name, base_url in PW_SECTIONS:
        print(f"\n  [PakWheels / {section_name}]")
        for pg in range(1, PW_MAX_PAGES + 1):
            url = base_url if pg == 1 else f"{base_url}?page={pg}"
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=35_000)
                page.wait_for_timeout(2500)
            except PWTimeout:
                print(f"    Timeout page {pg}")
                break

            pg_prods = _scrape_pw_page(page)
            all_products.extend(pg_prods)
            print(f"    page {pg:02d}: +{len(pg_prods):3d}  cumulative={len(all_products)}")

            if not pg_prods:
                print(f"    Empty — stopping {section_name}")
                break

            # Check for next page link
            nxt = page.query_selector("link[rel='next'], a[rel='next']")
            if not nxt and pg > 1:
                break

            time.sleep(1.0)
    return all_products


# ── Metro Online scraper ──────────────────────────────────────────────────────
# Metro Online (NF Shopping platform) is JS-rendered.
# Strategy: navigate to product detail page URL pattern to find category listing.
# Also try their internal search.

METRO_BASE    = "https://www.metro-online.pk"
METRO_TERMS   = [
    "engine oil", "havoline", "shell helix",
    "total quartz", "castrol", "mobil", "zic oil",
    "delo", "gear oil", "4t motorcycle oil",
]
METRO_CAT_URLS = [
    METRO_BASE + "/nf_shopping/automotive/car-care",
    METRO_BASE + "/nf_shopping/search?query=engine+oil",
    METRO_BASE + "/nf_shopping/search?q=engine+oil",
]


def _extract_metro_products(page, term: str) -> list[Product]:
    """Extract products from current Metro page."""
    products = []

    # JSON-LD (Metro often uses this for product data)
    ld = page.evaluate("""() => {
        const out = [];
        document.querySelectorAll('script[type="application/ld+json"]').forEach(s => {
            try {
                const d = JSON.parse(s.innerText || s.textContent);
                if (d['@type'] === 'Product') {
                    const offers = d.offers || {};
                    out.push({
                        name:  d.name || '',
                        price: String(offers.price || offers.lowPrice || ''),
                        url:   d.url || window.location.href,
                    });
                } else if (d['@type'] === 'ItemList') {
                    for (const it of d.itemListElement || []) {
                        const x = it.item || it;
                        out.push({
                            name:  x.name || '',
                            price: String((x.offers||{}).price || ''),
                            url:   x.url || '',
                        });
                    }
                }
            } catch(e) {}
        });
        return out;
    }""")
    for item in (ld or []):
        t = (item.get("name") or "").strip()
        p = _px(item.get("price", ""))
        if t and p > 50:
            products.append(make("metro-online", t, p, item.get("url", "")))

    if products:
        return products

    # DOM fallback — scroll and look for product cards
    for _ in range(3):
        page.evaluate("window.scrollBy(0, window.innerHeight)")
        page.wait_for_timeout(700)

    dom = page.evaluate("""() => {
        const out = [];
        const tried = [
            '[class*="product-card"]', '[class*="product-item"]',
            '[class*="item-card"]',    '[class*="nf-product"]',
            '[class*="tile"]',         'article',
            '[data-product]',          '[data-sku]',
        ];
        let cards = [];
        for (const s of tried) {
            cards = [...document.querySelectorAll(s)];
            if (cards.length > 1) break;
        }
        for (const c of cards.slice(0, 200)) {
            const tEl = c.querySelector('h1,h2,h3,h4,[class*="title"],[class*="name"]');
            const pEl = c.querySelector('[class*="price"],[class*="amount"],[class*="cost"]');
            const a   = c.querySelector('a[href]');
            const t   = (tEl?.innerText || '').trim();
            const p   = (pEl?.innerText || '').trim();
            if (t && p && t.length > 4)
                out.push({t, p, u: a?.href || ''});
        }
        return out;
    }""")
    for item in (dom or []):
        p = _px(item["p"])
        if p > 50:
            products.append(make("metro-online", item["t"], p, item.get("u", "")))

    return products


def scrape_metro(page) -> list[Product]:
    all_products: list[Product] = []
    seen: set[str] = set()

    # Try category listing pages first
    print("\n  [Metro Online] trying category pages...")
    for cat_url in METRO_CAT_URLS:
        try:
            page.goto(cat_url, wait_until="domcontentloaded", timeout=25_000)
            page.wait_for_timeout(4000)
            found = _extract_metro_products(page, "category")
            new = 0
            for pr in found:
                key = pr.title.lower()[:70]
                if key not in seen:
                    seen.add(key)
                    all_products.append(pr)
                    new += 1
            if new:
                print(f"    {cat_url} -> {new} products")
        except PWTimeout:
            pass

    # Search-based extraction
    print("  [Metro Online] search terms...")
    for term in METRO_TERMS:
        q = term.replace(" ", "+")
        search_urls = [
            f"{METRO_BASE}/nf_shopping/search?query={q}",
            f"{METRO_BASE}/nf_shopping/search?q={q}",
        ]
        loaded = False
        for su in search_urls:
            try:
                page.goto(su, wait_until="domcontentloaded", timeout=22_000)
                page.wait_for_timeout(4000)
                title = page.title()
                if "404" not in title and "Error" not in title:
                    loaded = True
                    break
            except PWTimeout:
                continue

        if not loaded:
            # Try through search box on homepage
            try:
                page.goto(METRO_BASE, wait_until="domcontentloaded", timeout=20_000)
                page.wait_for_timeout(2000)
                si = page.query_selector('input[type="search"],input[name="q"],input[name="query"],#search,input[placeholder*="search" i]')
                if si:
                    si.click()
                    si.fill(term)
                    si.press("Enter")
                    page.wait_for_timeout(4000)
                    loaded = True
            except Exception:
                pass

        if not loaded:
            print(f"    '{term}': could not load")
            continue

        found = _extract_metro_products(page, term)
        new = 0
        for pr in found:
            key = pr.title.lower()[:70]
            if key not in seen:
                seen.add(key)
                all_products.append(pr)
                new += 1
        print(f"    '{term}': {new} new ({len(all_products)} total)")
        time.sleep(1.5)

    return all_products


# ── CSV + summary ─────────────────────────────────────────────────────────────

FIELDS = ["platform","title","brand","grade","pack_l","oil_type",
          "price_pkr","price_per_l","original_price","discount_pct","url"]


def write_csv(products: list[Product]):
    with open(OUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        for p in products:
            w.writerow({k: getattr(p, k) for k in FIELDS})
    print(f"\nSaved {len(products)} rows -> {OUT_CSV}")


def summary(products: list[Product]):
    from collections import Counter
    print(f"\n{'='*60}")
    print(f"TOTAL: {len(products)} products")
    for label, fn, n in [
        ("Platform",       lambda p: p.platform,                       99),
        ("Brand (top 20)", lambda p: p.brand or "(unknown)",           20),
        ("Grade (top 10)", lambda p: p.grade or "(no grade detected)", 10),
        ("Oil Type",       lambda p: p.oil_type,                       99),
    ]:
        c = Counter(fn(p) for p in products)
        print(f"\n{label}:")
        for k, v in c.most_common(n):
            print(f"  {k:<32} {v:>4}")


def main():
    products: list[Product] = []

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        ctx = browser.new_context(user_agent=UA, viewport={"width": 1366, "height": 900})
        page = ctx.new_page()

        print("=== [1/2] PakWheels ===")
        pw_prods = scrape_pakwheels(page)
        print(f"\nPakWheels total: {len(pw_prods)}")
        products.extend(pw_prods)

        print("\n=== [2/2] Metro Online ===")
        metro_prods = scrape_metro(page)
        print(f"Metro total: {len(metro_prods)}")
        products.extend(metro_prods)

        ctx.close()
        browser.close()

    # Dedup
    seen: set[str] = set()
    unique: list[Product] = []
    for p in products:
        key = f"{p.platform}|{p.title.lower()[:80]}"
        if key not in seen:
            seen.add(key)
            unique.append(p)

    summary(unique)
    write_csv(unique)


if __name__ == "__main__":
    main()
