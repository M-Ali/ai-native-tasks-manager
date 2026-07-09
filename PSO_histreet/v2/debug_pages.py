# -*- coding: utf-8 -*-
"""Dumps page HTML and DOM info to debug scraper selectors."""
import io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from pathlib import Path
from playwright.sync_api import sync_playwright

OUT = Path("output/debug")
OUT.mkdir(parents=True, exist_ok=True)

with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=True)
    ctx = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/125.0.0.0 Safari/537.36",
        viewport={"width": 1366, "height": 900},
    )
    page = ctx.new_page()

    # ── PakWheels ──
    print("Fetching PakWheels engine oil page...")
    page.goto(
        "https://www.pakwheels.com/accessories-spare-parts/search/-/buynow_1/ctg_oils-lubricants/sctg_engine-oil/",
        wait_until="domcontentloaded", timeout=40_000
    )
    page.wait_for_timeout(5000)
    page.screenshot(path=str(OUT / "pakwheels.png"), full_page=False)

    html = page.content()
    (OUT / "pakwheels.html").write_text(html[:100000], encoding="utf-8")
    print(f"  HTML length: {len(html)}")
    print(f"  Title: {page.title()}")

    classes = page.evaluate("""() => {
        const cls = new Set();
        document.querySelectorAll('[class]').forEach(el => {
            (el.className || '').split(' ').forEach(c => { if(c) cls.add(c); });
        });
        return [...cls].sort();
    }""")
    print(f"  Unique classes (first 80): {classes[:80]}")

    # Find product-like elements
    info = page.evaluate("""() => {
        const results = [];
        const all = document.querySelectorAll('*');
        let count = 0;
        for (const el of all) {
            const t = (el.innerText || '').trim();
            if (count < 20 && t.length > 3 && t.length < 300
                && (t.indexOf('PKR') !== -1 || t.indexOf('Rs') !== -1)) {
                results.push({
                    tag: el.tagName,
                    cls: (el.className || '').slice(0, 80),
                    text: t.slice(0, 120)
                });
                count++;
            }
        }
        return results;
    }""")
    print(f"\n  Elements with PKR/Rs text:")
    for el in info:
        print(f"    <{el['tag']} class='{el['cls']}'> {el['text'][:100]}")

    # Dump first 5 list items or article elements
    items_info = page.evaluate("""() => {
        const results = [];
        const candidates = [
            ...document.querySelectorAll('li'),
            ...document.querySelectorAll('article'),
            ...document.querySelectorAll('[class*="item"]'),
            ...document.querySelectorAll('[class*="product"]'),
            ...document.querySelectorAll('[class*="listing"]'),
        ];
        let count = 0;
        for (const el of candidates) {
            const t = (el.innerText || '').trim();
            if (t.length > 20 && t.length < 500 && count < 10) {
                results.push({
                    tag: el.tagName,
                    cls: (el.className || '').slice(0, 100),
                    text: t.slice(0, 200),
                    childCount: el.children.length,
                    html: el.innerHTML.slice(0, 300)
                });
                count++;
            }
        }
        return results;
    }""")
    print(f"\n  Product-like elements (first 10):")
    for el in items_info:
        print(f"\n    <{el['tag']} class='{el['cls']}'>")
        print(f"    text: {el['text'][:150]}")
        print(f"    innerHTML: {el['html'][:200]}")

    # ── Metro Online product detail ──
    print("\n\nFetching Metro Online product detail page...")
    page.goto(
        "https://www.metro-online.pk/nf_shopping/detail/automotive/car-care/havoline-motor-oil-4l-w-ds-20w50/12634731",
        wait_until="domcontentloaded", timeout=40_000
    )
    page.wait_for_timeout(5000)
    page.screenshot(path=str(OUT / "metro_product.png"), full_page=False)

    metro_html = page.content()
    (OUT / "metro_product.html").write_text(metro_html[:100000], encoding="utf-8")
    print(f"  HTML length: {len(metro_html)}")
    print(f"  Title: {page.title()}")

    metro_classes = page.evaluate("""() => {
        const cls = new Set();
        document.querySelectorAll('[class]').forEach(el => {
            (el.className || '').split(' ').forEach(c => { if(c) cls.add(c); });
        });
        return [...cls].sort();
    }""")
    print(f"  Unique classes (first 60): {metro_classes[:60]}")

    metro_info = page.evaluate("""() => {
        const out = {};
        out.h1 = (document.querySelector('h1')?.innerText || '').slice(0, 100);
        out.windowKeys = Object.keys(window).filter(k => k.startsWith('__')).slice(0, 20);
        out.metaTags = [...document.querySelectorAll('meta[name],meta[property]')]
            .map(m => m.name || m.getAttribute('property') + '=' + m.content)
            .slice(0, 15);
        const bodyText = (document.body?.innerText || '').slice(0, 2000);
        out.bodySnippet = bodyText;
        return out;
    }""")
    print(f"  h1: {metro_info.get('h1')}")
    print(f"  Window keys: {metro_info.get('windowKeys')}")
    print(f"  Body snippet: {metro_info.get('bodySnippet', '')[:500]}")

    # Metro search page
    print("\nFetching Metro search page for 'havoline'...")
    page.goto(
        "https://www.metro-online.pk/nf_shopping/search?query=havoline",
        wait_until="domcontentloaded", timeout=30_000
    )
    page.wait_for_timeout(5000)
    page.screenshot(path=str(OUT / "metro_search.png"), full_page=False)
    s_html = page.content()
    (OUT / "metro_search.html").write_text(s_html[:100000], encoding="utf-8")
    print(f"  Title: {page.title()}, HTML len: {len(s_html)}")
    print(f"  URL after load: {page.url}")
    body = page.evaluate("() => (document.body?.innerText || '').slice(0, 1000)")
    print(f"  Body: {body[:500]}")

    ctx.close()
    browser.close()

print("\nDebug files: output/debug/")
