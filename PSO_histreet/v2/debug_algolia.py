# -*- coding: utf-8 -*-
"""Debug: dump full product card structure from PakWheels."""
import json, io, sys
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

    page.goto(
        "https://www.pakwheels.com/accessories-spare-parts/search/-/buynow_1/ctg_oils-lubricants/sctg_engine-oil/",
        wait_until="domcontentloaded", timeout=40_000
    )
    page.wait_for_timeout(4000)

    # Find product cards by tracing from price elements upward
    data = page.evaluate("""() => {
        const out = [];
        // Price elements are in <strong class='generic-white fs18'>
        const priceEls = document.querySelectorAll("strong.fs18");
        for (const pe of priceEls) {
            // Walk up to find the card container (li or article or div with product info)
            let card = pe;
            for (let i = 0; i < 8; i++) {
                card = card.parentElement;
                if (!card) break;
                // Check if this ancestor has a title/link
                const link = card.querySelector('a[title], h2 a, h3 a, [class*="title"] a, a[href*="accessories"]');
                if (link) {
                    const title = link.getAttribute('title') || link.innerText.trim();
                    const href  = link.href;
                    const origEl = card.querySelector('.discount-strike');
                    const orig   = origEl ? origEl.innerText.trim() : '';
                    if (title && title.length > 3) {
                        out.push({
                            title,
                            price: pe.innerText.trim(),
                            orig,
                            href,
                            cardTag: card.tagName,
                            cardClass: (card.className||'').slice(0,80),
                        });
                        break;
                    }
                }
            }
        }
        return out;
    }""")
    print(f"Products found via card walk: {len(data)}")
    for item in data[:10]:
        print(f"  [{item['cardTag']} .{item['cardClass'][:40]}]")
        print(f"    title: {item['title'][:80]}")
        print(f"    price: {item['price']}  orig: {item['orig']}")
        print(f"    href:  {item['href'][:80]}")

    # Also try direct title selector patterns
    print("\n\nTitle element candidates:")
    titles = page.evaluate("""() => {
        const sels = [
            'a[title]', 'h2 a', 'h3 a', 'h4 a',
            '[class*="title"]', '[class*="product-name"]',
            'li a[href*="accessories"]',
        ];
        const out = [];
        for (const sel of sels) {
            const els = document.querySelectorAll(sel);
            for (const el of [...els].slice(0, 5)) {
                const t = el.getAttribute('title') || el.innerText?.trim() || '';
                const h = el.href || '';
                if (t && t.length > 5 && h.includes('accessories'))
                    out.push({sel, t: t.slice(0,80), h: h.slice(0,80)});
            }
        }
        return out;
    }""")
    for t in titles[:20]:
        print(f"  [{t['sel']}] {t['t']}")
        print(f"    -> {t['h']}")

    ctx.close()
    browser.close()
