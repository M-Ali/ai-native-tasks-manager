"""
Daraz.pk scraper using Playwright.
Intercepts Daraz's internal search API response (JSON) rather than
parsing HTML class names — more reliable as Daraz frequently changes
its frontend CSS classes.
"""
import time
import json
import re
import os
from urllib.parse import quote_plus

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout
from rich.console import Console

from .base import BaseScraper, ScrapedProduct

console = Console()
MAX_RETRIES = 3


class DarazScraper(BaseScraper):
    platform_name = "daraz"

    def scrape(self, search_term: str) -> list[ScrapedProduct]:
        cfg = self.platform_cfg
        max_pages = cfg.get("max_pages", 3)
        delay = cfg.get("delay_seconds", 2)
        headless = cfg.get("headless", True)

        browsers_path = cfg.get("browsers_path")
        if browsers_path:
            os.environ["PLAYWRIGHT_BROWSERS_PATH"] = browsers_path

        products: list[ScrapedProduct] = []
        base_url = cfg.get("base_url", "https://www.daraz.pk")

        for attempt in range(1, MAX_RETRIES + 1):
            try:
                products = self._scrape_with_browser(
                    search_term, base_url, max_pages, delay, headless
                )
                break
            except Exception as e:
                err = str(e)
                if attempt < MAX_RETRIES and (
                    "ERR_NETWORK_CHANGED" in err
                    or "ERR_CONNECTION" in err
                    or "Timeout" in err
                ):
                    console.print(f"  [yellow]Retry {attempt}/{MAX_RETRIES} for '{search_term}': {err[:60]}[/yellow]")
                    time.sleep(3 * attempt)
                else:
                    console.print(f"  [red]Failed '{search_term}' after {attempt} attempts: {err[:80]}[/red]")
                    break

        console.print(f"  [green]Daraz: {len(products)} products for '{search_term}'[/green]")
        return products

    def _scrape_with_browser(
        self, search_term: str, base_url: str, max_pages: int, delay: float, headless: bool
    ) -> list[ScrapedProduct]:
        products = []

        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=headless)
            ctx = browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/125.0.0.0 Safari/537.36"
                ),
                viewport={"width": 1366, "height": 900},
            )
            page = ctx.new_page()

            for pg in range(1, max_pages + 1):
                q = quote_plus(search_term)
                url = f"{base_url}/catalog/?q={q}&page={pg}"
                console.print(f"  [dim]  [{search_term}] page {pg}[/dim]")

                # Capture API JSON via network interception
                captured = []

                def handle_response(resp):
                    try:
                        if (
                            "lazada-search-main-search" in resp.url
                            or ("catalog" in resp.url and "ajax" in resp.url)
                            or ("search" in resp.url and resp.headers.get("content-type", "").startswith("application/json"))
                        ):
                            data = resp.json()
                            captured.append(data)
                    except Exception:
                        pass

                page.on("response", handle_response)

                try:
                    page.goto(url, wait_until="domcontentloaded", timeout=30_000)
                    # Give JS time to fire XHR / render
                    page.wait_for_timeout(delay * 1000)
                except PlaywrightTimeout:
                    console.print(f"  [yellow]  Timeout page {pg} — skipping[/yellow]")
                    break

                page.remove_listener("response", handle_response)

                # Try API data first
                page_products = []
                for data in captured:
                    page_products.extend(_parse_api_response(data, search_term, base_url))

                # Fall back to DOM parsing if API capture missed
                if not page_products:
                    page_products = _parse_dom(page, search_term, base_url)

                products.extend(page_products)

                if len(page_products) < 5:
                    break  # likely last page

                time.sleep(delay)

            ctx.close()
            browser.close()

        return products


# ── API response parser ───────────────────────────────────────────

def _parse_api_response(data: dict, search_term: str, base_url: str) -> list[ScrapedProduct]:
    """Parse Daraz's internal search API JSON."""
    products = []

    # Daraz API structure varies; try common paths
    items = (
        data.get("mods", {}).get("listItems", [])
        or data.get("rgn", {}).get("conts", [])
        or data.get("items", [])
        or data.get("data", {}).get("items", [])
        or []
    )

    for item in items:
        try:
            title = (
                item.get("name")
                or item.get("title")
                or item.get("itemTitle", "")
            )
            if not title:
                continue

            price_raw = (
                item.get("price")
                or item.get("priceShow")
                or item.get("salePrice")
                or ""
            )
            price = _parse_price(str(price_raw))
            if not price or price <= 0:
                continue

            orig_raw = item.get("originalPrice") or item.get("originalPriceShow")
            orig_price = _parse_price(str(orig_raw)) if orig_raw else None

            item_url = item.get("itemUrl") or item.get("url") or ""
            if item_url and not item_url.startswith("http"):
                item_url = base_url + item_url

            seller = item.get("sellerName") or item.get("seller")
            rating = _parse_float(str(item.get("ratingScore") or item.get("rating") or ""))
            review_count = _parse_int(str(item.get("review") or item.get("reviewCount") or ""))

            products.append(ScrapedProduct(
                platform="daraz",
                brand_query=search_term,
                title=title,
                price=price,
                original_price=orig_price,
                url=item_url,
                seller=seller,
                rating=rating,
                review_count=review_count,
            ))
        except Exception:
            continue

    return products


# ── DOM fallback parser ───────────────────────────────────────────

_CARD_SELECTORS = [
    # Daraz search grid items — try several in order
    "[data-qa-locator='product-item']",
    ".gridItem--Yd0sa",
    "._17mcb",
    ".ant-col",
    "[class*='gridItem']",
    "[class*='product-item']",
]

_TITLE_SELECTORS = [
    "[class*='title']",
    "[class*='name']",
    "a[title]",
    "h3",
]

_PRICE_SELECTORS = [
    "[class*='price--'][class*='current']",
    ".ooOxS",
    "[class*='priceBox'] span:first-child",
    "[class*='price']",
]


def _parse_dom(page, search_term: str, base_url: str) -> list[ScrapedProduct]:
    products = []

    cards = []
    for sel in _CARD_SELECTORS:
        cards = page.query_selector_all(sel)
        if len(cards) > 2:
            break

    if not cards:
        # Last resort: extract structured data from JSON-LD or window.__INITIAL_DATA__
        products = _extract_from_page_data(page, search_term, base_url)
        return products

    for card in cards:
        try:
            title = None
            for sel in _TITLE_SELECTORS:
                el = card.query_selector(sel)
                if el:
                    title = (el.get_attribute("title") or el.inner_text()).strip()
                    if title:
                        break
            if not title:
                continue

            price_text = ""
            for sel in _PRICE_SELECTORS:
                el = card.query_selector(sel)
                if el:
                    price_text = el.inner_text().strip()
                    if price_text:
                        break

            price = _parse_price(price_text)
            if not price or price <= 0:
                continue

            link_el = card.query_selector("a")
            href = link_el.get_attribute("href") if link_el else ""
            full_url = href if (href and href.startswith("http")) else base_url + (href or "")

            products.append(ScrapedProduct(
                platform="daraz",
                brand_query=search_term,
                title=title,
                price=price,
                original_price=None,
                url=full_url,
                seller=None,
                rating=None,
                review_count=None,
            ))
        except Exception:
            continue

    return products


def _extract_from_page_data(page, search_term: str, base_url: str) -> list[ScrapedProduct]:
    """Extract product data from Daraz's window.__INITIAL_DATA__ or JSON-LD."""
    products = []
    try:
        raw = page.evaluate("""() => {
            try {
                const d = window.__INITIAL_DATA__ || window.pageData || {};
                return JSON.stringify(d);
            } catch(e) { return '{}'; }
        }""")
        data = json.loads(raw) if raw else {}
        products = _parse_api_response(data, search_term, base_url)
    except Exception:
        pass

    # Try JSON-LD as last resort
    if not products:
        try:
            ld_scripts = page.query_selector_all("script[type='application/ld+json']")
            for script in ld_scripts:
                content = script.inner_text()
                obj = json.loads(content)
                if obj.get("@type") in ("ItemList", "Product"):
                    items = obj.get("itemListElement", [obj])
                    for item in items:
                        offer = item.get("offers", {})
                        price = _parse_price(str(offer.get("price", "")))
                        title = item.get("name", "")
                        if title and price:
                            products.append(ScrapedProduct(
                                platform="daraz",
                                brand_query=search_term,
                                title=title,
                                price=price,
                                original_price=None,
                                url=item.get("url", ""),
                                seller=None,
                                rating=None,
                                review_count=None,
                            ))
        except Exception:
            pass

    return products


# ── Helpers ───────────────────────────────────────────────────────

def _parse_price(text: str) -> float | None:
    if not text:
        return None
    cleaned = re.sub(r"[^\d.]", "", str(text).replace(",", ""))
    try:
        v = float(cleaned)
        return v if v > 0 else None
    except ValueError:
        return None


def _parse_float(text: str) -> float | None:
    try:
        return float(re.sub(r"[^\d.]", "", text))
    except (ValueError, TypeError):
        return None


def _parse_int(text: str) -> int | None:
    try:
        return int(re.sub(r"[^\d]", "", text))
    except (ValueError, TypeError):
        return None
