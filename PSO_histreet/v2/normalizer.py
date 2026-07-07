"""
Normalizes raw ScrapedProduct titles into structured fields:
  - brand_detected
  - grade_detected  (e.g. 5W-30)
  - pack_size_l     (in litres)
  - price_per_litre
  - oil_type        (pcmo / hdeo / mco)
"""
import re
from scrapers.base import ScrapedProduct


# Regex patterns
_GRADE_RE = re.compile(
    r"\b(0W-?20|5W-?20|5W-?30|5W-?40|10W-?30|10W-?40|15W-?40|20W-?50|25W-?50)\b",
    re.IGNORECASE,
)

_PACK_PATTERNS = [
    (re.compile(r"\b(\d+(?:\.\d+)?)\s*(?:litre|liter|ltr|lt|L)\b", re.I), 1.0),
    (re.compile(r"\b(\d+(?:\.\d+)?)\s*(?:ml|ML)\b", re.I), 0.001),
    (re.compile(r"\b(\d+(?:\.\d+)?)\s*(?:kg|KG)\b", re.I), 1.0),  # approx for oils
]

# HDEO keywords suggest diesel engine oil
_HDEO_KEYWORDS = re.compile(
    r"\b(rimula|delo|rubia|tir|deo|diesel|commercial|fleet|hdeo|ci-4|ch-4|cj-4|cf-4)\b",
    re.I,
)
# MCO keywords
_MCO_KEYWORDS = re.compile(r"\b(4t|motorcycle|moto|blaze|scooter|two.wheel)\b", re.I)

# Brand detection — ordered by specificity (longer matches first)
_BRAND_PATTERNS = [
    (re.compile(r"\bshell\s+helix\s+ultra\b", re.I), "Shell Helix Ultra"),
    (re.compile(r"\bshell\s+helix\s+hx\s*8\b", re.I), "Shell Helix HX8"),
    (re.compile(r"\bshell\s+helix\s+hx\s*7\b", re.I), "Shell Helix HX7"),
    (re.compile(r"\bshell\s+helix\s+hx\s*5\b", re.I), "Shell Helix HX5"),
    (re.compile(r"\bshell\s+helix\s+hx\s*3\b", re.I), "Shell Helix HX3"),
    (re.compile(r"\bshell\s+rimula\s+r4\b", re.I), "Shell Rimula R4"),
    (re.compile(r"\bshell\s+rimula\s+r3\b", re.I), "Shell Rimula R3"),
    (re.compile(r"\bshell\s+rimula\b", re.I), "Shell Rimula"),
    (re.compile(r"\bshell\b", re.I), "Shell"),

    (re.compile(r"\bcaltex\s+havoline\s+ultra\b", re.I), "Caltex Havoline Ultra"),
    (re.compile(r"\bcaltex\s+havoline\s+pro.?ds\b", re.I), "Caltex Havoline Pro-DS"),
    (re.compile(r"\bcaltex\s+havoline\b", re.I), "Caltex Havoline"),
    (re.compile(r"\bcaltex\s+delo\s+400\b", re.I), "Caltex Delo 400"),
    (re.compile(r"\bcaltex\s+delo\s+300\b", re.I), "Caltex Delo 300"),
    (re.compile(r"\bcaltex\s+delo\b", re.I), "Caltex Delo"),
    (re.compile(r"\bcaltex\b", re.I), "Caltex"),

    (re.compile(r"\btotal\s+quartz\s+9000\b", re.I), "Total Quartz 9000"),
    (re.compile(r"\btotal\s+quartz\s+7000\b", re.I), "Total Quartz 7000"),
    (re.compile(r"\btotal\s+quartz\s+5000\b", re.I), "Total Quartz 5000"),
    (re.compile(r"\btotal\s+quartz\b", re.I), "Total Quartz"),
    (re.compile(r"\btotal\s+rubia\s+tir\b", re.I), "Total Rubia TIR"),
    (re.compile(r"\btotal\s+rubia\b", re.I), "Total Rubia"),
    (re.compile(r"\btotalenergies\b", re.I), "Total"),
    (re.compile(r"\btotal\b", re.I), "Total"),

    (re.compile(r"\baramco\b", re.I), "Aramco"),

    (re.compile(r"\bzic\s+x9\b", re.I), "ZIC X9"),
    (re.compile(r"\bzic\s+x7\b", re.I), "ZIC X7"),
    (re.compile(r"\bzic\s+x5\b", re.I), "ZIC X5"),
    (re.compile(r"\bzic\s+a\+\b", re.I), "ZIC A+"),
    (re.compile(r"\bzic\b", re.I), "ZIC"),

    (re.compile(r"\bkixx\s+pao\b", re.I), "Kixx PAO"),
    (re.compile(r"\bkixx\s+g1\b", re.I), "Kixx G1"),
    (re.compile(r"\bkixx\s+d1\b", re.I), "Kixx D1"),
    (re.compile(r"\bkixx\b", re.I), "Kixx"),

    (re.compile(r"\bequimoli\b", re.I), "Equimoli"),

    (re.compile(r"\bcarient\s+ultra\b", re.I), "PSO Carient Ultra"),
    (re.compile(r"\bcarient\s+fs\b", re.I), "PSO Carient FS"),
    (re.compile(r"\bcarient\s+plus\b", re.I), "PSO Carient Plus"),
    (re.compile(r"\bcarient\s+spro\b", re.I), "PSO Carient SPRO"),
    (re.compile(r"\bcarient\b", re.I), "PSO Carient"),
    (re.compile(r"\bpso\b", re.I), "PSO"),
]


def normalize(product: ScrapedProduct) -> ScrapedProduct:
    title = product.title or ""

    product.grade_detected = _detect_grade(title)
    product.pack_size_l = _detect_pack(title)
    product.brand_detected = _detect_brand(title)
    product.oil_type = _detect_type(title)

    if product.pack_size_l and product.pack_size_l > 0:
        product.price_per_litre = round(product.price / product.pack_size_l, 2)

    return product


def normalize_all(products: list[ScrapedProduct]) -> list[ScrapedProduct]:
    return [normalize(p) for p in products]


def _detect_grade(title: str) -> str | None:
    m = _GRADE_RE.search(title)
    if m:
        raw = m.group(1).upper()
        return raw.replace("W", "W-") if "W-" not in raw else raw
    return None


def _detect_pack(title: str) -> float | None:
    for pattern, multiplier in _PACK_PATTERNS:
        m = pattern.search(title)
        if m:
            val = float(m.group(1)) * multiplier
            if val > 0:
                return val
    return None


def _detect_brand(title: str) -> str | None:
    for pattern, brand in _BRAND_PATTERNS:
        if pattern.search(title):
            return brand
    return None


def _detect_type(title: str) -> str:
    if _MCO_KEYWORDS.search(title):
        return "mco"
    if _HDEO_KEYWORDS.search(title):
        return "hdeo"
    return "pcmo"  # default assumption for engine oils
