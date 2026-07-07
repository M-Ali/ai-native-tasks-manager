"""Abstract base scraper — extend this to add any new platform."""
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional


@dataclass
class ScrapedProduct:
    platform: str
    brand_query: str          # search term used
    title: str                # raw product title from site
    price: float              # PKR (or local currency)
    original_price: Optional[float]  # before discount, if shown
    url: str
    seller: Optional[str]
    rating: Optional[float]
    review_count: Optional[int]
    scraped_at: datetime = field(default_factory=datetime.utcnow)

    # Populated by normalizer
    brand_detected: Optional[str] = None
    grade_detected: Optional[str] = None
    pack_size_l: Optional[float] = None
    price_per_litre: Optional[float] = None
    oil_type: Optional[str] = None      # pcmo / hdeo / mco


class BaseScraper(ABC):
    def __init__(self, config: dict):
        self.config = config
        self.platform_cfg = config.get("platforms", {}).get(self.platform_name, {})

    @property
    @abstractmethod
    def platform_name(self) -> str:
        """Return the platform key matching config/lubricants.yaml platforms section."""

    @abstractmethod
    def scrape(self, search_term: str) -> list[ScrapedProduct]:
        """Scrape products for a given search term. Returns list of ScrapedProduct."""

    def scrape_all(self, search_terms: list[str]) -> list[ScrapedProduct]:
        results = []
        for term in search_terms:
            results.extend(self.scrape(term))
        return results
