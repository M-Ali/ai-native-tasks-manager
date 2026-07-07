from .base import BaseScraper, ScrapedProduct
from .daraz import DarazScraper

PLATFORM_MAP = {
    "daraz": DarazScraper,
}

__all__ = ["BaseScraper", "ScrapedProduct", "DarazScraper", "PLATFORM_MAP"]
