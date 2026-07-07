"""SQLite persistence — stores every scrape run with timestamps."""
import sqlite3
from pathlib import Path
from datetime import datetime
from scrapers.base import ScrapedProduct

DB_PATH = Path(__file__).parent / "prices.db"


def init_db():
    con = sqlite3.connect(DB_PATH)
    con.execute("""
        CREATE TABLE IF NOT EXISTS scraped_products (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            platform        TEXT,
            brand_query     TEXT,
            title           TEXT,
            price           REAL,
            original_price  REAL,
            url             TEXT,
            seller          TEXT,
            rating          REAL,
            review_count    INTEGER,
            brand_detected  TEXT,
            grade_detected  TEXT,
            pack_size_l     REAL,
            price_per_litre REAL,
            oil_type        TEXT,
            scraped_at      TEXT
        )
    """)
    con.execute("""
        CREATE INDEX IF NOT EXISTS idx_brand_grade
        ON scraped_products (brand_detected, grade_detected, oil_type)
    """)
    con.commit()
    con.close()


def save(products: list[ScrapedProduct]):
    init_db()
    con = sqlite3.connect(DB_PATH)
    rows = [
        (
            p.platform, p.brand_query, p.title, p.price, p.original_price,
            p.url, p.seller, p.rating, p.review_count,
            p.brand_detected, p.grade_detected, p.pack_size_l,
            p.price_per_litre, p.oil_type,
            p.scraped_at.isoformat() if isinstance(p.scraped_at, datetime) else p.scraped_at
        )
        for p in products
    ]
    con.executemany("""
        INSERT INTO scraped_products
          (platform, brand_query, title, price, original_price, url, seller,
           rating, review_count, brand_detected, grade_detected, pack_size_l,
           price_per_litre, oil_type, scraped_at)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, rows)
    con.commit()
    con.close()


def latest_run() -> list[dict]:
    """Return all products from the most recent scrape date."""
    init_db()
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    cur = con.execute("""
        SELECT * FROM scraped_products
        WHERE DATE(scraped_at) = (
            SELECT DATE(MAX(scraped_at)) FROM scraped_products
        )
        ORDER BY brand_detected, grade_detected, pack_size_l
    """)
    rows = [dict(r) for r in cur.fetchall()]
    con.close()
    return rows


def all_runs_summary() -> list[dict]:
    """Return one row per (run_date, brand, grade) showing price range — for trend tracking."""
    init_db()
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    cur = con.execute("""
        SELECT
            DATE(scraped_at)    AS run_date,
            brand_detected,
            grade_detected,
            oil_type,
            pack_size_l,
            COUNT(*)            AS listing_count,
            MIN(price_per_litre) AS min_ppl,
            AVG(price_per_litre) AS avg_ppl,
            MAX(price_per_litre) AS max_ppl
        FROM scraped_products
        WHERE price_per_litre IS NOT NULL
          AND brand_detected IS NOT NULL
          AND grade_detected IS NOT NULL
        GROUP BY run_date, brand_detected, grade_detected, oil_type, pack_size_l
        ORDER BY run_date DESC, brand_detected, grade_detected
    """)
    rows = [dict(r) for r in cur.fetchall()]
    con.close()
    return rows
