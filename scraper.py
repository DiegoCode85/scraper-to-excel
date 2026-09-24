"""Scrape a paginated product catalogue into Excel or CSV.

Demo target: https://books.toscrape.com (a sandbox built for scraping practice).
The same structure — fetch, parse, follow "next", export — adapts to most
listing sites in an hour: only `parse_page` and `next_page_url` change.

Usage:
    python scraper.py --out books.xlsx            # whole catalogue
    python scraper.py --out books.csv --pages 3   # first 3 pages only
"""
from __future__ import annotations

import argparse
import csv
import logging
import time
from dataclasses import astuple, dataclass
from pathlib import Path
from urllib.parse import urljoin

import httpx
from bs4 import BeautifulSoup

START_URL = "https://books.toscrape.com/catalogue/page-1.html"
HEADERS = ["title", "price_gbp", "rating", "in_stock", "url"]
_RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

log = logging.getLogger("scraper")


@dataclass
class Book:
    title: str
    price: float
    rating: int
    in_stock: bool
    url: str


def parse_page(html: str, page_url: str) -> list[Book]:
    soup = BeautifulSoup(html, "html.parser")
    books = []
    for pod in soup.select("article.product_pod"):
        link = pod.select_one("h3 a")
        price_text = pod.select_one(".price_color").get_text(strip=True)
        rating_cls = next((c for c in pod.select_one(".star-rating")["class"] if c in _RATINGS), None)
        books.append(Book(
            title=link["title"],
            price=float(price_text.lstrip("£Â")),
            rating=_RATINGS.get(rating_cls, 0),
            in_stock="In stock" in pod.select_one(".availability").get_text(),
            url=urljoin(page_url, link["href"]),
        ))
    return books


def next_page_url(html: str, page_url: str) -> str | None:
    nxt = BeautifulSoup(html, "html.parser").select_one("li.next a")
    return urljoin(page_url, nxt["href"]) if nxt else None


def fetch(client: httpx.Client, url: str, retries: int = 3) -> str:
    for attempt in range(1, retries + 1):
        try:
            r = client.get(url)
            r.raise_for_status()
            return r.text
        except httpx.HTTPError as exc:
            if attempt == retries:
                raise
            log.warning("retry %d for %s (%s)", attempt, url, exc)
            time.sleep(2 * attempt)
    raise RuntimeError("unreachable")


def scrape(start_url: str = START_URL, max_pages: int | None = None, delay_s: float = 0.5) -> list[Book]:
    books: list[Book] = []
    url, pages = start_url, 0
    with httpx.Client(headers={"User-Agent": "catalogue-scraper/1.0"}, timeout=20, follow_redirects=True) as client:
        while url and (max_pages is None or pages < max_pages):
            html = fetch(client, url)
            books.extend(parse_page(html, url))
            pages += 1
            log.info("page %d: %d books so far", pages, len(books))
            url = next_page_url(html, url)
            time.sleep(delay_s)          # be polite to the server
    return books


def write_output(books: list[Book], path: Path) -> None:
    path = Path(path)
    if path.suffix.lower() == ".xlsx":
        import openpyxl
        from openpyxl.styles import Font

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "books"
        ws.append(HEADERS)
        for c in ws[1]:
            c.font = Font(bold=True)
        for b in books:
            ws.append(list(astuple(b)))
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"
        wb.save(path)
    else:
        with path.open("w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(HEADERS)
            w.writerows(astuple(b) for b in books)


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--out", default="books.xlsx", help="output file (.xlsx or .csv)")
    ap.add_argument("--pages", type=int, default=None, help="max pages (default: all)")
    ap.add_argument("--start", default=START_URL)
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
    books = scrape(args.start, args.pages)
    write_output(books, Path(args.out))
    log.info("saved %d rows to %s", len(books), args.out)


if __name__ == "__main__":
    main()
