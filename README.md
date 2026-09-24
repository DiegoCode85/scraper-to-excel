# Catalogue scraper → Excel / CSV

Scrapes every product of a paginated catalogue and exports a clean spreadsheet.
Demo target: [books.toscrape.com](https://books.toscrape.com), a sandbox made for scraping practice
(1,000 books, 50 pages).

**What you get:** title, price, star rating, stock status and product URL for every item, in an
`.xlsx` with bold headers, frozen header row and filters — or a plain `.csv`.

```bash
pip install -r requirements.txt
python scraper.py --out books.xlsx           # full catalogue (~1 min)
python scraper.py --out books.csv --pages 3  # first 3 pages
```

## Built to be adapted
The site-specific part is two small functions:

| Function | Job |
|---|---|
| `parse_page(html, url)` | turn one listing page into rows |
| `next_page_url(html, url)` | find the "next" link (pagination) |

Everything else — polite delay between requests, retries with back-off, following pagination,
Excel/CSV export — stays the same for any listing site (products, real estate, job boards,
directories).

## Reliability
- retries each page up to 3 times with increasing wait
- 0.5 s pause between pages (configurable) to respect the target server
- tests on saved real pages: `pytest` (no network needed)

## Scope notes
Only public pages. Sites behind logins, CAPTCHAs or whose terms forbid automated access are out of
scope.
