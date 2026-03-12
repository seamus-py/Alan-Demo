# Alan-Demo
# Benefits Switch Predictor

A lead-scoring tool that identifies Canadian tech companies most likely to switch benefits providers. It pulls live signals from Indeed, Crunchbase, and Glassdoor, then outputs a colour-coded Excel dashboard ranked by switch likelihood.

---

## How It Works

Each company is scored out of 100 across five signals:

| Signal | Points | Source |
|---|---|---|
| Active HR / Benefits job posting | +30 | Indeed Canada |
| Funding round in last 12 months | +25 | Crunchbase |
| Glassdoor benefits rating < 3.5 | +20 | Glassdoor (via Google) |
| Company age 3–6 years | +10 | Seed data (first renewal window) |
| Headcount 50–300 employees | +5 | Seed data (Alan sweet spot) |

Scores are bucketed into three bands:

| Band | Score | Action |
|---|---|---|
| 🔴 HOT | 70–100 | Reach out immediately |
| 🟡 WARM | 40–69 | Monitor and nurture |
| 🟢 COLD | 0–39 | Low priority — revisit in 6 months |

---

## Installation

```bash
pip install requests beautifulsoup4 pandas openpyxl lxml
```

Python 3.9+ is recommended.

---

## Usage

```bash
python switch_predictor.py
```

The script will score all 25 seed companies, print a live summary to the console, and write two output files to the current directory:

- `switch_predictor_YYYYMMDD_HHMMSS.xlsx` — colour-coded Excel dashboard
- `switch_predictor_YYYYMMDD_HHMMSS.csv` — raw data export

**Note:** Crawl delays of 1.5–3 seconds per request are intentional and built in to avoid rate-limiting. A full run across 25 companies takes approximately 5–10 minutes.

---

## Output: Excel Dashboard

The workbook contains three sheets:

- **Lead Scores** — full results table, sorted HOT → WARM → COLD, with conditional colour formatting on the Score column, frozen header row, and auto-filter enabled.
- **Summary** — aggregated count and average score per band.
- **Signal Guide** — legend explaining each signal and scoring band.

---

## Customising the Company List

Edit the `COMPANIES` list at the top of `switch_predictor.py`. Each entry is a dictionary with the following fields:

```python
{
    "name":      "Company Name",
    "city":      "Toronto",
    "employees": 150,         # estimated headcount
    "founded":   2018,        # founding year
    "sector":    "FinTech",
    "website":   "example.com"
}
```

The `employees` and `founded` fields drive the static scoring signals (signals 4 and 5). The three live signals (Indeed, Crunchbase, Glassdoor) are fetched at runtime based on `name`.

---

## Limitations & Notes

- **Indeed scraping** relies on public job-listing HTML. Results may vary if Indeed changes its page structure.
- **Crunchbase scraping** targets public organisation pages only — no API key required, but data can be sparse for lesser-known companies.
- **Glassdoor ratings** are retrieved indirectly via Google search snippets and may not always surface a numeric rating; the script falls back to negative-keyword detection in that case.
- Crawling public websites may be subject to each site's Terms of Service. Use responsibly.
- For large company lists (50+), consider adding longer delays or batching runs to reduce the risk of IP-based rate limiting.

---

## Project Structure

```
switch_predictor.py   # main script (all logic in one file)
README.md             # this file
```

---

## License

Internal tool — not for public distribution.
