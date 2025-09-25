# brandCount

Tools to parse AI-generated ranked responses, extract brand names and URLs, and compute robust counts and metrics from CSV or Excel data.

## What this does

- Reads input from:
  - CSV file (default: `output_teams 1(in).csv`)
  - Excel workbook (`.xlsx/.xlsm/.xlsb/.xls`, all sheets)
- Detects any column with the word `Response` (case-insensitive)
- Parses ranked lists in each cell (1., 2), 3- etc.), extracting:
  - Brand name
  - URL(s) on the line
  - Rank index
- Produces two counting modes:
  - Lenient: counts every brand on the line, regardless of URL
  - Strict: counts only when the brand matches the URL’s domain token
- Outputs multiple CSV summaries and audits

## Why strict matching?
Generative responses often include multiple brands or mix unrelated URLs. Counting a brand only when its name matches the URL’s domain (`kpmg` in `home.kpmg/za/...`) dramatically improves precision by eliminating mismatches.

## Key outputs

- `brand_counts.csv` — Lenient brand totals with rank buckets (rank_1..k)
- `brand_counts_strict.csv` — Strict brand totals with rank buckets
- `brand_metrics.csv` — Lenient metrics per brand
- `brand_metrics_strict.csv` — Strict metrics per brand
- `brand_counts_by_role.csv` — Strict counts grouped by `Role`
- `brand_counts_by_model.csv` — Strict counts grouped by `Model`
- `brand_counts_by_provider.csv` — Strict counts grouped by `LLM Provider`
- `overall_rank_distribution_strict.csv` — Strict rank distribution across all items
- `domain_counts.csv` — Totals by URL domain across all URLs found on each line
- `parse_summary.csv` — Rows seen, rows with items, item count
- `brand_url_audit_all.csv` — All parsed items with brand, url, domain and match flag
- `brand_url_audit_strict.csv` — Only items that were counted (strict)
- `brand_url_unmatched.csv` — Items not counted (strict), for inspection
- `results_summary.json` — Quick machine‑readable run summary
- `run.log` — Timestamped log entries per run

## Metrics per brand
For each brand we compute (for both lenient and strict):
- top1_count/share, top3_count/share, top5_count/share, top10_count/share
- mean_rank, rank_std
- best_rank, worst_rank, median_rank, predominant_rank

## Matching details
- URL normalization: remove scheme, credentials, port, and common prefixes (www., www2., home., m.)
- Handles ccTLDs: `co.za`, `co.uk`, `com.au`, and common TLDs (`.com`, `.org`, `.net`, `.global`)
- Domain token: primary label, e.g., `kpmg` from `home.kpmg/za/en/home.html`, `pwc` from `www.pwc.co.za`
- Brand key normalization: lowercased alphanumerics with spaces, preserving `&` and `+`
- Strict match rule: brand tokens must contain the domain token or its compact variant

## How to run

- Show headers (CSV or Excel):
  ```
  /home/codespace/.python/current/bin/python /workspaces/brandCount/brand_counter.py --headers
  ```
- Process default input (CSV):
  ```
  /home/codespace/.python/current/bin/python /workspaces/brandCount/brand_counter.py
  ```
- Process a specific file and set strict output name:
  ```
  /home/codespace/.python/current/bin/python /workspaces/brandCount/brand_counter.py "/path/to/input.xlsx" -o brand_counts_strict.csv
  ```

## Customization
- Brand/domain mapping: If you have known exceptions (e.g., group-level sites), share a mapping and we’ll incorporate it to force a strict match.
- Filters: We can add flags to limit outputs by `Role`, `Model`, or `LLM Provider`.
- Lenient vs strict side-by-side: We already output both; we can also add a combined comparison if helpful.

## File structure
- `brand_counter.py` — Main logic
- `requirements.txt` — Python dependencies
- Outputs are written to repo root by default

## Troubleshooting
- If counts are unexpectedly low, inspect `brand_url_unmatched.csv` to see which items failed strict matching.
- For unexpected parsing, check `brand_url_audit_all.csv` to verify extracted brand names and URLs.
- If your data uses different headers, ensure response columns contain “Response”.

## License
Internal analytics tooling. Adjust as needed for your use case.
