"""
Brand and rank counter for Excel/CSV "Response" columns.

What it does
- Reads all sheets in the workbook and finds any column whose header contains "response" (case-insensitive).
- Parses numbered lists inside each cell like:
	1. Mazars - www.mazars.co.za
	2. Tax Consulting South Africa - www.taxconsulting.co.za
	...
- Extracts the brand name per item and counts:
	- total mentions per brand
	- how many times the brand appears at each rank (1, 2, 3, ...)
 Writes CSV summaries:
  - brand_counts.csv: LENIENT brand counts (every brand on the line is counted)
  - brand_counts_strict.csv: STRICT brand counts (brand + domain must match)
  - brand_metrics.csv and brand_metrics_strict.csv
  - brand_counts_by_role.csv/.strict.csv (same for Model and LLM Provider)
  - domain_counts.csv + grouped variants (counts by URL domain)
  - overall_rank_distribution.csv (+ _strict.csv)
  - parse_summary.csv
  - brand_url_audit_all.csv, brand_url_audit_strict.csv, brand_url_unmatched.csv

Notes on robustness
- Works across all sheets and any "response"-like column names (e.g., "Response", "AI Response", etc.).
- Accepts small format variations (1) Brand — URL, 1- Brand - URL, missing URL, different dash types, etc.
- Normalizes brand keys for grouping (lowercase, trim, collapse spaces, drop most punctuation) while preserving a readable display name.

Requires: pandas, openpyxl
"""

from __future__ import annotations

from pathlib import Path
from collections import Counter, defaultdict
from datetime import datetime
import json
import re
import sys

try:
	import pandas as pd
except ImportError:
	print("Missing dependency: pandas. Install with: pip install pandas openpyxl")
	sys.exit(1)


INPUT_NAME = "output_teams 1(in).csv"


def print_headers(path: Path) -> None:
	"""Utility to list sheet names and column headers for Excel or just headers for CSV."""
	suffix = path.suffix.lower()
	if suffix in {".xlsx", ".xlsm", ".xlsb", ".xls"}:
		xls = pd.ExcelFile(path)
		print(f"Workbook: {path}")
		for sheet in xls.sheet_names:
			try:
				df = pd.read_excel(path, sheet_name=sheet, nrows=0)
				cols = list(df.columns)
				print(f"\nSheet: {sheet}")
				if cols:
					for i, c in enumerate(cols, start=1):
						print(f"  {i}. {c}")
				else:
					print("  (No header row / zero columns)")
			except Exception as e:
				print(f"\nSheet: {sheet}")
				print(f"  Error reading columns: {e}")
	elif suffix == ".csv":
		try:
			df = pd.read_csv(path, nrows=0)
			cols = list(df.columns)
			print(f"CSV: {path}")
			if cols:
				for i, c in enumerate(cols, start=1):
					print(f"  {i}. {c}")
			else:
				print("  (No header row / zero columns)")
		except Exception as e:
			print(f"CSV: {path}")
			print(f"  Error reading columns: {e}")
	else:
		print(f"Unsupported file type for headers: {path}")


def normalize_key(name: str) -> str:
	"""Create a normalized key for grouping brand names.

	- Trim whitespace and surrounding punctuation
	- Collapse multiple spaces
	- Lowercase
	- Remove most punctuation except spaces and '&' and '+' (common in firm names)
	"""
	s = (name or "").strip()
	s = re.sub(r"\s+", " ", s)
	s = s.strip(" .-–—•*:_|")
	key = s.lower()
	key = re.sub(r"[^a-z0-9&+ ]+", "", key)
	key = re.sub(r"\s+", " ", key).strip()
	return key


_URL_RE = re.compile(r"(https?://\S+|www\.\S+|\S+\.(?:com|co\.za|net|org|global|za|uk|au)(?:/\S*)?)", re.IGNORECASE)


def extract_brand_and_url(segment: str) -> tuple[str, str | None]:
	"""Extract (brand, url) from a segment after the rank number.

	Handles formats like:
	- "Brand - www.brand.com"
	- "Brand — https://brand.co.za/path"
	- "Brand www.brand.com"
	- "Brand" (url=None)
	"""
	if not segment:
		return "", None
	s = segment.strip()

	# Try dash-separated first: split once
	for sep in (" - ", " – ", " — "):
		if sep in s:
			left, right = s.split(sep, 1)
			left = left.strip()
			right = right.strip()
			m = _URL_RE.search(right)
			url = m.group(0) if m else None
			if url:
				url = url.strip().strip("()[]{}")
			brand = left
			return brand.strip(), (url.strip() if url else None)

	# Otherwise search for URL anywhere
	m = _URL_RE.search(s)
	url = m.group(0) if m else None
	if m:
		brand = s[: m.start()].rstrip(" -–—").strip()
	else:
		brand = s.strip()

	# Clean trailing separators in brand
	brand = re.sub(r"\s+[–—-]\s*$", "", brand).strip()
	return brand, (url.strip().strip("()[]{}") if url else None)


def extract_all_urls(text: str) -> list[str]:
	"""Find all URL-like substrings in the text and return cleaned URLs (no surrounding ()[]{})."""
	if not isinstance(text, str):
		return []
	urls = []
	for m in _URL_RE.finditer(text):
		u = m.group(0).strip().strip("()[]{}")
		urls.append(u)
	return urls


_NUMBERED_ITEM_RE = re.compile(r"(?m)^\s*(\d+)\s*[\.\)\-:]\s*(.+)$")


def parse_response_items(text: str) -> list[dict]:
	"""Parse a response cell into items with rank, brand, and url.

	Returns a list of dicts: {rank: int, brand: str, url: Optional[str]}
	"""
	if not isinstance(text, str):
		return []
	raw = text.strip()
	if not raw:
		return []

	items: list[dict] = []
	for m in _NUMBERED_ITEM_RE.finditer(raw):
		try:
			rank = int(m.group(1))
		except Exception:
			continue
		rest = m.group(2).strip()
		brand, url = extract_brand_and_url(rest)
		urls = extract_all_urls(rest)
		if brand:
			items.append({"rank": rank, "brand": brand, "url": url, "urls": urls})

	if items:
		return items

	# Fallback: split into non-empty lines and treat order as rank
	lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
	for i, ln in enumerate(lines, start=1):
		ln = re.sub(r"^\s*\d+\s*[\.\)\-:]\s*", "", ln)  # leading numbers
		ln = re.sub(r"^[\-\*•]+\s*", "", ln)  # bullets
		brand, url = extract_brand_and_url(ln)
		urls = extract_all_urls(ln)
		if brand:
			items.append({"rank": i, "brand": brand, "url": url, "urls": urls})
	return items


def _normalize_domain(url: str | None) -> tuple[str | None, str | None]:
	"""Return (domain_full, domain_token) from a URL or domain-like string.

	- Strips scheme and common subdomain prefixes (www, www2, home, m)
	- Handles multi-label ccTLDs like co.za, co.uk where possible
	- domain_token is the key used for brand-domain matching
	"""
	if not url:
		return None, None
	s = url.strip()
	# Add scheme if missing for urlparse
	if not re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*://", s):
		s_for_parse = "http://" + s
	else:
		s_for_parse = s
	try:
		from urllib.parse import urlparse
		p = urlparse(s_for_parse)
		netloc = p.netloc or s
	except Exception:
		netloc = s
	host = netloc.lower()
	# Trim credentials and port if present
	host = re.sub(r"^.*@", "", host)
	host = host.split(":")[0]
	# Remove common prefixes repeatedly
	for pref in ("www.", "www2.", "home.", "m."):
		if host.startswith(pref):
			host = host[len(pref):]
	labels = [lbl for lbl in host.split(".") if lbl]
	if not labels:
		return None, None
	# Handle multi-label ccTLDs
	multi_suffixes = {("co", "za"), ("co", "uk"), ("com", "au")}
	domain_full = None
	domain_token = None
	if len(labels) >= 3 and (labels[-2], labels[-1]) in multi_suffixes:
		domain_full = ".".join(labels[-3:])
		domain_token = labels[-3]
	elif len(labels) >= 2 and labels[-1] in {"com", "org", "net", "global", "za", "uk", "au"}:
		domain_full = ".".join(labels[-2:])
		domain_token = labels[-2]
	else:
		# Brand TLD or single label
		domain_full = labels[-1]
		domain_token = labels[-1]
	# Normalize hyphens in token for matching convenience
	return domain_full, domain_token


def _brand_matches_domain(brand: str, domain_token: str | None) -> bool:
	if not brand or not domain_token:
		return False
	bk = normalize_key(brand)
	domain_token_norm = re.sub(r"[^a-z0-9]", "", domain_token.lower())
	if not domain_token_norm:
		return False
	# Token-based match
	tokens = re.findall(r"[a-z0-9]+", bk)
	if domain_token_norm in tokens:
		return True
	# Compact string contains
	compact = "".join(tokens)
	compact_dt = domain_token_norm.replace("-", "")
	if compact_dt and (compact_dt in compact or compact in compact_dt):
		return True
	return False


def find_response_columns(df: pd.DataFrame) -> list[str]:
	"""Return a list of column names that likely contain responses."""
	cols = []
	for c in df.columns:
		name = str(c)
		if "response" in name.lower():
			cols.append(name)
	return cols


def _get_datasets(path: Path) -> list[pd.DataFrame]:
	"""Load input into one or more DataFrames (sheets for Excel, single for CSV)."""
	suffix = path.suffix.lower()
	if suffix in {".xlsx", ".xlsm", ".xlsb", ".xls"}:
		xls = pd.ExcelFile(path)
		frames: list[pd.DataFrame] = []
		for sheet in xls.sheet_names:
			try:
				df = pd.read_excel(path, sheet_name=sheet)
				frames.append(df)
			except Exception as e:
				print(f"Skipping sheet '{sheet}' due to read error: {e}")
		return frames
	elif suffix == ".csv":
		try:
			df = pd.read_csv(path)
		except UnicodeDecodeError:
			for enc in ("utf-8-sig", "latin1"):
				try:
					df = pd.read_csv(path, encoding=enc)
					break
				except Exception:
					df = None  # type: ignore
			if df is None:  # type: ignore
				raise
		return [df]
	else:
		raise ValueError(f"Unsupported file type: {path}")


def count_brands_and_ranks(path: Path) -> pd.DataFrame:
	"""Read Excel/CSV and return strict counts (only when brand matches domain of URL)."""
	datasets = _get_datasets(path)
	records, _ = _records_from_datasets(datasets, ["Role", "Model", "LLM Provider"])
	# Filter to strict matches
	strict_records = [r for r in records if r.get("match_strict") is True]
	df, _ = _aggregate_brand_counts(strict_records)
	return df


def _records_from_datasets(datasets: list[pd.DataFrame], group_fields: list[str]) -> tuple[list[dict], dict]:
	"""Parse datasets into record dicts: {brand, rank, <groups...>} and a summary dict."""
	records: list[dict] = []
	rows_seen = 0
	rows_with_items = 0
	items_count = 0

	for df in datasets:
		if df is None or df.empty:
			continue
		response_cols = find_response_columns(df)
		if not response_cols:
			continue
		# Capture available group columns for this df
		df_cols_lower = {c.lower(): c for c in df.columns}
		mapped_groups = {g: df_cols_lower.get(g.lower()) for g in group_fields}

		for _, row in df.iterrows():
			rows_seen += 1
			group_vals = {g: (row[mapped_groups[g]] if mapped_groups.get(g) in row else None) for g in group_fields}
			row_items = 0
			for col in response_cols:
				val = row.get(col)
				if pd.isna(val):
					continue
				items = parse_response_items(val)
				for it in items:
					brand = it["brand"]
					rank = int(it["rank"])
					url = it.get("url")
					urls = it.get("urls") or ([] if url is None else [url])
					domain_full, domain_token = _normalize_domain(url)
					match_strict = _brand_matches_domain(brand, domain_token)
					records.append({
						"brand": brand,
						"rank": rank,
						"url": url,
						"urls": urls,
						"domain_full": domain_full,
						"domain_token": domain_token,
						"match_strict": match_strict,
						**group_vals
					})
					row_items += 1
			if row_items > 0:
				rows_with_items += 1
				items_count += row_items

	summary = {
		"rows_seen": rows_seen,
		"rows_with_items": rows_with_items,
		"items_count": items_count,
	}
	return records, summary


def _aggregate_brand_counts(records: list[dict]) -> tuple[pd.DataFrame, int]:
	brand_counts: dict[str, dict[str, object]] = {}
	display_name: dict[str, str] = {}
	max_rank = 0
	for rec in records:
		brand = rec["brand"]
		rank = int(rec["rank"])
		key = normalize_key(brand)
		if not key:
			continue
		if key not in brand_counts:
			brand_counts[key] = {"total": 0, "ranks": Counter()}
			display_name[key] = str(brand).strip()
		brand_counts[key]["total"] = int(brand_counts[key]["total"]) + 1  # type: ignore
		brand_counts[key]["ranks"][rank] += 1  # type: ignore
		if rank > max_rank:
			max_rank = rank

	rows = []
	rank_cols = [f"rank_{i}" for i in range(1, max_rank + 1)]
	for key, data in brand_counts.items():
		ranks: Counter = data["ranks"]  # type: ignore
		row = {
			"brand": display_name.get(key, key),
			"total": int(data["total"])  # type: ignore
		}
		for i in range(1, max_rank + 1):
			row[f"rank_{i}"] = int(ranks.get(i, 0))
		rows.append(row)
	if not rows:
		return pd.DataFrame(columns=["brand", "total"]), 0
	out_df = pd.DataFrame(rows)
	out_df.sort_values(by=["total", "brand"], ascending=[False, True], inplace=True)
	out_df.reset_index(drop=True, inplace=True)
	out_df = out_df[["brand", "total", *rank_cols]]
	return out_df, max_rank


def _aggregate_domain_counts(records: list[dict]) -> tuple[pd.DataFrame, int]:
	"""Aggregate counts by domain_full across all URLs present in each record."""
	dom_counts: dict[str, dict[str, object]] = {}
	max_rank = 0
	for rec in records:
		rank = int(rec["rank"])
		urls = rec.get("urls") or []
		# consider every URL on the line
		for u in urls:
			domain_full, _ = _normalize_domain(u)
			if not domain_full:
				continue
			key = domain_full.lower()
			if key not in dom_counts:
				dom_counts[key] = {"total": 0, "ranks": Counter()}
			dom_counts[key]["total"] = int(dom_counts[key]["total"]) + 1  # type: ignore
			dom_counts[key]["ranks"][rank] += 1  # type: ignore
			if rank > max_rank:
				max_rank = rank

	rows = []
	rank_cols = [f"rank_{i}" for i in range(1, max_rank + 1)]
	for key, data in dom_counts.items():
		ranks: Counter = data["ranks"]  # type: ignore
		row = {
			"domain": key,
			"total": int(data["total"])  # type: ignore
		}
		for i in range(1, max_rank + 1):
			row[f"rank_{i}"] = int(ranks.get(i, 0))
		rows.append(row)
	if not rows:
		return pd.DataFrame(columns=["domain", "total"]), 0
	out_df = pd.DataFrame(rows)
	out_df.sort_values(by=["total", "domain"], ascending=[False, True], inplace=True)
	out_df.reset_index(drop=True, inplace=True)
	out_df = out_df[["domain", "total", *rank_cols]]
	return out_df, max_rank


def _compute_metrics(df: pd.DataFrame) -> pd.DataFrame:
	if df.empty:
		return df
	rank_cols = sorted([c for c in df.columns if c.startswith("rank_")], key=lambda x: int(x.split("_")[1]))
	# precompute arrays
	def col_rank(c):
		return int(c.split("_")[1])

	totals = df["total"].astype(int).clip(lower=1)  # avoid div by zero
	counts = {c: df[c].astype(int) for c in rank_cols}

	# topN counts
	def sum_cols(n):
		cols = [c for c in rank_cols if col_rank(c) <= n]
		return sum((counts[c] for c in cols), start=df[rank_cols[0]].astype(int) * 0)

	dfm = pd.DataFrame({
		"brand": df["brand"],
		"total": df["total"].astype(int),
	})
	for n in (1, 3, 5, 10):
		topn = sum_cols(n)
		dfm[f"top{n}_count"] = topn.astype(int)
		dfm[f"top{n}_share"] = (topn / totals).round(4)

	# mean rank and moments
	weighted_sum = 0
	weighted_sq = 0
	for c in rank_cols:
		r = col_rank(c)
		cnt = counts[c]
		weighted_sum = weighted_sum + cnt * r
		weighted_sq = weighted_sq + cnt * (r * r)
	mean_rank = (weighted_sum / totals).round(4)
	dfm["mean_rank"] = mean_rank
	dfm["rank_std"] = ((weighted_sq / totals - mean_rank ** 2).clip(lower=0).pow(0.5)).round(4)

	# best/worst/median/predominant
	def best_rank_row(row):
		for c in rank_cols:
			if int(row[c]) > 0:
				return col_rank(c)
		return None

	def worst_rank_row(row):
		for c in reversed(rank_cols):
			if int(row[c]) > 0:
				return col_rank(c)
		return None

	def median_rank_row(row):
		total = int(row["total"]) if int(row["total"]) > 0 else 1
		half = (total + 1) // 2
		cum = 0
		for c in rank_cols:
			cum += int(row[c])
			if cum >= half:
				return col_rank(c)
		return None

	def predominant_rank_row(row):
		best_c = None
		best_v = -1
		for c in rank_cols:
			v = int(row[c])
			if v > best_v:
				best_v = v
				best_c = c
		return col_rank(best_c) if best_c else None

	dfm["best_rank"] = df.apply(best_rank_row, axis=1)
	dfm["worst_rank"] = df.apply(worst_rank_row, axis=1)
	dfm["median_rank"] = df.apply(median_rank_row, axis=1)
	dfm["predominant_rank"] = df.apply(predominant_rank_row, axis=1)

	return dfm


def _write_csv(path: Path, df: pd.DataFrame, name: str) -> None:
	try:
		df.to_csv(path, index=False)
		print(f"Saved {name} -> {path}")
	except Exception as e:
		print(f"Failed to write {name}: {e}")


def _append_log(msg: str) -> None:
	ts = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
	try:
		with open("run.log", "a", encoding="utf-8") as f:
			f.write(f"[{ts}] {msg}\n")
	except Exception:
		pass


def _write_summary_json(summary: dict) -> None:
	try:
		with open("results_summary.json", "w", encoding="utf-8") as f:
			json.dump(summary, f, indent=2, ensure_ascii=False)
		print("Saved results summary -> results_summary.json")
	except Exception as e:
		print(f"Failed to write results summary: {e}")


def _overall_rank_distribution(df: pd.DataFrame) -> pd.DataFrame:
	if df.empty:
		return pd.DataFrame(columns=["rank", "count", "share"])
	rank_cols = sorted([c for c in df.columns if c.startswith("rank_")], key=lambda x: int(x.split("_")[1]))
	counts = []
	total_mentions = int(df["total"].sum())
	for c in rank_cols:
		r = int(c.split("_")[1])
		cnt = int(df[c].sum())
		share = (cnt / total_mentions) if total_mentions else 0.0
		counts.append({"rank": r, "count": cnt, "share": round(share, 4)})
	return pd.DataFrame(counts)


def main() -> int:
	# Allow optional CLI args: [input_file] [--headers]
	import argparse

	parser = argparse.ArgumentParser(description="Count brand mentions and rank positions from response columns in an Excel workbook (all sheets) or a CSV file.")
	parser.add_argument("input_file", nargs="?", default=INPUT_NAME, help="Path to the input file (.xlsx/.xls/.csv)")
	parser.add_argument("--headers", action="store_true", help="Only list sheet headers (Excel) or CSV headers and exit")
	parser.add_argument("--output", "-o", default="brand_counts.csv", help="Output CSV filename")
	args = parser.parse_args()

	in_path = Path(args.input_file)
	if not in_path.exists():
		print(f"File not found: {in_path}")
		return 2

	if args.headers:
		print_headers(in_path)
		return 0

	print(f"Reading: {in_path}")
	try:
		# Original aggregate (strict counts)
		df = count_brands_and_ranks(in_path)

		# New: build records and compute grouped outputs/metrics
		datasets = _get_datasets(in_path)
		group_fields = ["Role", "Model", "LLM Provider"]
		records, parse_summary = _records_from_datasets(datasets, group_fields)
		# Strict-filtered records (brand-domain matched)
		strict_records = [r for r in records if r.get("match_strict") is True]

		# Overall strict counts for metrics and distributions
		df_overall_strict, _ = _aggregate_brand_counts(strict_records)
		metrics = _compute_metrics(df_overall_strict)

		# Grouped counts
		grouped_outputs = {}
		for g in group_fields:
			# pivot per group value
			sub_records = {}
			for r in strict_records:
				gval = r.get(g)
				sub_records.setdefault(gval, []).append(r)
			frames = []
			for gval, rs in sub_records.items():
				sub_df, _ = _aggregate_brand_counts(rs)
				if sub_df.empty:
					continue
				sub_df.insert(0, g, gval)
				frames.append(sub_df)
			grouped_outputs[g] = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

		# Rank distribution (overall strict)
		rank_dist = _overall_rank_distribution(df_overall_strict)

	except Exception as e:
		print(f"Error while counting brands: {e}")
		return 1

	if df.empty:
		print("No brand data found in any 'response' column.")
		return 0

	# Write outputs
	out_path = Path(args.output)
	try:
		# df is strict counts per current count_brands_and_ranks
		df.to_csv(Path("brand_counts_strict.csv"), index=False)
		print(f"\nTop brands (first 20):")
		print(df.head(20).to_string(index=False))
		print(f"\nSaved strict summary to: {Path('brand_counts_strict.csv').resolve()}")

		# Metrics and grouped outputs (strict)
		_write_csv(Path("brand_metrics_strict.csv"), metrics, "brand metrics (strict)")
		_write_csv(Path("brand_counts_by_role.csv"), grouped_outputs.get("Role", pd.DataFrame()), "brand counts by Role")
		_write_csv(Path("brand_counts_by_model.csv"), grouped_outputs.get("Model", pd.DataFrame()), "brand counts by Model")
		_write_csv(Path("brand_counts_by_provider.csv"), grouped_outputs.get("LLM Provider", pd.DataFrame()), "brand counts by LLM Provider")
		_write_csv(Path("overall_rank_distribution_strict.csv"), rank_dist, "overall rank distribution (strict)")

		# Parse summary
		ps = pd.DataFrame([parse_summary])
		_write_csv(Path("parse_summary.csv"), ps, "parse summary")

		# Audits for transparency
		audit_cols = [
			"brand", "rank", "url", "domain_full", "domain_token", "match_strict",
			"Role", "Model", "LLM Provider"
		]
		audit_df = pd.DataFrame(records)[audit_cols] if records else pd.DataFrame(columns=audit_cols)
		strict_audit_df = pd.DataFrame(strict_records)[audit_cols] if strict_records else pd.DataFrame(columns=audit_cols)
		unmatched_df = audit_df[audit_df["match_strict"] != True] if not audit_df.empty else audit_df
		_write_csv(Path("brand_url_audit_all.csv"), audit_df, "brand/url audit (all)")
		_write_csv(Path("brand_url_audit_strict.csv"), strict_audit_df, "brand/url audit (strict matches)")
		_write_csv(Path("brand_url_unmatched.csv"), unmatched_df, "brand/url unmatched")
		# Lenient brand counts (count every brand on the line)
		# Re-aggregate without the strict filter
		all_df, _ = _aggregate_brand_counts(records)
		_write_csv(Path("brand_counts.csv"), all_df, "brand counts (lenient)")
		all_metrics = _compute_metrics(all_df)
		_write_csv(Path("brand_metrics.csv"), all_metrics, "brand metrics (lenient)")

		# Domain counts from all URLs present on each line
		dom_df, _ = _aggregate_domain_counts(records)
		_write_csv(Path("domain_counts.csv"), dom_df, "domain counts")

		# Write a simple summary JSON and append run log
		summary = {
			"input": str(in_path),
			"rows_seen": parse_summary.get("rows_seen"),
			"rows_with_items": parse_summary.get("rows_with_items"),
			"items_count": parse_summary.get("items_count"),
			"outputs": [
				"brand_counts.csv",
				"brand_counts_strict.csv",
				"brand_metrics.csv",
				"brand_metrics_strict.csv",
				"brand_counts_by_role.csv",
				"brand_counts_by_model.csv",
				"brand_counts_by_provider.csv",
				"overall_rank_distribution_strict.csv",
				"domain_counts.csv",
				"parse_summary.csv",
				"brand_url_audit_all.csv",
				"brand_url_audit_strict.csv",
				"brand_url_unmatched.csv",
			],
		}
		_write_summary_json(summary)
		_append_log(f"Processed {in_path} with {summary['rows_with_items']} rows having items; outputs saved.")
		return 0
	except Exception as e:
		print(f"Failed to write output CSVs: {e}")
		return 1


if __name__ == "__main__":
	sys.exit(main())

