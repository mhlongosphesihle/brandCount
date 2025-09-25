"""
Brand and rank counter for Excel "response" columns.

What it does
- Reads all sheets in the workbook and finds any column whose header contains "response" (case-insensitive).
- Parses numbered lists inside each cell like:
	1. Mazars - www.mazars.co.za
	2. Tax Consulting South Africa - www.taxconsulting.co.za
	...
- Extracts the brand name per item and counts:
	- total mentions per brand
	- how many times the brand appears at each rank (1, 2, 3, ...)
- Writes a CSV summary: brand_counts.csv with columns [brand, total, rank_1, rank_2, ...]
- Also writes grouped summaries and metrics:
	- brand_metrics.csv: adds mean/median/best/worst rank, std dev, top1/top3/top5/top10 counts and shares
	- brand_counts_by_role.csv: per Role breakdown
	- brand_counts_by_model.csv: per Model breakdown
	- brand_counts_by_provider.csv: per LLM Provider breakdown
	- overall_rank_distribution.csv: how many mentions occur at each rank overall
	- parse_summary.csv: quick stats about parsed rows/items

Notes on robustness
- Works across all sheets and any "response"-like column names (e.g., "Response", "AI Response", etc.).
- Accepts small format variations (1) Brand — URL, 1- Brand - URL, missing URL, different dash types, etc.
- Normalizes brand keys for grouping (lowercase, trim, collapse spaces, drop most punctuation) while preserving a readable display name.

Requires: pandas, openpyxl
"""

from __future__ import annotations

from pathlib import Path
from collections import Counter, defaultdict
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


_URL_RE = re.compile(r"(https?://\S+|www\.\S+|\S+\.(?:com|co\.za|net|org|global)(?:/\S*)?)", re.IGNORECASE)


def extract_brand(segment: str) -> str:
	"""Given the text after the rank number, extract the brand portion before a URL or trailing dash.

	Examples:
	- "Mazars - www.mazars.co.za" -> "Mazars"
	- "PWC South Africa www.pwc.co.za" -> "PWC South Africa"
	- "EY South Africa — https://ey.com/en_za" -> "EY South Africa"
	- "BDO South Africa" -> "BDO South Africa"
	"""
	if not segment:
		return ""
	s = segment.strip()

	# If we have a dash-separated brand — URL, cut at the dash first
	dash_idx = None
	for sep in (" - ", " – ", " — "):
		if sep in s:
			dash_idx = s.index(sep)
			break
	if dash_idx is not None:
		s = s[:dash_idx]
	else:
		# Otherwise, cut right before a URL-like token
		m = _URL_RE.search(s)
		if m:
			s = s[: m.start()].rstrip(" -–—")

	# Clean any trailing separators
	s = re.sub(r"\s+[–—-]\s*$", "", s)
	return s.strip()


_NUMBERED_ITEM_RE = re.compile(r"(?m)^\s*(\d+)\s*[\.\)\-:]\s*(.+)$")


def parse_response_cell(text: str) -> list[tuple[int, str]]:
	"""Parse a response cell into a list of (rank, brand) pairs.

	Strategy:
	1) Prefer numbered lines like "1. Brand - URL" across the cell (multi-line supported)
	2) If none found, split by lines and assign implicit ranks 1..n
	"""
	if not isinstance(text, str):
		return []
	raw = text.strip()
	if not raw:
		return []

	items: list[tuple[int, str]] = []
	for m in _NUMBERED_ITEM_RE.finditer(raw):
		try:
			rank = int(m.group(1))
		except Exception:
			continue
		rest = m.group(2).strip()
		brand = extract_brand(rest)
		if brand:
			items.append((rank, brand))

	if items:
		return items

	# Fallback: split into non-empty lines and treat order as rank
	lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
	for i, ln in enumerate(lines, start=1):
		# remove any leading numbering or bullets
		ln = re.sub(r"^\s*\d+\s*[\.\)\-:]\s*", "", ln)
		ln = re.sub(r"^[\-\*•]+\s*", "", ln)
		brand = extract_brand(ln)
		if brand:
			items.append((i, brand))
	return items


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
	"""Read an Excel workbook (all sheets) or a CSV and produce a DataFrame with brand totals and rank counts.

	Returns columns: [brand, total, rank_1, rank_2, ..., rank_k]
	"""
	brand_counts: dict[str, dict[str, object]] = {}
	display_name: dict[str, str] = {}
	max_rank = 0

	datasets = _get_datasets(path)

	for df in datasets:
		if df is None or df.empty:
			continue

		response_cols = find_response_columns(df)
		if not response_cols:
			continue

		for col in response_cols:
			for val in df[col].dropna().tolist():
				items = parse_response_cell(val)
				for rank, brand in items:
					key = normalize_key(brand)
					if not key:
						continue
					if key not in brand_counts:
						brand_counts[key] = {"total": 0, "ranks": Counter()}
						display_name[key] = brand.strip()
					brand_counts[key]["total"] = int(brand_counts[key]["total"]) + 1
					brand_counts[key]["ranks"][rank] += 1
					if rank > max_rank:
						max_rank = rank

	# Build DataFrame
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
		return pd.DataFrame(columns=["brand", "total"])  # empty

	out_df = pd.DataFrame(rows)
	out_df.sort_values(by=["total", "brand"], ascending=[False, True], inplace=True)
	out_df.reset_index(drop=True, inplace=True)
	# Reorder columns: brand, total, rank_1..k
	out_df = out_df[["brand", "total", *rank_cols]]
	return out_df


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
				items = parse_response_cell(val)
				for rank, brand in items:
					records.append({
						"brand": brand,
						"rank": int(rank),
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
		# Original aggregate
		df = count_brands_and_ranks(in_path)

		# New: build records and compute grouped outputs/metrics
		datasets = _get_datasets(in_path)
		group_fields = ["Role", "Model", "LLM Provider"]
		records, parse_summary = _records_from_datasets(datasets, group_fields)

		# Overall counts (redundant with df but used for metrics)
		df_overall, _ = _aggregate_brand_counts(records)
		metrics = _compute_metrics(df_overall)

		# Grouped counts
		grouped_outputs = {}
		for g in group_fields:
			# pivot per group value
			sub_records = {}
			for r in records:
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

		# Rank distribution (overall)
		rank_dist = _overall_rank_distribution(df_overall)

	except Exception as e:
		print(f"Error while counting brands: {e}")
		return 1

	if df.empty:
		print("No brand data found in any 'response' column.")
		return 0

	# Write outputs
	out_path = Path(args.output)
	try:
		df.to_csv(out_path, index=False)
		print(f"\nTop brands (first 20):")
		print(df.head(20).to_string(index=False))
		print(f"\nSaved summary to: {out_path.resolve()}")

		# Metrics and grouped outputs
		_write_csv(Path("brand_metrics.csv"), metrics, "brand metrics")
		_write_csv(Path("brand_counts_by_role.csv"), grouped_outputs.get("Role", pd.DataFrame()), "brand counts by Role")
		_write_csv(Path("brand_counts_by_model.csv"), grouped_outputs.get("Model", pd.DataFrame()), "brand counts by Model")
		_write_csv(Path("brand_counts_by_provider.csv"), grouped_outputs.get("LLM Provider", pd.DataFrame()), "brand counts by LLM Provider")
		_write_csv(Path("overall_rank_distribution.csv"), rank_dist, "overall rank distribution")

		# Parse summary
		ps = pd.DataFrame([parse_summary])
		_write_csv(Path("parse_summary.csv"), ps, "parse summary")
		return 0
	except Exception as e:
		print(f"Failed to write output CSVs: {e}")
		return 1


if __name__ == "__main__":
	sys.exit(main())

