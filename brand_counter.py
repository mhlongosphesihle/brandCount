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


WORKBOOK_NAME = "All Questions - 25 Sept.xlsx"


def print_headers(xlsx_path: Path) -> None:
	"""Utility to list sheet names and column headers."""
	xls = pd.ExcelFile(xlsx_path)
	print(f"Workbook: {xlsx_path}")
	for sheet in xls.sheet_names:
		try:
			df = pd.read_excel(xlsx_path, sheet_name=sheet, nrows=0)
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


def count_brands_and_ranks(xlsx_path: Path) -> pd.DataFrame:
	"""Read the workbook and produce a DataFrame with brand totals and rank counts.

	Returns columns: [brand, total, rank_1, rank_2, ..., rank_k]
	"""
	brand_counts: dict[str, dict[str, object]] = {}
	display_name: dict[str, str] = {}
	max_rank = 0

	xls = pd.ExcelFile(xlsx_path)
	for sheet in xls.sheet_names:
		try:
			df = pd.read_excel(xlsx_path, sheet_name=sheet)
		except Exception as e:
			print(f"Skipping sheet '{sheet}' due to read error: {e}")
			continue

		response_cols = find_response_columns(df)
		if not response_cols:
			# No response-like columns in this sheet; skip
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
						# preserve the first-seen display name
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


def main() -> int:
	# Allow optional CLI args: [workbook] [--headers]
	import argparse

	parser = argparse.ArgumentParser(description="Count brand mentions and rank positions from Excel response columns.")
	parser.add_argument("workbook", nargs="?", default=WORKBOOK_NAME, help="Path to the Excel workbook (.xlsx)")
	parser.add_argument("--headers", action="store_true", help="Only list sheet headers and exit")
	parser.add_argument("--output", "-o", default="brand_counts.csv", help="Output CSV filename")
	args = parser.parse_args()

	xlsx_path = Path(args.workbook)
	if not xlsx_path.exists():
		print(f"File not found: {xlsx_path}")
		return 2

	if args.headers:
		print_headers(xlsx_path)
		return 0

	print(f"Reading: {xlsx_path}")
	try:
		df = count_brands_and_ranks(xlsx_path)
	except Exception as e:
		print(f"Error while counting brands: {e}")
		return 1

	if df.empty:
		print("No brand data found in any 'response' column.")
		return 0

	out_path = Path(args.output)
	try:
		df.to_csv(out_path, index=False)
		# Print a small sample to stdout
		print(f"\nTop brands (first 20):")
		print(df.head(20).to_string(index=False))
		print(f"\nSaved summary to: {out_path.resolve()}")
		return 0
	except Exception as e:
		print(f"Failed to write output CSV: {e}")
		return 1


if __name__ == "__main__":
	sys.exit(main())

