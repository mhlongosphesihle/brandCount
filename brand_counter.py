"""
Reads the Excel workbook and prints the column headers for each sheet.
Requires: pandas, openpyxl
"""

from pathlib import Path
import sys

try:
	import pandas as pd
except ImportError:
	print("Missing dependency: pandas. Install with 'py -m pip install pandas openpyxl'.")
	sys.exit(1)


def main() -> int:
	path = Path(r'All Questions - 25 Sept.xlsx')
	if not path.exists():
		print(f"File not found: {path}")
		return 2

	try:
		xls = pd.ExcelFile(path)
		print(f"Opened workbook: {path}")
		sheet_names = xls.sheet_names
		if not sheet_names:
			print("No sheets found in the workbook.")
			return 3

		print(f"Workbook: {path}")
		for sheet in sheet_names:
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
		return 0
	except Exception as e:
		print(f"Failed to open workbook: {e}")
		return 1


if __name__ == "__main__":
	print("Brand Counter - List Column Headers in Excel Sheets")
	main()

