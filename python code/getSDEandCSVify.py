#!/usr/bin/env python3
"""Convert every JSONL file in _sde into CSV files in testing/csv.

Usage:
	python testing/jsontocsv.py
"""

from __future__ import annotations

import csv
import json
import sys
from pathlib import Path


HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
SDE_DIR = ROOT / "_sde"
OUT_DIR = HERE / "csv"


def _to_cell(value):
	"""Convert JSON values to CSV-safe scalar text."""
	if value is None:
		return ""
	if isinstance(value, (dict, list)):
		return json.dumps(value, ensure_ascii=False, sort_keys=True)
	return value


def _load_jsonl(path: Path) -> list[dict]:
	"""Read a JSONL file into a list of dict rows."""
	rows: list[dict] = []
	with path.open("r", encoding="utf-8") as f:
		for line_no, line in enumerate(f, start=1):
			line = line.strip()
			if not line:
				continue
			try:
				obj = json.loads(line)
			except json.JSONDecodeError as exc:
				print(f"  WARN: {path.name}:{line_no} invalid JSON ({exc})", flush=True)
				continue

			if isinstance(obj, dict):
				rows.append(obj)
			else:
				# Keep non-dict rows by wrapping in a single value column.
				rows.append({"value": obj})
	return rows


def _write_csv(rows: list[dict], out_path: Path) -> None:
	"""Write rows to CSV using the union of keys across all rows."""
	fieldnames = sorted({k for row in rows for k in row.keys()})
	if not fieldnames:
		fieldnames = ["value"]

	out_path.parent.mkdir(parents=True, exist_ok=True)
	with out_path.open("w", encoding="utf-8", newline="") as f:
		writer = csv.DictWriter(f, fieldnames=fieldnames)
		writer.writeheader()
		for row in rows:
			writer.writerow({k: _to_cell(row.get(k)) for k in fieldnames})


def main() -> int:
	if not SDE_DIR.exists():
		print(f"ERROR: SDE directory not found: {SDE_DIR}", flush=True)
		return 1

	jsonl_files = sorted(SDE_DIR.glob("*.jsonl"))
	if not jsonl_files:
		print(f"No JSONL files found in {SDE_DIR}", flush=True)
		return 0

	OUT_DIR.mkdir(parents=True, exist_ok=True)

	converted = 0
	skipped = 0

	print(f"Found {len(jsonl_files)} JSONL files in {SDE_DIR}", flush=True)
	for src in jsonl_files:
		rows = _load_jsonl(src)
		if not rows:
			skipped += 1
			print(f"  SKIP: {src.name} (no valid rows)", flush=True)
			continue

		dst = OUT_DIR / f"{src.stem}.csv"
		_write_csv(rows, dst)
		converted += 1
		print(f"  OK: {src.name} -> {dst.name} ({len(rows):,} rows)", flush=True)

	print(
		f"Done. Converted {converted} file(s), skipped {skipped}. Output: {OUT_DIR}",
		flush=True,
	)
	return 0


if __name__ == "__main__":
	raise SystemExit(main())
