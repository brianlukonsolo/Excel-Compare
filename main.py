#!/usr/bin/env python3
"""
excel_seq_diff.py

Sequential Excel diff tool.

Features:
- Compare any number of spreadsheet files sequentially (file1 -> file2, file2 -> file3, ...)
- Supports xlsx/xls/csv/tsv; chooses engine per-file, with fallback attempts.
- Optional filtering of columns to compare.
- Configurable inclusion of additions/modifications/removals in the final text report.
- Optional CSV export of diffs (per comparison).
- Optional terminal printing.
- All env-vars at top; can be overridden with .env or CLI.
- Report saved to outputs/<REPORT_FILENAME>.txt
"""

from pathlib import Path
import pandas as pd
import sys
import argparse
from typing import Tuple, Dict, Any, Optional, List
from dotenv import load_dotenv
import os

# ----------------------------
# ENVIRONMENT / Configuration (defaults; override via .env or CLI)
# ----------------------------
DEFAULT_INPUTS = "inputs/sample_v1.xlsx,inputs/sample_v2.xlsx"
DEFAULT_SHEETS = ""            # comma-separated list of sheets (blank -> first sheet)
DEFAULT_INDEX_COL = "ID"
DEFAULT_COLUMNS = ""          # comma-separated list of columns to compare; blank = all
DEFAULT_OUTPUT_DIR = "outputs"
DEFAULT_REPORT_FILENAME = "excel_diff_report.txt"
DEFAULT_EXPORT_CSV = "true"
DEFAULT_PRINT_TERMINAL = "true"
DEFAULT_INCLUDE_ADDITIONS = "true"
DEFAULT_INCLUDE_MODIFICATIONS = "true"
DEFAULT_INCLUDE_REMOVALS = "true"

# Load .env (if present) to override environment variables
load_dotenv()

def _env(name: str, default: str) -> str:
    return os.getenv(name, default)

# Read configuration (these may be overridden again by CLI)
INPUTS = _env("INPUTS", DEFAULT_INPUTS)
SHEETS = _env("SHEETS", DEFAULT_SHEETS)
INDEX_COL = _env("INDEX_COL", DEFAULT_INDEX_COL)
COLUMNS = _env("COLUMNS", DEFAULT_COLUMNS)
OUTPUT_DIR = _env("OUTPUT_DIR", DEFAULT_OUTPUT_DIR)
REPORT_FILENAME = _env("REPORT_FILENAME", DEFAULT_REPORT_FILENAME)
EXPORT_CSV = _env("EXPORT_CSV", DEFAULT_EXPORT_CSV).lower() in ("1", "true", "yes", "y")
PRINT_TERMINAL = _env("PRINT_TERMINAL", DEFAULT_PRINT_TERMINAL).lower() in ("1", "true", "yes", "y")
INCLUDE_ADDITIONS = _env("INCLUDE_ADDITIONS", DEFAULT_INCLUDE_ADDITIONS).lower() in ("1", "true", "yes", "y")
INCLUDE_MODIFICATIONS = _env("INCLUDE_MODIFICATIONS", DEFAULT_INCLUDE_MODIFICATIONS).lower() in ("1", "true", "yes", "y")
INCLUDE_REMOVALS = _env("INCLUDE_REMOVALS", DEFAULT_INCLUDE_REMOVALS).lower() in ("1", "true", "yes", "y")

# ----------------------------
# Helpers
# ----------------------------
def parse_bool(value: Optional[str]) -> bool:
    if value is None:
        return False
    return str(value).lower() in ("1", "true", "yes", "y", "on")

def _choose_engines_for_suffix(suffix: str) -> List[str]:
    """
    Return preferred engines to try for an Excel file suffix.
    """
    suffix = suffix.lower()
    if suffix == ".xls":
        return ["xlrd", "openpyxl"]     # try xlrd first for .xls, then openpyxl as fallback
    if suffix in (".xlsx", ".xlsm", ".ods", ".odf"):
        return ["openpyxl", "xlrd"]     # try openpyxl first for modern formats
    return []

def _load_table(path: str, sheet: Optional[str] = None) -> Tuple[pd.DataFrame, str]:
    """
    Load a CSV/TSV or Excel file (.xls/.xlsx) and return (DataFrame, used_sheet_name_or_empty_for_csv).
    Tries multiple engines for Excel files (fallback) before giving up.
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"Input file not found: {path}")
    suffix = p.suffix.lower()

    # CSV / TSV
    if suffix in (".csv", ".tsv"):
        sep = "\t" if suffix == ".tsv" else ","
        df = pd.read_csv(path, sep=sep)
        return df, ""

    # Determine engine order
    engines = _choose_engines_for_suffix(suffix)
    if not engines:
        raise ValueError(f"Unsupported file extension: {suffix}")

    last_exc: Optional[Exception] = None
    for engine in engines:
        try:
            # Use ExcelFile to get sheet names with the chosen engine
            with pd.ExcelFile(path, engine=engine) as xls:
                sheet_names = xls.sheet_names
            # pick sheet to use
            used_sheet = sheet if (sheet and sheet in sheet_names) else sheet_names[0]
            if sheet and sheet not in sheet_names:
                print(f"Warning: sheet '{sheet}' not found in {path}. Using '{used_sheet}' instead.", file=sys.stderr)
            df = pd.read_excel(path, sheet_name=used_sheet, engine=engine)
            # success
            return df, used_sheet
        except ImportError as ie:
            # missing engine dependency: raise with helpful message
            if engine == "openpyxl":
                raise RuntimeError("Missing dependency 'openpyxl'. Install with: pip install openpyxl") from ie
            if engine == "xlrd":
                raise RuntimeError("Missing dependency 'xlrd' (use xlrd==1.2.0 for .xls support). Install with: pip install xlrd==1.2.0") from ie
        except Exception as e:
            # record and try next engine
            last_exc = e
            continue

    # If we fall through, all engines failed
    raise RuntimeError(f"Failed to read '{path}' with tried engines. Last error: {last_exc}") from last_exc

def _expose_index_column(df: pd.DataFrame, index_name: Optional[str]) -> pd.DataFrame:
    """
    Return a copy of df with its index exposed as the first column.
    If index_name is provided, use it; otherwise default to '_index'.
    """
    if df.empty:
        return df
    out = df.copy()
    name = index_name or "_index"
    # If index already has a name and matches requested, use it
    out.insert(0, name, out.index)
    out.reset_index(drop=True, inplace=True)
    return out

# ----------------------------
# Core diff logic
# ----------------------------
def diff_pair(file1: str, sheet1: Optional[str], file2: str, sheet2: Optional[str],
              index_col: Optional[str], columns_to_check: Optional[List[str]] = None) -> Dict[str, Any]:
    """
    Compare left (file1:sheet1) vs right (file2:sheet2).
    Returns dict with keys: 'added','removed','modified','meta'
    """
    df1, used_sheet1 = _load_table(file1, sheet1)
    df2, used_sheet2 = _load_table(file2, sheet2)

    # Apply column selection if provided (keep only columns that exist in each file)
    if columns_to_check:
        df1 = df1[[c for c in columns_to_check if c in df1.columns]].copy()
        df2 = df2[[c for c in columns_to_check if c in df2.columns]].copy()

    src1 = f"{Path(file1).name}:{used_sheet1 or 'CSV'}"
    src2 = f"{Path(file2).name}:{used_sheet2 or 'CSV'}"

    # Attempt to use index_col if present in both
    use_index = False
    if index_col and index_col in df1.columns and index_col in df2.columns:
        left_df = df1.set_index(index_col)
        right_df = df2.set_index(index_col)
        use_index = True
    else:
        # keep positional index
        left_df = df1.copy()
        right_df = df2.copy()

    # Align columns (outer) so both frames have same columns
    left_aligned, right_aligned = left_df.align(right_df, join="outer", axis=1)

    # Added/Removed (by index labels)
    added = right_aligned.loc[~right_aligned.index.isin(left_aligned.index)].copy()
    removed = left_aligned.loc[~left_aligned.index.isin(right_aligned.index)].copy()

    if not added.empty:
        added.insert(0, "_source", src2)
    if not removed.empty:
        removed.insert(0, "_source", src1)

    # Modified (for common indices)
    common_idx = left_aligned.index.intersection(right_aligned.index)
    left_common = left_aligned.loc[common_idx].copy()
    right_common = right_aligned.loc[common_idx].copy()

    sentinel = object()
    left_f = left_common.fillna(sentinel)
    right_f = right_common.fillna(sentinel)
    diff_mask = (left_f != right_f)

    modified_rows: List[Dict[str, Any]] = []
    for idx in diff_mask.index:
        row_mask = diff_mask.loc[idx]
        if row_mask.any():
            changes: Dict[str, Any] = {}
            for col in row_mask.index[row_mask]:
                old = left_common.at[idx, col]
                new = right_common.at[idx, col]
                old_repr = "<NA>" if pd.isna(old) else repr(old)
                new_repr = "<NA>" if pd.isna(new) else repr(new)
                changes[col] = f"{old_repr} → {new_repr}"
            changes["_index"] = idx
            changes["_old_source"] = src1
            changes["_new_source"] = src2
            modified_rows.append(changes)

    modified_df = pd.DataFrame(modified_rows).set_index("_index") if modified_rows else pd.DataFrame(columns=list(left_aligned.columns)+["_old_source","_new_source"])

    meta = {
        "file1": file1, "sheet1": used_sheet1 or "CSV",
        "file2": file2, "sheet2": used_sheet2 or "CSV",
        "src1": src1, "src2": src2, "use_index": use_index, "index_col": index_col
    }

    return {"added": added, "removed": removed, "modified": modified_df, "meta": meta}

# ----------------------------
# Reporting
# ----------------------------
def generate_sequential_report(inputs: List[str], sheets: List[Optional[str]], index_col: Optional[str],
                               columns_to_check: Optional[List[str]], out_dir: Path, report_filename: str,
                               include_additions: bool, include_mods: bool, include_removals: bool,
                               export_csv: bool, print_terminal: bool) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    report_path = out_dir / report_filename

    with report_path.open("w", encoding="utf-8") as rep:
        header_lines = [
            "Sequential Excel Diff Report",
            f"Inputs: {', '.join(inputs)}",
            f"Columns compared: {', '.join(columns_to_check) if columns_to_check else '(all)'}",
            f"Index column: {index_col or '(none)'}",
            "-" * 80,
            ""
        ]
        rep.write("\n".join(header_lines) + "\n")
        if print_terminal:
            print("\n".join(header_lines))

        # iterate sequentially
        for i in range(len(inputs) - 1):
            left = inputs[i]
            right = inputs[i+1]
            sheet_left = sheets[i] if i < len(sheets) else None
            sheet_right = sheets[i+1] if i+1 < len(sheets) else None

            rep.write("="*80 + "\n")
            comp_title = f"Comparison {i+1}: {Path(left).name}:{sheet_left or '(first)'}  --->  {Path(right).name}:{sheet_right or '(first)'}"
            rep.write(comp_title + "\n")
            rep.write("-"*80 + "\n")
            if print_terminal:
                print("="*80)
                print(comp_title)
                print("-"*80)

            result = diff_pair(left, sheet_left, right, sheet_right, index_col, columns_to_check)

            # ADDITIONS
            rep.write("ADDITIONS (rows present in RIGHT but not LEFT)\n")
            if include_additions:
                if result["added"].empty:
                    rep.write("  None\n\n")
                    if print_terminal: print("ADDITIONS: None")
                else:
                    added_df = _expose_index_column(result["added"].copy(), index_col)
                    rep.write(added_df.to_csv(index=False))
                    rep.write("\n")
                    if print_terminal:
                        print("ADDITIONS:")
                        print(added_df.to_string(index=False))
                    if export_csv:
                        csv_name = out_dir / f"diff_{i+1:02d}_added_{Path(right).stem}.csv"
                        added_df.to_csv(csv_name, index=False)
                        rep.write(f"[CSV saved: {csv_name.name}]\n\n")
                        if print_terminal:
                            print(f"  Saved CSV: {csv_name}")
            else:
                rep.write("  (excluded by config)\n\n")

            # REMOVALS
            rep.write("REMOVALS (rows present in LEFT but not RIGHT)\n")
            if include_removals:
                if result["removed"].empty:
                    rep.write("  None\n\n")
                    if print_terminal: print("REMOVALS: None")
                else:
                    rem_df = _expose_index_column(result["removed"].copy(), index_col)
                    rep.write(rem_df.to_csv(index=False))
                    rep.write("\n")
                    if print_terminal:
                        print("REMOVALS:")
                        print(rem_df.to_string(index=False))
                    if export_csv:
                        csv_name = out_dir / f"diff_{i+1:02d}_removed_{Path(left).stem}.csv"
                        rem_df.to_csv(csv_name, index=False)
                        rep.write(f"[CSV saved: {csv_name.name}]\n\n")
                        if print_terminal:
                            print(f"  Saved CSV: {csv_name}")
            else:
                rep.write("  (excluded by config)\n\n")

            # MODIFIED
            rep.write("MODIFIED (rows present in BOTH but with changed cells)\n")
            if include_mods:
                if result["modified"].empty:
                    rep.write("  None\n\n")
                    if print_terminal: print("MODIFIED: None")
                else:
                    mod_df = result["modified"].copy().reset_index()
                    # reorder columns for readability: index, changed cols..., _old_source, _new_source
                    cols = mod_df.columns.tolist()
                    # ensure _index is first column (reset_index always creates _index)
                    if cols[0] != "_index" and "_index" in cols:
                        cols = [c for c in cols if c != "_index"]
                        cols.insert(0, "_index")
                    # move source columns to the end
                    for s in ("_old_source", "_new_source"):
                        if s in cols:
                            cols = [c for c in cols if c != s] + [s]
                    mod_df = mod_df[cols]
                    rep.write(mod_df.to_csv(index=False))
                    rep.write("\n")
                    if print_terminal:
                        print("MODIFIED:")
                        print(mod_df.to_string(index=False))
                    if export_csv:
                        csv_name = out_dir / f"diff_{i+1:02d}_modified_{Path(left).stem}_to_{Path(right).stem}.csv"
                        mod_df.to_csv(csv_name, index=False)
                        rep.write(f"[CSV saved: {csv_name.name}]\n\n")
                        if print_terminal:
                            print(f"  Saved CSV: {csv_name}")
            else:
                rep.write("  (excluded by config)\n\n")

            rep.write("\n")  # spacer between comparisons
            if print_terminal:
                print("\n")

    # final notice
    if print_terminal:
        print(f"Report saved to {report_path.resolve()}")
    else:
        print(f"Report generated: {report_path.resolve()}")

# ----------------------------
# CLI
# ----------------------------
def cli():
    parser = argparse.ArgumentParser(description="Sequential Excel diff tool.")
    parser.add_argument("--inputs", "-i", help="Comma-separated list of input files in order (overrides INPUTS env)", default=None)
    parser.add_argument("--sheets", "-s", help="Comma-separated list of sheet names (same order), blank means first sheet", default=None)
    parser.add_argument("--index-col", help="Index column name to use as key (optional)", default=None)
    parser.add_argument("--columns", "-c", help="Comma-separated list of column names to compare (optional)", default=None)
    parser.add_argument("--output-dir", help="Output directory (overrides OUTPUT_DIR env)", default=None)
    parser.add_argument("--report", help="Report filename (overrides REPORT_FILENAME env)", default=None)
    parser.add_argument("--export-csv", help="Export CSVs (true/false)", default=None)
    parser.add_argument("--print-terminal", help="Print to terminal (true/false)", default=None)
    parser.add_argument("--include-additions", help="Include additions in report (true/false)", default=None)
    parser.add_argument("--include-modifications", help="Include modifications in report (true/false)", default=None)
    parser.add_argument("--include-removals", help="Include removals in report (true/false)", default=None)

    args = parser.parse_args()

    # Resolve config: CLI > env > defaults
    inputs_raw = args.inputs or os.getenv("INPUTS") or DEFAULT_INPUTS
    sheets_raw = args.sheets or os.getenv("SHEETS") or DEFAULT_SHEETS
    index_col = args.index_col if args.index_col is not None else (os.getenv("INDEX_COL") or DEFAULT_INDEX_COL)
    columns_raw = args.columns or os.getenv("COLUMNS") or DEFAULT_COLUMNS
    output_dir = args.output_dir or os.getenv("OUTPUT_DIR") or DEFAULT_OUTPUT_DIR
    report_name = args.report or os.getenv("REPORT_FILENAME") or DEFAULT_REPORT_FILENAME

    def decide_bool(cli_val, env_name, default_val):
        if cli_val is not None:
            return parse_bool(cli_val)
        env_val = os.getenv(env_name)
        if env_val is not None:
            return parse_bool(env_val)
        return default_val

    export_csv = decide_bool(args.export_csv, "EXPORT_CSV", EXPORT_CSV)
    print_terminal = decide_bool(args.print_terminal, "PRINT_TERMINAL", PRINT_TERMINAL)
    include_add = decide_bool(args.include_additions, "INCLUDE_ADDITIONS", INCLUDE_ADDITIONS)
    include_mod = decide_bool(args.include_modifications, "INCLUDE_MODIFICATIONS", INCLUDE_MODIFICATIONS)
    include_rem = decide_bool(args.include_removals, "INCLUDE_REMOVALS", INCLUDE_REMOVALS)

    inputs = [p.strip() for p in inputs_raw.split(",") if p.strip()]
    if not inputs or len(inputs) < 2:
        parser.error("At least two input files are required (provide --inputs or set INPUTS in .env).")

    if sheets_raw and sheets_raw.strip():
        sheets = [s.strip() if s.strip() else None for s in sheets_raw.split(",")]
    else:
        sheets = [None] * len(inputs)

    if len(sheets) < len(inputs):
        sheets.extend([None] * (len(inputs) - len(sheets)))

    columns_to_check = [c.strip() for c in columns_raw.split(",") if c.strip()] if columns_raw and columns_raw.strip() else None

    generate_sequential_report(inputs, sheets, index_col, columns_to_check, Path(output_dir), report_name,
                               include_add, include_mod, include_rem, export_csv, print_terminal)

if __name__ == "__main__":
    cli()
