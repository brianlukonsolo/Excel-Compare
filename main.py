#!/usr/bin/env python3
"""
excel_seq_diff.py

Sequential Excel diff tool.

Features:
- Compare any number of spreadsheet files sequentially (file1 -> file2, file2 -> file3, ...)
- Supports xlsx/xls/csv/tsv; chooses engine per-file.
- Configurable inclusion of additions/modifications/removals in the final text report.
- Optional CSV export of diffs (per comparison).
- Optional terminal printing.
- All env-vars at top; can be overridden with .env and/or CLI.
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
# ENVIRONMENT / Configuration
# ----------------------------
# Default values (can be overridden by .env or CLI)
# Place all env-vars here
DEFAULT_INPUTS = "inputs/sample_v1.xlsx,inputs/sample_v2.xlsx"  # comma-separated list of files
DEFAULT_SHEETS = ""  # comma-separated list of sheet names (same order as inputs). blank means first sheet for each
DEFAULT_INDEX_COL = "ID"  # logical key column name to use as index if present in both frames
OUTPUT_DIR = os.getenv("OUTPUT_DIR", "outputs")
REPORT_FILENAME = os.getenv("REPORT_FILENAME", "excel_diff_report.txt")
EXPORT_CSV = os.getenv("EXPORT_CSV", "true")  # "true"/"false"
PRINT_TERMINAL = os.getenv("PRINT_TERMINAL", "true")  # "true"/"false"
INCLUDE_ADDITIONS = os.getenv("INCLUDE_ADDITIONS", "true")
INCLUDE_MODIFICATIONS = os.getenv("INCLUDE_MODIFICATIONS", "true")
INCLUDE_REMOVALS = os.getenv("INCLUDE_REMOVALS", "true")
# You can create a .env in the working directory with these names to override.

# Load .env file if present (overrides os.environ defaults above)
load_dotenv()  # reads .env and updates environment variables

# Helper to re-read env values after loading .env
def _env(name: str, default: str) -> str:
    return os.getenv(name, default)

# Now final configuration values (they can be overridden with CLI args later)
INPUTS = _env("INPUTS", DEFAULT_INPUTS)
SHEETS = _env("SHEETS", DEFAULT_SHEETS)
INDEX_COL = _env("INDEX_COL", DEFAULT_INDEX_COL)
OUTPUT_DIR = _env("OUTPUT_DIR", OUTPUT_DIR)
REPORT_FILENAME = _env("REPORT_FILENAME", REPORT_FILENAME)
EXPORT_CSV = _env("EXPORT_CSV", EXPORT_CSV).lower() == "true"
PRINT_TERMINAL = _env("PRINT_TERMINAL", PRINT_TERMINAL).lower() == "true"
INCLUDE_ADDITIONS = _env("INCLUDE_ADDITIONS", INCLUDE_ADDITIONS).lower() == "true"
INCLUDE_MODIFICATIONS = _env("INCLUDE_MODIFICATIONS", INCLUDE_MODIFICATIONS).lower() == "true"
INCLUDE_REMOVALS = _env("INCLUDE_REMOVALS", INCLUDE_REMOVALS).lower() == "true"

# ----------------------------
# Utilities
# ----------------------------
def parse_bool(value: str) -> bool:
    return str(value).lower() in ("1", "true", "yes", "y", "on")

def _choose_engine_by_suffix(path: str) -> Optional[str]:
    suffix = Path(path).suffix.lower()
    if suffix in (".xlsx", ".xlsm", ".ods", ".odf"):
        # openpyxl is safe for .xlsx
        return "openpyxl"
    if suffix == ".xls":
        return "xlrd"
    return None

def _load_table(path: str, sheet: Optional[str] = None) -> Tuple[pd.DataFrame, str]:
    """
    Load a table from a file. Returns (DataFrame, actual_sheet_used_or_empty_for_csv).
    Supports .csv .tsv .xlsx .xls (uses engines per file).
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

    # For Excel choose engine per extension
    engine = _choose_engine_by_suffix(path)

    # read sheet names to find the first if needed
    try:
        if engine:
            with pd.ExcelFile(path, engine=engine) as xls:
                sheet_names = xls.sheet_names
        else:
            with pd.ExcelFile(path) as xls:
                sheet_names = xls.sheet_names
        if sheet is None or sheet == "":
            used_sheet = sheet_names[0]
            if engine:
                df = pd.read_excel(path, sheet_name=used_sheet, engine=engine)
            else:
                df = pd.read_excel(path, sheet_name=used_sheet)
            return df, used_sheet
        else:
            # try to read requested sheet
            if sheet in sheet_names:
                if engine:
                    df = pd.read_excel(path, sheet_name=sheet, engine=engine)
                else:
                    df = pd.read_excel(path, sheet_name=sheet)
                return df, sheet
            else:
                # fallback to first sheet and warn
                used_sheet = sheet_names[0]
                print(f"Warning: sheet '{sheet}' not found in {path}. Using first sheet '{used_sheet}'.", file=sys.stderr)
                if engine:
                    df = pd.read_excel(path, sheet_name=used_sheet, engine=engine)
                else:
                    df = pd.read_excel(path, sheet_name=used_sheet)
                return df, used_sheet
    except ImportError as ie:
        # suggest installation
        if engine == "openpyxl":
            raise RuntimeError("Missing dependency 'openpyxl'. Install with: pip install openpyxl") from ie
        if engine == "xlrd":
            raise RuntimeError("Missing dependency 'xlrd'. Install with: pip install xlrd") from ie
        raise
    except Exception as e:
        raise RuntimeError(f"Failed to read '{path}': {e}") from e

# Reusable function that compares two tables and returns dicts of DataFrames
def diff_pair(file1: str, sheet1: Optional[str], file2: str, sheet2: Optional[str], index_col: Optional[str]) -> Dict[str, pd.DataFrame]:
    """
    Compare file1:sheet1 (left) with file2:sheet2 (right).
    Returns dict with keys 'added','removed','modified' (DataFrames).
    'added' = rows in right but not left (source column indicates right file:sheet)
    'removed' = rows in left but not right (source column indicates left file:sheet)
    'modified' = DataFrame where each row has per-column "old -> new" strings and columns _old_source/_new_source.
    """
    df1, used_sheet1 = _load_table(file1, sheet1)
    df2, used_sheet2 = _load_table(file2, sheet2)

    src1 = f"{Path(file1).name}:{used_sheet1 or 'CSV'}"
    src2 = f"{Path(file2).name}:{used_sheet2 or 'CSV'}"

    # Attempt to set index if index_col exists in both
    use_index = False
    if index_col and index_col in df1.columns and index_col in df2.columns:
        df1i = df1.set_index(index_col)
        df2i = df2.set_index(index_col)
        use_index = True
    else:
        # fall back to implicit positional index (preserve alignment)
        df1i = df1.copy()
        df2i = df2.copy()

    # align columns (outer) to have same columns
    left, right = df1i.align(df2i, join='outer', axis=1)

    # Find added/removed by index labels
    added = right.loc[~right.index.isin(left.index)].copy()
    removed = left.loc[~left.index.isin(right.index)].copy()

    # attach source metadata
    if not added.empty:
        added.insert(0, "_source", src2)
    if not removed.empty:
        removed.insert(0, "_source", src1)

    # Compute modifications for common indices
    common_idx = left.index.intersection(right.index)
    common_left = left.loc[common_idx].copy()
    common_right = right.loc[common_idx].copy()

    sentinel = object()
    fleft = common_left.fillna(sentinel)
    fright = common_right.fillna(sentinel)
    diff_mask = (fleft != fright)

    modified_rows: List[Dict[str, Any]] = []
    for idx in diff_mask.index:
        mask_row = diff_mask.loc[idx]
        if mask_row.any():
            row_changes: Dict[str, Any] = {}
            for col in mask_row.index[mask_row]:
                old = common_left.at[idx, col]
                new = common_right.at[idx, col]
                old_repr = "<NA>" if pd.isna(old) else repr(old)
                new_repr = "<NA>" if pd.isna(new) else repr(new)
                row_changes[col] = f"{old_repr} → {new_repr}"
            # include index as explicit column for later clarity
            row_changes["_index"] = idx
            row_changes["_old_source"] = src1
            row_changes["_new_source"] = src2
            modified_rows.append(row_changes)

    if modified_rows:
        modified_df = pd.DataFrame(modified_rows).set_index("_index")
    else:
        # empty DF with expected columns
        modified_df = pd.DataFrame(columns=list(left.columns) + ["_old_source", "_new_source"])

    # Return also the used sheet names to build clear report headers
    return {
        "added": added,
        "removed": removed,
        "modified": modified_df,
        "meta": {
            "file1": file1, "sheet1": used_sheet1 or "CSV",
            "file2": file2, "sheet2": used_sheet2 or "CSV",
            "src1": src1, "src2": src2,
            "use_index": use_index,
            "index_col": index_col
        }
    }

# ----------------------------
# Reporting
# ----------------------------
def format_dataframe_for_csv_export(df: pd.DataFrame, index_name: Optional[str]) -> pd.DataFrame:
    """
    Prepare DataFrame so that index becomes explicit _index column (or the actual index_name).
    Returns a new DataFrame ready to to_csv(index=False).
    """
    if df.empty:
        return df
    out = df.copy()
    # ensure index is exposed as column
    if out.index.name is None:
        # unnamed index; call it '_index'
        out.insert(0, "_index", out.index)
    else:
        out.insert(0, out.index.name, out.index)
    out.reset_index(drop=True, inplace=True)
    return out

def generate_sequential_report(inputs: List[str], sheets: List[Optional[str]], index_col: Optional[str],
                               out_dir: Path, report_filename: str,
                               include_additions: bool, include_mods: bool, include_removals: bool,
                               export_csv: bool, print_terminal: bool) -> None:
    """
    Compare inputs sequentially and write a text report + optional CSVs.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    report_path = out_dir / report_filename

    with report_path.open("w", encoding="utf-8") as rep:
        header = f"Sequential Excel Diff Report\nInputs: {', '.join(inputs)}\n\n"
        rep.write(header)
        if print_terminal:
            print(header)

        # iterate pairs
        for i in range(len(inputs) - 1):
            left = inputs[i]
            right = inputs[i + 1]
            sheet_left = sheets[i] if i < len(sheets) else None
            sheet_right = sheets[i + 1] if i + 1 < len(sheets) else None

            rep.write("=" * 80 + "\n")
            title = f"Comparison {i+1}: {Path(left).name}:{sheet_left or '(first)'}  --->  {Path(right).name}:{sheet_right or '(first)'}\n"
            rep.write(title)
            rep.write("-" * 80 + "\n")
            if print_terminal:
                print("=" * 80)
                print(title.strip())
                print("-" * 80)

            result = diff_pair(left, sheet_left, right, sheet_right, index_col)

            # ADDED
            if include_additions:
                rep.write("ADDED (rows present in RIGHT but not LEFT)\n")
                if result["added"].empty:
                    rep.write("  None\n\n")
                    if print_terminal:
                        print("ADDED: None")
                else:
                    # write a readable table-like dump
                    added_df = result["added"].copy()
                    # expose index explicitly
                    added_df.insert(0, "_index", added_df.index)
                    rep.write(added_df.to_csv(index=False))
                    rep.write("\n")
                    if print_terminal:
                        print("ADDED:")
                        print(added_df.to_string(index=False))
                    if export_csv:
                        csv_name = out_dir / f"diff_{i+1:02d}_added_{Path(right).stem}.csv"
                        added_df.to_csv(csv_name, index=False)
                        rep.write(f"  [CSV saved: {csv_name.name}]\n\n")
                        if print_terminal:
                            print(f"  Saved CSV: {csv_name}")
            else:
                rep.write("ADDED (excluded by config)\n\n")

            # REMOVED
            if include_removals:
                rep.write("REMOVED (rows present in LEFT but not RIGHT)\n")
                if result["removed"].empty:
                    rep.write("  None\n\n")
                    if print_terminal:
                        print("REMOVED: None")
                else:
                    rem_df = result["removed"].copy()
                    rem_df.insert(0, "_index", rem_df.index)
                    rep.write(rem_df.to_csv(index=False))
                    rep.write("\n")
                    if print_terminal:
                        print("REMOVED:")
                        print(rem_df.to_string(index=False))
                    if export_csv:
                        csv_name = out_dir / f"diff_{i+1:02d}_removed_{Path(left).stem}.csv"
                        rem_df.to_csv(csv_name, index=False)
                        rep.write(f"  [CSV saved: {csv_name.name}]\n\n")
                        if print_terminal:
                            print(f"  Saved CSV: {csv_name}")
            else:
                rep.write("REMOVED (excluded by config)\n\n")

            # MODIFIED
            if include_mods:
                rep.write("MODIFIED (rows present in both but with changed cells)\n")
                if result["modified"].empty:
                    rep.write("  None\n\n")
                    if print_terminal:
                        print("MODIFIED: None")
                else:
                    mod_df = result["modified"].copy().reset_index()
                    # nice column order: index, _old_source, _new_source, then changed columns
                    cols = mod_df.columns.tolist()
                    # ensure index is first
                    # move _old_source/_new_source to the end of header block for clarity in CSV
                    rep.write(mod_df.to_csv(index=False))
                    rep.write("\n")
                    if print_terminal:
                        print("MODIFIED:")
                        print(mod_df.to_string(index=False))
                    if export_csv:
                        csv_name = out_dir / f"diff_{i+1:02d}_modified_{Path(left).stem}_to_{Path(right).stem}.csv"
                        mod_df.to_csv(csv_name, index=False)
                        rep.write(f"  [CSV saved: {csv_name.name}]\n\n")
                        if print_terminal:
                            print(f"  Saved CSV: {csv_name}")
            else:
                rep.write("MODIFIED (excluded by config)\n\n")

            # mark end of comparison block
            rep.write("\n")
            if print_terminal:
                print("\n")  # blank line between comparisons

    if print_terminal:
        print(f"Report saved to {report_path.resolve()}")
    else:
        # always inform user where report was written (non-verbose)
        print(f"Report generated: {report_path.resolve()}")

# ----------------------------
# CLI
# ----------------------------
def cli():
    parser = argparse.ArgumentParser(description="Sequential Excel diff tool.")
    parser.add_argument("--inputs", "-i", help="Comma-separated list of input files in order (overrides INPUTS env)", default=None)
    parser.add_argument("--sheets", "-s", help="Comma-separated list of sheet names (same order), blank for first sheet", default=None)
    parser.add_argument("--index-col", help="Index column name to use as key (optional)", default=None)
    parser.add_argument("--output-dir", help="Output directory (overrides OUTPUT_DIR env)", default=None)
    parser.add_argument("--report", help="Report filename (overrides REPORT_FILENAME env)", default=None)
    parser.add_argument("--export-csv", help="Export CSVs (true/false)", default=None)
    parser.add_argument("--print-terminal", help="Print to terminal (true/false)", default=None)
    parser.add_argument("--include-additions", help="Include additions in report (true/false)", default=None)
    parser.add_argument("--include-modifications", help="Include modifications in report (true/false)", default=None)
    parser.add_argument("--include-removals", help="Include removals in report (true/false)", default=None)
    args = parser.parse_args()

    # Resolve configuration order: CLI > env > defaults
    inputs_raw = args.inputs or os.getenv("INPUTS") or DEFAULT_INPUTS
    sheets_raw = args.sheets or os.getenv("SHEETS") or DEFAULT_SHEETS
    index_col = args.index_col if args.index_col is not None else (os.getenv("INDEX_COL") or INDEX_COL)
    output_dir = args.output_dir or os.getenv("OUTPUT_DIR") or OUTPUT_DIR
    report_name = args.report or os.getenv("REPORT_FILENAME") or REPORT_FILENAME

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

    # parse inputs/sheets into lists
    inputs = [p.strip() for p in inputs_raw.split(",") if p.strip()]
    if not inputs or len(inputs) < 2:
        parser.error("At least two input files are required (provide --inputs or set INPUTS in .env).")

    if sheets_raw and sheets_raw.strip():
        sheets = [s.strip() if s.strip() else None for s in sheets_raw.split(",")]
    else:
        sheets = [None] * len(inputs)

    # ensure sheets list length equals inputs length (pad with None if necessary)
    if len(sheets) < len(inputs):
        sheets.extend([None] * (len(inputs) - len(sheets)))

    # run generator
    generate_sequential_report(inputs, sheets, index_col, Path(output_dir), report_name,
                               include_add, include_mod, include_rem, export_csv, print_terminal)

if __name__ == "__main__":
    cli()
