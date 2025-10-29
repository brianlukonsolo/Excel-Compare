#!/usr/bin/env python3
"""
excel_diff_verbose.py  (updated to include explicit _index column in CSV outputs)
"""
from pathlib import Path
import pandas as pd
import sys
from typing import Tuple, Dict, Any, Optional

# ... keep the _load_table and excel_diff definitions above unchanged ...
# (paste your existing definitions for _load_table and excel_diff here)
# For brevity in this snippet I will re-include them exactly as you provided earlier.
# (If you already have them in your file, just replace the bottom section with the code below.)

def _load_table(path: str, sheet: Optional[str] = None) -> Tuple[pd.DataFrame, str]:
    p = Path(path)
    suffix = p.suffix.lower()
    if suffix in (".csv", ".tsv"):
        sep = "\t" if suffix == ".tsv" else ","
        df = pd.read_csv(path, sep=sep)
        return df, ""
    try:
        if sheet is None:
            with pd.ExcelFile(path) as xls:
                first_sheet = xls.sheet_names[0]
            df = pd.read_excel(path, sheet_name=first_sheet)
            return df, first_sheet
        else:
            try:
                df = pd.read_excel(path, sheet_name=sheet)
                return df, sheet
            except ValueError:
                with pd.ExcelFile(path) as xls:
                    first_sheet = xls.sheet_names[0]
                print(f"Warning: sheet '{sheet}' not found in {path}. Using first sheet '{first_sheet}' instead.", file=sys.stderr)
                df = pd.read_excel(path, sheet_name=first_sheet)
                return df, first_sheet
    except Exception as e:
        raise RuntimeError(f"Failed to read '{path}': {e}")

def excel_diff(file1: str,
               file2: str,
               sheet1: Optional[str] = None,
               sheet2: Optional[str] = None,
               index_col: Optional[str] = None,
               verbose: bool = True) -> Dict[str, pd.DataFrame]:
    df1, used_sheet1 = _load_table(file1, sheet1)
    df2, used_sheet2 = _load_table(file2, sheet2)

    src1 = f"{Path(file1).name}:{used_sheet1 or 'CSV'}"
    src2 = f"{Path(file2).name}:{used_sheet2 or 'CSV'}"

    use_index = False
    if index_col is not None and index_col in df1.columns and index_col in df2.columns:
        df1 = df1.set_index(index_col)
        df2 = df2.set_index(index_col)
        use_index = True
    elif index_col is not None:
        print(f"Warning: index_col '{index_col}' not present in both files — comparing by implicit index.", file=sys.stderr)

    df1, df2 = df1.align(df2, join='outer', axis=1)

    added = df2.loc[~df2.index.isin(df1.index)].copy()
    removed = df1.loc[~df1.index.isin(df2.index)].copy()

    if not added.empty:
        added.insert(0, "_source", src2)
    if not removed.empty:
        removed.insert(0, "_source", src1)

    common_idx = df1.index.intersection(df2.index)
    common1 = df1.loc[common_idx].copy()
    common2 = df2.loc[common_idx].copy()

    sentinel = object()
    filled1 = common1.fillna(sentinel)
    filled2 = common2.fillna(sentinel)
    diff_mask = (filled1 != filled2)

    modified_rows: list[dict[str, Any]] = []
    for idx in diff_mask.index:
        row_mask = diff_mask.loc[idx]
        if row_mask.any():
            row_changes: dict[str, Any] = {}
            for col in row_mask.index[row_mask]:
                old = common1.at[idx, col]
                new = common2.at[idx, col]
                old_repr = "<NA>" if pd.isna(old) else repr(old)
                new_repr = "<NA>" if pd.isna(new) else repr(new)
                row_changes[col] = f"{old_repr} → {new_repr}"
            row_changes["_index"] = idx
            row_changes["_old_source"] = src1
            row_changes["_new_source"] = src2
            modified_rows.append(row_changes)

    if modified_rows:
        modified_df = pd.DataFrame(modified_rows).set_index("_index")
    else:
        modified_df = pd.DataFrame(columns=list(df1.columns) + ["_old_source", "_new_source"])

    if verbose:
        print(f"\nComparing:\n  file1 = {file1} (sheet -> '{used_sheet1 or 'CSV'}')\n  file2 = {file2} (sheet -> '{used_sheet2 or 'CSV'}')\n")
        print("=== ADDED ROWS (present in file2 only) ===")
        if added.empty:
            print("None")
        else:
            print(added)

        print("\n=== REMOVED ROWS (present in file1 only) ===")
        if removed.empty:
            print("None")
        else:
            print(removed)

        print("\n=== MODIFIED CELLS (present in both but different) ===")
        if modified_df.empty:
            print("None")
        else:
            for idx, row in modified_df.iterrows():
                print(f"\nRow {idx} (old source: {row.get('_old_source')}, new source: {row.get('_new_source')}):")
                for col, val in row.drop(labels=["_old_source", "_new_source"], errors="ignore").items():
                    if pd.isna(val):
                        continue
                    print(f"  {col}: {val}")

    return {"added": added, "removed": removed, "modified": modified_df}


if __name__ == "__main__":
    # Demo example using the sample files in inputs/
    sample1 = "inputs/sample_v1.xlsx"  # sheet 'Data'
    sample2 = "inputs/sample_v2.xlsx"  # sheet 'Sheet1'
    diffs = excel_diff(sample1, sample2, sheet1="Data", sheet2="Sheet1", index_col="ID", verbose=True)

    # Save outputs to outputs/ with explicit _index column and without saving the DataFrame index
    out_dir = Path("outputs")
    out_dir.mkdir(parents=True, exist_ok=True)

    # Save 'added'
    if not diffs["added"].empty:
        added = diffs["added"].copy()
        # insert explicit _index column (copy of the index labels)
        added.insert(0, "_index", added.index)
        # Save without DataFrame index to keep CSV clean
        added.to_csv(out_dir / "diff_added.csv", index=False)
        print(f"Saved added rows to {out_dir / 'diff_added.csv'}")
    else:
        print("No added rows to save.")

    # Save 'removed'
    if not diffs["removed"].empty:
        removed = diffs["removed"].copy()
        removed.insert(0, "_index", removed.index)
        removed.to_csv(out_dir / "diff_removed.csv", index=False)
        print(f"Saved removed rows to {out_dir / 'diff_removed.csv'}")
    else:
        print("No removed rows to save.")

    # Save 'modified'
    if not diffs["modified"].empty:
        modified = diffs["modified"].copy()
        # modified currently has _index as its index; reset it to a column
        modified = modified.reset_index()  # this makes _index a normal column
        modified.to_csv(out_dir / "diff_modified.csv", index=False)
        print(f"Saved modified rows to {out_dir / 'diff_modified.csv'}")
    else:
        print("No modified rows to save.")

    print(f"\nSaved CSVs (if any changes) to {out_dir.resolve()}")
