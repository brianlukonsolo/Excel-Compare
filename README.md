# Excel Compare - A Sequential Excel Diff Tool

A Python tool for comparing multiple spreadsheet files sequentially and generating human-readable diff reports.

## Features

- **Sequential Comparisons**: Compare any number of files in sequence (file1→file2, file2→file3, ...)
- **Multi-Format Support**: Works with `.xlsx`, `.xls`, `.csv`, and `.tsv` files
- **Flexible Configuration**: Configure via `.env` file, environment variables, or CLI arguments
- **Smart Row Matching**: Use a key column (e.g., "ID") to match rows across files, or compare positionally
- **Column Filtering**: Optionally compare only specific columns
- **Custom Header Rows**: Specify which row contains column headers (0-indexed)
- **Detailed Change Tracking**: Detects additions, removals, and modifications with cell-level diffs
- **Multiple Output Formats**: Generate text reports and/or CSV files per comparison
- **Source Metadata**: Every change includes source file and sheet information

## Installation

### Prerequisites

- Python 3.7+
- pip

### Install Dependencies

```bash
pip install -r requirements.txt
```

Required packages:
- `pandas` - Data manipulation and analysis
- `python-dotenv` - Environment variable management
- `openpyxl` - Modern Excel file support (.xlsx, .xlsm)
- `xlrd==1.2.0` - Legacy Excel file support (.xls)

## Quick Start

1. **Create input directory and add your files:**
   ```bash
   mkdir inputs
   # Copy your Excel/CSV files to inputs/
   ```

2. **Create a `.env` file** (optional - see Configuration section):
   ```bash
   INPUTS=inputs/version1.xlsx,inputs/version2.xlsx
   INDEX_COL=ID
   ```

3. **Run the comparison:**
   ```bash
   python main.py
   ```

4. **Check the output:**
   - Report: `outputs/excel_diff_report.txt` (or your configured filename)
   - Optional CSVs: `outputs/diff_01_added_*.csv`, etc.

## Configuration

### Configuration Priority

Settings are resolved in this order (highest to lowest priority):
1. **CLI arguments** (highest priority)
2. **Environment variables** (from `.env` or system)
3. **Default values** (lowest priority)

### Available Settings

| Setting | Default | Description |
|---------|---------|-------------|
| `INPUTS` | `inputs/sample_v1.xlsx,inputs/sample_v2.xlsx` | Comma-separated list of input files in order |
| `SHEETS` | _(empty)_ | Comma-separated sheet names; blank entries use first sheet |
| `INDEX_COL` | `ID` | Column name to use as key for matching rows |
| `COLUMNS` | _(all)_ | Comma-separated list of columns to compare; blank = all |
| `COLUMN_HEADER_ROW_INDEX` | `0` | 0-indexed row number where column headers are located |
| `OUTPUT_DIR` | `outputs` | Directory for generated reports and CSVs |
| `REPORT_FILENAME` | `excel_diff_report.txt` | Name of the text report file |
| `EXPORT_CSV` | `true` | Export individual CSV files per comparison |
| `PRINT_TERMINAL` | `true` | Print results to terminal |
| `INCLUDE_ADDITIONS` | `true` | Include added rows in report |
| `INCLUDE_MODIFICATIONS` | `true` | Include modified rows in report |
| `INCLUDE_REMOVALS` | `true` | Include removed rows in report |

### Example `.env` File

```env
# Sequential Excel diff configuration
INPUTS=inputs/sample_v1.xlsx,inputs/sample_v2.xlsx,inputs/sample_v3.xlsx
SHEETS=Data,,Other
INDEX_COL=ID
COLUMNS=Name,Price,Quantity
COLUMN_HEADER_ROW_INDEX=0
EXPORT_CSV=true
PRINT_TERMINAL=false
INCLUDE_ADDITIONS=true
INCLUDE_MODIFICATIONS=true
INCLUDE_REMOVALS=true
OUTPUT_DIR=outputs
REPORT_FILENAME=changes.txt
```

## Usage Examples

### Basic Usage

Compare two Excel files using defaults:
```bash
python main.py --inputs "data_v1.xlsx,data_v2.xlsx"
```

### Specify Sheets

Compare specific sheets from each file:
```bash
python main.py --inputs "file1.xlsx,file2.xlsx" --sheets "Sheet1,Sheet2"
```

Leave a sheet entry blank to use the first sheet:
```bash
python main.py --inputs "file1.xlsx,file2.xlsx,file3.xlsx" --sheets "Data,,Summary"
# Uses "Data" from file1, first sheet from file2, "Summary" from file3
```

### Use Custom Index Column

Match rows by a key column (e.g., "ProductID"):
```bash
python main.py --index-col "ProductID"
```

### Compare Specific Columns Only

Only compare specific columns:
```bash
python main.py --columns "Name,Price,Stock"
```

### Headers on Different Row

If column headers are on row 3 (0-indexed = 2):
```bash
python main.py --column-header-row-index 2
```

### Compare CSV Files

Works the same way:
```bash
python main.py --inputs "data1.csv,data2.csv,data3.csv" --index-col "id"
```

### Mix File Types

Compare different formats together:
```bash
python main.py --inputs "v1.xlsx,v2.csv,v3.tsv"
```

### Exclude Certain Changes

Only show modifications, exclude additions and removals:
```bash
python main.py --include-additions false --include-removals false
```

### Export Only (No Terminal Output)

Generate report quietly:
```bash
python main.py --print-terminal false
```

### Skip CSV Export

Generate only the text report:
```bash
python main.py --export-csv false
```

## Understanding the Output

### Text Report Structure

```
Sequential Excel Diff Report
Inputs: inputs/sample_v1.xlsx, inputs/sample_v2.xlsx
Columns compared: (all)
Index column: ID
--------------------------------------------------------------------------------

================================================================================
Comparison 1: sample_v1.xlsx:Data  --->  sample_v2.xlsx:Data
--------------------------------------------------------------------------------

#### [ ADDITIONS ] (rows present in [sample_v2.xlsx] but not [sample_v1.xlsx])
ID,Name,Price,_source
105,New Product,99.99,sample_v2.xlsx:Data

#### [ REMOVALS ] (rows present in [sample_v1.xlsx] but not [sample_v2.xlsx])
ID,Name,Price,_source
101,Old Product,49.99,sample_v1.xlsx:Data

#### [ MODIFIED ] (rows present in BOTH [sample_v1.xlsx] and [sample_v2.xlsx] but with changed cells)
_index,Price,_old_source,_new_source
102,'19.99' → '24.99',sample_v1.xlsx:Data,sample_v2.xlsx:Data
```

### CSV Exports

When `EXPORT_CSV=true`, individual CSV files are generated:
- `diff_01_added_sample_v2.csv` - Rows added in the second file
- `diff_01_removed_sample_v1.csv` - Rows removed from the first file
- `diff_01_modified_sample_v1_to_sample_v2.csv` - Modified rows with changes

### Change Format

Modified cells show old → new format:
```
'19.99' → '24.99'
'John' → 'Jane'
<NA> → 'New Value'
```

## How It Works

### Row Matching

**With INDEX_COL**: Rows are matched by the specified key column value (e.g., matching "ID" values)
- ✅ Handles rows in different order
- ✅ Detects when a row is truly added or removed
- ⚠️ Column must exist in both files

**Without INDEX_COL**: Rows are matched by position (row 1 compares to row 1, etc.)
- ✅ Simple positional comparison
- ⚠️ Row order matters

### Column Handling

- If a column exists in one file but not the other, it's included in the comparison with `<NA>` values
- Use `COLUMNS` to restrict comparison to specific columns
- Columns that don't exist in a file are gracefully skipped

### Sheet Selection

- Specify sheets per file using `SHEETS` parameter
- Blank entries default to the first sheet
- If a specified sheet doesn't exist, the first sheet is used with a warning

### Sequential Comparisons

For 3+ files, comparisons are made sequentially:
```
file1 → file2  (Comparison 1)
file2 → file3  (Comparison 2)
file3 → file4  (Comparison 3)
```

Each comparison is independent and reported separately.

## Error Handling

The tool provides helpful error messages for common issues:

- **Missing files**: Clear error with file path
- **Missing dependencies**: Install instructions for `openpyxl` or `xlrd`
- **Unsupported formats**: Lists supported extensions
- **Missing sheets**: Falls back to first sheet with warning
- **Engine failures**: Automatically tries alternative Excel engines

## Advanced Tips

### Large Files

For very large files:
- Disable terminal output: `PRINT_TERMINAL=false`
- Disable CSV export if not needed: `EXPORT_CSV=false`
- Use column filtering to reduce memory: `COLUMNS=col1,col2,col3`

### Mixed Data Types

The tool handles different data types gracefully:
- Numbers, strings, booleans, dates
- Missing values (`NaN`, empty cells)
- Uses pandas' smart type detection

### Automation

Use in scripts or CI/CD pipelines:
```bash
#!/bin/bash
python main.py \
  --inputs "baseline.xlsx,current.xlsx" \
  --index-col "ID" \
  --export-csv false \
  --print-terminal false

# Check if report was generated
if [ -f "outputs/excel_diff_report.txt" ]; then
    echo "Diff report generated successfully"
fi
```

## Troubleshooting

### "Missing dependency 'openpyxl'"
```bash
pip install openpyxl
```

### "Missing dependency 'xlrd'"
For `.xls` files (legacy Excel):
```bash
pip install xlrd==1.2.0
```

### "KeyError: '['ID'] not in index'"
The specified `INDEX_COL` doesn't exist in one or both files. Check:
- Column name spelling and case
- Headers are on the correct row (use `COLUMN_HEADER_ROW_INDEX`)

### "At least two input files are required"
Provide at least 2 files to compare:
```bash
python main.py --inputs "file1.xlsx,file2.xlsx"
```

## CLI Reference

```
usage: main.py [-h] [--inputs INPUTS] [--sheets SHEETS] [--index-col INDEX_COL]
               [--columns COLUMNS] [--column-header-row-index COLUMN_HEADER_ROW_INDEX]
               [--output-dir OUTPUT_DIR] [--report REPORT] [--export-csv EXPORT_CSV]
               [--print-terminal PRINT_TERMINAL] [--include-additions INCLUDE_ADDITIONS]
               [--include-modifications INCLUDE_MODIFICATIONS]
               [--include-removals INCLUDE_REMOVALS]

optional arguments:
  -h, --help            Show this help message and exit
  --inputs, -i INPUTS   Comma-separated list of input files in order
  --sheets, -s SHEETS   Comma-separated list of sheet names (same order)
  --index-col INDEX_COL Index column name to use as key
  --columns, -c COLUMNS Comma-separated list of column names to compare
  --column-header-row-index COLUMN_HEADER_ROW_INDEX
                        0-indexed row number where column headers are located
  --output-dir OUTPUT_DIR Output directory
  --report REPORT       Report filename
  --export-csv EXPORT_CSV Export CSVs (true/false)
  --print-terminal PRINT_TERMINAL Print to terminal (true/false)
  --include-additions INCLUDE_ADDITIONS Include additions in report (true/false)
  --include-modifications INCLUDE_MODIFICATIONS Include modifications in report (true/false)
  --include-removals INCLUDE_REMOVALS Include removals in report (true/false)
```

## License

This project is provided as-is for use in comparing spreadsheet data.

## Contributing

Feel free to submit issues or pull requests for improvements.
