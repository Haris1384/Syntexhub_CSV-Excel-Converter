# CSV → Excel Converter

A Python command-line tool that reads a CSV file, cleans and normalizes the data, and exports it to a formatted Excel (`.xlsx`) workbook using **pandas** and **openpyxl**.

Built for **Syntecxhub Internship – Project 1**. Useful for data preparation and reporting automation.

## Features

- **Reads messy CSVs**: auto-detects the delimiter (`,` `;` tab `|`) and falls back across encodings (UTF-8, cp1252, Latin-1)
- **Cleans and normalizes**
  - Header names become `snake_case` (`" Order Date "` → `order_date`)
  - Trims extra whitespace in every cell
  - Treats `NA`, `N/A`, `null`, `none`, `-`, `?` and blank cells as missing
  - Drops fully empty rows and columns; optionally drops duplicate rows
  - Converts numeric text such as `"$1,200.00"` into real numbers
- **Missing values**: keep them, drop the rows, or fill them (text → `Not Available`, numbers → zero, mean or median)
- **Date parsing**: auto-detects date columns or takes the ones you list; handles mixed formats, with `--dayfirst` for DD/MM/YYYY
- **Column renames**: `--rename old:new`
- **Formatted Excel output**: coloured bold header, frozen header row, filters, auto column widths and date formatting
- **Cleaning Report sheet**: lists rows read and written, rows dropped, renamed columns, detected types and missing values filled
- **Logging and errors**: `-v` / `-q` / `--log-file`, with clear messages for missing, empty or unreadable files and bad options (exit code 1)

## Installation

```bash
git clone https://github.com/HammadDev-Z/Syntecxhub_CSV_Excel_Converter.git
cd Syntecxhub_CSV_Excel_Converter
pip install -r requirements.txt
```

## Usage

```bash
# Basic: writes sales_raw.xlsx next to the input
python csv_to_excel.py -i sample_data/sales_raw.csv

# Custom output path, remove duplicates, DD/MM dates
python csv_to_excel.py -i sample_data/sales_raw.csv -o output/sales_clean.xlsx --drop-duplicates --dayfirst

# Rename columns and choose the date columns yourself
python csv_to_excel.py -i data.csv -o report.xlsx -r "cust_nm:customer" -r "amt:amount" --date-columns order_date

# Drop rows with missing values and write a log file
python csv_to_excel.py -i data.csv -o report.xlsx --missing drop --log-file logs/convert.log
```

### CLI options

| Flag | Description | Default |
|------|-------------|---------|
| `-i, --input` | Input CSV file **(required)** | – |
| `-o, --output` | Output `.xlsx` file | input name with `.xlsx` |
| `-s, --sheet-name` | Data worksheet name | `Cleaned Data` |
| `-d, --delimiter` | CSV delimiter (`tab` for TSV) | auto-detect |
| `-e, --encoding` | File encoding | auto (utf-8 → cp1252 → latin-1) |
| `-r, --rename OLD:NEW` | Rename columns (repeatable / comma-separated) | – |
| `--keep-headers` | Don't convert headers to snake_case | off |
| `--date-columns COL…` | Columns to parse as dates | auto-detect |
| `--no-auto-dates` | Turn off date auto-detection | off |
| `--dayfirst` | Read ambiguous dates as DD/MM/YYYY | off |
| `--date-format` | Excel date display format | `yyyy-mm-dd` |
| `--missing {keep,drop,fill}` | How to handle missing values | `fill` |
| `--fill-text` | Fill value for text cells | `Not Available` |
| `--fill-numeric {zero,mean,median,none}` | Fill strategy for number cells | `zero` |
| `--drop-duplicates` | Remove duplicate rows | off |
| `--no-report` | Leave out the *Cleaning Report* sheet | off |
| `-v, --verbose` / `-q, --quiet` | More / less logging | info |
| `--log-file` | Also write logs to a file | – |

### Example

Input (`sample_data/sales_raw.csv`) has messy headers, `$1,200.00` style numbers, mixed date formats, `NA`/`null` tokens, a blank row and a duplicate row:

```text
 Order ID ,Customer Name,Order Date,Region,Product, Quantity ,Unit Price,Total ($),Ship Date,Notes
1001,  Alice Khan ,2026-01-05,North,Laptop,2,"$1,200.00","2,400.00",2026-01-08,
1002,Bilal Ahmed,01/07/2026,South,Mouse,5,25.5,127.5,N/A,Gift wrap
...
```

Console output:

```text
[INFO] Read 'sales_raw.csv' (encoding=utf-8-sig, delimiter=','): 8 rows x 10 columns
[INFO] Auto-detected date column: 'order_date'
[INFO] Auto-detected date column: 'ship_date'
[INFO] Cleaning done: 8 -> 6 rows, 4 numeric, 2 date column(s)
[INFO] Saved Excel file: .../output/sales_clean.xlsx
```

Errors:

```text
$ python csv_to_excel.py -i missing.csv
[ERROR] Input file not found: missing.csv
$ python csv_to_excel.py -i empty.csv
[ERROR] Input file is empty: empty.csv
```

## Using it from Python

```python
from pathlib import Path
from csv_to_excel import CleanOptions, convert

report = convert(Path("data.csv"), Path("data.xlsx"),
                 CleanOptions(renames={"cust_nm": "customer"}, drop_duplicates=True))
print(report.rows_in, "->", report.rows_out)
```

## Running the tests

```bash
pip install -r requirements-dev.txt
python -m pytest -q
```

## Project structure

```
haris/
├── csv_to_excel.py          # Converter: read → clean → write + CLI
├── sample_data/
│   ├── sales_raw.csv        # Messy example input
│   └── semicolon.csv        # Semicolon-delimited example
├── tests/
│   └── test_csv_to_excel.py # pytest suite
├── requirements.txt
├── requirements-dev.txt
└── README.md
```

## Tech stack

Python 3.9+, pandas, openpyxl, argparse, logging, pytest
