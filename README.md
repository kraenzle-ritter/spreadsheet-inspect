# Spreadsheet Inspect

[![Tests](https://github.com/kraenzle-ritter/spreadsheet-inspect/actions/workflows/tests.yml/badge.svg)](https://github.com/kraenzle-ritter/spreadsheet-inspect/actions/workflows/tests.yml)
[![codecov](https://codecov.io/gh/kraenzle-ritter/spreadsheet-inspect/branch/main/graph/badge.svg)](https://codecov.io/gh/kraenzle-ritter/spreadsheet-inspect)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PHP Version](https://img.shields.io/packagist/php-v/kraenzle-ritter/spreadsheet-inspect)](https://packagist.org/packages/kraenzle-ritter/spreadsheet-inspect)
[![Downloads](https://img.shields.io/github/downloads/kraenzle-ritter/spreadsheet-inspect/total)](https://github.com/kraenzle-ritter/spreadsheet-inspect/releases)

A CLI tool to inspect Excel (.xlsx, .xls) and LibreOffice (.ods) spreadsheet files. Analyze sheet structures, column statistics, find cross-sheet references, and extract embedded images.

## Features

- 📊 **Sheet Analysis** – List all sheets, view column headers, row counts, and fill rates
- 📈 **Value Statistics** – See distinct values per column with occurrence counts
- 🔗 **Cross-Sheet Reference Check** – Find where values from one sheet appear in others
- 🖼️ **Image Inspection** – Count and extract embedded images/drawings
- 📁 **Multiple Formats** – Supports Excel (.xlsx, .xls) and LibreOffice Calc (.ods)

## Installation

```bash
git clone https://github.com/kraenzle-ritter/spreadsheet-inspect.git
cd spreadsheet-inspect
composer install
```

## Usage

### List all sheets

```bash
php inspect spreadsheet myfile.xlsx --sheets
```

### Analyze a specific sheet

```bash
php inspect spreadsheet myfile.xlsx --sheet=1
# or by name
php inspect spreadsheet myfile.xlsx --sheet="Sheet Name"
```

**Output:**
```
## Available sheets
- **[1]** `Products`
- **[2]** `Categories`

# Sheet `Products` (Index: 1)

## Sheet statistics
- **Rows** (excluding header): `150`

### `ProductID`
- **Filled**: `150 / 150` (100%)
- **Distinct**: `150`

### `Category`
- **Filled**: `148 / 150` (98.67%)
- **Distinct**: `12`

  Values:
  - `Electronics` (45)
  - `Clothing` (32)
  ...
```

### Analyze images in a sheet

```bash
php inspect spreadsheet myfile.xlsx --sheet=1 --images
```

### Extract images to a directory

```bash
php inspect spreadsheet myfile.xlsx --sheet=1 --extract-images=./images
```

### Cross-sheet reference check

Find where values from a column appear in other sheets:

```bash
php inspect spreadsheet myfile.xlsx --sheet=1 --column=ProductID
```

Check against a specific target sheet:

```bash
php inspect spreadsheet myfile.xlsx --sheet=1 --column=ProductID --cross-sheet=2
```

Compare only against a specific column in target sheets:

```bash
php inspect spreadsheet myfile.xlsx --sheet=1 --column=ProductID --target-column=ID
```

### Options

| Option | Description |
|--------|-------------|
| `--sheets` | List all sheet names only |
| `--sheet=` | Inspect a specific sheet (by index or name) |
| `--column=` | Cross-search for values from this column |
| `--cross-sheet=` | Only check this target sheet |
| `--target-column=` | Only compare against this column in target sheets |
| `--images` | Count and list images in the sheet |
| `--extract-images=` | Extract images to specified directory |
| `--output=` | Output format: `console` (default), `html`, `pdf` |
| `--output-file=` | Output file path (required for html/pdf) |
| `--debug` | Show detailed matching values |
| `--memory=` | Memory limit in MB (default: 2000) |

### Export to HTML or PDF

Generate a styled HTML report:

```bash
php inspect spreadsheet myfile.xlsx --sheet=1 --output=html --output-file=report.html
```

Generate a PDF report:

```bash
php inspect spreadsheet myfile.xlsx --sheet=1 --output=pdf --output-file=report.pdf
```

## Supported Formats

- Microsoft Excel: `.xlsx`, `.xls`
- LibreOffice Calc: `.ods`

## Requirements

- PHP 8.3+
- Composer

## Testing

```bash
./vendor/bin/pest
```

## Built With

- [Laravel Zero](https://laravel-zero.com/) – Micro-framework for console applications
- [PhpSpreadsheet](https://phpspreadsheet.readthedocs.io/) – Library for reading/writing spreadsheet files
- [FastExcel](https://github.com/rap2hpoutre/fast-excel) – Fast Excel import/export
- [Dompdf](https://github.com/dompdf/dompdf) – HTML to PDF converter

## License

MIT License. See [LICENSE](LICENSE) for details.
