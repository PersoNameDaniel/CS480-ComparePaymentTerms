# Payment Terms Comparison Tool

A Python console application that reads payment terms from an Excel file, compares them with QuickBooks Desktop, and automatically synchronizes differences.

## Features

- **Read Excel Payment Terms**: Reads payment terms from the `payment_terms` sheet in an Excel file
- **Read QuickBooks Payment Terms**: Queries existing payment terms from QuickBooks Desktop
- **Compare Terms**: Compares terms by ID (stored in QB's `StdDiscountDays` field)
  - Reports terms with same ID but different names
  - Reports terms only in Excel (to be added to QB)
  - Reports terms only in QuickBooks
  - Counts matching terms (same ID and name)
- **Auto-Sync**: Automatically adds terms from Excel to QuickBooks if they don't exist

## Installation

```bash
# Install dependencies
poetry install
```

## Usage

```bash
# Run the comparison tool
python run_comparison.py <path_to_excel_file>

# Example
python run_comparison.py C:\datasets\paymentterms.xlsx
```

### Using Poetry

```bash
poetry run python run_comparison.py C:\datasets\paymentterms.xlsx
```

## Excel File Format

The Excel file must contain a sheet named `payment_terms` with the following structure:

| Name         | ID |
|--------------|----|
| Net 30       | 1  |
| Net 45       | 2  |
| Net 60       | 6  |
| down payment | 8  |
| EOM+60       | 11 |

- **Column A (Name)**: Payment term name (string)
- **Column B (ID)**: Unique identifier (integer)
- **Row 1**: Headers (will be skipped)

## QuickBooks Integration

The tool uses QuickBooks' `StdDiscountDays` field to store the Excel term ID. This allows matching terms between Excel and QuickBooks by ID rather than name.

### QuickBooks Setup
1. QuickBooks Desktop 2021 must be running
2. A company file must be open
3. You may need to grant permission for the application to access QuickBooks on first run

## Output Example

```
Starting payment terms comparison for: C:\datasets\paymentterms.xlsx

Found 5 payment terms in Excel
Reading payment terms from QuickBooks...
Found 3 payment terms in QuickBooks

=== Comparison Results ===

Matching terms (same ID and name): 2

Terms with same ID but different names (1):
  ID 2: Excel='Net 45' vs QB='Different Name'

Terms in QuickBooks but not in Excel (0):
No terms found only in QuickBooks

Adding 2 new terms to QuickBooks...
  - Net 60 (ID: 6)
  - down payment (ID: 8)

Successfully created 2 terms in QuickBooks

=== Summary ===
Completed successfully!
- Matching terms: 2
- Terms with same ID but different names: 1
- Terms only in Excel (added to QB): 2
- Terms only in QB: 0
```

## Project Structure

```
xlsx_reader/
├── __init__.py           # Package exports
└── excel_processor.py    # Core logic: Excel reading, QB integration, comparison

tests/
├── __init__.py
└── test_excel_processor.py  # Comprehensive tests

run_comparison.py         # Main entry point script
```

## Dependencies

- **Python 3.12+**
- **openpyxl**: For reading Excel files
- **pywin32**: For QuickBooks COM API integration
- **pytest**: For testing (dev dependency)

## Running Tests

```bash
poetry run pytest
```

Tests will create temporary Excel files automatically for testing.

## API Reference

For programmatic usage, see the functions in `xlsx_reader.excel_processor`:

- `read_payment_terms(file_path: str) -> list[PaymentTerm]`
- `get_qb_payment_terms() -> list[PaymentTerm]`
- `compare_payment_terms(excel_terms, qb_terms) -> TermComparison`
- `process_payment_terms(file_path: str) -> TermComparison` (main workflow)

See the module docstrings for detailed parameter and return value documentation.
