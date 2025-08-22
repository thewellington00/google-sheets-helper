# Google Sheets Python Library

A Python library for interacting with Google Sheets, providing a high-level interface for spreadsheet and worksheet operations. Built on top of the `gspread` library with additional functionality for data manipulation using pandas.

## Features

- **Spreadsheet Management**: Create, delete, rename, and copy worksheets
- **Data Operations**: Read data as pandas DataFrames or dictionaries
- **Named Ranges**: Automatically create named ranges from header rows

## Installation

```bash
pip install -r requirements.txt
```

## Setup

1. **Google Service Account**: You'll need a Google Service Account with access to the Google Sheets API
2. **Environment Variables**: Create a `.env` file in your project root and add your Google Service Account JSON:

```bash
# .env
GOOGLE_SERVICE_ACCOUNT_KEY={"type": "service_account", "project_id": "your-project", "private_key_id": "...", "private_key": "...", "client_email": "...", "client_id": "...", "auth_uri": "https://accounts.google.com/o/oauth2/auth", "token_uri": "https://oauth2.googleapis.com/token", "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs", "client_x509_cert_url": "..."}
```

**Note**: The `.env` file should not be committed to version control. Make sure it's in your `.gitignore`.

## Quick Start

```python
from google_sheets import Spreadsheet

# Initialize spreadsheet
spreadsheet = Spreadsheet("your_spreadsheet_key_here")

# List all worksheets
worksheets = spreadsheet.list_worksheets()
print(f"Available worksheets: {worksheets}")

# Get a worksheet
worksheet = spreadsheet.get_worksheet("Sheet1")

# Read all data as DataFrame
df = worksheet.read_all()
print(df.head())

# Create named ranges from headers
ranges = worksheet.create_named_ranges_from_headers()
print(f"Created ranges: {ranges}")
```

## API Reference

### Spreadsheet Class

#### `__init__(spreadsheet_key: str)`
Initialize the spreadsheet client.

#### `list_worksheets() -> List[str]`
Get a list of all worksheet names in the spreadsheet.

#### `get_worksheet(worksheet_name: str) -> Worksheet`
Get a Worksheet object by name. Validates that the worksheet exists.

#### `create_worksheet(worksheet_name: str, rows: int = 1000, cols: int = 26) -> Worksheet`
Create a new worksheet with formatting (bold headers, frozen first row).

#### `delete_worksheet(name: str) -> None`
Delete a worksheet by name.

#### `worksheet_exists(name: str) -> bool`
Check if a worksheet exists.

#### `rename_worksheet(old_name: str, new_name: str) -> None`
Rename a worksheet.

#### `copy_worksheet(source_name: str, destination_name: str) -> Worksheet`
Copy a worksheet to create a new one.

### Worksheet Class

#### `get_headers() -> List[str]`
Get the header row (first row) as a list of strings.

#### `read_all_records() -> List[Dict[str, Any]]`
Read all records as a list of dictionaries.

#### `read_all() -> pd.DataFrame`
Read all records as a pandas DataFrame with automatic date conversion for columns ending in `_at`.

#### `append_rows(values: List[List[Any]]) -> None`
Append multiple rows to the worksheet.

#### `create_named_ranges_from_headers(data_start_row: int = 2, data_end_row: int = None) -> Dict[str, str]`
Create named ranges for each column based on header names. Handles columns beyond Z (AA, AB, etc.).

## Examples

### Basic Data Operations

```python
# Read data
df = worksheet.read_all()

# Add new data
new_rows = [
    ["John Doe", "john@example.com", "2024-01-15"],
    ["Jane Smith", "jane@example.com", "2024-01-16"]
]
worksheet.append_rows(new_rows)
```

### Working with Multiple Worksheets

```python
# Create a new worksheet
new_sheet = spreadsheet.create_worksheet("Analysis", rows=500, cols=10)

# Copy existing data
copied_sheet = spreadsheet.copy_worksheet("Raw Data", "Raw Data Backup")

# Check if worksheet exists before operating
if spreadsheet.worksheet_exists("Reports"):
    reports = spreadsheet.get_worksheet("Reports")
    df = reports.read_all()
```

### Named Ranges

```python
# Create named ranges for easy formula reference
ranges = worksheet.create_named_ranges_from_headers()

# Example output: {'Name': 'A2:A100', 'Email': 'B2:B100', 'Date': 'C2:C100'}
# Now you can use =SUM(Sales) instead of =SUM(D2:D100) in formulas
```

### Error Handling

```python
try:
    worksheet = spreadsheet.get_worksheet("NonExistent")
except gspread.exceptions.WorksheetNotFound as e:
    print(f"Worksheet not found: {e}")

try:
    new_sheet = spreadsheet.create_worksheet("DuplicateName")
except gspread.exceptions.APIError as e:
    print(f"API Error: {e}")
```

## Requirements

- Python 3.7+
- Dependencies listed in `requirements.txt`
- Google Service Account with Sheets API access

## File Structure

```
your_project/
├── google_sheets.py    # Main library file
├── config.py          # Configuration helper (included)
├── requirements.txt   # Project dependencies
├── .env              # Your environment variables (not committed)
├── .gitignore        # Should include .env
└── README.md         # This file
```

## Contributing

Feel free to submit issues and enhancement requests!

## License

This project is open source. Please check the license file for details.