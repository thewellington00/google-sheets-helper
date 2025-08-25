import json
from typing import List, Dict, Any, Union

import gspread
import pandas as pd

from config import get_google_service_account_key


class Spreadsheet:
    """
    Low-level client for interacting with Google Sheets.

    This class provides basic operations for spreadsheet management
    without the business logic specific to the application.
    """

    def __init__(self, spreadsheet_key: str):
        """
        Initialize the spreadsheet client.

        Args:
            spreadsheet_key: Google Sheets spreadsheet key
        """
        try:
            # Set up authentication
            credentials_json = get_google_service_account_key()
            credentials_dict = json.loads(credentials_json)
            self.sheets_client = gspread.service_account_from_dict(credentials_dict)

            # Open the spreadsheet
            self.spreadsheet_gspread = self.sheets_client.open_by_key(spreadsheet_key)
        except Exception:
            raise

    def list_worksheets(self) -> List[str]:
        """
        List all worksheet (tab) names in the spreadsheet.

        Returns:
            List of worksheet names
        """
        try:
            worksheets = self.spreadsheet_gspread.worksheets()
            return [worksheet.title for worksheet in worksheets]
        except Exception:
            return []

    def get_worksheet(self, worksheet_name: str) -> 'Worksheet':
        """
        Get a Worksheet object by name.

        Args:
            worksheet_name: Name of the worksheet

        Returns:
            Worksheet object

        Raises:
            gspread.exceptions.WorksheetNotFound: If worksheet doesn't exist
        """
        try:
            # Check if worksheet exists first
            if not self.worksheet_exists(worksheet_name):
                raise gspread.exceptions.WorksheetNotFound(f"Worksheet '{worksheet_name}' not found")

            return Worksheet(self, worksheet_name)
        except Exception as e:
            raise

    def create_worksheet(self, worksheet_name: str, rows: int = 1000, cols: int = 26) -> 'Worksheet':
        """
        Create a new worksheet.

        Args:
            worksheet_name: Name for the new worksheet
            rows: Number of rows
            cols: Number of columns

        Returns:
            Worksheet object for the new worksheet

        Raises:
            gspread.exceptions.APIError: If worksheet creation fails
        """
        try:
            # Create worksheet
            worksheet_gspread = self.spreadsheet_gspread.add_worksheet(title=worksheet_name, rows=rows, cols=cols)

            # Make the first row bold, and add a freeze row for the top row
            worksheet_gspread.format('A1:Z1', {'textFormat': {'bold': True}})
            worksheet_gspread.freeze(rows=1)

            # Return our Worksheet wrapper
            return Worksheet(self, worksheet_name)
        except Exception as e:
            raise

    def delete_worksheet(self, name: str) -> None:
        """
        Delete a worksheet by name.

        Args:
            name: Name of the worksheet to delete

        Raises:
            gspread.exceptions.WorksheetNotFound: If worksheet doesn't exist
        """
        try:
            worksheet = self.get_worksheet(name)
            self.spreadsheet_gspread.del_worksheet(worksheet.worksheet_gspread)
        except Exception as e:
            raise

    def worksheet_exists(self, name: str) -> bool:
        """
        Check if a worksheet with the given name exists.

        Args:
            name: Name of the worksheet to check

        Returns:
            True if the worksheet exists, False otherwise
        """
        try:
            worksheets = self.list_worksheets()
            return name in worksheets
        except Exception:
            return False

    def rename_worksheet(self, old_name: str, new_name: str) -> None:
        """
        Rename a worksheet.

        Args:
            old_name: Current name of the worksheet
            new_name: New name for the worksheet

        Raises:
            gspread.exceptions.WorksheetNotFound: If worksheet doesn't exist
        """
        try:
            worksheet = self.get_worksheet(old_name)
            worksheet.worksheet_gspread.update_title(new_name)
        except Exception as e:
            raise

    def copy_worksheet(self, source_name: str, destination_name: str) -> 'Worksheet':
        """
        Copy a worksheet to a new worksheet.

        Args:
            source_name: Name of the source worksheet
            destination_name: Name for the new worksheet

        Returns:
            Worksheet object for the new worksheet

        Raises:
            gspread.exceptions.WorksheetNotFound: If source worksheet doesn't exist
            gspread.exceptions.APIError: If worksheet copying fails
        """
        try:
            # Get source worksheet
            source_worksheet = self.get_worksheet(source_name)

            # Create new worksheet
            new_worksheet = self.create_worksheet(destination_name)

            # Copy all data from source to destination
            data = source_worksheet.worksheet_gspread.get_all_values()
            if data:
                new_worksheet.worksheet_gspread.update(data)

            # Copy formatting (basic)
            # Note: gspread has limited formatting copy capabilities

            return new_worksheet
        except Exception as e:
            raise


class Worksheet:
    """
    Higher-level class for working with tabular data in Google Sheets.

    This class provides methods for CRUD operations on worksheet data,
    handling the conversion between DataFrame objects and Google Sheets.
    """

    def __init__(self, spreadsheet: Spreadsheet, sheet_name: str):
        """
        Initialize a worksheet.

        Args:
            spreadsheet: Spreadsheet instance
            sheet_name: Name of the worksheet
        """
        self.spreadsheet = spreadsheet
        self.sheet_name = sheet_name
        self._worksheet_gspread = None  # Cached worksheet object

    @property
    def worksheet_gspread(self):
        """
        Get the worksheet, with caching for performance.

        Returns:
            gspread worksheet object
        """
        if self._worksheet_gspread is None:
            # Call gspread directly, not our wrapper method
            self._worksheet_gspread = self.spreadsheet.spreadsheet_gspread.worksheet(self.sheet_name)
        return self._worksheet_gspread

    def get_headers(self) -> List[str]:
        """
        Get the header row from the worksheet.

        Returns:
            List of column headers
        """
        return self.worksheet_gspread.row_values(1)

    def read_all_records(self) -> List[Dict[str, Any]]:
        """
        Read all records from the worksheet as a list of dictionaries.

        Returns:
            List of dictionaries, each representing a row
        """
        return self.worksheet_gspread.get_all_records()

    def read_all(self) -> pd.DataFrame:
        """
        Read all records from the worksheet as a DataFrame.

        Returns:
            DataFrame containing all records
        """
        records = self.read_all_records()
        if records:
            df = pd.DataFrame(records)

            # Convert string date columns to datetime
            for col in df.columns:
                if col.endswith('_at') and df[col].dtype == 'object':
                    try:
                        df[col] = pd.to_datetime(df[col])
                    except:
                        pass  # Keep as string if conversion fails

            return df
        else:
            return pd.DataFrame(columns=self.get_headers())

    def append_rows(self, values: List[List[Any]]) -> None:
        """
        Append multiple rows to the worksheet.

        Args:
            values: List of row values to append
        """
        if values:
            self.worksheet_gspread.append_rows(values)

    def create_named_ranges_from_headers(self, data_start_row: int = 2, data_end_row: int = None) -> Dict[str, str]:
        """
        Create named ranges for each column based on the header row.

        This creates a named range for each column that includes all data rows,
        making it easy to reference columns by name in formulas.

        Args:
            data_start_row: First row containing data (default: 2, assuming row 1 is headers)
            data_end_row: Last row containing data (default: None, will use last row with data)

        Returns:
            Dictionary mapping header names to their A1 notation ranges

        Raises:
            Exception: If there are no headers or if named range creation fails
        """
        try:
            headers = self.get_headers()
            if not headers:
                raise ValueError("No headers found in the worksheet")

            # If data_end_row not specified, find the last row with data
            if data_end_row is None:
                all_values = self.worksheet_gspread.get_all_values()
                data_end_row = len(all_values) if all_values else data_start_row

            created_ranges = {}
            for col_index, header in enumerate(headers):
                if header.strip():  # Only create ranges for non-empty headers
                    # Use gspread's utils to convert column index to letter notation
                    # col_index is 0-based, but gspread expects 1-based
                    col_num = col_index + 1
                    range_notation = f"{gspread.utils.rowcol_to_a1(data_start_row, col_num)}:{gspread.utils.rowcol_to_a1(data_end_row, col_num)}"

                    # Clean the header name for use as a named range
                    # Remove spaces, special characters, and make it a valid name
                    range_name = header.strip().replace(' ', '_').replace('-', '_')
                    range_name = ''.join(c for c in range_name if c.isalnum() or c == '_')

                    # Ensure the name starts with a letter
                    if range_name and not range_name[0].isalpha():
                        range_name = f"col_{range_name}"

                    if range_name:  # Only create if we have a valid name
                        try:
                            self.worksheet_gspread.define_named_range(
                                name=range_notation,
                                range_name=range_name
                            )
                            created_ranges[header] = range_notation
                        except Exception as e:
                            print(f"Warning: Could not create named range for '{header}': {e}")
                            continue

            return created_ranges

        except Exception as e:
            raise Exception(f"Failed to create named ranges: {e}")

    def delete_all_named_ranges(self) -> Dict[str, Any]:
        """
        Delete all named ranges in the spreadsheet.

        This uses gspread's delete_named_range(named_range_id) for each
        named range returned by Spreadsheet.list_named_ranges().

        Returns:
            A dictionary with:
                - deleted_ids: list of successfully deleted named range IDs
                - responses: mapping of named range ID -> API response body
                - errors: mapping of named range ID -> error message (if deletion failed)
        """
        try:
            named_ranges = self.spreadsheet.spreadsheet_gspread.list_named_ranges() or []
            responses: Dict[str, Any] = {}
            errors: Dict[str, str] = {}

            for nr in named_ranges:
                # Be defensive about the structure returned by gspread
                nr_id = None
                if isinstance(nr, dict):
                    nr_id = (
                            nr.get("namedRangeId")
                            or nr.get("named_range_id")
                            or nr.get("id")
                            or nr.get("nameId")
                    )
                else:
                    nr_id = (
                            getattr(nr, "namedRangeId", None)
                            or getattr(nr, "id", None)
                    )

                if not nr_id:
                    # Skip entries we cannot identify
                    continue

                try:
                    resp = self.worksheet_gspread.delete_named_range(nr_id)
                    responses[nr_id] = resp
                except Exception as e:
                    errors[nr_id] = str(e)

            return {
                "deleted_ids": list(responses.keys()),
                "responses": responses,
                "errors": errors,
            }
        except Exception as e:
            raise Exception(f"Failed to delete named ranges: {e}")

    def cross_join_ranges_to_clipboard(self, range_a: str, range_b: str) -> List[List[Any]]:
        """
        Create a sorted cross join of the values contained in two ranges and copy it to the clipboard.

        Args:
            range_a: First range in A1 notation (e.g., "A2:A10" or "B2:D2")
            range_b: Second range in A1 notation

        Behavior:
            - Flattens both ranges to 1D lists (row-major), removing empty cells.
            - Produces all pairs [a, b] for a in range_a_values and b in range_b_values.
            - Sorts the pairs ascending by the first column, then the second.
              Sorting is numeric-aware when possible, otherwise case-insensitive textual.
            - Copies the resulting n x 2 table to the clipboard as TSV for easy pasting into Google Sheets.

        Returns:
            The sorted list of pairs with shape n x 2.
        """
        def _fetch_1d_list(a1: str) -> List[str]:
            # Returns flattened non-empty values (as strings) from the provided A1 range.
            values = self.worksheet_gspread.get(a1) or []
            flat = [("" if c is None else str(c)) for row in values for c in row]
            return [v for v in (s.strip() for s in flat) if v != ""]

        def _sort_key(v: Any):
            s = "" if v is None else str(v).strip()
            if s == "":
                return 0, ""
            try:
                return 1, float(s)
            except Exception:
                return 2, s.casefold()

        list_a = _fetch_1d_list(range_a)
        list_b = _fetch_1d_list(range_b)

        pairs: List[List[Any]] = []
        if list_a and list_b:
            pairs = [[a, b] for a in list_a for b in list_b]
            pairs.sort(key=lambda p: (_sort_key(p[0]), _sort_key(p[1])))

        # Copy via pandas to clipboard as TSV (n x 2) without headers or index
        df = pd.DataFrame(pairs)
        df.to_clipboard(index=False, header=False)
        return pairs
