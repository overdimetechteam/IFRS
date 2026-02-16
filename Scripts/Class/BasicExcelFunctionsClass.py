"""
Portfolio Automation Script using xlwings
Automates the process of:
1. Clearing data range A4:AB in Portfolio sheet of xlsb file
2. Reading CSV file columns A-AA and pasting to Portfolio sheet starting at B4
3. Filling column A with today's date
4. Clearing range A2:E in 1.DPD sheet
5. Reading xlsx file and pasting first 5 columns to 1.DPD sheet starting at A2
6. Saving as "Updated ECL Portfolio" in xlsb format
"""
import pandas as pd
import xlwings as xw
from datetime import datetime
import os
import glob
import shutil
from typing import Optional, Tuple, List
import re
import calendar


class ExcelPortfolioAutomation:
    """
    Excel Portfolio Automation class using xlwings for .xlsb file manipulation
    """

    def __init__(self, workbook_path: str, visible: bool = True):
        """
        Initialize the automation class

        Args:
            workbook_path: Path to the .xlsb Excel file
            visible: Whether to show Excel application (default: False)
        """
        self.workbook_path = workbook_path
        self.visible = visible
        self.app = None
        self.workbook = None

    def __enter__(self):
        """Context manager entry - opens the workbook"""
        self.open_workbook()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit - closes the workbook"""
        self.close_workbook()

    def open_workbook(self):
        """Open the Excel workbook"""
        print(f"Opening workbook: {self.workbook_path}")
        self.app = xw.App(visible=self.visible)
        self.workbook = self.app.books.open(self.workbook_path)
        print(f"    Workbook opened successfully")

    def close_workbook(self, save: bool = False):
        """
        Close the Excel workbook

        Args:
            save: Whether to save before closing (default: False)
        """
        if self.workbook:
            if save:
                self.workbook.save()
                print(f"    Workbook saved")
            self.workbook.close()
            print(f"    Workbook closed")
        if self.app:
            self.app.quit()

    def clear_range(self, sheet_name: str, range_address: str) -> None:
        """
        Clear data in a specific range of a worksheet

        Args:
            sheet_name: Name of the worksheet
            range_address: Range address (e.g., 'A4:AB10000')
        """
        print(f"   Clearing range {range_address} in sheet '{sheet_name}'...")
        sheet = self.workbook.sheets[sheet_name]
        sheet.range(range_address).clear_contents()
        print(f"    Range cleared successfully")

    def clear_range_dynamic(self, sheet_name: str, start_cell: str, end_column: str) -> None:
        """
        Clear data from start_cell to the last row with data in specified columns

        Args:
            sheet_name: Name of the worksheet
            start_cell: Starting cell (e.g., 'A4', 'AX3', 'AB10')
            end_column: Last column letter (e.g., 'AB', 'AX')
        """
        sheet = self.workbook.sheets[sheet_name]

        # Extract column letters and row number
        start_column = ''.join(filter(str.isalpha, start_cell))
        start_row = int(''.join(filter(str.isdigit, start_cell)))

        print(f"   Analyzing range starting from {start_cell} in sheet '{sheet_name}'...")

        try:
            # Find the last row with data in the start column
            # Check from the row after start_cell
            check_cell = sheet.range(f'{start_column}{start_row + 1}')
            
            # Use end('down') to find last row
            last_row = check_cell.end('down').row
            
            # Handle case where end('down') goes to very last row (means no data)
            # Check if the last_row actually has data
            if last_row > 1048576 or last_row < start_row:
                # No data found, just clear the start row
                last_row = start_row

            # If there's data, clear it
            if last_row >= start_row:
                range_address = f"{start_cell}:{end_column}{last_row}"
                print(f"   Clearing range {range_address} in sheet '{sheet_name}'...")
                sheet.range(range_address).clear_contents()
                print(f"    Range cleared successfully ({last_row - start_row + 1} rows)")
            else:
                print(f"   No data to clear in sheet '{sheet_name}'")
                
        except Exception as e:
            print(f"   Warning: {e}")
            # Fallback: try to clear used range
            try:
                used_range = sheet.used_range
                last_row = used_range.last_cell.row
                
                if last_row >= start_row:
                    range_address = f"{start_cell}:{end_column}{last_row}"
                    print(f"   Using fallback method, clearing {range_address}...")
                    sheet.range(range_address).clear_contents()
                    print(f"    Range cleared successfully")
                else:
                    print(f"   No data to clear")
            except Exception as e2:
                print(f"    Could not clear range: {e2}")


    def read_csv_data(self, csv_path: str, num_columns: int = 27) -> Tuple[pd.DataFrame, int, int]:
        """
        Read data from CSV file

        Args:
            csv_path: Path to the CSV file
            num_columns: Number of columns to read (default: 27 for A-AA)

        Returns:
            Tuple of (DataFrame, num_rows, num_cols)
        """
        print(f"   Reading CSV file: {csv_path}")
        df = pd.read_csv(csv_path, header=0)

        # Get specified number of columns
        if len(df.columns) >= num_columns:
            df = df.iloc[:, 0:num_columns]
        else:
            print(f"   Note: CSV has only {len(df.columns)} columns (expected {num_columns})")

        num_rows = len(df)
        num_cols = len(df.columns)
        print(f"    Successfully read {num_rows} rows and {num_cols} columns")

        return df, num_rows, num_cols

    def read_excel_data(self, excel_path: str, num_columns: int = 5, skip_rows: int = 0, sheet_name: str = None, engine='openpyxl') -> Tuple[pd.DataFrame, int, int]:
        """
        Read data from Excel file

        Args:
            excel_path: Path to the Excel file
            num_columns: Number of columns to read (default: 5 for A-E)

        Returns:
            Tuple of (DataFrame, num_rows, num_cols)
        """
        print(f"   Reading Excel file: {excel_path}")
        df = pd.read_excel(excel_path, header=0, skiprows=skip_rows, sheet_name=sheet_name, engine=engine)

        # Get specified number of columns
        if len(df.columns) >= num_columns:
            df = df.iloc[:, 0:num_columns]
        else:
            print(f"   Note: Excel has only {len(df.columns)} columns (expected {num_columns})")

        num_rows = len(df)
        num_cols = len(df.columns)
        print(f"    Successfully read {num_rows} rows and {num_cols} columns")

        return df, num_rows, num_cols

    def write_data_to_range(self, sheet_name: str, start_cell: str, data: list) -> None:
        """
        Write data to a worksheet starting at specified cell

        Args:
            sheet_name: Name of the worksheet
            start_cell: Starting cell (e.g., 'A4', 'B4')
            data: List of lists containing data to write
        """
        if not data:
            print(f"   No data to write to sheet '{sheet_name}'")
            return

        sheet = self.workbook.sheets[sheet_name]
        num_rows = len(data)
        num_cols = len(data[0]) if data else 0

        print(f"   Writing {num_rows} rows x {num_cols} columns to sheet '{sheet_name}' starting at {start_cell}...")

        # Write data to sheet
        sheet.range(start_cell).value = data

        print(f"    Data written successfully")

    def fill_column_with_value(self, sheet_name: str, column: str, start_row: int,
                                end_row: int, value: str) -> None:
        """
        Fill a column with a specific value

        Args:
            sheet_name: Name of the worksheet
            column: Column letter (e.g., 'A')
            start_row: Starting row number
            end_row: Ending row number
            value: Value to fill
        """
        sheet = self.workbook.sheets[sheet_name]
        range_address = f"{column}{start_row}:{column}{end_row}"

        print(f"   Filling {range_address} with value: {value}")
        sheet.range(range_address).value = value
        print(f"    Column filled successfully")

    def copy_formulas_to_range(self, sheet_name: str, source_range: str,
                                target_start_row: int, target_end_row: int) -> None:
        """
        Copy formulas from the last row with data in the column range to target rows

        Args:
            sheet_name: Name of the worksheet
            source_range: Column range to search (e.g., 'F2:O2' or 'F:O')
            target_start_row: Starting row for pasting formulas
            target_end_row: Ending row for pasting formulas
        """
        sheet = self.workbook.sheets[sheet_name]

        # Parse source range to get column range
        parts = source_range.split(':')
        if len(parts) != 2:
            raise ValueError(f"Invalid range format: {source_range}. Expected format like 'F2:O2'")

        start_cell = parts[0]
        end_cell = parts[1]

        # Extract column letters from cells
        start_col = ''.join([c for c in start_cell if c.isalpha()])
        end_col = ''.join([c for c in end_cell if c.isalpha()])

        # Find the last row with data in the start column of the range
        last_row_with_data = sheet.range(f'{start_col}5').end('down').row

        # Build the actual source range from the last row with data
        actual_source_range = f"{start_col}{last_row_with_data}:{end_col}{last_row_with_data}"

        print(f"   Found last row with data at row {last_row_with_data}")
        print(f"   Copying formulas from {actual_source_range} to rows {target_start_row}:{target_end_row}")

        # Copy the source range from the last row with data
        source_range_obj = sheet.range(actual_source_range)
        target_start_row = max(target_start_row, last_row_with_data + 1)  # Ensure we start below the source row
        # Paste to each target row (starting from the row after last data row)
        if target_start_row >= target_end_row:
            print("   No rows to copy formulas to (target start row is after target end row)")
        else:
            for row in range(target_start_row, target_end_row + 1):
                target_range = f"{start_col}{row}:{end_col}{row}"
                source_range_obj.copy(sheet.range(target_range))

            print(f"    Formulas copied successfully to {target_end_row - target_start_row + 1} rows")

    def save_as(self, output_path: str) -> None:
        """
        Save workbook as a new file

        Args:
            output_path: Path for the output file
        """
        print(f"   Saving workbook as: {output_path}")

        # Get full path
        full_path = os.path.abspath(output_path)

        # Save as new file
        self.workbook.save(full_path)
        print(f"    Workbook saved successfully")

    def copy_formula_and_paste_values(self, sheet_name: str, formula_cell: str, 
                                   target_start_cell: str, find_last_row_column: str = 'A') -> None:
        """
        Copy formula from a single cell and paste to a range, then convert to values
        Uses Excel's native PasteSpecial for better performance
        
        Args:
            sheet_name: Name of the worksheet
            formula_cell: Source cell with formula (e.g., 'A1')
            target_start_cell: Starting cell for pasting (e.g., 'A4')
            find_last_row_column: Column to check for last row with data (default: 'A')
        """
        sheet = self.workbook.sheets[sheet_name]
        
        # Find the last row with data in the specified column
        last_row = sheet.range(f'{find_last_row_column}4').end('down').row
        
        # Extract column letter from target_start_cell
        target_col = ''.join([c for c in target_start_cell if c.isalpha()])
        target_row = int(''.join([c for c in target_start_cell if c.isdigit()]))
        
        # Build target range
        target_range = f"{target_col}{target_row}:{target_col}{last_row}"
        
        print(f"   Copying formula from {formula_cell} to {target_range} in sheet '{sheet_name}'...")
        
        # Copy formula to range
        formula_source = sheet.range(formula_cell)
        target_range_obj = sheet.range(target_range)
        formula_source.copy(target_range_obj)
        
        print(f"   Converting formulas to values in {target_range}...")
        
        # Convert to values using Excel's native method
        # Copy the range
        target_range_obj.api.Copy()
        
        # Paste as values using Excel constant (xlPasteValues = -4163)
        target_range_obj.api.PasteSpecial(Paste=-4163)
        
        # Clear clipboard
        self.app.api.CutCopyMode = False
        
        print(f"    Formula copied and converted to values successfully")

    def refresh_pivot_table(self, sheet_name: str, pivot_table_name: str) -> None:
        """
        Refresh a specific pivot table in a worksheet

        Args:
            sheet_name: Name of the worksheet containing the pivot table
            pivot_table_name: Name of the pivot table to refresh
        """
        try:
            print(f"   Refreshing pivot table '{pivot_table_name}' in sheet '{sheet_name}'...")

            sheet = self.workbook.sheets[sheet_name]

            # Access the pivot table through Excel API
            pivot_table = sheet.api.PivotTables(pivot_table_name)

            # Refresh the pivot table
            pivot_table.RefreshTable()

            print(f"    Pivot table '{pivot_table_name}' refreshed successfully")

        except Exception as e:
            print(f"    Error refreshing pivot table: {e}")
            raise

    def update_pivot_source_and_refresh(self,
                                       sheet_name: str,
                                       pivot_table_name: str,
                                       data_sheet_name: str,
                                       start_cell: str = 'A1',
                                       end_column: str = None,
                                       include_headers: bool = True) -> None:
        """
        Dynamically detect data range, update pivot table source, and refresh it

        Args:
            sheet_name: Name of the worksheet containing the pivot table
            pivot_table_name: Name of the pivot table to update
            data_sheet_name: Name of the worksheet containing the source data
            start_cell: Starting cell of the data range (default: 'A1')
            end_column: Last column letter (e.g., 'Z'). If None, will auto-detect
            include_headers: Whether to include header row in the range (default: True)
        """
        try:
            print(f"   Updating data source for pivot table '{pivot_table_name}'...")

            # Get the data sheet
            data_sheet = self.workbook.sheets[data_sheet_name]

            # Extract start column and row from start_cell
            start_column = ''.join(filter(str.isalpha, start_cell))
            start_row = int(''.join(filter(str.isdigit, start_cell)))

            # Find the last row with data
            # Start checking from one row below the start_cell
            check_row = start_row + 1 if include_headers else start_row
            last_row = data_sheet.range(f'{start_column}{check_row}').end('down').row

            # Handle edge case where no data exists
            if last_row > 1048576 or last_row < start_row:
                last_row = start_row
                print(f"   Warning: No data found, using only header row")

            # If end_column not specified, find it dynamically
            if end_column is None:
                # Find last column with data in the header row
                end_column_obj = data_sheet.range(f'{start_column}{start_row}').end('right')
                end_column = end_column_obj.get_address(False, False).split(':')[0]
                end_column = ''.join(filter(str.isalpha, end_column))

            # Build the dynamic range address
            range_address = f"{start_cell}:{end_column}{last_row}"
            full_range_address = f"{data_sheet_name}!{range_address}"

            print(f"   Detected data range: {full_range_address}")
            print(f"   Range contains {last_row - start_row + 1} rows (including headers)")

            # Get the pivot table
            sheet = self.workbook.sheets[sheet_name]
            pivot_table = sheet.api.PivotTables(pivot_table_name)

            # Update the source data range
            # Build the full workbook reference for the range
            workbook_name = self.workbook.name
            full_source = f"'{workbook_name}'!{full_range_address}"

            print(f"   Updating pivot table source to: {full_source}")
            pivot_table.ChangePivotCache(
                self.workbook.api.PivotCaches().Create(
                    SourceType=1,  # xlDatabase
                    SourceData=full_source
                )
            )

            # Refresh the pivot table
            pivot_table.RefreshTable()

            print(f"    Pivot table '{pivot_table_name}' source updated and refreshed successfully")

        except Exception as e:
            print(f"    Error updating pivot table source: {e}")
            raise

    def read_sheet_range_to_dataframe(self, sheet_name: str, range_address: str = None) -> pd.DataFrame:
        """
        Read data from a specific range in a sheet to a pandas DataFrame
        If range_address is None, reads the entire used range

        Args:
            sheet_name: Name of the worksheet to read from
            range_address: Range to read (e.g., 'A1:Z100'). If None, reads used range

        Returns:
            pd.DataFrame: DataFrame containing the data from the range
        """
        print(f"   Reading data from sheet '{sheet_name}'...")

        sheet = self.workbook.sheets[sheet_name]

        if range_address:
            # Read specific range
            data = sheet.range(range_address).value
            print(f"    Read range {range_address}")
        else:
            # Read entire used range
            used_range = sheet.used_range
            data = used_range.value
            range_address = used_range.address
            print(f"    Read used range {range_address}")

        # Convert to DataFrame
        if data:
            # If data is a list of lists, convert to DataFrame
            if isinstance(data, list) and len(data) > 0:
                # Use first row as headers
                df = pd.DataFrame(data[1:], columns=data[0])
            else:
                # Single cell value
                df = pd.DataFrame([data])

            print(f"    Successfully read {len(df)} rows and {len(df.columns)} columns")
            return df
        else:
            print(f"    No data found in range")
            return pd.DataFrame()

    def write_dataframe_to_sheet(self, sheet_name: str, start_cell: str, df: pd.DataFrame,
                                  include_headers: bool = True, clear_existing: bool = False) -> None:
        """
        Write a pandas DataFrame to a specific location in a sheet

        Args:
            sheet_name: Name of the worksheet to write to
            start_cell: Starting cell for writing (e.g., 'A1')
            df: DataFrame to write
            include_headers: Whether to include column headers (default: True)
            clear_existing: Whether to clear existing data in the target range (default: False)
        """
        print(f"   Writing DataFrame to sheet '{sheet_name}' starting at {start_cell}...")

        sheet = self.workbook.sheets[sheet_name]

        # Prepare data
        if include_headers:
            # Include headers
            data = [df.columns.tolist()] + df.values.tolist()
        else:
            # Data only
            data = df.values.tolist()

        # Calculate target range if we need to clear
        if clear_existing and data:
            num_rows = len(data)
            num_cols = len(data[0]) if data else 0

            # Calculate end cell by offsetting from start cell
            end_cell = sheet.range(start_cell).offset(num_rows - 1, num_cols - 1)
            end_cell_address = end_cell.get_address(False, False)

            # Clear the range
            clear_range = f"{start_cell}:{end_cell_address}"
            print(f"   Clearing existing data in range {clear_range}...")
            sheet.range(clear_range).clear_contents()

        # Write data
        if data:
            sheet.range(start_cell).value = data
            print(f"    Successfully wrote {len(df)} rows and {len(df.columns)} columns")
        else:
            print(f"    No data to write")

    def delete_rows_after_last_data(self, sheet_name: str, check_column: str = 'A',
                                     start_row: int = 2, max_delete_rows: int = 1000000) -> None:
        """
        Delete all empty rows after the last row with data in a sheet
        This helps clean up sheets that have extra empty rows extending the used range

        Args:
            sheet_name: Name of the worksheet to clean up
            check_column: Column to check for last data (default: 'A')
            start_row: Row to start checking from (default: 2, to preserve headers)
            max_delete_rows: Maximum number of rows to delete in one go (default: 1000000)
        """
        print(f"   Deleting extra rows after last data in sheet '{sheet_name}'...")

        sheet = self.workbook.sheets[sheet_name]

        try:
            # Find the last row with data in the check column
            last_row_with_data = sheet.range(f'{check_column}{start_row}').end('down').row

            # Handle case where end('down') goes to the very last row
            if last_row_with_data > 1048576 or last_row_with_data < start_row:
                last_row_with_data = start_row
                print(f"   No data found below row {start_row}")
                return

            print(f"   Last row with data: {last_row_with_data}")

            # Find the actual used range last row
            used_range_last_row = sheet.used_range.last_cell.row

            if used_range_last_row > last_row_with_data:
                # Calculate how many rows to delete
                rows_to_delete = min(used_range_last_row - last_row_with_data, max_delete_rows)
                start_delete_row = last_row_with_data + 1

                print(f"   Deleting {rows_to_delete} empty rows starting from row {start_delete_row}...")

                # Delete the entire rows
                end_delete_row = start_delete_row + rows_to_delete - 1
                sheet.range(f'{start_delete_row}:{end_delete_row}').api.EntireRow.Delete()

                print(f"    Successfully deleted {rows_to_delete} empty rows")
            else:
                print(f"   No extra rows to delete (last data row matches used range)")

        except Exception as e:
            print(f"    Error deleting extra rows: {e}")
            raise

    def find_all_pivot_tables(self) -> dict:
        """
        Find all pivot tables in the entire workbook

        Returns:
            Dictionary with sheet names as keys and list of pivot table names as values
        """
        print(f"   Searching for all pivot tables in workbook...")

        pivot_tables_dict = {}

        for sheet in self.workbook.sheets:
            sheet_name = sheet.name
            try:
                pivot_tables = sheet.api.PivotTables()
                if pivot_tables.Count > 0:
                    pivot_list = []
                    for i in range(1, pivot_tables.Count + 1):
                        pivot_list.append(pivot_tables(i).Name)
                    pivot_tables_dict[sheet_name] = pivot_list
                    print(f"    Found {pivot_tables.Count} pivot table(s) in sheet '{sheet_name}': {pivot_list}")
            except Exception as e:
                # Sheet might not support pivot tables
                pass

        if not pivot_tables_dict:
            print(f"    No pivot tables found in workbook")

        return pivot_tables_dict

    # =============================================================================
    # NEW METHODS - Moved from Main.py and made dynamic
    # =============================================================================

    def read_portfolio_data(self, sheet_name: str) -> pd.DataFrame:
        """
        Read portfolio data from a sheet (reads all columns as-is).
        Instance method that uses the current workbook.

        Args:
            sheet_name: Name of the portfolio sheet to read

        Returns:
            pd.DataFrame: Portfolio data with all columns
        """
        sheet = self.workbook.sheets[sheet_name]
        last_row = sheet.range('A2').end('down').row

        # Handle empty sheet
        if last_row > 1000000:
            print(f"   {sheet_name}: Empty sheet detected")
            return pd.DataFrame()

        # Read entire used range to preserve all columns
        df = self.read_sheet_range_to_dataframe(sheet_name, None)
        return df

    def write_portfolio_data(self, sheet_name: str, df: pd.DataFrame, 
                            columns_to_write: List[str],
                            column_positions: dict = None) -> None:
        """
        Write specific columns to Portfolio sheet at specified positions.
        Leaves formula columns untouched.

        Args:
            sheet_name: Name of the portfolio sheet
            df: DataFrame containing data to write
            columns_to_write: List of column names to write (e.g., ['MONTH', 'CONTRACT_NO', ...])
            column_positions: Dict mapping column names to Excel columns (e.g., {'MONTH': 'A', 'CONTRACT_NO': 'B'})
                            If None, uses default mapping: A=MONTH, B=CONTRACT_NO, D=EQT_DESC, E=PD_CATEGORY, F=DPD
        """
        sheet = self.workbook.sheets[sheet_name]

        # Default column positions if not provided
        if column_positions is None:
            column_positions = {
                'MONTH': 'A',
                'CONTRACT_NO': 'B',
                'EQT_DESC': 'D',
                'PD_CATEGORY': 'E',
                'DPD': 'F'
            }

        # Clear only the specified data columns (skip formula columns)
        last_row = sheet.range('A2').end('down').row
        if last_row < 1000000:  # Has data
            try:
                # Clear columns A and B
                sheet.range(f'A2:B{last_row}').clear_contents()
                # Clear columns D, E, F (skip C which is formula)
                sheet.range(f'D2:F{last_row}').clear_contents()
            except:
                pass

        if df.empty:
            print(f"  {sheet_name}: No data to write")
            return

        # Write each column to its specified position
        num_rows = len(df)
        for col_name in columns_to_write:
            if col_name in df.columns and col_name in column_positions:
                excel_col = column_positions[col_name]
                sheet.range(f'{excel_col}2').value = df[col_name].values.reshape(-1, 1).tolist()

        print(f"  {sheet_name}: {num_rows} rows written ({len(columns_to_write)} columns: {', '.join(columns_to_write)})")

    # =============================================================================
    # STATIC UTILITY METHODS - Date and File Operations
    # =============================================================================

    @staticmethod
    def months_between(start_date: datetime, end_date: datetime) -> int:
        """
        Calculate number of months between two dates.

        Args:
            start_date: Starting date
            end_date: Ending date

        Returns:
            int: Number of months between dates
        """
        return (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month)

    @staticmethod
    def parse_month_string(month_str: str, date_format: str = '%m/%d/%Y') -> datetime:
        """
        Parse month string to datetime object.

        Args:
            month_str: Date string to parse
            date_format: Format of the date string (default: '%m/%d/%Y')

        Returns:
            datetime: Parsed datetime object
        """
        return datetime.strptime(month_str.strip(), date_format)

    @staticmethod
    def format_month_string(date: datetime, date_format: str = '%m/%d/%Y') -> str:
        """
        Format datetime to string.

        Args:
            date: Datetime object to format
            date_format: Desired output format (default: '%m/%d/%Y')

        Returns:
            str: Formatted date string
        """
        return date.strftime(date_format)

    @staticmethod
    def read_config_file(config_file_path: str, date_format: str = '%m/%d/%Y') -> datetime:
        """
        Read the latest recorded month from a config file.

        Args:
            config_file_path: Path to the config file
            date_format: Expected date format in config file (default: '%m/%d/%Y')

        Returns:
            datetime: Latest month from config file

        Raises:
            FileNotFoundError: If config file doesn't exist
            ValueError: If date format is invalid
        """
        if not os.path.exists(config_file_path):
            raise FileNotFoundError(f"Config file not found: {config_file_path}")

        with open(config_file_path, 'r') as f:
            date_str = f.read().strip()

        try:
            return ExcelPortfolioAutomation.parse_month_string(date_str, date_format)
        except ValueError:
            raise ValueError(f"Invalid date format in config file: {date_str}. Expected format: {date_format}")

    @staticmethod
    def save_config_file(config_file_path: str, latest_month: datetime, date_format: str = '%m/%d/%Y') -> None:
        """
        Save the latest month to a config file.

        Args:
            config_file_path: Path to the config file
            latest_month: Date to save
            date_format: Date format to use (default: '%m/%d/%Y')
        """
        with open(config_file_path, 'w') as f:
            f.write(ExcelPortfolioAutomation.format_month_string(latest_month, date_format))
        print(f"Config updated: {ExcelPortfolioAutomation.format_month_string(latest_month, date_format)}")

    @staticmethod
    def get_latest_file_in_folder(folder_path: str, file_pattern: str, 
                                  fallback_file: str = None) -> str:
        """
        Get the latest file matching a pattern in a folder.
        If no files found, returns fallback file.

        Args:
            folder_path: Path to search for files
            file_pattern: Glob pattern to match files (e.g., "*.xlsb", "PD_data_*_Updated_*.xlsb")
            fallback_file: File to return if no matches found (default: None)

        Returns:
            str: Path to the latest file or fallback file
        """
        # Search for files matching pattern
        search_pattern = os.path.join(folder_path, file_pattern)
        matching_files = glob.glob(search_pattern)

        if matching_files:
            # Sort by modification time (most recent first)
            matching_files.sort(key=os.path.getmtime, reverse=True)
            latest_file = matching_files[0]
            print(f"Using latest file: {os.path.basename(latest_file)}")
            return latest_file
        else:
            if fallback_file:
                print(f"Using fallback file: {os.path.basename(fallback_file)}")
                return fallback_file
            else:
                raise FileNotFoundError(f"No files found matching pattern: {file_pattern}")

    @staticmethod
    def find_summary_files_by_date_range(input_folder: str, start_month: datetime, 
                                        num_months: int, file_prefix: str = "3. Summary_",
                                        file_extension: str = ".xlsb",
                                        date_format_in_filename: str = '%Y-%m-%d') -> List[Tuple[str, datetime]]:
        """
        Find summary files for a specified date range.

        Args:
            input_folder: Folder containing summary files
            start_month: Starting month
            num_months: Number of months to find
            file_prefix: Prefix of summary files (default: "3. Summary_")
            file_extension: File extension (default: ".xlsb")
            date_format_in_filename: Date format used in filenames (default: '%Y-%m-%d')

        Returns:
            List of (file_path, month_date) tuples in chronological order
        """
        files = []

        for i in range(1, num_months + 1):
            target_date = ExcelPortfolioAutomation.add_months(start_month, i)
            # Summary files use specified date format in filename
            date_pattern = target_date.strftime(date_format_in_filename)

            pattern = os.path.join(input_folder, f"{file_prefix}{date_pattern}*{file_extension}")
            matches = glob.glob(pattern)

            if matches:
                files.append((matches[0], target_date))
                print(f"  Found: {os.path.basename(matches[0])}")
            else:
                print(f"  WARNING: No file for {date_pattern}")

        return files

    @staticmethod
    def extract_data_from_summary_files(file_paths: List[Tuple[str, datetime]], 
                                       column_mapping: dict,
                                       output_columns: List[str],
                                       sheet_name: str = 'SUMMARY',
                                       date_format: str = '%m/%d/%Y',
                                       max_header_search_rows: int = 10) -> pd.DataFrame:
        """
        Extract data from summary files into a DataFrame.

        Args:
            file_paths: List of (file_path, month_date) tuples
            column_mapping: Dict mapping source column names to target column names
                          e.g., {'CONTRACT NO': 'CONTRACT_NO', 'EQUIPMENT DESCRIPTION': 'EQT_DESC'}
            output_columns: List of columns to include in output (in order)
            sheet_name: Name of sheet to read (default: 'SUMMARY')
            date_format: Format for MONTH column (default: '%m/%d/%Y')
            max_header_search_rows: Maximum rows to search for headers (default: 10)

        Returns:
            pd.DataFrame: Consolidated data from all summary files
        """
        all_data = []

        for file_path, month_date in file_paths:
            try:
                print(f"  Extracting: {os.path.basename(file_path)}")

                # Find correct header row
                df = None
                for header_row in range(max_header_search_rows):
                    try:
                        temp_df = pd.read_excel(file_path, sheet_name=sheet_name,
                                               header=header_row, engine='pyxlsb')
                        # Check if any of the source columns exist
                        if any(col in temp_df.columns for col in column_mapping.keys()):
                            df = temp_df
                            break
                    except:
                        continue

                if df is None:
                    df = pd.read_excel(file_path, sheet_name=sheet_name,
                                      header=0, engine='pyxlsb')

                # Map columns
                extracted = pd.DataFrame()
                for src_col, tgt_col in column_mapping.items():
                    found = None
                    for col in df.columns:
                        if str(col).strip().upper() == src_col.upper():
                            found = col
                            break
                    extracted[tgt_col] = df[found] if found else None

                # Add MONTH column
                extracted['MONTH'] = ExcelPortfolioAutomation.format_month_string(month_date, date_format)

                # Filter empty rows
                first_col = list(column_mapping.values())[0]  # Use first target column for filtering
                extracted = extracted[extracted[first_col].notna()]
                extracted = extracted[~extracted[first_col].astype(str).isin(['', '-', 'nan', 'None'])]

                # Reorder columns
                extracted = extracted[output_columns]

                all_data.append(extracted)
                print(f"    Rows: {len(extracted)}")

            except Exception as e:
                print(f"    ERROR: {e}")

        if all_data:
            return pd.concat(all_data, ignore_index=True)
        return pd.DataFrame(columns=output_columns)

    @staticmethod
    def get_unique_months_from_dataframe(df: pd.DataFrame, month_column: str = 'MONTH',
                                        date_format: str = '%m/%d/%Y') -> List[datetime]:
        """
        Get sorted list of unique months from a DataFrame.

        Args:
            df: DataFrame containing month data
            month_column: Name of the column containing month data (default: 'MONTH')
            date_format: Format of date strings in the column (default: '%m/%d/%Y')

        Returns:
            List of unique month datetime objects in ascending order
        """
        if df.empty or month_column not in df.columns:
            return []

        months = pd.to_datetime(df[month_column], format=date_format, errors='coerce')
        unique_months = sorted(months.dropna().unique())
        return [pd.Timestamp(m).to_pydatetime() for m in unique_months]

    @staticmethod
    def filter_dataframe_by_months(df: pd.DataFrame, months_to_keep: List[datetime],
                                   month_column: str = 'MONTH',
                                   date_format: str = '%m/%d/%Y') -> pd.DataFrame:
        """
        Filter DataFrame to keep only rows matching specified months.

        Args:
            df: DataFrame to filter
            months_to_keep: List of datetime objects representing months to keep
            month_column: Name of the month column (default: 'MONTH')
            date_format: Format of date strings in the column (default: '%m/%d/%Y')

        Returns:
            pd.DataFrame: Filtered dataframe
        """
        if df.empty:
            return df

        df = df.copy()
        df['_MONTH_DT'] = pd.to_datetime(df[month_column], format=date_format, errors='coerce')

        # Convert months_to_keep to timestamps for comparison
        keep_timestamps = [pd.Timestamp(m) for m in months_to_keep]
        filtered = df[df['_MONTH_DT'].isin(keep_timestamps)]

        return filtered.drop(columns=['_MONTH_DT'])

    @staticmethod
    def consolidate_summary_files(input_folder: str, file_pattern: str = "3. Summary_*.xlsb",
                                   sheet_name: str = "SUMMARY", header_row: int = 0) -> pd.DataFrame:
        """
        Consolidate data from multiple summary Excel files into a single DataFrame

        Args:
            input_folder: Path to folder containing summary files
            file_pattern: Pattern to match summary files (default: "3. Summary_*.xlsb")
            sheet_name: Name of the sheet to read from (default: "SUMMARY")
            header_row: Row number containing column headers (default: 0)

        Returns:
            pd.DataFrame: Consolidated dataframe with columns CONTRACT_NO, EQT_DESC, PD_CATEGORY, DPD, MONTH
        """
        print(f"   Consolidating summary files from: {input_folder}")
        print(f"   File pattern: {file_pattern}")
        print()

        # Find all matching files
        search_pattern = os.path.join(input_folder, file_pattern)
        all_summary_files = sorted(glob.glob(search_pattern))

        # Take only the latest 6 files
        summary_files = all_summary_files[-6:] if len(all_summary_files) >= 6 else all_summary_files

        print(f"   Found {len(all_summary_files)} summary files, using latest {len(summary_files)}:")
        for file in summary_files:
            print(f"     - {os.path.basename(file)}")
        print()

        # Column mapping: source column name -> target column name
        column_mapping = {
            'CONTRACT NO': 'CONTRACT_NO',
            'EQUIPMENT DESCRIPTION': 'EQT_DESC',
            'PD/LGD CATEGORY': 'PD_CATEGORY',
            'CLIENT DPD': 'DPD'
        }

        consolidated_data = []

        for file_path in summary_files:
            try:
                print(f"   Processing: {os.path.basename(file_path)}")

                # Read the SUMMARY sheet using pandas (works with closed files)
                # Try to find the correct header row
                df = None
                for try_header in range(header_row, min(header_row + 10, 20)):
                    try:
                        temp_df = pd.read_excel(file_path, sheet_name=sheet_name, header=try_header, engine='pyxlsb')
                        # Check if this row contains the columns we're looking for
                        if 'CONTRACT NO' in temp_df.columns or 'CONTRACT_NO' in temp_df.columns:
                            df = temp_df
                            if try_header != header_row:
                                print(f"     Found headers at row {try_header} (tried starting from row {header_row})")
                            break
                    except:
                        continue

                if df is None:
                    # Fall back to specified header row
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, engine='pyxlsb')

                print(f"     Read {len(df)} rows")
                print(f"     Available columns: {list(df.columns)[:10]}...")  # Show first 10 columns only

                # Select and rename the required columns
                selected_data = pd.DataFrame()

                for source_col, target_col in column_mapping.items():
                    # Try to find the column (case-insensitive, with or without spaces)
                    found_col = None
                    for col in df.columns:
                        if str(col).strip().upper() == source_col.upper():
                            found_col = col
                            break

                    if found_col:
                        selected_data[target_col] = df[found_col]
                    else:
                        print(f"     Warning: Column '{source_col}' not found in {os.path.basename(file_path)}")
                        selected_data[target_col] = None

                # Extract date from filename (e.g., "3. Summary_2025-04-30_Final_V2.xlsb" -> "04/30/2025")
                filename = os.path.basename(file_path)
                date_match = re.search(r'(\d{4})-(\d{2})-(\d{2})', filename)
                if date_match:
                    year, month, day = date_match.groups()
                    month_date = f"{month}/{day}/{year}"
                else:
                    month_date = "Unknown"

                selected_data['MONTH'] = month_date
                print(f"     Extracted month: {month_date}")

                consolidated_data.append(selected_data)
                print(f"     Successfully extracted {len(selected_data)} rows with {len(selected_data.columns)} columns")

            except Exception as e:
                print(f"     Error processing {os.path.basename(file_path)}: {e}")
                import traceback
                traceback.print_exc()

        # Combine all dataframes
        if consolidated_data:
            final_df = pd.concat(consolidated_data, ignore_index=True)
            print()
            print(f"   Consolidation complete!")
            print(f"     Total rows: {len(final_df)}")
            print(f"     Total columns: {len(final_df.columns)}")
            print(f"     Columns: {list(final_df.columns)}")
            return final_df
        else:
            print()
            print(f"   No data consolidated")
            return pd.DataFrame()

    @staticmethod
    def add_months(date: datetime, months: int) -> datetime:
        """Add months to a date, handling month-end correctly.
        If original date is month-end, result will also be month-end."""
        # Check if original date is the last day of its month
        original_month_last_day = calendar.monthrange(date.year, date.month)[1]
        is_month_end = (date.day == original_month_last_day)

        # Calculate target month/year
        month = date.month - 1 + months
        year = date.year + month // 12
        month = month % 12 + 1

        # Get last day of target month
        target_month_last_day = calendar.monthrange(year, month)[1]

        # If original was month-end, use target month-end; otherwise use original day (capped)
        if is_month_end:
            day = target_month_last_day
        else:
            day = min(date.day, target_month_last_day)

        return datetime(year, month, day)

    # =============================================================================
    # PIVOT TABLE & HISTORIC PD METHODS
    # =============================================================================

    def extract_pivot_table_to_dataframe(self, sheet_name: str, 
                                        start_cell: str = 'A4',
                                        contract_col: str = 'A',
                                        pd_category_col: str = 'B',
                                        first_data_col: str = 'C') -> pd.DataFrame:
        """
        Extract pivot table data from a sheet to DataFrame.
        Assumes pivot structure:
        - Column A: CONTRACT_NO_NOLASTDIG
        - Column B: PD_CATEGORY  
        - Column C onwards: Date columns (e.g., 2024-09, 2024-10, etc.)

        Args:
            sheet_name: Name of the sheet containing pivot table
            start_cell: Starting cell of pivot data (default: 'A4')
            contract_col: Column letter for contract numbers (default: 'A')
            pd_category_col: Column letter for PD category (default: 'B')
            first_data_col: First data column letter (default: 'C')

        Returns:
            pd.DataFrame: Pivot data with multi-index (CONTRACT_NO_NOLASTDIG, PD_CATEGORY) 
                         and date columns
        """
        print(f"\n   Extracting pivot table from '{sheet_name}'...")
        
        sheet = self.workbook.sheets[sheet_name]
        
        # Read the entire used range
        used_range = sheet.used_range
        data = used_range.value
        
        if not data or len(data) < 4:
            print(f"   ERROR: No data found in sheet '{sheet_name}'")
            return pd.DataFrame()
        
        # Find header row (row with date columns like "2024-09")
        header_row_idx = None
        for i, row in enumerate(data):
            if row and len(row) > 2:
                # Check if row has date-like values (YYYY-MM format)
                for cell in row[2:]:  # Skip first 2 columns
                    if cell and isinstance(cell, str) and re.match(r'\d{4}-\d{2}', str(cell)):
                        header_row_idx = i
                        break
                if header_row_idx is not None:
                    break
        
        if header_row_idx is None:
            print(f"   ERROR: Could not find header row with date columns")
            return pd.DataFrame()
        
        print(f"   Found header row at index: {header_row_idx}")
        
        # Extract headers (date columns)
        headers = ['CONTRACT_NO_NOLASTDIG', 'PD_CATEGORY']
        date_headers = []
        for cell in data[header_row_idx][2:]:  # Skip first 2 columns
            if cell:
                headers.append(str(cell))
                date_headers.append(str(cell))
            else:
                break  # Stop at first empty column
        
        print(f"   Found {len(date_headers)} date columns: {date_headers[:5]}...")
        
        # Extract data rows (start from row after header)
        data_rows = []
        for row in data[header_row_idx + 1:]:
            if row and row[0]:  # If first column has value (contract number)
                # Take only the columns we need (up to length of headers)
                row_data = row[:len(headers)]
                data_rows.append(row_data)
        
        print(f"   Extracted {len(data_rows)} data rows")
        
        # Create DataFrame
        df = pd.DataFrame(data_rows, columns=headers)
        
        # Filter out empty rows and pivot artifacts like (blank) and Grand Total
        df = df[df['CONTRACT_NO_NOLASTDIG'].notna()]
        df = df[~df['CONTRACT_NO_NOLASTDIG'].astype(str).str.strip().isin(['(blank)', 'blank', '', 'Grand Total'])]
        df = df[df['PD_CATEGORY'].notna()]
        df = df[~df['PD_CATEGORY'].astype(str).str.strip().isin(['(blank)', 'blank', '', 'Grand Total'])]
        
        print(f"   Final DataFrame: {len(df)} rows x {len(df.columns)} columns")
        print(f"{df.head}")
        
        return df

    @staticmethod
    def convert_pivot_date_to_year_month(date_str: str) -> tuple:
        """
        Convert pivot date format to year and month name.
        
        Args:
            date_str: Date string in format "YYYY-MM" (e.g., "2024-09")
        
        Returns:
            tuple: (year, month_name) e.g., (2024, "Sep")
        """
        try:
            # Parse the date string
            parts = date_str.split('-')
            year = int(parts[0])
            month_num = int(parts[1])
            
            # Convert month number to name
            month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                          'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            month_name = month_names[month_num - 1]
            
            return (year, month_name)
        except:
            return (None, None)

    def write_historic_pd_format(self, sheet_name: str, pivot_df: pd.DataFrame,
                                year_row: int = 1, month_row: int = 2, data_start_row: int = 3,
                                contract_col: str = 'A', pd_category_col: str = 'B') -> None:
        """
        Write pivot data to Historic PD sheet format using direct positional mapping.

        Dynamically writes:
        - Row 1: Year headers (e.g., 2024 for C1:G1, 2025 for H1:O1)
        - Row 2: Month abbreviation headers (e.g., Aug, Sep, ..., Jul, Aug)
        - Row 3+: Data (CONTRACT_NO_NOLASTDIG in col A, PD_CATEGORY in col B, values in C:O)

        Mapping is positional: first pivot date column -> col C, second -> col D, etc.
        Total months is always 13.

        Args:
            sheet_name: Name of the sheet to update
            pivot_df: DataFrame with pivot data (from extract_pivot_table_to_dataframe)
            year_row: Row number for years (default: 1)
            month_row: Row number for month names (default: 2)
            data_start_row: First data row (default: 3)
            contract_col: Column letter for contract numbers (default: 'A')
            pd_category_col: Column letter for PD category (default: 'B')
        """
        print(f"\n   Writing data to '{sheet_name}'...")

        sheet = self.workbook.sheets[sheet_name]

        # Get date columns from pivot DataFrame (only YYYY-MM format columns)
        date_columns = [col for col in pivot_df.columns
                       if col not in ['CONTRACT_NO_NOLASTDIG', 'PD_CATEGORY']
                       and re.match(r'\d{4}-\d{2}', str(col))]

        if not date_columns:
            print(f"   ERROR: No date columns found in pivot data")
            return

        print(f"   Date columns ({len(date_columns)}): {date_columns}")

        # Parse each date column to extract year and month abbreviation
        month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                      'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

        year_headers = []
        month_headers = []

        for date_col in date_columns:
            parts = str(date_col).split('-')
            year = int(parts[0])
            month_num = int(parts[1])
            year_headers.append(year)
            month_headers.append(month_names[month_num - 1])

        print(f"   Year headers: {year_headers}")
        print(f"   Month headers: {month_headers}")

        # Column C = 3 (first data column)
        first_data_col_num = 3
        last_data_col_letter = self._col_number_to_letter(first_data_col_num + len(date_columns) - 1)

        # Disable screen updating, auto-calculation, alerts, and events for performance
        print(f"   Disabling screen updating and auto-calculation...")
        self.app.screen_updating = False
        original_calculation = self.app.api.Calculation
        self.app.api.Calculation = -4135  # xlCalculationManual
        self.app.api.DisplayAlerts = False
        self.app.api.EnableEvents = False

        try:
            # Step 1: Write year headers to Row 1 (C1 onwards)
            print(f"   Writing year headers to row {year_row} (C{year_row}:{last_data_col_letter}{year_row})...")

            # Clear contents first, then unmerge (avoids dialog prompts on merged cells)
            year_range = sheet.range(f'C{year_row}:{last_data_col_letter}{year_row}')
            year_range.clear_contents()
            try:
                year_range.api.UnMerge()
            except:
                pass

            # Write all year values at once (batch - cells are now unmerged)
            sheet.range(f'C{year_row}').value = [year_headers]
            print(f"   Year values written")

            # Merge consecutive columns that share the same year
            i = 0
            while i < len(year_headers):
                j = i + 1
                while j < len(year_headers) and year_headers[j] == year_headers[i]:
                    j += 1
                if j - i > 1:
                    start_col = self._col_number_to_letter(first_data_col_num + i)
                    end_col = self._col_number_to_letter(first_data_col_num + j - 1)
                    merge_range = f'{start_col}{year_row}:{end_col}{year_row}'
                    sheet.range(merge_range).api.Merge()
                    print(f"     Merged {merge_range} = {year_headers[i]}")
                i = j

            # Step 2: Write month abbreviation headers to Row 2 (C2 onwards)
            print(f"   Writing month headers to row {month_row} (C{month_row}:{last_data_col_letter}{month_row})...")
            sheet.range(f'C{month_row}').value = [month_headers]

            # Write DC Bucket header to S2 (only S2, preserve original P2:R2 headers)
            sheet.range(f'S{month_row}').value = 'DC Bucket'
            print(f"   DC Bucket header written to S{month_row}")

            # Step 3: Clear existing data from data_start_row down (including formula cols P-S)
            last_row = sheet.range(f'{contract_col}{data_start_row}').end('down').row
            if last_row < 1000000 and last_row >= data_start_row:
                clear_range = f'{contract_col}{data_start_row}:S{last_row}'
                print(f"   Clearing data range {clear_range}...")
                sheet.range(clear_range).clear_contents()

            if pivot_df.empty:
                print(f"   No data to write")
                return

            num_rows = len(pivot_df)
            print(f"   Writing {num_rows} rows of data starting at row {data_start_row}...")

            # Step 4: Write CONTRACT_NO_NOLASTDIG to column A
            print(f"   Writing column A (CONTRACT_NO)...")
            contract_values = pivot_df['CONTRACT_NO_NOLASTDIG'].values.reshape(-1, 1).tolist()
            sheet.range(f'{contract_col}{data_start_row}').value = contract_values

            # Step 5: Write PD_CATEGORY to column B
            print(f"   Writing column B (PD_CATEGORY)...")
            category_values = pivot_df['PD_CATEGORY'].values.reshape(-1, 1).tolist()
            sheet.range(f'{pd_category_col}{data_start_row}').value = category_values

            # Step 6: Write all date values at once (columns C through last_data_col)
            print(f"   Writing columns C-{last_data_col_letter} (date values)...")
            data_values = pivot_df[date_columns].values.tolist()
            sheet.range(f'C{data_start_row}').value = data_values

            print(f"   Successfully wrote {num_rows} rows x {len(date_columns) + 2} columns")
            print(f"   Year headers: Row {year_row} (C-{last_data_col_letter})")
            print(f"   Month headers: Row {month_row} (C-{last_data_col_letter})")
            print(f"   Data: Rows {data_start_row}-{data_start_row + num_rows - 1} (A-{last_data_col_letter})")

            # Step 7: Write formulas to columns P, Q, R, S
            last_data_row = data_start_row + num_rows - 1
            r = data_start_row  # first formula row
            lcl = last_data_col_letter  # e.g. 'O'
            print(f"   Writing formulas to columns P-S (rows {r}-{last_data_row})...")

            # Set formulas in the first data row
            sheet.range(f'P{r}').formula = f'=SUM(C{r}:{lcl}{r})'
            sheet.range(f'Q{r}').formula = f'=MAX(C{r}:{lcl}{r})'
            sheet.range(f'R{r}').formula = f'=INDEX(C{r}:{lcl}{r},MATCH(TRUE,INDEX((C{r}:{lcl}{r}<>""),0),0))'
            sheet.range(f'S{r}').formula = f'=LOOKUP(2,1/(C{r}:{lcl}{r}<>""),C{r}:{lcl}{r})'

            # AutoFill down to last data row (single native Excel call)
            if last_data_row > data_start_row:
                source = sheet.range(f'P{r}:S{r}')
                dest = sheet.range(f'P{r}:S{last_data_row}')
                source.api.AutoFill(Destination=dest.api, Type=0)  # xlFillDefault

            print(f"   Formulas written to P{r}:S{last_data_row}")

        finally:
            # Re-enable screen updating, auto-calculation, alerts, and events
            print(f"   Re-enabling screen updating and auto-calculation...")
            self.app.api.EnableEvents = True
            self.app.api.DisplayAlerts = True
            self.app.api.Calculation = original_calculation
            self.app.screen_updating = True

        print(f"   Historic PD update complete!")
    
    def _col_number_to_letter(self, col_num: int) -> str:
        """Convert column number to Excel column letter (e.g., 1->A, 27->AA)"""
        result = ""
        while col_num > 0:
            col_num -= 1
            result = chr(col_num % 26 + 65) + result
            col_num //= 26
        return result

    def setup_historic_pivot_tables(self, pivot_sheet_name: str, data_sheet_name: str,
                                 big_pivot_name: str, small_pivot_name: str,
                                 last_month_field: str, last_data_row: int) -> None:
        """
        Set up pivot tables in Historic PD file after data write.

        Steps:
        1. Delete 'Select PD Category' slicer
        2. Select E4 (PivotTable2) > Change Data Source to 02.Working
        3. Select O4 (PivotTable1) > Change Data Source to 02.Working
        4. Refresh PivotTable2 then PivotTable1
        5. Select O4 (PivotTable1) > Field List > drag last month to Rows
        6. Click O4 > filter > untick (blank)
        """
        import time
        import win32com.client

        print(f"\n   Setting up pivot tables in '{pivot_sheet_name}'...")

        pivot_sheet = self.workbook.sheets[pivot_sheet_name]
        source_range = f"'{data_sheet_name}'!$A$2:$S${last_data_row}"

        try:
            # ====================================================================
            # Step 1: Delete Select PD Category slicer
            # ====================================================================
            print(f"   Step 1: Deleting existing slicers...")
            try:
                while self.workbook.api.SlicerCaches.Count > 0:
                    self.workbook.api.SlicerCaches(1).Delete()
                print(f"     Slicers deleted")
            except Exception as e:
                print(f"     No slicers to delete: {e}")

            # ====================================================================
            # Step 2: Select E4 (PivotTable2) and change data source
            # ====================================================================
            print(f"   Step 2: Selecting E4 ({big_pivot_name}) > Change Data Source...")
            pivot_sheet.range("E4").select()
            big_pt = pivot_sheet.api.PivotTables(big_pivot_name)
            new_cache_big = self.workbook.api.PivotCaches().Create(
                SourceType=1, SourceData=source_range
            )
            big_pt.ChangePivotCache(new_cache_big)
            print(f"     {big_pivot_name} source updated to {source_range}")

            # ====================================================================
            # Step 3: Select O4 (PivotTable1) and change data source
            # ====================================================================
            print(f"   Step 3: Selecting O4 ({small_pivot_name}) > Change Data Source...")
            pivot_sheet.range("O4").select()
            small_pt = pivot_sheet.api.PivotTables(small_pivot_name)
            new_cache_small = self.workbook.api.PivotCaches().Create(
                SourceType=1, SourceData=source_range
            )
            small_pt.ChangePivotCache(new_cache_small)
            print(f"     {small_pivot_name} source updated to {source_range}")

            # ====================================================================
            # Step 4: Refresh PivotTable2 then PivotTable1
            # ====================================================================
            print(f"   Step 4: Refreshing pivots...")
            self.refresh_pivot_table(pivot_sheet_name, big_pivot_name)
            self.refresh_pivot_table(pivot_sheet_name, small_pivot_name)
            print(f"     Both pivots refreshed")

            # ====================================================================
            # Use win32com directly for Steps 5-6
            # (bypasses xlwings COMRetryMethodWrapper that blocks PivotFields)
            # ====================================================================
            excel_com = win32com.client.GetActiveObject("Excel.Application")
            ws_com = excel_com.ActiveWorkbook.Worksheets(pivot_sheet_name)
            small_pt_com = ws_com.PivotTables(small_pivot_name)

            def _get_field(pt, name):
                try:
                    return pt.PivotFields(name)
                except (TypeError, AttributeError):
                    return pt.PivotFields.Item(name)

            def _get_item(pivot_items, idx):
                try:
                    return pivot_items(idx)
                except (TypeError, AttributeError):
                    return pivot_items.Item(idx)

            # ====================================================================
            # Step 5: Select O4 (PivotTable1) > Field List > drag last month
            #         (e.g. Sep2) to Rows area, then wait 3 seconds
            # ====================================================================
            print(f"   Step 5: Adding '{last_month_field}' to {small_pivot_name} Rows...")
            pivot_sheet.range("O4").select()
            field = _get_field(small_pt_com, last_month_field)
            try:
                field.Orientation = 1  # xlRowField
                print(f"     '{last_month_field}' added to Rows")
            except Exception as e:
                print(f"     Error adding field: {e}")
                import traceback
                traceback.print_exc()

            print(f"     Waiting 3 seconds...")
            time.sleep(3)

            # ====================================================================
            # Step 6: Filter (blank) out of the small pivot via Filters trick
            #   6a. Click O4 (PivotTable1)
            #   6b. Drag Sep2 from Rows → Filters
            #   6c. Click P2 filter, tick "Select Multiple Items"
            #   6d. Untick (blank), hit OK
            #   6e. Drag Sep2 from Filters → back to Rows
            # ====================================================================
            print(f"   Step 6: Filtering (blank) via Filters on {small_pivot_name}...")

            # 6a. Click O4
            pivot_sheet.range("O4").select()
            time.sleep(1)

            # 6b. Move Sep2 from Rows to Filters (Page field)
            print(f"     Moving '{last_month_field}' from Rows to Filters...")
            try:
                field.Orientation = 3  # xlPageField (Filters area)
                print(f"     '{last_month_field}' moved to Filters")
            except Exception as e:
                print(f"     Error moving to Filters: {e}")
                import traceback
                traceback.print_exc()
            time.sleep(1)

            # 6c. Click P2 (filter dropdown) and enable Select Multiple Items
            print(f"     Enabling Select Multiple Items...")
            pivot_sheet.range("P2").select()
            try:
                field.EnableMultiplePageItems = True
                print(f"     Select Multiple Items enabled")
            except Exception as e:
                print(f"     Error enabling multi-select: {e}")
                import traceback
                traceback.print_exc()
            time.sleep(1)

            # 6d. Untick (blank) from items: 1, 2, 3, 4, 5, (blank)
            print(f"     Unticking (blank)...")
            try:
                try:
                    pi_items = field.PivotItems()
                except TypeError:
                    pi_items = field.PivotItems
                item_count = pi_items.Count
                print(f"     '{last_month_field}' has {item_count} items")

                blank_hidden = False
                for i in range(1, item_count + 1):
                    try:
                        pi = _get_item(pi_items, i)
                        name = str(pi.Name)
                        print(f"       Item {i}: '{name}'")
                        if name.lower() in ['(blank)', 'blank', '']:
                            print(f"     >>> Unticked: '{name}'")
                            pi.Visible = False
                            blank_hidden = True
                    except Exception as e:
                        print(f"       Error with item {i}: {e}")

                # Fallback: blank is typically the last item
                if not blank_hidden and item_count > 1:
                    try:
                        pi = _get_item(pi_items, item_count)
                        print(f"     >>> FALLBACK: Unticking item {item_count} '{pi.Name}'")
                        pi.Visible = False
                    except Exception as e:
                        print(f"     Fallback failed: {e}")

                print(f"     (blank) unticked")

            except Exception as e:
                print(f"     ERROR unticking (blank): {e}")
                import traceback
                traceback.print_exc()
            time.sleep(1)

            # 6e. Move Sep2 from Filters back to Rows
            print(f"     Moving '{last_month_field}' from Filters back to Rows...")
            try:
                field.Orientation = 1  # xlRowField
                print(f"     '{last_month_field}' back in Rows")
            except Exception as e:
                print(f"     Error moving back to Rows: {e}")
                import traceback
                traceback.print_exc()
            time.sleep(1)

            # ====================================================================
            # Step 7: Click big pivot > Insert Slicer > tick PD_CATEGORY
            #         Move slicer next to big pivot
            # ====================================================================
            print(f"   Step 7: Adding PD_CATEGORY slicer to {big_pivot_name}...")
            pivot_sheet.range("E4").select()
            try:
                slicer_cache = self.workbook.api.SlicerCaches.Add2(big_pt, "PD_CATEGORY")
                slicer = slicer_cache.Slicers.Add(pivot_sheet.api)
                slicer.Caption = "Select PD Category"

                # Position slicer next to the big pivot (to its right)
                big_pt_range = big_pt.TableRange2
                slicer.Width = 144
                slicer.Height = 200
                slicer.Left = big_pt_range.Left + big_pt_range.Width + 10
                slicer.Top = big_pt_range.Top

                print(f"     PD_CATEGORY slicer added next to {big_pivot_name}")

            except Exception as e:
                print(f"     ERROR adding slicer: {e}")
                import traceback
                traceback.print_exc()

        except Exception as e:
            print(f"\n   FATAL ERROR: {e}")
            import traceback
            traceback.print_exc()
            raise

        print(f"\n   Pivot setup complete!")

    @staticmethod
    def copy_pivot_to_historic(latest_pd_file: str,
                              input_folder: str,
                              output_folder: str,
                              historic_filename: str = "02. Historic PD Calculation 2024-25.xlsb",
                              pivot_sheet: str = "01.Pivoted_Portfolio",
                              historic_sheet: str = "02.Working") -> str:
        """
        Complete workflow to copy pivot data to historic PD format.
        
        Args:
            latest_pd_file: Path to the latest PD file with pivot data
            input_folder: Folder containing the original historic file
            output_folder: Folder to save the updated historic file
            historic_filename: Name of the historic file (default: "02. Historic PD Calculation 2024-25.xlsb")
            pivot_sheet: Sheet name with pivot data (default: "01.Pivoted_Portfolio")
            historic_sheet: Sheet name to write historic format (default: "02.Working")
        
        Returns:
            str: Path to the saved historic file
        """
        print("\n" + "="*80)
        print("COPYING PIVOT TO HISTORIC PD FORMAT")
        print("="*80)
        
        historic_input_file = os.path.join(input_folder, historic_filename)
        
        # Step 1: Extract pivot data from latest PD file
        print(f"\n1. Opening PD file to extract pivot...")
        print(f"   File: {os.path.basename(latest_pd_file)}")
        
        with ExcelPortfolioAutomation(latest_pd_file, visible=True) as excel:
            pivot_df = excel.extract_pivot_table_to_dataframe(pivot_sheet)
            
            if pivot_df.empty:
                print("\n   ERROR: No pivot data extracted!")
                return None
        
        print(f"   Extracted pivot data: {len(pivot_df)} rows")
        
        # Step 2: Open historic file and write data
        print(f"\n2. Opening Historic PD file...")
        print(f"   File: {os.path.basename(historic_input_file)}")
        
        with ExcelPortfolioAutomation(historic_input_file, visible=True) as excel:
            # Write in historic format
            excel.write_historic_pd_format(historic_sheet, pivot_df)

            # Step 3: Set up pivot tables in 03.PD_Pivot
            print(f"\n3. Setting up pivot tables...")

            # Compute last month field name from pivot date columns
            date_columns = [col for col in pivot_df.columns
                           if col not in ['CONTRACT_NO_NOLASTDIG', 'PD_CATEGORY']
                           and re.match(r'\d{4}-\d{2}', str(col))]
            last_date = date_columns[-1]  # e.g., '2025-09'
            month_num = int(last_date.split('-')[1])
            month_names_list = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                               'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            last_month_field = f"{month_names_list[month_num - 1]}2"  # e.g., 'Sep2'
            last_data_row = 3 + len(pivot_df) - 1  # data starts at row 3

            excel.setup_historic_pivot_tables(
                pivot_sheet_name="03.PD_Pivot",
                data_sheet_name=historic_sheet,
                big_pivot_name="PivotTable2",
                small_pivot_name="PivotTable1",
                last_month_field=last_month_field,
                last_data_row=last_data_row
            )

            # Save as new file with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = historic_filename.replace(".xlsb", f"_Updated_{timestamp}.xlsb")
            output_file = os.path.join(output_folder, output_filename)

            excel.save_as(output_file)
            print(f"\n4. Saved Historic file:")
            print(f"   {output_file}")
        
        print("\n" + "="*80)
        print("HISTORIC PD UPDATE COMPLETED!")
        print("="*80)
        
        return output_file