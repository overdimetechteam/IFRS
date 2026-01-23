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
from typing import Optional, Tuple
import re



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


