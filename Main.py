"""
Main script to automate Excel pivot table operations
Task: Copy pivot table from Portfolio_2 sheet to Portfolio_1 sheet
"""
import os
import sys
from datetime import datetime

# Add the Class directory to the path so we can import BasicExcelFunctionsClass
class_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'Class')
sys.path.insert(0, class_path)

from BasicExcelFunctionsClass import ExcelPortfolioAutomation
import pandas as pd


def consolidate_summary_files_to_test_file():
    """
    Consolidate data from 6 summary files and save to a test Excel file
    """
    # Define paths
    input_folder = r"C:\Users\Ashen Alwis\Desktop\Impairment Claculation\Impairment Claculation\Input Files\PD"
    output_folder = r"C:\Users\Ashen Alwis\Desktop\Impairment Claculation\Impairment Claculation\OutPut\PD"

    # Create output filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    test_output_file = os.path.join(output_folder, f"Test_Consolidated_Summary_{timestamp}.xlsx")

    # Ensure output directory exists
    os.makedirs(output_folder, exist_ok=True)

    print("="*80)
    print("Starting Summary Files Consolidation")
    print("="*80)
    print(f"Input folder: {input_folder}")
    print(f"Test output file: {test_output_file}")
    print()

    # Use the static method to consolidate summary files
    # Try different header rows (start from row 0 and search up to row 10)
    consolidated_df = ExcelPortfolioAutomation.consolidate_summary_files(
        input_folder=input_folder,
        file_pattern="3. Summary_*.xlsb",
        sheet_name="SUMMARY",
        header_row=0  # Will auto-search for the correct header row
    )

    if not consolidated_df.empty:
        # Save to test Excel file
        print(f"   Saving consolidated data to test file: {test_output_file}")
        with pd.ExcelWriter(test_output_file, engine='openpyxl') as writer:
            consolidated_df.to_excel(writer, sheet_name='Consolidated_Data', index=False)
        print(f"    Test file saved successfully!")
        print()

        print("="*80)
        print("Summary Consolidation completed successfully!")
        print(f"Total records: {len(consolidated_df)}")
        print(f"Test file saved to: {test_output_file}")
        print("="*80)
        print()

        return consolidated_df
    else:
        print("   No data to save")
        print()
        return None


def copy_data_between_sheets():
    """
    Copy table data from Portfolio_2 to Portfolio_1 sheet and refresh pivot tabl
    """
    # Define file paths
    input_file = r"C:\Users\Ashen Alwis\Desktop\Impairment Claculation\Impairment Claculation\Input Files\PD\01. PD_data_2024-25.xlsb"
    output_folder = r"C:\Users\Ashen Alwis\Desktop\Impairment Claculation\Impairment Claculation\OutPut\PD"

    # Create output filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(output_folder, f"01. PD_data_2024-25_Updated_{timestamp}.xlsb")

    # Ensure output directory exists
    os.makedirs(output_folder, exist_ok=True)

    print("="*80)
    print("Starting Data Copy Automation")
    print("="*80)
    print(f"Input file: {input_file}")
    print(f"Output file: {output_file}")
    print()

    # Use context manager to handle Excel operations
    with ExcelPortfolioAutomation(input_file, visible=True) as excel:

        # Step 1: Read table data from Portfolio_2 sheet (only actual data, not empty rows)
        print("Step 1: Reading table data from Portfolio_2 sheet...")

        # First, find the actual last row with data in Portfolio_2
        sheet_p2 = excel.workbook.sheets['Portfolio_2']
        last_row_p2 = sheet_p2.range('A2').end('down').row
        print(f"   Detected last row with data in Portfolio_2: {last_row_p2}")

        # Read only the range with actual data (A1:H + last_row)
        data_range = f"A1:H{last_row_p2}"
        print(f"   Reading range: {data_range}")

        portfolio_2_data = excel.read_sheet_range_to_dataframe(
            sheet_name='Portfolio_2',
            range_address=data_range
        )
        print(f"   Data shape: {portfolio_2_data.shape}")
        print()

        # Step 2: Clear existing data in Portfolio_1 sheet (A2 to H, last row)
        print("Step 2: Clearing existing data in Portfolio_1 sheet (A2:H)...")
        try:
            excel.clear_range_dynamic(
                sheet_name='Portfolio_1',
                start_cell='A2',
                end_column='H'
            )
        except Exception as e:
            print(f"   Note: {e}")
        print()

        # Step 3: Write the data to Portfolio_1 sheet
        print("Step 3: Writing data to Portfolio_1 sheet...")
        excel.write_dataframe_to_sheet(
            sheet_name='Portfolio_1',
            start_cell='A1',
            df=portfolio_2_data,
            include_headers=True,
            clear_existing=False  # Already cleared above
        )
        print()

        # Step 4: Refresh pivot table in 01.Pivoted_Portfolio sheet
        print("Step 4: Refreshing pivot table in 01.Pivoted_Portfolio...")
        try:
            excel.refresh_pivot_table(
                sheet_name='01.Pivoted_Portfolio',
                pivot_table_name='PivotTable1'
            )
        except Exception as e:
            print(f"   Error refreshing pivot table: {e}")
            import traceback
            traceback.print_exc()
        print()

        # Step 5: Save as new file
        print("Step 5: Saving modified workbook...")
        excel.save_as(output_file)
        print()

    print("="*80)
    print("Automation completed successfully!")
    print(f"Output saved to: {output_file}")
    print("="*80)


if __name__ == "__main__":
    try:
        # Step 0: Consolidate summary files first (test run)
        print("\n" + "="*80)
        print("STEP 0: CONSOLIDATING SUMMARY FILES (TEST)")
        print("="*80 + "\n")

        consolidated_df = consolidate_summary_files_to_test_file()

        if consolidated_df is not None:
            print("\n" + "="*80)
            print("STEP 1: COPYING DATA BETWEEN SHEETS")
            print("="*80 + "\n")

            # Proceed with the main workflow
            copy_data_between_sheets()
        else:
            print("\nError: Could not consolidate summary files. Aborting.")
            sys.exit(1)

    except Exception as e:
        print(f"\nError occurred: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
