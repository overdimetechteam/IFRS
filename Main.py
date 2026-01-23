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


def copy_pivot_between_sheets():
    """
    Copy pivot table data from Portfolio_2 to Portfolio_1 sheet
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
    print("Starting Pivot Table Copy Automation")
    print("="*80)
    print(f"Input file: {input_file}")
    print(f"Output file: {output_file}")
    print()

    # Use context manager to handle Excel operations
    with ExcelPortfolioAutomation(input_file, visible=True) as excel:

        # Step 1: Read pivot table data from Portfolio_2 sheet
        print("Step 1: Reading pivot table data from Portfolio_2 sheet...")
        portfolio_2_data = excel.read_sheet_range_to_dataframe(
            sheet_name='Portfolio_2',
            range_address=None  # Read entire used range (including pivot table)
        )
        print(f"   Data shape: {portfolio_2_data.shape}")
        print()

        # Step 2: Clear all existing data in Portfolio_1 sheet
        print("Step 2: Clearing all existing data in Portfolio_1 sheet...")
        try:
            excel.clear_range_dynamic(
                sheet_name='Portfolio_1',
                start_cell='A1',
                end_column='ZZ'  # Clear all columns
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

        # Step 4: Save as new file
        print("Step 4: Saving modified workbook...")
        excel.save_as(output_file)
        print()

    print("="*80)
    print("Automation completed successfully!")
    print(f"Output saved to: {output_file}")
    print("="*80)


if __name__ == "__main__":
    try:
        copy_pivot_between_sheets()
    except Exception as e:
        print(f"\nError occurred: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
