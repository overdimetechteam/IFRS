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

        # Step 1: Read pivot table data from Portfolio_2 sheet (only actual data, not empty rows)
        print("Step 1: Reading pivot table data from Portfolio_2 sheet...")

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

        # Step 5: Find all pivot tables in the workbook first
        print("Step 5: Finding all pivot tables in workbook...")
        all_pivots = excel.find_all_pivot_tables()
        print()

        # Step 6: Update pivot tables that reference Portfolio_1 data
        print("Step 6: Updating pivot tables that reference Portfolio_1...")
        try:
            # Get the actual last row with data in Portfolio_1
            sheet_p1 = excel.workbook.sheets['Portfolio_1']
            last_row_p1 = sheet_p1.range('A2').end('down').row
            print(f"   Last row with data in Portfolio_1: {last_row_p1}")

            workbook_name = excel.workbook.name
            updated_count = 0

            # Iterate through all sheets that have pivot tables
            if all_pivots:
                for sheet_name, pivot_names in all_pivots.items():
                    sheet = excel.workbook.sheets[sheet_name]
                    pivot_tables = sheet.api.PivotTables()

                    for i in range(1, pivot_tables.Count + 1):
                        pivot_table = pivot_tables(i)
                        pivot_table_name = pivot_table.Name

                        # Check if this pivot table's source references Portfolio_1
                        try:
                            source_data = pivot_table.SourceData
                            print(f"   Checking pivot '{pivot_table_name}' in sheet '{sheet_name}'")
                            print(f"     Current source: {source_data}")

                            if 'Portfolio_1' in source_data:
                                # Update the source range
                                new_source_range = f"Portfolio_1!$A$1:$H${last_row_p1}"
                                full_source = f"'{workbook_name}'!{new_source_range}"

                                print(f"     Updating to: {new_source_range}")

                                pivot_table.ChangePivotCache(
                                    excel.workbook.api.PivotCaches().Create(
                                        SourceType=1,  # xlDatabase
                                        SourceData=full_source
                                    )
                                )

                                # Refresh the pivot table
                                pivot_table.RefreshTable()
                                print(f"     Successfully updated and refreshed!")
                                updated_count += 1
                            else:
                                print(f"     Skipping (does not reference Portfolio_1)")
                        except Exception as e:
                            print(f"     Error checking/updating this pivot: {e}")

                print(f"\n   Updated {updated_count} pivot table(s)")
            else:
                print("   No pivot tables found in workbook")
        except Exception as e:
            print(f"   Error updating pivot tables: {e}")
            import traceback
            traceback.print_exc()
        print()

        # Step 7: Save as new file
        print("Step 7: Saving modified workbook...")
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
