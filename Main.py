"""
Main script to automate Excel pivot table operations
Workflow:
1. Update Portfolio_1: Delete first 6 months, keep most recent month + append all Portfolio_2 data
2. Extract 6 summary files into DataFrame
3. Paste mapped data to Portfolio_2 sheet (filter empty rows)
4. Refresh pivot table in 01.Pivoted_Portfolio
"""
import os
import sys
from datetime import datetime

# Add the Class directory to the path so we can import BasicExcelFunctionsClass
class_path = os.path.join(os.path.dirname(__file__), 'Scripts', 'Class')
sys.path.insert(0, class_path)

from BasicExcelFunctionsClass import ExcelPortfolioAutomation
import pandas as pd


def extract_summary_files_to_dataframe(input_folder: str) -> pd.DataFrame:
    """
    Step 2: Extract data from 6 summary files into a DataFrame

    Returns:
        DataFrame with columns: CONTRACT_NO, EQT_DESC, PD_CATEGORY, DPD, MONTH
    """
    print("="*80)
    print("STEP 2: EXTRACTING SUMMARY FILES TO DATAFRAME")
    print("="*80)
    print(f"Input folder: {input_folder}")
    print()

    # Use the static method to consolidate summary files
    consolidated_df = ExcelPortfolioAutomation.consolidate_summary_files(
        input_folder=input_folder,
        file_pattern="3. Summary_*.xlsb",
        sheet_name="SUMMARY",
        header_row=0
    )

    if not consolidated_df.empty:
        print(f"\nExtracted {len(consolidated_df)} rows from summary files")

        # Step 4: Filter out empty rows where CONTRACT_NO, EQT_DESC, PD_CATEGORY, DPD are all empty
        # but MONTH has a value (like "-, -, -, -, 09/30/2025")
        print("\nFiltering out empty data rows...")

        before_filter = len(consolidated_df)

        # Check if key data columns have actual values (not null, not empty string, not just dashes)
        def is_valid_value(val):
            if pd.isna(val):
                return False
            if str(val).strip() in ['', '-', 'nan', 'None']:
                return False
            return True

        # Keep rows where at least CONTRACT_NO has a valid value
        consolidated_df = consolidated_df[
            consolidated_df['CONTRACT_NO'].apply(is_valid_value)
        ]

        after_filter = len(consolidated_df)
        print(f"Filtered out {before_filter - after_filter} empty rows")
        print(f"Remaining rows: {after_filter}")

        return consolidated_df
    else:
        print("No data extracted from summary files")
        return pd.DataFrame()


def run_pd_automation():
    """
    Main automation workflow:
    1. Update Portfolio_1 (delete first 6 months, keep most recent + append Portfolio_2)
    2. Extract summary files to DataFrame
    3. Paste mapped data to Portfolio_2
    4. Refresh pivot table
    """
    # Define paths
    input_folder = r"C:\Users\Ashen Alwis\Desktop\Impairment Claculation\Input Files\PD"
    output_folder = r"C:\Users\Ashen Alwis\Desktop\Impairment Claculation\OutPut\PD"
    pd_file = os.path.join(input_folder, "01. PD_data_2024-25.xlsb")

    # Create output filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(output_folder, f"01. PD_data_2024-25_Updated_{timestamp}.xlsb")

    # Ensure output directory exists
    os.makedirs(output_folder, exist_ok=True)

    print("\n" + "="*80)
    print("PD AUTOMATION WORKFLOW")
    print("="*80)
    print(f"PD File: {pd_file}")
    print(f"Output File: {output_file}")
    print()

    # =========================================================================
    # STEP 2: Extract summary files to DataFrame (do this first, before opening Excel)
    # =========================================================================
    summary_df = extract_summary_files_to_dataframe(input_folder)

    if summary_df.empty:
        print("\nError: No data extracted from summary files. Aborting.")
        return False

    # Preview the data
    print("\nData preview (first 5 rows):")
    print(summary_df.head())
    print(f"\nColumns: {list(summary_df.columns)}")

    # =========================================================================
    # Open Excel workbook for Steps 1, 3, 4
    # =========================================================================
    with ExcelPortfolioAutomation(pd_file, visible=True) as excel:

        # =====================================================================
        # STEP 1: Update Portfolio_1 (delete first 6 months, keep most recent + append Portfolio_2)
        # =====================================================================
        print("\n" + "="*80)
        print("STEP 1: UPDATING PORTFOLIO_1")
        print("="*80)

        # Read CURRENT Portfolio_1 data
        print("Reading current Portfolio_1 data...")
        sheet_p1 = excel.workbook.sheets['Portfolio_1']
        last_row_p1 = sheet_p1.range('A2').end('down').row
        print(f"Last row with data in Portfolio_1: {last_row_p1}")

        data_range_p1 = f"A1:F{last_row_p1}"
        portfolio_1_data = excel.read_sheet_range_to_dataframe(
            sheet_name='Portfolio_1',
            range_address=data_range_p1
        )
        print(f"Read {len(portfolio_1_data)} rows from Portfolio_1")

        # Filter Portfolio_1 to keep ONLY the most recent month (delete first 6 months)
        portfolio_1_filtered = pd.DataFrame()
        if not portfolio_1_data.empty and 'MONTH' in portfolio_1_data.columns:
            # Convert MONTH column to datetime for proper sorting
            portfolio_1_data['MONTH_DATE'] = pd.to_datetime(portfolio_1_data['MONTH'], format='%m/%d/%Y', errors='coerce')
            
            # Find the most recent month in Portfolio_1
            most_recent_month_p1 = portfolio_1_data['MONTH_DATE'].max()
            print(f"\nMost recent month in Portfolio_1: {most_recent_month_p1.strftime('%m/%d/%Y') if pd.notna(most_recent_month_p1) else 'N/A'}")
            
            # Keep ONLY the most recent month's data (delete first 6 months)
            portfolio_1_filtered = portfolio_1_data[portfolio_1_data['MONTH_DATE'] == most_recent_month_p1].copy()
            
            # Drop the temporary MONTH_DATE column
            portfolio_1_filtered = portfolio_1_filtered.drop(columns=['MONTH_DATE'])
            
            print(f"Keeping {len(portfolio_1_filtered)} rows from most recent month")
            print(f"Deleting {len(portfolio_1_data) - len(portfolio_1_filtered)} rows from older months")
        else:
            print("Warning: Could not filter Portfolio_1 by month")

        # Read ALL Portfolio_2 data
        print("\nReading Portfolio_2 data...")
        sheet_p2 = excel.workbook.sheets['Portfolio_2']
        last_row_p2 = sheet_p2.range('A2').end('down').row
        print(f"Last row with data in Portfolio_2: {last_row_p2}")

        data_range_p2 = f"A1:F{last_row_p2}"
        portfolio_2_data = excel.read_sheet_range_to_dataframe(
            sheet_name='Portfolio_2',
            range_address=data_range_p2
        )
        print(f"Read {len(portfolio_2_data)} rows from Portfolio_2")

        # Combine: Portfolio_1 most recent month + ALL Portfolio_2 data
        print("\nCombining Portfolio_1 (most recent month) + ALL Portfolio_2 data...")
        combined_portfolio_1 = pd.concat([portfolio_1_filtered, portfolio_2_data], ignore_index=True)
        print(f"Combined: {len(portfolio_1_filtered)} (P1) + {len(portfolio_2_data)} (P2) = {len(combined_portfolio_1)} total rows")

        # Clear Portfolio_1 and write the combined data
        print("\nClearing Portfolio_1...")
        try:
            excel.clear_range_dynamic(
                sheet_name='Portfolio_1',
                start_cell='A2',
                end_column='F'
            )
        except Exception as e:
            print(f"Note: {e}")

        print("Writing updated data to Portfolio_1...")
        excel.write_dataframe_to_sheet(
            sheet_name='Portfolio_1',
            start_cell='A1',
            df=combined_portfolio_1,
            include_headers=True,
            clear_existing=False
        )
        print(f"Portfolio_1 updated successfully with {len(combined_portfolio_1)} rows!")

        # =====================================================================
        # STEP 3: Paste mapped data to Portfolio_2
        # =====================================================================
        print("\n" + "="*80)
        print("STEP 3: PASTING SUMMARY DATA TO PORTFOLIO_2")
        print("="*80)

        # Portfolio_2 table column order:
        # A: MONTH, B: CONTRACT_NO, C: CONTRACT_NO_NOLASTDI, D: EQT_DESC, E: PD_CATEGORY, F: DPD

        # Create CONTRACT_NO_NOLASTDI (CONTRACT_NO without last digit)
        summary_df['CONTRACT_NO_NOLASTDI'] = summary_df['CONTRACT_NO'].apply(
            lambda x: str(x)[:-1] if pd.notna(x) and len(str(x)) > 0 else ''
        )

        # Prepare data for Portfolio_2 - select and order columns to match table structure
        portfolio_2_columns = ['MONTH', 'CONTRACT_NO', 'CONTRACT_NO_NOLASTDI', 'EQT_DESC', 'PD_CATEGORY', 'DPD']

        # Select only the mapped columns in the correct order
        mapped_df = summary_df[portfolio_2_columns].copy()

        print(f"Mapped DataFrame has {len(mapped_df)} rows")
        print(f"Columns to write: {list(mapped_df.columns)}")

        # Clear existing data in Portfolio_2 (keep the table structure/headers)
        print("\nClearing existing data in Portfolio_2 (A2:F)...")
        try:
            excel.clear_range_dynamic(
                sheet_name='Portfolio_2',
                start_cell='A2',
                end_column='F'
            )
        except Exception as e:
            print(f"Note: {e}")

        # Write the mapped data to Portfolio_2 (starting at A2 to preserve headers)
        print("Writing summary data to Portfolio_2...")
        sheet_p2 = excel.workbook.sheets['Portfolio_2']

        # Write data without headers (table already has headers)
        data_values = mapped_df.values.tolist()
        if data_values:
            sheet_p2.range('A2').value = data_values
            print(f"Successfully wrote {len(data_values)} rows to Portfolio_2")

        # =====================================================================
        # STEP 4: Refresh pivot table
        # =====================================================================
        print("\n" + "="*80)
        print("STEP 4: REFRESHING PIVOT TABLE")
        print("="*80)

        try:
            excel.refresh_pivot_table(
                sheet_name='01.Pivoted_Portfolio',
                pivot_table_name='PivotTable1'
            )
            print("Pivot table refreshed successfully!")
        except Exception as e:
            print(f"Error refreshing pivot table: {e}")
            import traceback
            traceback.print_exc()

        # =====================================================================
        # Save the file
        # =====================================================================
        print("\n" + "="*80)
        print("SAVING FILE")
        print("="*80)

        excel.save_as(output_file)
        print(f"File saved to: {output_file}")

    print("\n" + "="*80)
    print("AUTOMATION COMPLETED SUCCESSFULLY!")
    print("="*80)

    return True


if __name__ == "__main__":
    try:
        success = run_pd_automation()
        if not success:
            sys.exit(1)
    except Exception as e:
        print(f"\nError occurred: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)