"""
PD Portfolio Roll-Forward Automation
=====================================
Dynamic month-driven rollover maintaining:
- Portfolio_2: Latest 6 months
- Portfolio_1: Older 7 months
- Total: Always 13 months

Flow:
1. Read latest month from config file
2. Get user input for new end month
3. Calculate N = months to process
4. Extract N summary files (1 file = 1 month)
5. Shift oldest N months from Portfolio_2 → Portfolio_1
6. Remove oldest N months from Portfolio_1
7. Append N new months to Portfolio_2
8. Refresh pivot table and save
"""
import os
import sys
from datetime import datetime

# Add the Class directory to the path
class_path = os.path.join(os.path.dirname(__file__), 'Scripts', 'Class')
sys.path.insert(0, class_path)

from BasicExcelFunctionsClass import ExcelPortfolioAutomation
import pandas as pd


# =============================================================================
# CONFIGURATION
# =============================================================================
INPUT_FOLDER = r"C:\Users\Ashen Alwis\Desktop\Impairment Claculation\Input Files\PD"
OUTPUT_FOLDER = r"C:\Users\Ashen Alwis\Desktop\Impairment Claculation\OutPut\PD"
ORIGINAL_PD_FILE = os.path.join(INPUT_FOLDER, "01. PD_data_2024-25.xlsb")
CONFIG_FILE = os.path.join(INPUT_FOLDER, "latest_month.txt")

PORTFOLIO_2_MONTHS = 6   # Latest months in Portfolio_2
PORTFOLIO_1_MONTHS = 7   # Older months in Portfolio_1
TOTAL_MONTHS = 13        # Total across both portfolios

# Column mapping: Summary file column -> DataFrame column
COLUMN_MAPPING = {
    'CONTRACT NO': 'CONTRACT_NO',
    'EQUIPMENT DESCRIPTION': 'EQT_DESC',
    'PD/LGD CATEGORY': 'PD_CATEGORY',
    'CLIENT DPD': 'DPD'
}

# Columns to write from DataFrame to Portfolio sheets
DF_TO_PORTFOLIO_COLUMNS = ['MONTH', 'CONTRACT_NO', 'EQT_DESC', 'PD_CATEGORY', 'DPD']

# Column positions in Excel (for write_portfolio_data method)
COLUMN_POSITIONS = {
    'MONTH': 'A',
    'CONTRACT_NO': 'B',
    'EQT_DESC': 'D',
    'PD_CATEGORY': 'E',
    'DPD': 'F'
}


# =============================================================================
# USER INPUT FUNCTION (Remains in Main.py as it's specific to this script)
# =============================================================================
def get_end_month_from_args() -> datetime:
    """Get the new end month from command line argument."""
    if len(sys.argv) < 2:
        print("\nUsage: python Main.py MM/DD/YYYY")
        print("Example: python Main.py 09/30/2025")
        sys.exit(1)

    date_str = sys.argv[1].strip()
    try:
        return ExcelPortfolioAutomation.parse_month_string(date_str)
    except ValueError:
        print(f"\nERROR: Invalid date format: {date_str}")
        print("Expected format: MM/DD/YYYY (e.g., 09/30/2025)")
        sys.exit(1)


# =============================================================================
# MAIN AUTOMATION - PART 1: PORTFOLIO ROLL-FORWARD
# =============================================================================
def run_automation():
    """Main automation workflow - Portfolio Roll-Forward."""

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    print("\n" + "="*80)
    print("PD PORTFOLIO ROLL-FORWARD AUTOMATION")
    print("="*80)

    # Step 1: Read current latest month from config
    try:
        current_latest = ExcelPortfolioAutomation.read_config_file(CONFIG_FILE)
        print(f"\nCurrent latest month: {ExcelPortfolioAutomation.format_month_string(current_latest)}")
    except FileNotFoundError:
        print(f"\nERROR: Config file not found: {CONFIG_FILE}")
        print(f"Please create the file with the latest month in MM/DD/YYYY format.")
        sys.exit(1)
    except ValueError as e:
        print(f"\nERROR: {e}")
        sys.exit(1)

    # Step 2: Get new end month from command line argument
    new_end_month = get_end_month_from_args()
    print(f"New end month: {ExcelPortfolioAutomation.format_month_string(new_end_month)}")

    # Step 3: Calculate months to process
    num_new_months = ExcelPortfolioAutomation.months_between(current_latest, new_end_month)

    if num_new_months <= 0:
        print("\nERROR: New end month must be after current latest month.")
        return False, None

    if num_new_months > PORTFOLIO_2_MONTHS:
        print(f"\nWARNING: {num_new_months} months requested, max is {PORTFOLIO_2_MONTHS}.")
        num_new_months = PORTFOLIO_2_MONTHS

    print(f"\nMonths to process: {num_new_months}")

    # Step 4: Find and extract summary files
    print("\n" + "-"*60)
    print(f"FINDING {num_new_months} SUMMARY FILES")
    print("-"*60)

    summary_files = ExcelPortfolioAutomation.find_summary_files_by_date_range(
        input_folder=INPUT_FOLDER,
        start_month=current_latest,
        num_months=num_new_months,
        file_prefix="3. Summary_",
        file_extension=".xlsb",
        date_format_in_filename='%Y-%m-%d'
    )

    if not summary_files:
        print("\nERROR: No summary files found.")
        return False, None

    print(f"\nExtracting data...")
    new_data = ExcelPortfolioAutomation.extract_data_from_summary_files(
        file_paths=summary_files,
        column_mapping=COLUMN_MAPPING,
        output_columns=DF_TO_PORTFOLIO_COLUMNS,
        sheet_name='SUMMARY',
        date_format='%m/%d/%Y'
    )

    if new_data.empty:
        print("\nERROR: No data extracted.")
        return False, None

    print(f"\nNew data: {len(new_data)} rows")

    # Step 5-7: Open Excel and perform roll-forward
    print("\n" + "-"*60)
    print("PERFORMING ROLL-FORWARD")
    print("-"*60)

    # Get the latest PD file (output file if exists, otherwise original)
    pd_file = ExcelPortfolioAutomation.get_latest_file_in_folder(
        folder_path=OUTPUT_FOLDER,
        file_pattern="01. PD_data_2024-25_Updated_*.xlsb",
        fallback_file=ORIGINAL_PD_FILE
    )

    output_file = None

    with ExcelPortfolioAutomation(pd_file, visible=True) as excel:

        # Read current portfolios using the new instance method
        print("\nReading current portfolios...")
        p1_data = excel.read_portfolio_data('Portfolio_1')
        p2_data = excel.read_portfolio_data('Portfolio_2')

        p1_months = ExcelPortfolioAutomation.get_unique_months_from_dataframe(p1_data)
        p2_months = ExcelPortfolioAutomation.get_unique_months_from_dataframe(p2_data)

        print(f"  Portfolio_1: {len(p1_months)} months, {len(p1_data)} rows")
        print(f"  Portfolio_2: {len(p2_months)} months, {len(p2_data)} rows")

        # Calculate new month distributions
        new_months = ExcelPortfolioAutomation.get_unique_months_from_dataframe(new_data)
        all_months = sorted(set(p1_months + p2_months + new_months))

        # Keep only latest 13 months total
        if len(all_months) > TOTAL_MONTHS:
            all_months = all_months[-TOTAL_MONTHS:]

        # Split: Portfolio_1 gets older 7, Portfolio_2 gets latest 6
        new_p1_months = all_months[:PORTFOLIO_1_MONTHS] if len(all_months) > PORTFOLIO_2_MONTHS else []
        new_p2_months = all_months[-PORTFOLIO_2_MONTHS:] if len(all_months) >= PORTFOLIO_2_MONTHS else all_months

        print(f"\nNew distribution:")
        print(f"  Portfolio_1: {len(new_p1_months)} months")
        print(f"  Portfolio_2: {len(new_p2_months)} months")
        print(f"  Total: {len(new_p1_months) + len(new_p2_months)} months")

        # Combine all data
        all_data = pd.concat([p1_data, p2_data, new_data], ignore_index=True)

        # Remove duplicates (keep latest)
        all_data = all_data.drop_duplicates(subset=['MONTH', 'CONTRACT_NO'], keep='last')

        # Split data by new month distributions using static method
        new_p1_data = ExcelPortfolioAutomation.filter_dataframe_by_months(all_data, new_p1_months)
        new_p2_data = ExcelPortfolioAutomation.filter_dataframe_by_months(all_data, new_p2_months)

        # Sort by month
        new_p1_data = new_p1_data.sort_values('MONTH', key=lambda x: pd.to_datetime(x, format='%m/%d/%Y'))
        new_p2_data = new_p2_data.sort_values('MONTH', key=lambda x: pd.to_datetime(x, format='%m/%d/%Y'))

        # Write to portfolios using the new instance method
        print("\nWriting portfolios...")
        excel.write_portfolio_data('Portfolio_1', new_p1_data, DF_TO_PORTFOLIO_COLUMNS, COLUMN_POSITIONS)
        excel.write_portfolio_data('Portfolio_2', new_p2_data, DF_TO_PORTFOLIO_COLUMNS, COLUMN_POSITIONS)

        # Refresh pivot table
        print("\nRefreshing pivot table...")
        try:
            excel.refresh_pivot_table('01.Pivoted_Portfolio', 'PivotTable1')
            print("  Pivot table refreshed!")
        except Exception as e:
            print(f"  Warning: {e}")

        # Save
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(OUTPUT_FOLDER, f"01. PD_data_2024-25_Updated_{timestamp}.xlsb")
        excel.save_as(output_file)
        print(f"\nSaved: {output_file}")

    # Update config using static method
    ExcelPortfolioAutomation.save_config_file(CONFIG_FILE, new_end_month)

    print("\n" + "="*80)
    print("PORTFOLIO AUTOMATION COMPLETED!")
    print("="*80)

    return True, output_file


# =============================================================================
# MAIN AUTOMATION - PART 2: HISTORIC PD UPDATE
# =============================================================================
def run_historic_update(latest_pd_file: str):
    """
    Second automation - Copy pivot data to Historic PD format.
    
    Args:
        latest_pd_file: Path to the PD file generated by first automation
    
    Returns:
        bool: Success status
    """
    try:
        # Call the static method to perform the historic update
        historic_output_file = ExcelPortfolioAutomation.copy_pivot_to_historic(
            latest_pd_file=latest_pd_file,
            input_folder=INPUT_FOLDER,
            output_folder=OUTPUT_FOLDER,
            historic_filename="02. Historic PD Calculation 2024-25.xlsb",
            pivot_sheet="01.Pivoted_Portfolio",
            historic_sheet="02.Working"
        )
        
        if historic_output_file:
            print(f"\n✅ Historic PD file created: {os.path.basename(historic_output_file)}")
            return True
        else:
            print(f"\n❌ Failed to create Historic PD file")
            return False
            
    except Exception as e:
        print(f"\n❌ ERROR in Historic PD update: {e}")
        import traceback
        traceback.print_exc()
        return False


# =============================================================================
# ENTRY POINT
# =============================================================================
if __name__ == "__main__":
    try:
        print("\n" + "█"*80)
        print("█" + " "*78 + "█")
        print("█" + " "*20 + "PD AUTOMATION - COMPLETE WORKFLOW" + " "*25 + "█")
        print("█" + " "*78 + "█")
        print("█"*80)
        
        # ==============================================================
        # PART 1: Portfolio Roll-Forward Automation
        # ==============================================================
        print("\n▶ Starting Part 1: Portfolio Roll-Forward...")
        success, pd_output_file = run_automation()
        
        if not success or not pd_output_file:
            print("\n❌ Part 1 failed. Stopping automation.")
            sys.exit(1)
        
        print("\n✅ Part 1 completed successfully!")
        
        # ==============================================================
        # PART 2: Historic PD Update Automation
        # ==============================================================
        print("\n▶ Starting Part 2: Historic PD Update...")
        historic_success = run_historic_update(pd_output_file)
        
        if not historic_success:
            print("\n⚠️  Part 2 failed, but Part 1 was successful.")
            print("   PD file was created successfully.")
            sys.exit(1)
        
        print("\n✅ Part 2 completed successfully!")
        
        # ==============================================================
        # FINAL SUMMARY
        # ==============================================================
        print("\n" + "█"*80)
        print("█" + " "*78 + "█")
        print("█" + " "*15 + "🎉 ALL AUTOMATIONS COMPLETED SUCCESSFULLY! 🎉" + " "*18 + "█")
        print("█" + " "*78 + "█")
        print("█"*80)
        print("\n📁 Output files saved in:")
        print(f"   {OUTPUT_FOLDER}")
        print("\n✅ Portfolio Roll-Forward: Complete")
        print("✅ Historic PD Update: Complete")
        print("\n")
        
        sys.exit(0)
        
    except Exception as e:
        print("\n" + "="*80)
        print(f"❌ FATAL ERROR: {e}")
        print("="*80)
        import traceback
        traceback.print_exc()
        sys.exit(1)
        