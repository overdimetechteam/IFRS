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
import calendar
from datetime import datetime
import glob as glob_module

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
PD_FILE = os.path.join(INPUT_FOLDER, "01. PD_data_2024-25.xlsb")
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

# Columns to write from DataFrame to Portfolio sheets (only these 5, leave other columns alone)
DF_TO_PORTFOLIO_COLUMNS = ['MONTH', 'CONTRACT_NO', 'EQT_DESC', 'PD_CATEGORY', 'DPD']


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================
def add_months(date: datetime, months: int) -> datetime:
    """Add months to a date, handling month-end correctly."""
    month = date.month - 1 + months
    year = date.year + month // 12
    month = month % 12 + 1
    day = min(date.day, calendar.monthrange(year, month)[1])
    return datetime(year, month, day)


def months_between(start_date: datetime, end_date: datetime) -> int:
    """Calculate number of months between two dates."""
    return (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month)


def parse_month_string(month_str: str) -> datetime:
    """Parse MM/DD/YYYY string to datetime."""
    return datetime.strptime(month_str.strip(), '%m/%d/%Y')


def format_month_string(date: datetime) -> str:
    """Format datetime to MM/DD/YYYY string."""
    return date.strftime('%m/%d/%Y')


# =============================================================================
# CONFIG FILE OPERATIONS
# =============================================================================
def read_config() -> datetime:
    """Read the latest recorded month from config file."""
    if not os.path.exists(CONFIG_FILE):
        print(f"ERROR: Config file not found: {CONFIG_FILE}")
        print(f"Please create the file with the latest month in MM/DD/YYYY format.")
        sys.exit(1)

    with open(CONFIG_FILE, 'r') as f:
        date_str = f.read().strip()

    try:
        return parse_month_string(date_str)
    except ValueError:
        print(f"ERROR: Invalid date format in config file: {date_str}")
        print("Expected format: MM/DD/YYYY (e.g., 12/31/2024)")
        sys.exit(1)


def save_config(latest_month: datetime):
    """Save the latest month to config file."""
    with open(CONFIG_FILE, 'w') as f:
        f.write(format_month_string(latest_month))
    print(f"Config updated: {format_month_string(latest_month)}")


# =============================================================================
# USER INPUT
# =============================================================================
def get_end_month_from_args() -> datetime:
    """Get the new end month from command line argument."""
    if len(sys.argv) < 2:
        print("\nUsage: python Main.py MM/DD/YYYY")
        print("Example: python Main.py 09/30/2025")
        sys.exit(1)

    date_str = sys.argv[1].strip()
    try:
        return parse_month_string(date_str)
    except ValueError:
        print(f"\nERROR: Invalid date format: {date_str}")
        print("Expected format: MM/DD/YYYY (e.g., 09/30/2025)")
        sys.exit(1)


# =============================================================================
# SUMMARY FILE OPERATIONS
# =============================================================================
def find_summary_files(start_month: datetime, num_months: int) -> list:
    """
    Find summary files for the specified number of months after start_month.
    Returns list of (file_path, month_date) tuples in chronological order.
    """
    files = []

    for i in range(1, num_months + 1):
        target_date = add_months(start_month, i)
        # Summary files use YYYY-MM-DD format in filename
        date_pattern = target_date.strftime('%Y-%m-%d')

        pattern = os.path.join(INPUT_FOLDER, f"3. Summary_{date_pattern}*.xlsb")
        matches = glob_module.glob(pattern)

        if matches:
            files.append((matches[0], target_date))
            print(f"  Found: {os.path.basename(matches[0])}")
        else:
            print(f"  WARNING: No file for {date_pattern}")

    return files


def extract_summary_data(file_paths: list) -> pd.DataFrame:
    """
    Extract data from summary files into a DataFrame.

    Args:
        file_paths: List of (file_path, month_date) tuples

    Returns:
        DataFrame with columns matching PORTFOLIO_COLUMNS
    """
    all_data = []

    for file_path, month_date in file_paths:
        try:
            print(f"  Extracting: {os.path.basename(file_path)}")

            # Find correct header row
            df = None
            for header_row in range(10):
                try:
                    temp_df = pd.read_excel(file_path, sheet_name='SUMMARY',
                                           header=header_row, engine='pyxlsb')
                    if 'CONTRACT NO' in temp_df.columns:
                        df = temp_df
                        break
                except:
                    continue

            if df is None:
                df = pd.read_excel(file_path, sheet_name='SUMMARY',
                                  header=0, engine='pyxlsb')

            # Map columns
            extracted = pd.DataFrame()
            for src_col, tgt_col in COLUMN_MAPPING.items():
                found = None
                for col in df.columns:
                    if str(col).strip().upper() == src_col.upper():
                        found = col
                        break
                extracted[tgt_col] = df[found] if found else None

            # Add MONTH column (MM/DD/YYYY format)
            extracted['MONTH'] = format_month_string(month_date)

            # Filter empty rows
            extracted = extracted[extracted['CONTRACT_NO'].notna()]
            extracted = extracted[~extracted['CONTRACT_NO'].astype(str).isin(['', '-', 'nan', 'None'])]

            # Reorder columns (only the 5 we need)
            extracted = extracted[DF_TO_PORTFOLIO_COLUMNS]

            all_data.append(extracted)
            print(f"    Rows: {len(extracted)}")

        except Exception as e:
            print(f"    ERROR: {e}")

    if all_data:
        return pd.concat(all_data, ignore_index=True)
    return pd.DataFrame(columns=DF_TO_PORTFOLIO_COLUMNS)


# =============================================================================
# PORTFOLIO OPERATIONS
# =============================================================================
def read_portfolio(excel: ExcelPortfolioAutomation, sheet_name: str) -> pd.DataFrame:
    """Read portfolio data from sheet (reads all columns as-is)."""
    sheet = excel.workbook.sheets[sheet_name]
    last_row = sheet.range('A2').end('down').row

    # Handle empty sheet
    if last_row > 1000000:
        return pd.DataFrame(columns=DF_TO_PORTFOLIO_COLUMNS)

    # Read entire used range to preserve all columns
    df = excel.read_sheet_range_to_dataframe(sheet_name, None)
    return df


def get_unique_months(df: pd.DataFrame) -> list:
    """Get sorted list of unique months as datetime objects."""
    if df.empty or 'MONTH' not in df.columns:
        return []

    months = pd.to_datetime(df['MONTH'], format='%m/%d/%Y', errors='coerce')
    unique_months = sorted(months.dropna().unique())
    return [pd.Timestamp(m).to_pydatetime() for m in unique_months]


def filter_by_months(df: pd.DataFrame, months_to_keep: list) -> pd.DataFrame:
    """Filter DataFrame to keep only rows matching specified months."""
    if df.empty:
        return df

    df = df.copy()
    df['_MONTH_DT'] = pd.to_datetime(df['MONTH'], format='%m/%d/%Y', errors='coerce')

    # Convert months_to_keep to timestamps for comparison
    keep_timestamps = [pd.Timestamp(m) for m in months_to_keep]
    filtered = df[df['_MONTH_DT'].isin(keep_timestamps)]

    return filtered.drop(columns=['_MONTH_DT'])


def write_portfolio(excel: ExcelPortfolioAutomation, sheet_name: str, df: pd.DataFrame):
    """
    Write only the 5 mapped columns to Portfolio sheet.
    Column positions: A=MONTH, B=CONTRACT_NO, C=skip(formula), D=EQT_DESC, E=PD_CATEGORY, F=DPD
    Other columns (like DC_BUCKET) are left untouched - they are formulas.
    """
    sheet = excel.workbook.sheets[sheet_name]

    # Clear only the 5 data columns (skip C which is formula column CONTRACT_NO_NOLASTDI)
    # Also don't touch columns beyond F (like DC_BUCKET which is also a formula)
    last_row = sheet.range('A2').end('down').row
    if last_row < 1000000:  # Has data
        try:
            sheet.range(f'A2:B{last_row}').clear_contents()  # Clear A and B
            sheet.range(f'D2:F{last_row}').clear_contents()  # Clear D, E, F (skip C)
        except:
            pass

    if df.empty:
        print(f"  {sheet_name}: No data to write")
        return

    # Write only the 5 columns we need, to their specific positions
    # A=MONTH, B=CONTRACT_NO, D=EQT_DESC, E=PD_CATEGORY, F=DPD
    # Skip column C (CONTRACT_NO_NOLASTDI) - it's a formula
    num_rows = len(df)

    if 'MONTH' in df.columns:
        sheet.range('A2').value = df['MONTH'].values.reshape(-1, 1).tolist()
    if 'CONTRACT_NO' in df.columns:
        sheet.range('B2').value = df['CONTRACT_NO'].values.reshape(-1, 1).tolist()
    if 'EQT_DESC' in df.columns:
        sheet.range('D2').value = df['EQT_DESC'].values.reshape(-1, 1).tolist()
    if 'PD_CATEGORY' in df.columns:
        sheet.range('E2').value = df['PD_CATEGORY'].values.reshape(-1, 1).tolist()
    if 'DPD' in df.columns:
        sheet.range('F2').value = df['DPD'].values.reshape(-1, 1).tolist()

    print(f"  {sheet_name}: {num_rows} rows written (5 columns: MONTH, CONTRACT_NO, EQT_DESC, PD_CATEGORY, DPD)")


# =============================================================================
# MAIN AUTOMATION
# =============================================================================
def run_automation():
    """Main automation workflow."""

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    print("\n" + "="*80)
    print("PD PORTFOLIO ROLL-FORWARD AUTOMATION")
    print("="*80)

    # Step 1: Read current latest month from config
    current_latest = read_config()
    print(f"\nCurrent latest month: {format_month_string(current_latest)}")

    # Step 2: Get new end month from command line argument
    new_end_month = get_end_month_from_args()
    print(f"New end month: {format_month_string(new_end_month)}")

    # Step 3: Calculate months to process
    num_new_months = months_between(current_latest, new_end_month)

    if num_new_months <= 0:
        print("\nERROR: New end month must be after current latest month.")
        return False

    if num_new_months > PORTFOLIO_2_MONTHS:
        print(f"\nWARNING: {num_new_months} months requested, max is {PORTFOLIO_2_MONTHS}.")
        num_new_months = PORTFOLIO_2_MONTHS

    print(f"\nMonths to process: {num_new_months}")

    # Step 4: Find and extract summary files
    print("\n" + "-"*60)
    print(f"FINDING {num_new_months} SUMMARY FILES")
    print("-"*60)

    summary_files = find_summary_files(current_latest, num_new_months)

    if not summary_files:
        print("\nERROR: No summary files found.")
        return False

    print(f"\nExtracting data...")
    new_data = extract_summary_data(summary_files)

    if new_data.empty:
        print("\nERROR: No data extracted.")
        return False

    print(f"\nNew data: {len(new_data)} rows")

    # Step 5-7: Open Excel and perform roll-forward
    print("\n" + "-"*60)
    print("PERFORMING ROLL-FORWARD")
    print("-"*60)

    with ExcelPortfolioAutomation(PD_FILE, visible=True) as excel:

        # Read current portfolios
        print("\nReading current portfolios...")
        p1_data = read_portfolio(excel, 'Portfolio_1')
        p2_data = read_portfolio(excel, 'Portfolio_2')

        p1_months = get_unique_months(p1_data)
        p2_months = get_unique_months(p2_data)

        print(f"  Portfolio_1: {len(p1_months)} months, {len(p1_data)} rows")
        print(f"  Portfolio_2: {len(p2_months)} months, {len(p2_data)} rows")

        # Calculate new month distributions
        new_months = get_unique_months(new_data)
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

        # Split data by new month distributions
        new_p1_data = filter_by_months(all_data, new_p1_months)
        new_p2_data = filter_by_months(all_data, new_p2_months)

        # Sort by month
        new_p1_data = new_p1_data.sort_values('MONTH', key=lambda x: pd.to_datetime(x, format='%m/%d/%Y'))
        new_p2_data = new_p2_data.sort_values('MONTH', key=lambda x: pd.to_datetime(x, format='%m/%d/%Y'))

        # Write to portfolios
        print("\nWriting portfolios...")
        write_portfolio(excel, 'Portfolio_1', new_p1_data)
        write_portfolio(excel, 'Portfolio_2', new_p2_data)

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

    # Update config
    save_config(new_end_month)

    print("\n" + "="*80)
    print("AUTOMATION COMPLETED!")
    print("="*80)

    return True


# =============================================================================
# ENTRY POINT
# =============================================================================
if __name__ == "__main__":
    try:
        success = run_automation()
        sys.exit(0 if success else 1)
    except Exception as e:
        print(f"\nFATAL ERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
