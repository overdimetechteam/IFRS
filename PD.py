"""
PD Portfolio Roll-Forward Automation - 2-Run Cycle
====================================================
RUN 1: Processes 7 months (running month + 6 prior months).
        Portfolio_1 = 7 new months | Portfolio_2 = EMPTY
        Big pivot sees Portfolio_1 only (7 months)

RUN 2: Processes next 6 months (completes the year).
        Portfolio_1 = UNTOUCHED (7 months) | Portfolio_2 = 6 new months
        Big pivot sees BOTH portfolios = 13 months combined

Then cycle repeats: Run 1 -> Run 2 -> Run 1 -> ...
"""
import os
import sys
import glob
import calendar
from datetime import datetime

# Add the Class directory to the path
class_path = os.path.join(os.path.dirname(__file__), 'Scripts', 'Class')
sys.path.insert(0, class_path)

from BasicExcelFunctionsClass import ExcelPortfolioAutomation
import pandas as pd


class InsufficientSummaryFilesError(Exception):
    """Raised when Run 2 does not find exactly 6 new summary files."""
    pass


# =============================================================================
# CONFIGURATION
# =============================================================================
INPUT_FOLDER  = r"C:\Users\Ashen Alwis\Desktop\Impairment Claculation\Input Files\PD"
OUTPUT_FOLDER = r"C:\Users\Ashen Alwis\Desktop\Impairment Claculation\OutPut\PD"
ORIGINAL_PD_FILE = os.path.join(INPUT_FOLDER, "01. PD_data_2024-25.xlsb")
CONFIG_FILE      = os.path.join(INPUT_FOLDER, "latest_month.txt")
RUN_CYCLE_FILE   = os.path.join(INPUT_FOLDER, "run_cycle.txt")
P1_CACHE_FILE             = os.path.join(INPUT_FOLDER, "portfolio1_run1_cache.pkl")
PD_CAT_OVERRIDE_FILE      = os.path.join(INPUT_FOLDER, "Test_PD_Cate.xlsx")

# Run 2 minimum: exactly this many new summary files required
RUN2_REQUIRED_FILES = 6

# Column mapping: Summary file column -> DataFrame column
COLUMN_MAPPING = {
    'CONTRACT NO':           'CONTRACT_NO',
    'EQUIPMENT DESCRIPTION': 'EQT_DESC',
    'PD/LGD CATEGORY':       'PD_CATEGORY',
    'CLIENT DPD':             'DPD',
    'AG':                     'STAGE',   # primary: Summary STAGE column header
    'STAGE':                  'STAGE',   # fallback: in case header is literally 'STAGE'
}

# Columns extracted from summary files
DF_TO_PORTFOLIO_COLUMNS = ['MONTH', 'CONTRACT_NO', 'EQT_DESC', 'PD_CATEGORY', 'DPD', 'STAGE']

# Columns written to Portfolio sheets (includes computed CONTRACT_NO_NOLASTDIG ->col C)
PORTFOLIO_WRITE_COLUMNS = ['MONTH', 'CONTRACT_NO', 'CONTRACT_NO_NOLASTDIG', 'EQT_DESC', 'PD_CATEGORY', 'DPD', 'STAGE']

# Column positions in Excel (C = CONTRACT_NO_NOLASTDIG written as value, no formula needed)
COLUMN_POSITIONS = {
    'MONTH':                 'A',
    'CONTRACT_NO':           'B',
    'CONTRACT_NO_NOLASTDIG': 'C',
    'EQT_DESC':              'D',
    'PD_CATEGORY':           'E',
    'DPD':                   'F',
    'STAGE':                 'H'
}


# =============================================================================
# RUN CYCLE HELPERS
# =============================================================================
def read_run_cycle() -> int:
    """Read current run cycle (1 or 2) from file. Defaults to 1."""
    try:
        with open(RUN_CYCLE_FILE, 'r') as f:
            val = int(f.read().strip())
            if val not in (1, 2):
                return 1
            return val
    except Exception:
        return 1


def save_run_cycle(run_num: int) -> None:
    """Save run cycle number to file."""
    with open(RUN_CYCLE_FILE, 'w') as f:
        f.write(str(run_num))


def next_run_cycle(current: int) -> int:
    """Return next run cycle number (2 wraps back to 1)."""
    return 1 if current == 2 else current + 1


def add_months_eom(dt: datetime, months: int) -> datetime:
    """Add N months to a date and return the last day of the resulting month."""
    m = dt.month - 1 + months
    year  = dt.year + m // 12
    month = m % 12 + 1
    last_day = calendar.monthrange(year, month)[1]
    return dt.replace(year=year, month=month, day=last_day)


# =============================================================================
# PD CATEGORY OVERRIDE
# =============================================================================
def apply_pd_category_overrides(df: pd.DataFrame) -> pd.DataFrame:
    """
    Apply PD category corrections from the override file.

    The override file (Test_PD_Cate.xlsx) contains two columns:
        CONTRACT_NO_NOLASTDIG  |  PD_CATEGORY
    For any row in df where CONTRACT_NO_NOLASTDIG appears in the override file
    AND the stored PD_CATEGORY differs, the value is replaced with the override.

    Requires CONTRACT_NO_NOLASTDIG to already exist in df.
    Safe to call even if the override file is missing (returns df unchanged).
    """
    if not os.path.exists(PD_CAT_OVERRIDE_FILE):
        print(f"  [Override] File not found — skipping: {os.path.basename(PD_CAT_OVERRIDE_FILE)}")
        return df

    try:
        override_df = pd.read_excel(PD_CAT_OVERRIDE_FILE, usecols=['CONTRACT_NO_NOLASTDIG', 'PD_CATEGORY'])
        override_df = override_df.dropna(subset=['CONTRACT_NO_NOLASTDIG', 'PD_CATEGORY'])
        override_df['CONTRACT_NO_NOLASTDIG'] = override_df['CONTRACT_NO_NOLASTDIG'].astype(str).str.strip()
        override_df['PD_CATEGORY']           = override_df['PD_CATEGORY'].astype(str).str.strip()

        # Build lookup dict: CONTRACT_NO_NOLASTDIG -> correct PD_CATEGORY
        override_map = dict(zip(override_df['CONTRACT_NO_NOLASTDIG'], override_df['PD_CATEGORY']))

        # Apply only where the contract exists in the override AND the value differs
        df = df.copy()
        df['CONTRACT_NO_NOLASTDIG'] = df['CONTRACT_NO_NOLASTDIG'].astype(str).str.strip()
        mapped = df['CONTRACT_NO_NOLASTDIG'].map(override_map)
        changed_mask = mapped.notna() & (mapped != df['PD_CATEGORY'].astype(str).str.strip())
        change_count = changed_mask.sum()
        df.loc[changed_mask, 'PD_CATEGORY'] = mapped[changed_mask]

        print(f"  [Override] Applied: {change_count} PD_CATEGORY value(s) corrected "
              f"from {os.path.basename(PD_CAT_OVERRIDE_FILE)}")
        return df

    except Exception as e:
        print(f"  [Override] WARNING — could not apply overrides: {e}")
        return df


# =============================================================================
# SEQUENTIAL MONTH FILE FINDER
# =============================================================================
def find_required_summary_files(input_folder: str, after_month: datetime,
                                num_months: int = 6,
                                file_prefix: str = "3. Summary_",
                                file_extension: str = ".xlsb") -> list:
    """
    Find exactly num_months consecutive summary files starting from the month
    immediately after after_month. Checks each month in sequence — stops and
    raises InsufficientSummaryFilesError the moment any month's file is missing.

    Returns list of (file_path, month_date) in chronological order.
    """
    results = []
    for i in range(1, num_months + 1):
        target   = add_months_eom(after_month, i)
        yr_month = target.strftime('%Y-%m')
        pattern  = os.path.join(input_folder, f"{file_prefix}{yr_month}-*{file_extension}")
        matches  = glob.glob(pattern)

        if not matches:
            raise InsufficientSummaryFilesError(
                f"Missing summary file for {target.strftime('%B %Y')} ({yr_month}). "
                f"Found {len(results)} of {num_months} required files. "
                f"Add the missing file and re-run."
            )

        fp = sorted(matches)[0]   # if somehow multiple matches, take first alphabetically
        print(f"  Found [{i}/{num_months}]: {os.path.basename(fp)}")
        results.append((fp, target))

    return results


# =============================================================================
# MAIN AUTOMATION - PART 1: PORTFOLIO ROLL-FORWARD
# =============================================================================
def run_automation():
    """Main automation workflow - 2-run cycle portfolio roll-forward."""

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    print("\n" + "="*80)
    print("PD PORTFOLIO ROLL-FORWARD AUTOMATION  (2-RUN CYCLE)")
    print("="*80)

    # ------------------------------------------------------------------
    # Read current state
    # ------------------------------------------------------------------
    try:
        current_latest = ExcelPortfolioAutomation.read_config_file(CONFIG_FILE)
        print(f"\nCurrent latest month : {ExcelPortfolioAutomation.format_month_string(current_latest)}")
    except FileNotFoundError:
        print(f"\nERROR: Config file not found: {CONFIG_FILE}")
        sys.exit(1)
    except ValueError as e:
        print(f"\nERROR: {e}")
        sys.exit(1)

    run_cycle = read_run_cycle()
    print(f"Run cycle            : {run_cycle} of 2")

    # ------------------------------------------------------------------
    # Dynamically scan for all new summary files after current_latest
    # ------------------------------------------------------------------
    print("\n" + "-"*60)
    print("SCANNING FOR NEW SUMMARY FILES")
    print("-"*60)

    summary_files = find_required_summary_files(INPUT_FOLDER, current_latest)

    new_end_month = summary_files[-1][1]   # derived from latest file found
    print(f"New end month        : {ExcelPortfolioAutomation.format_month_string(new_end_month)}")
    print(f"Files to process     : {len(summary_files)}")

    print(f"\nExtracting data from {len(summary_files)} file(s)...")
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

    # Sort new data by month
    new_data = new_data.sort_values(
        'MONTH', key=lambda x: pd.to_datetime(x, format='%m/%d/%Y')
    )
    # Compute CONTRACT_NO_NOLASTDIG (remove last character) — written as value to col C
    # This replaces the Excel formula so col C can be safely cleared on Portfolio clears
    new_data['CONTRACT_NO_NOLASTDIG'] = new_data['CONTRACT_NO'].apply(
        lambda x: str(x)[:-1] if pd.notna(x) and str(x) else None
    )

    # Apply PD category overrides from client-supplied correction file
    new_data = apply_pd_category_overrides(new_data)

    print(f"\nNew data extracted: {len(new_data)} rows")

    # ------------------------------------------------------------------
    # Get latest PD file
    # ------------------------------------------------------------------
    pd_file = ExcelPortfolioAutomation.get_latest_file_in_folder(
        folder_path=OUTPUT_FOLDER,
        file_pattern="01. PD_data_2024-25_Updated_*.xlsb",
        fallback_file=ORIGINAL_PD_FILE
    )

    output_file = None

    # ------------------------------------------------------------------
    # Open Excel and apply run-specific portfolio strategy
    # ------------------------------------------------------------------
    print("\n" + "-"*60)
    print(f"PERFORMING ROLL-FORWARD  [RUN {run_cycle}]")
    print("-"*60)

    with ExcelPortfolioAutomation(pd_file, visible=True) as excel:

        # ==============================================================
        # RUN 1 — Portfolio_1 = last old month + 6 new = 7 months
        #         Portfolio_2 = empty
        # ==============================================================
        if run_cycle == 1:
            print("\nRun 1 strategy:")
            print("  -Portfolio_1 : keep last old month + append 6 new months = 7 months")
            print("  -Portfolio_2 : CLEAR (stays empty)")
            print("  -Pivot       : reads Portfolio_1 only (7 months)\n")

            # Find and extract the LAST OLD month's summary file (= current_latest month)
            # current_latest is the last processed month — its summary file sits one step back
            keep_start = add_months_eom(current_latest, -1)
            print(f"  Finding last-month summary file ({ExcelPortfolioAutomation.format_month_string(current_latest)})...")
            keep_files = ExcelPortfolioAutomation.find_summary_files_by_date_range(
                input_folder=INPUT_FOLDER,
                start_month=keep_start,
                num_months=1,
                file_prefix="3. Summary_",
                file_extension=".xlsb",
                date_format_in_filename='%Y-%m-%d'
            )

            if keep_files:
                p1_keep = ExcelPortfolioAutomation.extract_data_from_summary_files(
                    file_paths=keep_files,
                    column_mapping=COLUMN_MAPPING,
                    output_columns=DF_TO_PORTFOLIO_COLUMNS,
                    sheet_name='SUMMARY',
                    date_format='%m/%d/%Y'
                )
                p1_keep['CONTRACT_NO_NOLASTDIG'] = p1_keep['CONTRACT_NO'].apply(
                    lambda x: str(x)[:-1] if pd.notna(x) and str(x) else None
                )
                p1_keep = apply_pd_category_overrides(p1_keep)
                print(f"  Last old month loaded: {len(p1_keep)} rows")
            else:
                p1_keep = pd.DataFrame()
                print("  WARNING: Last-month summary file not found — Portfolio_1 will have 6 months only")

            # Combine: last old month + 6 new months = 7 months
            p1_combined = pd.concat([p1_keep, new_data], ignore_index=True)
            p1_combined = p1_combined.sort_values(
                'MONTH', key=lambda x: pd.to_datetime(x, format='%m/%d/%Y')
            )
            unique_months = ExcelPortfolioAutomation.get_unique_months_from_dataframe(p1_combined)
            print(f"  Portfolio_1 will contain {len(unique_months)} months, {len(p1_combined)} rows")

            excel.write_portfolio_data(
                'Portfolio_1', p1_combined,
                PORTFOLIO_WRITE_COLUMNS, COLUMN_POSITIONS
            )
            excel.write_portfolio_data(
                'Portfolio_2', pd.DataFrame(),
                PORTFOLIO_WRITE_COLUMNS, COLUMN_POSITIONS,
                full_clear=True
            )

            # Write DC_BUCKET formula to Portfolio_1
            print("  Writing DC_BUCKET formula to Portfolio_1...")
            try:
                p1_sheet   = excel.workbook.sheets['Portfolio_1']
                num_rows   = len(p1_combined)
                last_row   = num_rows + 1
                dc_formula = '=IF(H2=3,5,IF(F2<=0,1,IF(F2<=30,2,IF(F2<=60,3,IF(F2<=90,4,IF(F2>90,5))))))'
                p1_sheet.range('G2').formula = dc_formula
                if num_rows > 1:
                    p1_sheet.range('G2').api.AutoFill(
                        Destination=p1_sheet.range(f'G2:G{last_row}').api, Type=0
                    )
                print(f"  DC_BUCKET formula written to Portfolio_1 G2:G{last_row}")
            except Exception as e:
                print(f"  WARNING: Could not write DC_BUCKET formula to Portfolio_1: {e}")

            excel.refresh_pivot_table('01.Pivoted_Portfolio', 'PivotTable1')
            print("  Pivot refreshed — showing 7 months from Portfolio_1")

        # ==============================================================
        # RUN 2 — Portfolio_1 = untouched | Portfolio_2 = 6 new months
        #         Pivot sees Portfolio_2 ONLY (P1 temp-cleared & restored)
        # ==============================================================
        elif run_cycle == 2:
            print("\nRun 2 strategy:")
            print("  -Portfolio_1 : UNTOUCHED (7 months)")
            print("  -Portfolio_2 : paste 6 new months")
            print("  -Pivot       : reads BOTH portfolios = 13 months combined\n")

            # Write 6 new months to Portfolio_2
            excel.write_portfolio_data(
                'Portfolio_2', new_data,
                PORTFOLIO_WRITE_COLUMNS, COLUMN_POSITIONS
            )

            # Write DC_BUCKET formula directly to Portfolio_2
            print("  Writing DC_BUCKET formula to Portfolio_2...")
            try:
                p2_sheet    = excel.workbook.sheets['Portfolio_2']
                num_p2_rows = len(new_data)
                last_p2_row = num_p2_rows + 1
                dc_formula  = '=IF(H2=3,5,IF(F2<=0,1,IF(F2<=30,2,IF(F2<=60,3,IF(F2<=90,4,IF(F2>90,5))))))'
                p2_sheet.range('G2').formula = dc_formula
                if num_p2_rows > 1:
                    p2_sheet.range('G2').api.AutoFill(
                        Destination=p2_sheet.range(f'G2:G{last_p2_row}').api, Type=0
                    )
                print(f"  DC_BUCKET formula written to Portfolio_2 G2:G{last_p2_row}")
            except Exception as e:
                print(f"  WARNING: Could not write DC_BUCKET formula to Portfolio_2: {e}")

            # Refresh pivot — both Portfolio_1 (7 months) and Portfolio_2 (6 months) visible
            excel.refresh_pivot_table('01.Pivoted_Portfolio', 'PivotTable1')
            print("  Pivot refreshed — showing 13 months from both portfolios")

        # Save output file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(
            OUTPUT_FOLDER, f"01. PD_data_2024-25_Updated_{timestamp}.xlsb"
        )
        excel.save_as(output_file)
        print(f"\nSaved: {output_file}")

    # Update configs
    ExcelPortfolioAutomation.save_config_file(CONFIG_FILE, new_end_month)
    next_cycle = next_run_cycle(run_cycle)
    save_run_cycle(next_cycle)

    print(f"\nlatest_month.txt updated -> {ExcelPortfolioAutomation.format_month_string(new_end_month)}")
    print(f"run_cycle.txt updated    -> {next_cycle} (next run will be Run {next_cycle})")

    print("\n" + "="*80)
    print(f"RUN {run_cycle} of 2 PORTFOLIO ROLL-FORWARD COMPLETED!")
    print("="*80)

    return True, output_file


# =============================================================================
# MAIN AUTOMATION - PART 2: HISTORIC PD UPDATE
# =============================================================================
def run_historic_update(latest_pd_file: str) -> bool:
    """Second automation — copy pivot data to Historic PD format.

    Run 1: pivot has 7 months (Portfolio_1 only).
    Run 2: pivot has 13 months (both portfolios combined) — no merging needed.
    """
    try:
        historic_output_file = ExcelPortfolioAutomation.copy_pivot_to_historic(
            latest_pd_file=latest_pd_file,
            input_folder=INPUT_FOLDER,
            output_folder=OUTPUT_FOLDER,
            historic_filename="02. Historic PD Calculation 2024-25.xlsb",
            pivot_sheet="01.Pivoted_Portfolio",
            historic_sheet="02.Working"
        )

        if historic_output_file:
            print(f"\n Historic PD file created: {os.path.basename(historic_output_file)}")
            return True
        else:
            print(f"\n Failed to create Historic PD file")
            return False

    except Exception as e:
        print(f"\n ERROR in Historic PD update: {e}")
        import traceback
        traceback.print_exc()
        return False


# =============================================================================
# ENTRY POINT
# =============================================================================
if __name__ == "__main__":
    try:
        print("\n" + "="*80)
        print("   PD AUTOMATION - COMPLETE WORKFLOW (2-RUN CYCLE)")
        print("="*80)

        # ==============================================================
        # PART 1: Portfolio Roll-Forward
        # ==============================================================
        print("\nStarting Part 1: Portfolio Roll-Forward...")
        try:
            success, pd_output_file = run_automation()
        except InsufficientSummaryFilesError as e:
            print("\n" + "="*80)
            print("RUN 2 HALTED — INSUFFICIENT SUMMARY FILES")
            print("="*80)
            print(f"\n{e}")
            print("\nNo output files were written. Add the missing files and re-run.")
            print("="*80)
            sys.exit(1)

        if not success or not pd_output_file:
            print("\nPart 1 failed. Stopping automation.")
            sys.exit(1)

        print("\nPart 1 completed successfully!")

        # ==============================================================
        # PART 2: Historic PD Update
        # ==============================================================
        print("\nStarting Part 2: Historic PD Update...")
        historic_success = run_historic_update(pd_output_file)

        if not historic_success:
            print("\nPart 2 failed, but Part 1 was successful.")
            sys.exit(1)

        print("\nPart 2 completed successfully!")

        # ==============================================================
        # FINAL SUMMARY
        # ==============================================================
        print("\n" + "="*80)
        print("   ALL AUTOMATIONS COMPLETED SUCCESSFULLY!")
        print("="*80)
        print("\nOutput files saved in:")
        print(f"   {OUTPUT_FOLDER}")
        print("\nPortfolio Roll-Forward: Complete")
        print("Historic PD Update: Complete\n")

        sys.exit(0)

    except Exception as e:
        print("\n" + "="*80)
        print(f"FATAL ERROR: {e}")
        print("="*80)
        import traceback
        traceback.print_exc()
        sys.exit(1)
