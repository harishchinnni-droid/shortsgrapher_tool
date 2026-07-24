import os
import shutil

# [CHANGED -- Task 66, 21-Jul-26, office-laptop portability] BASE_DIR now
# resolves from the ALGO_BASE_DIR environment variable first; if that's
# not set, it's auto-detected as the PARENT of the folder this file lives
# in (e.g. this file at 'D:\06_Claude_Automation\05_Codes\file_mgmt.py'
# resolves BASE_DIR to 'D:\06_Claude_Automation'), instead of a hardcoded
# 'F:\05_Claude_Automation'. The hardcoded fallback is exactly what broke
# LIVE on Harish's office laptop: the code folder there is
# 'D:\06_Claude_Automation\05_Codes' (different drive letter AND different
# folder name), so the old fallback pointed at a path that doesn't exist
# on that machine, and os.makedirs() in token_mgmt.py crashed trying to
# create drive 'F:\' itself (WinError 3). Auto-detecting from __file__
# means this now works unmodified on ANY machine/drive/folder-name
# without needing to set ALGO_BASE_DIR by hand -- the env var is kept
# only as an explicit override for non-Windows setups (Colab etc.) where
# auto-detection from a local script path wouldn't make sense anyway.
BASE_DIR = os.environ.get(
    "ALGO_BASE_DIR",
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
)
SOURCE_FILE = os.path.join(BASE_DIR, "01_SourceFile.xlsx")

# [ADDED -- Phase B] The Reference/symbol template now lives natively in
# Google Sheets (Harish's 01_SourceFile spreadsheet -- the same file
# sheets_sync.py mirrors the daily output back into) instead of requiring
# a manually re-exported .xlsx copy kept in this folder by hand. That
# manual step is exactly what broke this run ("Source template missing"
# -- the .xlsx copy was never re-uploaded). SOURCE_FILE above is still
# tried FIRST for zero behavior change on any setup that keeps using a
# local/Drive .xlsx copy; the Google Sheet is only reached for when that
# local copy isn't there.
SOURCE_SHEET_ID = os.environ.get(
    "SOURCE_SHEET_ID", "1ObfXwDkTZTCkiJIVzYDP1NAl1n19TLqgucZp-mOGIdY"
)
SOURCE_SHEET_TAB = "Reference"


def _coerce_numeric_columns(df, pd):
    """gspread returns every cell as a string. Restore numeric dtype for
    any column that was genuinely all-numeric (ignoring blanks) so
    downstream code that expects e.g. 'Option Price Difference' as a
    float behaves the same as when this came from an .xlsx source."""
    for col in df.columns:
        converted = pd.to_numeric(df[col], errors='coerce')
        is_blank = df[col].astype(str).str.strip() == ''
        if converted[~is_blank].notna().all():
            df[col] = converted
    return df


def _build_reference_from_google_sheet(new_filename):
    """Pulls the 'Reference' tab straight from the Google Sheet template
    and writes it as the seed 'Reference' sheet of a brand-new local/
    Drive workbook -- the Sheets-native replacement for shutil.copy2()ing
    a manually-exported 01_SourceFile.xlsx. Reuses the same service
    account credentials sheets_sync.py already authenticates with (same
    JSON key, same one-time setup), so nothing new needs configuring."""
    import sheets_sync
    import gspread
    from google.oauth2.service_account import Credentials
    import pandas as pd

    if not os.path.exists(sheets_sync.SERVICE_ACCOUNT_FILE):
        raise FileNotFoundError(
            f"[CRITICAL] Neither a local source template ('{SOURCE_FILE}') nor "
            f"the Google service account key ('{sheets_sync.SERVICE_ACCOUNT_FILE}') "
            "were found -- cannot provision today's tracker from either source."
        )

    scopes = ["https://www.googleapis.com/auth/spreadsheets",
              "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file(sheets_sync.SERVICE_ACCOUNT_FILE, scopes=scopes)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(SOURCE_SHEET_ID)
    worksheet = spreadsheet.worksheet(SOURCE_SHEET_TAB)
    rows = worksheet.get_all_values()
    if not rows:
        raise RuntimeError(f"[CRITICAL] '{SOURCE_SHEET_TAB}' tab in the source Google Sheet is empty.")

    df = pd.DataFrame(rows[1:], columns=rows[0])
    df = _coerce_numeric_columns(df, pd)
    df.to_excel(new_filename, sheet_name='Reference', index=False)


def provision_daily_trade_file(target_date, mode=None):
    """[CHANGED -- Task 72, 22-Jul-26] mode ('LIVE'/'BACKTEST', i.e.
    calendar_mgmt.LIVE/calendar_mgmt.BACKTEST -- compared as plain
    strings here to avoid importing calendar_mgmt into this module)
    appends '-L' or '-BT' to the filename, per Harish's request so a
    LIVE run and a BACKTEST run for the SAME calendar date no longer
    collide/overwrite each other and can be compared side-by-side
    without either one having to be manually renamed first (exactly the
    manual step Harish was doing by hand before every LIVE-vs-BACKTEST
    audit so far). mode=None (default) keeps the old unsuffixed name for
    any caller that hasn't been updated to pass it."""
    date_str = target_date.strftime('%d-%b-%y')
    suffix = {"LIVE": "-L", "BACKTEST": "-BT"}.get(mode, "")
    new_filename = os.path.join(BASE_DIR, f"{date_str} FNO{suffix}.xlsx")

    if not os.path.exists(new_filename):
        if os.path.exists(SOURCE_FILE):
            print(f"[SYSTEM] Creating daily tracker: '{new_filename}' from local source.")
            shutil.copy2(SOURCE_FILE, new_filename)
        else:
            print(f"[SYSTEM] No local source template found -- pulling '{SOURCE_SHEET_TAB}' "
                  f"from the Google Sheet template instead.")
            _build_reference_from_google_sheet(new_filename)
            print(f"[SYSTEM] Daily tracker '{new_filename}' created from Google Sheet source.")
    else:
        print(f"[SYSTEM] Using existing daily tracker: '{new_filename}'.")

    return new_filename
