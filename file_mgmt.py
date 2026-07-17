import os
import shutil

# [CHANGED -- cloud/Colab portability] BASE_DIR now resolves from the
# ALGO_BASE_DIR environment variable first, falling back to the original
# hardcoded Windows path so nothing changes for existing desktop runs
# where that variable isn't set. Set ALGO_BASE_DIR in Colab (or any
# non-Windows environment) to your mounted Drive folder, e.g.
# '/content/drive/MyDrive/02_Claude_Trading'. Every other module that
# needs this pipeline's root folder imports BASE_DIR from HERE instead of
# hardcoding its own copy -- one place to point at a new environment,
# not six.
BASE_DIR = os.environ.get("ALGO_BASE_DIR", r"F:\05_Claude_Automation")
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


def provision_daily_trade_file(target_date):
    date_str = target_date.strftime('%d-%b-%y')
    new_filename = os.path.join(BASE_DIR, f"{date_str} FNO.xlsx")

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
