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

def provision_daily_trade_file(target_date):
    date_str = target_date.strftime('%d-%b-%y')
    new_filename = os.path.join(BASE_DIR, f"{date_str} FNO.xlsx")

    if not os.path.exists(new_filename):
        if os.path.exists(SOURCE_FILE):
            print(f"[SYSTEM] Creating daily tracker: '{new_filename}' from source.")
            shutil.copy2(SOURCE_FILE, new_filename)
        else:
            raise FileNotFoundError(f"[CRITICAL] Source template missing at '{SOURCE_FILE}'.")
    else:
        print(f"[SYSTEM] Using existing daily tracker: '{new_filename}'.")
        
    return new_filename