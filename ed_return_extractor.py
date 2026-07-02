# =============================================================================
# ED Return Extractor — Google Colab script
# -----------------------------------------------------------------------------
# What this does:
#   1. Reads the raw ED list from "ED - Store!A20:L900" in the source sheet.
#   2. Applies the same filter logic as your LET/QUERY/IMPORTRANGE formula:
#        - Col1 (SKU ID) is not null
#        - Col11 (Qty UOM TMP) > 0
#        - Col11 is a whole number  (equivalent to MOD(Col11,1)=0)
#        - Col7 (Return Status) contains one of the 3 target statuses
#   3. Selects/orders columns the same way your query does:
#        Store, SKU ID, Batch, Product Name, UOM TMP, Qty (UOM TMP),
#        Qty (UOM Inofarma), Expiry Date, Return Policy, Return Status
#   4. Groups by Store (Column C) and writes:
#        - One Google Sheet per store  -> "ED Return - dd_mmm_yyyy - <Store>"
#        - One combined Google Sheet (all stores) -> "ED Return - dd_mmm_yyyy"
#   5. Puts all output files inside one new Drive folder named
#      "ED Return - dd_mmm_yyyy".
#
# Run this in Google Colab. It will prompt you to log in / authorize once.
# =============================================================================

# ---- 0. Setup -------------------------------------------------------------
!pip install --quiet gspread gspread_dataframe

from google.colab import auth, files
auth.authenticate_user()

import gspread
import pandas as pd
import json
import re
from datetime import datetime
from google.auth import default
from googleapiclient.discovery import build
from gspread_dataframe import set_with_dataframe

creds, _ = default()
gc = gspread.authorize(creds)
drive_service = build("drive", "v3", credentials=creds)

# ---- 1. Config --------------------------------------------------------------
SOURCE_SHEET_ID = "1R955518MNXsmTF77v2IrYIvd6K-qWPoOH57TxvERzuY"
SOURCE_TAB = "ED - Store"
DATA_RANGE = "A20:L900"   # raw data range, no header row

# Optional: put outputs into an EXISTING Drive folder instead of creating a
# new one. Leave as None to auto-create a dated folder.
DESTINATION_FOLDER_ID = None

TARGET_STATUSES = [
    "Expired (policy)",
    "Near ED — Return window closed",
    "Near ED — Return window open",
]

RAW_COLUMNS = [
    "SKU_ID", "Batch", "Warehouse", "Qty_UOM_Inofarma", "Expiry_Date",
    "Custom_Return_Policy", "Return_Status", "COGS", "Product_Name",
    "UOM_TMP", "Qty_UOM_TMP", "Modulo",
]

MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def today_str():
    d = datetime.now()
    return f"{d.day:02d}_{MONTH_ABBR[d.month - 1]}_{d.year}"


# ---- 2. Pull raw data --------------------------------------------------------
def fetch_raw_data():
    sh = gc.open_by_key(SOURCE_SHEET_ID)
    ws = sh.worksheet(SOURCE_TAB)
    values = ws.get(DATA_RANGE)

    # Pad/truncate every row to exactly 12 columns so it lines up with RAW_COLUMNS
    fixed_rows = []
    for row in values:
        row = row + [""] * (12 - len(row)) if len(row) < 12 else row[:12]
        fixed_rows.append(row)

    df = pd.DataFrame(fixed_rows, columns=RAW_COLUMNS)
    return df


# ---- 3. Apply the same filter as the QUERY formula ---------------------------
def filter_and_shape(df):
    df = df.copy()

    # Numeric qty (strip thousand separators just in case)
    df["Qty_UOM_TMP_num"] = pd.to_numeric(
        df["Qty_UOM_TMP"].astype(str).str.replace(",", "").str.strip(),
        errors="coerce",
    )

    def status_match(status_value):
        s = str(status_value)
        return any(target in s for target in TARGET_STATUSES)

    mask = (
        df["SKU_ID"].notna()
        & (df["SKU_ID"].astype(str).str.strip() != "")
        & df["Qty_UOM_TMP_num"].notna()
        & (df["Qty_UOM_TMP_num"] > 0)
        & (df["Qty_UOM_TMP_num"] % 1 == 0)          # whole number, like MOD(...,1)=0
        & df["Return_Status"].apply(status_match)
    )

    filtered = df[mask].copy()
    filtered["COGS_num"] = filtered["COGS"].astype(str).apply(
        lambda s: pd.to_numeric(re.sub(r"[^0-9.\-]", "", s), errors="coerce")
    ).fillna(0)

    # Reorder/rename to match: Col3, Col1, Col2, Col9, Col10, Col11, Col4, Col5, Col6, Col7
    ordered = filtered[[
        "Warehouse", "SKU_ID", "Batch", "Product_Name", "UOM_TMP",
        "Qty_UOM_TMP", "Qty_UOM_Inofarma", "Expiry_Date",
        "Custom_Return_Policy", "Return_Status",
    ]].rename(columns={
        "Warehouse": "Store",
        "SKU_ID": "SKU ID",
        "Batch": "Batch",
        "Product_Name": "Product Name",
        "UOM_TMP": "UOM TMP",
        "Qty_UOM_TMP": "Qty (UOM TMP)",
        "Qty_UOM_Inofarma": "Qty (UOM Inofarma)",
        "Expiry_Date": "Expiry Date",
        "Custom_Return_Policy": "Return Policy",
        "Return_Status": "Return Status",
    })

    # keep Store/SKU/COGS/Status alongside for the dashboard summary (COGS isn't
    # part of your query's selected columns, so it's dropped from `ordered`)
    for_summary = filtered[["Warehouse", "SKU_ID", "COGS_num", "Return_Status"]].rename(
        columns={"Warehouse": "Store", "SKU_ID": "SKU ID", "COGS_num": "COGS", "Return_Status": "Return Status"}
    )

    return ordered.reset_index(drop=True), for_summary.reset_index(drop=True)


# ---- 3b. Build the small JSON summary the dashboard reads -------------------
def build_dashboard_summary(for_summary, date_tag, folder_url):
    by_store = (
        for_summary.groupby("Store")
        .agg(uniqueSkus=("SKU ID", "nunique"), value=("COGS", "sum"), rows=("SKU ID", "count"))
        .reset_index()
        .sort_values("value", ascending=False)
    )
    status_mix = for_summary.groupby("Return Status")["COGS"].sum().to_dict()

    summary = {
        "generatedAt": datetime.now().isoformat(timespec="seconds"),
        "dateTag": date_tag,
        "folderUrl": folder_url,
        "totals": {
            "grossValueAtRisk": round(float(for_summary["COGS"].sum()), 2),
            "uniqueSkus": int(for_summary["SKU ID"].nunique()),
            "branches": int(for_summary["Store"].nunique()),
            "rows": int(len(for_summary)),
        },
        "statusMix": {k: round(float(v), 2) for k, v in status_mix.items()},
        "byStore": [
            {
                "store": row.Store,
                "uniqueSkus": int(row.uniqueSkus),
                "value": round(float(row.value), 2),
                "rows": int(row.rows),
            }
            for row in by_store.itertuples()
        ],
    }
    return summary


# ---- 4. Write to Google Sheets in Drive --------------------------------------
def get_or_create_folder(name):
    if DESTINATION_FOLDER_ID:
        return DESTINATION_FOLDER_ID
    folder = drive_service.files().create(
        body={"name": name, "mimeType": "application/vnd.google-apps.folder"},
        fields="id",
    ).execute()
    return folder["id"]


def write_sheet(df, title, folder_id):
    sheet = gc.create(title, folder_id=folder_id)
    ws = sheet.sheet1
    set_with_dataframe(ws, df)
    print(f"Saved: {title}  ->  {sheet.url}")


# ---- 5. Main ------------------------------------------------------------------
def main():
    date_tag = today_str()
    folder_name = f"ED Return - {date_tag}"
    folder_id = get_or_create_folder(folder_name)
    folder_url = f"https://drive.google.com/drive/folders/{folder_id}"

    raw = fetch_raw_data()
    result, for_summary = filter_and_shape(raw)

    if result.empty:
        print("No rows matched the filter — double check DATA_RANGE / SOURCE_TAB.")
        return

    # Combined file (all stores together)
    write_sheet(result, folder_name, folder_id)

    # One file per store
    for store_name, group in result.groupby("Store"):
        title = f"{folder_name} - {store_name}"
        write_sheet(group.reset_index(drop=True), title, folder_id)

    print(f"\nDone. {result['Store'].nunique()} store files + 1 combined file created in folder '{folder_name}'.")

    # ---- Dashboard summary JSON ----
    # This is the ONLY file the GitHub-hosted dashboard needs. Copy it into
    # your repo (replacing the old one) each time you run this, e.g.:
    #   cp ~/Downloads/dashboard-data.json /path/to/ED-Separator/dashboard-data.json
    #   git add dashboard-data.json && git commit -m "Update data" && git push
    summary = build_dashboard_summary(for_summary, date_tag, folder_url)
    json_path = "dashboard-data.json"
    with open(json_path, "w") as f:
        json.dump(summary, f, indent=2)

    print(f"\ndashboard-data.json written — download it and copy it into your GitHub repo.")
    try:
        files.download(json_path)
    except Exception:
        pass  # download() only works when run interactively in Colab's browser session


if __name__ == "__main__":
    main()
