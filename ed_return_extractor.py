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
#   3. Aggregates data by unique:
#        - Warehouse
#        - SKU ID
#        - Batch
#      Quantities and COGS are summed.
#   4. Selects/orders columns:
#        Store, SKU ID, Batch, Product Name, UOM TMP,
#        Qty (UOM TMP), Qty (UOM Inofarma), Expiry Date,
#        Return Policy, Return Status
#   5. Groups by Store and writes:
#        - One Google Sheet per store
#        - One combined Google Sheet (all stores)
#   6. Puts all output files inside one new Drive folder.
#   7. Builds dashboard-data.json and pushes it to GitHub.
#
# Run this in Google Colab.
# =============================================================================


# ---- 0. Setup ---------------------------------------------------------------

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

drive_service = build(
    "drive",
    "v3",
    credentials=creds
)


# ---- 1. Config ---------------------------------------------------------------

SOURCE_SHEET_ID = "1R955518MNXsmTF77v2IrYIvd6K-qWPoOH57TxvERzuY"

SOURCE_TAB = "ED - Store"

DATA_RANGE = "A20:L"


# Optional:
# Put outputs into an EXISTING Drive folder instead of creating
# a new dated folder.
#
# Leave as None to auto-create a dated folder.

DESTINATION_FOLDER_ID = None


# ---- GitHub auto-push --------------------------------------------------------

# Fine-grained PAT with:
# Contents: Read and write
# on this repository only.

GITHUB_TOKEN = ""


GITHUB_REPO = "tiffanynagari/ED-Separator"

GITHUB_FILE_PATH = "dashboard-data.json"

GITHUB_BRANCH = "main"


# ---- Target Return Status ----------------------------------------------------

TARGET_STATUSES = [
    "Expired (policy)",
    "Near ED — Return window closed",
    "Near ED — Return window open",
]


# ---- Raw columns -------------------------------------------------------------

RAW_COLUMNS = [
    "SKU_ID",
    "Batch",
    "Warehouse",
    "Qty_UOM_Inofarma",
    "Expiry_Date",
    "Custom_Return_Policy",
    "Return_Status",
    "COGS",
    "Product_Name",
    "UOM_TMP",
    "Qty_UOM_TMP",
    "Modulo",
]


# ---- Month abbreviation ------------------------------------------------------

MONTH_ABBR = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
]


# =============================================================================
# 2. Helper Functions
# =============================================================================


def today_str():
    """
    Return today's date in format:
    dd_Mmm_yyyy

    Example:
    21_Sep_2026
    """

    d = datetime.now()

    return f"{d.day:02d}_{MONTH_ABBR[d.month - 1]}_{d.year}"


# =============================================================================
# 3. Pull raw data
# =============================================================================


def fetch_raw_data():

    sh = gc.open_by_key(SOURCE_SHEET_ID)

    ws = sh.worksheet(SOURCE_TAB)

    values = ws.get(DATA_RANGE)


    # Pad/truncate every row to exactly 12 columns
    # so it lines up with RAW_COLUMNS.

    fixed_rows = []

    for row in values:

        if len(row) < 12:

            row = row + [""] * (12 - len(row))

        else:

            row = row[:12]

        fixed_rows.append(row)


    df = pd.DataFrame(
        fixed_rows,
        columns=RAW_COLUMNS
    )

    return df


# =============================================================================
# 4. Apply filter + aggregate + shape data
# =============================================================================


def filter_and_shape(df):

    df = df.copy()


    # -------------------------------------------------------------------------
    # 4.1 Convert Qty UOM TMP to numeric
    # -------------------------------------------------------------------------

    # Strip thousand separators just in case.

    df["Qty_UOM_TMP_num"] = pd.to_numeric(
        df["Qty_UOM_TMP"]
        .astype(str)
        .str.replace(",", "")
        .str.strip(),
        errors="coerce"
    )


    # -------------------------------------------------------------------------
    # 4.2 Return Status matching
    # -------------------------------------------------------------------------

    def status_match(status_value):

        s = str(status_value)

        return any(
            target in s
            for target in TARGET_STATUSES
        )


    # -------------------------------------------------------------------------
    # 4.3 Apply filter
    # -------------------------------------------------------------------------

    mask = (

        # SKU ID is not blank
        df["SKU_ID"].notna()

        & (
            df["SKU_ID"]
            .astype(str)
            .str.strip()
            != ""
        )

        # Qty UOM TMP is numeric
        & df["Qty_UOM_TMP_num"].notna()

        # Qty UOM TMP > 0
        & (
            df["Qty_UOM_TMP_num"] > 0
        )

        # Qty UOM TMP must be a whole number
        & (
            df["Qty_UOM_TMP_num"] % 1 == 0
        )

        # Return Status matches target statuses
        & df["Return_Status"].apply(status_match)
    )


    filtered = df[mask].copy()


    # -------------------------------------------------------------------------
    # 4.4 Convert COGS to numeric
    # -------------------------------------------------------------------------

    filtered["COGS_num"] = (
        filtered["COGS"]
        .astype(str)
        .apply(
            lambda s: pd.to_numeric(
                re.sub(r"[^0-9.\-]", "", s),
                errors="coerce"
            )
        )
        .fillna(0)
    )


    # -------------------------------------------------------------------------
    # 4.5 AGGREGATION
    # -------------------------------------------------------------------------
    #
    # IMPORTANT:
    #
    # One row = one unique combination of:
    #
    #     Warehouse
    #     SKU_ID
    #     Batch
    #
    # Quantitative columns are SUMMED.
    #
    # Descriptive columns use FIRST.
    #
    # -------------------------------------------------------------------------

    aggregated = (

        filtered

        .groupby(
            [
                "Warehouse",
                "SKU_ID",
                "Batch"
            ],

            as_index=False,

            dropna=False
        )

        .agg({

            # Descriptive information
            "Product_Name": "first",

            "UOM_TMP": "first",

            "Expiry_Date": "first",

            "Custom_Return_Policy": "first",

            "Return_Status": "first",


            # Quantities
            "Qty_UOM_TMP": "sum",

            "Qty_UOM_Inofarma": "sum",


            # COGS
            "COGS_num": "sum",
        })
    )


    # -------------------------------------------------------------------------
    # 4.6 Reorder + rename output columns
    # -------------------------------------------------------------------------

    ordered = aggregated[
        [
            "Warehouse",
            "SKU_ID",
            "Batch",
            "Product_Name",
            "UOM_TMP",
            "Qty_UOM_TMP",
            "Qty_UOM_Inofarma",
            "Expiry_Date",
            "Custom_Return_Policy",
            "Return_Status",
        ]
    ].rename(
        columns={

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
        }
    )


    # -------------------------------------------------------------------------
    # 4.7 Data for dashboard summary
    # -------------------------------------------------------------------------

    for_summary = aggregated[
        [
            "Warehouse",
            "SKU_ID",
            "COGS_num",
            "Return_Status",
        ]
    ].rename(
        columns={

            "Warehouse": "Store",

            "SKU_ID": "SKU ID",

            "COGS_num": "COGS",

            "Return_Status": "Return Status",
        }
    )


    return (
        ordered.reset_index(drop=True),

        for_summary.reset_index(drop=True)
    )


# =============================================================================
# 5. Build dashboard summary JSON
# =============================================================================


def build_dashboard_summary(
    for_summary,
    date_tag,
    folder_url
):


    # -------------------------------------------------------------------------
    # Summary by Store
    # -------------------------------------------------------------------------

    by_store = (

        for_summary

        .groupby("Store")

        .agg(

            uniqueSkus=(
                "SKU ID",
                "nunique"
            ),

            value=(
                "COGS",
                "sum"
            ),

            rows=(
                "SKU ID",
                "count"
            )
        )

        .reset_index()

        .sort_values(
            "value",
            ascending=False
        )
    )


    # -------------------------------------------------------------------------
    # Summary by Return Status
    # -------------------------------------------------------------------------

    status_mix = (

        for_summary

        .groupby("Return Status")["COGS"]

        .sum()

        .to_dict()
    )


    # -------------------------------------------------------------------------
    # Final JSON structure
    # -------------------------------------------------------------------------

    summary = {

        "generatedAt":
            datetime.now().isoformat(
                timespec="seconds"
            ),

        "dateTag":
            date_tag,

        "folderUrl":
            folder_url,


        "totals": {

            "grossValueAtRisk":
                round(
                    float(
                        for_summary["COGS"].sum()
                    ),
                    2
                ),

            "uniqueSkus":
                int(
                    for_summary["SKU ID"].nunique()
                ),

            "branches":
                int(
                    for_summary["Store"].nunique()
                ),

            "rows":
                int(
                    len(for_summary)
                ),
        },


        "statusMix":
            {
                k: round(
                    float(v),
                    2
                )

                for k, v in status_mix.items()
            },


        "byStore": [

            {

                "store":
                    row.Store,

                "uniqueSkus":
                    int(
                        row.uniqueSkus
                    ),

                "value":
                    round(
                        float(row.value),
                        2
                    ),

                "rows":
                    int(
                        row.rows
                    ),
            }

            for row
            in by_store.itertuples()
        ],
    }


    return summary


# =============================================================================
# 6. Google Drive / Google Sheets Functions
# =============================================================================


def get_or_create_folder(name):

    # If an existing folder ID is provided,
    # use that folder.

    if DESTINATION_FOLDER_ID:

        return DESTINATION_FOLDER_ID


    # Otherwise create a new folder.

    folder = drive_service.files().create(

        body={
            "name": name,
            "mimeType":
                "application/vnd.google-apps.folder"
        },

        fields="id",
    ).execute()


    return folder["id"]


# -----------------------------------------------------------------------------


def write_sheet(
    df,
    title,
    folder_id
):

    sheet = gc.create(
        title,
        folder_id=folder_id
    )

    ws = sheet.sheet1


    set_with_dataframe(
        ws,
        df
    )


    print(
        f"Saved: {title}  ->  {sheet.url}"
    )


# =============================================================================
# 7. Push dashboard-data.json to GitHub
# =============================================================================


def push_to_github(json_str):

    import requests
    import base64


    # -------------------------------------------------------------------------
    # Skip if token is empty
    # -------------------------------------------------------------------------

    if not GITHUB_TOKEN:

        print(
            "\nGITHUB_TOKEN is empty — "
            "skipping auto-push. "
            "Paste your token into the config section "
            "to enable this."
        )

        return False


    # -------------------------------------------------------------------------
    # GitHub API URL
    # -------------------------------------------------------------------------

    api_url = (
        f"https://api.github.com/repos/"
        f"{GITHUB_REPO}/contents/"
        f"{GITHUB_FILE_PATH}"
    )


    headers = {

        "Authorization":
            f"Bearer {GITHUB_TOKEN}",

        "Accept":
            "application/vnd.github+json",
    }


    # -------------------------------------------------------------------------
    # Get existing file SHA
    # -------------------------------------------------------------------------

    get_res = requests.get(

        api_url,

        headers=headers,

        params={
            "ref": GITHUB_BRANCH
        }
    )


    sha = (
        get_res.json().get("sha")
        if get_res.status_code == 200
        else None
    )


    # -------------------------------------------------------------------------
    # Prepare upload
    # -------------------------------------------------------------------------

    body = {

        "message":
            (
                "Update dashboard data — "
                f"{datetime.now().strftime('%Y-%m-%d %H:%M')}"
            ),

        "content":
            base64.b64encode(
                json_str.encode("utf-8")
            ).decode("utf-8"),

        "branch":
            GITHUB_BRANCH,
    }


    # GitHub requires SHA when updating
    # an existing file.

    if sha:

        body["sha"] = sha


    # -------------------------------------------------------------------------
    # Upload
    # -------------------------------------------------------------------------

    put_res = requests.put(

        api_url,

        headers=headers,

        json=body
    )


    if put_res.status_code in (200, 201):

        print(
            f"\nPushed dashboard-data.json "
            f"to {GITHUB_REPO} "
            f"({GITHUB_BRANCH}). "
            f"GitHub Pages will update in ~1 minute."
        )

        return True


    else:

        print(
            f"\nGitHub push failed "
            f"({put_res.status_code}): "
            f"{put_res.text[:300]}"
        )

        return False


# =============================================================================
# 8. Main
# =============================================================================


def main():

    # -------------------------------------------------------------------------
    # Date / folder
    # -------------------------------------------------------------------------

    date_tag = today_str()

    folder_name = (
        f"ED Return - {date_tag}"
    )


    folder_id = get_or_create_folder(
        folder_name
    )


    folder_url = (
        f"https://drive.google.com/drive/folders/"
        f"{folder_id}"
    )


    # -------------------------------------------------------------------------
    # Pull raw data
    # -------------------------------------------------------------------------

    print("\nFetching raw data...")

    raw = fetch_raw_data()


    print(
        f"Raw rows: {len(raw):,}"
    )


    # -------------------------------------------------------------------------
    # Filter + aggregate + shape
    # -------------------------------------------------------------------------

    print(
        "\nFiltering and aggregating..."
    )


    result, for_summary = filter_and_shape(
        raw
    )


    # -------------------------------------------------------------------------
    # Check empty result
    # -------------------------------------------------------------------------

    if result.empty:

        print(
            "No rows matched the filter — "
            "double check DATA_RANGE / SOURCE_TAB."
        )

        return


    # -------------------------------------------------------------------------
    # Show aggregation result
    # -------------------------------------------------------------------------

    print(
        f"Final rows after "
        f"Warehouse + SKU ID + Batch aggregation: "
        f"{len(result):,}"
    )


    # -------------------------------------------------------------------------
    # Combined file
    # -------------------------------------------------------------------------

    print(
        "\nCreating combined file..."
    )


    write_sheet(
        result,
        folder_name,
        folder_id
    )


    # -------------------------------------------------------------------------
    # One file per store
    # -------------------------------------------------------------------------

    print(
        "\nCreating store files..."
    )


    for store_name, group in result.groupby(
        "Store"
    ):

        title = (
            f"{folder_name} - {store_name}"
        )


        write_sheet(

            group.reset_index(drop=True),

            title,

            folder_id
        )


    print(
        f"\nDone. "
        f"{result['Store'].nunique()} "
        f"store files + 1 combined file "
        f"created in folder "
        f"'{folder_name}'."
    )


    # =============================================================================
    # Dashboard summary JSON
    # =============================================================================

    print(
        "\nBuilding dashboard summary..."
    )


    summary = build_dashboard_summary(

        for_summary,

        date_tag,

        folder_url
    )


    json_str = json.dumps(
        summary,
        indent=2
    )


    # -------------------------------------------------------------------------
    # Save JSON locally
    # -------------------------------------------------------------------------

    with open(
        "dashboard-data.json",
        "w"
    ) as f:

        f.write(json_str)


    # -------------------------------------------------------------------------
    # Push to GitHub
    # -------------------------------------------------------------------------

    pushed = push_to_github(
        json_str
    )


    # -------------------------------------------------------------------------
    # Fallback download
    # -------------------------------------------------------------------------

    if not pushed:

        try:

            files.download(
                "dashboard-data.json"
            )

        except Exception:

            pass


# =============================================================================
# Run
# =============================================================================

if __name__ == "__main__":

    main()
