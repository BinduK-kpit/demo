import requests
import pandas as pd
import os
import sys
import json
import html
from datetime import datetime, timezone, timedelta
from pathlib import Path
from dotenv import load_dotenv
load_dotenv()

# ======================
# Configuration
# ======================

URL = "https://devstack.vwgroup.com/confluence/rest/confiforms/1.0/search/1093179152/dr"
TOKEN=os.getenv("CONFLUENCE_API_TOKEN")
SCRIPT_DIR = Path(__file__).parent
OUTPUT_DIR = SCRIPT_DIR / "output"
OUTPUT_EXCEL = OUTPUT_DIR / "WA_deployment_request.xlsx"
SCRAPED_JSON = SCRIPT_DIR / "structured_DR_table.json"

# ======================
# Time Fields
# ======================

TIME_FIELDS = {
    "created",
    "drdeploymentdate",
    "drdeploymentdatefinal"
}

# ======================
# Excel Columns
# ======================

COLUMNS = [
    "requestid",
    "drappname",
    "recordId",
    "createdBy",
    "created",
    "id",
    "drhmi",
    "drsmart",
    "drsales",
    "drstatus",
    "drfra1",
    "drchanges3",
    "drperformance1",
    "drchanges1",
    "drchanges2",
    "drmeta2",
    "drmeta1",
    "drlegal1",
    "drlegal2",
    "drdeploymentdatefinal",
    "drba",
    "drdeploymentdate",
    "drrnsytem",
    "dradmin1",
    "drmeta3",
    "dradmin2",
    "drdelta5",
    "drfosswarning",
    "drnewversion2",
    "drdelta2",
    "drnewversion1",
    "drpageurl",
    "drvwn1",
    "drvwn2",
    "drbrand1",
    "dract1",
    "dract2",
    "drpcm2",
    "drpcm",
    "drself",
    "droperation",
    "drmebrelease1",
    "warnssoamanifest",
    "drreleasedocu",
    "ownedBy"
]

# ======================
# Helper Functions
# ======================

def normalize_text(value):
    """
    Normalize text for matching:
    - HTML unescape
    - strip spaces
    - uppercase
    - remove all spaces
    """
    return "".join(html.unescape(str(value)).split()).upper()


def safe_str(value):
    if value is None:
        return ""
    return str(value).strip()


def ts_to_date_string(value):
    """
    Convert timestamp in milliseconds to a date string (YYYY-MM-DD).

    Logic:
    - Convert timestamp to UTC datetime
    - If hour is 22 or 23, move to next day
    - Return only the date part
    - If not a valid timestamp, return original value as string
    """
    if value in ("", None):
        return ""

    try:
        ts_ms = int(value)
        dt = datetime.fromtimestamp(ts_ms / 1000, tz=timezone.utc)

        # If hour is 22 or 23, move to the next day
        if dt.hour in (22, 23):
            dt = dt + timedelta(days=1)

        return dt.strftime("%Y-%m-%d")
    except Exception:
        return safe_str(value)

# ======================
# Load App Lookup
# ======================

def load_dr_app_lookup():
    """
    Reads structured_DR_table.json and builds:
        requestid -> app name

    Expected JSON structure:
    {
      "headers": [...],
      "rows": [
        ["Created", "Request ID", "App &amp; Version", ...],  # repeated header row
        ["2026-06-29", "WDR319", "carapp_maintenance 5.5.0", ...],
        ...
      ]
    }
    """
    lookup = {}

    try:
        if not SCRAPED_JSON.exists():
            raise FileNotFoundError(f"Scraped JSON file not found: {SCRAPED_JSON}")

        with open(SCRAPED_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)

        print("\nLoaded scraped JSON type:", type(data))

        if not isinstance(data, dict):
            raise ValueError(f"Unsupported JSON root type: {type(data)}")

        headers = data.get("headers", [])
        rows = data.get("rows", [])

        print("Headers found:")
        print(headers)
        print(f"Total scraped rows found: {len(rows)}")

        if not headers or not rows:
            raise ValueError("Missing 'headers' or 'rows' in scraped JSON.")

        # Build header -> index map
        header_index = {
            html.unescape(str(h)).strip(): idx
            for idx, h in enumerate(headers)
        }

        request_idx = header_index.get("Request ID")
        app_idx = header_index.get("App & Version")

        # Handle HTML-escaped header names too
        if request_idx is None:
            request_idx = header_index.get("Request ID".replace("&amp;", "&amp;amp;"))

        if app_idx is None:
            app_idx = header_index.get("App &amp;amp; Version")

        print(f"Request ID column index: {request_idx}")
        print(f"App & Version column index: {app_idx}")

        if request_idx is None or app_idx is None:
            raise ValueError(
                "Could not find required columns 'Request ID' and 'App & Version' in headers."
            )

        # Parse rows
        for row in rows:
            if not isinstance(row, list):
                continue

            # Skip repeated header row inside rows
            if len(row) >= 2 and str(row[0]).strip() == "Created" and str(row[1]).strip() == "Request ID":
                continue

            if len(row) <= max(request_idx, app_idx):
                continue

            request_id = normalize_text(row[request_idx])
            app_name = html.unescape(safe_str(row[app_idx]))

            if request_id and app_name:
                lookup[request_id] = app_name

        print(f"\nLoaded {len(lookup)} request-id -> app-name mappings")

        print("\nSample mappings:")
        for i, (k, v) in enumerate(lookup.items()):
            print(f"{k} -> {v}")
            if i >= 5:
                break

    except Exception as e:
        print(f"Lookup error: {e}")

    return lookup

# ======================
# Main
# ======================

def main():
    if not TOKEN:
        print("CONFLUENCE_API_TOKEN not set", file=sys.stderr)
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)

    # -------------------
    # Load DR mapping
    # -------------------
    dr_app_lookup = load_dr_app_lookup()

    # -------------------
    # Fetch API Data
    # -------------------
    print("\nFetching DR API...")

    api_headers = {
        "Authorization": f"Bearer {TOKEN}",
        "Accept": "application/json"
    }

    try:
        response = requests.get(URL, headers=api_headers, timeout=60)
        response.raise_for_status()
    except Exception as e:
        print(f"API Error: {e}", file=sys.stderr)
        sys.exit(1)

    data = response.json()

    entries = data.get("list", {}).get("entry", [])
    print(f"Fetched {len(entries)} records")

    rows = []

    # -------------------
    # Process Records
    # -------------------
    for entry in entries:
        row = {}
        fields = entry.get("fields", {})

        request_id_raw = (
            fields.get("requestid")
            or entry.get("requestid")
            or fields.get("Request ID")
            or entry.get("Request ID")
            or fields.get("requestId")
            or ""
        )

        request_id_norm = normalize_text(request_id_raw)

        for col in COLUMNS:
            if col == "drappname":
                value = dr_app_lookup.get(request_id_norm, "")

            elif col == "requestid":
                value = (
                    fields.get("requestid")
                    or entry.get("requestid")
                    or fields.get("Request ID")
                    or entry.get("Request ID")
                    or ""
                )

            elif col in entry:
                value = entry.get(col, "")

            elif col in fields:
                value = fields.get(col, "")

            else:
                value = ""

            # Apply date logic for time fields
            if col in TIME_FIELDS and value not in ("", None):
                value = ts_to_date_string(value)

            row[col] = value

        # Keep only rows where drappname is present
        if str(row.get("drappname", "")).strip():
            rows.append(row)

    # -------------------
    # Generate Excel
    # -------------------
    print("\nGenerating Excel...")

    df = pd.DataFrame(rows, columns=COLUMNS)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    df.to_excel(OUTPUT_EXCEL, index=False)

    print("\nExcel Created:")
    print(OUTPUT_EXCEL)
    print(f"Total Matching Rows: {len(df)}")

    if "drappname" in df.columns:
        filled_count = df["drappname"].fillna("").astype(str).str.strip().ne("").sum()
        print(f"App names populated: {filled_count} / {len(df)}")

        print("\nMatched rows:")
        print(df[["requestid", "drappname"]].to_string(index=False))

# ======================
# Run
# ======================

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected Error: {e}")
        import traceback
        traceback.print_exc()