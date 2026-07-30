import requests
import pandas as pd
import os
import sys
import json
from datetime import datetime, timezone
from pathlib import Path
from datetime import datetime, timezone, timedelta
from dotenv import load_dotenv
load_dotenv()
# =====================================================
# CONFIGURATION
# =====================================================

URL = "https://devstack.vwgroup.com/confluence/rest/confiforms/1.0/search/1035867849/acr"


TOKEN=os.getenv("CONFLUENCE_API_TOKEN")
SCRIPT_DIR = Path(__file__).parent

OUTPUT_DIR = SCRIPT_DIR / "output"

OUTPUT_EXCEL = (
    OUTPUT_DIR /
    "WA_app_classification_request.xlsx"
)

SCRAPED_JSON = (
    SCRIPT_DIR /
    "structured_ACR_table.json"
)

# =====================================================
# TIMESTAMP FIELDS
# =====================================================

TIME_FIELDS = {
    "created",
    "acractdate",
    "acrdeploymentdate",
    "acractdate2"
}

# =====================================================
# EXCEL COLUMN ORDER
# =====================================================
COLUMNS = [
    "recordId",
    "createdBy",
    "created",
    "id",
    "acractdate2",
    "acrapptype",
    "acrbrand",
    "acrpcm2",
    "acradmin2",
    "apslaw1",
    "acradmin1",
    "acrvwn",
    "apslaw4",
    "apslaw2",
    "apslaw3",
    "apslaw",
    "apsdown",
    "acrsales",
    "acrmebrelease",
    "apsevaluation",
    "acrreleasetype",
    "acrmotivation",
    "acrapsdecision2",
    "apsevaluation6",
    "acrapsdecision3",
    "apsevaluation4",
    "apsevaluation5",
    "acrdeploymentdate",
    "apsevaluation1",
    "acrfosswarning",
    "requestid",
    "acrlegal2",
    "acrlegal1",
    "apsorl10",
    "acrvwn1",
    "acrvwn2",
    "acrplatform",
    "acrindexlevel6",
    "acrba",
    "acrindexlevel9",
    "acrindexlevel2",
    "acrversion",
    "acrindexlevel3",
    "acrindexlevel4",
    "acrindexlevel5",
    "apssafety1",
    "apsdown2",
    "apsdown3",
    "acrstatus",
    "apssafety",
    "acrappname",
    "apssecurity1",
    "apsdown1",
    "apssecurity2",
    "acradmin",
    "apssecurity3",
    "apssafety2",
    "acroperation",
    "acrhmi",
    "acrpageurl",
    "acrmod",
    "apsimage1",
    "apsimage2",
    "apsimage3",
    "acrpcm",
    "apsquality",
    "apsorl4",
    "apsorl3",
    "apsorl2",
    "apsorl1",
    "acrperformance1",
    "apsorl6",
    "acrrnsytem",
    "apsorl5",
    "acrself",
    "apsimage",
    "apsquality1",
    "apsquality2",
    "acrcrboard",
    "warnssoamanifest",
    "apssecurity",
    "acractdate",
    "acrapsdecision"
]
  

# =====================================================
# CONVERT TIMESTAMP
# =====================================================

def ts_to_utc_string(ts_ms):
    try:
        ts_ms = int(ts_ms)

        dt = datetime.fromtimestamp(
            ts_ms / 1000,
            tz=timezone.utc
        )

        # If hour is 22 or 23, move to next day
        if dt.hour in (22, 23):
            dt = dt + timedelta(days=1)

        return dt.strftime("%Y-%m-%d")

    except Exception:
        return ""
# =====================================================
# LOAD APP NAME LOOKUP
# Request ID -> App & Version
# =====================================================

def load_app_lookup():

    lookup = {}

    if not SCRAPED_JSON.exists():

        print(
            f"WARNING: {SCRAPED_JSON} not found"
        )
        return lookup

    try:

        with open(
            SCRAPED_JSON,
            "r",
            encoding="utf-8"
        ) as f:

            data = json.load(f)

        headers = data.get(
            "headers",
            []
        )

        rows = data.get(
            "rows",
            []
        )

        if "Request ID" not in headers:

            raise ValueError(
                "Request ID column not found"
            )

        if "App & Version" not in headers:

            raise ValueError(
                "App & Version column not found"
            )

        request_idx = headers.index(
            "Request ID"
        )

        app_idx = headers.index(
            "App & Version"
        )

        for row in rows:

            if len(row) <= max(
                request_idx,
                app_idx
            ):
                continue

            request_id = str(
                row[request_idx]
            ).strip()

            app_name = str(
                row[app_idx]
            ).strip()

            if request_id:

                lookup[
                    request_id
                ] = app_name

        print(
            f"Loaded {len(lookup)} App mappings"
        )

    except Exception as e:

        print(
            f"Error reading scraped JSON: {e}"
        )

    return lookup

# =====================================================
# MAIN
# =====================================================

def main():

    if not TOKEN:

        print(
            "CONFLUENCE_API_TOKEN is not set",
            file=sys.stderr
        )

        sys.exit(1)

    OUTPUT_DIR.mkdir(
        exist_ok=True
    )

    # -------------------------------------------
    # LOAD SCRAPED REQUEST ID -> APP NAME
    # -------------------------------------------

    app_lookup = load_app_lookup()

    # -------------------------------------------
    # FETCH API DATA
    # -------------------------------------------

    print(
        "Fetching data from Confluence..."
    )

    api_headers = {
        "Authorization": f"Bearer {TOKEN}",
        "Accept": "application/json"
    }

    try:

        response = requests.get(
            URL,
            headers=api_headers,
            timeout=60
        )

        response.raise_for_status()

    except Exception as e:

        print(
            f"API Error: {e}",
            file=sys.stderr
        )

        sys.exit(1)

    data = response.json()

    entries = (
        data.get("list", {})
            .get("entry", [])
    )

    print(
        f"Fetched {len(entries)} records"
    )

    # -------------------------------------------
    # PROCESS DATA
    # -------------------------------------------

    rows = []

    for entry in entries:

        row = {}

        fields = entry.get(
            "fields",
            {}
        )

        request_id = str(
            fields.get(
                "requestid",
                ""
            )
        ).strip()

        for col in COLUMNS:

            # -----------------------
            # SPECIAL CASE:
            # APP NAME FROM SCRAPED DATA
            # -----------------------

            if col == "acrappname":

                scraped_app_name = (
                    app_lookup.get(
                        request_id,
                        ""
                    )
                )

                if scraped_app_name:

                    value = scraped_app_name

                else:

                    value = fields.get(
                        "acrappname",
                        ""
                    )

            # -----------------------
            # NORMAL FIELD PROCESSING
            # -----------------------

            elif col in entry:

                value = entry.get(col)

            elif col in fields:

                value = fields.get(col)

            else:

                value = ""

            # -----------------------
            # TIMESTAMP CONVERSION
            # -----------------------

            if (
                col in TIME_FIELDS
                and value not in ("", None)
            ):

                value = ts_to_utc_string(
                    value
                )

            row[col] = value

        rows.append(row)

    # -------------------------------------------
    # GENERATE EXCEL
    # -------------------------------------------

    print(
        "Generating Excel..."
    )

    df = pd.DataFrame(
        rows,
        columns=COLUMNS
    )

    df.to_excel(
        OUTPUT_EXCEL,
        index=False
    )

    print(
        f"Excel created successfully"
    )

    print(
        f"File: {OUTPUT_EXCEL}"
    )

    print(
        f"Total rows: {len(df)}"
    )

    return 0

# =====================================================
# RUN
# =====================================================

if __name__ == "__main__":

    try:

        sys.exit(main())

    except KeyboardInterrupt:

        print(
            "\nOperation cancelled",
            file=sys.stderr
        )

        sys.exit(1)

    except Exception as e:

        print(
            f"Unexpected Error: {e}",
            file=sys.stderr
        )

        import traceback
        traceback.print_exc()

        sys.exit(1)