import re
import sys
import logging
import subprocess
from pathlib import Path
from io import BytesIO
from datetime import datetime

import pandas as pd
from azure.storage.blob import BlobServiceClient, ContentSettings
from dotenv import load_dotenv
import os
# ==================
# Configuration
# ==================

SCRIPT_DIR = Path(__file__).parent
OUTPUT_DIR = SCRIPT_DIR / "output"
LOCAL_FILE = OUTPUT_DIR / "Deployment_Dashboard_WebApps.xlsx"
TIMESTAMP = datetime.now().strftime("%Y%m%d")

load_dotenv()

AZURE_STORAGE_CONNECTION_STRING = os.getenv("CONNECTION_STRING")

CONTAINER_NAME = os.getenv("CONTAINER_NAME")
BLOB_PREFIX = "downloads/manual/WA/"
OUTPUT_BLOB_NAME = f"{BLOB_PREFIX}{TIMESTAMP}_WA_Deployment_Dashboard.xlsx"

DASHBOARD_PATTERN = re.compile(r"^\d{8}_WA_Deployment_Dashboard\.xlsx$", re.IGNORECASE)

MAX_ROWS_PER_FILE = 150

PREFERRED_ORDER_COLUMNS = [
    "ACR App Name",
    "ACR ID",
    "DR ID",
    "ACR Link",
    "DR Link",
    "ACR Create Date",
    "DR Create Date",
    "Desired ACT Meeting Date",
    "Desired Deployment Date",
    "ACR Deployment Date",
    "Final Deployment Date",
    "Final DR Version",
    "App Type",
    "Release Type",
    "Brand",
    "Target Platform",
    "MEB Release",
    "MOD Generation",
    "ACR Status",
    "DR Status",
]

# ==================
# Logging
# ==================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)


# ==================
# Helper functions
# ==================

def list_blob_names(container_client, prefix: str):
    return [blob.name for blob in container_client.list_blobs(name_starts_with=prefix)]


def is_dashboard_file(filename: str) -> bool:
    return bool(DASHBOARD_PATTERN.match(Path(filename).name))


def extract_dashboard_date(blob_name: str):
    """
    Extract YYYYMMDD from blob filename like:
    20260720_WA_Deployment_Dashboard.xlsx
    """
    name = Path(blob_name).name
    m = re.match(r"^(\d{8})_WA_Deployment_Dashboard\.xlsx$", name, re.IGNORECASE)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), "%Y%m%d")
    except ValueError:
        return None


def read_excel_from_blob(container_client, blob_name: str) -> pd.DataFrame:
    log.info("Downloading blob: %s", blob_name)
    blob_client = container_client.get_blob_client(blob_name)
    blob_data = blob_client.download_blob().readall()
    df = pd.read_excel(BytesIO(blob_data))
    df.columns = df.columns.astype(str).str.strip()
    return df


def read_local_excel(file_path: Path) -> pd.DataFrame:
    if not file_path.exists():
        raise FileNotFoundError(f"Local file not found: {file_path}")
    df = pd.read_excel(file_path)
    df.columns = df.columns.astype(str).str.strip()
    return df


def normalize_text_series(series: pd.Series) -> pd.Series:
    return (
        series.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(r"\s+", " ", regex=True)
        .str.replace(r"[^\w\s\-]", "", regex=True)
    )


def prepare_app_key(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if "ACR App Name" in df.columns:
        df["APP_KEY"] = normalize_text_series(df["ACR App Name"])
    elif "App Name" in df.columns:
        df["APP_KEY"] = normalize_text_series(df["App Name"])
    else:
        raise KeyError("Input file must contain either 'ACR App Name' or 'App Name'.")

    df = df[df["APP_KEY"].notna() & (df["APP_KEY"] != "")].copy()
    return df


def to_datetime_safe(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series.astype(str).str.strip(), errors="coerce")


def get_latest_date(df: pd.DataFrame, column_name: str):
    if column_name not in df.columns:
        return pd.NaT
    return to_datetime_safe(df[column_name]).max()


def normalize_acr_id(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.lower()


def clean_output_columns(df: pd.DataFrame) -> pd.DataFrame:
    existing_cols = [c for c in PREFERRED_ORDER_COLUMNS if c in df.columns]
    return df[existing_cols].copy()


def run_merge_excel():
    merge_script = SCRIPT_DIR / "merge_excel.py"

    if not merge_script.exists():
        raise FileNotFoundError(f"merge_excel.py not found: {merge_script}")

    log.info("Starting merge_excel.py")

    result = subprocess.run(
        [sys.executable, str(merge_script)],
        capture_output=True,
        text=True,
    )

    if result.returncode != 0:
        log.error("merge_excel.py failed")
        if result.stdout:
            log.error(result.stdout)
        if result.stderr:
            log.error(result.stderr)
        raise RuntimeError("merge_excel.py execution failed")

    log.info("merge_excel.py completed successfully")


def delete_intermediate_file():
    try:
        if LOCAL_FILE.exists():
            LOCAL_FILE.unlink()
            log.info("Deleted intermediate file: %s", LOCAL_FILE)
    except Exception as exc:
        log.warning("Could not delete intermediate file %s. Error: %s", LOCAL_FILE, exc)


def save_output_to_blob(df: pd.DataFrame, container_client, blob_name: str) -> None:
    output_buffer = BytesIO()
    df.to_excel(output_buffer, index=False)
    output_buffer.seek(0)

    blob_client = container_client.get_blob_client(blob_name)
    blob_client.upload_blob(
        output_buffer.getvalue(),
        overwrite=True,
        content_settings=ContentSettings(
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
    log.info("Output uploaded to Azure blob: %s", blob_name)


def normalize_value_for_excel(value):
    """
    Convert values into Excel-safe values.
    """
    if pd.isna(value):
        return None

    if isinstance(value, (pd.Timestamp, datetime)):
        dt = pd.to_datetime(value, errors="coerce")
        if pd.isna(dt):
            return None
        return dt.strftime("%Y-%m-%d")

    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass

    return value


def build_output_blob_name(base_date: datetime) -> str:
    return f"{BLOB_PREFIX}{base_date.strftime('%Y%m%d')}_WA_Deployment_Dashboard.xlsx"


def get_output_files_sorted(container_client, prefix: str):
    blob_names = list_blob_names(container_client, prefix)
    dashboard_files = [b for b in blob_names if is_dashboard_file(b)]

    dashboard_files_with_dates = []
    for b in dashboard_files:
        dt = extract_dashboard_date(b)
        if dt is not None:
            dashboard_files_with_dates.append((b, dt))

    # Oldest -> newest
    dashboard_files_with_dates.sort(key=lambda x: x[1])
    return dashboard_files_with_dates


def upload_dataframe_to_blob(df: pd.DataFrame, container_client, blob_name: str) -> None:
    output_buffer = BytesIO()
    df.to_excel(output_buffer, index=False)
    output_buffer.seek(0)

    blob_client = container_client.get_blob_client(blob_name)
    blob_client.upload_blob(
        output_buffer.getvalue(),
        overwrite=True,
        content_settings=ContentSettings(
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
    log.info("Output uploaded to Azure blob: %s", blob_name)


def save_output_with_row_limit(df: pd.DataFrame, container_client) -> None:
    """
    Save output according to these rules:
    - No new rows -> do nothing
    - If latest file has room and total <= 150, append and rename/update to today's date
    - If appending would exceed 150, fill latest file up to 150 and create a new file for remaining rows
    - If latest file already has 150 rows, create a new file for incoming rows
    """
    if df.empty:
        return

    today = datetime.now()
    existing_files = get_output_files_sorted(container_client, BLOB_PREFIX)

    if not existing_files:
        # No previous file exists -> create today's file
        output_blob_name = build_output_blob_name(today)
        rows_to_write = df.iloc[:MAX_ROWS_PER_FILE].copy()
        upload_dataframe_to_blob(rows_to_write, container_client, output_blob_name)

        if len(df) > MAX_ROWS_PER_FILE:
            log.warning(
                "More than %d rows to save, but no multi-file suffix logic is enabled. "
                "Only the first %d rows were written.",
                MAX_ROWS_PER_FILE,
                MAX_ROWS_PER_FILE,
            )
        return

    latest_blob_name = existing_files[-1][0]
    current_df = read_excel_from_blob(container_client, latest_blob_name)
    current_df.columns = current_df.columns.astype(str).str.strip()
    current_count = len(current_df)

    # If current file already has 150 rows -> create today's new file
    if current_count >= MAX_ROWS_PER_FILE:
        output_blob_name = build_output_blob_name(today)
        rows_to_write = df.iloc[:MAX_ROWS_PER_FILE].copy()
        upload_dataframe_to_blob(rows_to_write, container_client, output_blob_name)

        if len(df) > MAX_ROWS_PER_FILE:
            log.warning(
                "More than %d rows to save, but no multi-file suffix logic is enabled. "
                "Only the first %d rows were written.",
                MAX_ROWS_PER_FILE,
                MAX_ROWS_PER_FILE,
            )
        return

    remaining_capacity = MAX_ROWS_PER_FILE - current_count

    # Case 1: all new rows fit in current file -> append and rename/update to today's date
    if len(df) <= remaining_capacity:
        combined_df = pd.concat([current_df, df], ignore_index=True)
        target_blob_name = build_output_blob_name(today)

        upload_dataframe_to_blob(combined_df, container_client, target_blob_name)

        # Delete old file if filename changed
        if latest_blob_name != target_blob_name:
            try:
                container_client.delete_blob(latest_blob_name)
                log.info("Deleted old blob after rename: %s", latest_blob_name)
            except Exception as exc:
                log.warning("Could not delete old blob %s. Error: %s", latest_blob_name, exc)

        return

    # Case 2: file will overflow -> fill old file up to 150, create new file for remaining rows
    rows_to_fill_current = df.iloc[:remaining_capacity].copy()
    remaining_rows = df.iloc[remaining_capacity:].copy()

    combined_current = pd.concat([current_df, rows_to_fill_current], ignore_index=True)

    # Keep the existing file name for the filled file
    upload_dataframe_to_blob(combined_current, container_client, latest_blob_name)

    # Create a new today's file for overflow rows
    if not remaining_rows.empty:
        new_blob_name = build_output_blob_name(today)
        upload_dataframe_to_blob(remaining_rows, container_client, new_blob_name)

    log.info(
        "Output split completed: %d row(s) added to existing file, %d row(s) written to new file.",
        len(rows_to_fill_current),
        len(remaining_rows),
    )


# ==================
# Azure historical file logic
# ==================

def get_dashboard_files_sorted(container_client, prefix: str):
    blob_names = list_blob_names(container_client, prefix)
    dashboard_files = [b for b in blob_names if is_dashboard_file(b)]

    dashboard_files_with_dates = []
    for b in dashboard_files:
        dt = extract_dashboard_date(b)
        if dt is not None:
            dashboard_files_with_dates.append((b, dt))

    # Oldest -> newest
    dashboard_files_with_dates.sort(key=lambda x: x[1])
    return dashboard_files_with_dates


def build_azure_acr_index(container_client, dashboard_files_with_dates) -> dict:
    """
    Build index of all historical ACR IDs.
    Latest blob wins if same ACR ID appears in multiple files.
    Structure:
    {
        normalized_acr_id: {
            "blob_name": ...,
            "row_idx": ...,
            "blob_date": ...
        }
    }
    """
    index = {}

    for blob_name, blob_dt in dashboard_files_with_dates:
        try:
            df = read_excel_from_blob(container_client, blob_name)
            df = prepare_app_key(df)

            if "ACR ID" not in df.columns:
                log.warning("Skipping index build for %s because 'ACR ID' is missing", blob_name)
                continue

            acr_ids = normalize_acr_id(df["ACR ID"])

            for row_idx, acr_id in acr_ids.items():
                if not acr_id or acr_id == "nan":
                    continue

                if acr_id not in index or blob_dt >= index[acr_id]["blob_date"]:
                    index[acr_id] = {
                        "blob_name": blob_name,
                        "row_idx": row_idx,
                        "blob_date": blob_dt,
                    }

        except Exception as exc:
            log.exception("Failed building ACR index from %s: %s", blob_name, exc)

    return index


def update_historical_blob_row(container_client, blob_name: str, row_idx: int, local_row: pd.Series) -> None:
    """
    Update the matching row in the historical Azure blob file.
    Only columns existing in the target blob are updated.
    Datetime values are converted safely before writing.
    """
    df = read_excel_from_blob(container_client, blob_name)

    if row_idx not in df.index:
        raise IndexError(f"Row index {row_idx} not found in blob {blob_name}")

    local_row_dict = local_row.to_dict()

    for col in df.columns:
        if col not in local_row_dict:
            continue

        value = normalize_value_for_excel(local_row_dict[col])

        if value is None:
            continue

        try:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df.at[row_idx, col] = pd.to_datetime(value, errors="coerce")
            else:
                df.at[row_idx, col] = value
        except Exception:
            df.at[row_idx, col] = str(value)

    output_buffer = BytesIO()
    df.to_excel(output_buffer, index=False)
    output_buffer.seek(0)

    blob_client = container_client.get_blob_client(blob_name)
    blob_client.upload_blob(
        output_buffer.getvalue(),
        overwrite=True,
        content_settings=ContentSettings(
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

    log.info("Updated historical blob %s at row %s", blob_name, row_idx)


def get_rows_matching_latest_dep_date(df: pd.DataFrame, latest_dep_date) -> pd.DataFrame:
    """
    Returns all rows in df where Desired Deployment Date equals latest_dep_date.
    """
    if "Desired Deployment Date" not in df.columns or pd.isna(latest_dep_date):
        return pd.DataFrame(columns=df.columns)

    date_series = to_datetime_safe(df["Desired Deployment Date"])
    return df[date_series == latest_dep_date].copy()


# ==================
# Main processing logic
# ==================

def build_comparison_output(container_client) -> pd.DataFrame:
    local_df = read_local_excel(LOCAL_FILE)
    local_df = prepare_app_key(local_df)

    if "ACR ID" not in local_df.columns:
        raise KeyError("Local file must contain 'ACR ID' column.")

    if "Desired ACT Meeting Date" in local_df.columns:
        local_df["Desired ACT Meeting Date"] = to_datetime_safe(local_df["Desired ACT Meeting Date"])
    if "Desired Deployment Date" in local_df.columns:
        local_df["Desired Deployment Date"] = to_datetime_safe(local_df["Desired Deployment Date"])

    local_act_latest = get_latest_date(local_df, "Desired ACT Meeting Date")
    local_dep_latest = get_latest_date(local_df, "Desired Deployment Date")

    log.info("Local file loaded: %s", LOCAL_FILE)
    log.info("Local rows: %d", len(local_df))
    log.info("Local latest Desired ACT Meeting Date: %s", local_act_latest)
    log.info("Local latest Desired Deployment Date: %s", local_dep_latest)

    dashboard_files_with_dates = get_dashboard_files_sorted(container_client, BLOB_PREFIX)

    if not dashboard_files_with_dates:
        raise FileNotFoundError(
            f"No dashboard files found under prefix {BLOB_PREFIX} matching pattern YYYYMMDD_WA_Deployment_Dashboard.xlsx"
        )

    log.info("Dashboard files found in Azure: %d", len(dashboard_files_with_dates))
    for blob_name, _ in dashboard_files_with_dates:
        log.info("  %s", blob_name)

    # Find latest Azure dates
    azure_act_latest = pd.NaT
    azure_dep_latest = pd.NaT

    for blob_name, _ in dashboard_files_with_dates:
        try:
            blob_df = read_excel_from_blob(container_client, blob_name)

            if "Desired ACT Meeting Date" in blob_df.columns:
                act_latest = get_latest_date(blob_df, "Desired ACT Meeting Date")
                if pd.isna(azure_act_latest) or (pd.notna(act_latest) and act_latest > azure_act_latest):
                    azure_act_latest = act_latest

            if "Desired Deployment Date" in blob_df.columns:
                dep_latest = get_latest_date(blob_df, "Desired Deployment Date")
                if pd.isna(azure_dep_latest) or (pd.notna(dep_latest) and dep_latest > azure_dep_latest):
                    azure_dep_latest = dep_latest

        except Exception as exc:
            log.exception("Failed reading latest dates from %s: %s", blob_name, exc)

    log.info("Overall Azure latest Desired ACT Meeting Date: %s", azure_act_latest)
    log.info("Overall Azure latest Desired Deployment Date: %s", azure_dep_latest)

    azure_acr_index = build_azure_acr_index(container_client, dashboard_files_with_dates)
    log.info("Built Azure ACR ID index with %d entries", len(azure_acr_index))

    # ------------------------------------------------------------
    # 1) Find brand-new ACR IDs -> write them to a new output file
    # 2) If local latest deployment date is newer than Azure latest:
    #    get ACR ID(s) from local latest-date row(s),
    #    find those ACR IDs in Azure history,
    #    update only the matching row(s)
    # ------------------------------------------------------------

    rows_for_new_output = []
    seen_acr_ids = set()

    # Build new output file for ACR IDs not present in Azure
    for _, row in local_df.iterrows():
        acr_id = str(row.get("ACR ID", "")).strip().lower()

        if not acr_id or acr_id == "nan":
            continue

        if acr_id in seen_acr_ids:
            continue
        seen_acr_ids.add(acr_id)

        safe_row = row.copy()
        for col in safe_row.index:
            safe_row[col] = normalize_value_for_excel(safe_row[col])

        if acr_id not in azure_acr_index:
            rows_for_new_output.append(safe_row)

    # Historical update only if local latest deployment date is newer than Azure latest
    if pd.notna(local_dep_latest) and (pd.isna(azure_dep_latest) or local_dep_latest > azure_dep_latest):
        log.info(
            "Local latest Desired Deployment Date (%s) is newer than Azure latest (%s). Checking matching ACR ID(s) for update.",
            local_dep_latest,
            azure_dep_latest,
        )

        latest_dep_rows = get_rows_matching_latest_dep_date(local_df, local_dep_latest)

        if latest_dep_rows.empty:
            log.warning("No rows found in local file with the latest Desired Deployment Date.")
        else:
            for _, row in latest_dep_rows.iterrows():
                acr_id = str(row.get("ACR ID", "")).strip().lower()

                if not acr_id or acr_id == "nan":
                    continue

                safe_row = row.copy()
                for col in safe_row.index:
                    safe_row[col] = normalize_value_for_excel(safe_row[col])

                if acr_id in azure_acr_index:
                    info = azure_acr_index[acr_id]
                    log.info(
                        "ACR ID %s found in historical file %s. Updating only that matching row.",
                        row.get("ACR ID"),
                        info["blob_name"],
                    )
                    update_historical_blob_row(
                        container_client=container_client,
                        blob_name=info["blob_name"],
                        row_idx=info["row_idx"],
                        local_row=safe_row,
                    )
                else:
                    log.info(
                        "ACR ID %s from latest local deployment-date row not found in Azure history.",
                        row.get("ACR ID"),
                    )
    else:
        log.info(
            "Local latest Desired Deployment Date is not newer than Azure latest. No historical row update will be done."
        )

    # Prepare output for new ACR IDs only
    if not rows_for_new_output:
        log.warning("No brand-new ACR IDs found. Historical files may have been updated only.")
        return pd.DataFrame(columns=PREFERRED_ORDER_COLUMNS)

    final_df = pd.DataFrame(rows_for_new_output)

    for col in ["Desired ACT Meeting Date", "Desired Deployment Date"]:
        if col in final_df.columns:
            final_df[col] = pd.to_datetime(final_df[col], errors="coerce").dt.strftime("%Y-%m-%d")

    final_df = clean_output_columns(final_df)
    return final_df


# ==================
# Main
# ==================

def main() -> None:
    try:
        run_merge_excel()

        blob_service_client = BlobServiceClient.from_connection_string(
            AZURE_STORAGE_CONNECTION_STRING
        )
        container_client = blob_service_client.get_container_client(CONTAINER_NAME)

        df = build_comparison_output(container_client)

        if not df.empty:
            save_output_with_row_limit(df, container_client)
            log.info("Done. %d new record(s) written and uploaded.", len(df))
        else:
            log.info("Done. No new records to upload.")

        delete_intermediate_file()

    except FileNotFoundError as exc:
        log.error("Input file not found: %s", exc)
        delete_intermediate_file()
        sys.exit(1)

    except Exception as exc:
        log.exception("Unexpected error: %s", exc)
        delete_intermediate_file()
        sys.exit(1)


if __name__ == "__main__":
    main()