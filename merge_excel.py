import re
import sys
import logging
from pathlib import Path

import pandas as pd

# ==================
# Configuration
# ==================

SCRIPT_DIR = Path(__file__).parent
INPUT_DIR = SCRIPT_DIR / "input"
OUTPUT_DIR = SCRIPT_DIR / "output"
OUTPUT_FILE = OUTPUT_DIR / "Deployment_Dashboard_WebApps.xlsx"

# Source file names
ACR_FILE = OUTPUT_DIR / "WA_app_classification_request.xlsx"
DR_FILE = OUTPUT_DIR / "WA_deployment_request.xlsx"

# Columns to retain from each source file
ACR_COLS = [
    "requestid",
    "created",
    "acrpageurl",
    "acractdate",
    "acrappname",
    "acrapptype",
    "acrdeploymentdate",
    "acrstatus",
    "acrmebrelease",
    "acrreleasetype",
    "acrplatform",
    "acrmod",
    "acrbrand",
]

DR_COLS = [
    "requestid",
    "created",
    "drpageurl",
    "drdeploymentdate",
    "drdeploymentdatefinal",
    "drnewversion2",
    "drappname",
    "drstatus"  
]

# Rename maps
ACR_RENAME = {
    "requestid": "ACR ID",
    "acrpageurl": "ACR Link",
    "acrapptype": "App Type",
    "acractdate": "Desired ACT Meeting Date",
    "acrdeploymentdate": "ACR Deployment Date",
    "acrappname": "ACR App Name",
    "acrreleasetype": "Release Type",
    "acrmod": "MOD Generation",
    "acrbrand": "Brand",
    "acrplatform": "Target Platform",
    "acrmebrelease": "MEB Release",
    "created": "ACR Create Date",
    "acrstatus": "ACR Status",
}

DR_RENAME = {
    "requestid": "DR ID",
    "drstatus": "DR Status",
    "drdeploymentdate": "Desired Deployment Date",
    "drdeploymentdatefinal": "Final Deployment Date",
    "created": "DR Create Date",
    "drnewversion2": "Final DR Version",
    "drappname": "DR App Name",
    "drpageurl": "DR Link"
}

# Status normalization
STATUS_MAPPING = {
    "success": "completed",
    "current": "requested",
    "moved": "draft",
    "error": "cancelled",
    "complete": "Accepted"
}

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

def load_excel(path: Path, required_cols: list[str]) -> pd.DataFrame:
    """Read Excel file and keep only columns that exist."""
    log.info("Loading %s", path.name)
    df = pd.read_excel(path)

    # Normalize column names just in case there are spaces
    df.columns = df.columns.astype(str).str.strip()

    present = [c for c in required_cols if c in df.columns]
    missing = set(required_cols) - set(present)

    if missing:
        log.warning("Columns not found in %s - skipped: %s", path.name, sorted(missing))

    return df[present].copy().reset_index(drop=True)


def clean_text(series: pd.Series) -> pd.Series:
    """Lowercase, strip, remove extra spaces."""
    return (
        series.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(r"\s+", " ", regex=True)
    )


def normalize_app_name(series: pd.Series) -> pd.Series:
    """Normalize app names so matching is more reliable."""
    return (
        series.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(r"\s+", " ", regex=True)
        .str.replace(r"[^\w\s\-]", "", regex=True)  # remove punctuation
    )


def to_date_str(series: pd.Series) -> pd.Series:
    """Convert date-like series to YYYY-MM-DD."""
    return pd.to_datetime(series, errors="coerce").dt.strftime("%Y-%m-%d")


def unix_ms_to_date_str(series: pd.Series) -> pd.Series:
    """Convert Unix millisecond timestamp to YYYY-MM-DD."""
    numeric = pd.to_numeric(series, errors="coerce")
    return pd.to_datetime(numeric, unit="ms", errors="coerce").dt.strftime("%Y-%m-%d")


def normalise_status(series: pd.Series, mapping: dict) -> pd.Series:
    """Lowercase, strip, then replace keywords with canonical statuses."""
    return (
        series.astype(str)
        .str.lower()
        .str.strip()
        .replace(mapping)
    )


def extract_labels(value):
    """Extract all 'label':'text' values from JSON-like string."""
    if pd.isna(value):
        return None
    labels = re.findall(r'"label"\s*:\s*"([^"]+)"', str(value))
    return ", ".join(labels) if labels else None


def extract_mod_values(value):

    if pd.isna(value):
        return None

    text = str(value)
    mods = set()

    # Extract all label values
    labels = re.findall(r'"label"\s*:\s*"([^"]+)"', text, flags=re.IGNORECASE)

    for label in labels:
        clean_label = label.strip().upper().replace("MOD", "").strip()

        if clean_label.startswith("3"):
            mods.add("MOD3")
        elif clean_label.startswith("4"):
            mods.add("MOD4")

    return ", ".join(sorted(mods)) if mods else None
# ==================
# Main pipeline
# ==================

def build_dashboard() -> pd.DataFrame:
    """Load, clean, and merge ACR and DR data on app name."""

    # Load files
    acr = load_excel(ACR_FILE, ACR_COLS)
    dr = load_excel(DR_FILE, DR_COLS)

    # Rename columns
    acr.rename(columns=ACR_RENAME, inplace=True)
    dr.rename(columns=DR_RENAME, inplace=True)

    # Check if required columns exist after renaming
    if "ACR App Name" not in acr.columns:
        raise KeyError("ACR file must contain 'acrappname' column.")
    if "DR App Name" not in dr.columns:
        raise KeyError("DR file must contain 'drappname' column.")

    # Clean app names for matching
    acr["APP_KEY"] = normalize_app_name(acr["ACR App Name"])
    dr["APP_KEY"] = normalize_app_name(dr["DR App Name"])

    # Optional: remove blank keys
    acr = acr[acr["APP_KEY"].notna() & (acr["APP_KEY"] != "")].copy()
    dr = dr[dr["APP_KEY"].notna() & (dr["APP_KEY"] != "")].copy()

    # Merge by app name
    df = pd.merge(
        acr,
        dr,
        on="APP_KEY",
        how="outer",
        suffixes=("_ACR", "_DR")
    )

    log.info("Merged shape: %d rows x %d columns", *df.shape)

    # Format ACR dates
    for col in ("ACR Create Date", "Desired ACT Meeting Date", "ACR Deployment Date"):
        if col in df.columns:
            df[col] = to_date_str(df[col])
            log.info("Formatted date column: %s", col)

    # Format DR dates
    for col in ("DR Create Date", "Desired Deployment Date", "Final Deployment Date"):
        if col in df.columns:
            # Try unix ms first; if that fails, fall back to normal date parsing
            converted = unix_ms_to_date_str(df[col])
            if converted.isna().all():
                converted = to_date_str(df[col])
            df[col] = converted
            log.info("Formatted date column: %s", col)

    # Normalize status columns
    for col in ("ACR Status", "DR Status"):
        if col in df.columns:
            df[col] = normalise_status(df[col], STATUS_MAPPING)
            log.info("Normalized status column: %s", col)

    # Extract structured labels
    for col in ("Release Type", "Brand", "MEB Release", "Target Platform"):
        if col in df.columns:
            df[col] = df[col].apply(extract_labels)
            log.info("Extracted labels from column: %s", col)

    # Extract MOD values
    if "MOD Generation" in df.columns:
        df["MOD Generation"] = df["MOD Generation"].apply(extract_mod_values)
        log.info("Extracted MOD values from: MOD Generation")

    # Create clean final columns for display
    # Prefer ACR App Name, but if missing use DR App Name
    df["App Name"] = df["ACR App Name"].combine_first(df["DR App Name"])

    # Reorder columns (optional)
    preferred_order = [
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
        "DR Status"
    ]

    existing_cols = [c for c in preferred_order if c in df.columns]
    #remaining_cols = [c for c in df.columns if c not in existing_cols]
    df = df[existing_cols]

    return df


def save_dashboard(df: pd.DataFrame) -> None:
    """Write final DataFrame to Excel."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Merged Excel file created successfully: {OUTPUT_FILE}")


# ==================
# Entry point
# ==================

def main() -> None:
    try:
        dashboard = build_dashboard()
        save_dashboard(dashboard)
        log.info("Done. %d records written.", len(dashboard))
    except FileNotFoundError as exc:
        log.error("Input file not found: %s", exc)
        sys.exit(1)
    except Exception as exc:
        log.exception("Unexpected error: %s", exc)
        sys.exit(1)


if __name__ == "__main__":
    main()