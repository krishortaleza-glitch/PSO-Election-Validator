import re
from io import BytesIO

import pandas as pd
import streamlit as st


st.set_page_config(page_title="PSO Election Validator", page_icon="✅", layout="wide")


# -----------------------------
# Election normalization
# -----------------------------
def extract_99b_election(zone):
    """
    Extract the 99B election from Store List column B (zone).
    The zone can have many naming formats, so we look for known election
    indicators rather than assuming a fixed underscore position.
    """
    if pd.isna(zone):
        return ""

    z = str(zone).strip().upper()

    # Explicit / special cases first
    if "NO TOBACCO" in z or "NO_DISCOUNT" in z:
        return "OPT-OUT"
    if "CVF" in z:
        return "CVF"

    # Enhanced variants
    if "ENHANCED" in z or "ENHAN" in z:
        return "Enhanced Margin"

    # Standard elections
    if "RESERVE" in z:
        return "Reserve"
    if "MARGIN" in z:
        return "Margin"
    if "BALANCED" in z or re.search(r"\bBAL[A-Z]*", z):
        return "Balanced"
    if "BLENDED" in z:
        return "Blended"

    return ""


def normalize_election(value):
    """Normalize both portal and report election values to a common label."""
    if pd.isna(value):
        return ""

    v = str(value).strip().upper()

    if not v or v == "NAN":
        return ""

    if "OPT" in v or "NO TOBACCO" in v or "NO DISCOUNT" in v:
        return "OPT-OUT"
    if "CVF" in v:
        return "CVF"
    if "ENHANCED" in v:
        return "Enhanced Margin"
    if "RESERVE" in v:
        return "Reserve"
    if "MARGIN" in v:
        return "Margin"
    if "BALANCED" in v:
        return "Balanced"
    if "BLENDED" in v:
        return "Blended"

    return str(value).strip()


def clean_store(value):
    """Keep digits only. Used for the matching key."""
    if pd.isna(value):
        return ""
    digits = re.sub(r"\D", "", str(value))
    return digits


# -----------------------------
# Read Store List
# -----------------------------
def read_store_list(uploaded_file):
    """
    Reads every tab in Storelist.xlsx.

    User requirement:
      - Column B = zone
      - Extract the 99B election from Column B
      - Store is the first column
      - State column is used for output state
    """
    xls = pd.ExcelFile(uploaded_file)
    records = []

    for sheet in xls.sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=sheet)

        if df.shape[1] < 2:
            continue

        store_col = df.columns[0]
        zone_col = df.columns[1]

        state_col = next(
            (c for c in df.columns if str(c).strip().lower() == "state"),
            None,
        )

        for _, row in df.iterrows():
            store = clean_store(row.get(store_col))
            if not store:
                continue

            zone = row.get(zone_col)
            election = extract_99b_election(zone)

            state = ""
            if state_col is not None and pd.notna(row.get(state_col)):
                state = str(row.get(state_col)).strip().upper()

            records.append(
                {
                    "Store": store,
                    "State": state,
                    "99B Election": election,
                    "Zone": "" if pd.isna(zone) else str(zone).strip(),
                    "Source Tab": sheet,
                }
            )

    portal = pd.DataFrame(records)

    if portal.empty:
        return portal

    # Ignore rows where no election could be identified.
    # The validator is intended to compare the election attached to each store.
    portal = portal[portal["99B Election"] != ""].copy()

    # Normalize state/election.
    portal["99B Election"] = portal["99B Election"].map(normalize_election)

    # Each store can appear multiple times because of vendor zones.
    # Collapse to one election per store.
    grouped = (
        portal.groupby("Store", as_index=False)
        .agg(
            State=("State", lambda s: next((x for x in s if x), "")),
            Elections=("99B Election", lambda s: sorted(set(x for x in s if x))),
        )
    )

    grouped["99B Election"] = grouped["Elections"].apply(
        lambda x: x[0] if len(x) == 1 else " / ".join(x)
    )
    grouped["Multiple Elections"] = grouped["Elections"].apply(lambda x: len(x) > 1)

    return grouped.drop(columns=["Elections"])


# -----------------------------
# Read Election Report
# -----------------------------
def read_election_report(uploaded_file):
    """
    Election Report starts at Excel row 5, meaning pandas header=4.

    Required:
      D = Account Name -> digits only -> matching Store key
      E = State
      F = Election
    """
    report = pd.read_excel(uploaded_file, header=4)

    if report.shape[1] < 6:
        raise ValueError("Election Report must have at least columns A:F.")

    # Use physical columns D/E/F as requested.
    account_name = report.iloc[:, 3]
    state = report.iloc[:, 4]
    election = report.iloc[:, 5]

    out = pd.DataFrame(
        {
            "Store": account_name.map(clean_store),
            "State": state.fillna("").astype(str).str.strip().str.upper(),
            "Report Election": election.map(normalize_election),
        }
    )

    # Only valid store numbers are matching keys.
    out = out[out["Store"].str.len() >= 6].copy()

    # If the report contains duplicate rows for the same store with the
    # same election, keep one. If there are different elections, retain
    # the most recent occurrence for a deterministic validation.
    out = out.drop_duplicates(subset=["Store", "State", "Report Election"])

    duplicate_store_counts = out.groupby("Store")["Report Election"].nunique()
    conflicting_stores = set(
        duplicate_store_counts[duplicate_store_counts > 1].index
    )

    # For normal stores, one row is enough.
    # For conflicting report elections, combine them so the issue is visible.
    report_grouped = (
        out.groupby("Store", as_index=False)
        .agg(
            State=("State", lambda s: next((x for x in s if x), "")),
            Report_Elections=("Report Election", lambda s: sorted(set(x for x in s if x))),
        )
    )

    report_grouped["Report Election"] = report_grouped["Report_Elections"].apply(
        lambda x: x[0] if len(x) == 1 else " / ".join(x)
    )
    report_grouped["Multiple Report Elections"] = report_grouped[
        "Report_Elections"
    ].apply(lambda x: len(x) > 1)

    return report_grouped.drop(columns=["Report_Elections"])


# -----------------------------
# Validation
# -----------------------------
def validate(portal, report):
    result = portal.merge(
        report,
        on="Store",
        how="outer",
        suffixes=("_99B", "_Report"),
    )

    result["State"] = result["State_99B"].replace("", pd.NA)
    result["State"] = result["State"].fillna(result["State_Report"]).fillna("")

    result["99B Election"] = result["99B Election"].fillna("")
    result["Report Election"] = result["Report Election"].fillna("")

    def get_status(row):
        p = row["99B Election"]
        r = row["Report Election"]

        if row.get("Multiple Elections", False):
            return "MULTIPLE 99B ELECTIONS"

        if row.get("Multiple Report Elections", False):
            return "MULTIPLE REPORT ELECTIONS"

        if p and r:
            return "MATCH" if p == r else "MISMATCH"

        if p:
            return "MISSING IN REPORT"

        if r:
            return "MISSING IN 99B"

        return "NO ELECTION"

    result["Status"] = result.apply(get_status, axis=1)

    final = result[
        ["State", "Store", "99B Election", "Report Election", "Status"]
    ].copy()

    final["Store"] = final["Store"].astype(str)

    return final


def make_excel(result):
    output = BytesIO()

    summary = (
        result["Status"]
        .value_counts()
        .rename_axis("Status")
        .reset_index(name="Count")
    )

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="Validation")
        summary.to_excel(writer, index=False, sheet_name="Summary")

        # Only exceptions
        exceptions = result[result["Status"] != "MATCH"].copy()
        exceptions.to_excel(writer, index=False, sheet_name="Exceptions")

    output.seek(0)
    return output


# -----------------------------
# UI
# -----------------------------
st.title("PSO Election Validator")
st.caption("Validate 99 Bottles store elections against the supplier Election Report.")

st.info(
    "Upload the Store List and Election Report. The tool compares stores using "
    "the store number and validates the 99B election against the supplier election."
)

col1, col2 = st.columns(2)

with col1:
    store_file = st.file_uploader(
        "1. Store List",
        type=["xlsx", "xls"],
        key="store_list",
        help="Column B on every tab is the Zone used to determine the 99B Election.",
    )

with col2:
    report_file = st.file_uploader(
        "2. Election Report",
        type=["xlsx", "xls"],
        key="election_report",
        help="Header starts on row 5. Column D = Account Name, E = State, F = Election.",
    )

if store_file and report_file:
    try:
        with st.spinner("Reading and validating files..."):
            portal = read_store_list(store_file)
            report = read_election_report(report_file)
            result = validate(portal, report)

        st.success("Validation complete.")

        # Summary cards
        counts = result["Status"].value_counts()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Stores Checked", f"{len(result):,}")
        c2.metric("Matches", f"{counts.get('MATCH', 0):,}")
        c3.metric("Mismatches", f"{counts.get('MISMATCH', 0):,}")
        c4.metric(
            "Missing",
            f"{counts.get('MISSING IN 99B', 0) + counts.get('MISSING IN REPORT', 0):,}",
        )

        st.subheader("Validation Results")

        status_filter = st.multiselect(
            "Filter Status",
            options=sorted(result["Status"].unique()),
            default=sorted(result["Status"].unique()),
        )

        filtered = result[result["Status"].isin(status_filter)].copy()

        st.dataframe(
            filtered,
            use_container_width=True,
            hide_index=True,
        )

        st.download_button(
            "Download Validation Excel",
            data=make_excel(result),
            file_name="PSO_Election_Validation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Unable to process the files: {e}")
        st.exception(e)
else:
    st.markdown(
        """
### Expected file structure

**Store List**
- Reads **all tabs**
- **Column B = Zone**
- Extracts the 99B Election from the Zone
- Store number is the matching key

**Election Report**
- Header begins on **row 5**
- **Column D = Account Name** → numbers only
- **Column E = State**
- **Column F = Election**

### Output

`State | Store | 99B Election | Report Election | Status`
"""
    )
