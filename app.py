
import re
from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="PSO Election Validator", page_icon="✅", layout="wide")

APP_VERSION = "v1.2 — Store List Controls Output Population"


def clean_store(value):
    if pd.isna(value):
        return ""
    return re.sub(r"\D", "", str(value))


def extract_99b_election(zone):
    if pd.isna(zone):
        return ""

    z = str(zone).strip().upper()

    # Order matters for more specific names.
    if "ENHANCED" in z:
        return "Enhanced Margin"
    if "RESERVE" in z:
        return "Reserve"
    if "BALANCED" in z:
        return "Balanced"
    if "BLENDED" in z:
        return "Blended"
    if "MARGIN" in z:
        return "Margin"

    return ""


def normalize_election(value):
    if pd.isna(value):
        return ""

    v = str(value).strip().upper()
    if not v or v == "NAN":
        return ""

    if "ENHANCED" in v:
        return "Enhanced Margin"
    if "RESERVE" in v:
        return "Reserve"
    if "BALANCED" in v:
        return "Balanced"
    if "BLENDED" in v:
        return "Blended"
    if "MARGIN" in v:
        return "Margin"

    return str(value).strip()


# ============================================================
# STORE LIST — CSV ONLY
# ============================================================
def read_store_list_csv(uploaded_file):
    """
    Store List is CSV only.
    Column B = Zone.
    The first column is the Store.
    A State column is used if present.
    """
    df = pd.read_csv(uploaded_file, dtype=str)

    if df.shape[1] < 2:
        raise ValueError("Store List CSV must contain at least columns A and B.")

    store_col = df.columns[0]
    zone_col = df.columns[1]

    state_col = next(
        (c for c in df.columns if str(c).strip().lower() == "state"),
        None,
    )

    records = []

    for _, row in df.iterrows():
        store = clean_store(row.iloc[0])
        if not store:
            continue

        zone = row.iloc[1]
        election = extract_99b_election(zone)

        state = ""
        if state_col is not None and pd.notna(row[state_col]):
            state = str(row[state_col]).strip().upper()

        records.append(
            {
                "Store": store,
                "State": state,
                "99B Election": election,
            }
        )

    portal = pd.DataFrame(records)

    if portal.empty:
        return pd.DataFrame(
            columns=["Store", "State", "99B Election", "Multiple Elections"]
        )

    portal = portal[portal["99B Election"] != ""].copy()

    grouped = (
        portal.groupby("Store", as_index=False)
        .agg(
            State=("State", lambda s: next((x for x in s if x), "")),
            Elections=("99B Election", lambda s: sorted(set(s))),
        )
    )

    grouped["Multiple Elections"] = grouped["Elections"].apply(
        lambda x: len(x) > 1
    )
    grouped["99B Election"] = grouped["Elections"].apply(
        lambda x: x[0] if len(x) == 1 else " / ".join(x)
    )

    return grouped.drop(columns=["Elections"])


# ============================================================
# ELECTION REPORT — XLSX ONLY
# ============================================================
def read_election_report_xlsx(uploaded_file):
    """
    Election Report is XLSX.
    Header is on Excel row 5 => pandas header=4.

    Column D = matching store source
    Column E = State
    Column F = Election
    """
    df = pd.read_excel(uploaded_file, header=4, dtype=str)

    if df.shape[1] < 6:
        raise ValueError("Election Report XLSX must contain at least columns A:F.")

    account = df.iloc[:, 3]
    state = df.iloc[:, 4]
    election = df.iloc[:, 5]

    report = pd.DataFrame(
        {
            "Store": account.map(clean_store),
            "State": state.fillna("").astype(str).str.strip().str.upper(),
            "Report Election": election.map(normalize_election),
        }
    )

    report = report[report["Store"] != ""].copy()

    grouped = (
        report.groupby("Store", as_index=False)
        .agg(
            State=("State", lambda s: next((x for x in s if x), "")),
            Elections=("Report Election", lambda s: sorted(set(x for x in s if x))),
        )
    )

    # Duplicate store rows are collapsed. Since duplicate rows in the
    # supplier report do not represent a separate validation result,
    # use the first election value for the store.
    grouped["Report Election"] = grouped["Elections"].apply(
        lambda x: x[0] if x else ""
    )

    return grouped.drop(columns=["Elections"])


def validate(portal, report):
    # IMPORTANT:
    # The Store List is the controlling population.
    # Only stores present in the Store List should appear in the output.
    # Stores that exist only in the Election Report are excluded.
    result = portal.merge(
        report,
        on="Store",
        how="left",
        suffixes=("_99B", "_Report"),
    )

    result["State"] = result["State_99B"].replace("", pd.NA)
    result["State"] = result["State"].fillna(result["State_Report"]).fillna("")

    result["99B Election"] = result["99B Election"].fillna("")
    result["Report Election"] = result["Report Election"].fillna("")

    def status(row):
        if row.get("Multiple Elections", False):
            return "MULTIPLE 99B ELECTIONS"

        p = row["99B Election"]
        r = row["Report Election"]

        if p and r:
            return "MATCH" if p == r else "MISMATCH"
        if p:
            return "MISSING IN REPORT"
        if r:
            return "MISSING IN 99B"
        return "NO ELECTION"

    result["Status"] = result.apply(status, axis=1)

    return result[
        ["State", "Store", "99B Election", "Report Election", "Status"]
    ].copy()


def export_excel(result):
    output = BytesIO()

    summary = (
        result["Status"]
        .value_counts()
        .rename_axis("Status")
        .reset_index(name="Count")
    )

    exceptions = result[result["Status"] != "MATCH"].copy()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="Validation")
        summary.to_excel(writer, index=False, sheet_name="Summary")
        exceptions.to_excel(writer, index=False, sheet_name="Exceptions")

    output.seek(0)
    return output


# ============================================================
# UI
# ============================================================
st.title("PSO Election Validator")
st.caption(APP_VERSION)

st.markdown(
    """
Compare the **99 Bottles store election** against the **supplier Election Report**.

**Matching key:** Store number

**Output population:** Only stores from the **Store List CSV** are included.  
Stores that exist only in the Election Report are ignored.

**Duplicate Election Report rows:** Duplicate rows for the same store are collapsed and do not create a separate status.
"""
)

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Store List")
    st.caption("CSV ONLY — one file / one dataset")
    store_file = st.file_uploader(
        "Upload Store List CSV",
        type=["csv"],
        key="store_csv",
    )

with col2:
    st.subheader("2. Election Report")
    st.caption("XLSX ONLY — header starts on Excel row 5")
    report_file = st.file_uploader(
        "Upload Election Report XLSX",
        type=["xlsx"],
        key="report_xlsx",
    )

st.divider()

st.markdown(
    """
### Expected columns

**Store List CSV**
- Column A = Store
- **Column B = Zone**
- State column, if included, is used for the output
- Election is extracted from Column B

**Election Report XLSX**
- Header = Excel row 5
- **Column D = Store source** → numbers only
- **Column E = State**
- **Column F = Election**

### Output

`State | Store | 99B Election | Report Election | Status`
"""
)

if store_file and report_file:
    try:
        with st.spinner("Validating elections..."):
            portal = read_store_list_csv(store_file)
            report = read_election_report_xlsx(report_file)
            result = validate(portal, report)

        st.success("Validation complete.")

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

        statuses = sorted(result["Status"].unique())
        selected = st.multiselect(
            "Filter Status",
            statuses,
            default=statuses,
        )

        filtered = result[result["Status"].isin(selected)]

        st.dataframe(
            filtered,
            use_container_width=True,
            hide_index=True,
        )

        st.download_button(
            "Download Validation Excel",
            data=export_excel(result),
            file_name="PSO_Election_Validation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as exc:
        st.error(f"Unable to process the files: {exc}")
else:
    st.info("Upload both files to start the validation.")
