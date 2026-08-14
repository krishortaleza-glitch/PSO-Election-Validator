
import re
import csv
from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="PSO Election Validator", page_icon="✅", layout="wide")



def clean_store(value):
    if pd.isna(value):
        return ""
    return re.sub(r"\D", "", str(value))


def extract_99b_election(zone):
    """
    Extract the 99B Election from Store List Column B (Zone).

    Rules:
      1. If the zone contains a state abbreviation, the election is the
         first recognized election designation appearing AFTER the first
         state abbreviation.
      2. If there is no state abbreviation, use the first recognized election
         designation anywhere in the zone.
      3. Later election-like words are ignored.

    This prevents examples such as:
        MO Blended_EDLP Balanced_StL
    from being incorrectly classified as Balanced.
    The result is Blended.
    """
    if pd.isna(zone):
        return ""

    z = str(zone).strip()
    if not z:
        return ""

    if re.search(r"\bNO[\s_-]*(?:TOBACCO|DISCOUNT)\b", z, re.I):
        return "OPT-OUT"

    elections = [
        ("Enhanced Margin", re.compile(r"(?<![A-Z])ENHANCED(?:\s*[_-]?\s*MARGIN|\s*\d+)?(?![A-Z])", re.I)),
        ("Blended", re.compile(r"(?<![A-Z])BLENDED(?![A-Z])", re.I)),
        ("Reserve", re.compile(r"(?<![A-Z])RESERVE(?![A-Z])", re.I)),
        ("Balanced", re.compile(r"(?<![A-Z])BALANCED(?![A-Z])", re.I)),
        ("Margin", re.compile(r"(?<![A-Z])MARGIN(?![A-Z])", re.I)),
        ("CVF", re.compile(r"(?<![A-Z])CVF(?![A-Z])", re.I)),
    ]

    state_matches = list(
        re.finditer(r"(?<![A-Z])([A-Z]{2})(?![A-Z])", z.upper())
    )

    if state_matches:
        after_state = z[state_matches[0].end():]
        matches = []

        for label, pattern in elections:
            match = pattern.search(after_state)
            if match:
                matches.append((match.start(), label))

        if matches:
            return min(matches, key=lambda x: x[0])[1]

    matches = []
    for label, pattern in elections:
        match = pattern.search(z)
        if match:
            matches.append((match.start(), label))

    return min(matches, key=lambda x: x[0])[1] if matches else ""


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
def extract_state_from_zone(zone):
    """Return the first valid US state abbreviation found in the zone."""
    if pd.isna(zone):
        return ""

    z = str(zone).strip().upper()
    states = {
        "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN",
        "IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV",
        "NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN",
        "TX","UT","VT","VA","WA","WV","WI","WY","DC"
    }

    for match in re.finditer(r"(?<![A-Z])([A-Z]{2})(?![A-Z])", z):
        if match.group(1) in states:
            return match.group(1)

    return ""


def read_store_list_csv(uploaded_file):
    """
    Store List is CSV only.

    The validator only uses physical CSV Column A (Store) and Column B (Zone).
    The uploaded Store List may contain an unquoted comma in a later field,
    so pandas can shift those later columns. Reading the raw first two fields
    keeps Store and Zone aligned correctly.
    """
    uploaded_file.seek(0)
    raw = uploaded_file.read()

    if isinstance(raw, str):
        raw = raw.encode("utf-8")

    text = raw.decode("utf-8-sig", errors="replace")
    rows = csv.reader(text.splitlines())

    header = next(rows, None)
    if not header or len(header) < 2:
        raise ValueError("Store List CSV must contain Column A = Store and Column B = Zone.")

    records = []

    for row in rows:
        if len(row) < 2:
            continue

        store = clean_store(row[0])
        zone = row[1].strip() if row[1] else ""

        if not store or not zone:
            continue

        election = extract_99b_election(zone)
        if not election:
            continue

        records.append(
            {
                "Store": store,
                "State": extract_state_from_zone(zone),
                "99B Election": election,
            }
        )

    portal = pd.DataFrame(records)

    if portal.empty:
        return pd.DataFrame(
            columns=["Store", "State", "99B Election", "Multiple Elections"]
        )

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

st.caption("Compare 99B elections against the supplier Election Report.")

col1, col2 = st.columns(2)

with col1:
    store_file = st.file_uploader(
        "Store List CSV",
        type=["csv"],
        key="store_csv",
    )

with col2:
    report_file = st.file_uploader(
        "Election Report XLSX",
        type=["xlsx"],
        key="report_xlsx",
    )

if store_file and report_file:
    try:
        with st.spinner("Validating elections..."):
            portal = read_store_list_csv(store_file)
            report = read_election_report_xlsx(report_file)

            if portal.empty:
                raise ValueError(
                    "No valid stores were found in the Store List CSV. "
                    "Column A must be Store and Column B must be Zone."
                )

            result = validate(portal, report)

        st.success("Validation complete.")

        counts = result["Status"].value_counts()

        c1, c2, c3 = st.columns(3)
        c1.metric("Stores Checked", f"{len(result):,}")
        c2.metric("Matches", f"{counts.get('MATCH', 0):,}")
        c3.metric("Missing in Report", f"{counts.get('MISSING IN REPORT', 0):,}")

        st.subheader("Validation Results")

        st.dataframe(
            result,
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
