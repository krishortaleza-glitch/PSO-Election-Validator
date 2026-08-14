
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
    Extract the 99B election from the Zone.

    The Store List contains many naming styles, so the election is identified
    by the first recognized election token AFTER the first valid state.
    If no state exists, the first recognized election token is used.

    Supported naming variations include:
      - Enhanced, Enhanced0, EnhancedMargin, Enhanced_Margin, EnhanPrCam65
      - Blended / Blended...
      - Reserve / Reserve...
      - Balanced / Bal...
      - Margin
      - CVF

    Later election-like terms are ignored. Example:
      MO Blended_EDLP Balanced_StL -> Blended
    """
    if pd.isna(zone):
        return ""

    z = str(zone).strip()
    if not z:
        return ""

    if re.search(r"\bNO[\s_-]*(?:TOBACCO|DISCOUNT)", z, re.I):
        return "OPT-OUT"

    election_patterns = [
        # Enhanced variants, including EnhanPrCam65 and EnhancedMargin.
        (
            "Enhanced Margin",
            re.compile(
                r"(?<![A-Z])ENHAN(?:CED)?(?:[_\-\s]*MARGIN)?",
                re.I,
            ),
        ),
        # Blended may be followed immediately by zone text.
        ("Blended", re.compile(r"(?<![A-Z])BLENDED", re.I)),
        # Reserve may be followed immediately by zone text, e.g. ReserveCamSLP.
        ("Reserve", re.compile(r"(?<![A-Z])RESERVE", re.I)),
        # Balanced shorthand appears as BalPremCamel.
        ("Balanced", re.compile(r"(?<![A-Z])BAL(?:ANCED)?", re.I)),
        ("Margin", re.compile(r"(?<![A-Z])MARGIN", re.I)),
        ("CVF", re.compile(r"(?<![A-Z])CVF", re.I)),
    ]

    valid_states = {
        "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN",
        "IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV",
        "NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN",
        "TX","UT","VT","VA","WA","WV","WI","WY","DC",
    }

    state_matches = [
        m for m in re.finditer(r"(?<![A-Z])([A-Z]{2})(?![A-Z])", z.upper())
        if m.group(1) in valid_states
    ]

    search_text = z[state_matches[0].end():] if state_matches else z

    candidates = []
    for label, pattern in election_patterns:
        match = pattern.search(search_text)
        if match:
            candidates.append((match.start(), label))

    if not candidates:
        return ""

    return min(candidates, key=lambda item: item[0])[1]


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
