import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timezone

EXCEL_FILE = "Anushil_Organization_Skip_Level_Pairings_v4.xlsx"

st.set_page_config(
    page_title="Cross-functional Skip-Level Meeting Pairings",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #f8fbff 0%, #eef4ff 45%, #f7f4ff 100%);
}
.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
    max-width: 1200px;
}
.main-title {
    font-size: 2.3rem;
    font-weight: 800;
    color: #1f2a44;
    margin-bottom: 0.2rem;
}
.sub-title {
    font-size: 1rem;
    color: #5b6475;
    margin-bottom: 1.5rem;
}
.metric-card {
    background: linear-gradient(180deg, #ffffff 0%, #f6f9ff 100%);
    border: 1px solid rgba(98, 118, 255, 0.14);
    border-radius: 16px;
    padding: 0.9rem 1rem;
    box-shadow: 0 6px 16px rgba(31, 42, 68, 0.05);
}
.metric-label {
    font-size: 0.9rem;
    color: #667085;
    margin-bottom: 0.35rem;
}
.metric-value {
    font-size: 1.6rem;
    font-weight: 700;
    color: #1f2a44;
}
div[data-testid="stRadio"] > div {
    background: rgba(255,255,255,0.72);
    padding: 0.65rem 0.9rem;
    border-radius: 14px;
    border: 1px solid rgba(120, 140, 180, 0.18);
}
div[data-testid="stTextInput"] > div > div input {
    background-color: white;
    border-radius: 12px;
}
div[data-testid="stDataEditor"] {
    border-radius: 16px;
    overflow: hidden;
    border: 1px solid rgba(120, 140, 180, 0.18);
}
.stDownloadButton button, .stButton button {
    border-radius: 10px;
    font-weight: 600;
}
hr {
    margin-top: 1rem;
    margin-bottom: 1rem;
    border: none;
    height: 1px;
    background: linear-gradient(to right, transparent, rgba(100,100,140,0.35), transparent);
}
</style>
""", unsafe_allow_html=True)

# ---------- Config from Streamlit secrets ----------
GSHEET_URL = st.secrets["connections"]["gsheets"]["spreadsheet"]
UPDATES_WORKSHEET = "updates"

def get_gspread_client():
    info = {
        "type": st.secrets["connections"]["gsheets"]["type"],
        "project_id": st.secrets["connections"]["gsheets"]["project_id"],
        "private_key_id": st.secrets["connections"]["gsheets"]["private_key_id"],
        "private_key": st.secrets["connections"]["gsheets"]["private_key"],
        "client_email": st.secrets["connections"]["gsheets"]["client_email"],
        "client_id": st.secrets["connections"]["gsheets"]["client_id"],
        "auth_uri": st.secrets["connections"]["gsheets"]["auth_uri"],
        "token_uri": st.secrets["connections"]["gsheets"]["token_uri"],
        "auth_provider_x509_cert_url": st.secrets["connections"]["gsheets"]["auth_provider_x509_cert_url"],
        "client_x509_cert_url": st.secrets["connections"]["gsheets"]["client_x509_cert_url"],
    }

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds)

def get_updates_worksheet():
    client = get_gspread_client()
    spreadsheet = client.open_by_url(GSHEET_URL)
    return spreadsheet.worksheet(UPDATES_WORKSHEET)

def load_updates() -> pd.DataFrame:
    try:
        ws = get_updates_worksheet()
        records = ws.get_all_records()
        if not records:
            return pd.DataFrame(columns=["row_key", "model", "status", "comments", "last_updated"])
        df = pd.DataFrame(records)
        for col in ["row_key", "model", "status", "comments", "last_updated"]:
            if col not in df.columns:
                df[col] = ""
        return df[["row_key", "model", "status", "comments", "last_updated"]].copy()
    except Exception as e:
        st.error("Failed to read updates sheet.")
        st.exception(e)
        return pd.DataFrame(columns=["row_key", "model", "status", "comments", "last_updated"])

def save_updates(df_updates: pd.DataFrame) -> None:
    ws = get_updates_worksheet()
    df_updates = df_updates.fillna("")
    data = [df_updates.columns.tolist()] + df_updates.astype(str).values.tolist()
    ws.clear()
    ws.update("A1", data)

def upsert_updates(model_name: str, edited_df: pd.DataFrame) -> None:
    updates = load_updates()

    existing_other_models = updates[updates["model"] != model_name].copy()

    current_model_updates = edited_df[["row_key", "Status", "Comments"]].copy()
    current_model_updates = current_model_updates.rename(
        columns={"Status": "status", "Comments": "comments"}
    )
    current_model_updates["model"] = model_name
    current_model_updates["last_updated"] = datetime.now(timezone.utc).isoformat()

    merged = pd.concat([existing_other_models, current_model_updates], ignore_index=True)
    merged = merged.drop_duplicates(subset=["model", "row_key"], keep="last")
    save_updates(merged)

# ---------- Load pairing data ----------
df_11 = pd.read_excel(EXCEL_FILE, sheet_name="Q1 1-1 Allowed Dir Only")
df_12 = pd.read_excel(EXCEL_FILE, sheet_name="Q1 1-2 Allowed Dir Only")

REMOVE_PATTERNS = [
    "full stack developer intern",
    "sales product line intern",
    "unfilled",
]

def should_remove(series: pd.Series) -> pd.Series:
    mask = pd.Series(False, index=series.index)
    for pattern in REMOVE_PATTERNS:
        mask = mask | series.astype(str).str.contains(pattern, case=False, na=False)
    return mask

df_11 = df_11[~should_remove(df_11["IC Name"])].copy()
df_12 = df_12[
    ~should_remove(df_12["IC1 Name"])
    & ~should_remove(df_12["IC2 Name"])
].copy()

df_11["IC Team"] = df_11["IC Team"].fillna("").astype(str).str.strip()
df_11.loc[df_11["IC Team"] == "", "IC Team"] = df_11["IC Title"]

for prefix in ["IC1", "IC2"]:
    df_12[f"{prefix} Team"] = df_12[f"{prefix} Team"].fillna("").astype(str).str.strip()
    df_12.loc[df_12[f"{prefix} Team"] == "", f"{prefix} Team"] = df_12[f"{prefix} Title"]

STATUS_OPTIONS = ["", "Scheduled", "Done", "Cancelled"]

# ---------- Header ----------
st.markdown('<div class="main-title">Cross-functional Skip-Level Meeting Pairings</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">Search by IC or Director, track meeting progress, and add comments in one place.</div>',
    unsafe_allow_html=True
)

col1, col2 = st.columns([1, 2])
with col1:
    model = st.radio(
        "Select Model",
        ["1:1 Model", "1:2 Model"],
        horizontal=True
    )
with col2:
    name = st.text_input("Enter employee name (IC or Director)")

st.markdown("---")

# ---------- Helpers ----------
def merge_updates(results: pd.DataFrame, model_name: str) -> pd.DataFrame:
    updates = load_updates()
    updates_model = updates[updates["model"] == model_name].copy()
    if updates_model.empty:
        results["Comments"] = ""
        results["Status"] = ""
        return results

    results = results.merge(
        updates_model[["row_key", "status", "comments"]],
        on="row_key",
        how="left",
    )
    results["Comments"] = results["comments"].fillna("")
    results["Status"] = results["status"].fillna("")
    results = results.drop(columns=["status", "comments"], errors="ignore")
    return results

def add_edit_columns_11(results: pd.DataFrame) -> pd.DataFrame:
    results = results.copy()
    results["row_key"] = (
        results["Director Name"].astype(str) + " | " +
        results["IC Name"].astype(str)
    )
    return merge_updates(results, "1:1")

def add_edit_columns_12(results: pd.DataFrame) -> pd.DataFrame:
    results = results.copy()
    results["row_key"] = (
        results["Director Name"].astype(str) + " | " +
        results["IC1 Name"].astype(str) + " | " +
        results["IC2 Name"].astype(str)
    )
    return merge_updates(results, "1:2")

def metric_box(label: str, value: int):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

def show_progress_dashboard_11():
    base = df_11.copy()
    base["row_key"] = (
        base["Director Name"].astype(str) + " | " +
        base["IC Name"].astype(str)
    )
    base = merge_updates(base, "1:1")
    statuses = base["Status"].fillna("")

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        metric_box("Total Meetings", len(base))
    with c2:
        metric_box("Scheduled", int((statuses == "Scheduled").sum()))
    with c3:
        metric_box("Done", int((statuses == "Done").sum()))
    with c4:
        metric_box("Cancelled", int((statuses == "Cancelled").sum()))
    with c5:
        metric_box("Blank", int((statuses == "").sum()))

def show_progress_dashboard_12():
    base = df_12.copy()
    base["row_key"] = (
        base["Director Name"].astype(str) + " | " +
        base["IC1 Name"].astype(str) + " | " +
        base["IC2 Name"].astype(str)
    )
    base = merge_updates(base, "1:2")
    statuses = base["Status"].fillna("")

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        metric_box("Total Meetings", len(base))
    with c2:
        metric_box("Scheduled", int((statuses == "Scheduled").sum()))
    with c3:
        metric_box("Done", int((statuses == "Done").sum()))
    with c4:
        metric_box("Cancelled", int((statuses == "Cancelled").sum()))
    with c5:
        metric_box("Blank", int((statuses == "").sum()))

st.subheader("Progress Dashboard")
if model == "1:1 Model":
    show_progress_dashboard_11()
else:
    show_progress_dashboard_12()

st.markdown("---")
st.subheader("Pairing Results")

# ---------- Results ----------
if name:
    if model == "1:1 Model":
        results = df_11[
            df_11["IC Name"].astype(str).str.contains(name, case=False, na=False)
            | df_11["Director Name"].astype(str).str.contains(name, case=False, na=False)
        ].copy()

        if results.empty:
            st.warning("No matching mapping found.")
        else:
            results = add_edit_columns_11(results)

            display_cols = [
                "Director Name",
                "Director Title",
                "Director Team",
                "IC Name",
                "IC Title",
                "IC Team",
                "Comments",
                "Status",
            ]

            edited = st.data_editor(
                results[display_cols + ["row_key"]],
                hide_index=True,
                use_container_width=True,
                disabled=[
                    "Director Name",
                    "Director Title",
                    "Director Team",
                    "IC Name",
                    "IC Title",
                    "IC Team",
                    "row_key",
                ],
                column_config={
                    "Comments": st.column_config.TextColumn("Comments"),
                    "Status": st.column_config.SelectboxColumn(
                        "Status",
                        options=STATUS_OPTIONS,
                        required=False,
                    ),
                    "row_key": None,
                },
                key="editor_11",
            )

            if st.button("Save 1:1 updates"):
                upsert_updates("1:1", edited)
                st.success("Saved.")
                st.rerun()

            export_df = edited.drop(columns=["row_key"])
            st.download_button(
                "Download current 1:1 results as CSV",
                export_df.to_csv(index=False).encode("utf-8"),
                file_name="skip_level_pairings_1_1.csv",
                mime="text/csv",
            )

    else:
        results = df_12[
            df_12["IC1 Name"].astype(str).str.contains(name, case=False, na=False)
            | df_12["IC2 Name"].astype(str).str.contains(name, case=False, na=False)
            | df_12["Director Name"].astype(str).str.contains(name, case=False, na=False)
        ].copy()

        if results.empty:
            st.warning("No matching mapping found.")
        else:
            results = add_edit_columns_12(results)

            display_cols = [
                "Director Name",
                "Director Title",
                "Director Team",
                "IC1 Name",
                "IC1 Title",
                "IC1 Team",
                "IC2 Name",
                "IC2 Title",
                "IC2 Team",
                "Comments",
                "Status",
            ]

            edited = st.data_editor(
                results[display_cols + ["row_key"]],
                hide_index=True,
                use_container_width=True,
                disabled=[
                    "Director Name",
                    "Director Title",
                    "Director Team",
                    "IC1 Name",
                    "IC1 Title",
                    "IC1 Team",
                    "IC2 Name",
                    "IC2 Title",
                    "IC2 Team",
                    "row_key",
                ],
                column_config={
                    "Comments": st.column_config.TextColumn("Comments"),
                    "Status": st.column_config.SelectboxColumn(
                        "Status",
                        options=STATUS_OPTIONS,
                        required=False,
                    ),
                    "row_key": None,
                },
                key="editor_12",
            )

            if st.button("Save 1:2 updates"):
                upsert_updates("1:2", edited)
                st.success("Saved.")
                st.rerun()

            export_df = edited.drop(columns=["row_key"])
            st.download_button(
                "Download current 1:2 results as CSV",
                export_df.to_csv(index=False).encode("utf-8"),
                file_name="skip_level_pairings_1_2.csv",
                mime="text/csv",
            )
else:
    st.info("Search for an employee name to view their pairing details.")
