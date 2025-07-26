import pandas as pd
from datetime import timedelta
import streamlit as st
import calendar
import plotly.express as px
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1
import json

st.set_page_config(page_title="Smart Records",
                    page_icon='📚',
                    layout='wide',
                    initial_sidebar_state='expanded')

st.image("assets/diligent_header.png")
# st.title("🏗️ Diligent Supplies Limited - Store Management System")
with open("assets/styles.css") as f:
    st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

@st.cache_resource
def get_worksheet(sheet_id: str, worksheet_name: str):
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=[
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/spreadsheets"
    ])
    client = gspread.authorize(creds)
    return client.open(sheet_id).worksheet(worksheet_name)

worksheet = get_worksheet(sheet_id="Material Algorithm", worksheet_name="Data")

# @st.cache_data()
def load_data_sheet():
    data_ws = worksheet.spreadsheet.worksheet("Data")
    df = pd.DataFrame(data_ws.get_all_records())
    df.columns = df.columns.str.strip()
    return df

def next_workday(start_date, days):
    current = start_date
    added_days = 0
    while added_days < days:
        current += timedelta(days=1)
        if current.weekday() != 6:  # Skip Sundays
            added_days += 1
    return current


def workdays_only(start, end):
    """Generate list of dates excluding Sundays between start and end inclusive."""
    current = start
    dates = []
    while current <= end:
        if current.weekday() != 6:
            dates.append(current)
        current += timedelta(days=1)
    return dates


def generate_work_schedule(data, start_date_str, status_durations, priority_order):
    start_date = pd.to_datetime(start_date_str)
    schedule_rows = []
    team_tracker = {}

    def get_priority(row):
        constituency_priority = priority_order.get("Constituency", [])
        status_priority = priority_order.get("Status", [])
        constituency_rank = constituency_priority.index(row["Constituency"]) if row["Constituency"] in constituency_priority else len(constituency_priority)
        status_rank = status_priority.index(row["Status"]) if row["Status"] in status_priority else len(status_priority)
        return (row["Team"], constituency_rank, status_rank, row["Scheme Name"])

    data = data.copy()
    data["_priority"] = data.apply(get_priority, axis=1)
    data = data.sort_values(by="_priority")

    for _, row in data.iterrows():
        team = row["Team"]
        constituency = row["Constituency"]
        scheme = row["Scheme Name"]

        if "Status" in row and row["Status"] in status_durations:
            duration = status_durations[row["Status"]]
        elif "Duration (Days)" in row and not pd.isna(row["Duration (Days)"]):
            duration = int(row["Duration (Days)"])
        else:
            duration = 2

        last_end = team_tracker.get(team, start_date - timedelta(days=1))
        start = next_workday(last_end, 1)
        end = start
        days_added = 1
        while days_added < duration:
            end += timedelta(days=1)
            if end.weekday() != 6:
                days_added += 1

        schedule_rows.append({
            "Team": team,
            "Constituency": constituency,
            "Scheme Name": scheme,
            "Duration (Days)": duration,
            "Start Date": start,
            "End Date": end,
            "Status": row.get("Status", "Unknown")
        })

        team_tracker[team] = end

    return pd.DataFrame(schedule_rows)

def append_stay_totals(output_df):
    stay_df = load_data_sheet().dropna(how="all")
    
    stay_cols = [
        "NORMAL STAY", "FLYING STAY", "UNDER WP HT NORMAL STAY", "UNDER CP HT NORMAL STAY",
        "UNDER WP HT FLYING STAY", "UNDER CP HT FLYING STAY", "MV NORMAL STAY", "MV FLYING STAY"
    ]
    stay_df[stay_cols] = stay_df[stay_cols].apply(pd.to_numeric, errors="coerce")
    totals = stay_df.groupby("SCHEME NAME")[stay_cols].sum().reset_index()
    totals["TOTAL STAYS"] = totals[stay_cols].sum(axis=1)
    merged = output_df.merge(totals, left_on="Scheme Name", right_on="SCHEME NAME", how="left")
    merged.drop(columns=["SCHEME NAME"], inplace=True)
    merged[stay_cols] = merged[stay_cols].fillna(0).astype(int)
    return merged

st.title("Work Schedule Generator")

uploaded_file = st.file_uploader("Upload Excel file with columns: 'Team', 'Constituency', 'Scheme Name', 'Status' or 'Duration (Days)'", type=["xlsx"])
start_date_input = st.date_input("Select the starting date")

st.sidebar.header("Custom Duration Mapping")
not_done_duration = st.sidebar.number_input("Not Done Duration", min_value=1, value=7)
pending_duration = st.sidebar.number_input("Pending Duration", min_value=1, value=5)
in_progress_duration = st.sidebar.number_input("In Progress Duration", min_value=1, value=3)
complete_duration = st.sidebar.number_input("Complete Duration", min_value=1, value=1)

status_durations = {
    "Not Done": not_done_duration,
    "Pending": pending_duration,
    "In Progress": in_progress_duration,
    "Complete": complete_duration
}

if uploaded_file and start_date_input:
    input_df = pd.read_excel(uploaded_file)
    constituency_options = sorted(input_df["Constituency"].dropna().unique())
    status_options = sorted(input_df["Status"].dropna().unique())

    st.sidebar.header("Prioritization Order")
    constituency_priority = st.sidebar.multiselect("Prioritize Constituencies (Top to Bottom)", constituency_options, default=constituency_options)
    status_priority = st.sidebar.multiselect("Prioritize Status (Top to Bottom)", status_options, default=status_options)

    priority_order = {
        "Constituency": constituency_priority,
        "Status": status_priority
    }
    
    output_df = generate_work_schedule(input_df, start_date_input.strftime('%Y-%m-%d'), status_durations, priority_order)
    full_df = append_stay_totals(output_df)
    full_df["Start Date"] = full_df["Start Date"].dt.strftime("%Y-%m-%d")
    full_df["End Date"] = full_df["End Date"].dt.strftime("%Y-%m-%d")
    st.write("### Generated Work Schedule")
    st.dataframe(full_df)

    with st.expander("### Calendar View by Scheme Name"):
        # st.write("### Calendar View by Scheme Name")
        cal_df = output_df.copy()
        cal_df["Start Date"] = pd.to_datetime(cal_df["Start Date"])
        cal_df["End Date"] = pd.to_datetime(cal_df["End Date"])
        cal_expanded = []
        for _, row in cal_df.iterrows():
            for d in workdays_only(row["Start Date"], row["End Date"]):
                cal_expanded.append({"Date": d, "Scheme Name": row["Scheme Name"], "Team": row["Team"]})
        calendar_df = pd.DataFrame(cal_expanded)
        calendar_df = calendar_df.sort_values("Date")
        
        pivot_df = calendar_df.pivot_table(index="Date", columns="Team", values="Scheme Name", aggfunc=lambda x: ', '.join(x))
        pivot_df.reset_index(inplace=True) 
        pivot_df["Date"] = pd.to_datetime(pivot_df["Date"])
        pivot_df["Day"] = pivot_df["Date"].dt.day_name()
        cols = pivot_df.columns.tolist()
        if "Day" in cols and "Date" in cols:
            cols.insert(cols.index("Date"), cols.pop(cols.index("Day")))
            pivot_df = pivot_df[cols]
        pivot_df["Date"] = pivot_df["Date"].dt.strftime("%Y-%m-%d")
        st.dataframe(pivot_df)

    output_file = "Work_Schedule_Output.xlsx"
    full_df.to_excel(output_file, index=False)
    with open(output_file, "rb") as f:
        st.download_button("Download Excel File", f, file_name=output_file)
