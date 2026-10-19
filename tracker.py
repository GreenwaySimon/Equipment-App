import datetime
import os
import pandas as pd
import streamlit as st

EXCEL_FILE = "Chainsaw_Equipment_Tracker.xlsx"
EXCEL_PATH = EXCEL_FILE

COLUMNS = [
    "Timestamp",
    "Asset ID",
    "Status",
    "Instructor",
    "Assessment Date",
    "Inspection",
    "Notes",
]


def load_data():
  if os.path.exists(EXCEL_FILE):
    try:
      df = pd.read_excel(EXCEL_FILE)
      for col in COLUMNS:
        if col not in df.columns:
          df[col] = ""
      return df
    except Exception:
      return pd.DataFrame(columns=COLUMNS)
  else:
    return pd.DataFrame(columns=COLUMNS)


def save_data(df):
  df.to_excel(EXCEL_FILE, index=False)


def run():
  # Mobile-Friendly CSS Styling & Large Touch Targets
  st.markdown(
      """
    <style>
    div.stButton > button, div.stFormSubmitButton > button {
        font-size: 18px !important;
        padding: 0.75rem 1rem !important;
        width: 100% !important;
        min-height: 50px !important;
    }
    div.row-widget.stRadio > div {
        font-size: 18px !important;
    }
    h1, h2, h3 {
        word-break: break-word;
    }
    table {
        width: 100% !important;
        font-size: 14px !important;
    }
    th, td {
        word-break: break-word !important;
        padding: 8px 4px !important;
    }
    </style>
    """,
      unsafe_allow_html=True,
  )

  df = load_data()

  # Left Sidebar Navigation Menu
  st.sidebar.title("Menu")
  nav_option = st.sidebar.radio(
      "Navigation", ["Equipment Log Action", "Live Dashboard"]
  )

  st.markdown("---")

  # Section 1: Equipment Log Action
  if nav_option == "Equipment Log Action":
    st.header("Equipment Log Action")
    st.write("Fill out the form below to record an entry.")

    with st.form("log_form"):
      asset_id = st.text_input(
          "Asset ID (Enter equipment number/tag)", placeholder="e.g. SAW-01"
      )
      status = st.selectbox(
          "Status", ["Operational", "Needs Maintenance", "Out of Service"]
      )
      instructor = st.text_input("Instructor Name")
      assessment_date = st.date_input(
          "Assessment Date", value=datetime.date.today()
      )

      # Inspection dropdown defaulting to "Not Inspected"
      inspection = st.selectbox(
          "Inspection Result",
          ["Not Inspected", "Inspected - Passed", "Inspected - Failed"],
          index=0,
      )

      notes = st.text_area("Notes / Details")

      submitted = st.form_submit_button("Submit Entry")

      if submitted:
        if asset_id and instructor:
          new_entry = pd.DataFrame([{
              "Timestamp": datetime.datetime.now().strftime(
                  "%Y-%m-%d %H:%M:%S"
              ),
              "Asset ID": asset_id.strip(),
              "Status": status,
              "Instructor": instructor.strip(),
              "Assessment Date": assessment_date.strftime("%Y-%m-%d"),
              "Inspection": inspection,
              "Notes": notes,
          }])
          df = pd.concat([df, new_entry], ignore_index=True)
          save_data(df)
          st.success("Log submitted successfully!")
          st.rerun()
        else:
          st.error("Please fill in both Asset ID and Instructor Name.")

  # Section 2: Live Dashboard (Excel spreadsheet replication view)
  elif nav_option == "Live Dashboard":
    st.header("Live Dashboard")
    st.write("Complete log history replicating the Excel spreadsheet view.")

    if not df.empty:
      st.table(df)
    else:
      st.info("No equipment logs recorded yet.")
