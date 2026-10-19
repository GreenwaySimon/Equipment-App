import datetime
import os
import pandas as pd
import streamlit as st

EXCEL_FILE = "Chainsaw_Equipment_Tracker.xlsx"
EXCEL_PATH = EXCEL_FILE


def load_inventory_and_logs():
  # Load master Inventory sheet
  try:
    df_inv = pd.read_excel(EXCEL_FILE, sheet_name="Inventory")
  except Exception:
    df_inv = pd.DataFrame(columns=["Asset ID", "Category"])

  # Load log history sheet
  try:
    df_logs = pd.read_excel(EXCEL_FILE, sheet_name="Sheet1")
  except Exception:
    try:
      df_logs = pd.read_excel(EXCEL_FILE)
    except Exception:
      df_logs = pd.DataFrame(
          columns=[
              "Timestamp",
              "Asset ID",
              "Status",
              "Instructor",
              "Assessment Date",
              "Inspection",
              "Notes",
          ]
      )

  return df_inv, df_logs


def save_log(new_row_df):
  df_inv, df_logs = load_inventory_and_logs()
  df_logs = pd.concat([df_logs, new_row_df], ignore_index=True)

  with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
    df_inv.to_excel(writer, sheet_name="Inventory", index=False)
    df_logs.to_excel(writer, sheet_name="Sheet1", index=False)


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

  df_inv, df_logs = load_inventory_and_logs()

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
      inspection = st.text_input(
          "Inspection", placeholder="e.g. Passed / Failed"
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
          save_log(new_entry)
          st.success("Log submitted successfully!")
          st.rerun()
        else:
          st.error("Please fill in both Asset ID and Instructor Name.")

  # Section 2: Live Dashboard (Merging Inventory master sheet with latest logs)
  elif nav_option == "Live Dashboard":
    st.header("Live Dashboard")
    st.write("Complete log history replicating the Excel spreadsheet view.")

    if not df_inv.empty:
      if not df_logs.empty:
        df_logs_sorted = (
            df_logs.sort_values("Timestamp", ascending=True)
            if "Timestamp" in df_logs.columns
            else df_logs
        )
        latest_logs = df_logs_sorted.drop_duplicates(
            subset=["Asset ID"], keep="last"
        )
        display_df = pd.merge(df_inv, latest_logs, on="Asset ID", how="left")
      else:
        display_df = df_inv
      st.table(display_df)
    elif not df_logs.empty:
      st.table(df_logs)
    else:
      st.info("No equipment logs recorded yet.")
