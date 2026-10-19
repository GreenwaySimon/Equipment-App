import datetime
import os
import pandas as pd
import streamlit as st

EXCEL_FILE = "Chainsaw_Equipment_Tracker.xlsx"


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
              "Action Type",
              "Instructor Name",
              "Customer / Candidate Name",
              "Equipment IDs",
              "Course Date",
              "Assessment Date",
              "Inspection Result",
              "Notes / Damage Details",
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
  # Custom CSS Styling
  st.markdown(
      """
    <style>
    div.stButton > button, div.stFormSubmitButton > button {
        font-size: 18px !important;
        padding: 0.75rem 1rem !important;
        width: 100% !important;
        min-height: 50px !important;
    }
    .metric-card {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 15px;
        text-align: center;
        border: 1px solid #dee2e6;
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

  # Process equipment statuses based on latest logs
  asset_status_map = {}
  if not df_logs.empty and "Equipment IDs" in df_logs.columns:
    df_sorted = (
        df_logs.sort_values("Timestamp", ascending=True)
        if "Timestamp" in df_logs.columns
        else df_logs
    )
    for _, row in df_sorted.iterrows():
      eq_str = str(row.get("Equipment IDs", ""))
      eq_list = [e.strip() for e in eq_str.split(",") if e.strip()]
      action = str(row.get("Action Type", ""))

      for eq_id in eq_list:
        if action == "Check Out":
          asset_status_map[eq_id] = "🔴 IN USE"
        elif action == "Report Damage / Maintenance":
          asset_status_map[eq_id] = "⚠️ MAINTENANCE"
        elif action == "Return Equipment":
          insp = str(row.get("Inspection Result", ""))
          if "Fail" in insp or "Maintenance" in insp:
            asset_status_map[eq_id] = "⚠️ MAINTENANCE"
          else:
            asset_status_map[eq_id] = "🟢 AVAILABLE"

  # Determine all assets from master inventory
  all_asset_ids = []
  if not df_inv.empty and "Asset ID" in df_inv.columns:
    all_asset_ids = df_inv["Asset ID"].dropna().astype(str).tolist()

  # Exclude 'IN USE' assets for Check Out option
  available_for_checkout = [
      aid for aid in all_asset_ids if asset_status_map.get(aid) != "🔴 IN USE"
  ]

  # ----------------------------------------------------
  # SECTION 1: EQUIPMENT LOG ACTION
  # ----------------------------------------------------
  if nav_option == "Equipment Log Action":
    st.header("Equipment Log Action")
    st.write("Fill out the form below to record an entry.")

    with st.form("log_form"):
      # 1. Action Type dropdown in the first position
      action_type = st.selectbox(
          "Action Type",
          ["Check Out", "Return Equipment", "Report Damage / Maintenance"],
      )

      instructor_name = st.text_input("Instructor Name")
      customer_name = st.text_input("Customer / Candidate Name")

      # Equipment IDs selection (omits in-use gear if checking out)
      if action_type == "Check Out":
        equipment_ids = st.multiselect(
            "Equipment IDs (Available Equipment)",
            options=available_for_checkout,
        )
      else:
        equipment_ids = st.multiselect(
            "Equipment IDs",
            options=all_asset_ids if all_asset_ids else [],
        )

      col_d1, col_d2 = st.columns(2)
      with col_d1:
        course_date = st.date_input("Course Date", value=datetime.date.today())
      with col_d2:
        assessment_date = st.date_input(
            "Assessment Date", value=datetime.date.today()
        )

      inspection_result = st.selectbox(
          "Inspection Result",
          ["Pass", "Fail / Maintenance Required"],
      )

      notes = st.text_area("Notes / Damage Details")

      submitted = st.form_submit_button("Submit Entry")

      if submitted:
        if instructor_name and equipment_ids:
          eq_formatted = ", ".join(equipment_ids)
          new_entry = pd.DataFrame([{
              "Timestamp": datetime.datetime.now().strftime(
                  "%Y-%m-%d %H:%M:%S"
              ),
              "Action Type": action_type,
              "Instructor Name": instructor_name.strip(),
              "Customer / Candidate Name": customer_name.strip(),
              "Equipment IDs": eq_formatted,
              "Course Date": course_date.strftime("%Y-%m-%d"),
              "Assessment Date": assessment_date.strftime("%Y-%m-%d"),
              "Inspection Result": inspection_result,
              "Notes / Damage Details": notes,
          }])
          save_log(new_entry)
          st.success("Log submitted successfully!")
          st.rerun()
        else:
          st.error("Please provide Instructor Name and select Equipment IDs.")

  # ----------------------------------------------------
  # SECTION 2: LIVE DASHBOARD
  # ----------------------------------------------------
  elif nav_option == "Live Dashboard":
    st.header("Live Dashboard")

    # Build master live dashboard dataset
    dashboard_rows = []
    for aid in all_asset_ids:
      category = ""
      if not df_inv.empty and "Category" in df_inv.columns:
        cat_match = df_inv[df_inv["Asset ID"].astype(str) == aid]["Category"]
        if not cat_match.empty:
          category = cat_match.values[0]

      status = asset_status_map.get(aid, "🟢 AVAILABLE")

      dashboard_rows.append(
          {"Asset ID": aid, "Category": category, "Status": status}
      )

    df_dashboard = pd.DataFrame(dashboard_rows)

    # Calculate KPI summary totals
    total_avail = (
        (df_dashboard["Status"] == "🟢 AVAILABLE").sum()
        if not df_dashboard.empty
        else 0
    )
    total_in_use = (
        (df_dashboard["Status"] == "🔴 IN USE").sum()
        if not df_dashboard.empty
        else 0
    )
    total_maint = (
        (df_dashboard["Status"] == "⚠️ MAINTENANCE").sum()
        if not df_dashboard.empty
        else 0
    )

    # Display KPI Blocks at Top
    kpi1, kpi2, kpi3 = st.columns(3)
    with kpi1:
      st.metric(label="Total Available", value=total_avail)
    with kpi2:
      st.metric(label="Total In Use", value=total_in_use)
    with kpi3:
      st.metric(label="Maintenance Required", value=total_maint)

    st.markdown("### Asset Inventory Status")

    # Styling helper for conditional status formatting
    def highlight_status(val):
      if val == "🟢 AVAILABLE":
        return "background-color: #d4edda; color: #155724; font-weight: bold;"
      elif val == "🔴 IN USE":
        return "background-color: #f8d7da; color: #721c24; font-weight: bold;"
      elif val == "⚠️ MAINTENANCE":
        return "background-color: #fff3cd; color: #856404; font-weight: bold;"
      return ""

    if not df_dashboard.empty:
      styled_df = df_dashboard.style.applymap(
          highlight_status, subset=["Status"]
      )
      st.dataframe(styled_df, use_container_width=True)
    else:
      st.info("No asset inventory found.")
