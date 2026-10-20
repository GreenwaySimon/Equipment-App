import datetime
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import streamlit as st

EXCEL_FILE = "Chainsaw_Equipment_Tracker.xlsx"
LOGO_FILE = "greenway_training_logo.jpg"

st.set_page_config(
    page_title="Greenway Training Inventory Hub",
    page_icon="🪵",
    layout="wide",
)


def load_inventory_and_logs():
    try:
        df_inv = pd.read_excel(EXCEL_FILE, sheet_name="Inventory")
    except Exception:
        df_inv = pd.DataFrame(columns=["Asset ID", "Category"])

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


def save_inventory(df_inv):
    _, df_logs = load_inventory_and_logs()
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_inv.to_excel(writer, sheet_name="Inventory", index=False)
        df_logs.to_excel(writer, sheet_name="Sheet1", index=False)


def send_maintenance_email(log_details):
    recipients = ["simon@greenwaytraining.co.uk", "rob@greenwaytraining.co.uk"]
    subject = f"⚠️ Equipment Damage / Maintenance Reported - {log_details['Equipment IDs']}"

    body = f"""
An equipment damage or maintenance report has been submitted:

- Timestamp: {log_details['Timestamp']}
- Action Type: {log_details['Action Type']}
- Instructor Name: {log_details['Instructor Name']}
- Customer / Candidate Name: {log_details['Customer / Candidate Name'] or 'N/A'}
- Equipment IDs: {log_details['Equipment IDs']}
- Assessment Date: {log_details['Assessment Date'] or 'N/A'}
- Inspection Result: {log_details['Inspection Result']}
- Notes / Damage Details: {log_details['Notes / Damage Details']}
"""

    try:
        smtp_server = st.secrets.get("SMTP_SERVER", "smtp.gmail.com")
        smtp_port = int(st.secrets.get("SMTP_PORT", 587))
        sender_email = st.secrets.get("SENDER_EMAIL", "")
        sender_password = st.secrets.get("SENDER_PASSWORD", "")

        if not sender_email or not sender_password:
            return

        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = ", ".join(recipients)
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipients, msg.as_string())
        server.quit()
    except Exception:
        pass


def main():
    if "show_instructor" not in st.session_state:
        st.session_state.show_instructor = True
    if "show_customer" not in st.session_state:
        st.session_state.show_customer = True
    if "show_dates" not in st.session_state:
        st.session_state.show_dates = True

    # Custom mobile, UI, and High-Contrast Sidebar Styling
    st.markdown(
        """
    <style>
    [data-testid="stHeader"] {
        background: transparent;
    }
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 2rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }
    
    div.stButton > button {
        font-size: 18px !important;
        padding: 0.75rem 1rem;
        min-height: 50px;
        width: 100% !important;
        border-radius: 8px;
    }
    .metric-card {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 15px;
        text-align: center;
        border: 1px solid #dee2e6;
    }
    @media print {
        body { background: white; color: black; }
        .stSidebar, nav, header, button { display: none !important; }
    }
    @media (max-width: 768px) {
        .stDataFrame { 
            font-size: 14px !important; 
        }
        .stSelectbox, .stMultiSelect, .stTextInput, .stDateInput, .stTextArea {
            font-size: 16px !important;
        }
        h1 { font-size: 24px !important; }
        h2 { font-size: 20px !important; }
        h3 { font-size: 18px !important; }
    }
    </style>
    """,
        unsafe_allow_html=True,
    )

    df_inv, df_logs = load_inventory_and_logs()

    # Sidebar Navigation
    st.sidebar.title("Menu")
    nav_options = ["Equipment Hire Log", "Live Dashboard", "Admin Portal"]

    if st.session_state.get("redirect_to_dashboard", False):
        st.session_state.redirect_to_dashboard = False
        st.session_state.nav_choice = "Live Dashboard"

    default_ix = 0
    if "nav_choice" in st.session_state and st.session_state.nav_choice in nav_options:
        default_ix = nav_options.index(st.session_state.nav_choice)

    nav_option = st.sidebar.radio("Navigation", nav_options, index=default_ix, key="nav_choice")

    today = datetime.date.today()

    asset_data_map = {}
    customer_to_equipment_map = {}

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
            inst = str(row.get("Instructor Name", ""))
            cust = str(row.get("Customer / Candidate Name", ""))
            ts = str(row.get("Timestamp", ""))
            date_val = ts[:10] if ts and ts != "nan" else ""
            ass_dt_log = str(row.get("Assessment Date", ""))

            for eq_id in eq_list:
                if eq_id not in asset_data_map:
                    asset_data_map[eq_id] = {
                        "status": "🟢 AVAILABLE",
                        "instructor": "",
                        "customer": "",
                        "checkout_date": "",
                        "return_date": "",
                        "assessment_date": "",
                        "reported_date": "",
                        "repaired_date": "",
                    }

                d = asset_data_map[eq_id]
                d["instructor"] = inst
                d["customer"] = cust
                if ass_dt_log and ass_dt_log != "nan" and ass_dt_log.strip():
                    d["assessment_date"] = ass_dt_log.strip()

                if action in ["Check Out", "Scheduled Kit Hire"]:
                    d["status"] = "🔴 IN USE"
                    d["checkout_date"] = date_val
                    d["return_date"] = "Pending"
                elif action == "Return Equipment":
                    insp = str(row.get("Inspection Result", ""))
                    d["return_date"] = date_val
                    d["customer"] = ""
                    if "Fail" in insp or "Maintenance" in insp:
                        d["status"] = "⚠️ MAINTENANCE"
                        d["reported_date"] = date_val
                        d["repaired_date"] = "Pending"
                    else:
                        is_reserved = False
                        if d["assessment_date"]:
                            try:
                                ass_parsed = datetime.datetime.strptime(
                                    d["assessment_date"], "%d/%m/%Y"
                                ).date()
                                if today <= ass_parsed:
                                    is_reserved = True
                            except Exception:
                                pass

                        if is_reserved:
                            d["status"] = f"📅 RESERVED ({d['assessment_date']})"
                        else:
                            d["status"] = "🟢 AVAILABLE"
                elif action == "Report Damage / Maintenance":
                    d["status"] = "⚠️ MAINTENANCE"
                    d["reported_date"] = date_val
                    d["repaired_date"] = "Pending"

    for aid, info in asset_data_map.items():
        if "RESERVED" in info["status"]:
            if info["assessment_date"]:
                try:
                    ass_parsed = datetime.datetime.strptime(
                        info["assessment_date"], "%d/%m/%Y"
                    ).date()
                    if today > ass_parsed:
                        info["status"] = "🟢 AVAILABLE"
                except Exception:
                    pass

    for aid, info in asset_data_map.items():
        if info["status"] == "🔴 IN USE" and info["customer"]:
            c_name = info["customer"]
            if c_name not in customer_to_equipment_map:
                customer_to_equipment_map[c_name] = []
            if aid not in customer_to_equipment_map[c_name]:
                customer_to_equipment_map[c_name].append(aid)

    all_asset_ids = []
    if not df_inv.empty and "Asset ID" in df_inv.columns:
        all_asset_ids = df_inv["Asset ID"].dropna().astype(str).tolist()

    def render_header(title_text):
        if os.path.exists(LOGO_FILE):
            col_logo, col_title = st.columns([1, 4])
            with col_logo:
                st.image(LOGO_FILE, width=130)
            with col_title:
                st.title(title_text)
        else:
            st.title(title_text)

    if nav_option == "Equipment Hire Log":
        render_header("Equipment Hire Log")
        st.write("Select the operation tab below to record an entry.")

        log_tab_checkout, log_tab_return, log_tab_damage = st.tabs([
            "📤 Check Out",
            "📥 Return Equipment",
            "⚠️ Report Damage",
        ])

        with log_tab_checkout:
            st.subheader("Check Out Equipment")
            with st.form("checkout_form_sub"):
                co_instructor = st.text_input("Instructor Name", key="co_inst")
                co_customer = st.text_input("Customer / Candidate Name", key="co_cust")
                co_equipment = st.multiselect(
                    "Equipment IDs (Available Equipment)",
                    options=all_asset_ids if all_asset_ids else [],
                    key="co_eq",
                )

                co_col_cb, co_col_dt = st.columns([1, 1])
                with co_col_cb:
                    co_include_ass = st.checkbox("Include Assessment Date (Optional)", value=False, key="co_inc_ass")
                
                co_ass_date_str = ""
                if co_include_ass:
                    with co_col_dt:
                        co_ass_val = st.date_input("Assessment Date", value=datetime.date.today(), format="DD/MM/YYYY", key="co_ass_val")
                        co_ass_date_str = co_ass_val.strftime("%d/%m/%Y")

                co_notes = st.text_area("Notes / Booking Details", key="co_notes")
                co_submitted = st.form_submit_button("Submit Check Out")

                has_reservation_warning = False
                if co_submitted and co_instructor:
                    for eq_id in co_equipment:
                        eq_info = asset_data_map.get(eq_id, {})
                        if "RESERVED" in eq_info.get("status", ""):
                            ass_dt = eq_info.get("assessment_date", "")
                            st.warning(f"Warning: Equipment {eq_id} is reserved for assessment on {ass_dt} and cannot be booked out!")
                            has_reservation_warning = True

                if co_submitted:
                    if has_reservation_warning:
                        st.error("Submission blocked: One or more selected items are reserved for assessment.")
                    elif not co_instructor:
                        st.error("Please provide the Instructor Name.")
                    elif not co_customer:
                        st.error("Please provide the Customer / Candidate Name.")
                    elif not co_equipment:
                        st.error("Please select Equipment IDs.")
                    else:
                        eq_formatted = ", ".join(co_equipment)
                        new_entry_dict = {
                            "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Action Type": "Check Out",
                            "Instructor Name": co_instructor.strip(),
                            "Customer / Candidate Name": co_customer.strip(),
                            "Equipment IDs": eq_formatted,
                            "Assessment Date": co_ass_date_str,
                            "Inspection Result": "Not Inspected",
                            "Notes / Damage Details": co_notes.strip(),
                        }
                        save_log(pd.DataFrame([new_entry_dict]))
                        st.success("✅ Success! Your equipment has been successfully logged.")
                        st.balloons()
                        st.session_state.redirect_to_dashboard = True
                        st.rerun()

        with log_tab_return:
            st.subheader("Return Equipment")
            active_customers = list(customer_to_equipment_map.keys())
            
            if active_customers:
                with st.form("return_form_sub"):
                    ret_instructor = st.text_input("Instructor Name", key="ret_inst")
                    ret_customer = st.selectbox("Customer / Candidate Name", options=active_customers, key="ret_cust_sel")
                    
                    default_checked_out = customer_to_equipment_map.get(ret_customer, [])
                    ret_equipment = st.multiselect(
                        "Equipment IDs to Return (Pre-selected automatically)",
                        options=default_checked_out,
                        default=default_checked_out,
                        key="ret_eq_sel"
                    )

                    ret_inspection = st.selectbox(
                        "Inspection Result",
                        ["Pass", "Fail / Maintenance Required", "Not Inspected"],
                        index=0,
                        key="ret_insp"
                    )
                    ret_notes = st.text_area("Return Notes / Inspection Details", key="ret_notes")
                    ret_submitted = st.form_submit_button("Submit Return")

                    if ret_submitted:
                        if not ret_instructor:
                            st.error("Please provide the Instructor Name.")
                        elif not ret_equipment:
                            st.error("Please select Equipment IDs to return.")
                        else:
                            eq_formatted = ", ".join(ret_equipment)
                            new_entry_dict = {
                                "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "Action Type": "Return Equipment",
                                "Instructor Name": ret_instructor.strip(),
                                "Customer / Candidate Name": ret_customer.strip(),
                                "Equipment IDs": eq_formatted,
                                "Assessment Date": "",
                                "Inspection Result": ret_inspection,
                                "Notes / Damage Details": ret_notes.strip(),
                            }
                            save_log(pd.DataFrame([new_entry_dict]))

                            if "Fail" in ret_inspection or "Maintenance" in ret_inspection:
                                send_maintenance_email(new_entry_dict)

                            st.success("✅ Success! Your equipment has been successfully logged.")
                            st.balloons()
                            st.session_state.redirect_to_dashboard = True
                            st.rerun()
            else:
                st.info("No customers currently have equipment checked out to return.")

        with log_tab_damage:
            st.subheader("Report Damage or Maintenance")
            with st.form("damage_form_sub"):
                dmg_instructor = st.text_input("Instructor Name", key="dmg_inst")
                dmg_customer = st.text_input("Customer / Candidate Name (Optional)", key="dmg_cust")
                dmg_equipment = st.multiselect(
                    "Equipment IDs (Required)",
                    options=all_asset_ids if all_asset_ids else [],
                    key="dmg_eq"
                )
                dmg_
