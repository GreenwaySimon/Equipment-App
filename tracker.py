import datetime
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
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
    subject = (
        f"⚠️ Equipment Damage / Maintenance Reported -"
        f" {log_details['Equipment IDs']}"
    )

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
            st.warning(
                "Log saved successfully, but email notification could not be sent"
                " because SMTP credentials (SENDER_EMAIL and SENDER_PASSWORD) are not"
                " configured in Streamlit Secrets."
            )
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
        st.success("Maintenance report email sent successfully to Simon and Rob!")
    except Exception as e:
        st.warning(
            f"Log saved successfully, but failed to send email notification: {e}"
        )


def run():
    # Initialize session state for Dashboard view options if not set
    if "show_instructor" not in st.session_state:
        st.session_state.show_instructor = True
    if "show_customer" not in st.session_state:
        st.session_state.show_customer = True
    if "show_dates" not in st.session_state:
        st.session_state.show_dates = True

    # Responsive CSS Styling: Hides Admin Portal on mobile screens (<= 768px)
    st.markdown(
        """
    <style>
    div.stButton > button {
        font-size: 16px;
        padding: 0.5rem 1rem;
        min-height: 40px;
    }
    .metric-card {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 15px;
        text-align: center;
        border: 1px solid #dee2e6;
    }

    /* Mobile & Tablet Optimizations & Admin Portal Hiding */
    @media (max-width: 768px) {
        div.stRadio div[role="radiogroup"] label:nth-child(3) {
            display: none !important;
        }
        div.stButton > button {
            font-size: 18px !important;
            padding: 0.85rem 1rem !important;
            width: 100% !important;
            min-height: 55px !important;
        }
        .stSelectbox, .stMultiSelect, .stTextInput, .stDateInput {
            font-size: 16px !important;
        }
        h1 { font-size: 24px !important; }
        h2 { font-size: 20px !important; }
        h3 { font-size: 18px !important; }
        .stDataFrame { font-size: 13px !important; }
    }
    </style>
    """,
        unsafe_allow_html=True,
    )

    df_inv, df_logs = load_inventory_and_logs()

    # Left Sidebar Navigation Menu
    st.sidebar.title("Menu")
    nav_option = st.sidebar.radio(
        "Navigation",
        ["Equipment Hire Log", "Live Dashboard", "Admin Portal"],
    )

    st.markdown("---")

    today = datetime.date.today()

    # Process equipment statuses, customer mappings, and dynamic dates based on history
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

    # Re-evaluate statuses for assets whose assessment date has passed today
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

    # Build map of active customers and the equipment they currently hold
    for aid, info in asset_data_map.items():
        if info["status"] == "🔴 IN USE" and info["customer"]:
            c_name = info["customer"]
            if c_name not in customer_to_equipment_map:
                customer_to_equipment_map[c_name] = []
            if aid not in customer_to_equipment_map[c_name]:
                customer_to_equipment_map[c_name].append(aid)

    # Determine all assets from master inventory
    all_asset_ids = []
    if not df_inv.empty and "Asset ID" in df_inv.columns:
        all_asset_ids = df_inv["Asset ID"].dropna().astype(str).tolist()

    # ----------------------------------------------------
    # SECTION 1: EQUIPMENT HIRE LOG
    # ----------------------------------------------------
    if nav_option == "Equipment Hire Log":
        st.header("Equipment Hire Log")
        st.write("Fill out the form below to record an entry.")

        action_type = st.selectbox(
            "Action Type",
            [
                "Check Out",
                "Return Equipment",
                "Report Damage / Maintenance",
            ],
            key="input_action_type",
        )

        instructor_name = st.text_input("Instructor Name", key="input_instructor")

        if action_type == "Check Out":
            customer_name = st.text_input(
                "Customer / Candidate Name", key="input_customer_text"
            )
            equipment_ids = st.multiselect(
                "Equipment IDs (Available Equipment)",
                options=all_asset_ids if all_asset_ids else [],
                key="input_eq_checkout",
            )

        elif action_type == "Return Equipment":
            active_customers = list(customer_to_equipment_map.keys())
            if active_customers:
                customer_name = st.selectbox(
                    "Customer / Candidate Name",
                    options=active_customers,
                    key="input_customer_select",
                )
                allowed_equipment = customer_to_equipment_map.get(customer_name, [])
                equipment_ids = st.multiselect(
                    "Equipment IDs to Return",
                    options=allowed_equipment,
                    default=allowed_equipment,
                    key="input_eq_return",
                )
            else:
                st.info("No customers currently have equipment checked out to return.")
                customer_name = ""
                equipment_ids = []

        else:  # Report Damage / Maintenance
            customer_name = st.text_input(
                "Customer / Candidate Name (Optional)", key="input_customer_maint"
            )
            equipment_ids = st.multiselect(
                "Equipment IDs (Required)",
                options=all_asset_ids if all_asset_ids else [],
                key="input_eq_maint",
            )

        col_cb, col_dt = st.columns([1, 1])
        with col_cb:
            include_assessment = st.checkbox(
                "Include Assessment Date (Optional)",
                value=False,
                key="input_include_assessment",
            )

        assessment_date_str = ""
        if include_assessment:
            with col_dt:
                assessment_date_val = st.date_input(
                    "Assessment Date",
                    value=datetime.date.today(),
                    format="DD/MM/YYYY",
                    key="input_assessment_date",
                )
                assessment_date_str = assessment_date_val.strftime("%d/%m/%Y")

        inspection_result = st.selectbox(
            "Inspection Result",
            ["Not Inspected", "Pass", "Fail / Maintenance Required"],
            index=0,
            key="input_inspection",
        )

        notes = st.text_area(
            "Notes / Damage Details"
            + (
                " (Required for Maintenance)"
                if action_type == "Report Damage / Maintenance"
                else ""
            ),
            key="input_notes",
        )

        submitted = st.button("Submit Entry", key="submit_log_btn")

        has_reservation_warning = False
        if submitted and action_type == "Check Out" and instructor_name:
            for eq_id in equipment_ids:
                eq_info = asset_data_map.get(eq_id, {})
                if "RESERVED" in eq_info.get("status", ""):
                    ass_dt = eq_info.get("assessment_date", "")
                    st.warning(
                        f"Warning: Equipment {eq_id} is reserved for assessment on"
                        f" {ass_dt} and cannot be booked out!"
                    )
                    has_reservation_warning = True

        if submitted:
            if has_reservation_warning:
                st.error(
                    "Submission blocked: One or more selected items are reserved for"
                    " assessment."
                )
            elif action_type == "Report Damage / Maintenance" and (
                not equipment_ids or not notes.strip()
            ):
                st.error(
                    "For 'Report Damage / Maintenance', you must select Equipment IDs"
                    " and provide Notes / Damage Details."
                )
            elif not instructor_name:
                st.error("Please provide the Instructor Name.")
            elif not equipment_ids:
                st.error("Please select Equipment IDs.")
            else:
                eq_formatted = ", ".join(equipment_ids)
                new_entry_dict = {
                    "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Action Type": action_type,
                    "Instructor Name": instructor_name.strip(),
                    "Customer / Candidate Name": (
                        str(customer_name).strip() if customer_name else
