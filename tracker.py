import datetime
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import streamlit as st

EXCEL_FILE = "Chainsaw_Equipment_Tracker.xlsx"
LOGO_FILE = "greenway_training_logo.jpg"


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
            st.warning(
                "Log saved successfully, but email notification could not be sent "
                "because SMTP credentials are not configured in Streamlit Secrets."
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
        st.warning(f"Log saved successfully, but failed to send email notification: {e}")


def run():
    if "show_instructor" not in st.session_state:
        st.session_state.show_instructor = True
    if "show_customer" not in st.session_state:
        st.session_state.show_customer = True
    if "show_dates" not in st.session_state:
        st.session_state.show_dates = True

    if "welcome_dismissed" not in st.session_state:
        ua = ""
        try:
            ua = st.context.headers.get("User-Agent", "").lower()
        except Exception:
            pass
        is_mobile_device = any(k in ua for k in ["mobi", "android", "iphone", "ipad", "ipod"])
        st.session_state.show_welcome = is_mobile_device

    st.markdown(
        """
    <style>
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 2rem !important;
    }
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
    @media print {
        body { background: white; color: black; }
        .stSidebar, nav, header, button { display: none !important; }
        .printable-card {
            border: 2px solid #333 !important;
            padding: 20px !important;
            margin: 20px 0 !important;
            border-radius: 8px !important;
        }
    }
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
        h1 { font-size: 22px !important; }
        h2 { font-size: 20px !important; }
        h3 { font-size: 18px !important; }
        .stDataFrame { font-size: 13px !important; }
    }
    </style>
    """,
        unsafe_allow_html=True,
    )

    if st.session_state.get("show_welcome", False):
        if os.path.exists(LOGO_FILE):
            st.image(LOGO_FILE, width=220)
        
        st.title("Greenway Training Inventory Hub")
        st.write("Welcome! Streamlined equipment tracking and safety logs for field instructors.")
        
        st.info("📱 **Quick Tip:** Add this app directly to your mobile phone's home screen so you can access it instantly like a native app without typing web addresses.")
        
        with st.expander("📖 Step-by-Step Instructions: How to Add to Mobile Home Screen", expanded=True):
            st.markdown("""
            **For iPhone / iPad (Safari):**
            1. Tap the **Share** button at the bottom of your screen (square with an arrow pointing upward).
            2. Scroll down and tap **"Add to Home Screen"**.
            3. Tap **Add** in the top-right corner.

            **For Android (Chrome):**
            1. Tap the **three vertical dots** in the top-right corner of Chrome.
            2. Tap **"Add to Home screen"** or **"Install app"**.
            3. Tap **Install** or **Add** to confirm.
            """)
        
        if st.button("🚀 Continue to Inventory App", use_container_width=True):
            st.session_state.show_welcome = False
            st.session_state.welcome_dismissed = True
            st.rerun()
            
        return

    df_inv, df_logs = load_inventory_and_logs()

    st.sidebar.title("Menu")
    nav_option = st.sidebar.radio(
        "Navigation",
        ["Equipment Hire Log", "Live Dashboard", "Admin Portal"],
    )

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

        else:
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
                        f"Warning: Equipment {eq_id} is reserved for assessment on {ass_dt} and cannot be booked out!"
                    )
                    has_reservation_warning = True

        if submitted:
            if has_reservation_warning:
                st.error("Submission blocked: One or more selected items are reserved for assessment.")
            elif action_type == "Report Damage / Maintenance" and (not equipment_ids or not notes.strip()):
                st.error("For 'Report Damage / Maintenance', you must select Equipment IDs and provide Notes / Damage Details.")
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
                    "Customer / Candidate Name": str(customer_name).strip() if customer_name else "",
                    "Equipment IDs": eq_formatted,
                    "Assessment Date": assessment_date_str,
                    "Inspection Result": inspection_result,
                    "Notes / Damage Details": notes.strip(),
                }
                new_entry = pd.DataFrame([new_entry_dict])
                save_log(new_entry)

                if action_type == "Report Damage / Maintenance":
                    send_maintenance_email(new_entry_dict)

                st.success("Log submitted successfully!")
                st.rerun()

    elif nav_option == "Live Dashboard":
        render_header("Live Dashboard")

        dashboard_rows = []
        for aid in all_asset_ids:
            info = asset_data_map.get(
                aid,
                {
                    "status": "🟢 AVAILABLE",
                    "instructor": "",
                    "customer": "",
                    "checkout_date": "",
                    "return_date": "",
                    "assessment_date": "",
                    "reported_date": "",
                    "repaired_date": "",
                },
            )

            status = info["status"]
            instructor = info["instructor"]
            customer = info["customer"]

            if status == "⚠️ MAINTENANCE":
                rep_dt = info["reported_date"] or "N/A"
                repa_dt = info["repaired_date"] or "Pending"
                date_info = f"Reported: {rep_dt} | Repaired: {repa_dt}"
            elif "RESERVED" in status:
                ass_dt = info["assessment_date"] or "N/A"
                date_info = f"Assessment Date: {ass_dt}"
            else:
                co_dt = info["checkout_date"] or "N/A"
                ret_dt = info["return_date"] or ("Pending" if status == "🔴 IN USE" else "N/A")
                date_info = f"Check Out: {co_dt} | Return: {ret_dt}"

            row_dict = {"Asset ID": aid, "Status": status}
            if st.session_state.show_instructor:
                row_dict["Instructor Name"] = instructor
            if st.session_state.show_customer:
                row_dict["Customer / Candidate Name"] = customer
            if st.session_state.show_dates:
                row_dict["Dates"] = date_info

            dashboard_rows.append(row_dict)

        df_dashboard = pd.DataFrame(dashboard_rows)

        total_avail = (df_dashboard["Status"] == "🟢 AVAILABLE").sum() if not df_dashboard.empty else 0
        total_in_use = (df_dashboard["Status"] == "🔴 IN USE").sum() if not df_dashboard.empty else 0
        total_reserved = df_dashboard["Status"].str.contains("RESERVED").sum() if not df_dashboard.empty else 0
        total_maint = (df_dashboard["Status"] == "⚠️ MAINTENANCE").sum() if not df_dashboard.empty else 0

        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        with kpi1:
            st.metric(label="Total Available", value=total_avail)
        with kpi2:
            st.metric(label="Total In Use", value=total_in_use)
        with kpi3:
            st.metric(label="Reserved for Assessment", value=total_reserved)
        with kpi4:
            st.metric(label="Maintenance Required", value=total_maint)

        st.markdown("### Asset Inventory Status")

        def highlight_status(val):
            if val == "🟢 AVAILABLE":
                return "background-color: #d4edda; color: #155724; font-weight: bold;"
            elif val == "🔴 IN USE":
                return "background-color: #f8d7da; color: #721c24; font-weight: bold;"
            elif val == "⚠️ MAINTENANCE":
                return "background-color: #fff3cd; color: #856404; font-weight: bold;"
            elif "RESERVED" in val:
                return "background-color: #cce5ff; color: #004085; font-weight: bold;"
            return ""

        if not df_dashboard.empty:
            if hasattr(df_dashboard.style, "map"):
                styled_df = df_dashboard.style.map(highlight_status, subset=["Status"])
            else:
                styled_df = df_dashboard.style.applymap(highlight_status, subset=["Status"])
            st.dataframe(styled_df, use_container_width=True)
        else:
            st.info("No asset inventory found.")

    elif nav_option == "Admin Portal":
        render_header("Admin Management Portal")
        st.write("Administrative controls for managing inventory, assessment dates, dashboard views, scheduled kit hires, maintenance history, and mobile guides.")

        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "Manage Equipment",
            "Assessment Dates",
            "Dashboard Views",
            "Scheduled Kit Hire",
            "Maintenance History",
            "Mobile App Guide",
        ])

        with tab1:
            st.subheader("Add or Remove Equipment")
            col_add, col_rem = st.columns(2)

            with col_add:
                st.markdown("#### Add New Equipment")
                with st.form("add_equip_form"):
                    new_asset_id = st.text_input("New Asset ID")
                    new_category = st.text_input("Category (e.g., Chainsaw, PPE)")
                    add_submitted = st.form_submit_button("Add Equipment")

                    if add_submitted:
                        if new_asset_id.strip():
                            if new_asset_id.strip() in all_asset_ids:
                                st.error("Asset ID already exists!")
                            else:
                                new_row = pd.DataFrame([{"Asset ID": new_asset_id.strip(), "Category": new_category}])
                                df_inv = pd.concat([df_inv, new_row], ignore_index=True)
                                save_inventory(df_inv)
                                st.success(f"Successfully added {new_asset_id.strip()}!")
                                st.rerun()
                        else:
                            st.error("Please enter a valid Asset ID.")

            with col_rem:
                st.markdown("#### Remove Equipment")
                with st.form("rem_equip_form"):
                    rem_asset_id = st.selectbox("Select Asset ID to Remove", options=all_asset_ids)
                    rem_submitted = st.form_submit_button("Remove Equipment")

                    if rem_submitted:
                        if rem_asset_id:
                            df_inv = df_inv[df_inv["Asset ID"].astype(str) != str(rem_asset_id)]
                            save_inventory(df_inv)
                            st.success(f"Successfully removed {rem_asset_id}!")
                            st.rerun()

        with tab2:
            st.subheader("Manually Assign Assessment Date")
            with st.form("admin_assessment_form"):
                target_asset = st.selectbox("Select Asset ID", options=all_asset_ids)
                admin_ass_date = st.date_input("Assessment Date", value=datetime.date.today(), format="DD/MM/YYYY")
                admin_ass_notes = st.text_area("Notes / Reason")
                admin_ass_submitted = st.form_submit_button("Save Assessment Date")

                if admin_ass_submitted:
                    if target_asset:
                        formatted_date = admin_ass_date.strftime("%d/%m/%Y")
                        new_entry = pd.DataFrame([{
                            "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Action Type": "Admin Assessment Setup",
                            "Instructor Name": "Admin",
                            "Customer / Candidate Name": "Scheduled Assessment",
                            "Equipment IDs": str(target_asset),
                            "Assessment Date": formatted_date,
                            "Inspection Result": "Not Inspected",
                            "Notes / Damage Details": admin_ass_notes,
                        }])
                        save_log(new_entry)
                        st.success(f"Assessment date {formatted_date} set for asset {target_asset}!")
                        st.rerun()

        with tab3:
            st.subheader("Live Dashboard Display Settings")
            st.write("Select which information columns you want displayed on the Live Dashboard:")

            st.session_state.show_instructor = st.checkbox("Show Instructor Name", value=st.session_state.show_instructor)
            st.session_state.show_customer = st.checkbox("Show Customer / Candidate Name", value=st.session_state.show_customer)
            st.session_state.show_dates = st.checkbox("Show Dates (Check Out / Return / Assessment / Maintenance)", value=st.session_state.show_dates)
            st.success("Dashboard view preferences updated successfully!")

        with tab4:
            st.subheader("Schedule Predetermined Kit Hire")
            st.write("Book a kit hire in advance. The system will automatically apply the 'IN USE' status on the selected start date.")

            with st.form("scheduled_hire_form"):
                sched_instructor = st.text_input("Instructor Name")
                sched_customer = st.text_input("Customer / Candidate Name")
                sched_equipment = st.multiselect("Equipment IDs to Schedule", options=all_asset_ids)
                sched_date = st.date_input("Kit Hire Start Date", value=datetime.date.today(), format="DD/MM/YYYY")
                sched_notes = st.text_area("Booking Notes")
                sched_submitted = st.form_submit_button("Schedule Kit Hire")

                if sched_submitted:
                    if sched_instructor and sched_equipment and sched_customer:
                        eq_formatted = ", ".join(sched_equipment)
                        new_entry = pd.DataFrame([{
                            "Timestamp": sched_date.strftime("%Y-%m-%d 00:00:00"),
                            "Action Type": "Scheduled Kit Hire",
                            "Instructor Name": sched_instructor.strip(),
                            "Customer / Candidate Name": sched_customer.strip(),
                            "Equipment IDs": eq_formatted,
                            "Assessment Date": "",
                            "Inspection Result": "Not Inspected",
                            "Notes / Damage Details": sched_notes,
                        }])
                        save_log(new_entry)
                        st.success("Predetermined kit hire scheduled successfully!")
                        st.rerun()
                    else:
                        st.error("Please fill in Instructor Name, Customer Name, and select Equipment IDs.")

        with tab5:
            st.subheader("Damage & Maintenance History")
            st.write("View the history of damage and maintenance reports logged for equipment.")

            selected_asset_filter = st.selectbox(
                "Filter by Asset ID",
                options=["All Equipment"] + all_asset_ids,
                key="maint_history_filter"
            )

            maint_logs = []
            if not df_logs.empty:
                for _, row in df_logs.iterrows():
                    action = str(row.get("Action Type", ""))
                    insp = str(row.get("Inspection Result", ""))
                    notes = str(row.get("Notes / Damage Details", ""))
                    if (
                        action == "Report Damage / Maintenance"
                        or ("Fail" in insp or "Maintenance" in insp)
                        or (action == "Return Equipment" and notes.strip() and notes != "nan")
                    ):
                        maint_logs.append(row)

            df_maint_history = pd.DataFrame(maint_logs) if maint_logs else pd.DataFrame(columns=df_logs.columns)

            if not df_maint_history.empty:
                if selected_asset_filter != "All Equipment":
                    filtered_rows = []
                    for _, row in df_maint_history.iterrows():
                        eq_str = str(row.get("Equipment IDs", ""))
                        eq_list = [e.strip() for e in eq_str.split(",") if e.strip()]
                        if selected_asset_filter in eq_list:
                            filtered_rows.append(row)
                    df_display = pd.DataFrame(filtered_rows) if filtered_rows else pd.DataFrame(columns=df_maint_history.columns)
                else:
                    df_display = df_maint_history

                if not df_display.empty:
                    st.dataframe(df_display, use_container_width=True)
                else:
                    st.info(f"No damage or maintenance records found
