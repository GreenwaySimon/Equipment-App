import streamlit as st
import pandas as pd
import tracker
import json
from datetime import datetime

# Page Configuration (Wide layout, Greenway logo as favicon)
st.set_page_config(
    page_title="Greenway Training - Chainsaw & PPE Tracker",
    page_icon="greenway_training_logo.jpg",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Greenway Theme Styling (Green and White) with adjusted top padding to prevent clipping
st.markdown("""
    <style>
    /* Main container styling with comfortable top breathing room */
    .block-container { padding-top: 2rem; padding-bottom: 0rem; padding-left: 2rem; padding-right: 2rem; background-color: #ffffff; }
    
    /* Headers */
    h1, h2, h3 { color: #145A32 !important; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; }
    h1 { font-size: 1.8rem; margin: 0rem; line-height: 1.2; }
    h3 { font-size: 1.2rem; margin-top: 0rem; }
    
    /* Metric cards styling */
    .stMetric { background-color: #f4f9f4; padding: 10px; border-radius: 6px; border-left: 4px solid #145A32; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
    .stMetric label { color: #145A32 !important; font-weight: 600; }
    
    /* Buttons */
    .stButton>button { background-color: #145A32; color: white; border-radius: 6px; border: none; font-weight: 600; }
    .stButton>button:hover { background-color: #0e421d; color: white; }
    
    /* Sidebar */
    [data-testid="stSidebar"] { background-color: #f8faf8; border-right: 1px solid #e2e8e2; }
    
    /* Tabs and Inputs */
    .stTextInput>div>div>input, .stSelectbox>div>div>select, .stTextArea>div>div>textarea { border-color: #c8d8c8; }
    .stTextInput>div>div>input:focus, .stSelectbox>div>div>select:focus { border-color: #145A32; box-shadow: 0 0 0 1px #145A32; }
    </style>
""", unsafe_allow_html=True)

# Header layout with Greenway Logo and Title vertically centered to prevent clipping
col_logo, col_title = st.columns([1, 7], vertical_alignment="center")
with col_logo:
    try:
        st.image("greenway_training_logo.jpg", width=110)
    except:
        st.markdown("### 🟢")
with col_title:
    st.markdown("<h1>Chainsaw & PPE Live Inventory Hub</h1>", unsafe_allow_html=True)

# Load current data
df_inv = pd.read_excel(tracker.EXCEL_PATH, sheet_name='Inventory')
df_resp = pd.read_excel(tracker.EXCEL_PATH, sheet_name='Form Responses 1')
allocations = tracker.get_current_allocations(df_resp, df_inv)

data_rows = []
for aid, state in allocations.items():
    cat_match = df_inv.loc[df_inv['Asset ID'].str.strip() == aid, 'Category']
    cat = cat_match.values[0] if not cat_match.empty else 'General'
    data_rows.append({
        'Asset ID': aid,
        'Category': cat,
        'Status': state['Status'],
        'Instructor': state['Instructor'],
        'Customer': state['Customer'],
        'Assessment Date': state['Assessment Date'],
        'Last Inspection': state['Last Inspection'],
        'Notes': state['Notes']
    })
df_dashboard = pd.DataFrame(data_rows)

# Sidebar Navigation
app_mode = st.sidebar.radio("Navigation:", ["📊 Live Dashboard", "📱 Log Equipment Action"])

if app_mode == "📊 Live Dashboard":
    st.markdown("💡 *Tip: On desktop, you can directly click and edit the **Assessment Date** column in the table below.*")
    
    # Metrics Row
    total_items = len(df_dashboard)
    available_items = len(df_dashboard[df_dashboard['Status'].str.contains('AVAILABLE')])
    in_use_items = len(df_dashboard[df_dashboard['Status'].str.contains('IN USE')])
    reserved_items = len(df_dashboard[df_dashboard['Status'].str.contains('RESERVED')])
    maint_items = len(df_dashboard[df_dashboard['Status'].str.contains('MAINTENANCE')])
    
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Total", total_items)
    col2.metric("Available", available_items)
    col3.metric("In Use", in_use_items)
    col4.metric("Reserved", reserved_items)
    col5.metric("Maintenance", maint_items)
    
    # Compact Filter Row
    fcol1, fcol2 = st.columns([2, 1])
    with fcol1:
        search_query = st.text_input("🔍 Search Asset / Customer / Instructor:", "", placeholder="Type to filter...")
    with fcol2:
        status_filter = st.selectbox("Filter Status:", ["All", "🟢 AVAILABLE", "🔴 IN USE", "🟡 RESERVED FOR ASSESSMENT", "⚠️ MAINTENANCE"])
    
    filtered_df = df_dashboard.copy()
    if search_query:
        filtered_df = filtered_df[filtered_df.apply(lambda row: row.astype(str).str.contains(search_query, case=False).any(), axis=1)]
    if status_filter != "All":
        filtered_df = filtered_df[filtered_df['Status'] == status_filter]
        
    # Editable DataFrame (Desktop-friendly: allows editing Assessment Date directly)
    edited_df = st.data_editor(
        filtered_df,
        use_container_width=True,
        hide_index=True,
        height=500,
        disabled=["Asset ID", "Category", "Status", "Instructor", "Customer", "Last Inspection", "Notes"],
        key="dashboard_editor"
    )
    
    # Check if any Assessment Date was modified in the editor
    if not edited_df.equals(df_dashboard):
        for idx, row in edited_df.iterrows():
            orig_row = df_dashboard.iloc[idx]
            if row['Assessment Date'] != orig_row['Assessment Date']:
                aid = row['Asset ID']
                new_date = row['Assessment Date']
                tracker.update_assessment_date(aid, new_date)
                st.rerun()

elif app_mode == "📱 Log Equipment Action":
    st.header("Instructor Action Portal")
    
    # Collapsible instruction box for first-time mobile users
    with st.expander("📱 First time here? How to add this app to your phone's Home Screen"):
        st.markdown("""
        **For iPhone (Safari):**
        1. Tap the **Share** button (the square icon with an arrow pointing upward) at the bottom of your screen.
        2. Scroll down the menu and tap **Add to Home Screen**.
        3. Tap **Add** in the top-right corner.
        
        **For Android (Chrome):**
        1. Tap the **three vertical dots** in the top-right corner of your browser.
        2. Tap **Add to Home screen** (or **Install app**).
        3. Confirm by tapping **Add**.
        """)
    
    with st.form("action_form"):
        action_type = st.selectbox("Action Type", [
            "Check Out for Course",
            "Return / Check In",
            "Report Maintenance / Damage"
        ])
        
        all_ids = df_inv['Asset ID'].str.strip().tolist()
        equipment_selected = st.multiselect("Select Equipment IDs", all_ids)
        
        col_a, col_b = st.columns(2)
        with col_a:
            trainer = st.text_input("Instructor / Assessor Name")
        with col_b:
            candidate = st.text_input("Candidate / Customer Name")
            
        # Assessment Date input left empty unless entered manually by the instructor
        assessment_date_input = st.text_input("Assessment Date (DD/MM/YYYY - Optional)", value="", placeholder="e.g. 25/10/2026")
            
        inspection_result = st.selectbox("Inspection Result", ["Pass", "Fail / Maintenance Required", "Not Inspected"])
        notes = st.text_area("Notes / Damage Details (Optional)")
        
        submit_btn = st.form_submit_button("Submit Transaction")
        
        if submit_btn:
            if not equipment_selected:
                st.error("Please select at least one piece of equipment.")
            elif not trainer:
                st.error("Please enter the Instructor Name.")
            else:
                success = tracker.process_transaction(
                    action_type, 
                    equipment_selected, 
                    trainer, 
                    candidate, 
                    assessment_date_input, 
                    inspection_result, 
                    notes
                )
                if success:
                    st.success("✅ Transaction logged successfully! Dashboard updated.")
                    st.balloons()
                else:
                    st.error("❌ Conflict detected! Check terminal logs (e.g. equipment already in use).")