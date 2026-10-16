import openpyxl
import pandas as pd
from datetime import datetime, date
import json
import smtplib
from email.message import EmailMessage

EXCEL_PATH = 'Chainsaw_Equipment_Tracker.xlsx'

def send_maintenance_email(equipment_id, trainer, candidate, notes):
    sender_email = "simon@greenwaytraining.co.uk"
    app_password = "YOUR_EMAIL_APP_PASSWORD" # Configure your SMTP app password here
    receiver_email = "simon@greenwaytraining.co.uk"
    
    msg = EmailMessage()
    msg.set_subject(f"⚠️ MAINTENANCE ALERT: Equipment {equipment_id} Flagged Damaged")
    msg.set_from(sender_email)
    msg.set_to(receiver_email)
    
    msg.set_content(f"""
Hello Simon,

An item of equipment has been flagged for maintenance or marked as failed during inspection.

- Equipment ID: {equipment_id}
- Reported By (Instructor): {trainer}
- Customer / Candidate: {candidate if candidate else 'N/A'}
- Damage / Maintenance Notes: {notes}
- Timestamp: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}

Please check the inventory dashboard and arrange for maintenance immediately.
""")
    
    try:
        with smtplib.SMTP('smtp.office365.com', 587) as server:
            server.starttls()
            server.login(sender_email, app_password)
            server.send_message(msg)
        print("[EMAIL SENT] Maintenance alert sent to simon@greenwaytraining.co.uk")
    except Exception as e:
        print(f"[EMAIL NOTICE] Maintenance logged successfully. (SMTP note: {e})")

def process_transaction(action_type, equipment_ids, trainer, candidate, assessment_date_str, inspection_result, notes):
    wb = openpyxl.load_workbook(EXCEL_PATH)
    df_inv = pd.read_excel(EXCEL_PATH, sheet_name='Inventory')
    df_resp = pd.read_excel(EXCEL_PATH, sheet_name='Form Responses 1')
    
    if isinstance(equipment_ids, str):
        try:
            eq_list = json.loads(equipment_ids)
        except:
            eq_list = [e.strip() for e in equipment_ids.split(',')]
    else:
        eq_list = equipment_ids
        
    eq_list = [e.strip().replace('\xa0', '') for e in eq_list]
    
    active_allocations = get_current_allocations(df_resp, df_inv)
    
    new_date = None
    if assessment_date_str and assessment_date_str.strip() not in ['', 'NaT', 'None']:
        date_clean = assessment_date_str.strip()
        try:
            # Try parsing DD/MM/YYYY first
            new_date = datetime.strptime(date_clean[:10], '%d/%m/%Y').date()
        except:
            try:
                # Fallback to YYYY-MM-DD
                new_date = datetime.strptime(date_clean[:10], '%Y-%m-%d').date()
            except:
                pass

    if 'Check Out' in action_type:
        for eq in eq_list:
            if eq in active_allocations:
                curr = active_allocations[eq]
                curr_status = curr['Status']
                if curr_status == '🔴 IN USE':
                    print(f"\n[ERROR CONFLICT] Asset '{eq}' is currently IN USE by {curr['Customer']}.")
                    return False

    formatted_date_str = ""
    if new_date:
        formatted_date_str = new_date.strftime('%d/%m/%Y')

    new_row = {
        'Timestamp': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
        'Action Type': action_type,
        'Equipment ID': json.dumps(eq_list),
        'Trainer Name': trainer,
        'Candidate Name': candidate,
        'Assessment Date': formatted_date_str if formatted_date_str else None,
        'Inspection Result': inspection_result,
        'Notes / Damage Details': notes
    }
    
    ws_resp = wb['Form Responses 1']
    ws_resp.append([
        new_row['Timestamp'],
        new_row['Action Type'],
        new_row['Equipment ID'],
        new_row['Trainer Name'],
        new_row['Candidate Name'],
        new_row['Assessment Date'],
        new_row['Inspection Result'],
        new_row['Notes / Damage Details']
    ])
    
    wb.save(EXCEL_PATH)
    update_live_dashboard()
    
    if 'Maintenance' in action_type or 'Fail' in str(inspection_result):
        for eq in eq_list:
            send_maintenance_email(eq, trainer, candidate, notes)
            
    print(f"\n[SUCCESS] Transaction '{action_type}' recorded successfully for {len(eq_list)} item(s)!")
    return True

def update_assessment_date(asset_id, new_date_str):
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws_resp = wb['Form Responses 1']
    
    timestamp = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    eq_json = json.dumps([asset_id])
    
    formatted_date = ""
    if new_date_str:
        date_clean = str(new_date_str).strip()
        try:
            if '/' in date_clean:
                dt = datetime.strptime(date_clean[:10], '%d/%m/%Y')
            else:
                dt = datetime.strptime(date_clean[:10], '%Y-%m-%d')
            formatted_date = dt.strftime('%d/%m/%Y')
        except:
            formatted_date = date_clean
            
    ws_resp.append([
        timestamp,
        'Update Assessment Date',
        eq_json,
        '',
        '',
        formatted_date if formatted_date else None,
        'Pass',
        'Assessment date updated via live dashboard'
    ])
    wb.save(EXCEL_PATH)
    update_live_dashboard()

def get_current_allocations(df_resp, df_inv):
    asset_states = {}
    today_date = date.today()
    
    for _, row in df_inv.iterrows():
        aid = str(row['Asset ID']).strip()
        asset_states[aid] = {'Status': '🟢 AVAILABLE', 'Instructor': '', 'Customer': '', 'Assessment Date': '', 'Last Inspection': 'Not Inspected', 'Notes': ''}
        
    if not df_resp.empty and 'Timestamp' in df_resp.columns:
        for _, row in df_resp.iterrows():
            action = str(row.get('Action Type', ''))
            trainer = str(row.get('Trainer Name', ''))
            candidate = str(row.get('Candidate Name', ''))
            assessment_date = row.get('Assessment Date')
            inspection = str(row.get('Inspection Result', 'Pass'))
            notes = str(row.get('Notes / Damage Details', ''))
            if notes == 'nan': notes = ''
            
            formatted_as_date = ""
            parsed_assess_date = None
            if pd.notna(assessment_date) and str(assessment_date) != 'NaT':
                s_date = str(assessment_date).strip()
                try:
                    if '/' in s_date:
                        parsed_assess_date = datetime.strptime(s_date[:10], '%d/%m/%Y').date()
                    else:
                        parsed_assess_date = datetime.strptime(s_date[:10], '%Y-%m-%d').date()
                    formatted_as_date = parsed_assess_date.strftime('%d/%m/%Y')
                except:
                    formatted_as_date = s_date[:10]
            
            eq_raw = str(row.get('Equipment ID', '[]'))
            try:
                eq_list = json.loads(eq_raw)
            except:
                eq_list = [eq_raw]
                
            for eq in eq_list:
                eid = eq.strip().replace('\xa0', '')
                if eid in asset_states:
                    if 'Check Out' in action or 'Update Assessment Date' in action:
                        status = '🔴 IN USE'
                        # Automatically changes to RESERVED on the exact assessment date
                        if parsed_assess_date and parsed_assess_date == today_date:
                            status = '🟡 RESERVED FOR ASSESSMENT'
                            
                        asset_states[eid]['Status'] = status
                        if trainer: asset_states[eid]['Instructor'] = trainer
                        if candidate: asset_states[eid]['Customer'] = candidate
                        if formatted_as_date: 
                            asset_states[eid]['Assessment Date'] = formatted_as_date
                        elif action == 'Update Assessment Date' and not formatted_as_date:
                            asset_states[eid]['Assessment Date'] = ''
                            
                        asset_states[eid]['Last Inspection'] = inspection
                        if notes: asset_states[eid]['Notes'] = notes
                    elif 'Return' in action:
                        asset_states[eid]['Status'] = '🟢 AVAILABLE'
                        asset_states[eid]['Instructor'] = ''
                        asset_states[eid]['Customer'] = ''
                        asset_states[eid]['Assessment Date'] = ''
                        asset_states[eid]['Notes'] = ''
                        # Last Inspection is intentionally preserved here
                    elif 'Maintenance' in action or 'Fail' in inspection:
                        asset_states[eid]['Status'] = '⚠️ MAINTENANCE'
                        asset_states[eid]['Instructor'] = trainer
                        asset_states[eid]['Customer'] = candidate if candidate != 'nan' else ''
                        asset_states[eid]['Last Inspection'] = inspection
                        asset_states[eid]['Notes'] = notes
    return asset_states

def update_live_dashboard():
    wb = openpyxl.load_workbook(EXCEL_PATH)
    df_inv = pd.read_excel(EXCEL_PATH, sheet_name='Inventory')
    df_resp = pd.read_excel(EXCEL_PATH, sheet_name='Form Responses 1')
    
    asset_states = get_current_allocations(df_resp, df_inv)
    
    ws = wb['Live Dashboard']
    for row in range(2, ws.max_row + 1):
        for col in range(1, 9):
            ws.cell(row=row, column=col, value=None)
            
    row_idx = 2
    for aid, state in asset_states.items():
        cat_match = df_inv.loc[df_inv['Asset ID'].str.strip() == aid, 'Category']
        cat_val = cat_match.values[0] if not cat_match.empty else ''
        
        ws.cell(row=row_idx, column=1, value=aid)
        ws.cell(row=row_idx, column=2, value=cat_val)
        ws.cell(row=row_idx, column=3, value=state['Status'])
        ws.cell(row=row_idx, column=4, value=state['Instructor'])
        ws.cell(row=row_idx, column=5, value=state['Customer'])
        ws.cell(row=row_idx, column=6, value=state['Assessment Date'])
        ws.cell(row=row_idx, column=7, value=state['Last Inspection'])
        ws.cell(row=row_idx, column=8, value=state['Notes'])
        row_idx += 1
    wb.save(EXCEL_PATH)