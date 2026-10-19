import os
import pandas as pd
import streamlit as st
import tracker

# Page Configuration
st.set_page_config(
    page_title="Chainsaw & PPE Live Inventory Hub",
    layout="centered",
    initial_sidebar_state="expanded",
)

# Ensure the Excel file and 'Inventory' sheet exist to prevent FileNotFoundError
if not os.path.exists(tracker.EXCEL_FILE):
  with pd.ExcelWriter(tracker.EXCEL_FILE, engine="openpyxl") as writer:
    # Create an initial empty Inventory template sheet
    pd.DataFrame(
        columns=["Asset ID", "Status", "Category"]
    ).to_excel(  # Adjust columns as needed for your inventory
        writer, sheet_name="Inventory", index=False
    )
    # Also ensure log sheet structure exists if needed
    pd.DataFrame(columns=tracker.COLUMNS).to_excel(
        writer, sheet_name="Sheet1", index=False
    )

# Run the main app logic from tracker.py
tracker.run()
