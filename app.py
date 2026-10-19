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

# Original Header Layout with Greenway Training Logo and Title
col1, col2 = st.columns([1, 5])
with col1:
  if os.path.exists("greenway_training_logo.jpg"):
    st.image("greenway_training_logo.jpg", width=110)
with col2:
  st.title("Chainsaw & PPE Live Inventory Hub")

st.markdown("---")

# Load Inventory sheet from your Excel tracker file as originally structured
try:
  if os.path.exists(tracker.EXCEL_PATH):
    df_inv = pd.read_excel(tracker.EXCEL_PATH, sheet_name="Inventory")
except Exception:
  pass

# Run the core tracking logic from tracker.py
tracker.run()
