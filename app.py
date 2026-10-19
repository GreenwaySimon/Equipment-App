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

# Original Header Layout with Logo and Title
col1, col2 = st.columns([1, 5])
with col1:
  if os.path.exists("greenway_training_logo.jpg"):
    st.image("greenway_training_logo.jpg", width=110)
with col2:
  st.title("Chainsaw & PPE Live Inventory Hub")

st.markdown("---")

# Run the core application and dashboard logic from tracker.py
tracker.run()
