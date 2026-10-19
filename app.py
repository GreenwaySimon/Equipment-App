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

# Display Greenway Training Logo if present
if os.path.exists("greenway_training_logo.jpg"):
  st.image("greenway_training_logo.jpg", width=120)

# App Title Header
st.title("Chainsaw & PPE Live Inventory Hub")

# Run the core tracking logic from tracker.py
tracker.run()
