import streamlit as st
import tracker

# Page Configuration
st.set_page_config(
    page_title="Chainsaw & PPE Live Inventory Hub",
    layout="centered",
    initial_sidebar_state="expanded",
)

# Run the application logic from tracker.py
tracker.run()
