import streamlit as st
import tracker

# Page Configuration
st.set_page_config(
    page_title="Instructor Action Portal",
    layout="centered",
    initial_sidebar_state="expanded",
)

# Run the main app logic from tracker.py
tracker.run()
