import streamlit as st
import tracker

st.set_page_config(
    page_title="Live Equipment Inventory Hub",
    page_icon="🪵",
    layout="wide",
)

if __name__ == "__main__":
    tracker.run()
