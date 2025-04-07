import streamlit as st

st.set_page_config(page_title="Pipeline Predict App", layout="wide")

st.title("📊 Pipeline Predict Dashboard")

st.markdown("""
Welcome to the **Pipeline Predict** app!

Use the sidebar or page tabs to navigate:
- **Pipeline Predict**: Upload pipeline data, apply conversion logic, and view forecasts.
- **Pacing Tracker & Predictor**: Upload your Q2 pacing file, compare actual vs predicted, and calculate how many more opps are needed.
- **Lead Generation Analysis**: (Coming soon!) Analyze lead quality, source attribution, and conversion insights.

Upload your files and use each tool to align marketing with pipeline performance.
""")
