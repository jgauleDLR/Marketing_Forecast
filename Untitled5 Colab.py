import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches

st.set_page_config(page_title="Pipeline Predict Dashboard", layout="wide")
st.title("Pipeline Predict - Forecasting Tool")

st.markdown("""
### Instructions
To get started:
1. Upload your **Pipeline Data** file (CSV or Excel).
2. Upload your **Weekly Pacing Tracker** (CSV or Excel).
3. This app will auto-populate predicted pipeline, targets, and pacing visuals.
""")

# --- Upload Pipeline Section ---
st.header("1. Upload Pipeline Data")
pipeline_file = st.file_uploader("Upload Pipeline File (CSV or Excel)", type=["csv", "xlsx"], key="pipeline")

# --- Upload Pacing Tracker Section ---
st.header("2. Upload Pacing Targets")
pacing_file = st.file_uploader("Upload Pacing Tracker (CSV or Excel)", type=["csv", "xlsx"], key="pacing")

# --- Placeholder containers ---
total_pipeline = predicted_total = total_opps = 0

if pipeline_file:
    if pipeline_file.name.endswith(".csv"):
        df = pd.read_csv(pipeline_file)
    else:
        df = pd.read_excel(pipeline_file)

    df = df.dropna(subset=['GAAP', 'Forecast'])
    df['Forecast'] = df['Forecast'].str.title()
    df['Close Quarter'] = df['Close Quarter'].astype(str)

    def quarter_sort_key(q):
        try:
            parts = q.split('-')
            return int(parts[1]) * 10 + int(parts[0].replace('Q', ''))
        except:
            return 9999

    df['Close_Quarter_Sort'] = df['Close Quarter'].apply(quarter_sort_key)

    commit_rate = st.sidebar.slider("Commit Conversion %", 0, 100, 80) / 100
    upside_rate = st.sidebar.slider("Upside Conversion %", 0, 100, 50) / 100
    pipeline_rate = st.sidebar.slider("Pipeline Conversion %", 0, 100, 30) / 100

    rate_map = {"Commit": commit_rate, "Upside": upside_rate, "Pipeline": pipeline_rate}
    df['Conversion Rate'] = df['Forecast'].map(rate_map).fillna(0)
    df['Predicted Value'] = df['GAAP'] * df['Conversion Rate']

    allowed_segmentations = ['Enterprise', 'Commercial', 'Global']
    segmentation_options = sorted([s for s in df['Coverage Segmentation'].dropna().unique() if s in allowed_segmentations])
    cro_line_options = sorted(df['1st Line from CRO'].dropna().unique())
    quarter_options = sorted(df['Close Quarter'].dropna().unique(), key=quarter_sort_key)

    selected_segmentation = st.sidebar.multiselect("Coverage Segmentation", options=segmentation_options, default=segmentation_options)
    selected_cro = st.sidebar.multiselect("1st Line from CRO", options=cro_line_options, default=cro_line_options)
    selected_quarter = st.sidebar.multiselect("Close Quarter", options=quarter_options, default=quarter_options)

    filtered_df = df[df['Coverage Segmentation'].isin(selected_segmentation) &
                     df['1st Line from CRO'].isin(selected_cro) &
                     df['Close Quarter'].isin(selected_quarter)]

    total_pipeline = filtered_df['GAAP'].sum()
    predicted_total = filtered_df['Predicted Value'].sum()
    total_opps = len(filtered_df)

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Opportunities", f"{total_opps}")
    col2.metric("Total Pipeline Value", f"${total_pipeline:,.0f}")
    col3.metric("Predicted Closed Value", f"${predicted_total:,.0f}")

if pacing_file:
    pacing = pd.read_csv(pacing_file) if pacing_file.name.endswith(".csv") else pd.read_excel(pacing_file)

    st.markdown("### Q2 Targets and Pacing Overview")
    try:
        creation_targets = pacing[(pacing['Source'] == 'Q2 Target') & (pacing['Metric Group'] == 'Creation') & (pacing['Metric Type'] == '$')]
        creation_pacing = pacing[(pacing['Source'] == 'Week 1 Pacing') & (pacing['Metric Group'] == 'Creation') & (pacing['Metric Type'] == '$')]

        segments = ['ALL', 'Enterprise', 'Commercial', 'Global']
        bar_data = pd.DataFrame({
            'Segment': segments,
            'Target': [creation_targets[seg].values[0] for seg in segments],
            'Week 1': [creation_pacing[seg].values[0] for seg in segments]
        })
        bar_data.set_index('Segment', inplace=True)

        st.bar_chart(bar_data)
        st.dataframe(bar_data.reset_index().assign(**{
            '% to Target': lambda df_: (df_['Week 1'] / df_['Target'] * 100).round(1)
        }))

    except Exception as e:
        st.warning(f"Could not process pacing data: {e}")
