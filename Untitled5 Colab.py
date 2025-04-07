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
1. Download the raw open funnel details from Power BI and upload below.
2. Upload your updated weekly pacing Excel sheet.
3. This app will auto-populate predicted pipeline, targets, and pacing visuals.
""")

# --- Uploaders ---
pipeline_file = st.file_uploader("Upload Pipeline File (CSV or Excel)", type=["csv", "xlsx"], key="pipeline")
pacing_file = st.file_uploader("Upload Pacing Tracker (Excel)", type=["csv", "xlsx"], key="pacing")

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
    pacing = pd.read_csv(pacing_file) if pacing_file.name.endswith(".csv") else pd.read_excel(pacing_file, header=None)

    try:
        targets_row = pacing[pacing[18] == 'Target'].index[0] + 1
        target_amount = float(pacing.iloc[targets_row, 18])
        current_amount = float(pacing.iloc[targets_row, 19])
        target_count = float(pacing.iloc[targets_row + 1, 18])
        current_count = float(pacing.iloc[targets_row + 1, 19])

        st.markdown("### Marketing Q2 Target vs Current Pacing")
        st.write(f"Amount Pacing: ${current_amount:,.0f} / ${target_amount:,.0f} ({(current_amount/target_amount)*100:.1f}%)")
        st.write(f"Count Pacing: {int(current_count)} / {int(target_count)} ({(current_count/target_count)*100:.1f}%)")

        fig, ax = plt.subplots(figsize=(5, 3))
        ax.bar(['Target Amount', 'Current Amount'], [target_amount, current_amount], color=['#ccc', '#2a9d8f'])
        ax.set_title("Amount Pacing vs Target", fontsize=10)
        for i, v in enumerate([target_amount, current_amount]):
            ax.text(i, v + max([target_amount, current_amount])*0.02, f'${v:,.0f}', ha='center', fontsize=7)
        st.pyplot(fig)

    except Exception as e:
        st.warning(f"Could not extract pacing summary: {e}")

    try:
        weekly_start = pacing[pacing.apply(lambda row: row.astype(str).str.contains('Enterprise').any(), axis=1)].index.min() - 1
        weekly_end = weekly_start + 4
        weekly_pacing = pacing.iloc[weekly_start:weekly_end]
        st.markdown("### Weekly Creation Pacing by Segment")
        st.dataframe(weekly_pacing.reset_index(drop=True))
    except Exception as e:
        st.warning(f"Could not extract weekly pacing: {e}")
