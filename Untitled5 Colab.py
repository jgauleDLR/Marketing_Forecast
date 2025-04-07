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

    st.subheader("Forecast Category Breakdown")
    forecast_group = filtered_df.groupby('Forecast').agg({'GAAP': 'sum', 'Predicted Value': 'sum'}).reset_index()
    fig, ax = plt.subplots(figsize=(6, 4))
    bar_width = 0.35
    x = range(len(forecast_group))
    bars1 = ax.bar(x, forecast_group['GAAP'], width=bar_width, label='GAAP')
    bars2 = ax.bar([p + bar_width for p in x], forecast_group['Predicted Value'], width=bar_width, label='Predicted', alpha=0.7)
    ax.set_xticks([p + bar_width / 2 for p in x])
    ax.set_xticklabels(forecast_group['Forecast'])
    ax.set_ylabel("Value ($)")
    ax.set_title("GAAP vs. Predicted Value by Forecast")
    ax.legend()
    for bar in bars1:
        height = bar.get_height()
        ax.annotate(f'${height:,.0f}', xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3), textcoords="offset points", ha='center', fontsize=8)
    for bar in bars2:
        height = bar.get_height()
        ax.annotate(f'${height:,.0f}', xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3), textcoords="offset points", ha='center', fontsize=8)
    st.pyplot(fig)

    col4, col5 = st.columns(2)
    with col4:
        st.subheader("Predicted Value by Segmentation")
        seg_group = filtered_df.groupby('Coverage Segmentation')['Predicted Value'].sum().sort_values(ascending=False)
        fig, ax = plt.subplots(figsize=(5, 3))
        bars = ax.bar(seg_group.index, seg_group.values)
        for bar in bars:
            height = bar.get_height()
            ax.annotate(f'${height:,.0f}', xy=(bar.get_x() + bar.get_width()/2, height),
                        xytext=(0, 3), textcoords="offset points", ha='center', fontsize=7)
        ax.set_title("By Segmentation")
        st.pyplot(fig)

    with col5:
        st.subheader("Predicted Value by 1st Line CRO")
        cro_group = filtered_df.groupby('1st Line from CRO')['Predicted Value'].sum().sort_values(ascending=False)
        fig, ax = plt.subplots(figsize=(5, 3))
        bars = ax.bar(cro_group.index, cro_group.values)
        for bar in bars:
            height = bar.get_height()
            ax.annotate(f'${height:,.0f}', xy=(bar.get_x() + bar.get_width()/2, height),
                        xytext=(0, 3), textcoords="offset points", ha='center', fontsize=7)
        ax.set_title("By 1st Line CRO")
        st.pyplot(fig)

    st.subheader("Predicted Value by Close Quarter")
    quarter_group = df.groupby('Close Quarter')['Predicted Value'].sum().reset_index()
    quarter_group = quarter_group.sort_values(by='Close Quarter', key=lambda x: x.map(quarter_sort_key))
    fig, ax = plt.subplots(figsize=(8, 3))
    bars = ax.bar(quarter_group['Close Quarter'], quarter_group['Predicted Value'])
    for bar in bars:
        height = bar.get_height()
        ax.annotate(f'${height:,.0f}', xy=(bar.get_x() + bar.get_width()/2, height),
                    xytext=(0, 3), textcoords="offset points", ha='center', fontsize=7)
    ax.set_title("Predicted Value by Quarter")
    st.pyplot(fig)

    st.subheader("Filtered Opportunity Table")
    st.dataframe(filtered_df[['Account Name', 'Forecast', 'GAAP', 'Conversion Rate', 'Predicted Value', 'Coverage Segmentation', '1st Line from CRO', 'Close Quarter']])

if pacing_file:
    pacing = pd.read_csv(pacing_file) if pacing_file.name.endswith(".csv") else pd.read_excel(pacing_file)

    st.header("3. Q2 Targets and Pacing Overview")
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
