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

# --- Upload Sections ---
st.header("1. Upload Your Data")
pipeline_file = st.file_uploader("Upload Pipeline File (CSV or Excel)", type=["csv", "xlsx"], key="pipeline")
pacing_file = st.file_uploader("Upload Pacing Tracker (CSV or Excel)", type=["csv", "xlsx"], key="pacing")

# --- Process Pipeline Data ---
total_pipeline = predicted_total = total_opps = 0
predicted_close_total = 0

if pipeline_file:
    st.header("2. Pipeline Predict Dashboard")
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
    predicted_close_total = df['Predicted Value'].sum()

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

if pacing_file:
    st.header("3. Q2 Pacing Tracker + Forecast Visuals")
    pacing = pd.read_csv(pacing_file) if pacing_file.name.endswith(".csv") else pd.read_excel(pacing_file)

    try:
        creation_targets = pacing[(pacing['Source'] == 'Q2 Target') & (pacing['Metric Group'] == 'Creation') & (pacing['Metric Type'] == '$')]
        creation_pacing = pacing[(pacing['Source'] == 'Week 1 Pacing') & (pacing['Metric Group'] == 'Creation') & (pacing['Metric Type'] == '$')]
        segments = ['ALL', 'Enterprise', 'Commercial', 'Global']
        bar_data = pd.DataFrame({
            'Segment': segments,
            'Target': [creation_targets[seg].values[0] for seg in segments],
            'Week 1': [creation_pacing[seg].values[0] for seg in segments]
        }).set_index('Segment')

        st.subheader("Target vs Week 1 Pacing")
        fig, ax = plt.subplots(figsize=(6, 4))
        bar_width = 0.35
        x = range(len(bar_data))
        bars1 = ax.bar(x, bar_data['Target'], width=bar_width, label='Target', color='green')
        bars2 = ax.bar([p + bar_width for p in x], bar_data['Week 1'], width=bar_width, label='Week 1 Pacing', color='steelblue')
        ax.set_xticks([p + bar_width / 2 for p in x])
        ax.set_xticklabels(bar_data.index)
        ax.set_ylabel("Creation Amount ($M)")
        ax.set_title("Q2 Creation Target vs Week 1 Pacing by Segment")
        ax.legend()
        for bar in bars1:
            height = bar.get_height()
            ax.annotate(f'{height:,.0f}', xy=(bar.get_x() + bar.get_width() / 2, height),
                        xytext=(0, 3), textcoords="offset points", ha='center', fontsize=8)
        for bar in bars2:
            height = bar.get_height()
            ax.annotate(f'{height:,.0f}', xy=(bar.get_x() + bar.get_width() / 2, height),
                        xytext=(0, 3), textcoords="offset points", ha='center', fontsize=8)
        st.pyplot(fig)

        st.subheader("Forecast vs Target Over Time")
        q2_target_total = 115_000_000
        actual_weekly_pacing = [8_000_000]
        weeks_in_q2 = 13
        weeks_completed = len(actual_weekly_pacing)
        weeks_remaining = weeks_in_q2 - weeks_completed
        weekly_projection = predicted_close_total / weeks_remaining if weeks_remaining > 0 else 0
        projected_weekly = [weekly_projection] * weeks_remaining
        weeks = [f"Week {i+1}" for i in range(weeks_in_q2)]
        target_line = [q2_target_total / weeks_in_q2 * (i + 1) for i in range(weeks_in_q2)]
        actual_line = actual_weekly_pacing + [None] * weeks_remaining
        projection_line = actual_weekly_pacing + list(pd.Series(projected_weekly).cumsum() + actual_weekly_pacing[-1])

        fig2, ax2 = plt.subplots(figsize=(10, 5))
        ax2.plot(weeks, target_line, label="Q2 Target", linestyle='--', color='green')
        ax2.plot(weeks, actual_line, label="Actual Pacing", marker='o', color='blue')
        ax2.plot(weeks, projection_line, label="Projected Close (Pipeline)", marker='o', linestyle='-', color='orange')
        ax2.set_title("Q2 Opportunity Creation Forecast vs Target")
        ax2.set_xlabel("Week")
        ax2.set_ylabel("Cumulative Creation Amount ($)")
        ax2.set_xticks(range(0, weeks_in_q2))
        ax2.set_xticklabels(weeks, rotation=45)
        ax2.legend()
        ax2.grid(True)
        st.pyplot(fig2)

        st.subheader("Gap Closure Calculator")
        gap_amount = q2_target_total - predicted_close_total
        st.write(f"Target: ${q2_target_total:,.0f}")
        st.write(f"Projected from Pipeline: ${predicted_close_total:,.0f}")
        st.write(f"Gap Remaining: ${gap_amount:,.0f}")
        avg_opp_size = st.slider("Average Opportunity Size ($)", 10000, 1000000, 250000, step=10000)
        if avg_opp_size > 0:
            opps_needed = int(gap_amount / avg_opp_size)
            st.metric("Opportunities Needed to Close Gap", f"{opps_needed} new opps")

    except Exception as e:
        st.warning(f"Could not process pacing data: {e}")
