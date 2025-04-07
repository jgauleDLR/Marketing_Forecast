import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.title("📈 Pipeline Predict")

pipeline_file = st.file_uploader("Upload Pipeline File (CSV or Excel)", type=["csv", "xlsx"], key="pipeline")

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

    st.subheader("Metrics Overview")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Opportunities", f"{len(filtered_df)}")
    col2.metric("Total Pipeline Value", f"${filtered_df['GAAP'].sum():,.0f}")
    col3.metric("Predicted Closed Value", f"${filtered_df['Predicted Value'].sum():,.0f}")

    st.subheader("GAAP vs. Predicted Value by Forecast")
    forecast_group = filtered_df.groupby('Forecast').agg({'GAAP': 'sum', 'Predicted Value': 'sum'}).reset_index()
    fig, ax = plt.subplots(figsize=(6, 4))
    bar_width = 0.35
    x = range(len(forecast_group))
    ax.bar(x, forecast_group['GAAP'], width=bar_width, label='GAAP')
    ax.bar([p + bar_width for p in x], forecast_group['Predicted Value'], width=bar_width, label='Predicted', alpha=0.7)
    ax.set_xticks([p + bar_width / 2 for p in x])
    ax.set_xticklabels(forecast_group['Forecast'])
    ax.set_ylabel("Value ($)")
    ax.legend()
    st.pyplot(fig)

    st.subheader("Predicted Value by Segmentation")
    seg_group = filtered_df.groupby('Coverage Segmentation').agg({'Predicted Value': 'sum'}).reset_index()
    fig2, ax2 = plt.subplots(figsize=(6, 3))
    ax2.bar(seg_group['Coverage Segmentation'], seg_group['Predicted Value'], color='steelblue')
    for i, v in enumerate(seg_group['Predicted Value']):
        ax2.text(i, v + 0.01 * max(seg_group['Predicted Value']), f"${v:,.0f}", ha='center', fontsize=8)
    st.pyplot(fig2)

    st.subheader("Predicted Value by 1st Line CRO")
    cro_group = filtered_df.groupby('1st Line from CRO').agg({'Predicted Value': 'sum'}).reset_index()
    fig3, ax3 = plt.subplots(figsize=(8, 3))
    ax3.bar(cro_group['1st Line from CRO'], cro_group['Predicted Value'], color='orange')
    ax3.tick_params(axis='x', rotation=45)
    for i, v in enumerate(cro_group['Predicted Value']):
        ax3.text(i, v + 0.01 * max(cro_group['Predicted Value']), f"${v:,.0f}", ha='center', fontsize=8)
    st.pyplot(fig3)

    st.subheader("Predicted Value by Close Quarter")
    q_group = filtered_df.groupby(['Close Quarter', 'Close_Quarter_Sort']).agg({'Predicted Value': 'sum'}).reset_index().sort_values('Close_Quarter_Sort')
    fig4, ax4 = plt.subplots(figsize=(8, 3))
    ax4.bar(q_group['Close Quarter'], q_group['Predicted Value'], color='green')
    for i, v in enumerate(q_group['Predicted Value']):
        ax4.text(i, v + 0.01 * max(q_group['Predicted Value']), f"${v:,.0f}", ha='center', fontsize=8)
    st.pyplot(fig4)

    st.subheader("Filtered Opportunity Table")
    st.dataframe(filtered_df[['Account Name', 'Forecast', 'GAAP', 'Predicted Value', 'Coverage Segmentation', '1st Line from CRO', 'Close Quarter']])
