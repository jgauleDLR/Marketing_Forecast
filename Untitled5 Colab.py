import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches

st.set_page_config(page_title="Pipeline Predict Dashboard", layout="wide")
st.title("📈 Pipeline Predict - Forecasting Tool")

st.markdown("""
### 📝 Instructions
To get started:
1. Go to Power BI and download the **raw open funnel details** as a CSV or Excel file.
2. Upload that file below.
3. This app will auto-populate with predicted pipeline metrics, charts, and filters.
""")

# --- File Upload ---
uploaded_file = st.file_uploader("Upload your pipeline data (CSV or Excel)", type=["csv", "xlsx"])

if uploaded_file:
    # Read file
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    # Clean and prep
    df = df.dropna(subset=['GAAP', 'Forecast'])
    df['Forecast'] = df['Forecast'].str.title()
    df['Close Quarter'] = df['Close Quarter'].astype(str)

    st.sidebar.header("Adjust Forecast Conversion Rates")
    commit_rate = st.sidebar.slider("Commit Conversion %", 0, 100, 80) / 100
    upside_rate = st.sidebar.slider("Upside Conversion %", 0, 100, 50) / 100
    pipeline_rate = st.sidebar.slider("Pipeline Conversion %", 0, 100, 30) / 100

    rate_map = {
        "Commit": commit_rate,
        "Upside": upside_rate,
        "Pipeline": pipeline_rate
    }

    df['Conversion Rate'] = df['Forecast'].map(rate_map).fillna(0)
    df['Predicted Value'] = df['GAAP'] * df['Conversion Rate']

    # Filters
    st.sidebar.header("Filter Data")
    segmentation_options = df['Coverage Segmentation'].dropna().unique()
    cro_line_options = df['1st Line from CRO'].dropna().unique()
    quarter_options = sorted(df['Close Quarter'].dropna().unique())

    selected_segmentation = st.sidebar.selectbox("Coverage Segmentation", options=segmentation_options)
    selected_cro = st.sidebar.selectbox("1st Line from CRO", options=cro_line_options)
    selected_quarter = st.sidebar.selectbox("Close Quarter", options=quarter_options)

    filtered_df = df[
        (df['Coverage Segmentation'] == selected_segmentation) &
        (df['1st Line from CRO'] == selected_cro) &
        (df['Close Quarter'] == selected_quarter)
    ]

    total_pipeline = filtered_df['GAAP'].sum()
    predicted_total = filtered_df['Predicted Value'].sum()
    total_opps = len(filtered_df)

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Opportunities", f"{total_opps}")
    col2.metric("Total Pipeline Value", f"${total_pipeline:,.0f}")
    col3.metric("Predicted Closed Value", f"${predicted_total:,.0f}")

    st.subheader("📊 Forecast Category Breakdown")
    forecast_group = filtered_df.groupby('Forecast').agg({
        'GAAP': 'sum',
        'Predicted Value': 'sum'
    }).reset_index()

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

    st.subheader("📌 Predicted Value by Segmentation")
    seg_group = filtered_df.groupby('Coverage Segmentation')['Predicted Value'].sum().sort_values(ascending=False)
    st.bar_chart(seg_group)

    st.subheader("📌 Predicted Value by 1st Line CRO")
    cro_group = filtered_df.groupby('1st Line from CRO')['Predicted Value'].sum().sort_values(ascending=False)
    st.bar_chart(cro_group)

    st.subheader("📆 Predicted Value by Close Quarter")
    quarter_group = df.groupby('Close Quarter')['Predicted Value'].sum().reset_index()
    quarter_group = quarter_group.sort_values(by='Close Quarter')
    st.line_chart(quarter_group.set_index('Close Quarter'))

    st.subheader("📋 Filtered Opportunity Table")
    st.dataframe(filtered_df[['Account Name', 'Forecast', 'GAAP', 'Conversion Rate', 'Predicted Value', 'Coverage Segmentation', '1st Line from CRO', 'Close Quarter']])

    st.subheader("📎 Full Dataset (Raw View)")
    st.dataframe(df)

    # Export to PPTX
    def generate_ppt(data_summary):
        prs = Presentation()
        slide_layout = prs.slide_layouts[5]  # Title Only
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = "Pipeline Predict Summary"

        content = f"""
        Total Opportunities: {total_opps}
        Total Pipeline Value: ${total_pipeline:,.0f}
        Predicted Closed Value: ${predicted_total:,.0f}
        Selected Filters: {selected_segmentation}, {selected_cro}, {selected_quarter}
        """
        textbox = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(8), Inches(4))
        tf = textbox.text_frame
        tf.text = content

        output = BytesIO()
        prs.save(output)
        output.seek(0)
        return output

    if st.button("📤 Export Summary to PowerPoint"):
        pptx_file = generate_ppt(filtered_df)
        st.download_button(label="Download PPTX", data=pptx_file, file_name="Pipeline_Predict_Summary.pptx")

else:
    st.info("👆 Upload a pipeline file to begin.")
