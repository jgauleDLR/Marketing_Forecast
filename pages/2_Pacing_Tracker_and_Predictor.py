import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.title("📊 Pacing Tracker & Forecasting")

pacing_file = st.file_uploader("Upload Q2 Pacing File (CSV or Excel)", type=["csv", "xlsx"], key="pacing")

if pacing_file:
    pacing_df = pd.read_csv(pacing_file) if pacing_file.name.endswith(".csv") else pd.read_excel(pacing_file)
    creation_targets = pacing_df[(pacing_df['Source'] == 'Q2 Target') & (pacing_df['Metric Group'] == 'Creation') & (pacing_df['Metric Type'] == '$')]
    creation_pacing = pacing_df[(pacing_df['Source'].str.contains('Week')) & (pacing_df['Metric Group'] == 'Creation') & (pacing_df['Metric Type'] == '$')]

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
    ax.bar(x, bar_data['Target'], width=bar_width, label='Target', color='green')
    ax.bar([p + bar_width for p in x], bar_data['Week 1'], width=bar_width, label='Week 1 Pacing', color='steelblue')
    ax.set_xticks([p + bar_width / 2 for p in x])
    ax.set_xticklabels(bar_data.index)
    ax.set_ylabel("Creation Amount ($M)")
    ax.set_title("Q2 Creation Target vs Week 1 Pacing by Segment")
    ax.legend()
    st.pyplot(fig)

    # Forecast Simulation
    st.subheader("Forecast vs Target Over Time")
    predicted_close_total = 56700000  # this would be passed from shared state or file in full app
    q2_target_total = 115000000
    actual_weekly_pacing = [8000000]
    weeks_in_q2 = 13
    weeks_completed = len(actual_weekly_pacing)
    weeks_remaining = weeks_in_q2 - weeks_completed
    weekly_projection = predicted_close_total / weeks_remaining
    projected_weekly = [weekly_projection] * weeks_remaining
    weeks = [f"Week {i+1}" for i in range(weeks_in_q2)]
    target_line = [q2_target_total / weeks_in_q2 * (i + 1) for i in range(weeks_in_q2)]
    actual_line = actual_weekly_pacing + [None] * weeks_remaining
    projection_line = actual_weekly_pacing + list(pd.Series(projected_weekly).cumsum() + actual_weekly_pacing[-1])

    fig2, ax2 = plt.subplots(figsize=(10, 5))
    ax2.plot(weeks, target_line, label="Q2 Target", linestyle='--', color='green')
    ax2.plot(weeks, actual_line, label="Actual Pacing", marker='o', color='blue')
    ax2.plot(weeks, projection_line, label="Projected Close (Pipeline)", marker='o', color='orange')
    ax2.set_title("Q2 Opportunity Creation Forecast vs Target")
    ax2.set_xlabel("Week")
    ax2.set_ylabel("Cumulative Creation Amount ($)")
    ax2.set_xticks(range(0, weeks_in_q2))
    ax2.set_xticklabels(weeks, rotation=45)
    ax2.legend()
    ax2.grid(True)
    st.pyplot(fig2)

    # Gap Closure Tool
    st.subheader("Gap Closure Calculator")
    gap_amount = q2_target_total - predicted_close_total
    st.write(f"Target: ${q2_target_total:,.0f}")
    st.write(f"Projected from Pipeline: ${predicted_close_total:,.0f}")
    st.write(f"Gap Remaining: ${gap_amount:,.0f}")
    avg_opp_size = st.slider("Average Opportunity Size ($)", 10000, 1000000, 250000, step=10000)
    if avg_opp_size > 0:
        opps_needed = int(gap_amount / avg_opp_size)
        st.metric("Opportunities Needed to Close Gap", f"{opps_needed} new opps")
