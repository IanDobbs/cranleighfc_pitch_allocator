import streamlit as st
import pandas as pd
import io
from datetime import datetime
import os

# Import your existing modules
from CranleighFC_Pitch_Allocation_PROD import (
    load_and_validate_fixtures,
    solve_allocation,
    generate_excel_schedule,
    generate_html_schedule,
    pitches,
    valid_teams
)

# Page configuration
st.set_page_config(
    page_title="Cranleigh FC Pitch Allocator",
    page_icon="⚽",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1e3a8a;
        margin-bottom: 1rem;
    }
    .stButton>button {
        background-color: #2563eb;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 0.5rem 2rem;
    }
    .success-box {
        padding: 1rem;
        background-color: #d1fae5;
        border-left: 4px solid #10b981;
        border-radius: 4px;
        margin: 1rem 0;
    }
    .warning-box {
        padding: 1rem;
        background-color: #fef3c7;
        border-left: 4px solid #f59e0b;
        border-radius: 4px;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown('<p class="main-header">⚽ Cranleigh FC Pitch Allocator</p>', unsafe_allow_html=True)
st.markdown("Automated pitch allocation system for home fixtures")


# ----------------------------------------
# LOAD DEFAULT FIXTURES AUTOMATICALLY
# ----------------------------------------

DEFAULT_FILE = "cranleigh_home_fixtures.csv"

if not os.path.exists(DEFAULT_FILE):
    st.error(
        f"❌ **'{DEFAULT_FILE}' not found.**\n\n"
        "Please add the file to your project root in the GitHub repository."
    )
    st.stop()

try:
    df = pd.read_csv(DEFAULT_FILE)

    st.markdown("## Loaded Default Home Fixtures")
    st.info(f"Loaded automatically from `{DEFAULT_FILE}`")
    st.dataframe(df.head(10), use_container_width=True)

except Exception as e:
    st.error(f"❌ Failed to read `{DEFAULT_FILE}`: {str(e)}")
    st.stop()


# ----------------------------------------
# BASIC STATISTICS & VALIDATION
# ----------------------------------------

st.markdown("### Fixture Summary")

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Total Fixtures", len(df))
with col2:
    cup_count = df['prefix'].str.contains('cup', case=False, na=False).sum()
    st.metric("Cup Fixtures", cup_count)
with col3:
    st.metric("Match Days", df['match_date'].nunique())
with col4:
    st.metric("Teams", df['home_team_clean'].nunique())

# Validation checks
errors = []

required_cols = ['match_date', 'match_time', 'home_team_clean']
missing = [c for c in required_cols if c not in df.columns]
if missing:
    errors.append("Missing columns: " + ", ".join(missing))

unknown_teams = [
    t for t in df['home_team_clean'].unique()
    if t not in valid_teams
]
if unknown_teams:
    errors.append("Unknown teams: " + ", ".join(unknown_teams[:5]))

if errors:
    st.markdown('<div class="warning-box">', unsafe_allow_html=True)
    st.warning("⚠️ Validation Warnings")
    for e in errors:
        st.markdown(f"- {e}")
    st.markdown('</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="success-box">', unsafe_allow_html=True)
    st.success("✅ All validations passed!")
    st.markdown('</div>', unsafe_allow_html=True)


# ----------------------------------------
# ALLOCATION CONTROLS
# ----------------------------------------

st.markdown("## Run Pitch Allocation")

col1, col2 = st.columns([1, 1])
with col1:
    timeout = st.number_input(
        "Solver Timeout (seconds)",
        min_value=5, max_value=120, value=30
    )

with col2:
    st.markdown("<br>", unsafe_allow_html=True)
    allocate_button = st.button("🚀 Allocate Pitches", use_container_width=True)


# ----------------------------------------
# RUN ALLOCATION
# ----------------------------------------

if allocate_button:
    with st.spinner("🔄 Allocating pitches… this may take 30–60 seconds..."):

        try:
            fixtures, slots_by_date = load_and_validate_fixtures(DEFAULT_FILE)
            result = solve_allocation(fixtures, slots_by_date, timeout=timeout)

            if result is None or len(result) == 0:
                st.error("❌ Allocation failed — no feasible solution found.")
                st.stop()

            st.session_state['allocation_result'] = result
            st.session_state['fixtures'] = fixtures

        except Exception as e:
            st.error(f"❌ Allocation error: {str(e)}")
            st.stop()

    # SUCCESS
    st.markdown('<div class="success-box">', unsafe_allow_html=True)
    st.success("🎉 Allocation Complete!")
    st.markdown('</div>', unsafe_allow_html=True)

    result = st.session_state['allocation_result']

    # Summary metrics
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        allocated = len(result)
        total = len(fixtures)
        st.metric("Allocated", f"{allocated}/{total}", f"{(allocated/total)*100:.1f}%")

    with col2:
        tm = result['matched_pref_time'].sum()
        st.metric("Time Matches", f"{tm}/{allocated}", f"{(tm/allocated)*100:.1f}%")

    with col3:
        if 'is_cup' in result.columns:
            c = result['is_cup'].sum()
        else:
            c = result['fixture_id'].str.contains("cup", case=False, na=False).sum()
        st.metric("Cup Fixtures", c)

    with col4:
        st.metric("Match Days", result['date'].nunique())

    st.markdown("### Allocation Preview")
    st.dataframe(
        result[['team', 'date', 'time', 'pitch', 'age_group']].head(20),
        use_container_width=True
    )


    # ----------------------------------------
    # DOWNLOADS
    # ----------------------------------------

    st.markdown("### Download Schedule")

    col1, col2, col3 = st.columns(3)

    # CSV
    with col1:
        csv_buffer = io.StringIO()
        result.to_csv(csv_buffer, index=False)
        st.download_button(
            "📄 Download CSV",
            csv_buffer.getvalue(),
            file_name=f"pitch_alloc_{datetime.now():%Y%m%d}.csv",
            mime="text/csv",
            use_container_width=True
        )

    # Excel
    with col2:
        excel_file = f"temp_{datetime.now():%Y%m%d%H%M%S}.xlsx"
        generate_excel_schedule(result, fixtures, excel_file)

        with open(excel_file, "rb") as f:
            st.download_button(
                "📊 Download Excel",
                f.read(),
                file_name=f"pitch_schedule_{datetime.now():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        os.remove(excel_file)

    # HTML
    with col3:
        html_file = "temp_schedule.html"
        generate_html_schedule(result, fixtures, html_file)

        with open(html_file, "r", encoding="utf-8") as f:
            st.download_button(
                "🌐 Download HTML",
                f.read(),
                file_name=f"pitch_schedule_{datetime.now():%Y%m%d}.html",
                mime="text/html",
                use_container_width=True
            )

        os.remove(html_file)


# ----------------------------------------
# ANALYTICS TAB
# ----------------------------------------

st.markdown("---")
st.markdown("## 📊 Allocation Analytics")

if "allocation_result" in st.session_state:
    result = st.session_state['allocation_result']

    st.markdown("### Fixtures by Date")
    st.bar_chart(result.groupby("date").size())

    st.markdown("### Pitch Utilisation")
    st.bar_chart(result.groupby("pitch").size())

    st.markdown("### Fixtures by Age Group")
    st.bar_chart(result.groupby("age_group").size())

    st.markdown("### Time Slot Distribution")
    st.bar_chart(result.groupby("time").size())

else:
    st.info("Run an allocation to see analytics.")


# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; font-size: 0.9rem;'>
    <p>Cranleigh FC Pitch Allocator | Version 52 | Built with Streamlit</p>
</div>
""", unsafe_allow_html=True)
