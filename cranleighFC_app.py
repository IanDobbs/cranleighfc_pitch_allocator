import streamlit as st
import pandas as pd
import io
from datetime import datetime
import sys
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

# Sidebar
with st.sidebar:
    st.image("https://via.placeholder.com/200x100/1e3a8a/ffffff?text=Cranleigh+FC", width='stretch')
    st.markdown("### About")
    st.info("""
    This tool automatically assigns home fixtures to available pitches at:
    - Snoxhall Fields (P1-P9)
    - Cranleigh Cricket Club (CCC1-4)
    - Glebelands 3G (G1-G2)
    """)
    
    st.markdown("### Quick Guide")
    st.markdown("""
    1. Upload fixture CSV
    2. Review and adjust settings
    3. Click 'Allocate Pitches'
    4. Download schedule
    """)

# Main content
tab1, tab2, tab3 = st.tabs(["📤 Upload & Allocate", "⚙️ Settings", "📊 Analytics"])

with tab1:
    st.markdown("## Upload Fixtures")
    
    # File upload
    uploaded_file = st.file_uploader(
        "Upload cranleigh_home_fixtures.csv",
        type=['csv'],
        help="CSV file with columns: match_date, match_time, home_team_clean, prefix, away_team, etc."
    )
    
    if uploaded_file is not None:
        # Save uploaded file temporarily
        temp_file = "temp_fixtures.csv"
        with open(temp_file, "wb") as f:
            f.write(uploaded_file.getvalue())
        
        # Load and preview
        try:
            df = pd.read_csv(temp_file)
            
            st.markdown("### Preview Uploaded Fixtures")
            st.dataframe(df.head(10), width='stretch')
            
            # Statistics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Fixtures", len(df))
            with col2:
                cup_count = df['prefix'].str.contains('cup', case=False, na=False).sum()
                st.metric("Cup Fixtures", cup_count)
            with col3:
                dates = df['match_date'].nunique()
                st.metric("Match Days", dates)
            with col4:
                teams = df['home_team_clean'].nunique()
                st.metric("Teams", teams)
            
            # Validation
            st.markdown("### Validation")
            errors = []
            
            # Check required columns
            required_cols = ['match_date', 'match_time', 'home_team_clean']
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                errors.append(f"Missing columns: {', '.join(missing_cols)}")
            
            # Check for unknown teams
            if 'home_team_clean' in df.columns:
                unknown_teams = [team for team in df['home_team_clean'].unique() 
                                if team not in valid_teams]
                if unknown_teams:
                    errors.append(f"Unknown teams: {', '.join(unknown_teams[:5])}")
            
            if errors:
                st.markdown('<div class="warning-box">', unsafe_allow_html=True)
                st.warning("⚠️ Validation Warnings")
                for error in errors:
                    st.markdown(f"- {error}")
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="success-box">', unsafe_allow_html=True)
                st.success("✅ All validations passed!")
                st.markdown('</div>', unsafe_allow_html=True)
            
            # Allocation button
            st.markdown("---")
            
            col1, col2, col3 = st.columns([1, 1, 2])
            with col1:
                timeout = st.number_input("Solver Timeout (seconds)", min_value=5, max_value=120, value=30)
            with col2:
                st.markdown("<br>", unsafe_allow_html=True)
                allocate_button = st.button("🚀 Allocate Pitches", width='stretch', type="primary")
            
            # Run allocation
            if allocate_button:
                with st.spinner("🔄 Allocating pitches... This may take 30-60 seconds..."):
                    try:
                        # Run allocation
                        fixtures, slots_by_date = load_and_validate_fixtures(temp_file)
                        result = solve_allocation(fixtures, slots_by_date, timeout=timeout)
                        
                        if result is not None and len(result) > 0:
                            # Store in session state
                            st.session_state['allocation_result'] = result
                            st.session_state['fixtures'] = fixtures
                            
                            # Success message
                            st.markdown('<div class="success-box">', unsafe_allow_html=True)
                            st.success("🎉 Allocation Complete!")
                            
                            # Summary metrics
                            col1, col2, col3, col4 = st.columns(4)
                            with col1:
                                allocated = len(result)
                                total = len(fixtures)
                                pct = (allocated/total)*100
                                st.metric("Allocated", f"{allocated}/{total}", f"{pct:.1f}%")
                            with col2:
                                time_matches = result['matched_pref_time'].sum()
                                pct = (time_matches/allocated)*100
                                st.metric("Time Matches", f"{time_matches}/{allocated}", f"{pct:.1f}%")
                            # After "Summary metrics" section, replace the Cup fixtures metric:
                            with col3:
                                # ✅ FIX: Check is_cup column instead of fixture_id
                                if 'is_cup' in result.columns:
                                    cup_count = result['is_cup'].sum()
                                else:
                                    # Fallback: check prefix in fixture_id
                                    cup_count = result['fixture_id'].str.contains('Cup', case=False, na=False).sum()
                                st.metric("Cup Fixtures", cup_count)
                            with col4:
                                dates = result['date'].nunique()
                                st.metric("Match Days", dates)
                            
                            st.markdown('</div>', unsafe_allow_html=True)
                            
                            # Show allocation preview
                            st.markdown("### Allocation Preview")
                            st.dataframe(
                                result[['team', 'date', 'time', 'pitch', 'age_group']].head(20),
                                width='stretch'
                            )
                            
                            # ✅ NEW: Cup Fixture Analysis Section
                            if 'is_cup' in result.columns:
                                cup_fixtures = result[result['is_cup'] == True]
                                
                                if len(cup_fixtures) > 0:
                                    st.markdown("### 🏆 Cup Fixture Allocations")
                                    
                                    col1, col2, col3 = st.columns(3)
                                    
                                    with col1:
                                        st.metric("Total Cup Fixtures", len(cup_fixtures))
                                    
                                    with col2:
                                        cup_0930 = cup_fixtures[cup_fixtures['time'] == '09:30']
                                        pct = (len(cup_0930) / len(cup_fixtures)) * 100
                                        st.metric("At 09:30 (Preferred)", f"{len(cup_0930)}/{len(cup_fixtures)}", f"{pct:.0f}%")
                                    
                                    with col3:
                                        cup_other = cup_fixtures[cup_fixtures['time'] != '09:30']
                                        st.metric("At Other Times", len(cup_other))
                                    
                                    # Show cup fixtures at 09:30
                                    if len(cup_0930) > 0:
                                        with st.expander("✅ Cup fixtures at preferred 09:30 slot", expanded=True):
                                            st.dataframe(
                                                cup_0930[['date', 'team', 'pitch', 'age_group']].sort_values('date'),
                                                width='stretch'
                                            )
                                    
                                    # Show cup fixtures at other times
                                    if len(cup_other) > 0:
                                        with st.expander("⚠️ Cup fixtures at other times (capacity constraints)"):
                                            st.dataframe(
                                                cup_other[['date', 'team', 'time', 'pitch', 'age_group']].sort_values('date'),
                                                width='stretch'
                                            )
                            
                            # Download buttons
                            st.markdown("### Download Schedule")
                            
                            col1, col2, col3 = st.columns(3)
                            # ... rest of download code
                            
                            # CSV download
                            with col1:
                                csv_buffer = io.StringIO()
                                result.to_csv(csv_buffer, index=False)
                                st.download_button(
                                    label="📄 Download CSV",
                                    data=csv_buffer.getvalue(),
                                    file_name=f"pitch_allocations_{datetime.now().strftime('%Y%m%d')}.csv",
                                    mime="text/csv",
                                    width='stretch'
                                )
                            
                            # Excel download
                            with col2:
                                excel_file = f"temp_schedule_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
                                generate_excel_schedule(result, fixtures, excel_file)
                                with open(excel_file, 'rb') as f:
                                    st.download_button(
                                        label="📊 Download Excel",
                                        data=f.read(),
                                        file_name=f"pitch_schedule_{datetime.now().strftime('%Y%m%d')}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        width='stretch'
                                    )
                                os.remove(excel_file)
                            
                            # HTML download
                            with col3:
                                html_file = "temp_schedule.html"
                                generate_html_schedule(result, fixtures, html_file)
                                with open(html_file, 'r', encoding='utf-8') as f:
                                    st.download_button(
                                        label="🌐 Download HTML",
                                        data=f.read(),
                                        file_name=f"pitch_schedule_{datetime.now().strftime('%Y%m%d')}.html",
                                        mime="text/html",
                                        width='stretch'
                                    )
                                os.remove(html_file)
                        
                        else:
                            st.error("❌ Allocation failed - no solution found")
                            st.markdown("""
                            **Possible reasons:**
                            - Too many fixtures for available capacity
                            - Conflicting constraints
                            - Try increasing the timeout or reducing fixture count
                            """)
                    
                    except Exception as e:
                        st.error(f"❌ Error during allocation: {str(e)}")
                        import traceback
                        with st.expander("Show error details"):
                            st.code(traceback.format_exc())
                
                # Clean up
                if os.path.exists(temp_file):
                    os.remove(temp_file)
        
        except Exception as e:
            st.error(f"❌ Error reading CSV: {str(e)}")
            st.markdown("Please ensure your CSV has the correct format and column names.")

with tab2:
    st.markdown("## Allocation Settings")
    
    st.markdown("### Pitch Configuration")
    
    # Show current pitches
    pitch_df = pd.DataFrame([
        {
            'Pitch': name,
            'Format': info['format'],
            'Lights': '✓' if info['lights'] else '✗',
            'Location': info['location'],
            'Priority': info['priority']
        }
        for name, info in pitches.items()
    ])
    
    st.dataframe(pitch_df, width='stretch')
    
    st.markdown("### Time Slots")
    st.info("""
    **Current time slots:**
    - 11v11 pitches: 09:30, 11:00, 14:00
    - 9v9/7v7/5v5 pitches: 09:30, 11:00
    
    To modify, edit the `pitch_allocator_v52.py` file.
    """)
    
    st.markdown("### Allocation Priorities")
    st.info("""
    **Current priorities:**
    1. Cup fixtures → 09:30 kickoff (very high priority)
    2. Senior teams → P6 11v11 (Seniors) pitch
    3. U13/U14 → P3 11v11 (Middle) small pitch
    4. Minimize back-to-back matches on same pitch
    5. Glebelands 3G → overflow only
    """)

with tab3:
    st.markdown("## Analytics Dashboard")
    
    if 'allocation_result' in st.session_state:
        result = st.session_state['allocation_result']
        
        # Allocation rate by date
        st.markdown("### Fixtures by Date")
        date_counts = result.groupby('date').size().reset_index(name='Fixtures')
        st.bar_chart(date_counts.set_index('date'))
        
        # Pitch utilization
        st.markdown("### Pitch Utilization")
        pitch_counts = result.groupby('pitch').size().reset_index(name='Fixtures')
        pitch_counts = pitch_counts.sort_values('Fixtures', ascending=False)
        st.bar_chart(pitch_counts.set_index('pitch'))
        
        # Age group distribution
        st.markdown("### Fixtures by Age Group")
        age_counts = result.groupby('age_group').size().reset_index(name='Fixtures')
        age_counts = age_counts.sort_values('Fixtures', ascending=False)
        st.bar_chart(age_counts.set_index('age_group'))
        
        # Time slot distribution
        st.markdown("### Fixtures by Time Slot")
        time_counts = result.groupby('time').size().reset_index(name='Fixtures')
        st.bar_chart(time_counts.set_index('time'))
    
    else:
        st.info("📊 Run an allocation to see analytics")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; font-size: 0.9rem;'>
    <p>Cranleigh FC Pitch Allocator | Version 52 | Built with Streamlit</p>
</div>
""", unsafe_allow_html=True)