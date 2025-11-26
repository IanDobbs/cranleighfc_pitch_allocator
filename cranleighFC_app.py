import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta
import os
import requests

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
    page_icon="logo.png",
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
    .weather-card {
        padding: 1.5rem;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 12px;
        margin: 1rem 0;
    }
    .weather-severe {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
    }
    .weather-good {
        background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
    }
</style>
""", unsafe_allow_html=True)

# ----------------------------------------
# WEATHER FUNCTIONS
# ----------------------------------------

@st.cache_data(ttl=3600)  # Cache for 1 hour
def get_weather_forecast(latitude=51.14209, longitude=-0.48374, days=7):
    """
    Fetch weather forecast from Open-Meteo API
    Default coordinates are for Cranleigh, England
    """
    try:
        url = "https://api.open-meteo.com/v1/forecast"
        params = {
            "latitude": latitude,
            "longitude": longitude,
            "hourly": "temperature_2m,precipitation,precipitation_probability,weather_code,wind_speed_10m",
            "daily": "weather_code,temperature_2m_max,temperature_2m_min,precipitation_sum,precipitation_probability_max,wind_speed_10m_max",
            "timezone": "Europe/London",
            "forecast_days": days
        }
        
        response = requests.get(url, params=params)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        st.error(f"Weather API error: {str(e)}")
        return None

def get_weather_code_description(code):
    """Convert weather code to description"""
    weather_codes = {
        0: ("Clear sky", "☀️"),
        1: ("Mainly clear", "🌤️"),
        2: ("Partly cloudy", "⛅"),
        3: ("Overcast", "☁️"),
        45: ("Foggy", "🌫️"),
        48: ("Foggy", "🌫️"),
        51: ("Light drizzle", "🌦️"),
        53: ("Moderate drizzle", "🌦️"),
        55: ("Heavy drizzle", "🌧️"),
        61: ("Light rain", "🌧️"),
        63: ("Moderate rain", "🌧️"),
        65: ("Heavy rain", "⛈️"),
        71: ("Light snow", "🌨️"),
        73: ("Moderate snow", "🌨️"),
        75: ("Heavy snow", "❄️"),
        80: ("Light showers", "🌦️"),
        81: ("Moderate showers", "🌧️"),
        82: ("Heavy showers", "⛈️"),
        95: ("Thunderstorm", "⛈️"),
    }
    return weather_codes.get(code, ("Unknown", "🌡️"))

def assess_playing_conditions(temp, precip_prob, wind_speed, weather_code):
    """Assess if conditions are suitable for football"""
    issues = []
    
    if precip_prob > 70:
        issues.append("High rain probability")
    if wind_speed > 35:
        issues.append("High winds")
    if temp < 2:
        issues.append("Near freezing")
    if weather_code in [71, 73, 75]:
        issues.append("Snow expected")
    if weather_code == 95:
        issues.append("Thunderstorm risk")
    
    if not issues:
        return "✅ Good", "weather-good"
    elif len(issues) == 1:
        return f"⚠️ Caution: {issues[0]}", "weather-card"
    else:
        return f"❌ Poor: {', '.join(issues)}", "weather-severe"

# ----------------------------------------
# HEADER & WEATHER WIDGET
# ----------------------------------------

st.image("logo.png", width=60)
# st.markdown('<p class="main-header">⚽ Cranleigh FC Pitch Allocator</p>', unsafe_allow_html=True)
st.markdown("Automated pitch allocation system for home fixtures")

# Weather Widget in Sidebar
with st.sidebar:
    st.markdown("## 🌤️ Local Weather")
    
    # Location settings
    with st.expander("📍 Location Settings"):
        lat = st.number_input("Latitude", value=51.14209, format="%.5f", help="Cranleigh, England")
        lon = st.number_input("Longitude", value=-0.48374, format="%.5f")
        forecast_days = st.slider("Forecast Days", 1, 7, 7)
    
    if st.button("🔄 Refresh Weather", width='stretch'):
        st.cache_data.clear()
    
    # Fetch weather
    weather_data = get_weather_forecast(lat, lon, forecast_days)
    
    if weather_data:
        daily = weather_data.get('daily', {})
        
        # Today's summary
        if daily:
            today_code = daily['weather_code'][0]
            today_desc, today_emoji = get_weather_code_description(today_code)
            today_temp_max = daily['temperature_2m_max'][0]
            today_temp_min = daily['temperature_2m_min'][0]
            today_precip = daily['precipitation_probability_max'][0]
            today_wind = daily['wind_speed_10m_max'][0]
            
            st.markdown(f"### Today {today_emoji}")
            st.markdown(f"**{today_desc}**")
            st.markdown(f"🌡️ {today_temp_min:.0f}°C - {today_temp_max:.0f}°C")
            st.markdown(f"💧 Rain: {today_precip}%")
            st.markdown(f"💨 Wind: {today_wind:.0f} km/h")
            
            # Playing conditions
            condition, css_class = assess_playing_conditions(
                today_temp_min, today_precip, today_wind, today_code
            )
            st.markdown(f"**{condition}**")
            
            st.markdown("---")
            
            # 7-day forecast
            st.markdown("### 7-Day Forecast")
            for i in range(min(7, len(daily['time']))):
                date = daily['time'][i]
                code = daily['weather_code'][i]
                desc, emoji = get_weather_code_description(code)
                temp_max = daily['temperature_2m_max'][i]
                temp_min = daily['temperature_2m_min'][i]
                precip = daily['precipitation_probability_max'][i]
                
                # Format date
                date_obj = datetime.fromisoformat(date)
                day_name = date_obj.strftime("%a %d/%m")
                
                st.markdown(f"**{day_name}** {emoji}")
                st.markdown(f"{temp_min:.0f}°-{temp_max:.0f}°C | 💧{precip}%")
                
                if i < 6:
                    st.markdown("---")

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

    st.markdown("## Loaded Published Home Fixtures on FA Full-Time")
    st.info(f"Loaded automatically from `{DEFAULT_FILE}`")
    st.dataframe(df.head(10), width='stretch')

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
# WEATHER-AWARE FIXTURE ANALYSIS
# ----------------------------------------

if weather_data and 'match_date' in df.columns:
    st.markdown("## 🌦️ Weather Impact on Upcoming Fixtures")
    
    # Parse fixture dates
    df['parsed_date'] = pd.to_datetime(df['match_date'], errors='coerce')
    
    # Get next 7 days of fixtures
    today = datetime.now().date()
    next_week = today + timedelta(days=7)
    
    upcoming = df[
        (df['parsed_date'].dt.date >= today) & 
        (df['parsed_date'].dt.date <= next_week)
    ].copy()
    
    if len(upcoming) > 0:
        st.info(f"Found {len(upcoming)} fixtures in the next 7 days")
        
        # Match weather data to fixtures
        daily = weather_data.get('daily', {})
        weather_lookup = {}
        
        for i, date_str in enumerate(daily.get('time', [])):
            date_obj = datetime.fromisoformat(date_str).date()
            weather_lookup[date_obj] = {
                'code': daily['weather_code'][i],
                'temp_max': daily['temperature_2m_max'][i],
                'temp_min': daily['temperature_2m_min'][i],
                'precip_prob': daily['precipitation_probability_max'][i],
                'wind': daily['wind_speed_10m_max'][i]
            }
        
        # Show fixtures with weather
        for _, fixture in upcoming.iterrows():
            fixture_date = fixture['parsed_date'].date()
            
            if fixture_date in weather_lookup:
                w = weather_lookup[fixture_date]
                desc, emoji = get_weather_code_description(w['code'])
                condition, css_class = assess_playing_conditions(
                    w['temp_min'], w['precip_prob'], w['wind'], w['code']
                )
                
                col1, col2, col3 = st.columns([2, 2, 1])
                
                with col1:
                    st.markdown(f"**{fixture['home_team_clean']}**")
                    st.markdown(f"{fixture_date.strftime('%A, %d %B')}")
                
                with col2:
                    st.markdown(f"{emoji} {desc}")
                    st.markdown(f"🌡️ {w['temp_min']:.0f}-{w['temp_max']:.0f}°C | 💧 {w['precip_prob']}% | 💨 {w['wind']:.0f} km/h")
                
                with col3:
                    st.markdown(condition)
                
                st.markdown("---")
    else:
        st.info("No fixtures scheduled in the next 7 days")


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
    allocate_button = st.button("🚀 Allocate Pitches", width='stretch')


# ----------------------------------------
# RUN ALLOCATION
# ----------------------------------------

if allocate_button:
    with st.spinner("Allocating pitches… this may take 30–60 seconds...."):

        try:
          fixtures, slots_by_date, removed_duplicates = load_and_validate_fixtures(DEFAULT_FILE)
          result = solve_allocation(fixtures, slots_by_date, timeout=timeout)

          if result is None or len(result) == 0:
              st.error("❌ Allocation failed – no feasible solution found.")
              st.stop()

          st.session_state['allocation_result'] = result
          st.session_state['fixtures'] = fixtures
          st.session_state['removed_duplicates'] = removed_duplicates
            
        except Exception as e:
            st.error(f"❌ Allocation error: {str(e)}")
            import traceback
            traceback.print_exc()
            st.stop()

    # SUCCESS
    st.markdown('<div class="success-box">', unsafe_allow_html=True)
    st.success("🎉 Allocation Complete!")
    st.markdown('</div>', unsafe_allow_html=True)
# ----------------------------------------
# DISPLAY REMOVED DUPLICATES (always visible after run)
# ----------------------------------------

removed = st.session_state.get("removed_duplicates")

st.markdown("### 🔁 Duplicate Fixtures Removed")

if removed is not None and not removed.empty:

    st.warning(f"⚠️ {len(removed)} duplicate fixtures were removed!")

    summary = (
        removed.groupby(['team_name', 'date'])
        .size()
        .reset_index(name='Removed Count')
        .sort_values(['date', 'team_name'])
    )
    st.dataframe(summary)

    # Checkbox works now because this code always runs
    if st.checkbox("Show full removed fixture details"):
        st.dataframe(removed.reset_index(drop=True))

else:
    st.success("No duplicate fixtures were removed.")


result = st.session_state.get("allocation_result")
fixtures = st.session_state.get("fixtures")

if result is not None and not result.empty:
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
        width='stretch'
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
            width='stretch'
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
                width='stretch'
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
                width='stretch'
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
    <p>Cranleigh FC Pitch Allocator | Version 53 | Built with Streamlit</p>
    <p>Weather data from Open-Meteo.com</p>
</div>
""", unsafe_allow_html=True)