# ===============================================
# Production Scheduling App - Streamlit Version
# ===============================================
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io

# ================== CONFIG =====================
class Config:
    WORKING_HOURS_PER_DAY = 12
    WORK_START_HOUR = 8
    WORK_END_HOUR = 20
    WORKING_DAYS = [6,0,1,2,3]  # Sunday-Thursday

# ================= HELPERS ====================
def is_working_day(date):
    if pd.isna(date): return False
    return date.weekday() in Config.WORKING_DAYS

def get_next_working_day(date):
    if pd.isna(date): return None
    next_day = date + timedelta(days=1)
    next_day = next_day.replace(hour=0, minute=0, second=0, microsecond=0)
    while not is_working_day(next_day):
        next_day += timedelta(days=1)
    return next_day

def calculate_completion_from_hours(start_datetime, lead_time_hours):
    if pd.isna(start_datetime) or pd.isna(lead_time_hours) or lead_time_hours <=0:
        return start_datetime if not pd.isna(start_datetime) else None
    lead_time_hours = float(lead_time_hours)

    # Adjust start_datetime within work hours
    if start_datetime.hour < Config.WORK_START_HOUR:
        start_datetime = start_datetime.replace(hour=Config.WORK_START_HOUR, minute=0)
    elif start_datetime.hour >= Config.WORK_END_HOUR:
        next_day = get_next_working_day(start_datetime)
        if next_day:
            start_datetime = next_day.replace(hour=Config.WORK_START_HOUR, minute=0)

    current_datetime = start_datetime
    remaining_hours = lead_time_hours

    while remaining_hours > 0:
        end_of_day = current_datetime.replace(hour=Config.WORK_END_HOUR, minute=0)
        hours_left_today = (end_of_day - current_datetime).total_seconds()/3600

        if remaining_hours <= hours_left_today:
            return current_datetime + timedelta(hours=remaining_hours)
        else:
            remaining_hours -= hours_left_today
            next_day = get_next_working_day(current_datetime)
            if next_day:
                current_datetime = next_day.replace(hour=Config.WORK_START_HOUR, minute=0)
            else:
                return current_datetime
    return current_datetime

# ================ CALCULATION LOGIC =================
def calculate_completion_dates(df):
    df = df.copy()
    activities = [
        ('Internal Testing','Internal Testing Completion'),
        ('Blasting','Blasting Completion'),
        ('Shell Test','Shell Test Completion'),
        ('Shell Test TPI review','Shell Test TPI review Completion'),
        ('Painting','Painting Completion'),
        ('Paint Curing','Paint Curing Completion'),
        ('Accessory Mounting & Tubing','Accessory Mounting Completion'),
        ('Calibration & Testing','Calibration Completion'),
        ('Document handover & verification by QC (System)','QC Documentation Completion'),
        ('TPI review','TPI Review Completion'),
        ('Packing','Packing Completion')
    ]

    for _, comp_col in activities:
        df[comp_col] = None

    for idx, row in df.iterrows():
        start_dt = row.get('Start Date', pd.Timestamp.today())
        if isinstance(start_dt, str):
            start_dt = pd.to_datetime(start_dt, errors='coerce')
        current_start = start_dt

        for lead_col, comp_col in activities:
            lead_time = row.get(lead_col)
            completion = calculate_completion_from_hours(current_start, lead_time if pd.notna(lead_time) else 0)
            df.at[idx, comp_col] = completion
            if completion: current_start = completion

    return df

def calculate_milestone_dates(df):
    df = df.copy()
    for idx, row in df.iterrows():
        blast = row.get('Blasting Completion')
        shell_tpi = row.get('Shell Test TPI review Completion')
        qc = row.get('QC Documentation Completion')
        tpi_hours = row.get('TPI review')

        if pd.notna(blast): df.at[idx, 'TPI witness Start Date'] = blast
        if pd.notna(shell_tpi): df.at[idx, 'TPI witness Finish Date'] = shell_tpi

        # Final TPI
        if pd.notna(qc):
            df.at[idx, 'Final TPI witness Start Date'] = qc
            finish = calculate_completion_from_hours(qc, tpi_hours if pd.notna(tpi_hours) else 0)
            df.at[idx, 'Final TPI witness Finish Date'] = finish

        # Final Packing TPI
        pack = row.get('Packing Completion')
        if pd.notna(pack):
            df.at[idx, 'Final Packing TPI witness Start Date'] = pack
            finish_pack = calculate_completion_from_hours(pack, tpi_hours if pd.notna(tpi_hours) else 0)
            df.at[idx, 'Final Packing TPI witness Finish Date'] = finish_pack

        # Dispatch
        df.at[idx, 'Dispatch Date'] = df.at[idx, 'Final Packing TPI witness Finish Date']

    return df

# =================== EXCEL EXPORT ===================
def export_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Production Schedule')
        workbook  = writer.book
        worksheet = writer.sheets['Production Schedule']

        # Example styling: Header bold
        header_format = workbook.add_format({'bold': True, 'bg_color':'#D7E4BC'})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)

    output.seek(0)
    return output

# ===================== STREAMLIT UI =====================
st.set_page_config(page_title="Production Scheduler", layout="wide")
st.title("Production Scheduler & Milestone Calculator 🚀")

uploaded_file = st.file_uploader("Upload Excel file with lead times & start dates", type=["xlsx"])
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.subheader("Preview of uploaded data")
        st.dataframe(df.head())

        if st.button("Calculate Schedule"):
            df1 = calculate_completion_dates(df)
            df2 = calculate_milestone_dates(df1)
            st.subheader("Calculated Production Schedule")
            st.dataframe(df2)

            excel_file = export_to_excel(df2)
            st.download_button(
                label="📥 Download Excel",
                data=excel_file,
                file_name="Production_Schedule.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Error reading Excel: {e}")