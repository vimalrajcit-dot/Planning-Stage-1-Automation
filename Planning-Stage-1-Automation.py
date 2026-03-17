"""
Production Planning Schedule Generator - Streamlit Deployment Version
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import warnings
warnings.filterwarnings('ignore')

from openpyxl.styles import PatternFill, Font

# ============================================================================

class Config:
    WORKING_HOURS_PER_DAY = 12
    WORK_START_HOUR = 8
    WORK_END_HOUR = 20
    WORKING_DAYS = [6,0,1,2,3]

    LEAD_TIME_COLUMNS = [
        'Internal Testing','Blasting','Shell Test','Shell Test TPI review','Painting','Paint Curing',
        'Accessory Mounting & Tubing','Calibration & Testing','Document handover & verification by QC (System)',
        'TPI review','Packing'
    ]

    SCHEDULE_COLUMN_ORDER = [
        'Sales Order No','Customer Name','Tag No.','Work Order No.','Inspection Y/N',
        'Valve size  (inch)','Rating','Body material ',
        'Internal Testing','Blasting','Shell Test','Shell Test TPI review','Painting','Paint Curing',
        'Accessory Mounting & Tubing','Calibration & Testing','Document handover & verification by QC (System)',
        'TPI review','Packing','Start Date',
        'Internal Testing Completion','Blasting Completion','Shell Test Completion','Shell Test TPI review Completion','TPI witness Start Date','TPI witness Finish Date',
        'Painting Completion','Paint Curing Completion','Accessory Mounting Completion','Calibration Completion',
        'QC Documentation Completion','Final TPI witness Start Date','Final TPI witness Finish Date',
        'Packing Completion','Final Packing TPI witness Start Date','Final Packing TPI witness Finish Date',
        'Dispatch Date'
    ]

# ============================================================================

def add_display_columns(df):
    return df

def is_working_day(date):
    if pd.isna(date): return False
    return date.weekday() in Config.WORKING_DAYS

def get_next_working_day(date):
    if pd.isna(date): return None
    next_day = date + timedelta(days=1)
    next_day = next_day.replace(hour=0,minute=0,second=0,microsecond=0)
    while not is_working_day(next_day):
        next_day += timedelta(days=1)
    return next_day

def calculate_completion_from_hours(start_datetime, lead_time_hours):
    if pd.isna(start_datetime) or pd.isna(lead_time_hours) or lead_time_hours <=0:
        return start_datetime
    lead_time_hours = float(lead_time_hours)

    if start_datetime.hour < Config.WORK_START_HOUR:
        start_datetime = start_datetime.replace(hour=Config.WORK_START_HOUR)
    elif start_datetime.hour >= Config.WORK_END_HOUR:
        start_datetime = get_next_working_day(start_datetime).replace(hour=Config.WORK_START_HOUR)

    current_datetime = start_datetime
    remaining_hours = lead_time_hours

    while remaining_hours > 0:
        end_of_day = current_datetime.replace(hour=Config.WORK_END_HOUR)
        hours_left_today = (end_of_day - current_datetime).total_seconds()/3600

        if remaining_hours <= hours_left_today:
            return current_datetime + timedelta(hours=remaining_hours)
        else:
            remaining_hours -= hours_left_today
            current_datetime = get_next_working_day(current_datetime).replace(hour=Config.WORK_START_HOUR)

    return current_datetime

# ============================================================================

def lookup_lead_times(user_df, master_df):
    user_df['Valve size  (inch)'] = user_df['Valve size  (inch)'].astype(str).str.strip()
    user_df['Rating'] = user_df['Rating'].astype(str).str.strip()
    master_df['Valve size  (inch)'] = master_df['Valve size  (inch)'].astype(str).str.strip()
    master_df['Rating'] = master_df['Rating'].astype(str).str.strip()

    user_keys = user_df['Valve size  (inch)'] + '|' + user_df['Rating']
    master_keys = master_df['Valve size  (inch)'] + '|' + master_df['Rating']
    lookup_dict = {key:idx for idx,key in enumerate(master_keys)}

    for col in Config.LEAD_TIME_COLUMNS:
        user_df[col] = None

    for idx,user_key in enumerate(user_keys):
        if user_key in lookup_dict:
            master_idx = lookup_dict[user_key]
            for col in Config.LEAD_TIME_COLUMNS:
                val = master_df.iloc[master_idx][col]
                if pd.notna(val):
                    user_df.at[idx,col] = float(val)

    return user_df

# ============================================================================

def calculate_completion_dates(df):
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

    for _,comp_col in activities:
        df[comp_col] = None

    for idx,row in df.iterrows():
        if pd.isna(row.get('Start Date')): continue
        current_start = row['Start Date']

        for lead_col, comp_col in activities:
            lead_time = row.get(lead_col)
            completion = calculate_completion_from_hours(current_start, lead_time)
            df.at[idx, comp_col] = completion
            current_start = completion

    return df

# ============================================================================

def calculate_milestone_dates(df):
    for idx,row in df.iterrows():
        df.at[idx,'Dispatch Date'] = row.get('Packing Completion')
    return df

# ============================================================================

def save_output(df, file_path):
    dir_name = os.path.dirname(file_path)
    if dir_name:
        os.makedirs(dir_name, exist_ok=True)

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Production Schedule", index=False)

# ============================================================================

# ====================== STREAMLIT UI ==========================
st.set_page_config(layout="wide")
st.title("📈 Production Planning Scheduler")

master_file = st.file_uploader("Upload Master File", type=['xlsx'])
user_file = st.file_uploader("Upload Production Plan", type=['xlsx'])

if master_file and user_file:
    with st.spinner("Processing..."):
        master_df = pd.read_excel(master_file)
        user_df = pd.read_excel(user_file)

        if 'Start Date' in user_df.columns:
            user_df['Start Date'] = pd.to_datetime(user_df['Start Date'], dayfirst=True)

        df = lookup_lead_times(user_df.copy(), master_df)
        df = calculate_completion_dates(df)
        df = add_display_columns(df)
        df = calculate_milestone_dates(df)

        output_file = "output.xlsx"
        save_output(df, output_file)

    st.success("✅ Done!")

    st.subheader("Preview")
    st.dataframe(df.head(20))

    with open(output_file, "rb") as f:
        st.download_button("⬇ Download Excel", f, file_name="Production_Output.xlsx")