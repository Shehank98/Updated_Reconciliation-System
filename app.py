import streamlit as st
import pandas as pd
from io import BytesIO
from fuzzywuzzy import fuzz
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
import re

def extract_duration_from_theme(theme_name):
    """Extract duration in seconds from theme names like 'Cash Bonanza_2025 (10)(Sin)'"""
    if pd.isna(theme_name):
        return None

    # Look for patterns like (10), (15), (30) in theme names
    duration_match = re.search(r'\((\d+)\)', str(theme_name))
    if duration_match:
        return int(duration_match.group(1))
    return None

def preprocess_time_format(time_value):
    """Extract and standardize time format from various inputs."""
    try:
        if isinstance(time_value, str):
            time_pattern = re.search(r'(\d{1,2}:\d{2}(?::\d{2})?)', str(time_value))
            if time_pattern:
                time_str = time_pattern.group(1)
                if time_str.count(':') == 1:
                    time_str += ':00'
                return time_str
            else:
                return '00:00:00'
        elif hasattr(time_value, 'strftime'):
            return time_value.strftime('%H:%M:%S')
        return '00:00:00'
    except Exception:
        return '00:00:00'

def time_to_seconds(time_str):
    """Convert time string (HH:MM:SS) to seconds."""
    try:
        if pd.isna(time_str):
            return 0

        time_str = str(time_str)
        time_parts = time_str.split(':')

        if len(time_parts) >= 3:
            return int(time_parts[0]) * 3600 + int(time_parts[1]) * 60 + int(time_parts[2])
        elif len(time_parts) == 2:
            return int(time_parts[0]) * 3600 + int(time_parts[1]) * 60
        else:
            return 0
    except (ValueError, IndexError):
        return 0

def normalize_theme_name(theme):
    """Normalize theme name for consistent comparison by removing whitespace and lowercase."""
    if pd.isna(theme):
        return ""
    return str(theme).lower().strip()

def filter_by_channel(df, channel):
    """Filter data by selected channel."""
    if "Channel" not in df.columns:
        st.error("Channel column not found in the dataframe.")
        return df
    return df[df["Channel"] == channel].reset_index(drop=True)

def remove_duplicates(current_df, previous_df):
    """Remove duplicates between current and previous month's data."""
    current_df = current_df.reset_index(drop=True)
    previous_df = previous_df.reset_index(drop=True)
    match_columns = ["Advt_Theme", "Channel", "Program", "Advt_time", "Dd", "Mn", "Yr"]

    for col in match_columns:
        if col not in current_df.columns or col not in previous_df.columns:
            st.warning(f"Column '{col}' missing in one of the dataframes. Skipping duplicate removal.")
            return current_df, 0

    merged_df = pd.merge(
        current_df,
        previous_df,
        on=match_columns,
        how="left",
        indicator=True
    )
    cleaned_df = merged_df[merged_df["_merge"] == "left_only"].drop(columns=["_merge"]).reset_index(drop=True)
    duplicate_count = len(current_df) - len(cleaned_df)

    return cleaned_df, duplicate_count

def clean_column_names(df):
    """Clean up column names by removing _x and _y suffixes from merged dataframes."""
    col_mapping = {}
    seen_base_cols = set()

    for col in df.columns:
        if col.endswith('_x') or col.endswith('_y'):
            base_name = col[:-2]
            if base_name not in seen_base_cols:
                col_mapping[col] = base_name
                seen_base_cols.add(base_name)
        else:
            col_mapping[col] = col

    if 'Advt_time' in df.columns and df.columns.tolist().count('Advt_time') > 1:
        advt_time_indices = [i for i, col in enumerate(df.columns) if col == 'Advt_time']
        for i, idx in enumerate(advt_time_indices[1:], 1):
            original_name = df.columns[idx]
            new_name = f'Advt_time_{i}'
            col_mapping[original_name] = new_name

    return df.rename(columns=col_mapping)

def standardize_dataframe(df, date_col=None, time_col=None, theme_col=None, program_col=None):
    """Standardize DataFrame columns and formats for consistent processing."""
    df = clean_column_names(df)
    result_df = df.copy().reset_index(drop=True)

    rename_dict = {}
    if theme_col and theme_col in df.columns:
        rename_dict[theme_col] = 'Advt_Theme'
    if program_col and program_col in df.columns:
        rename_dict[program_col] = 'Program'
    if time_col and time_col in df.columns:
        rename_dict[time_col] = 'Advt_time'

    if rename_dict:
        result_df = result_df.rename(columns=rename_dict)

    if date_col and date_col in df.columns:
        result_df['Date'] = pd.to_datetime(df[date_col], errors='coerce')
        result_df['Dd'] = result_df['Date'].dt.day
        result_df['Mn'] = result_df['Date'].dt.month
        result_df['Yr'] = result_df['Date'].dt.year

    if 'Advt_time' in result_df.columns:
        result_df['Advt_time'] = result_df['Advt_time'].apply(preprocess_time_format)

    if 'Advt_Theme' in result_df.columns:
        result_df['Normalized_Theme'] = result_df['Advt_Theme'].apply(normalize_theme_name)
        # Extract duration from theme names
        result_df['Theme_Duration'] = result_df['Advt_Theme'].apply(extract_duration_from_theme)

    return result_df

def match_media_watch_with_tc(media_watch_df, tc_df, theme_mapping, ignore_date=False, time_tolerance=30):
    """Match Media Watch data with TC data based on theme mapping."""
    def local_time_to_seconds(time_str):
        try:
            if pd.isna(time_str):
                return 0
            time_str = str(time_str)
            time_parts = time_str.split(':')
            if len(time_parts) >= 3:
                return int(time_parts[0]) * 3600 + int(time_parts[1]) * 60 + int(time_parts[2])
            elif len(time_parts) == 2:
                return int(time_parts[0]) * 3600 + int(time_parts[1]) * 60
            else:
                return 0
        except (ValueError, IndexError):
            return 0
            
    if media_watch_df.empty or tc_df.empty:
        st.error("One or both DataFrames are empty.")
        return pd.DataFrame(), [], pd.DataFrame()

    # Debug TC columns and find/fix Advt_time
    with st.expander("Debug TC Data Columns"):
        st.write(f"TC DataFrame columns: {tc_df.columns.tolist()}")
        time_related_cols = [col for col in tc_df.columns if any(term in str(col).lower() for term in ['time', 'spot', 'air'])]
        st.write(f"Potential time columns: {time_related_cols}")
        
        # Check for Advt_time with exact capitalization
        if 'Advt_time' not in tc_df.columns:
            # Try to find case-insensitive match
            for col in tc_df.columns:
                if col.lower() == 'advt_time':
                    tc_df = tc_df.rename(columns={col: 'Advt_time'})
                    st.success(f"Renamed '{col}' to 'Advt_time'")
                    break
            else:
                # If no exact match found, use the first time-related column
                if time_related_cols:
                    tc_df = tc_df.rename(columns={time_related_cols[0]: 'Advt_time'})
                    st.success(f"Using '{time_related_cols[0]}' as 'Advt_time'")

    media_watch_df = media_watch_df.copy().reset_index(drop=True)
    tc_df = tc_df.copy().reset_index(drop=True)

    with st.expander("Debug DataFrame Info"):
        st.write(f"Media Watch DF shape: {media_watch_df.shape}")
        st.write(f"Media Watch columns: {media_watch_df.columns.tolist()}")
        if 'Advt_time' in media_watch_df.columns:
            st.write(f"First few time values: {media_watch_df['Advt_time'].head(5).values.tolist()}")

    if 'Advt_time' in media_watch_df.columns:
        media_watch_df['Time_Seconds'] = media_watch_df['Advt_time'].astype(str).apply(local_time_to_seconds)
    else:
        st.error("Advt_time column not found in Media Watch data")
        return pd.DataFrame(), [], pd.DataFrame()

    if 'Advt_time' in tc_df.columns:
        tc_df['Time_Seconds'] = tc_df['Advt_time'].astype(str).apply(local_time_to_seconds)
    else:
        st.error("Advt_time column not found in TC data")
        return pd.DataFrame(), [], pd.DataFrame()

    mw_to_tc = {}
    for m in theme_mapping:
        if 'tc_theme' in m and m['tc_theme']:
            mw_theme = normalize_theme_name(m['media_watch_theme'])
            tc_theme = normalize_theme_name(m['tc_theme'])
            mw_to_tc[mw_theme] = tc_theme

    with st.expander("Debug Theme Mappings"):
        st.write("Normalized Theme Mappings:")
        for mw, tc in mw_to_tc.items():
            st.write(f"Media Watch: '{mw}' → TC: '{tc}'")

    progress_bar = st.progress(0)
    total_rows = len(media_watch_df)

    matched_results = []
    matched_mw_indices = []
    time_matched_program_mismatched = []

    filter_stats = {
        "total_media_watch_rows": len(media_watch_df),
        "no_normalized_mapping": 0,
        "no_matches_after_theme_filter": 0,
        "no_matches_after_date_filter": 0,
        "no_matches_after_time_filter": 0,
        "matches_found": 0,
        "failed_program_similarity": 0,
        "program_mismatched_but_time_matched": 0
    }

    for i in range(len(media_watch_df)):
        progress_bar.progress(min(1.0, (i + 1) / total_rows))
        mw_row = media_watch_df.iloc[i]

        if pd.isna(mw_row['Advt_Theme']) or pd.isna(mw_row['Program']) or pd.isna(mw_row['Advt_time']):
            continue

        mw_theme_norm = normalize_theme_name(mw_row['Advt_Theme'])

        if mw_theme_norm not in mw_to_tc:
            filter_stats["no_normalized_mapping"] += 1
            continue

        tc_theme_norm = mw_to_tc[mw_theme_norm]
        tc_matches = tc_df[tc_df['Normalized_Theme'] == tc_theme_norm].copy().reset_index(drop=True)

        if tc_matches.empty:
            filter_stats["no_matches_after_theme_filter"] += 1
            continue

        # Extract duration from theme
        mw_duration = extract_duration_from_theme(mw_row['Advt_Theme'])
            
        # Filter by duration if available
        if mw_duration is not None:
            if 'Theme_Duration' in tc_matches.columns:
                # First try theme duration from extracted theme name 
                duration_matches = tc_matches[tc_matches['Theme_Duration'] == mw_duration]
                if not duration_matches.empty:
                    tc_matches = duration_matches.reset_index(drop=True)
            elif 'Dur' in tc_matches.columns:
                # Then try explicit duration column
                duration_matches = tc_matches[tc_matches['Dur'] == mw_duration]
                if not duration_matches.empty:
                    tc_matches = duration_matches.reset_index(drop=True)

        if not ignore_date:
            date_filter = (
                (tc_matches['Dd'] == mw_row['Dd']) &
                (tc_matches['Mn'] == mw_row['Mn']) &
                (tc_matches['Yr'] == mw_row['Yr'])
            )
            tc_matches = tc_matches[date_filter].reset_index(drop=True)
            
            if tc_matches.empty:
                filter_stats["no_matches_after_date_filter"] += 1
                continue

        mw_time_seconds = mw_row['Time_Seconds']
        tc_matches['Time_Diff'] = abs(tc_matches['Time_Seconds'] - mw_time_seconds)
        
        # Increased time tolerance to 30 seconds (configurable)
        close_time_matches = tc_matches[tc_matches['Time_Diff'] <= time_tolerance].reset_index(drop=True)

        if close_time_matches.empty:
            filter_stats["no_matches_after_time_filter"] += 1
            continue

        close_time_matches = close_time_matches.sort_values('Time_Diff').reset_index(drop=True)
        close_time_matches['Program_Similarity'] = close_time_matches['Program'].astype(str).apply(
            lambda x: fuzz.token_set_ratio(str(mw_row['Program']).lower(), x.lower())
        )

        # Handle multiple potential matches
        if len(close_time_matches) > 1:
            # Sort by program similarity first, then time difference
            close_time_matches = close_time_matches.sort_values(
                ['Program_Similarity', 'Time_Diff'], 
                ascending=[False, True]
            ).reset_index(drop=True)

        if close_time_matches.empty:
            filter_stats["no_matches_after_time_filter"] += 1
            continue
            
        best_match = close_time_matches.iloc[0]
        
        # Check if time matches but program doesn't match well
        if best_match['Program_Similarity'] < 50:
            # This is a time-matched but program-mismatched case
            filter_stats["program_mismatched_but_time_matched"] += 1
            
            # Create a result for program mismatched but time matched cases
            mismatch_row = {
                'Media_Watch_Theme': mw_row['Advt_Theme'],
                'TC_Theme': best_match['Advt_Theme'],
                'Date': f"{mw_row['Dd']}/{mw_row['Mn']}/{mw_row['Yr']}",
                'Dd': mw_row['Dd'],
                'Mn': mw_row['Mn'], 
                'Yr': mw_row['Yr'],
                'Program_LMRB': mw_row['Program'],
                'Program_TC': best_match['Program'],
                'Media_Watch_Time': mw_row['Advt_time'],
                'TC_Time': best_match['Advt_time'],
                'Program_Similarity': best_match['Program_Similarity'],
                'Time_Difference_Seconds': best_match['Time_Diff'],
                'Channel': mw_row.get('Channel', ''),
                'Match_Status': 'Time Matched / Program Mismatched'
            }
            time_matched_program_mismatched.append(mismatch_row)
            continue
            
        result_row = {
            'Media_Watch_Theme': mw_row['Advt_Theme'],
            'TC_Theme': best_match['Advt_Theme'],
            'Date': f"{mw_row['Dd']}/{mw_row['Mn']}/{mw_row['Yr']}",
            'Dd': mw_row['Dd'],
            'Mn': mw_row['Mn'], 
            'Yr': mw_row['Yr'],
            'Program': mw_row['Program'],
            'TC_Program': best_match['Program'],
            'Media_Watch_Time': mw_row['Advt_time'],
            'TC_Time': best_match['Advt_time'],
            'Program_Similarity': best_match['Program_Similarity'],
            'Time_Difference_Seconds': best_match['Time_Diff'],
            'Channel': mw_row.get('Channel', ''),
            'Advt_time': mw_row['Advt_time'],
            'Match_Status': 'Full Match',
            'Dur': mw_row.get('Dur', None)  # Add duration for schedule matching
        }

        # Copy all original LMRB columns
        for col in media_watch_df.columns:
            if col not in result_row and col not in ['Normalized_Theme', 'Time_Seconds', 'Theme_Duration']:
                col_name = f'LMRB_{col}' if col in result_row else col
                result_row[col_name] = mw_row[col]

        # Copy all TC columns
        for col in tc_df.columns:
            if col not in result_row and col not in ['Normalized_Theme', 'Time_Seconds', 'Time_Diff', 'Program_Similarity', 'Theme_Duration']:
                col_name = f'TC_{col}' if col in result_row else col
                result_row[col_name] = best_match[col]

        filter_stats["matches_found"] += 1
        matched_results.append(result_row)
        matched_mw_indices.append(i)

    progress_bar.empty()

    with st.expander("TC to LMRB Matching Statistics"):
        st.write(filter_stats)

    matched_df = pd.DataFrame(matched_results) if matched_results else pd.DataFrame()
    program_mismatched_df = pd.DataFrame(time_matched_program_mismatched) if time_matched_program_mismatched else pd.DataFrame()

    return matched_df, matched_mw_indices, program_mismatched_df

def match_media_watch_with_schedule(media_watch_df, schedule_df, theme_mapping, ignore_date=False):
    """Match Media Watch data directly with Schedule data when TC theme is empty."""
    def local_time_to_seconds(time_str):
        try:
            if pd.isna(time_str):
                return 0
            time_str = str(time_str)
            time_parts = time_str.split(':')
            if len(time_parts) >= 3:
                return int(time_parts[0]) * 3600 + int(time_parts[1]) * 60 + int(time_parts[2])
            elif len(time_parts) == 2:
                return int(time_parts[0]) * 3600 + int(time_parts[1]) * 60
            else:
                return 0
        except (ValueError, IndexError):
            return 0
            
    if media_watch_df.empty or schedule_df.empty:
        st.error("One or both DataFrames are empty.")
        return pd.DataFrame(), []

    media_watch_df = media_watch_df.copy().reset_index(drop=True)
    schedule_df = schedule_df.copy().reset_index(drop=True)

    with st.expander("Debug DataFrame Info for Schedule Matching"):
        st.write(f"Media Watch DF shape: {media_watch_df.shape}")
        st.write(f"Schedule DF shape: {schedule_df.shape}")

    if 'Advt_time' in media_watch_df.columns:
        media_watch_df['Time_Seconds'] = media_watch_df['Advt_time'].astype(str).apply(local_time_to_seconds)

    if 'Advt_time' in schedule_df.columns:
        schedule_df['Time_Seconds'] = schedule_df['Advt_time'].astype(str).apply(local_time_to_seconds)

    mw_to_schedule = {}
    for m in theme_mapping:
        if 'schedule_theme' in m and m['schedule_theme'] and ('tc_theme' not in m or not m['tc_theme']):
            mw_theme = normalize_theme_name(m['media_watch_theme'])
            schedule_theme = normalize_theme_name(m['schedule_theme'])
            mw_to_schedule[mw_theme] = schedule_theme

    with st.expander("Debug Media Watch to Schedule Theme Mappings"):
        st.write("Normalized Theme Mappings (Direct to Schedule):")
        for mw, sch in mw_to_schedule.items():
            st.write(f"Media Watch: '{mw}' → Schedule: '{sch}'")

    progress_bar = st.progress(0)
    total_rows = len(media_watch_df)

    matched_results = []
    matched_mw_indices = []

    filter_stats = {
        "total_media_watch_rows": len(media_watch_df),
        "no_normalized_mapping": 0,
        "no_matches_after_theme_filter": 0,
        "no_matches_after_date_filter": 0,
        "no_matches_after_duration_filter": 0,
        "matches_found": 0,
        "failed_program_similarity": 0
    }

    for i in range(len(media_watch_df)):
        progress_bar.progress(min(1.0, (i + 1) / total_rows))
        mw_row = media_watch_df.iloc[i]

        if pd.isna(mw_row['Advt_Theme']) or pd.isna(mw_row['Program']):
            continue

        mw_theme_norm = normalize_theme_name(mw_row['Advt_Theme'])

        if mw_theme_norm not in mw_to_schedule:
            filter_stats["no_normalized_mapping"] += 1
            continue

        schedule_theme_norm = mw_to_schedule[mw_theme_norm]
        schedule_matches = schedule_df[schedule_df['Normalized_Theme'] == schedule_theme_norm].copy().reset_index(drop=True)

        if schedule_matches.empty:
            filter_stats["no_matches_after_theme_filter"] += 1
            continue

        # Extract duration from theme
        mw_duration = extract_duration_from_theme(mw_row['Advt_Theme'])
            
        # Filter by duration if available
        if mw_duration is not None:
            if 'Theme_Duration' in schedule_matches.columns:
                # First try theme duration from extracted theme name 
                duration_matches = schedule_matches[schedule_matches['Theme_Duration'] == mw_duration]
                if not duration_matches.empty:
                    schedule_matches = duration_matches.reset_index(drop=True)
            elif 'Dur' in schedule_matches.columns:
                # Then try explicit duration column
                duration_matches = schedule_matches[schedule_matches['Dur'] == mw_duration]
                if not duration_matches.empty:
                    schedule_matches = duration_matches.reset_index(drop=True)

        if not ignore_date:
            date_filter = (
                (schedule_matches['Dd'] == mw_row['Dd']) &
                (schedule_matches['Mn'] == mw_row['Mn']) &
                (schedule_matches['Yr'] == mw_row['Yr'])
            )
            schedule_matches = schedule_matches[date_filter].reset_index(drop=True)
            
            if schedule_matches.empty:
                filter_stats["no_matches_after_date_filter"] += 1
                continue

        if 'Dur' in mw_row and 'Dur' in schedule_matches.columns:
            mw_duration = mw_row['Dur']
            schedule_matches['Duration_Diff'] = abs(schedule_matches['Dur'] - mw_duration)
            close_dur_matches = schedule_matches[schedule_matches['Duration_Diff'] <= 1].reset_index(drop=True)
            
            if close_dur_matches.empty:
                filter_stats["no_matches_after_duration_filter"] += 1
                continue
                
            schedule_matches = close_dur_matches.sort_values('Duration_Diff').reset_index(drop=True)

        schedule_matches['Program_Similarity'] = schedule_matches['Program'].astype(str).apply(
            lambda x: fuzz.token_set_ratio(str(mw_row['Program']).lower(), x.lower())
        )

        if 'Duration_Diff' in schedule_matches.columns:
            schedule_matches = schedule_matches.sort_values(
                ['Program_Similarity', 'Duration_Diff'], 
                ascending=[False, True]
            ).reset_index(drop=True)
        else:
            schedule_matches = schedule_matches.sort_values(
                'Program_Similarity', 
                ascending=False
            ).reset_index(drop=True)

        if schedule_matches.empty:
            continue
            
        best_match = schedule_matches.iloc[0]

        if best_match['Program_Similarity'] < 50:
            filter_stats["failed_program_similarity"] += 1
            continue
            
        result_row = {
            'Media_Watch_Theme': mw_row['Advt_Theme'],
            'TC_Theme': '',
            'Schedule_Theme': best_match['Advt_Theme'],
            'Date': f"{mw_row['Dd']}/{mw_row['Mn']}/{mw_row['Yr']}",
            'Dd': mw_row['Dd'],
            'Mn': mw_row['Mn'], 
            'Yr': mw_row['Yr'],
            'Program': mw_row['Program'],
            'Schedule_Program': best_match['Program'],
            'Media_Watch_Time': mw_row.get('Advt_time', ''),
            'Program_Similarity': best_match['Program_Similarity'],
            'Channel': mw_row.get('Channel', ''),
            'Match_Status': 'Direct Schedule Match'
        }

        if 'Duration_Diff' in best_match:
            result_row['Duration_Difference'] = best_match['Duration_Diff']

        # Add original LMRB columns
        for col in media_watch_df.columns:
            if col not in result_row and col not in ['Normalized_Theme', 'Time_Seconds', 'Theme_Duration']:
                col_name = f'LMRB_{col}' if col in result_row else col
                result_row[col_name] = mw_row[col]

        # Add schedule columns
        for col in schedule_df.columns:
            if col not in result_row and col not in ['Normalized_Theme', 'Duration_Diff', 'Program_Similarity', 'Time_Seconds', 'Theme_Duration']:
                col_name = f'Schedule_{col}' if col in result_row else col
                result_row[col_name] = best_match[col]

        filter_stats["matches_found"] += 1
        matched_results.append(result_row)
        matched_mw_indices.append(i)

    progress_bar.empty()

    with st.expander("Debug Schedule Match Filter Statistics"):
        st.write(filter_stats)

    matched_df = pd.DataFrame(matched_results) if matched_results else pd.DataFrame()
    return matched_df, matched_mw_indices

def match_with_schedule(matched_mw_tc_df, schedule_df, theme_mapping, ignore_date=False):
    """Match the MW+TC matched data with Schedule data based on theme mapping."""
    def local_time_to_seconds(time_str):
        try:
            if pd.isna(time_str):
                return 0
            time_str = str(time_str)
            time_parts = time_str.split(':')
            if len(time_parts) >= 3:
                return int(time_parts[0]) * 3600 + int(time_parts[1]) * 60 + int(time_parts[2])
            elif len(time_parts) == 2:
                return int(time_parts[0]) * 3600 + int(time_parts[1]) * 60
            else:
                return 0
        except (ValueError, IndexError):
            return 0
            
    if matched_mw_tc_df.empty or schedule_df.empty:
        st.error("One or both DataFrames are empty.")
        return pd.DataFrame(), pd.DataFrame()

    matched_mw_tc_df = matched_mw_tc_df.copy().reset_index(drop=True)
    schedule_df = schedule_df.copy().reset_index(drop=True)

    with st.expander("Debug DataFrame Info for Combined Matching"):
        st.write(f"MW+TC DF shape: {matched_mw_tc_df.shape}")
        st.write(f"Schedule DF shape: {schedule_df.shape}")

    if 'Advt_time' in matched_mw_tc_df.columns and 'Time_Seconds' not in matched_mw_tc_df.columns:
        matched_mw_tc_df['Time_Seconds'] = matched_mw_tc_df['Advt_time'].astype(str).apply(local_time_to_seconds)

    if 'Advt_time' in schedule_df.columns and 'Time_Seconds' not in schedule_df.columns:
        schedule_df['Time_Seconds'] = schedule_df['Advt_time'].astype(str).apply(local_time_to_seconds)

    mw_to_schedule = {}
    tc_to_schedule = {}

    for m in theme_mapping:
        if 'schedule_theme' in m and m['schedule_theme']:
            if 'media_watch_theme' in m and m['media_watch_theme']:
                mw_theme = normalize_theme_name(m['media_watch_theme'])
                schedule_theme = normalize_theme_name(m['schedule_theme'])
                mw_to_schedule[mw_theme] = schedule_theme

            if 'tc_theme' in m and m['tc_theme']:
                tc_theme = normalize_theme_name(m['tc_theme'])
                schedule_theme = normalize_theme_name(m['schedule_theme'])
                tc_to_schedule[tc_theme] = schedule_theme

    progress_bar = st.progress(0)
    total_rows = len(matched_mw_tc_df)

    matched_with_schedule = []
    unmatched_with_schedule = []

    for i in range(len(matched_mw_tc_df)):
        progress_bar.progress(min(1.0, (i + 1) / total_rows))
        mw_tc_row = matched_mw_tc_df.iloc[i]

        tc_theme_norm = normalize_theme_name(mw_tc_row['TC_Theme'])
        schedule_theme_norm = None

        if tc_theme_norm in tc_to_schedule:
            schedule_theme_norm = tc_to_schedule[tc_theme_norm]
        elif 'Media_Watch_Theme' in mw_tc_row:
            mw_theme_norm = normalize_theme_name(mw_tc_row['Media_Watch_Theme'])
            if mw_theme_norm in mw_to_schedule:
                schedule_theme_norm = mw_to_schedule[mw_theme_norm]

        if not schedule_theme_norm:
            # No mapping found - add to unmatched
            row_copy = dict(mw_tc_row)
            row_copy['Match_Status'] = 'No Schedule Theme Mapping'
            unmatched_with_schedule.append(row_copy)
            continue

        schedule_matches = schedule_df[schedule_df['Normalized_Theme'] == schedule_theme_norm].copy().reset_index(drop=True)

        if schedule_matches.empty:
            row_copy = dict(mw_tc_row)
            row_copy['Match_Status'] = 'No Matching Schedule Theme'
            unmatched_with_schedule.append(row_copy)
            continue

        # Extract duration from theme
        if 'TC_Theme' in mw_tc_row:
            tc_duration = extract_duration_from_theme(mw_tc_row['TC_Theme'])
            
            # Filter by duration if available
            if tc_duration is not None:
                if 'Theme_Duration' in schedule_matches.columns:
                    # First try theme duration from extracted theme name 
                    duration_matches = schedule_matches[schedule_matches['Theme_Duration'] == tc_duration]
                    if not duration_matches.empty:
                        schedule_matches = duration_matches.reset_index(drop=True)
                elif 'Dur' in schedule_matches.columns:
                    # Then try explicit duration column
                    duration_matches = schedule_matches[schedule_matches['Dur'] == tc_duration]
                    if not duration_matches.empty:
                        schedule_matches = duration_matches.reset_index(drop=True)

        if not ignore_date:
            date_filter = (
                (schedule_matches['Dd'] == mw_tc_row['Dd']) &
                (schedule_matches['Mn'] == mw_tc_row['Mn']) &
                (schedule_matches['Yr'] == mw_tc_row['Yr'])
            )
            schedule_matches = schedule_matches[date_filter].reset_index(drop=True)
            
            if schedule_matches.empty:
                row_copy = dict(mw_tc_row)
                row_copy['Match_Status'] = 'No Schedule Match on Date'
                unmatched_with_schedule.append(row_copy)
                continue

        if 'Dur' in mw_tc_row and 'Dur' in schedule_matches.columns:
            mw_duration = mw_tc_row['Dur']
            schedule_matches['Duration_Diff'] = abs(schedule_matches['Dur'] - mw_duration)
            close_dur_matches = schedule_matches[schedule_matches['Duration_Diff'] <= 1].reset_index(drop=True)
            
            if not close_dur_matches.empty:
                schedule_matches = close_dur_matches
            else:
                row_copy = dict(mw_tc_row)
                row_copy['Match_Status'] = 'Duration Mismatch with Schedule'
                unmatched_with_schedule.append(row_copy)
                continue

        schedule_matches['Program_Similarity'] = schedule_matches['Program'].astype(str).apply(
            lambda x: fuzz.token_set_ratio(str(mw_tc_row['Program']).lower(), x.lower())
        )

        if 'Duration_Diff' in schedule_matches.columns:
            schedule_matches = schedule_matches.sort_values(
                ['Program_Similarity', 'Duration_Diff'], 
                ascending=[False, True]
            ).reset_index(drop=True)
        else:
            schedule_matches = schedule_matches.sort_values(
                'Program_Similarity', 
                ascending=False
            ).reset_index(drop=True)

        if schedule_matches.empty:
            row_copy = dict(mw_tc_row)
            row_copy['Match_Status'] = 'No Matching Schedule After Filtering'
            unmatched_with_schedule.append(row_copy)
            continue
            
        best_match = schedule_matches.iloc[0]
            
        if best_match['Program_Similarity'] < 50:
            row_copy = dict(mw_tc_row)
            row_copy['Match_Status'] = 'Low Program Similarity with Schedule'
            row_copy['Schedule_Theme'] = best_match['Advt_Theme']
            row_copy['Schedule_Program'] = best_match['Program']
            row_copy['Program_Similarity_Schedule'] = best_match['Program_Similarity']
            unmatched_with_schedule.append(row_copy)
            continue

        result_row = dict(mw_tc_row)
        result_row['Schedule_Theme'] = best_match['Advt_Theme']
        result_row['Schedule_Program'] = best_match['Program']
        result_row['Program_Similarity_Schedule'] = best_match['Program_Similarity']
        result_row['Match_Status'] = 'Full Match with Schedule'
        
        if 'Duration_Diff' in best_match:
            result_row['Duration_Difference'] = best_match['Duration_Diff']

        for col in schedule_df.columns:
            if col not in result_row and col not in ['Normalized_Theme', 'Duration_Diff', 'Program_Similarity', 'Time_Seconds', 'Theme_Duration']:
                col_name = f'Schedule_{col}'
                result_row[col_name] = best_match[col]

        matched_with_schedule.append(result_row)

    progress_bar.empty()
    matched_df = pd.DataFrame(matched_with_schedule) if matched_with_schedule else pd.DataFrame()
    unmatched_df = pd.DataFrame(unmatched_with_schedule) if unmatched_with_schedule else pd.DataFrame()

    return matched_df, unmatched_df

def generate_summary_report(matched_df, original_media_watch, original_schedule):
    """Generate a reconciliation summary report."""
    # Create a summary by theme
    if matched_df.empty:
        return None
        
    # Group by theme
    if 'Media_Watch_Theme' in matched_df.columns:
        summary = matched_df.groupby('Media_Watch_Theme').size().reset_index(name='Aired_Count')  # Changed from 'Count' to 'Aired_Count'
    else:
        return None
        
    # Calculate total planned vs aired (if schedule data available)
    if 'Schedule_Theme' in matched_df.columns and original_schedule is not None:
        # Map from Media_Watch_Theme to Schedule_Theme
        schedule_themes = {}
        for i, row in enumerate(matched_df.iterrows()):
            mw_theme = row[1]['Media_Watch_Theme']
            schedule_theme = row[1].get('Schedule_Theme', '')
            if schedule_theme:
                schedule_themes[mw_theme] = schedule_theme
        
        # Add Schedule_Theme to summary
        summary['Schedule_Theme'] = summary['Media_Watch_Theme'].map(schedule_themes)
        
        # Group schedule by theme and duration
        if 'Dur' in original_schedule.columns:
            schedule_counts = original_schedule.groupby(['Advt_Theme', 'Dur']).size().reset_index(name='Dur_Count')
            # Then sum by theme
            schedule_theme_counts = schedule_counts.groupby('Advt_Theme')['Dur_Count'].sum().reset_index(name='Planned_Count')
        else:
            # Just group by theme if no duration column
            schedule_theme_counts = original_schedule.groupby('Advt_Theme').size().reset_index(name='Planned_Count')
        
        schedule_theme_counts.rename(columns={'Advt_Theme': 'Schedule_Theme'}, inplace=True)
        
        # Merge with summary
        summary = pd.merge(summary, schedule_theme_counts, on='Schedule_Theme', how='left')
        summary['Planned_Count'].fillna(0, inplace=True)
        summary['Planned_Count'] = summary['Planned_Count'].astype(int)
        
        # Calculate aired percentage
        summary['Aired_Percentage'] = (summary['Aired_Count'] / summary['Planned_Count'] * 100).fillna(0)
        summary.loc[summary['Aired_Percentage'] > 100, 'Aired_Percentage'] = 100  # Cap at 100%
        
    # Add match status breakdown if available
    if 'Match_Status' in matched_df.columns:
        status_counts = matched_df.groupby(['Media_Watch_Theme', 'Match_Status']).size().reset_index(name='Status_Count')
        
        # Get unique status types
        status_types = matched_df['Match_Status'].unique()
        
        # Create pivot to spread status counts across columns
        status_pivot = status_counts.pivot(
            index='Media_Watch_Theme', 
            columns='Match_Status', 
            values='Status_Count'
        ).reset_index().fillna(0)
        
        # Merge with summary
        summary = pd.merge(summary, status_pivot, on='Media_Watch_Theme', how='left')

    # Add duration information if available
    if 'Dur' in matched_df.columns:
        dur_sum = matched_df.groupby('Media_Watch_Theme')['Dur'].sum().reset_index(name='Total_Duration')
        summary = pd.merge(summary, dur_sum, on='Media_Watch_Theme', how='left')

    return summary

def find_unmatched_schedule_spots(schedule_df, matched_schedule_themes):
    """Find spots in the schedule that weren't matched with LMRB or TC data."""
    if schedule_df is None or schedule_df.empty:
        return pd.DataFrame()

    # Create a copy of the schedule
    schedule_copy = schedule_df.copy()

    # If no matches, all schedule spots are unmatched
    if not matched_schedule_themes:
        schedule_copy['Match_Status'] = 'Not Found in TC or LMRB'
        return schedule_copy

    # Check which schedule spots aren't in the matched data
    schedule_copy['Is_Matched'] = False

    for theme, dates in matched_schedule_themes.items():
        for date_tuple in dates:
            mask = (
                (schedule_copy['Normalized_Theme'] == theme) &
                (schedule_copy['Dd'] == date_tuple[0]) &
                (schedule_copy['Mn'] == date_tuple[1]) &
                (schedule_copy['Yr'] == date_tuple[2])
            )
            schedule_copy.loc[mask, 'Is_Matched'] = True

    # Filter to only unmatched spots
    unmatched_spots = schedule_copy[~schedule_copy['Is_Matched']].copy()
    unmatched_spots['Match_Status'] = 'Missing Spot - Not Found in TC or LMRB'

    return unmatched_spots.drop('Is_Matched', axis=1)

def create_tc_vs_lmrb_excel_report(matched_df, program_mismatched_df, unmatched_df, time_matched_program_mismatched_df, media_watch_filtered, filename="tc_vs_lmrb_report.xlsx"):
    """Create a comprehensive Excel report with multiple sheets for TC vs LMRB matching."""
    wb = Workbook()

    # Create Summary Sheet
    summary_sheet = wb.active
    summary_sheet.title = "Summary"

    # Add title and header
    summary_sheet.merge_cells('A1:H1')
    title_cell = summary_sheet['A1']
    title_cell.value = "TC vs LMRB Reconciliation Report"
    title_cell.font = Font(size=16, bold=True)
    title_cell.alignment = Alignment(horizontal='center')

    # Add metadata
    summary_sheet['A3'] = "Report Date:"
    summary_sheet['B3'] = pd.Timestamp.now().strftime("%Y-%m-%d")

    # Add match statistics
    summary_sheet['A5'] = "Match Statistics"
    summary_sheet['A5'].font = Font(bold=True)

    summary_sheet['A6'] = "Total LMRB Records:"
    summary_sheet['B6'] = len(media_watch_filtered) if media_watch_filtered is not None else 0

    summary_sheet['A7'] = "Full Matches Found:"
    summary_sheet['B7'] = len(matched_df) if matched_df is not None else 0

    summary_sheet['A8'] = "Program Mismatches:"
    summary_sheet['B8'] = len(program_mismatched_df) if program_mismatched_df is not None else 0

    summary_sheet['A9'] = "Time Matched / Program Mismatched:"
    summary_sheet['B9'] = len(time_matched_program_mismatched_df) if time_matched_program_mismatched_df is not None else 0

    summary_sheet['A10'] = "Unmatched TC Entries:"
    summary_sheet['B10'] = len(unmatched_df) if unmatched_df is not None else 0

    # Calculate match rate
    if media_watch_filtered is not None and len(media_watch_filtered) > 0:
        match_rate = (len(matched_df) / len(media_watch_filtered)) * 100 if matched_df is not None else 0
        summary_sheet['A11'] = "Match Rate:"
        summary_sheet['B11'] = f"{match_rate:.2f}%"

    # Add matched records summary by theme
    if matched_df is not None and not matched_df.empty:
        theme_counts = matched_df.groupby('Media_Watch_Theme').size().reset_index(name='Count')
        
        summary_sheet['A13'] = "Matches by Theme"
        summary_sheet['A13'].font = Font(bold=True)
        
        summary_sheet['A14'] = "Theme"
        summary_sheet['B14'] = "Count"
        summary_sheet['A14'].font = Font(bold=True)
        summary_sheet['B14'].font = Font(bold=True)
        
        for i, (theme, count) in enumerate(zip(theme_counts['Media_Watch_Theme'], theme_counts['Count']), 15):
            summary_sheet[f'A{i}'] = theme
            summary_sheet[f'B{i}'] = count

    # Format the summary sheet
    for col in ['A', 'B']:
        for row in range(1, 20):
            cell = summary_sheet[f'{col}{row}']
            cell.alignment = Alignment(vertical='center')
            if row in [5, 13, 14]:
                cell.font = Font(bold=True)

    # Adjust column widths - FIXED VERSION
    for col_idx in range(1, summary_sheet.max_column + 1):
        max_length = 0
        column_letter = get_column_letter(col_idx)
        
        # Check all cells in this column
        for row_idx in range(1, summary_sheet.max_row + 1):
            cell = summary_sheet.cell(row=row_idx, column=col_idx)
            if hasattr(cell, 'value') and cell.value:
                max_length = max(max_length, len(str(cell.value)))
        
        adjusted_width = max_length + 2
        summary_sheet.column_dimensions[column_letter].width = adjusted_width

    # Create Full Matches Sheet
    if matched_df is not None and not matched_df.empty:
        full_matches_sheet = wb.create_sheet("Full Matches")
        
        # Add headers
        for col_idx, column in enumerate(matched_df.columns, 1):
            cell = full_matches_sheet.cell(row=1, column=col_idx, value=column)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
        
        # Add data
        for row_idx, row in enumerate(matched_df.itertuples(index=False), 2):
            for col_idx, value in enumerate(row, 1):
                cell = full_matches_sheet.cell(row=row_idx, column=col_idx, value=value)
        
        # Adjust column widths
        for col_idx, column in enumerate(matched_df.columns, 1):
            full_matches_sheet.column_dimensions[get_column_letter(col_idx)].width = max(len(str(column)) + 2, 12)

    # Create Program Mismatches Sheet
    if program_mismatched_df is not None and not program_mismatched_df.empty:
        program_mismatch_sheet = wb.create_sheet("Program Mismatches")
        
        # Add headers
        for col_idx, column in enumerate(program_mismatched_df.columns, 1):
            cell = program_mismatch_sheet.cell(row=1, column=col_idx, value=column)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
        
        # Add data
        for row_idx, row in enumerate(program_mismatched_df.itertuples(index=False), 2):
            for col_idx, value in enumerate(row, 1):
                cell = program_mismatch_sheet.cell(row=row_idx, column=col_idx, value=value)
        
        # Adjust column widths
        for col_idx, column in enumerate(program_mismatched_df.columns, 1):
            program_mismatch_sheet.column_dimensions[get_column_letter(col_idx)].width = max(len(str(column)) + 2, 12)

    # Create Time Matched/Program Mismatched Sheet
    if time_matched_program_mismatched_df is not None and not time_matched_program_mismatched_df.empty:
        time_program_sheet = wb.create_sheet("Time Matched-Program Mismatched")
        
        # Add headers
        for col_idx, column in enumerate(time_matched_program_mismatched_df.columns, 1):
            cell = time_program_sheet.cell(row=1, column=col_idx, value=column)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
        
        # Add data
        for row_idx, row in enumerate(time_matched_program_mismatched_df.itertuples(index=False), 2):
            for col_idx, value in enumerate(row, 1):
                cell = time_program_sheet.cell(row=row_idx, column=col_idx, value=value)
        
        # Adjust column widths
        for col_idx, column in enumerate(time_matched_program_mismatched_df.columns, 1):
            time_program_sheet.column_dimensions[get_column_letter(col_idx)].width = max(len(str(column)) + 2, 12)

    # Create Unmatched TC Entries Sheet
    if unmatched_df is not None and not unmatched_df.empty:
        unmatched_sheet = wb.create_sheet("Unmatched TC Entries")
        
        # Add headers
        for col_idx, column in enumerate(unmatched_df.columns, 1):
            cell = unmatched_sheet.cell(row=1, column=col_idx, value=column)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
        
        # Add data
        for row_idx, row in enumerate(unmatched_df.itertuples(index=False), 2):
            for col_idx, value in enumerate(row, 1):
                cell = unmatched_sheet.cell(row=row_idx, column=col_idx, value=value)
        
        # Adjust column widths
        for col_idx, column in enumerate(unmatched_df.columns, 1):
            unmatched_sheet.column_dimensions[get_column_letter(col_idx)].width = max(len(str(column)) + 2, 12)

    return wb

def highlight_matched_rows_excel(df, highlight_color="#FFFF00"):
    """Highlight theme columns in the Excel output."""
    wb = Workbook()
    ws = wb.active

    # Add headers
    for col_idx, column in enumerate(df.columns, 1):
        ws.cell(row=1, column=col_idx, value=column)

    # Add data
    for row_idx, row in enumerate(df.itertuples(index=False), 2):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)

            if df.columns[col_idx-1] in ['Media_Watch_Theme', 'TC_Theme', 'Schedule_Theme']:
                cell.fill = PatternFill(
                    start_color=highlight_color.lstrip('#'),
                    end_color=highlight_color.lstrip('#'),
                    fill_type="solid"
                )
                
    # Adjust column widths safely
    for col_idx in range(1, ws.max_column + 1):
        column_letter = get_column_letter(col_idx)
        max_length = len(str(df.columns[col_idx-1])) + 2  # Start with header length
        
        # Check values in each row
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if hasattr(cell, 'value') and cell.value:
                max_length = max(max_length, len(str(cell.value)))
        
        ws.column_dimensions[column_letter].width = max_length
        
    return wb

def create_schedule_compliance_report(matched_df, schedule_df, unmatched_schedule_df):
    """Create a comprehensive Excel report for schedule compliance."""
    wb = Workbook()

    # Create Summary Sheet
    summary_sheet = wb.active
    summary_sheet.title = "Schedule Compliance Summary"

    # Add title and header
    summary_sheet.merge_cells('A1:H1')
    title_cell = summary_sheet['A1']
    title_cell.value = "Media Schedule Compliance Report"
    title_cell.font = Font(size=16, bold=True)
    title_cell.alignment = Alignment(horizontal='center')

    # Add metadata
    summary_sheet['A3'] = "Report Date:"
    summary_sheet['B3'] = pd.Timestamp.now().strftime("%Y-%m-%d")

    # Add match statistics
    summary_sheet['A5'] = "Schedule Compliance Statistics"
    summary_sheet['A5'].font = Font(bold=True)

    summary_sheet['A6'] = "Total Scheduled Spots:"
    summary_sheet['B6'] = len(schedule_df) if schedule_df is not None else 0

    summary_sheet['A7'] = "Verified Aired Spots:"
    summary_sheet['B7'] = len(matched_df) if matched_df is not None else 0

    summary_sheet['A8'] = "Missed Spots:"
    summary_sheet['B8'] = len(unmatched_schedule_df) if unmatched_schedule_df is not None else 0

    # Calculate compliance rate
    if schedule_df is not None and len(schedule_df) > 0:
        compliance_rate = (len(matched_df) / len(schedule_df)) * 100 if matched_df is not None else 0
        summary_sheet['A9'] = "Compliance Rate:"
        summary_sheet['B9'] = f"{compliance_rate:.2f}%"
    
    # Add theme-based summary
    if schedule_df is not None and matched_df is not None:
        # Group by theme
        if 'Schedule_Theme' in matched_df.columns and 'Advt_Theme' in schedule_df.columns:
            matched_theme_counts = matched_df.groupby('Schedule_Theme').size().reset_index(name='Aired_Count')
            schedule_theme_counts = schedule_df.groupby('Advt_Theme').size().reset_index(name='Scheduled_Count')
            schedule_theme_counts.rename(columns={'Advt_Theme': 'Theme'}, inplace=True)
            matched_theme_counts.rename(columns={'Schedule_Theme': 'Theme'}, inplace=True)
            
            # Merge to get comparison
            theme_summary = pd.merge(schedule_theme_counts, matched_theme_counts, on='Theme', how='left')
            theme_summary['Aired_Count'].fillna(0, inplace=True)
            theme_summary['Aired_Count'] = theme_summary['Aired_Count'].astype(int)
            theme_summary['Missed_Count'] = theme_summary['Scheduled_Count'] - theme_summary['Aired_Count']
            theme_summary['Compliance_Rate'] = (theme_summary['Aired_Count'] / theme_summary['Scheduled_Count'] * 100).round(2)
            
            # Add to sheet
            summary_sheet['A11'] = "Compliance by Theme"
            summary_sheet['A11'].font = Font(bold=True)
            
            headers = ['Theme', 'Scheduled', 'Aired', 'Missed', 'Compliance %']
            for col_idx, header in enumerate(headers, 1):
                cell = summary_sheet.cell(row=12, column=col_idx, value=header)
                cell.font = Font(bold=True)
            
            for row_idx, theme_row in enumerate(theme_summary.itertuples(index=False), 13):
                summary_sheet.cell(row=row_idx, column=1, value=theme_row.Theme)
                summary_sheet.cell(row=row_idx, column=2, value=theme_row.Scheduled_Count)
                summary_sheet.cell(row=row_idx, column=3, value=theme_row.Aired_Count)
                summary_sheet.cell(row=row_idx, column=4, value=theme_row.Missed_Count)
                summary_sheet.cell(row=row_idx, column=5, value=f"{theme_row.Compliance_Rate}%")
    
    # Create Matched Schedule Sheet
    if matched_df is not None and not matched_df.empty:
        matched_sheet = wb.create_sheet("Aired as Scheduled")
        
        # Add headers
        important_columns = [
            'Schedule_Theme', 'Media_Watch_Theme', 'Date', 
            'Schedule_Program', 'Program', 'Media_Watch_Time', 
            'Duration_Difference', 'Match_Status'
        ]
        
        # Find available columns
        available_columns = [col for col in important_columns if col in matched_df.columns]
        remaining_columns = [col for col in matched_df.columns if col not in available_columns]
        display_columns = available_columns + remaining_columns
        
        for col_idx, column in enumerate(display_columns, 1):
            if column in matched_df.columns:
                cell = matched_sheet.cell(row=1, column=col_idx, value=column)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
        
        # Add data
        for row_idx, row in enumerate(matched_df.itertuples(index=False), 2):
            for col_idx, column in enumerate(display_columns, 1):
                if column in matched_df.columns:
                    col_pos = matched_df.columns.get_loc(column)
                    value = row[col_pos]
                    matched_sheet.cell(row=row_idx, column=col_idx, value=value)
        
        # Adjust column widths
        for col_idx, column in enumerate(display_columns, 1):
            if column in matched_df.columns:
                matched_sheet.column_dimensions[get_column_letter(col_idx)].width = max(len(column) + 2, 15)
    
    # Create Missed Schedule Sheet
    if unmatched_schedule_df is not None and not unmatched_schedule_df.empty:
        missed_sheet = wb.create_sheet("Missed Schedule Spots")
        
        # Add headers
        important_columns = [
            'Advt_Theme', 'Date', 'Program', 'Advt_time', 'Dur', 'Match_Status'
        ]
        
        # Find available columns
        available_columns = [col for col in important_columns if col in unmatched_schedule_df.columns]
        remaining_columns = [col for col in unmatched_schedule_df.columns if col not in available_columns]
        display_columns = available_columns + remaining_columns
        
        for col_idx, column in enumerate(display_columns, 1):
            if column in unmatched_schedule_df.columns:
                cell = missed_sheet.cell(row=1, column=col_idx, value=column)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="FFDDDD", end_color="FFDDDD", fill_type="solid")  # Light red for missed
        
        # Add data
        for row_idx, row in enumerate(unmatched_schedule_df.itertuples(index=False), 2):
            for col_idx, column in enumerate(display_columns, 1):
                if column in unmatched_schedule_df.columns:
                    col_pos = unmatched_schedule_df.columns.get_loc(column)
                    value = row[col_pos]
                    missed_sheet.cell(row=row_idx, column=col_idx, value=value)
        
        # Adjust column widths
        for col_idx, column in enumerate(display_columns, 1):
            if column in unmatched_schedule_df.columns:
                missed_sheet.column_dimensions[get_column_letter(col_idx)].width = max(len(column) + 2, 15)
    
    return wb

def read_uploaded_file(uploaded_file):
 """Read an uploaded file into a pandas DataFrame."""
 if uploaded_file is None:
     return None

 try:
     file_extension = uploaded_file.name.split('.')[-1].lower()

     if file_extension == 'xlsx':
         df = pd.read_excel(uploaded_file, engine='openpyxl')
     elif file_extension == 'xls':
         df = pd.read_excel(uploaded_file, engine='xlrd')
     elif file_extension == 'csv':
         df = pd.read_csv(uploaded_file)
     else:
         st.error(f"Unsupported file format: {file_extension}")
         return None

     df = df.reset_index(drop=True)
     df = clean_column_names(df)

     return df
 except Exception as e:
     st.error(f"Error reading file: {str(e)}")
     return None

def main():
 st.set_page_config(page_title="Media Data Matcher", layout="wide")

 st.title("Media Reconciliation")
 st.write("Upload data files, select channel, map themes, and find matching records between LMRB, TC and Schedule.")

 with st.sidebar:
     st.header("Upload Files")
     media_watch_file = st.file_uploader("Upload LMRB Data", type=["xlsx", "xls", "csv"])
     tc_file = st.file_uploader("Upload TC Data", type=["xlsx", "xls", "csv"])
     schedule_file = st.file_uploader("Upload Schedule Data", type=["xlsx", "xls", "csv"])
     prev_month_file = st.file_uploader("Upload Previous Month LMRB Data (optional)", type=["xlsx", "xls", "csv"])
     
     st.header("Matching Configuration")
     time_tolerance = st.slider("Time Matching Tolerance (seconds)", min_value=5, max_value=60, value=30, step=5,
                             help="Maximum time difference (in seconds) allowed for matching times between datasets")

 media_watch_df = read_uploaded_file(media_watch_file)
 tc_df = read_uploaded_file(tc_file)
 schedule_df = read_uploaded_file(schedule_file)
 prev_month_df = read_uploaded_file(prev_month_file)

 if 'mw_tc_matched_data' not in st.session_state:
     st.session_state.mw_tc_matched_data = None

 if 'mw_schedule_matched_data' not in st.session_state:
     st.session_state.mw_schedule_matched_data = None

 if 'final_matched_data' not in st.session_state:
     st.session_state.final_matched_data = None
     
 if 'program_mismatched_data' not in st.session_state:
     st.session_state.program_mismatched_data = None

 if 'theme_mapping' not in st.session_state:
     st.session_state.theme_mapping = []

 if 'matched_mw_indices' not in st.session_state:
     st.session_state.matched_mw_indices = []

 if media_watch_df is not None:
     st.success("LMRB data file uploaded successfully!")

     if "Channel" in media_watch_df.columns:
         available_channels = sorted(media_watch_df["Channel"].unique())
         channel = st.selectbox("Select Channel", options=available_channels)
         
         media_watch_filtered = filter_by_channel(media_watch_df, channel)
         
         if prev_month_df is not None:
             prev_month_filtered = filter_by_channel(prev_month_df, channel)
             cleaned_mw_df, duplicate_count = remove_duplicates(media_watch_filtered, prev_month_filtered)
             
             st.info(f"Removed {duplicate_count} duplicate records from previous month's data.")
             media_watch_filtered = cleaned_mw_df
         
         st.subheader("Theme Mapping")
         
         media_watch_theme_col = st.selectbox(
             "Select LMRB Theme Column",
             options=[col for col in media_watch_filtered.columns if 'theme' in col.lower() or 'advt' in col.lower()],
             key="mw_theme_col"
         )
         
         media_watch_std = standardize_dataframe(
             media_watch_filtered,
             theme_col=media_watch_theme_col,
             program_col=next((col for col in media_watch_filtered.columns if 'program' in col.lower()), None),
             time_col=next((col for col in media_watch_filtered.columns if 'time' in col.lower() and 'advt' in col.lower()), None),
             date_col=next((col for col in media_watch_filtered.columns if 'date' in col.lower()), None)
         )
         
         if tc_df is not None:
             tc_theme_col = st.selectbox(
                 "Select TC Theme Column",
                 options=[col for col in tc_df.columns if 'theme' in col.lower() or 'advt' in col.lower()],
                 key="tc_theme_col"
             )
             
             tc_std = standardize_dataframe(
                 tc_df,
                 theme_col=tc_theme_col,
                 program_col=next((col for col in tc_df.columns if 'program' in col.lower()), None),
                 time_col=next((col for col in tc_df.columns if 'time' in col.lower() and 'spot' in col.lower() or 'air' in col.lower()), None),
                 date_col=next((col for col in tc_df.columns if 'date' in col.lower()), None)
             )
         else:
             tc_std = None
         
         if schedule_df is not None:
             schedule_theme_col = st.selectbox(
                 "Select Schedule Theme Column",
                 options=[col for col in schedule_df.columns if 'theme' in col.lower() or 'advt' in col.lower()],
                 key="schedule_theme_col"
             )
             
             schedule_std = standardize_dataframe(
                 schedule_df,
                 theme_col=schedule_theme_col,
                 program_col=next((col for col in schedule_df.columns if 'program' in col.lower()), None),
                 time_col=next((col for col in schedule_df.columns if 'time' in col.lower()), None),
                 date_col=next((col for col in schedule_df.columns if 'date' in col.lower()), None)
             )
         else:
             schedule_std = None
         
         col1, col2, col3 = st.columns(3)
         
         with col1:
             st.subheader("LMRB Theme")
             media_watch_themes = sorted(media_watch_std['Advt_Theme'].dropna().unique())
             selected_mw_theme = st.selectbox("Select LMRB Theme", options=[""] + media_watch_themes)
         
         with col2:
             st.subheader("TC Theme")
             if tc_std is not None:
                 tc_themes = sorted(tc_std['Advt_Theme'].dropna().unique())
                 selected_tc_theme = st.selectbox("Select TC Theme (leave empty if not in TC)", options=[""] + tc_themes)
             else:
                 selected_tc_theme = ""
                 st.info("No TC data uploaded")
         
         with col3:
             st.subheader("Schedule Theme")
             if schedule_std is not None:
                 schedule_themes = sorted(schedule_std['Advt_Theme'].dropna().unique())
                 selected_schedule_theme = st.selectbox("Select Schedule Theme", options=[""] + schedule_themes)
             else:
                 selected_schedule_theme = ""
                 st.info("No Schedule data uploaded")
         
         if st.button("Add Theme Mapping"):
             if selected_mw_theme:
                 is_duplicate = False
                 for mapping in st.session_state.theme_mapping:
                     if normalize_theme_name(mapping['media_watch_theme']) == normalize_theme_name(selected_mw_theme):
                         is_duplicate = True
                         break
                 
                 if not is_duplicate:
                     mapping_entry = {
                         'media_watch_theme': selected_mw_theme,
                     }
                     
                     if selected_tc_theme:
                         mapping_entry['tc_theme'] = selected_tc_theme
                     else:
                         mapping_entry['tc_theme'] = ""
                         
                     if selected_schedule_theme:
                         mapping_entry['schedule_theme'] = selected_schedule_theme
                         
                     st.session_state.theme_mapping.append(mapping_entry)
                     
                     mapping_msg = f"Added mapping: {selected_mw_theme}"
                     if selected_tc_theme:
                         mapping_msg += f" → TC: {selected_tc_theme}"
                     if selected_schedule_theme:
                         mapping_msg += f" → Schedule: {selected_schedule_theme}"
                         
                     st.success(mapping_msg)
                 else:
                     st.warning(f"A mapping for '{selected_mw_theme}' already exists.")
             else:
                 st.warning("Please select at least an LMRB theme to create a mapping.")
         
         if st.session_state.theme_mapping:
             st.subheader("Current Theme Mappings")
             
             cols = st.columns([3, 3, 3, 1])
             cols[0].markdown("**LMRB Theme**")
             cols[1].markdown("**TC Theme**")
             cols[2].markdown("**Schedule Theme**")
             cols[3].markdown("**Action**")
             
             st.markdown("---")
             
             for i, mapping in enumerate(st.session_state.theme_mapping):
                 cols = st.columns([3, 3, 3, 1])
                 cols[0].write(mapping['media_watch_theme'])
                 cols[1].write(mapping.get('tc_theme', ''))
                 cols[2].write(mapping.get('schedule_theme', ''))
                 if cols[3].button("Delete", key=f"del_{i}"):
                     st.session_state.theme_mapping.pop(i)
                     st.rerun()
         
         st.subheader("Matching Options")
         ignore_date = st.checkbox("Ignore Date (Match Only by Theme, Program, and Duration)", value=False)
         
         st.subheader("Data Matching")
         
         match_all = st.checkbox("Perform all possible matches", value=True)
         
         if match_all:
             if st.button("Match All Data"):
                 if not st.session_state.theme_mapping:
                     st.warning("Please create theme mappings before matching data.")
                 else:
                     all_results = []
                     unmatched_mw_indices = list(range(len(media_watch_std)))
                     
                     with st.spinner("Matching LMRB with TC data..."):
                         if tc_std is not None:
                             try:
                                 matched_tc_results, matched_tc_indices, program_mismatched_df = match_media_watch_with_tc(
                                     media_watch_std, 
                                     tc_std, 
                                     st.session_state.theme_mapping,
                                     ignore_date=ignore_date,
                                     time_tolerance=time_tolerance
                                 )
                                 
                                 if not matched_tc_results.empty:
                                     st.session_state.mw_tc_matched_data = matched_tc_results
                                     st.session_state.matched_mw_indices = matched_tc_indices
                                     st.session_state.program_mismatched_data = program_mismatched_df
                                     
                                     st.success(f"Found {len(matched_tc_results)} matching records between LMRB and TC data!")
                                     if not program_mismatched_df.empty:
                                         st.info(f"Found {len(program_mismatched_df)} records with time matched but program mismatched.")
                                     
                                     # Generate comprehensive TC vs LMRB report
                                     tc_lmrb_report = create_tc_vs_lmrb_excel_report(
                                         matched_tc_results,
                                         program_mismatched_df,
                                         pd.DataFrame(),  # Unmatched (we don't have this yet)
                                         program_mismatched_df,  # Time matched/program mismatched
                                         media_watch_filtered,
                                         f"{channel}_TC_vs_LMRB_report.xlsx"
                                     )
                                     
                                     # Convert report to binary for download
                                     tc_lmrb_buffer = BytesIO()
                                     tc_lmrb_report.save(tc_lmrb_buffer)
                                     tc_lmrb_buffer.seek(0)
                                     
                                     st.download_button(
                                         label="Download TC vs LMRB Reconciliation Report",
                                         data=tc_lmrb_buffer,
                                         file_name=f"{channel}_TC_vs_LMRB_report.xlsx",
                                         mime="application/vnd.ms-excel"
                                     )
                                     
                                     matched_mw_original = media_watch_filtered.iloc[matched_tc_indices].copy()
                                     
                                     excel_buffer = BytesIO()
                                     matched_mw_original.to_excel(excel_buffer, index=False)
                                     excel_buffer.seek(0)
                                     
                                     st.download_button(
                                         label="Download Original LMRB Data (TC Matched)",
                                         data=excel_buffer,
                                         file_name=f"{channel}_original_lmrb_tc_matched.xlsx",
                                         mime="application/vnd.ms-excel"
                                     )
                                     
                                     all_results.append(matched_tc_results)
                                     
                                     for idx in matched_tc_indices:
                                         if idx in unmatched_mw_indices:
                                             unmatched_mw_indices.remove(idx)
                             except Exception as e:
                                 st.error(f"Error during TC matching: {str(e)}")
                                 st.exception(e)
                     
                     if st.session_state.mw_tc_matched_data is not None and not st.session_state.mw_tc_matched_data.empty and schedule_std is not None:
                         with st.spinner("Matching TC results with Schedule data..."):
                             try:
                                 schedule_matched_results, schedule_unmatched_results = match_with_schedule(
                                     st.session_state.mw_tc_matched_data,
                                     schedule_std,
                                     st.session_state.theme_mapping,
                                     ignore_date=ignore_date
                                 )
                                 
                                 if not schedule_matched_results.empty:
                                     st.success(f"Found {len(schedule_matched_results)} matching records between TC results and Schedule data!")
                                     
                                     # Extract matched schedule themes and dates for finding unmatched spots
                                     matched_schedule_themes = {}
                                     
                                     for _, row in schedule_matched_results.iterrows():
                                         if 'Schedule_Theme' in row:
                                             theme = normalize_theme_name(row['Schedule_Theme'])
                                             date_tuple = (row['Dd'], row['Mn'], row['Yr'])
                                             
                                             if theme not in matched_schedule_themes:
                                                 matched_schedule_themes[theme] = []
                                                 
                                             matched_schedule_themes[theme].append(date_tuple)
                                             
                                     # Find unmatched schedule spots
                                     unmatched_schedule_df = find_unmatched_schedule_spots(
                                         schedule_std, 
                                         matched_schedule_themes
                                     )
                                     
                                     # Create schedule compliance report
                                     schedule_report = create_schedule_compliance_report(
                                         schedule_matched_results,
                                         schedule_std,
                                         unmatched_schedule_df
                                     )
                                     
                                     # Convert report to binary for download
                                     schedule_buffer = BytesIO()
                                     schedule_report.save(schedule_buffer)
                                     schedule_buffer.seek(0)
                                     
                                     st.download_button(
                                         label="Download Schedule Compliance Report",
                                         data=schedule_buffer,
                                         file_name=f"{channel}_schedule_compliance_report.xlsx",
                                         mime="application/vnd.ms-excel"
                                     )
                                     
                                     all_results.append(schedule_matched_results)
                             except Exception as e:
                                 st.error(f"Error during Schedule matching: {str(e)}")
                                 st.exception(e)
                     
                     if schedule_std is not None and unmatched_mw_indices:
                         with st.spinner("Matching unmatched LMRB data directly with Schedule..."):
                             try:
                                 unmatched_rows = []
                                 for idx in unmatched_mw_indices:
                                     if idx < len(media_watch_std):
                                         unmatched_rows.append(media_watch_std.iloc[idx])
                                 
                                 if unmatched_rows:
                                     unmatched_mw_df = pd.DataFrame(unmatched_rows).reset_index(drop=True)
                                     
                                     direct_schedule_results, direct_schedule_indices = match_media_watch_with_schedule(
                                         unmatched_mw_df,
                                         schedule_std,
                                         st.session_state.theme_mapping,
                                         ignore_date=ignore_date
                                     )
                                     
                                     if not direct_schedule_results.empty:
                                         st.success(f"Found {len(direct_schedule_results)} direct matches between LMRB and Schedule data!")
                                         
                                         direct_original_indices = [unmatched_mw_indices[i] for i in direct_schedule_indices if i < len(unmatched_mw_indices)]
                                         if direct_original_indices:
                                             direct_mw_original = media_watch_filtered.iloc[direct_original_indices].copy()
                                             
                                             excel_buffer = BytesIO()
                                             direct_mw_original.to_excel(excel_buffer, index=False)
                                             excel_buffer.seek(0)
                                             
                                             st.download_button(
                                                 label="Download Original LMRB Data (Schedule Direct Matched)",
                                                 data=excel_buffer,
                                                 file_name=f"{channel}_original_lmrb_schedule_matched.xlsx",
                                                 mime="application/vnd.ms-excel"
                                             )
                                         
                                         all_results.append(direct_schedule_results)
                             except Exception as e:
                                 st.error(f"Error during direct Schedule matching: {str(e)}")
                                 st.exception(e)
                     
                     if all_results:
                         for i in range(len(all_results)):
                             all_results[i] = all_results[i].reset_index(drop=True)
                             
                         final_results = pd.concat(all_results, ignore_index=True)
                         st.session_state.final_matched_data = final_results
                         
                         st.subheader("Combined Matching Results")
                         st.dataframe(final_results)
                         
                         # Generate summary report
                         summary = generate_summary_report(final_results, media_watch_filtered, schedule_std)
                         if summary is not None:
                             st.subheader("Reconciliation Summary")
                             st.dataframe(summary)
                             
                             # Export summary
                             summary_buffer = BytesIO()
                             summary.to_excel(summary_buffer, index=False)
                             summary_buffer.seek(0)
                             
                             st.download_button(
                                 label="Download Reconciliation Summary",
                                 data=summary_buffer,
                                 file_name=f"{channel}_reconciliation_summary.xlsx",
                                 mime="application/vnd.ms-excel"
                             )
                         
                         highlight_color = st.color_picker("Choose highlight color for Excel", "#FFFF00")
                         
                         excel_buffer = BytesIO()
                         wb = highlight_matched_rows_excel(final_results, highlight_color)
                         wb.save(excel_buffer)
                         excel_buffer.seek(0)
                         
                         st.download_button(
                             label="Download Combined Matching Results",
                             data=excel_buffer,
                             file_name=f"{channel}_combined_matching_results.xlsx",
                             mime="application/vnd.ms-excel"
                         )
                     else:
                         st.warning("No matches found in any of the matching processes.")
         else:
             col1, col2, col3 = st.columns(3)
             
             with col1:
                 st.subheader("Step 1: Match with TC")
                 if st.button("Match LMRB with TC"):
                     if not st.session_state.theme_mapping or tc_std is None:
                         st.warning("Please create theme mappings and upload TC data.")
                     else:
                         with st.spinner("Matching LMRB with TC data..."):
                             try:
                                 matched_tc_results, matched_indices, program_mismatched_df = match_media_watch_with_tc(
                                     media_watch_std, 
                                     tc_std, 
                                     st.session_state.theme_mapping,
                                     ignore_date=ignore_date,
                                     time_tolerance=time_tolerance
                                 )
                                 
                                 if not matched_tc_results.empty:
                                     st.session_state.mw_tc_matched_data = matched_tc_results
                                     st.session_state.matched_mw_indices = matched_indices
                                     st.session_state.program_mismatched_data = program_mismatched_df
                                     
                                     st.success(f"Found {len(matched_tc_results)} matching records between LMRB and TC data!")
                                     if not program_mismatched_df.empty:
                                         st.info(f"Found {len(program_mismatched_df)} records with time matched but program mismatched.")
                                     
                                     # Display results
                                     st.dataframe(matched_tc_results)
                                     
                                     # Generate comprehensive TC vs LMRB report
                                     tc_lmrb_report = create_tc_vs_lmrb_excel_report(
                                         matched_tc_results,
                                         program_mismatched_df,
                                         pd.DataFrame(),  # Unmatched (we don't have this yet)
                                         program_mismatched_df,  # Time matched/program mismatched
                                         media_watch_filtered,
                                         f"{channel}_TC_vs_LMRB_report.xlsx"
                                     )
                                     
                                     # Convert report to binary for download
                                     tc_lmrb_buffer = BytesIO()
                                     tc_lmrb_report.save(tc_lmrb_buffer)
                                     tc_lmrb_buffer.seek(0)
                                     
                                     st.download_button(
                                         label="Download TC vs LMRB Reconciliation Report",
                                         data=tc_lmrb_buffer,
                                         file_name=f"{channel}_TC_vs_LMRB_report.xlsx",
                                         mime="application/vnd.ms-excel"
                                     )
                                     
                                     matched_mw_original = media_watch_filtered.iloc[matched_indices].copy()
                                     
                                     excel_buffer = BytesIO()
                                     matched_mw_original.to_excel(excel_buffer, index=False)
                                     excel_buffer.seek(0)
                                     
                                     st.download_button(
                                         label="Download Original LMRB Data (Matched Rows)",
                                         data=excel_buffer,
                                         file_name=f"{channel}_original_lmrb_matched.xlsx",
                                         mime="application/vnd.ms-excel"
                                     )
                                     
                                     excel_buffer2 = BytesIO()
                                     matched_tc_results.to_excel(excel_buffer2, index=False)
                                     excel_buffer2.seek(0)
                                     
                                     st.download_button(
                                         label="Download LMRB - TC Matched Results",
                                         data=excel_buffer2,
                                         file_name=f"{channel}_lmrb_tc_matched_results.xlsx",
                                         mime="application/vnd.ms-excel"
                                     )
                                     
                                     if not program_mismatched_df.empty:
                                         excel_buffer3 = BytesIO()
                                         program_mismatched_df.to_excel(excel_buffer3, index=False)
                                         excel_buffer3.seek(0)
                                         
                                         st.download_button(
                                             label="Download Time Matched/Program Mismatched Results",
                                             data=excel_buffer3,
                                             file_name=f"{channel}_time_matched_program_mismatched.xlsx",
                                             mime="application/vnd.ms-excel"
                                         )
                                 else:
                                     st.warning("No matches found between LMRB and TC data.")
                             except Exception as e:
                                 st.error(f"Error during TC matching: {str(e)}")
                                 st.exception(e)
             
             with col2:
                 st.subheader("Step 2: Match with Schedule")
                 if st.button("Match TC Results with Schedule"):
                     if not st.session_state.theme_mapping or schedule_std is None:
                         st.warning("Please create theme mappings and upload Schedule data.")
                     elif st.session_state.mw_tc_matched_data is None:
                         st.warning("Please match LMRB with TC data first.")
                     else:
                         with st.spinner("Matching with Schedule data..."):
                             try:
                                 matched_schedule, unmatched_schedule = match_with_schedule(
                                     st.session_state.mw_tc_matched_data,
                                     schedule_std,
                                     st.session_state.theme_mapping,
                                     ignore_date=ignore_date
                                 )
                                 
                                 if not matched_schedule.empty:
                                     st.session_state.final_matched_data = matched_schedule
                                     st.success(f"Found {len(matched_schedule)} matching records with Schedule data!")
                                     
                                     # Display results
                                     st.dataframe(matched_schedule)
                                     
                                     # Extract matched schedule themes and dates for finding unmatched spots
                                     matched_schedule_themes = {}
                                     
                                     for _, row in matched_schedule.iterrows():
                                         if 'Schedule_Theme' in row:
                                             theme = normalize_theme_name(row['Schedule_Theme'])
                                             date_tuple = (row['Dd'], row['Mn'], row['Yr'])
                                             
                                             if theme not in matched_schedule_themes:
                                                 matched_schedule_themes[theme] = []
                                                 
                                             matched_schedule_themes[theme].append(date_tuple)
                                             
                                     # Find unmatched schedule spots
                                     unmatched_schedule_df = find_unmatched_schedule_spots(
                                         schedule_std, 
                                         matched_schedule_themes
                                     )
                                     
                                     # Create schedule compliance report
                                     schedule_report = create_schedule_compliance_report(
                                         matched_schedule,
                                         schedule_std,
                                         unmatched_schedule_df
                                     )
                                     
                                     # Convert report to binary for download
                                     schedule_buffer = BytesIO()
                                     schedule_report.save(schedule_buffer)
                                     schedule_buffer.seek(0)
                                     
                                     st.download_button(
                                         label="Download Schedule Compliance Report",
                                         data=schedule_buffer,
                                         file_name=f"{channel}_schedule_compliance_report.xlsx",
                                         mime="application/vnd.ms-excel"
                                     )
                                     
                                     highlight_color = st.color_picker("Choose highlight color for Excel", "#FFFF00")
                                     
                                     excel_buffer = BytesIO()
                                     wb = highlight_matched_rows_excel(matched_schedule, highlight_color)
                                     wb.save(excel_buffer)
                                     excel_buffer.seek(0)
                                     
                                     st.download_button(
                                         label="Download Final Matched Results",
                                         data=excel_buffer,
                                         file_name=f"{channel}_final_matched_results.xlsx",
                                         mime="application/vnd.ms-excel"
                                     )
                                     
                                     # Show unmatched results
                                     if not unmatched_schedule.empty:
                                         st.warning(f"Found {len(unmatched_schedule)} records that matched TC but not Schedule")
                                         with st.expander("Show TC Entries Not Matching Schedule"):
                                             st.dataframe(unmatched_schedule)
                                             
                                             excel_buffer2 = BytesIO()
                                             unmatched_schedule.to_excel(excel_buffer2, index=False)
                                             excel_buffer2.seek(0)
                                             
                                             st.download_button(
                                                 label="Download Unmatched Schedule Results",
                                                 data=excel_buffer2,
                                                 file_name=f"{channel}_tc_not_matching_schedule.xlsx",
                                                 mime="application/vnd.ms-excel"
                                             )
                                 else:
                                     st.warning("No matches found with Schedule data.")
                             except Exception as e:
                                 st.error(f"Error during Schedule matching: {str(e)}")
                                 st.exception(e)
             
             with col3:
                 st.subheader("Step 3: Direct LMRB to Schedule")
                 if st.button("Match Unmatched LMRB to Schedule"):
                     if not st.session_state.theme_mapping or schedule_std is None:
                         st.warning("Please create theme mappings and upload Schedule data.")
                     else:
                         with st.spinner("Matching unmatched LMRB data with Schedule..."):
                             try:
                                 unmatched_mw = media_watch_std.copy()
                                 
                                 if hasattr(st.session_state, 'matched_mw_indices') and st.session_state.matched_mw_indices:
                                     matched_indices_list = list(st.session_state.matched_mw_indices)
                                     mask = ~unmatched_mw.index.isin(matched_indices_list)
                                     unmatched_mw = unmatched_mw[mask].reset_index(drop=True)
                                 
                                 direct_schedule_results, direct_matched_indices = match_media_watch_with_schedule(
                                     unmatched_mw,
                                     schedule_std,
                                     st.session_state.theme_mapping,
                                     ignore_date=ignore_date
                                 )
                                 
                                 if not direct_schedule_results.empty:
                                     st.session_state.mw_schedule_matched_data = direct_schedule_results
                                     st.success(f"Found {len(direct_schedule_results)} direct matches between LMRB and Schedule!")
                                     st.dataframe(direct_schedule_results)
                                     
                                     original_indices = []
                                     current_idx = 0
                                     
                                     for i in range(len(media_watch_std)):
                                         if i not in st.session_state.matched_mw_indices:
                                             if current_idx in direct_matched_indices:
                                                 original_indices.append(i)
                                             current_idx += 1
                                     
                                     if original_indices:
                                         direct_mw_original = media_watch_filtered.iloc[original_indices].copy()
                                         
                                         excel_buffer = BytesIO()
                                         direct_mw_original.to_excel(excel_buffer, index=False)
                                         excel_buffer.seek(0)
                                         
                                         st.download_button(
                                             label="Download Original LMRB Data (Schedule Matches)",
                                             data=excel_buffer,
                                             file_name=f"{channel}_original_lmrb_schedule_matched.xlsx",
                                             mime="application/vnd.ms-excel"
                                         )
                                     
                                     excel_buffer2 = BytesIO()
                                     direct_schedule_results.to_excel(excel_buffer2, index=False)
                                     excel_buffer2.seek(0)
                                     
                                     st.download_button(
                                         label="Download LMRB - Schedule Direct Matches",
                                         data=excel_buffer2,
                                         file_name=f"{channel}_lmrb_schedule_direct_matched.xlsx",
                                         mime="application/vnd.ms-excel"
                                     )
                                 else:
                                     st.warning("No direct matches found between LMRB and Schedule data.")
                             except Exception as e:
                                 st.error(f"Error during direct Schedule matching: {str(e)}")
                                 st.exception(e)
     else:
         st.error("Channel column not found in LMRB data.")
 else:
     st.info("Please upload LMRB data file to continue.")

if __name__ == "__main__":
 main()