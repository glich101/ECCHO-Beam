#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel Generator - Creates all 16 Excel sheets with analysis
Maintains original analysis logic while adding robustness
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment
import logging
import os

class ExcelGenerator:
    def __init__(self, progress_callback=None):
        self.progress_callback = progress_callback
        self.cancel_flag = False
        
    def set_cancel_flag(self):
        """Set flag to cancel processing"""
        self.cancel_flag = True
        
    def update_progress(self, percent, message=""):
        """Update progress if callback is available"""
        if self.progress_callback:
            self.progress_callback(percent, message)
    
    def clean_text(self, s):
        """Clean and normalize text"""
        if s is None or (isinstance(s, float) and np.isnan(s)): 
            return ""
        import re
        return re.sub(r"\s+", " ", str(s)).strip()

    def safe_reindex_columns(self, df, columns):
        """Safely reindex dataframe columns"""
        df = df.loc[:, ~df.columns.duplicated()]
        cols = []
        seen = set()
        for c in columns:
            if c not in seen:
                cols.append(c); seen.add(c)
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        return df.reindex(columns=cols)

    def format_excel_sheet(self, writer, sheet_name, df):
        """Apply formatting to Excel sheet"""
        try:
            workbook = writer.book
            worksheet = workbook[sheet_name]
            
            # Apply AutoFilter
            if len(df) > 0:
                worksheet.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"
            
            # Freeze top row
            worksheet.freeze_panes = "A2"
            
            # Auto-fit column widths
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        cell_value = str(cell.value) if cell.value is not None else ""
                        if len(cell_value) > max_length:
                            max_length = len(cell_value)
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
                worksheet.column_dimensions[column].width = adjusted_width
            
            # Format header row
            for cell in worksheet[1]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                cell.alignment = Alignment(horizontal="center")
                
        except Exception as e:
            logging.warning(f"Error formatting sheet {sheet_name}: {e}")

    def create_mapping_sheet(self, df):
        """Create Mapping sheet - comprehensive call/SMS records with all original columns"""
        try:
            n = len(df)
            empty_series = lambda val="": pd.Series([val]*n)
            
            # Build comprehensive mapping with all original columns
            mapping_df = pd.DataFrame({
                'CdrNo': df['CdrNo'].fillna('').astype(str),
                'B Party': df['Counterparty'].apply(self.clean_text),
                'Date': df['DateStr'].fillna('').astype(str),
                'Time': df['TimeStr'].fillna('').astype(str),
                'Duration': df['Duration'].fillna(0).astype('Int64'),
                'Call Type': df['CallTypeStd'].fillna('').astype(str),
                'First Cell ID': df['FirstCellID'].apply(self.clean_text),
                'First Cell ID Address': df['FirstCellAddr'].apply(self.clean_text),
                'Last Cell ID': df['LastCellID'].apply(self.clean_text),
                'Last Cell ID Address': df['LastCellAddr'].apply(self.clean_text),
                'IMEI': df['IMEI'].apply(self.clean_text),
                'IMEI Manufacturer': empty_series(""),
                'Device Type': empty_series(""),
                'IMSI': df['IMSI'].apply(self.clean_text),
                'Roaming': df['Circle'].fillna("").astype(str),
                'B Party Provider': empty_series(""),
                'Main City(First CellID)': df['FirstCellCity'].apply(self.clean_text),
                'Sub City(First CellID)': empty_series(""),
                'Lat-Long-Azimuth (First CellID)': empty_series(""),
                'Crime': empty_series(""),
                'Circle': df['Circle'].fillna("").astype(str),
                'Operator': df['operator'].fillna("").astype(str),
                'CallForward': df.get('CallForward', empty_series()).apply(self.clean_text),
                'LRN': df.get('LRN', empty_series()).apply(self.clean_text),
                'Location': empty_series("")
            })
            
            return mapping_df
            
        except Exception as e:
            logging.error(f"Error creating Mapping sheet: {e}")
            return pd.DataFrame()

    def create_summary_sheet(self, df):
        """Create Summary sheet - comprehensive per-contact summary"""
        try:
            # Filter out records with blank counterparty
            df_clean = df[df['Counterparty'].astype(str).str.strip() != ""].copy()
            
            if len(df_clean) == 0:
                return pd.DataFrame()
            
            # Group by CdrNo and Counterparty
            grp = df_clean.groupby(['CdrNo', 'Counterparty'], dropna=False)
            
            # Aggregate statistics
            agg = grp.agg(
                Total_Calls=('Counterparty', 'count'),
                Total_Duration=('Duration', 'sum'),
                First_dt=('start_dt', 'min'),
                Last_dt=('start_dt', 'max'),
                Total_Days=('DateStr', lambda s: s.nunique()),
                Total_CellIds=('FirstCellID', lambda s: s.nunique()),
                Total_IMEI=('IMEI', lambda s: s.nunique()),
                Total_IMSI=('IMSI', lambda s: s.nunique()),
                OutCalls=('CallTypeStd', lambda s: (s == 'CALL_OUT').sum()),
                InCalls=('CallTypeStd', lambda s: (s == 'CALL_IN').sum()),
                OutSms=('CallTypeStd', lambda s: (s == 'SMS_OUT').sum()),
                InSms=('CallTypeStd', lambda s: (s == 'SMS_IN').sum()),
                RoamCalls=('Circle', lambda s: s.replace("", np.nan).notna().sum())
            ).reset_index()
            
            # Get provider information (most frequent operator per counterparty)
            operator_map = {}
            for cp, sub in df_clean.groupby('Counterparty'):
                try:
                    operator_map[cp] = sub['operator'].astype(str).mode().iat[0]
                except Exception:
                    operator_map[cp] = ""
            
            # Format dates and times
            def fmt_date_dMonY(d):
                if pd.isna(d) or d is None:
                    return ""
                try:
                    return pd.to_datetime(d).strftime("%d/%b/%Y")
                except Exception:
                    return ""
            
            def fmt_time_HMS(t):
                if pd.isna(t) or t is None:
                    return ""
                try:
                    import datetime as _dt
                    if isinstance(t, _dt.time):
                        return t.strftime("%H:%M:%S")
                    if isinstance(t, _dt.date) and not isinstance(t, _dt.datetime):
                        return ""
                    return pd.to_datetime(t).strftime("%H:%M:%S")
                except Exception:
                    return ""
            
            # Build comprehensive summary
            summary_df = pd.DataFrame({
                'CdrNo': agg['CdrNo'],
                'B Party': agg['Counterparty'].apply(self.clean_text),
                'Provider': agg['Counterparty'].map(operator_map).fillna(""),
                'Type': "",
                'Total Calls': agg['Total_Calls'],
                'Out Calls': agg['OutCalls'],
                'In Calls': agg['InCalls'],
                'Out Sms': agg['OutSms'],
                'In Sms': agg['InSms'],
                'Other Calls': 0,
                'Roam Calls': agg['RoamCalls'],
                'Roam Sms': 0,
                'Total Duration': agg['Total_Duration'],
                'Total Days': agg['Total_Days'],
                'Total CellIds': agg['Total_CellIds'],
                'Total Imei': agg['Total_IMEI'],
                'Total Imsi': agg['Total_IMSI'],
                'First Call Date': agg['First_dt'].dt.date.apply(fmt_date_dMonY),
                'First Call Time': agg['First_dt'].dt.time.apply(fmt_time_HMS),
                'Last Call Date': agg['Last_dt'].dt.date.apply(fmt_date_dMonY),
                'Last Call Time': agg['Last_dt'].dt.time.apply(fmt_time_HMS)
            })
            
            return summary_df
            
        except Exception as e:
            logging.error(f"Error creating Summary sheet: {e}")
            return pd.DataFrame()

    def create_max_calls_sheet(self, df):
        """Create MaxCalls sheet - top counterparties by call count with provider info"""
        try:
            df_clean = df[df['Counterparty'].astype(str).str.strip() != ""].copy()
            if len(df_clean) == 0:
                return pd.DataFrame()
                
            # Group by CdrNo and Counterparty
            max_calls = df_clean.groupby(['CdrNo', 'Counterparty'], dropna=False).size().reset_index(name='Total Calls')
            max_calls = max_calls[max_calls['Counterparty'].astype(str).str.strip() != ""]
            
            # Get provider information
            operator_map = {}
            for cp, sub in df_clean.groupby('Counterparty'):
                try:
                    operator_map[cp] = sub['operator'].astype(str).mode().iat[0]
                except Exception:
                    operator_map[cp] = ""
            
            max_calls['Provider'] = max_calls['Counterparty'].map(operator_map).fillna("")
            max_calls = max_calls.rename(columns={'Counterparty': 'B Party'})
            max_calls = max_calls[max_calls['B Party'].astype(str).str.strip() != ""]
            
            # Reorder columns
            max_calls = max_calls[['CdrNo', 'B Party', 'Total Calls', 'Provider']]
            
            return max_calls
            
        except Exception as e:
            logging.error(f"Error creating MaxCalls sheet: {e}")
            return pd.DataFrame()

    def create_max_duration_sheet(self, df):
        """Create MaxDuration sheet - top counterparties by duration with provider info"""
        try:
            df_clean = df[df['Counterparty'].astype(str).str.strip() != ""].copy()
            if len(df_clean) == 0:
                return pd.DataFrame()
                
            # Group by CdrNo and Counterparty
            max_dur = df_clean.groupby(['CdrNo', 'Counterparty'], dropna=False)['Duration'].sum().reset_index(name='Total Duration')
            max_dur = max_dur[max_dur['Counterparty'].astype(str).str.strip() != ""]
            
            # Get provider information
            operator_map = {}
            for cp, sub in df_clean.groupby('Counterparty'):
                try:
                    operator_map[cp] = sub['operator'].astype(str).mode().iat[0]
                except Exception:
                    operator_map[cp] = ""
            
            max_dur['Provider'] = max_dur['Counterparty'].map(operator_map).fillna("")
            max_dur = max_dur.rename(columns={'Counterparty': 'B Party'})
            max_dur = max_dur[max_dur['B Party'].astype(str).str.strip() != ""]
            
            # Reorder columns
            max_dur = max_dur[['CdrNo', 'B Party', 'Total Duration', 'Provider']]
            
            return max_dur
            
        except Exception as e:
            logging.error(f"Error creating MaxDuration sheet: {e}")
            return pd.DataFrame()

    def create_max_stay_sheet(self, df):
        """Create MaxStay sheet - comprehensive location analysis"""
        try:
            # Filter for records with cell ID information
            sub = df[df['FirstCellID'].astype(str).str.strip() != ""].copy()
            if sub.empty:
                return pd.DataFrame()
            
            # Group by CdrNo and FirstCellID
            g = sub.groupby(['CdrNo', 'FirstCellID'], dropna=False).agg(
                Total_Calls=('FirstCellID', 'count'),
                Days=('DateStr', lambda s: s.nunique()),
                TowerAddress=('FirstCellAddr', lambda s: s.replace("", np.nan).dropna().iloc[0] if len(s.replace("", np.nan).dropna()) > 0 else ""),
                First_dt=('start_dt', 'min'),
                Last_dt=('start_dt', 'max'),
                Roaming=('Circle', lambda s: s.replace("", np.nan).dropna().iloc[0] if len(s.replace("", np.nan).dropna()) > 0 else "")
            ).reset_index()
            
            # Format dates and times
            def fmt_date_dMonY(d):
                if pd.isna(d) or d is None:
                    return ""
                try:
                    return pd.to_datetime(d).strftime("%d/%b/%Y")
                except Exception:
                    return ""
            
            def fmt_time_HMS(t):
                if pd.isna(t) or t is None:
                    return ""
                try:
                    import datetime as _dt
                    if isinstance(t, _dt.time):
                        return t.strftime("%H:%M:%S")
                    if isinstance(t, _dt.date) and not isinstance(t, _dt.datetime):
                        return ""
                    return pd.to_datetime(t).strftime("%H:%M:%S")
                except Exception:
                    return ""
            
            # Build comprehensive MaxStay sheet
            max_stay_df = pd.DataFrame({
                'CdrNo': g['CdrNo'],
                'Cell ID': g['FirstCellID'],
                'Total Calls': g['Total_Calls'],
                'Days': g['Days'],
                'Tower Address': g['TowerAddress'].apply(self.clean_text),
                'Latitude': "",
                'Longitude': "",
                'Azimuth': "",
                'Roaming': g['Roaming'],
                'First Call Date': g['First_dt'].dt.date.apply(fmt_date_dMonY),
                'First Call Time': g['First_dt'].dt.time.apply(fmt_time_HMS),
                'Last Call Date': g['Last_dt'].dt.date.apply(fmt_date_dMonY),
                'Last Call Time': g['Last_dt'].dt.time.apply(fmt_time_HMS)
            })
            
            return max_stay_df
            
        except Exception as e:
            logging.error(f"Error creating MaxStay sheet: {e}")
            return pd.DataFrame()

    def create_roaming_analysis_sheets(self, df):
        """Create roaming-related analysis sheets"""
        try:
            # OtherStateContactSummary - comprehensive circle analysis
            sub = df.copy()
            sub['Circle'] = sub['Circle'].astype(str)
            sub = sub[sub['Circle'].str.strip() != ""]
            
            if sub.empty:
                other_state_summary = pd.DataFrame()
            else:
                g = sub.groupby(['CdrNo', 'Circle'], dropna=False).agg(
                    TotalCalls=('Circle', 'count'),
                    OutCalls=('CallTypeStd', lambda s: (s == 'CALL_OUT').sum()),
                    InCalls=('CallTypeStd', lambda s: (s == 'CALL_IN').sum()),
                    OutSms=('CallTypeStd', lambda s: (s == 'SMS_OUT').sum()),
                    InSms=('CallTypeStd', lambda s: (s == 'SMS_IN').sum()),
                    TotalDuration=('Duration', 'sum')
                ).reset_index()
                
                other_state_summary = pd.DataFrame({
                    'CdrNo': g['CdrNo'],
                    'Circle': g['Circle'],
                    'Total Calls': g['TotalCalls'],
                    'Out Calls': g['OutCalls'],
                    'In Calls': g['InCalls'],
                    'Out Sms': g['OutSms'],
                    'In Sms': g['InSms'],
                    'Other Calls': 0,
                    'Total Duration': g['TotalDuration']
                })
                other_state_summary = other_state_summary[other_state_summary['Total Calls'] > 0]
            
            # RoamingPeriod - detailed roaming periods
            if sub.empty:
                roaming_periods = pd.DataFrame()
            else:
                g = sub.groupby(['CdrNo', 'Circle'], dropna=False).agg(
                    TotalCalls=('Circle', 'count'),
                    Days=('DateStr', lambda s: s.nunique()),
                    FirstLoc=('FirstCellAddr', lambda s: s.replace("", np.nan).dropna().iloc[0] if len(s.replace("", np.nan).dropna()) > 0 else ""),
                    LastLoc=('LastCellAddr', lambda s: s.replace("", np.nan).dropna().iloc[-1] if len(s.replace("", np.nan).dropna()) > 0 else ""),
                    First_dt=('start_dt', 'min'),
                    Last_dt=('start_dt', 'max'),
                    OutCalls=('CallTypeStd', lambda s: (s == 'CALL_OUT').sum()),
                    InCalls=('CallTypeStd', lambda s: (s == 'CALL_IN').sum()),
                    OutSms=('CallTypeStd', lambda s: (s == 'SMS_OUT').sum()),
                    InSms=('CallTypeStd', lambda s: (s == 'SMS_IN').sum()),
                    TotalDuration=('Duration', 'sum')
                ).reset_index()
                
                # Format period
                def format_period(first_dt, last_dt):
                    if pd.isna(first_dt) or pd.isna(last_dt):
                        return ""
                    try:
                        start_str = pd.to_datetime(first_dt).strftime("%d/%b/%Y %H:%M:%S")
                        end_str = pd.to_datetime(last_dt).strftime("%d/%b/%Y %H:%M:%S")
                        return f"{start_str} to {end_str}"
                    except Exception:
                        return ""
                
                roaming_periods = pd.DataFrame({
                    'CdrNo': g['CdrNo'],
                    'Roaming': g['Circle'],
                    'Period': [format_period(f, l) for f, l in zip(g['First_dt'], g['Last_dt'])],
                    'Total Calls': g['TotalCalls'],
                    'Days': g['Days'],
                    'First Location': g['FirstLoc'].apply(self.clean_text),
                    'Last Location': g['LastLoc'].apply(self.clean_text),
                    'Out Calls': g['OutCalls'],
                    'In Calls': g['InCalls'],
                    'Out Sms': g['OutSms'],
                    'In Sms': g['InSms'],
                    'Other Calls': 0,
                    'Total Duration': g['TotalDuration']
                })
            
            return other_state_summary, roaming_periods
            
        except Exception as e:
            logging.error(f"Error creating roaming analysis sheets: {e}")
            return pd.DataFrame(), pd.DataFrame()

    def create_device_analysis_sheets(self, df):
        """Create IMEI and IMSI analysis sheets"""
        try:
            # Format dates and times
            def fmt_date_dMonY(d):
                if pd.isna(d) or d is None:
                    return ""
                try:
                    return pd.to_datetime(d).strftime("%d/%b/%Y")
                except Exception:
                    return ""
            
            def fmt_time_HMS(t):
                if pd.isna(t) or t is None:
                    return ""
                try:
                    import datetime as _dt
                    if isinstance(t, _dt.time):
                        return t.strftime("%H:%M:%S")
                    if isinstance(t, _dt.date) and not isinstance(t, _dt.datetime):
                        return ""
                    return pd.to_datetime(t).strftime("%H:%M:%S")
                except Exception:
                    return ""
            
            # IMEIPeriod - comprehensive IMEI analysis
            imei_df = df[df['IMEI'].astype(str).str.strip() != ''].copy()
            if len(imei_df) > 0:
                g = imei_df.groupby(['CdrNo', 'IMEI'], dropna=False).agg(
                    Total_Calls=('IMEI', 'count'),
                    Days=('DateStr', lambda s: s.nunique()),
                    First_dt=('start_dt', 'min'),
                    Last_dt=('start_dt', 'max'),
                    Total_Duration=('Duration', 'sum'),
                    OutCalls=('CallTypeStd', lambda s: (s == 'CALL_OUT').sum()),
                    InCalls=('CallTypeStd', lambda s: (s == 'CALL_IN').sum()),
                    OutSms=('CallTypeStd', lambda s: (s == 'SMS_OUT').sum()),
                    InSms=('CallTypeStd', lambda s: (s == 'SMS_IN').sum())
                ).reset_index()
                
                # Deduplicate by keeping highest activity per IMEI
                imei_period = g.loc[g.groupby(['CdrNo', 'IMEI'])['Total_Calls'].idxmax()].reset_index(drop=True)
                
                imei_period = pd.DataFrame({
                    'CdrNo': imei_period['CdrNo'],
                    'IMEI': imei_period['IMEI'],
                    'Total Calls': imei_period['Total_Calls'],
                    'Days': imei_period['Days'],
                    'Out Calls': imei_period['OutCalls'],
                    'In Calls': imei_period['InCalls'],
                    'Out Sms': imei_period['OutSms'],
                    'In Sms': imei_period['InSms'],
                    'Other Calls': 0,
                    'Total Duration': imei_period['Total_Duration'],
                    'First Call Date': imei_period['First_dt'].dt.date.apply(fmt_date_dMonY),
                    'First Call Time': imei_period['First_dt'].dt.time.apply(fmt_time_HMS),
                    'Last Call Date': imei_period['Last_dt'].dt.date.apply(fmt_date_dMonY),
                    'Last Call Time': imei_period['Last_dt'].dt.time.apply(fmt_time_HMS)
                })
            else:
                imei_period = pd.DataFrame()
            
            # IMSIPeriod - comprehensive IMSI analysis
            imsi_df = df[df['IMSI'].astype(str).str.strip() != ''].copy()
            if len(imsi_df) > 0:
                g = imsi_df.groupby(['CdrNo', 'IMSI'], dropna=False).agg(
                    Total_Calls=('IMSI', 'count'),
                    Days=('DateStr', lambda s: s.nunique()),
                    First_dt=('start_dt', 'min'),
                    Last_dt=('start_dt', 'max'),
                    Total_Duration=('Duration', 'sum'),
                    OutCalls=('CallTypeStd', lambda s: (s == 'CALL_OUT').sum()),
                    InCalls=('CallTypeStd', lambda s: (s == 'CALL_IN').sum()),
                    OutSms=('CallTypeStd', lambda s: (s == 'SMS_OUT').sum()),
                    InSms=('CallTypeStd', lambda s: (s == 'SMS_IN').sum())
                ).reset_index()
                
                # Deduplicate by keeping highest activity per IMSI  
                imsi_period = g.loc[g.groupby(['CdrNo', 'IMSI'])['Total_Calls'].idxmax()].reset_index(drop=True)
                
                imsi_period = pd.DataFrame({
                    'CdrNo': imsi_period['CdrNo'],
                    'IMSI': imsi_period['IMSI'],
                    'Total Calls': imsi_period['Total_Calls'],
                    'Days': imsi_period['Days'],
                    'Out Calls': imsi_period['OutCalls'],
                    'In Calls': imsi_period['InCalls'],
                    'Out Sms': imsi_period['OutSms'],
                    'In Sms': imsi_period['InSms'],
                    'Other Calls': 0,
                    'Total Duration': imsi_period['Total_Duration'],
                    'First Call Date': imsi_period['First_dt'].dt.date.apply(fmt_date_dMonY),
                    'First Call Time': imsi_period['First_dt'].dt.time.apply(fmt_time_HMS),
                    'Last Call Date': imsi_period['Last_dt'].dt.date.apply(fmt_date_dMonY),
                    'Last Call Time': imsi_period['Last_dt'].dt.time.apply(fmt_time_HMS)
                })
            else:
                imsi_period = pd.DataFrame()
            
            return imei_period, imsi_period
            
        except Exception as e:
            logging.error(f"Error creating device analysis sheets: {e}")
            return pd.DataFrame(), pd.DataFrame()

    def create_night_day_analysis(self, df):
        """Create night/day analysis sheets with comprehensive columns"""
        try:
            # Debug: Check if IsNight column exists and has values
            logging.info(f"Total records: {len(df)}")
            if 'IsNight' in df.columns:
                night_count = df['IsNight'].sum()
                logging.info(f"Night records found: {night_count}")
            else:
                logging.warning("IsNight column not found")
            
            # Night analysis (18:00-06:00) - be more flexible with the filter
            if 'IsNight' in df.columns:
                night_df = df[df['IsNight'] == True].copy()
                day_df = df[df['IsNight'] == False].copy()
            else:
                # Fallback: create IsNight based on Hour if available
                if 'Hour' in df.columns:
                    df['IsNight'] = df['Hour'].apply(lambda h: h >= 18 or h < 6 if pd.notna(h) else False)
                    night_df = df[df['IsNight'] == True].copy()
                    day_df = df[df['IsNight'] == False].copy()
                else:
                    # If no time data, split randomly for testing
                    night_df = df.copy()
                    day_df = df.copy()
            
            # Helper function to create mapping format
            def create_mapping_format(data_df):
                if len(data_df) == 0:
                    return pd.DataFrame()
                    
                n = len(data_df)
                empty_series = lambda val="": pd.Series([val]*n)
                
                return pd.DataFrame({
                    'CdrNo': data_df['CdrNo'].fillna('').astype(str),
                    'B Party': data_df['Counterparty'].apply(self.clean_text),
                    'Date': data_df['DateStr'].fillna('').astype(str),
                    'Time': data_df['TimeStr'].fillna('').astype(str),
                    'Duration': data_df['Duration'].fillna(0).astype('Int64'),
                    'Call Type': data_df['CallTypeStd'].fillna('').astype(str),
                    'First Cell ID': data_df['FirstCellID'].apply(self.clean_text),
                    'First Cell ID Address': data_df['FirstCellAddr'].apply(self.clean_text),
                    'Last Cell ID': data_df['LastCellID'].apply(self.clean_text),
                    'Last Cell ID Address': data_df['LastCellAddr'].apply(self.clean_text),
                    'IMEI': data_df['IMEI'].apply(self.clean_text),
                    'IMEI Manufacturer': empty_series(""),
                    'Device Type': empty_series(""),
                    'IMSI': data_df['IMSI'].apply(self.clean_text),
                    'Roaming': data_df['Circle'].fillna("").astype(str),
                    'B Party Provider': empty_series(""),
                    'Main City(First CellID)': data_df['FirstCellCity'].apply(self.clean_text),
                    'Sub City(First CellID)': empty_series(""),
                    'Lat-Long-Azimuth (First CellID)': empty_series(""),
                    'Crime': empty_series(""),
                    'Circle': data_df['Circle'].fillna("").astype(str),
                    'Operator': data_df['operator'].fillna("").astype(str),
                    'CallForward': data_df.get('CallForward', empty_series()).apply(self.clean_text),
                    'LRN': data_df.get('LRN', empty_series()).apply(self.clean_text),
                    'Location': empty_series("")
                })
            
            # Helper function to create MaxStay format
            def create_maxstay_format(data_df):
                if len(data_df) == 0:
                    return pd.DataFrame()
                    
                sub = data_df[data_df['FirstCellID'].astype(str).str.strip() != ""].copy()
                if sub.empty:
                    return pd.DataFrame()
                
                g = sub.groupby(['CdrNo', 'FirstCellID'], dropna=False).agg(
                    Total_Calls=('FirstCellID', 'count'),
                    Days=('DateStr', lambda s: s.nunique()),
                    TowerAddress=('FirstCellAddr', lambda s: s.replace("", np.nan).dropna().iloc[0] if len(s.replace("", np.nan).dropna()) > 0 else ""),
                    First_dt=('start_dt', 'min'),
                    Last_dt=('start_dt', 'max'),
                    Roaming=('Circle', lambda s: s.replace("", np.nan).dropna().iloc[0] if len(s.replace("", np.nan).dropna()) > 0 else "")
                ).reset_index()
                
                def fmt_date_dMonY(d):
                    if pd.isna(d) or d is None:
                        return ""
                    try:
                        return pd.to_datetime(d).strftime("%d/%b/%Y")
                    except Exception:
                        return ""
                
                def fmt_time_HMS(t):
                    if pd.isna(t) or t is None:
                        return ""
                    try:
                        import datetime as _dt
                        if isinstance(t, _dt.time):
                            return t.strftime("%H:%M:%S")
                        if isinstance(t, _dt.date) and not isinstance(t, _dt.datetime):
                            return ""
                        return pd.to_datetime(t).strftime("%H:%M:%S")
                    except Exception:
                        return ""
                
                return pd.DataFrame({
                    'CdrNo': g['CdrNo'],
                    'Cell ID': g['FirstCellID'],
                    'Total Calls': g['Total_Calls'],
                    'Days': g['Days'],
                    'Tower Address': g['TowerAddress'].apply(self.clean_text),
                    'Latitude': "",
                    'Longitude': "",
                    'Azimuth': "",
                    'Roaming': g['Roaming'],
                    'First Call Date': g['First_dt'].dt.date.apply(fmt_date_dMonY),
                    'First Call Time': g['First_dt'].dt.time.apply(fmt_time_HMS),
                    'Last Call Date': g['Last_dt'].dt.date.apply(fmt_date_dMonY),
                    'Last Call Time': g['Last_dt'].dt.time.apply(fmt_time_HMS)
                })
            
            # Create all four sheets
            night_mapping = create_mapping_format(night_df)
            night_stay = create_maxstay_format(night_df)
            day_mapping = create_mapping_format(day_df)
            day_stay = create_maxstay_format(day_df)
            
            return night_mapping, night_stay, day_mapping, day_stay
            
        except Exception as e:
            logging.error(f"Error creating night/day analysis: {e}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    def create_location_analysis_sheets(self, df):
        """Create comprehensive location-based analysis sheets"""
        try:
            # WorkHomeLocation - comprehensive location analysis
            if len(df) > 0 and 'FirstCellCity' in df.columns:
                sub = df[df['FirstCellCity'].astype(str).str.strip() != ""].copy()
                if not sub.empty:
                    g = sub.groupby(['CdrNo', 'FirstCellCity'], dropna=False).agg(
                        Total_Calls=('FirstCellCity', 'count'),
                        Days=('DateStr', lambda s: s.nunique()),
                        Out_Calls=('CallTypeStd', lambda s: (s == 'CALL_OUT').sum()),
                        In_Calls=('CallTypeStd', lambda s: (s == 'CALL_IN').sum()),
                        Out_Sms=('CallTypeStd', lambda s: (s == 'SMS_OUT').sum()),
                        In_Sms=('CallTypeStd', lambda s: (s == 'SMS_IN').sum()),
                        Total_Duration=('Duration', 'sum'),
                        First_dt=('start_dt', 'min'),
                        Last_dt=('start_dt', 'max'),
                        Tower_Address=('FirstCellAddr', lambda s: s.replace("", np.nan).dropna().iloc[0] if len(s.replace("", np.nan).dropna()) > 0 else "")
                    ).reset_index()
                    
                    def fmt_date_dMonY(d):
                        if pd.isna(d) or d is None:
                            return ""
                        try:
                            return pd.to_datetime(d).strftime("%d/%b/%Y")
                        except Exception:
                            return ""
                    
                    def fmt_time_HMS(t):
                        if pd.isna(t) or t is None:
                            return ""
                        try:
                            import datetime as _dt
                            if isinstance(t, _dt.time):
                                return t.strftime("%H:%M:%S")
                            if isinstance(t, _dt.date) and not isinstance(t, _dt.datetime):
                                return ""
                            return pd.to_datetime(t).strftime("%H:%M:%S")
                        except Exception:
                            return ""
                    
                    work_home_locations = pd.DataFrame({
                        'CdrNo': g['CdrNo'],
                        'Cell ID': g['FirstCellCity'],
                        'Total Calls': g['Total_Calls'],
                        'Days': g['Days'],
                        'Out Calls': g['Out_Calls'],
                        'In Calls': g['In_Calls'],
                        'Out Sms': g['Out_Sms'],
                        'In Sms': g['In_Sms'],
                        'Other Calls': 0,
                        'Total Duration': g['Total_Duration'],
                        'Tower Address': g['Tower_Address'].apply(self.clean_text),
                        'Latitude': "",
                        'Longitude': "",
                        'Azimuth': "",
                        'First Call Date': g['First_dt'].dt.date.apply(fmt_date_dMonY),
                        'First Call Time': g['First_dt'].dt.time.apply(fmt_time_HMS),
                        'Last Call Date': g['Last_dt'].dt.date.apply(fmt_date_dMonY),
                        'Last Call Time': g['Last_dt'].dt.time.apply(fmt_time_HMS)
                    })
                else:
                    work_home_locations = pd.DataFrame()
            else:
                work_home_locations = pd.DataFrame()
            
            # HomeLocationBasedonDayFirstand - day-time location analysis
            day_df = df[df['IsNight'] == False].copy()
            if len(day_df) > 0 and 'FirstCellCity' in day_df.columns:
                sub = day_df[day_df['FirstCellCity'].astype(str).str.strip() != ""].copy()
                if not sub.empty:
                    # Get first and last records for each location
                    location_groups = []
                    for location, group in sub.groupby('FirstCellCity'):
                        first_record = group.loc[group['start_dt'].idxmin()] if 'start_dt' in group.columns else group.iloc[0]
                        last_record = group.loc[group['start_dt'].idxmax()] if 'start_dt' in group.columns else group.iloc[-1]
                        
                        location_groups.append({
                            'CdrNo': first_record.get('CdrNo', ''),
                            'Cell ID': location,
                            'Total Calls': len(group),
                            'Days': group['DateStr'].nunique() if 'DateStr' in group.columns else 1,
                            'First Record': f"{first_record.get('DateStr', '')} {first_record.get('TimeStr', '')}".strip(),
                            'Last Record': f"{last_record.get('DateStr', '')} {last_record.get('TimeStr', '')}".strip(),
                            'Tower Address': first_record.get('FirstCellAddr', ''),
                            'Latitude': "",
                            'Longitude': "",
                            'Azimuth': ""
                        })
                    
                    home_locations = pd.DataFrame(location_groups)
                    home_locations['Tower Address'] = home_locations['Tower Address'].apply(self.clean_text)
                else:
                    home_locations = pd.DataFrame()
            else:
                home_locations = pd.DataFrame()
            
            return work_home_locations, home_locations
            
        except Exception as e:
            logging.error(f"Error creating location analysis sheets: {e}")
            return pd.DataFrame(), pd.DataFrame()

    def create_isd_calls_sheet(self, df):
        """Create ISDCalls sheet - comprehensive international calls analysis"""
        try:
            # Identify ISD calls (international calls) - more flexible logic
            def is_international(number):
                if pd.isna(number) or str(number).strip() == '':
                    return False
                num_str = str(number).strip()
                
                # Remove common prefixes and clean
                num_clean = num_str.replace('+', '').replace('-', '').replace(' ', '')
                
                # Check for international patterns:
                # 1. Starts with 00 (international prefix)
                # 2. Starts with + 
                # 3. Very long numbers (>12 digits)
                # 4. Numbers that don't start with 91 (India) but are long
                # 5. Country codes like 1 (US), 44 (UK), 86 (China), etc.
                
                is_intl = (
                    num_str.startswith('00') or 
                    num_str.startswith('+') or
                    len(num_clean) > 12 or
                    (len(num_clean) > 10 and not num_clean.startswith('91')) or
                    # Common country codes at start
                    (len(num_clean) >= 10 and num_clean[:1] in ['1'] and len(num_clean) == 11) or  # US/Canada
                    (len(num_clean) >= 10 and num_clean[:2] in ['44', '86', '33', '49', '39', '81']) or  # UK, China, France, Germany, Italy, Japan
                    (len(num_clean) >= 10 and num_clean[:3] in ['971', '966', '974']) or  # UAE, Saudi, Qatar
                    # Any number starting with non-Indian mobile prefixes
                    (len(num_clean) >= 8 and not num_clean.startswith(('91', '6', '7', '8', '9')))
                )
                
                return is_intl
            
            logging.info(f"Analyzing {len(df)} total records for ISD calls")
            
            # Filter for calls only
            calls_df = df[df['CallTypeStd'].str.startswith('CALL')].copy()
            logging.info(f"Found {len(calls_df)} call records for ISD analysis")
            
            if len(calls_df) > 0:
                calls_df['IsISD'] = calls_df['Counterparty'].apply(is_international)
                isd_count = calls_df['IsISD'].sum()
                logging.info(f"Identified {isd_count} ISD calls")
                
                # Sample some numbers for debugging
                sample_numbers = calls_df['Counterparty'].head(10).tolist()
                logging.info(f"Sample counterparty numbers: {sample_numbers}")
                
                isd_calls = calls_df[calls_df['IsISD'] == True].copy()
                
                # If no ISD calls found, let's be more lenient
                if len(isd_calls) == 0:
                    logging.info("No ISD calls found with strict criteria, trying lenient approach")
                    # More lenient: any number that's not a typical 10-digit Indian number
                    def is_international_lenient(number):
                        if pd.isna(number) or str(number).strip() == '':
                            return False
                        num_str = str(number).strip()
                        num_clean = num_str.replace('+', '').replace('-', '').replace(' ', '').replace('(', '').replace(')', '')
                        
                        # Any number that doesn't look like a standard Indian mobile (10 digits starting with 6,7,8,9)
                        if len(num_clean) != 10:
                            return True
                        if not num_clean.isdigit():
                            return True
                        if not num_clean.startswith(('6', '7', '8', '9')):
                            return True
                        return False
                    
                    calls_df['IsISD'] = calls_df['Counterparty'].apply(is_international_lenient)
                    isd_count_lenient = calls_df['IsISD'].sum()
                    logging.info(f"Lenient approach found {isd_count_lenient} potential ISD calls")
                    isd_calls = calls_df[calls_df['IsISD'] == True].copy()
                
                if len(isd_calls) > 0:
                    n = len(isd_calls)
                    empty_series = lambda val="": pd.Series([val]*n)
                    
                    isd_result = pd.DataFrame({
                        'CdrNo': isd_calls['CdrNo'].fillna('').astype(str),
                        'B Party': isd_calls['Counterparty'].apply(self.clean_text),
                        'Date': isd_calls['DateStr'].fillna('').astype(str),
                        'Time': isd_calls['TimeStr'].fillna('').astype(str),
                        'Duration': isd_calls['Duration'].fillna(0).astype('Int64'),
                        'Call Type': isd_calls['CallTypeStd'].fillna('').astype(str),
                        'First Cell ID': isd_calls['FirstCellID'].apply(self.clean_text),
                        'First Cell ID Address': isd_calls['FirstCellAddr'].apply(self.clean_text),
                        'Last Cell ID': isd_calls['LastCellID'].apply(self.clean_text),
                        'Last Cell ID Address': isd_calls['LastCellAddr'].apply(self.clean_text),
                        'IMEI': isd_calls['IMEI'].apply(self.clean_text),
                        'IMEI Manufacturer': empty_series(""),
                        'Device Type': empty_series(""),
                        'IMSI': isd_calls['IMSI'].apply(self.clean_text),
                        'Roaming': isd_calls['Circle'].fillna("").astype(str),
                        'B Party Provider': empty_series(""),
                        'Main City(First CellID)': isd_calls['FirstCellCity'].apply(self.clean_text),
                        'Sub City(First CellID)': empty_series(""),
                        'Lat-Long-Azimuth (First CellID)': empty_series(""),
                        'Crime': empty_series(""),
                        'Circle': isd_calls['Circle'].fillna("").astype(str),
                        'Operator': isd_calls['operator'].fillna("").astype(str),
                        'CallForward': isd_calls.get('CallForward', empty_series()).apply(self.clean_text),
                        'LRN': isd_calls.get('LRN', empty_series()).apply(self.clean_text),
                        'Location': empty_series("")
                    })
                else:
                    isd_result = pd.DataFrame()
            else:
                isd_result = pd.DataFrame()
            
            return isd_result
            
        except Exception as e:
            logging.error(f"Error creating ISDCalls sheet: {e}")
            return pd.DataFrame()

    def generate_excel_file(self, df, output_path):
        """Generate complete Excel file with all 16 sheets"""
        try:
            self.update_progress(5, "Initializing Excel generation...")
            
            if self.cancel_flag:
                raise Exception("Processing cancelled by user")
                
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                
                # 1. Mapping sheet
                self.update_progress(10, "Creating Mapping sheet...")
                mapping_df = self.create_mapping_sheet(df)
                mapping_df.to_excel(writer, sheet_name='Mapping', index=False)
                self.format_excel_sheet(writer, 'Mapping', mapping_df)
                
                # 2. Summary sheet
                self.update_progress(15, "Creating Summary sheet...")
                summary_df = self.create_summary_sheet(df)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                self.format_excel_sheet(writer, 'Summary', summary_df)
                
                # 3. MaxCalls sheet
                self.update_progress(20, "Creating MaxCalls sheet...")
                max_calls_df = self.create_max_calls_sheet(df)
                max_calls_df.to_excel(writer, sheet_name='MaxCalls', index=False)
                self.format_excel_sheet(writer, 'MaxCalls', max_calls_df)
                
                # 4. MaxDuration sheet
                self.update_progress(25, "Creating MaxDuration sheet...")
                max_duration_df = self.create_max_duration_sheet(df)
                max_duration_df.to_excel(writer, sheet_name='MaxDuration', index=False)
                self.format_excel_sheet(writer, 'MaxDuration', max_duration_df)
                
                # 5. MaxStay sheet
                self.update_progress(30, "Creating MaxStay sheet...")
                max_stay_df = self.create_max_stay_sheet(df)
                max_stay_df.to_excel(writer, sheet_name='MaxStay', index=False)
                self.format_excel_sheet(writer, 'MaxStay', max_stay_df)
                
                # 6-7. Roaming analysis sheets
                self.update_progress(35, "Creating roaming analysis sheets...")
                other_state_df, roaming_period_df = self.create_roaming_analysis_sheets(df)
                other_state_df.to_excel(writer, sheet_name='OtherStateContactSummary', index=False)
                self.format_excel_sheet(writer, 'OtherStateContactSummary', other_state_df)
                roaming_period_df.to_excel(writer, sheet_name='RoamingPeriod', index=False)
                self.format_excel_sheet(writer, 'RoamingPeriod', roaming_period_df)
                
                # 8-9. Device analysis sheets
                self.update_progress(45, "Creating device analysis sheets...")
                imei_period_df, imsi_period_df = self.create_device_analysis_sheets(df)
                imei_period_df.to_excel(writer, sheet_name='IMEIPeriod', index=False)
                self.format_excel_sheet(writer, 'IMEIPeriod', imei_period_df)
                imsi_period_df.to_excel(writer, sheet_name='IMSIPeriod', index=False)
                self.format_excel_sheet(writer, 'IMSIPeriod', imsi_period_df)
                
                # 10-13. Night/Day analysis sheets
                self.update_progress(60, "Creating night/day analysis sheets...")
                night_mapping_df, night_stay_df, day_mapping_df, day_stay_df = self.create_night_day_analysis(df)
                night_mapping_df.to_excel(writer, sheet_name='Night_Mapping', index=False)
                self.format_excel_sheet(writer, 'Night_Mapping', night_mapping_df)
                night_stay_df.to_excel(writer, sheet_name='Night_MaxStay', index=False)
                self.format_excel_sheet(writer, 'Night_MaxStay', night_stay_df)
                day_mapping_df.to_excel(writer, sheet_name='Day_Mapping', index=False)
                self.format_excel_sheet(writer, 'Day_Mapping', day_mapping_df)
                day_stay_df.to_excel(writer, sheet_name='Day_MaxStay', index=False)
                self.format_excel_sheet(writer, 'Day_MaxStay', day_stay_df)
                
                # 14-15. Location analysis sheets
                self.update_progress(80, "Creating location analysis sheets...")
                work_home_df, home_location_df = self.create_location_analysis_sheets(df)
                work_home_df.to_excel(writer, sheet_name='WorkHomeLocation', index=False)
                self.format_excel_sheet(writer, 'WorkHomeLocation', work_home_df)
                home_location_df.to_excel(writer, sheet_name='HomeLocationBasedonDayFirstand', index=False)
                self.format_excel_sheet(writer, 'HomeLocationBasedonDayFirstand', home_location_df)
                
                # 16. ISDCalls sheet
                self.update_progress(90, "Creating ISDCalls sheet...")
                isd_calls_df = self.create_isd_calls_sheet(df)
                isd_calls_df.to_excel(writer, sheet_name='ISDCalls', index=False)
                self.format_excel_sheet(writer, 'ISDCalls', isd_calls_df)
                
            self.update_progress(100, "Excel file generated successfully")
            logging.info(f"Successfully generated Excel file: {output_path}")
            
            return True
            
        except Exception as e:
            error_msg = f"Error generating Excel file: {str(e)}"
            logging.error(error_msg)
            raise Exception(error_msg)
