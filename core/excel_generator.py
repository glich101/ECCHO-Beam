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
        """Create Mapping sheet - basic call/SMS records"""
        try:
            cols = ['DateStr','TimeStr','Counterparty','CallTypeStd','Duration','FirstCellCity','LastCellCity','IMEI']
            mapping_df = self.safe_reindex_columns(df, cols)
            
            # Clean and format data
            for col in mapping_df.columns:
                if col in mapping_df.columns:
                    mapping_df[col] = mapping_df[col].astype(str).str.strip()
            
            return mapping_df
            
        except Exception as e:
            logging.error(f"Error creating Mapping sheet: {e}")
            return pd.DataFrame()

    def create_summary_sheet(self, df):
        """Create Summary sheet - aggregated statistics"""
        try:
            summary_data = []
            
            # Basic statistics
            total_records = len(df)
            calls_df = df[df['CallTypeStd'].str.startswith('CALL')]
            sms_df = df[df['CallTypeStd'].str.startswith('SMS')]
            
            summary_data.append(['Total Records', total_records])
            summary_data.append(['Total Calls', len(calls_df)])
            summary_data.append(['Total SMS', len(sms_df)])
            summary_data.append(['Incoming Calls', len(df[df['CallTypeStd'] == 'CALL_IN'])])
            summary_data.append(['Outgoing Calls', len(df[df['CallTypeStd'] == 'CALL_OUT'])])
            summary_data.append(['Incoming SMS', len(df[df['CallTypeStd'] == 'SMS_IN'])])
            summary_data.append(['Outgoing SMS', len(df[df['CallTypeStd'] == 'SMS_OUT'])])
            
            # Duration statistics
            if len(calls_df) > 0:
                total_duration = calls_df['Duration'].sum()
                avg_duration = calls_df['Duration'].mean()
                summary_data.append(['Total Call Duration (seconds)', total_duration])
                summary_data.append(['Average Call Duration (seconds)', f"{avg_duration:.2f}"])
            
            # Date range
            if 'start_dt' in df.columns and not df['start_dt'].isna().all():
                min_date = df['start_dt'].min()
                max_date = df['start_dt'].max()
                summary_data.append(['Date Range Start', min_date.strftime('%Y-%m-%d %H:%M:%S') if pd.notna(min_date) else 'N/A'])
                summary_data.append(['Date Range End', max_date.strftime('%Y-%m-%d %H:%M:%S') if pd.notna(max_date) else 'N/A'])
            
            # Unique counts
            summary_data.append(['Unique Counterparties', df['Counterparty'].nunique()])
            summary_data.append(['Unique IMEIs', df['IMEI'].nunique()])
            summary_data.append(['Unique Cell Sites', df['FirstCellCity'].nunique()])
            
            return pd.DataFrame(summary_data, columns=['Metric', 'Value'])
            
        except Exception as e:
            logging.error(f"Error creating Summary sheet: {e}")
            return pd.DataFrame()

    def create_max_calls_sheet(self, df):
        """Create MaxCalls sheet - top counterparties by call count"""
        try:
            calls_df = df[df['CallTypeStd'].str.startswith('CALL')]
            if len(calls_df) == 0:
                return pd.DataFrame()
                
            max_calls = calls_df.groupby('Counterparty').size().reset_index(name='CallCount')
            max_calls = max_calls.sort_values('CallCount', ascending=False).head(50)
            
            return max_calls
            
        except Exception as e:
            logging.error(f"Error creating MaxCalls sheet: {e}")
            return pd.DataFrame()

    def create_max_duration_sheet(self, df):
        """Create MaxDuration sheet - top counterparties by duration"""
        try:
            calls_df = df[df['CallTypeStd'].str.startswith('CALL')]
            if len(calls_df) == 0:
                return pd.DataFrame()
                
            max_dur = calls_df.groupby('Counterparty')['Duration'].sum().reset_index()
            max_dur = max_dur.sort_values('Duration', ascending=False).head(50)
            max_dur['Duration_Minutes'] = (max_dur['Duration'] / 60).round(2)
            
            return max_dur[['Counterparty', 'Duration', 'Duration_Minutes']]
            
        except Exception as e:
            logging.error(f"Error creating MaxDuration sheet: {e}")
            return pd.DataFrame()

    def create_max_stay_sheet(self, df):
        """Create MaxStay sheet - locations with longest stays"""
        try:
            # Group by cell location and calculate time spans
            location_groups = df.groupby('FirstCellCity')['start_dt'].agg(['min', 'max', 'count']).reset_index()
            location_groups['Duration_Hours'] = ((location_groups['max'] - location_groups['min']).dt.total_seconds() / 3600).round(2)
            location_groups = location_groups.sort_values('Duration_Hours', ascending=False).head(50)
            location_groups.columns = ['Location', 'First_Seen', 'Last_Seen', 'Record_Count', 'Duration_Hours']
            
            return location_groups
            
        except Exception as e:
            logging.error(f"Error creating MaxStay sheet: {e}")
            return pd.DataFrame()

    def create_roaming_analysis_sheets(self, df):
        """Create roaming-related analysis sheets"""
        try:
            # OtherStateContactSummary - contacts from different circles
            other_state = df[df['Circle'] != df['HomeCircle']] if 'Circle' in df.columns else pd.DataFrame()
            
            if len(other_state) > 0:
                other_state_summary = other_state.groupby(['Circle', 'Counterparty']).size().reset_index(name='ContactCount')
                other_state_summary = other_state_summary.sort_values('ContactCount', ascending=False).head(100)
            else:
                other_state_summary = pd.DataFrame()
            
            # RoamingPeriod - time periods in different circles
            if 'Circle' in df.columns and not df['start_dt'].isna().all():
                roaming_periods = df.groupby('Circle')['start_dt'].agg(['min', 'max', 'count']).reset_index()
                roaming_periods.columns = ['Circle', 'Period_Start', 'Period_End', 'Record_Count']
            else:
                roaming_periods = pd.DataFrame()
            
            return other_state_summary, roaming_periods
            
        except Exception as e:
            logging.error(f"Error creating roaming analysis sheets: {e}")
            return pd.DataFrame(), pd.DataFrame()

    def create_device_analysis_sheets(self, df):
        """Create IMEI and IMSI analysis sheets"""
        try:
            # IMEIPeriod - unique IMEI usage periods
            if 'IMEI' in df.columns and not df['IMEI'].str.strip().eq('').all():
                imei_df = df[df['IMEI'].str.strip() != ''].copy()
                if len(imei_df) > 0:
                    # Group by IMEI and CdrNo, keep highest activity
                    imei_groups = imei_df.groupby(['IMEI', 'CdrNo']).agg({
                        'start_dt': ['min', 'max'],
                        'CallTypeStd': 'count'
                    }).reset_index()
                    
                    imei_groups.columns = ['IMEI', 'CdrNo', 'First_Used', 'Last_Used', 'Activity_Count']
                    
                    # Keep only highest activity per IMEI per CdrNo
                    imei_period = imei_groups.groupby(['IMEI', 'CdrNo']).apply(
                        lambda x: x.loc[x['Activity_Count'].idxmax()]
                    ).reset_index(drop=True)
                else:
                    imei_period = pd.DataFrame()
            else:
                imei_period = pd.DataFrame()
            
            # IMSIPeriod - similar for IMSI
            if 'IMSI' in df.columns and not df['IMSI'].str.strip().eq('').all():
                imsi_df = df[df['IMSI'].str.strip() != ''].copy()
                if len(imsi_df) > 0:
                    imsi_groups = imsi_df.groupby(['IMSI', 'CdrNo']).agg({
                        'start_dt': ['min', 'max'],
                        'CallTypeStd': 'count'
                    }).reset_index()
                    
                    imsi_groups.columns = ['IMSI', 'CdrNo', 'First_Used', 'Last_Used', 'Activity_Count']
                    imsi_period = imsi_groups.groupby(['IMSI', 'CdrNo']).apply(
                        lambda x: x.loc[x['Activity_Count'].idxmax()]
                    ).reset_index(drop=True)
                else:
                    imsi_period = pd.DataFrame()
            else:
                imsi_period = pd.DataFrame()
            
            return imei_period, imsi_period
            
        except Exception as e:
            logging.error(f"Error creating device analysis sheets: {e}")
            return pd.DataFrame(), pd.DataFrame()

    def create_night_day_analysis(self, df):
        """Create night/day analysis sheets"""
        try:
            # Night analysis (18:00-06:00)
            night_df = df[df['IsNight'] == True].copy()
            day_df = df[df['IsNight'] == False].copy()
            
            # Night_Mapping
            night_cols = ['DateStr','TimeStr','Counterparty','CallTypeStd','Duration','FirstCellCity','LastCellCity','IMEI']
            night_mapping = self.safe_reindex_columns(night_df, night_cols)
            
            # Night_MaxStay
            if len(night_df) > 0:
                night_stay = night_df.groupby('FirstCellCity')['start_dt'].agg(['min', 'max', 'count']).reset_index()
                night_stay['Duration_Hours'] = ((night_stay['max'] - night_stay['min']).dt.total_seconds() / 3600).round(2)
                night_stay = night_stay.sort_values('Duration_Hours', ascending=False).head(50)
                night_stay.columns = ['Location', 'First_Seen', 'Last_Seen', 'Record_Count', 'Duration_Hours']
            else:
                night_stay = pd.DataFrame()
            
            # Day_Mapping
            day_cols = ['DateStr','TimeStr','Counterparty','CallTypeStd','Duration','FirstCellCity','LastCellCity','IMEI']
            day_mapping = self.safe_reindex_columns(day_df, day_cols)
            
            # Day_MaxStay
            if len(day_df) > 0:
                day_stay = day_df.groupby('FirstCellCity')['start_dt'].agg(['min', 'max', 'count']).reset_index()
                day_stay['Duration_Hours'] = ((day_stay['max'] - day_stay['min']).dt.total_seconds() / 3600).round(2)
                day_stay = day_stay.sort_values('Duration_Hours', ascending=False).head(50)
                day_stay.columns = ['Location', 'First_Seen', 'Last_Seen', 'Record_Count', 'Duration_Hours']
            else:
                day_stay = pd.DataFrame()
            
            return night_mapping, night_stay, day_mapping, day_stay
            
        except Exception as e:
            logging.error(f"Error creating night/day analysis: {e}")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    def create_location_analysis_sheets(self, df):
        """Create location-based analysis sheets"""
        try:
            # WorkHomeLocation - most frequent locations
            if len(df) > 0 and 'FirstCellCity' in df.columns:
                location_freq = df['FirstCellCity'].value_counts().reset_index()
                location_freq.columns = ['Location', 'Frequency']
                work_home_locations = location_freq.head(20)
            else:
                work_home_locations = pd.DataFrame()
            
            # HomeLocationBasedonDayFirstand - locations during day hours with first/last analysis
            day_df = df[df['IsNight'] == False].copy()
            if len(day_df) > 0 and 'FirstCellCity' in day_df.columns:
                day_locations = day_df.groupby('FirstCellCity').agg({
                    'start_dt': ['min', 'max', 'count'],
                    'CallTypeStd': lambda x: x.mode().iloc[0] if not x.empty else 'N/A'
                }).reset_index()
                
                day_locations.columns = ['Location', 'First_Activity', 'Last_Activity', 'Total_Records', 'Common_Activity']
                home_locations = day_locations.sort_values('Total_Records', ascending=False).head(20)
            else:
                home_locations = pd.DataFrame()
            
            return work_home_locations, home_locations
            
        except Exception as e:
            logging.error(f"Error creating location analysis sheets: {e}")
            return pd.DataFrame(), pd.DataFrame()

    def create_isd_calls_sheet(self, df):
        """Create ISDCalls sheet - international calls analysis"""
        try:
            # Identify ISD calls (international calls)
            # Assuming international numbers have country codes or specific patterns
            def is_international(number):
                if pd.isna(number) or str(number).strip() == '':
                    return False
                num_str = str(number).strip()
                # Check for common international patterns
                return (len(num_str) > 12 or 
                       num_str.startswith('00') or 
                       num_str.startswith('+') or
                       (len(num_str) > 10 and not num_str.startswith('91')))
            
            calls_df = df[df['CallTypeStd'].str.startswith('CALL')].copy()
            if len(calls_df) > 0:
                calls_df['IsISD'] = calls_df['Counterparty'].apply(is_international)
                isd_calls = calls_df[calls_df['IsISD'] == True].copy()
                
                if len(isd_calls) > 0:
                    isd_cols = ['DateStr', 'TimeStr', 'Counterparty', 'CallTypeStd', 'Duration', 'FirstCellCity']
                    isd_result = self.safe_reindex_columns(isd_calls, isd_cols)
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
