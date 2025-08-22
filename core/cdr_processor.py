#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CDR Processor - Core data processing logic
Maintains the original processing logic while adding robustness
"""

import pandas as pd
import numpy as np
import re
import os
from datetime import datetime
import logging

# Configuration constants
NIGHT_START = 18
NIGHT_END = 6

# Alias map for common column name variants (from original script)
ALIASES = {
    "target_number": ["target /a party number","target no","cdr party no","target"],
    "a_party": ["calling party telephone number","a party number","msisdn a","a_number","calling party"],
    "b_party": ["called party telephone number","b party number","b party no","called party"],
    "call_forwarding": ["call forwarding","call fow no","call forwarding number","callforward"],
    "lrn_called": ["lrn called no","lrn no","translation of lrn","lrn"],
    "call_date": ["call date","date"],
    "call_time": ["call time","time","call initiation time","start time"],
    "call_term_time": ["call termination time","end time","release time"],
    "duration": ["call duration","dur(s)","duration","hold time (sec)","durations"],
    "call_type": ["call type","service type","event type","dir","toc","type of connection"],
    "toc": ["toc","type of connection"],
    "first_cell_id": ["first cell id","first cgi","first cell global id","firstcellid"],
    "last_cell_id": ["last cell id","last cgi","last cell global id","lastcellid"],
    "first_cell_addr": ["first bts location","first cell site address","first cell address","first cell site location"],
    "last_cell_addr": ["last bts location","last cell site address","last cell address","last cell site location"],
    "first_cell_city": ["first cell site name-city","first cell site name","first cell city","first site name"],
    "last_cell_city": ["last cell site name-city","last cell site name","last cell city","last site name"],
    "first_lat_long": ["first lat/long","first cgi lat/long","first latitude/longitude","first lat long"],
    "last_lat_long": ["last lat/long","last cgi lat/long","last latitude/longitude","last lat long"],
    "smsc": ["sms center number","smsc no","smsc","sms center"],
    "imei": ["imei","esn_imei_a","device imei"],
    "imsi": ["imsi","imsi_a","subscriber imsi"],
    "circle": ["roaming circle name","roam circle","circle","roaming"],
    "home_circle": ["home circle","home region","home state"],
    "operator": ["operator","service provider","sp","provider"]
}

class CDRProcessor:
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
            
    def _lower_map(self, cols):
        """Create lowercase mapping of columns with duplicate handling"""
        seen = {}
        for c in cols:
            cl = str(c).lower().strip()
            if cl not in seen:
                seen[cl] = c
            else:
                # If we encounter a duplicate lowercase key, keep the first one
                logging.debug(f"Duplicate lowercase column key '{cl}' found, keeping first occurrence")
        return seen

    def _pick(self, cols_map, df, candidates):
        """Pick column from dataframe based on candidates"""
        for c in candidates:
            if c in cols_map:
                return df[cols_map[c]]
        return pd.Series([np.nan]*len(df))

    def detect_header_start(self, path):
        """Detect where the actual header starts in CSV file"""
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                lines = f.readlines()
            for i, line in enumerate(lines):
                l = line.lower()
                if ("calling party telephone number" in l) or ("target /a party number" in l) or ("target no" in l):
                    return i
        except Exception as e:
            logging.warning(f"Error detecting header start: {e}")
        return 0

    def load_csv_file(self, path):
        """Load CSV file with robust error handling"""
        try:
            self.update_progress(10, f"Loading file: {os.path.basename(path)}")
            
            start = self.detect_header_start(path)
            df = pd.read_csv(
                path, engine="python", sep=",", header=0, skiprows=start,
                on_bad_lines="skip", dtype=str
            )
            
            # Handle duplicate columns more robustly
            original_cols = list(df.columns)
            duplicate_mask = pd.Index(df.columns).duplicated()
            
            if duplicate_mask.any():
                logging.warning(f"Found {duplicate_mask.sum()} duplicate columns in {path}")
                duplicate_cols = [col for col, is_dup in zip(original_cols, duplicate_mask) if is_dup]
                logging.warning(f"Duplicate columns: {duplicate_cols}")
                
                # Rename duplicate columns by adding suffix
                new_columns = []
                col_counts = {}
                
                for col in original_cols:
                    if col in col_counts:
                        col_counts[col] += 1
                        new_col = f"{col}_duplicate_{col_counts[col]}"
                        new_columns.append(new_col)
                        logging.info(f"Renamed duplicate column '{col}' to '{new_col}'")
                    else:
                        col_counts[col] = 0
                        new_columns.append(col)
                
                df.columns = new_columns
            
            # Final check - remove any remaining duplicates
            df = df.loc[:, ~pd.Index(df.columns).duplicated()]
            
            self.update_progress(20, f"Loaded {len(df)} records with {len(df.columns)} columns")
            logging.info(f"Successfully loaded {len(df)} records from {path}")
            logging.info(f"Column names: {list(df.columns)[:10]}{'...' if len(df.columns) > 10 else ''}")
            
            return df
            
        except Exception as e:
            error_msg = f"Error loading CSV file {path}: {str(e)}"
            logging.error(error_msg)
            raise Exception(error_msg)

    def to_seconds(self, x):
        """Convert duration to seconds"""
        if pd.isna(x): return 0
        s = str(x).strip().strip("'")
        if s == "": return 0
        if s.isdigit():
            return int(s)
        try:
            parts = s.split(":")
            if len(parts) == 3:
                h,m,ss = map(int, parts); return h*3600 + m*60 + ss
            if len(parts) == 2:
                m,ss = map(int, parts); return m*60 + ss
        except Exception:
            pass
        try:
            return int(float(s))
        except Exception:
            return 0

    def parse_time_field(self, x):
        """Parse time field from various formats"""
        s = str(x).strip().strip("'")
        for fmt in ("%H:%M:%S","%H:%M","%I:%M:%S %p","%I:%M %p"):
            try:
                return pd.to_datetime(s, format=fmt).time()
            except Exception:
                continue
        if re.fullmatch(r"\d{6}", s):
            try:
                return pd.to_datetime(s, format="%H%M%S").time()
            except:
                return None
        if re.fullmatch(r"\d{4}", s):
            try:
                return pd.to_datetime(s, format="%H%M").time()
            except:
                return None
        return None

    def parse_date_field(self, x):
        """Parse date field from various formats"""
        s = str(x).strip().strip("'").replace(".", "/")
        for dayfirst in (True, False):
            try:
                return pd.to_datetime(s, dayfirst=dayfirst, errors="raise").date()
            except Exception:
                continue
        return None

    def normalize_msisdn(self, num):
        """Normalize MSISDN number"""
        s = re.sub(r"\D", "", str(num))
        if s.startswith("0"):
            s = s[1:]
        if len(s) > 10 and s.startswith("91"):
            s = s[2:]
        return s

    def contains_sender_code(self, s):
        """Check if string contains alphanumeric sender-id"""
        if s is None: return False
        st = str(s).strip()
        if st == "": return False
        return bool(re.search(r"[A-Za-z]", st))

    def clean_text(self, s):
        """Clean and normalize text"""
        if s is None or (isinstance(s,float) and np.isnan(s)): return ""
        return re.sub(r"\s+", " ", str(s)).strip()

    def is_night_hour(self, hour):
        """Check if hour falls in night window (18:00-06:00)"""
        if pd.isna(hour): 
            return False
        # Convert to int to handle any float hours
        try:
            hour_int = int(hour)
            # Night time: 18, 19, 20, 21, 22, 23, 0, 1, 2, 3, 4, 5
            is_night = (hour_int >= NIGHT_START) or (hour_int < NIGHT_END)
            return is_night
        except (ValueError, TypeError):
            return False

    def standardize_rows(self, df):
        """Standardize raw CDR data - main processing function"""
        try:
            self.update_progress(30, "Standardizing data columns...")
            
            if self.cancel_flag:
                raise Exception("Processing cancelled by user")
            
            # Ensure no duplicate columns before processing
            if df.columns.duplicated().any():
                logging.warning("Found duplicate columns during standardization, removing...")
                df = df.loc[:, ~df.columns.duplicated()]
                
            cols = self._lower_map(df.columns)
            pick = lambda k: self._pick(cols, df, ALIASES[k])
            
            logging.info(f"Processing {len(df)} rows with {len(df.columns)} unique columns")
            logging.debug(f"Available columns: {list(df.columns)[:20]}{'...' if len(df.columns) > 20 else ''}")

            std = pd.DataFrame({
                'TargetRaw': pick('target_number').astype(str),
                'Araw': pick('a_party').astype(str),
                'Braw': pick('b_party').astype(str),
                'CallForward': pick('call_forwarding').astype(str),
                'LRN': pick('lrn_called').astype(str),
                'CallDateRaw': pick('call_date').astype(str),
                'CallTimeRaw': pick('call_time').astype(str),
                'CallTermTimeRaw': pick('call_term_time').astype(str),
                'DurationRaw': pick('duration').astype(str),
                'CallTypeRaw': pick('call_type').astype(str),
                'TOC': pick('toc').astype(str),
                'FirstCellID': pick('first_cell_id').astype(str),
                'LastCellID': pick('last_cell_id').astype(str),
                'FirstCellAddr': pick('first_cell_addr').astype(str),
                'LastCellAddr': pick('last_cell_addr').astype(str),
                'FirstCellCity': pick('first_cell_city').astype(str),
                'LastCellCity': pick('last_cell_city').astype(str),
                'FirstLatLong': self._pick(cols, df, ALIASES.get('first_lat_long', [])).astype(str),
                'LastLatLong': self._pick(cols, df, ALIASES.get('last_lat_long', [])).astype(str),
                'IMEI': pick('imei').astype(str),
                'IMSI': pick('imsi').astype(str),
                'Circle': pick('circle').astype(str),
                'HomeCircle': pick('home_circle').astype(str),
                'operator': pick('operator').astype(str),
                'SMSC': self._pick(cols, df, ALIASES.get('smsc', [])).astype(str)
            })

            self.update_progress(40, "Parsing dates and times...")
            
            # Parse dates and times
            std['DateObj'] = std['CallDateRaw'].apply(self.parse_date_field)
            std['TimeObj'] = std['CallTimeRaw'].apply(self.parse_time_field)
            std['start_dt'] = pd.to_datetime(
                [pd.NaT if (d is None or t is None) else f"{d} {t}" for d,t in zip(std['DateObj'], std['TimeObj'])],
                errors="coerce"
            )
            std['Duration'] = std['DurationRaw'].apply(self.to_seconds).astype('Int64')

            self.update_progress(50, "Determining call types...")

            # Derive call types
            def derive_call_type(ct, toc):
                s = f"{ct} {toc}".lower()
                is_sms = 'sms' in s
                if is_sms:
                    if 'smsin' in s or 'sms_in' in s or 'inbound' in s or 'terminat' in s or 'mt' in s:
                        d = 'IN'
                    elif 'smsout' in s or 'sms_out' in s or 'mo' in s or 'orig' in s:
                        d = 'OUT'
                    else:
                        d = 'IN' if (' in' in s or s.strip().endswith('in')) else 'OUT'
                else:
                    if any(k in s for k in ['incoming','mt','terminating','a_in','term',' in','-in']):
                        d = 'IN'
                    elif any(k in s for k in ['outgoing','mo','originating','a_out','orig',' out','-out']):
                        d = 'OUT'
                    else:
                        d = 'OUT'
                return ('SMS' if is_sms else 'CALL') + '_' + d
            
            std['CallTypeStd'] = [derive_call_type(ct,toc) for ct,toc in zip(std['CallTypeRaw'], std['TOC'])]

            self.update_progress(60, "Normalizing phone numbers...")

            # Normalized numbers
            std['A_norm'] = std['Araw'].apply(self.normalize_msisdn).astype(str)
            std['B_norm'] = std['Braw'].apply(self.normalize_msisdn).astype(str)
            std['Target_norm'] = std['TargetRaw'].apply(lambda x: self.normalize_msisdn(x) if pd.notna(x) else "")

            # Determine monitored number (CdrNo)
            if std['Target_norm'].str.strip().replace('', np.nan).dropna().shape[0] > 0:
                try:
                    top_target = std['Target_norm'].replace('', np.nan).mode().iat[0]
                except Exception:
                    top_target = ""
            else:
                combined = pd.concat([std['A_norm'], std['B_norm']], ignore_index=True)
                combined = combined[combined.astype(str) != ""]
                try:
                    top_target = combined.mode().iat[0]
                except Exception:
                    top_target = ""
            
            top_target = str(top_target) if top_target is not None else ""
            std['CdrNo'] = top_target

            self.update_progress(70, "Processing counterparty information...")

            # Counterparty selection with SMS sender-code logic
            def pick_counterparty(row):
                a_raw = self.clean_text(row['Araw'])
                b_raw = self.clean_text(row['Braw'])
                a_num = str(row['A_norm'])
                b_num = str(row['B_norm'])
                t = str(row['CdrNo'])
                ctype = str(row['CallTypeStd'])

                # If SMS and any side has an alphanumeric sender-id, use that as "B Party"
                if ctype.startswith("SMS"):
                    if self.contains_sender_code(b_raw):
                        return b_raw
                    if self.contains_sender_code(a_raw):
                        return a_raw

                # Otherwise choose the opposite number of target (call cases)
                if a_num == t and b_num:
                    return b_num
                if b_num == t and a_num:
                    return a_num
                if b_num:
                    return b_num
                if a_num:
                    return a_num
                return b_raw or a_raw or ""

            std['Counterparty'] = std.apply(pick_counterparty, axis=1).astype(str)

            self.update_progress(80, "Adding derived fields...")

            # Helpful derived fields
            std['Hour'] = std['start_dt'].dt.hour
            std['IsNight'] = std['Hour'].apply(self.is_night_hour)
            
            # Debug night time detection
            if 'Hour' in std.columns:
                hour_counts = std['Hour'].value_counts().sort_index()
                night_count = std['IsNight'].sum()
                logging.info(f"Hour distribution: {dict(hour_counts.head(24))}")
                logging.info(f"Total night records (18:00-06:00): {night_count}/{len(std)}")
                
                # Show some sample records with their hour and IsNight values
                sample_data = std[['start_dt', 'Hour', 'IsNight']].head(10)
                logging.info(f"Sample hour/night data:\n{sample_data}")
            
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
                    
            std['DateStr'] = std['start_dt'].dt.date.apply(fmt_date_dMonY)
            std['TimeStr'] = std['start_dt'].dt.time.apply(fmt_time_HMS)

            # Clean IMEI & IMSI strings consistently
            std['IMEI'] = std['IMEI'].apply(self.clean_text)
            std['IMSI'] = std['IMSI'].apply(self.clean_text)
            std['operator'] = std['operator'].astype(str).fillna("")

            self.update_progress(90, "Data standardization complete")
            
            # Clean up any remaining index issues and validate final dataframe
            std = std.reset_index(drop=True)
            
            # Final validation
            if std.index.duplicated().any():
                logging.warning("Found duplicate indices, cleaning up...")
                std = std[~std.index.duplicated()]
                
            # Ensure all columns are properly named and no duplicates
            if std.columns.duplicated().any():
                logging.warning("Found duplicate column names in final dataframe, cleaning up...")
                std = std.loc[:, ~std.columns.duplicated()]
            
            logging.info(f"Successfully standardized {len(std)} records with {len(std.columns)} columns")
            
            return std
            
        except Exception as e:
            error_msg = f"Error during data standardization: {str(e)}"
            logging.error(error_msg)
            raise Exception(error_msg)

    def process_files(self, file_paths):
        """Process multiple CSV files and combine them"""
        try:
            self.update_progress(5, "Starting file processing...")
            
            all_data = []
            file_count = len(file_paths)
            
            for i, file_path in enumerate(file_paths):
                if self.cancel_flag:
                    raise Exception("Processing cancelled by user")
                    
                self.update_progress(5 + (i * 85 // file_count), f"Processing file {i+1} of {file_count}")
                
                # Load and standardize each file
                raw_df = self.load_csv_file(file_path)
                std_df = self.standardize_rows(raw_df)
                all_data.append(std_df)
                
                logging.info(f"Processed file {i+1}/{file_count}: {os.path.basename(file_path)}")
            
            # Combine all data
            self.update_progress(95, "Combining all data...")
            combined_df = pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()
            
            self.update_progress(100, f"Processing complete - {len(combined_df)} total records")
            logging.info(f"Successfully processed {file_count} files, total records: {len(combined_df)}")
            
            return combined_df
            
        except Exception as e:
            error_msg = f"Error processing files: {str(e)}"
            logging.error(error_msg)
            raise Exception(error_msg)
