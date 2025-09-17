  #!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""CDR Processing and Standardization Logic"""

import pandas as pd
import numpy as np
import re
import os
from datetime import datetime
import logging

# constants
NIGHT_START = 21
NIGHT_END = 7

# aliases for column detection
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
        self.cancel_flag = True

    def update_progress(self, percent, message=""):
        if self.progress_callback:
            self.progress_callback(percent, message)

    def _lower_map(self, cols):
        return {str(c).lower().strip(): c for c in cols}

    def _pick(self, cols_map, df, candidates):
        for c in candidates:
            if c in cols_map:
                return df[cols_map[c]]
        return pd.Series([np.nan]*len(df), index=df.index)

    def detect_header_start(self, path):
        try:
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                for i, line in enumerate(f):
                    l = line.lower()
                    if ("calling party telephone number" in l) or ("target /a party number" in l) or ("target no" in l):
                        return i
        except Exception:
            pass
        return 0

    def load_csv_file(self, path):
        try:
            self.update_progress(10, f"Loading {os.path.basename(path)}")
            start = self.detect_header_start(path)
            df = pd.read_csv(path, engine="python", sep=",", header=0, skiprows=start, dtype=str, on_bad_lines="skip")
            df = df.loc[:, ~pd.Index(df.columns).duplicated()]
            self.update_progress(20, f"Loaded {len(df)} rows")
            return df
        except Exception as e:
            logging.error(f"Error loading {path}: {e}")
            raise

    def to_seconds(self, x):
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
        s = str(x).strip().strip("'").replace(".", "/")
        try:
            return pd.to_datetime(s, dayfirst=True, errors="coerce").date()
        except Exception:
            try:
                return pd.to_datetime(s, dayfirst=False, errors="coerce").date()
            except Exception:
                return None

    def normalize_msisdn(self, num):
        if pd.isna(num): return ""
        s = re.sub(r"\D", "", str(num))
        if s.startswith("0"):
            s = s.lstrip("0")
        if s.startswith("91") and len(s) > 10:
            s = s[2:]
        return s

    def contains_sender_code(self, s):
        return bool(s and re.search(r"[A-Za-z]", str(s)))

    def clean_text(self, s):
        if s is None or (isinstance(s, float) and np.isnan(s)): return ""
        return re.sub(r"\s+", " ", str(s)).strip()

    def is_night_hour(self, hour):
        try:
            if pd.isna(hour): return False
            h = int(hour)
            return (h >= NIGHT_START) or (h < NIGHT_END)
        except Exception:
            return False

    def standardize_rows(self, df):
        try:
            self.update_progress(30, "Standardizing DATA ...")
            cols = self._lower_map(df.columns)
            pick = lambda k: self._pick(cols, df, ALIASES[k])
            def to_int_safe(series, name=""):
                series = series.fillna("")  # treat missing as empty string
                numeric = pd.to_numeric(series, errors='coerce')
                mask_bad = numeric.isna() & series.ne("")  # rows that failed numeric conversion
                if mask_bad.any():
                    logging.warning(f"[WARN] {mask_bad.sum()} non-numeric values in column '{name}', kept as string.")
                    # Convert to object dtype and put original values back
                    numeric = numeric.astype(object)
                    numeric[mask_bad] = series[mask_bad]
                return numeric  # may be mixed int+str but no data loss

          
            df = df.applymap(lambda x: str(x).strip() if pd.notnull(x) else "")

            Raw = pd.DataFrame({
                # Big numeric fields → safe numeric conversion
                'TargetRaw': to_int_safe(pick('target_number'), name="Target Number"),
                'Araw': to_int_safe(pick('a_party'), name="A Party"),
                'Braw': to_int_safe(pick('b_party'), name="B Party"),
                'IMEI': to_int_safe(pick('imei'), name="IMEI"),
                'IMSI': to_int_safe(pick('imsi'), name="IMSI"),
                'FirstCellID': to_int_safe(pick('first_cell_id'), name="First Cell ID"),
                'LastCellID': to_int_safe(pick('last_cell_id'), name="Last Cell ID"),

                # Other fields remain as clean strings
                'CallDateRaw': pick('call_date'),
                'CallTimeRaw': pick('call_time'),
                'DurationRaw': pick('duration'),
                'CallTypeRaw': pick('call_type'),
                'TOC': pick('toc'),
                'FirstCellAddr': pick('first_cell_addr'),
                'LastCellAddr': pick('last_cell_addr'),
                'FirstCellCity': pick('first_cell_city'),
                'LastCellCity': pick('last_cell_city'),
                'FirstLatLong': self._pick(cols, df, ALIASES.get('first_lat_long', [])),
                'LastLatLong': self._pick(cols, df, ALIASES.get('last_lat_long', [])),
                'Circle': pick('circle'),
                'HomeCircle': pick('home_circle'),
                'operator': pick('operator'),
                'SMSC': self._pick(cols, df, ALIASES.get('smsc', []))
            })

            # ✅ Parse date & time safely
            Raw['DateObj'] = Raw['CallDateRaw'].apply(self.parse_date_field) #type: ignore
            Raw['TimeObj'] = Raw['CallTimeRaw'].apply(self.parse_time_field) #type: ignore
            Raw['start_dt'] = pd.to_datetime(
                [pd.NaT if (d is None or pd.isna(d) or t is None) else f"{d} {t}" #type: ignore
                for d, t in zip(Raw['DateObj'], Raw['TimeObj'])],
                errors="coerce"
            )

            # ✅ Call duration stays numeric
            Raw['CALL_DURATION'] = pd.to_numeric(Raw['DurationRaw'], errors='coerce').fillna(0).astype('Int64')

            # ✅ Normalize numbers (convert to str first, since some may be strings)
            Raw['A_norm'] = Raw['Araw'].astype(str).apply(self.normalize_msisdn)
            Raw['B_norm'] = Raw['Braw'].astype(str).apply(self.normalize_msisdn)
            Raw['Target_norm'] = Raw['TargetRaw'].astype(str).apply(self.normalize_msisdn)

            # ✅ Pick main number safely
            if Raw['Target_norm'].replace('', np.nan).dropna().shape[0] > 0:
                top = Raw['Target_norm'].replace('', np.nan).mode().iat[0]
            else:
                combined = pd.concat([Raw['A_norm'], Raw['B_norm']], ignore_index=True)
                combined = combined[combined != ""]
                top = combined.mode().iat[0] if not combined.empty else ""
            Raw['CdrNo'] = top

          
            def derive_call_type(ct, toc):
                s = f"{ct} {toc}".lower()
                is_sms = 'sms' in s
                if is_sms:
                    d = 'IN' if any(k in s for k in ['inbound','mt','terminat','smsin','sms_in']) else 'OUT'
                else:
                    if any(k in s for k in ['incoming','mt','terminating',' in']):
                        d = 'IN'
                    elif any(k in s for k in ['outgoing','mo','originating',' out']):
                        d = 'OUT'
                    else:
                        d = 'OUT'
                return ('SMS' if is_sms else 'CALL') + '_' + d

            Raw['CallTypeStd'] = [derive_call_type(ct, toc) for ct, toc in zip(Raw['CallTypeRaw'], Raw['TOC'])]

            def pick_counterparty(row):
                a_raw, b_raw = str(row['Araw']), str(row['Braw'])
                a_num, b_num = row['A_norm'], row['B_norm']
                t = str(row['CdrNo'])
                ctype = str(row['CallTypeStd'])
                if ctype.startswith("SMS"):
                    if self.contains_sender_code(b_raw): return b_raw
                    if self.contains_sender_code(a_raw): return a_raw
                if a_num == t and b_num: return b_num
                if b_num == t and a_num: return a_num
                return b_num or a_num or b_raw or a_raw or ""

            Raw['Counterparty'] = Raw.apply(pick_counterparty, axis=1)

            Raw['CALL_DATE'] = Raw['DateObj'].apply(lambda d: d.strftime("%Y-%m-%d") if pd.notna(d) else "")
            Raw['CALL_TIME'] = Raw['TimeObj'].apply(lambda t: t.strftime("%H:%M:%S") if pd.notna(t) else "")
            Raw['IsNight'] = Raw['start_dt'].dt.hour.apply(self.is_night_hour)

            # ✅ Build standardized output (convert mixed types to str just for output)
            std = pd.DataFrame({
                'CDR Party No': Raw['CdrNo'],
                'Opposite Party No': Raw['Counterparty'],
                'Opp Party-Name': Raw['Counterparty'],
                'Opp Party-Full Address': "",
                'Opp Party-SP State': Raw['Circle'],
                'CALL_DATE': Raw['CALL_DATE'],
                'CALL_TIME': Raw['CALL_TIME'],
                'CallTypeStd': Raw['CallTypeStd'],
                'CALL_DURATION': Raw['CALL_DURATION'],
                'FIRST_CELL_ID_A': Raw['FirstCellID'].astype(str),
                'First_Cell_Site_Address': Raw['FirstCellAddr'],
                'First_Cell_Site_Name-City': Raw['FirstCellCity'],
                'First_Lat_Long': Raw['FirstLatLong'],
                'LAST_CELL_ID_A': Raw['LastCellID'].astype(str),
                'Last_Cell_Site_Address': Raw['LastCellAddr'],
                'Last_Cell_Site_Name-City': Raw['LastCellCity'],
                'Last_Lat_Long': Raw['LastLatLong'],
                'ESN_IMEI_A': Raw['IMEI'].astype(str),
                'IMSI_A': Raw['IMSI'].astype(str),
                'CUST_TYPE': "",
                'SMSC_CENTER': Raw['SMSC'],
                'Home Circle': Raw['HomeCircle'],
                'ROAM_CIRCLE': Raw['Circle'],
                'Opp Party-Activation Date': "",
                'Opp Party-Service Provider': Raw['operator'],
                'ID': pd.RangeIndex(start=1, stop=len(Raw)+1).astype(int)
            })

            std['start_dt'] = Raw['start_dt']
            std['IsNight'] = Raw['IsNight']
            std['DurationSeconds'] = Raw['CALL_DURATION']
            return std.reset_index(drop=True)

        except Exception as e:
            logging.error(f"Error standardizing rows: {e}")
            raise

    def process_files(self, file_paths):
        try:
            self.update_progress(5, "Starting processing files...")
            all_dfs = []
            for i, p in enumerate(file_paths):
                if self.cancel_flag: raise Exception("Cancelled")
                raw = self.load_csv_file(p)
                std = self.standardize_rows(raw)
                all_dfs.append(std)
            combined = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()
            self.update_progress(100, f"Processing complete: {len(combined)} records")
            return combined
        except Exception as e:
            logging.error(f"Error processing files: {e}")
            raise