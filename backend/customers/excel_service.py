import os
import json
import shutil
import datetime
import re
import uuid
from pathlib import Path
from django.conf import settings
import openpyxl

# Human-readable label mappings for common rail & customer headers
LABEL_MAPPINGS = {
    'TRAIN NO.': 'Train Number',
    'TRAIN NO': 'Train Number',
    'TRAIN NUMBER': 'Train Number',
    'TRAIN': 'Train Number',
    'NO.': 'Sr. Number',
    'NO': 'Sr. Number',
    'SR NO.': 'Sr. Number',
    'SR. NO.': 'Sr. Number',
    'SR NO': 'Sr. Number',
    'NAME': 'Customer Name',
    'NAME OF PASSENGER': 'Customer Name',
    'PASSENGER NAME': 'Customer Name',
    'CUSTOMER NAME': 'Customer Name',
    'AGE': 'Age',
    'SEX': 'Gender',
    'GENDER': 'Gender',
    'SC': 'SC',
    'DOJ': 'Date of Journey',
    'DATE OF JOURNEY': 'Date of Journey',
    'TRAVEL DATE': 'Date of Journey',
    'DEP (T)': 'Departure',
    'DEPARTURE': 'Departure',
    'DEPARTURE TIME': 'Departure',
    'DEP': 'Departure',
    'ARR (T)': 'Arrival',
    'ARRIVAL': 'Arrival',
    'ARRIVAL TIME': 'Arrival',
    'ARR': 'Arrival',
    'SEAT NO/STATUS': 'Seat / Status',
    'SEAT NO.': 'Seat / Status',
    'SEAT NO': 'Seat / Status',
    'SEAT': 'Seat / Status',
    'STATUS': 'Status',
    'TICKET NO.': 'Ticket Number',
    'TICKET NO': 'Ticket Number',
    'TKT NO.': 'Ticket Number',
    'TKT NO': 'Ticket Number',
    'TICKET NUMBER': 'Ticket Number',
    'RS.': 'Amount',
    'RS': 'Amount',
    'AMOUNT': 'Amount',
    'FARE': 'Amount',
    'PRICE': 'Amount',
    'TCKT SR NO.': 'Ticket Serial Number',
    'TCKT SR NO': 'Ticket Serial Number',
    'TKT SR NO': 'Ticket Serial Number',
    'TKT SR NO.': 'Ticket Serial Number',
    'ADHAAR NO.': 'Aadhaar Number',
    'ADHAAR NO': 'Aadhaar Number',
    'AADHAAR NO.': 'Aadhaar Number',
    'AADHAAR NO': 'Aadhaar Number',
    'AADHAAR': 'Aadhaar Number',
    'ADHAAR': 'Aadhaar Number',
    'MOBILE': 'Mobile Number',
    'MOBILE NO.': 'Mobile Number',
    'MOBILE NO': 'Mobile Number',
    'ADDRESS': 'Address',
    'PNR': 'PNR Number',
    'PNR NO.': 'PNR Number',
    'COACH': 'Coach',
    'BERTH': 'Berth',
}

STATION_ABBR = {
    'HRDW': 'Haridwar',
    'HDW': 'Haridwar',
    'AHMD': 'Ahmedabad',
    'AHM': 'Ahmedabad',
    'ADI': 'Ahmedabad',
    'SAB': 'Sabarmati',
    'DEL': 'Delhi',
    'ASR': 'Asarwa',
    'UDAI': 'Udaipur',
    'UDZ': 'Udaipur',
    'AJMER': 'Ajmer',
    'AII': 'Ajmer',
    'BAR': 'Baroda',
    'BRC': 'Vadodara',
}


class ExcelDataManager:
    _instance = None

    def __init__(self):
        self.customer_dir = Path(settings.MEDIA_ROOT) / 'customer_data'
        self.temp_dir = self.customer_dir / 'temp'
        self.files_dir = self.customer_dir / 'files'
        self.customer_dir.mkdir(parents=True, exist_ok=True)
        self.temp_dir.mkdir(parents=True, exist_ok=True)
        self.files_dir.mkdir(parents=True, exist_ok=True)
        self.photos_dir = self.customer_dir / 'photos'
        self.photos_dir.mkdir(parents=True, exist_ok=True)

        self.excel_file_path = self.customer_dir / 'customers.xlsx'
        self.metadata_file_path = self.customer_dir / 'metadata.json'

        self._cached_records = None
        self._cached_grouped_customers = None
        self._cached_name_to_cluster = None
        self._cached_sheets_meta = None
        self._cached_metadata = None
        self._cached_customer_profiles = None

    @classmethod
    def get_instance(cls):
        if cls._instance is None:
            cls._instance = ExcelDataManager()
        return cls._instance

    def _format_cell_value(self, cell):
        """
        Preserve cell values cleanly without loss:
        - Dates formatted cleanly (e.g. 2026-09-26)
        - Integers preserved without decimals (e.g. 19032.0 -> '19032')
        - Leading zeros preserved if string
        """
        val = cell.value
        if val is None:
            return ""

        if isinstance(val, (datetime.datetime, datetime.date)):
            if isinstance(val, datetime.datetime) and (val.hour != 0 or val.minute != 0 or val.second != 0):
                return val.strftime("%Y-%m-%d %H:%M:%S")
            return val.strftime("%Y-%m-%d")

        if isinstance(val, datetime.time):
            return val.strftime("%H:%M")

        if isinstance(val, float):
            if val.is_integer():
                return str(int(val))
            return str(val)

        return str(val).strip()

    def _deduplicate_headers(self, raw_headers):
        """
        Deduplicates header names while maintaining original base names:
        e.g. ['TICKET NO.', 'TICKET NO.'] -> ['TICKET NO.', 'TICKET NO. (2)']
        """
        seen = {}
        unique_headers = []
        for h in raw_headers:
            base = h.strip() if h else "Column"
            if base in seen:
                seen[base] += 1
                unique_headers.append(f"{base} ({seen[base]})")
            else:
                seen[base] = 1
                unique_headers.append(base)
        return unique_headers

    def _get_display_label(self, original_header):
        """
        Maps technical / abbreviated Excel column names to human-friendly display labels.
        Supports duplicate numbered columns:
        e.g. 'TICKET NO. (2)' -> 'Ticket Number 2'
        """
        clean_upper = original_header.strip().upper()
        if clean_upper in LABEL_MAPPINGS:
            return LABEL_MAPPINGS[clean_upper]

        # Handle numbered duplicates like 'TICKET NO. (2)'
        match = re.match(r'^(.*?)\s*\((\d+)\)$', clean_upper)
        if match:
            base, num = match.groups()
            base_clean = base.strip()
            if base_clean in LABEL_MAPPINGS:
                return f"{LABEL_MAPPINGS[base_clean]} {num}"
            return f"{base_clean.title()} {num}"

        return original_header.title()

    def _format_display_value(self, key, value):
        """
        Formats friendly display value:
        - NEVER displays header names as values
        - RS / Amount -> ₹ 7,420.00
        - DOJ -> 03 Oct 2026 or 03 Oct
        """
        if not value:
            return ""

        val_str = str(value).strip()
        key_upper = str(key).strip().upper()

        if val_str.upper() == key_upper:
            return ""

        header_rejections = {
            'TRAIN NO.', 'TRAIN NO', 'TRAIN NUMBER', 'NO.', 'SR NO.', 'SR. NO.',
            'NAME', 'CUSTOMER NAME', 'PASSENGER NAME', 'NAME OF PASSENGER',
            'AGE', 'SEX', 'GENDER', 'SC', 'DOJ', 'DATE OF JOURNEY',
            'DEP (T)', 'ARR (T)', 'DEPARTURE', 'ARRIVAL', 'SEAT NO/STATUS',
            'TICKET NO.', 'RS.', 'AMOUNT', 'TCKT SR NO.', 'TICKET NO. (2)'
        }
        if val_str.upper() in header_rejections:
            return ""

        # Format Currency / Amount for RS. or AMOUNT columns
        if any(w in key_upper for w in ['RS', 'AMOUNT', 'FARE', 'PRICE']):
            clean_num = val_str.replace('₹', '').replace(',', '').strip()
            try:
                num_val = float(clean_num)
                return f"₹ {num_val:,.2f}"
            except ValueError:
                pass

        # Format Date of Journey (DOJ)
        if any(w in key_upper for w in ['DOJ', 'DATE OF JOURNEY', 'TRAVEL DATE']):
            for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y", "%d-%m-%Y"):
                try:
                    dt = datetime.datetime.strptime(val_str.split()[0], fmt.split()[0])
                    return dt.strftime("%d %b %Y")
                except Exception:
                    pass

        return val_str

    def _extract_sheet_title_metadata(self, sheet, header_row_idx, sheet_name="", filename=""):
        """
        Inspects rows above the detected header row for sheet-level title metadata
        such as train number ("Train no.19032"), route ("Haridwar to Ahmedabad on 3RD OCT"),
        date ("3rd Oct"), departure ("Haridwar"), and arrival ("Ahmedabad").
        """
        sheet_train = ""
        sheet_route = ""
        sheet_date = ""
        sheet_dep = ""
        sheet_arr = ""

        title_texts = []
        title_rows = []
        for r_idx in range(min(header_row_idx if header_row_idx > 0 else 5, 10)):
            row_cells = list(sheet.iter_rows(min_row=r_idx + 1, max_row=r_idx + 1, values_only=False))
            if row_cells:
                r_vals = [self._format_cell_value(c) for c in row_cells[0] if self._format_cell_value(c)]
                if r_vals:
                    title_rows.append(r_vals)
                    for val in r_vals:
                        title_texts.append(val)

        # 1. Search for Train Number in title texts
        for text in title_texts:
            train_m = re.search(r'train\s*(?:no\.?|number)?\s*[:.-]?\s*(\d{4,6})', text, re.I)
            if train_m:
                sheet_train = train_m.group(1).strip()
                break

        # Fallback train number from filename if not in title (e.g. 19032)
        if not sheet_train and filename:
            train_m = re.search(r'\b(\d{5})\b', filename)
            if train_m:
                sheet_train = train_m.group(1).strip()

        # 2. Search for (HEAD) rows with [..., Origin, Destination]
        for r_vals in title_rows:
            if any('head' in str(v).lower() for v in r_vals) and len(r_vals) >= 3:
                c1 = r_vals[-2].strip()
                c2 = r_vals[-1].strip()
                if len(c1) >= 2 and len(c2) >= 2:
                    sheet_dep = STATION_ABBR.get(c1.upper(), c1.title())
                    sheet_arr = STATION_ABBR.get(c2.upper(), c2.title())
                    sheet_route = f"{sheet_dep} → {sheet_arr}"
                    break

        # 3. Search for Route in title texts (e.g. "Haridwar to Ahmedabad on 3RD OCT")
        if not sheet_route:
            for text in title_texts:
                route_m = re.search(r'([A-Za-z\s]+?)\s+(?:to|-|->)\s+([A-Za-z\s]+?)(?:\s+on\b|\s*\d|\s*$)', text, re.I)
                if route_m:
                    c1 = route_m.group(1).strip()
                    c2 = route_m.group(2).strip()
                    ignore_words = ['train', 'list', 'passenger', 'sleeper', 'general', 'shri', 'passengers']
                    if not any(w in c1.lower() for w in ignore_words) and len(c1) >= 3 and len(c2) >= 3:
                        c1_clean = STATION_ABBR.get(c1.upper(), c1.title())
                        c2_clean = STATION_ABBR.get(c2.upper(), c2.title())
                        sheet_dep = c1_clean
                        sheet_arr = c2_clean
                        sheet_route = f"{c1_clean} → {c2_clean}"
                        break

        # 4. Check sheet name (e.g. "delhi to haridwar", "Asarwa to udaipur", "ASR-UDAI")
        if not sheet_route and sheet_name:
            if ' to ' in sheet_name.lower():
                parts = re.split(r'\s+to\s+', sheet_name, flags=re.I)
                if len(parts) == 2 and len(parts[0].strip()) >= 2 and len(parts[1].strip()) >= 2:
                    p1 = parts[0].strip().upper()
                    p2 = parts[1].strip().upper()
                    sheet_dep = STATION_ABBR.get(p1, parts[0].strip().title())
                    sheet_arr = STATION_ABBR.get(p2, parts[1].strip().title())
                    sheet_route = f"{sheet_dep} → {sheet_arr}"
            elif '-' in sheet_name:
                parts = sheet_name.split('-')
                if len(parts) == 2:
                    p1 = parts[0].strip().upper()
                    p2 = parts[1].strip().upper()
                    sheet_dep = STATION_ABBR.get(p1, parts[0].strip().title())
                    sheet_arr = STATION_ABBR.get(p2, parts[1].strip().title())
                    if len(sheet_dep) >= 2 and len(sheet_arr) >= 2:
                        sheet_route = f"{sheet_dep} → {sheet_arr}"

        # 5. Check filename (e.g. "TRAIN HRDW TO AHMD FINAL.xlsx")
        if not sheet_route and filename:
            m = re.search(r'([A-Za-z]+)\s*(?:-|to)\s*([A-Za-z]+)', filename, re.I)
            if m:
                p1 = m.group(1).strip().upper()
                p2 = m.group(2).strip().upper()
                c1 = STATION_ABBR.get(p1, m.group(1).strip().title())
                c2 = STATION_ABBR.get(p2, m.group(2).strip().title())
                if len(c1) >= 2 and len(c2) >= 2:
                    sheet_dep = c1
                    sheet_arr = c2
                    sheet_route = f"{c1} → {c2}"

        # 6. Search for Date of Journey in title texts and filename (e.g. "on 3RD OCT", "3rd Oct")
        def clean_date_str(val):
            if not val:
                return ""
            s = re.sub(r'(\d+)(st|nd|rd|th)', lambda m: f"{m.group(1)}{m.group(2).lower()}", val.strip(), flags=re.I)
            parts = s.split()
            if len(parts) >= 2:
                return f"{parts[0]} {parts[1].capitalize()}" + (f" {' '.join(parts[2:])}" if len(parts) > 2 else "")
            return s.title()

        for text in title_texts:
            date_m = re.search(r'(?:on\s+)?(\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]{3,9}(?:\s+\d{2,4})?)', text, re.I)
            if date_m:
                d_val = date_m.group(1).strip()
                if len(d_val) >= 3:
                    sheet_date = clean_date_str(d_val)
                    break

        if not sheet_date and filename:
            date_m = re.search(r'(?:on\s+)?(\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]{3,9}(?:\s+\d{2,4})?)', filename, re.I)
            if date_m:
                d_val = date_m.group(1).strip()
                if len(d_val) >= 3:
                    sheet_date = clean_date_str(d_val)

        return sheet_train, sheet_route, sheet_date, sheet_dep, sheet_arr

    def _is_invalid_or_header_row(self, row_dict, headers, name_col):
        """
        Globally validates whether an Excel row is a genuine customer passenger record.
        Strictly rejects:
        1. Completely blank rows or rows with only whitespace/dashes
        2. Rows with missing, invalid, or placeholder customer names
        3. Table header or sub-header rows (including repeated column labels)
        4. Section banner rows (e.g. (HEAD), (HEAD) -, TOTAL, LIST OF PASSENGERS, TICKETS DATA)
        5. Rows with zero passenger booking substance (no age, sex, seat, ticket, train, fare, sr_no, etc.)
        """
        # 1. Blank check: Ensure at least one cell has meaningful content
        non_empty_values = [
            str(v).strip() for v in row_dict.values()
            if v is not None and str(v).strip() and str(v).strip() not in ('—', '-', '--', 'N/A')
        ]
        if not non_empty_values:
            return True

        # 2. Section Banner / Total / (HEAD) keywords in any cell
        for k, v in row_dict.items():
            v_up = str(v).strip().upper()
            if '(HEAD)' in v_up or v_up.startswith('(HEAD)') or v_up.startswith('TOTAL') or 'LIST OF PASSENGER' in v_up or 'TICKETS DATA' in v_up:
                return True

        # 3. Customer name check
        raw_name = str(row_dict.get(name_col, "")).strip()
        if not raw_name or raw_name.lower() in ('none', 'nan', 'null', 'n/a', '-', '—', '--'):
            return True

        name_upper = raw_name.upper()
        header_name_rejections = {
            'NAME', 'CUSTOMER NAME', 'PASSENGER NAME', 'NAME OF PASSENGER', 'PASSENGER',
            'TRAIN NO.', 'TRAIN NO', 'TRAIN NUMBER', 'SR NO.', 'SR. NO.', 'NO.',
            'DEP (T)', 'ARR (T)', 'DOJ', 'DATE OF JOURNEY', 'SEAT NO/STATUS', 'TICKET NO.',
            'RS.', 'AMOUNT', 'FARE', 'PRICE', 'TCKT SR NO.', 'TICKET NO. (2)', 'COLUMN'
        }
        if name_upper in header_name_rejections or any(name_upper.startswith(h + ' ') for h in header_name_rejections):
            return True

        # 4. Repeated table header row
        header_match_count = 0
        for h in headers:
            v = str(row_dict.get(h, "")).strip().upper()
            h_clean = str(h).strip().upper()
            if v and (v == h_clean or v in header_name_rejections) and len(h_clean) >= 2:
                header_match_count += 1
        if header_match_count >= 2:
            return True

        # 5. Passenger booking substance check
        # A valid passenger record must have at least one booking detail:
        # Age, Sex/Gender, Seat/Status/Berth, Ticket/PNR, numeric Train Number, Fare/Amount, Passenger Sr. No., Aadhaar/Mobile, or DOJ
        has_booking_substance = False
        for k, v in row_dict.items():
            if k == name_col:
                continue
            v_str = str(v).strip()
            if not v_str or v_str in ('—', '-', '--', 'N/A', 'NONE'):
                continue
            k_up = str(k).strip().upper()

            # Age
            if 'AGE' in k_up and re.search(r'\b\d{1,3}\b', v_str):
                has_booking_substance = True
                break

            # Sex / Gender
            if ('SEX' in k_up or 'GENDER' in k_up) and v_str.upper() in ('M', 'F', 'MALE', 'FEMALE'):
                has_booking_substance = True
                break

            # Seat / Status / Berth / Coach
            if any(kw in k_up for kw in ['SEAT', 'STATUS', 'BERTH', 'COACH']) and v_str.upper() not in ('SEAT', 'STATUS', 'BERTH', 'COACH'):
                has_booking_substance = True
                break

            # Ticket Number / PNR
            if any(kw in k_up for kw in ['TICKET', 'TKT', 'TCKT', 'PNR']) and re.search(r'\d{3,}', v_str):
                has_booking_substance = True
                break

            # Train Number (numeric 4 to 6 digits)
            if 'TRAIN' in k_up and re.search(r'\d{4,6}', v_str):
                has_booking_substance = True
                break

            # Fare / Amount / RS
            if any(kw in k_up for kw in ['RS', 'AMOUNT', 'FARE', 'PRICE']) and re.search(r'\d+', v_str):
                has_booking_substance = True
                break

            # Passenger serial number (1, 2, 3...)
            if any(kw in k_up for kw in ['NO.', 'SR', 'SR. NO.']) and not any(kw in k_up for kw in ['TRAIN', 'TICKET', 'TCKT']) and v_str.isdigit():
                has_booking_substance = True
                break

            # Aadhaar or Mobile Number
            if any(kw in k_up for kw in ['ADHAAR', 'AADHAAR', 'MOBILE']) and re.search(r'\d{4,}', v_str):
                has_booking_substance = True
                break

            # Date of Journey (valid date format)
            if any(kw in k_up for kw in ['DOJ', 'DATE']) and v_str.upper() not in ('DOJ', 'DATE'):
                has_booking_substance = True
                break

            # Missing Field(s) or Reason (for Missing Data sheets)
            if any(kw in k_up for kw in ['MISSING', 'REASON']) and len(v_str) > 1:
                has_booking_substance = True
                break

        if not has_booking_substance:
            return True

        return False

    def _is_data_row_masquerading_as_header(self, row_cells):
        """
        Checks if a candidate header row is actually a passenger data row.
        e.g. [1, 19609, 'Kiritbhai T Chudasama', 64, 'M', 'SC', ...]
        """
        vals = [self._format_cell_value(c) for c in row_cells if self._format_cell_value(c)]
        if len(vals) >= 4:
            c0 = str(vals[0]).strip()
            c1 = str(vals[1]).strip()
            if (c0.isdigit() or not c0) and re.match(r'^\d{4,5}$', c1):
                return True
            for i in range(2, min(len(vals), 6)):
                v = str(vals[i]).strip()
                if v.isdigit() and 1 <= int(v) <= 110:
                    if i + 1 < len(vals) and str(vals[i+1]).strip().upper() in ('M', 'F', 'MALE', 'FEMALE'):
                        return True
        return False

    def _resolve_row_route(self, row_data, sheet_route=""):
        origin = row_data.get('ORIGIN') or row_data.get('FROM') or row_data.get('SOURCE') or ''
        dest = row_data.get('DESTINATION') or row_data.get('TO') or ''
        if origin and dest:
            o_clean = str(origin).strip()
            d_clean = str(dest).strip()
            if o_clean.upper() not in ('ORIGIN', 'FROM', 'SOURCE') and d_clean.upper() not in ('DESTINATION', 'TO'):
                return f"{o_clean.title()} → {d_clean.title()}"

        dep = row_data.get('DEP (T)') or row_data.get('DEPARTURE') or row_data.get('DEP') or ''
        arr = row_data.get('ARR (T)') or row_data.get('ARRIVAL') or row_data.get('ARR') or ''
        time_regex = r'^\d{1,2}:\d{2}(?::\d{2})?$'

        if dep and arr:
            dep_str = str(dep).strip()
            arr_str = str(arr).strip()
            header_rejects = {'DEP (T)', 'ARR (T)', 'DEPARTURE', 'ARRIVAL', 'DEP', 'ARR', 'DEP(T)', 'ARR(T)'}
            if dep_str.upper() not in header_rejects and arr_str.upper() not in header_rejects:
                if not re.match(time_regex, dep_str) and not re.match(time_regex, arr_str):
                    c1 = STATION_ABBR.get(dep_str.upper(), dep_str.title())
                    c2 = STATION_ABBR.get(arr_str.upper(), arr_str.title())
                    if len(c1) >= 2 and len(c2) >= 2:
                        return f"{c1} → {c2}"

        if sheet_route and 'dep (t)' not in sheet_route.lower() and 'arr (t)' not in sheet_route.lower() and sheet_route != "Route information not available":
            return sheet_route

        return ""

    def _is_passenger_block_sheet(self, sheet):
        """
        Detects if a sheet uses the Passenger Block format (e.g. Final_Passenger_Details.xlsx),
        where each passenger has a header line 'Name | Age: ... | Sex: ... | Aadhaar No.: ...'
        followed by journey rows.
        """
        for r in sheet.iter_rows(values_only=False, max_row=30):
            c0 = str(self._format_cell_value(r[0])).strip() if len(r) > 0 else ""
            if '|' in c0 and ('Age:' in c0 or 'Sex:' in c0 or 'Aadhaar' in c0):
                return True
        return False

    def _parse_passenger_block_sheet(self, sheet, sheet_name="", filename=""):
        """
        Parses a passenger block sheet (such as 'Passenger Details' in Final_Passenger_Details.xlsx),
        extracting all valid passenger journey records with exact row numbers and demographics attached.
        """
        headers = ['NAME', 'AGE', 'SEX', 'AADHAAR NO.', 'Date of Journey', 'Train No.', 'Departure', 'Arrival', 'Ticket No.', 'Seat No. / Status']
        data_rows = []
        current_p = None
        sub_headers = None

        all_rows = sheet.iter_rows(values_only=False)
        for row_idx, r in enumerate(all_rows):
            excel_row_num = row_idx + 1
            c0 = str(self._format_cell_value(r[0])).strip() if len(r) > 0 else ""

            # Check if row is completely empty
            has_val = any(self._format_cell_value(cell).strip() != "" for cell in r)
            if not has_val or c0.upper() == 'FINAL PASSENGER DETAILS':
                continue

            # Check if passenger header row: Name | Age: ... | Sex: ... | Aadhaar No.: ...
            if '|' in c0 and ('Age:' in c0 or 'Sex:' in c0 or 'Aadhaar' in c0):
                parts = [p.strip() for p in c0.split('|')]
                p_name = parts[0]
                p_age = ""
                p_sex = ""
                p_aadhaar = ""
                for pt in parts[1:]:
                    if pt.startswith('Age:'):
                        p_age = pt.replace('Age:', '').strip()
                    elif pt.startswith('Sex:'):
                        p_sex = pt.replace('Sex:', '').strip()
                    elif pt.startswith('Aadhaar No.:') or pt.startswith('Aadhaar'):
                        p_aadhaar = pt.split(':', 1)[-1].strip()
                if p_aadhaar in ('—', '-', '--', 'N/A', 'NA'):
                    p_aadhaar = ""

                current_p = {
                    "name": p_name,
                    "age": p_age,
                    "sex": p_sex,
                    "aadhaar": p_aadhaar
                }
                sub_headers = None
                continue

            # Check if sub-header row (e.g. Date of Journey | Train No. | ...)
            if c0.lower().startswith('date of journey'):
                sub_headers = [str(self._format_cell_value(cell)).strip() for cell in r]
                while sub_headers and not sub_headers[-1]:
                    sub_headers.pop()
                continue

            # If we have a current passenger, this is a journey row!
            if current_p is not None:
                j_headers = sub_headers or ['Date of Journey', 'Train No.', 'Departure', 'Arrival', 'Ticket No.', 'Seat No. / Status']
                row_dict = {
                    "NAME": current_p["name"],
                    "AGE": current_p["age"],
                    "SEX": current_p["sex"],
                    "AADHAAR NO.": current_p["aadhaar"]
                }
                for ci, h in enumerate(j_headers):
                    val = str(self._format_cell_value(r[ci])).strip() if ci < len(r) else ""
                    if val.upper() == h.upper():
                        val = ""
                    row_dict[h] = val

                data_rows.append({
                    "row_number": excel_row_num,
                    "data": row_dict
                })

        return 0, headers, data_rows, "", "", "", "", ""

    def detect_sheet_headers_and_rows(self, sheet, sheet_name="", filename=""):
        """
        Detects true header row for a sheet (skipping multi-row titles),
        extracts deduplicated headers and all non-empty data rows.
        Handles sheets without explicit headers seamlessly.
        """
        if self._is_passenger_block_sheet(sheet):
            return self._parse_passenger_block_sheet(sheet, sheet_name=sheet_name, filename=filename)

        rows = list(sheet.iter_rows(values_only=False, max_row=25))
        if not rows:
            return 0, [], [], "", "", "", "", ""

        best_header_idx = -1
        best_score = -1
        detected_headers = []
        keywords = {'name', 'customer', 'passenger', 'id', 'no', 'sr', 'mobile', 'address', 'ticket', 'adhaar', 'train', 'age', 'sex', 'doj', 'seat', 'dep', 'arr'}

        for idx, row in enumerate(rows):
            non_empty_cells = []
            score = 0
            for cell in row:
                val = self._format_cell_value(cell)
                if val:
                    non_empty_cells.append(val)
                    lower_val = val.lower()
                    # Only score if the cell looks like a header label (not pure number and not an explanation sentence)
                    if not val.isdigit() and len(val) < 60 and len(val.split()) <= 6 and val not in ('—', '-', '--', 'N/A', 'NA'):
                        for kw in keywords:
                            if kw in lower_val:
                                score += 6
                        # Bonus if exact name header
                        if any(lower_val == kw_name for kw_name in ['name', 'passenger name', 'customer name', 'passenger_name']):
                            score += 15

            if len(non_empty_cells) >= 2:
                score += len(non_empty_cells) * 2
                if score > best_score:
                    best_score = score
                    best_header_idx = idx
                    headers = []
                    col_counter = 1
                    for cell in row:
                        val = self._format_cell_value(cell)
                        if not val:
                            val = f"Column_{col_counter}"
                        headers.append(val)
                        col_counter += 1
                    detected_headers = headers

        # Check if the detected header row is actually a data row masquerading as header
        if best_header_idx >= 0 and best_header_idx < len(rows):
            if self._is_data_row_masquerading_as_header(rows[best_header_idx]):
                detected_headers = [
                    'NO.', 'TRAIN NO.', 'NAME', 'AGE', 'SEX', 'SC', 'DOJ', 'DEP (T)', 'ARR (T)',
                    'SEAT NO/STATUS', 'TICKET NO.', 'RS.', 'TICKET NO. (2)', 'TCKT SR NO.'
                ]
                best_header_idx = -1

        if not detected_headers and rows:
            detected_headers = [self._format_cell_value(c) or f"Column_{i+1}" for i, c in enumerate(rows[0])]
            best_header_idx = 0

        while detected_headers and detected_headers[-1].startswith("Column_"):
            detected_headers.pop()

        headers = self._deduplicate_headers(detected_headers)
        if not headers:
            return 0, [], [], "", "", "", "", ""

        sheet_train, sheet_route, sheet_date, sheet_dep, sheet_arr = self._extract_sheet_title_metadata(
            sheet=sheet,
            header_row_idx=max(best_header_idx, 0),
            sheet_name=sheet_name,
            filename=filename
        )
        if sheet_train and sheet_train.upper() in ('TRAIN NO.', 'TRAIN NO', 'TRAIN'):
            sheet_train = ""
        if sheet_route and ('dep (t)' in sheet_route.lower() or 'arr (t)' in sheet_route.lower()):
            sheet_route = ""

        # Determine primary name column for this sheet
        candidate_name_cols = [
            h for h in headers if any(w in h.lower() for w in ['name', 'customer', 'passenger'])
        ]
        name_col = candidate_name_cols[0] if candidate_name_cols else (headers[1] if len(headers) > 1 else (headers[0] if headers else ""))

        # Extract all valid data rows
        data_rows = []
        all_rows = sheet.iter_rows(values_only=False)

        for row_idx, row in enumerate(all_rows):
            excel_row_number = row_idx + 1
            if row_idx <= best_header_idx and best_header_idx >= 0:
                continue

            row_dict = {}
            is_empty = True
            for col_i, h in enumerate(headers):
                if col_i < len(row):
                    val = self._format_cell_value(row[col_i])
                    if val.strip().upper() == h.strip().upper():
                        val = ""
                    row_dict[h] = val
                    if val != "":
                        is_empty = False
                else:
                    row_dict[h] = ""

            if is_empty:
                continue

            if self._is_invalid_or_header_row(row_dict, headers, name_col):
                continue

            data_rows.append({
                "row_number": excel_row_number,
                "data": row_dict
            })

        return best_header_idx, headers, data_rows, sheet_train, sheet_route, sheet_date, sheet_dep, sheet_arr

    def parse_full_workbook(self, file_path, original_filename):
        """
        Parses ALL sheets of an Excel workbook.
        Returns multi-sheet structure, physical row count, and unique customer names count.
        """
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        sheet_names = wb.sheetnames

        sheets_data = []
        total_physical_rows = 0
        all_unique_names = set()

        for name in sheet_names:
            sheet = wb[name]
            header_idx, headers, rows, sheet_train, sheet_route, sheet_date, sheet_dep, sheet_arr = self.detect_sheet_headers_and_rows(
                sheet=sheet,
                sheet_name=name,
                filename=original_filename
            )
            records_count = len(rows)
            total_physical_rows += records_count

            candidate_name_cols = [
                h for h in headers if any(w in h.lower() for w in ['name', 'customer', 'passenger'])
            ]
            name_col = candidate_name_cols[0] if candidate_name_cols else (headers[1] if len(headers) > 1 else (headers[0] if headers else ""))

            candidate_train_cols = [
                h for h in headers if any(w in h.lower() for w in ['train no', 'train', 'train number'])
            ]
            train_col = candidate_train_cols[0] if candidate_train_cols else ""

            sheet_unique_names = set()
            sheet_preview = []
            for r in rows:
                c_name = r["data"].get(name_col, "").strip()
                if c_name:
                    norm = re.sub(r'\s+', ' ', c_name.lower())
                    all_unique_names.add(norm)
                    sheet_unique_names.add(norm)

            for r in rows[:5]:
                row_train = r["data"].get(train_col) or sheet_train or ""
                if str(row_train).strip().upper() in ('TRAIN NO.', 'TRAIN NO', 'TRAIN'):
                    row_train = ""
                row_route = self._resolve_row_route(r["data"], sheet_route)
                sheet_preview.append({
                    "record_id": f"{name}_row_{r['row_number']}",
                    "row_number": r["row_number"],
                    "name": r["data"].get(name_col, ""),
                    "train_no": row_train,
                    "train_route": row_route,
                    "doj": sheet_date,
                    "departure": sheet_dep,
                    "arrival": sheet_arr,
                    "data": r["data"]
                })

            sheets_data.append({
                "name": name,
                "header_row_index": header_idx,
                "headers": headers,
                "name_column": name_col,
                "train_column": train_col,
                "sheet_train": sheet_train,
                "sheet_route": sheet_route or "Route information not available",
                "sheet_date": sheet_date,
                "sheet_dep": sheet_dep,
                "sheet_arr": sheet_arr,
                "records_count": records_count,
                "unique_customers_count": len(sheet_unique_names),
                "preview_rows": sheet_preview
            })

        wb.close()

        return {
            "filename": original_filename,
            "sheet_count": len(sheet_names),
            "customer_count": len(all_unique_names),
            "total_records": total_physical_rows,
            "sheets": sheets_data
        }

    def save_multiple_excel_permanently(self, files_list):
        """
        Atomically saves multiple workbooks (e.g. exactly 3 files) into media/customer_data/files/
        and creates unified metadata.json.
        files_list = [{"temp_file_path": Path, "original_filename": str}, ...]
        """
        self.files_dir.mkdir(parents=True, exist_ok=True)

        # Clean out any previous files to ensure NO old data remains
        for old_f in self.files_dir.glob("*"):
            if old_f.is_file():
                try:
                    old_f.unlink()
                except Exception:
                    pass
        if self.excel_file_path.exists():
            try:
                self.excel_file_path.unlink()
            except Exception:
                pass
        backup_p = self.customer_dir / "customers.backup.xlsx"
        if backup_p.exists():
            try:
                backup_p.unlink()
            except Exception:
                pass

        saved_files_meta = []
        total_records_all = 0
        total_sheets_all = 0

        # Validate all workbooks first
        parsed_files = []
        for idx, item in enumerate(files_list):
            temp_path = Path(item["temp_file_path"])
            fname = item["original_filename"]
            info = self.parse_full_workbook(temp_path, fname)
            file_id = f"file_{idx + 1}"
            parsed_files.append({
                "file_id": file_id,
                "temp_path": temp_path,
                "filename": fname,
                "info": info
            })

        # Save each workbook permanently into files_dir
        for pf in parsed_files:
            file_id = pf["file_id"]
            orig_name = pf["filename"]
            safe_name = re.sub(r'[^\w\.-]', '_', orig_name)
            target_file_rel = f"files/{file_id}_{safe_name}"
            target_path = self.customer_dir / target_file_rel

            shutil.copy2(pf["temp_path"], target_path)

            file_sheets = pf["info"]["sheets"]
            total_sheets_all += len(file_sheets)
            total_records_all += pf["info"]["total_records"]

            saved_files_meta.append({
                "file_id": file_id,
                "filename": orig_name,
                "saved_file": target_file_rel,
                "sheet_count": len(file_sheets),
                "records_count": pf["info"]["total_records"],
                "customer_count": pf["info"]["customer_count"],
                "sheets": [
                    {
                        "name": s["name"],
                        "header_row_index": s["header_row_index"],
                        "headers": s["headers"],
                        "name_column": s["name_column"],
                        "train_column": s["train_column"],
                        "sheet_train": s["sheet_train"],
                        "sheet_route": s["sheet_route"],
                        "sheet_date": s.get("sheet_date", ""),
                        "sheet_dep": s.get("sheet_dep", ""),
                        "sheet_arr": s.get("sheet_arr", ""),
                        "records_count": s["records_count"],
                        "unique_customers_count": s.get("unique_customers_count", 0),
                        "file_id": file_id,
                        "file_name": orig_name
                    }
                    for s in file_sheets
                ]
            })

        flat_sheets = []
        for fm in saved_files_meta:
            for s in fm["sheets"]:
                flat_sheets.append(s)

        metadata = {
            "total_files": len(saved_files_meta),
            "filenames": [fm["filename"] for fm in saved_files_meta],
            "filename": saved_files_meta[0]["filename"] if len(saved_files_meta) == 1 else f"{len(saved_files_meta)} Excel Files",
            "uploaded_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "sheet_count": total_sheets_all,
            "total_records": total_records_all,
            "customer_count": 0,
            "files": saved_files_meta,
            "sheets": flat_sheets
        }

        with open(self.metadata_file_path, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, indent=2)

        # Invalidate cache and reload all data to compute true merged unique customer count
        self.invalidate_cache()
        records, grouped, sheets_meta, loaded_meta = self.load_all_data()

        metadata["customer_count"] = len(grouped)
        with open(self.metadata_file_path, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, indent=2)

        return metadata

    def save_excel_permanently(self, temp_file_path, original_filename):
        """Single file permanent save (backward compatible wrapper)."""
        return self.save_multiple_excel_permanently([
            {"temp_file_path": temp_file_path, "original_filename": original_filename}
        ])

    def replace_with_single_excel(self, source_file_path, filename="Final_Passenger_Details.xlsx"):
        """
        Completely replaces all existing customer data with the provided single Excel file.
        Purges old files and rebuilds everything from this new file only.
        """
        source_path = Path(source_file_path)
        if not source_path.exists():
            raise FileNotFoundError(f"Source file '{source_file_path}' does not exist.")

        self.files_dir.mkdir(parents=True, exist_ok=True)
        # 1. Clean out all old files
        for old_f in self.files_dir.glob("*"):
            if old_f.is_file():
                try:
                    old_f.unlink()
                except Exception:
                    pass
        if self.excel_file_path.exists():
            try:
                self.excel_file_path.unlink()
            except Exception:
                pass
        backup_path = self.customer_dir / "customers.backup.xlsx"
        if backup_path.exists():
            try:
                backup_path.unlink()
            except Exception:
                pass

        # 2. Copy source into files_dir and customers.xlsx
        safe_name = re.sub(r'[^\w\.-]', '_', filename)
        target_in_files = self.files_dir / f"file_1_{safe_name}"
        shutil.copy2(source_path, target_in_files)
        shutil.copy2(source_path, self.excel_file_path)

        # 3. Parse full workbook
        info = self.parse_full_workbook(target_in_files, filename)

        # 4. Construct metadata
        file_sheets = [
            {**s, "file_id": "file_1", "file_name": filename}
            for s in info["sheets"]
        ]

        metadata = {
            "total_files": 1,
            "filenames": [filename],
            "filename": filename,
            "uploaded_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "sheet_count": len(file_sheets),
            "total_records": info["total_records"],
            "customer_count": 0,
            "files": [
                {
                    "file_id": "file_1",
                    "filename": filename,
                    "saved_file": f"files/file_1_{safe_name}",
                    "sheet_count": len(file_sheets),
                    "records_count": info["total_records"],
                    "customer_count": info["customer_count"],
                    "sheets": file_sheets
                }
            ],
            "sheets": file_sheets
        }

        with open(self.metadata_file_path, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, indent=2)

        # 5. Invalidate and reload
        self.invalidate_cache()
        records, grouped, sheets_meta, loaded_meta = self.load_all_data()

        metadata["customer_count"] = len(grouped)
        with open(self.metadata_file_path, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, indent=2)

        return metadata

    def rebuild_dataset_metadata(self):
        """
        Re-scans the active dataset using strict record validation,
        pruning all blank, header, and invalid rows. Rebuilds accurate per-sheet
        records counts, unique customer counts, total records, and updates metadata.json.
        """
        self.invalidate_cache()
        records, grouped, sheets_meta, _ = self.load_all_data()

        if not self.metadata_file_path.exists():
            return None

        try:
            with open(self.metadata_file_path, 'r', encoding='utf-8') as f:
                metadata = json.load(f)

            sheet_counts = {}
            sheet_unique_customers = {}
            file_counts = {}
            file_unique_customers = {}

            for r in records.values():
                f_id = r.get("file_id", "file_1")
                s_name = r.get("sheet_name", "")
                norm_name = r.get("normalized_name", "")

                key = (f_id, s_name)
                sheet_counts[key] = sheet_counts.get(key, 0) + 1
                if key not in sheet_unique_customers:
                    sheet_unique_customers[key] = set()
                if norm_name:
                    sheet_unique_customers[key].add(norm_name)

                file_counts[f_id] = file_counts.get(f_id, 0) + 1
                if f_id not in file_unique_customers:
                    file_unique_customers[f_id] = set()
                if norm_name:
                    file_unique_customers[f_id].add(norm_name)

            total_recs = 0
            if "files" in metadata and isinstance(metadata["files"], list):
                for fl in metadata["files"]:
                    f_id = fl.get("file_id", "file_1")
                    fl["records_count"] = file_counts.get(f_id, 0)
                    fl["customer_count"] = len(file_unique_customers.get(f_id, set()))
                    total_recs += fl["records_count"]

                    for sh in fl.get("sheets", []):
                        s_name = sh.get("name", "")
                        sh["records_count"] = sheet_counts.get((f_id, s_name), 0)
                        sh["unique_customers_count"] = len(sheet_unique_customers.get((f_id, s_name), set()))

            if "sheets" in metadata and isinstance(metadata["sheets"], list):
                for sh in metadata["sheets"]:
                    f_id = sh.get("file_id", "file_1")
                    s_name = sh.get("name", "")
                    sh["records_count"] = sheet_counts.get((f_id, s_name), 0)
                    sh["unique_customers_count"] = len(sheet_unique_customers.get((f_id, s_name), set()))

            metadata["total_records"] = len(records)
            metadata["customer_count"] = len(grouped)

            with open(self.metadata_file_path, 'w', encoding='utf-8') as f:
                json.dump(metadata, f, indent=2)

            self.invalidate_cache()
            return metadata
        except Exception as e:
            print(f"Error rebuilding metadata: {e}")
            return None

    def get_status(self):
        """
        Returns status of saved Excel dataset across all files and sheets.
        """
        if not self.metadata_file_path.exists():
            return {
                "available": False,
                "message": "No customer data has been configured."
            }

        try:
            with open(self.metadata_file_path, 'r', encoding='utf-8') as f:
                metadata = json.load(f)

            records, grouped, sheets_meta, _ = self.load_all_data()

            total_files = metadata.get("total_files") or (len(metadata.get("files", [])) if "files" in metadata else (1 if metadata.get("filename") else 0))
            filenames = metadata.get("filenames") or ([metadata.get("filename")] if metadata.get("filename") else [])

            return {
                "available": True if (records or metadata.get("total_records", 0) > 0 or metadata.get("files") or metadata.get("saved_file")) else False,
                "total_files": total_files,
                "filenames": filenames,
                "filename": filenames[0] if len(filenames) == 1 else (f"{total_files} Excel Files" if total_files > 1 else metadata.get("filename", "")),
                "uploaded_at": metadata.get("uploaded_at", ""),
                "sheet_count": metadata.get("sheet_count", len(metadata.get("sheets", []))),
                "customer_count": len(grouped) if grouped else metadata.get("customer_count", 0),
                "total_records": len(records) if records else metadata.get("total_records", 0),
                "files": metadata.get("files", []),
                "sheets": metadata.get("sheets", [])
            }
        except Exception as e:
            return {
                "available": False,
                "error": str(e)
            }

    def normalize_customer_name(self, name):
        """
        Per Requirement: Case-insensitive grouping.
        Trim, lowercase, and collapse repeated whitespace.
        """
        if not name:
            return ""
        return re.sub(r'\s+', ' ', str(name).strip().lower())

    def _normalize_age(self, val):
        if val is None:
            return None
        val_str = str(val).strip()
        if not val_str:
            return None
        try:
            f = float(val_str)
            if 0 < f < 130:
                return int(round(f))
        except (ValueError, TypeError):
            pass
        m = re.search(r'\b(\d{1,3})\b', val_str)
        if m:
            try:
                num = int(m.group(1))
                if 0 < num < 130:
                    return num
            except (ValueError, TypeError):
                pass
        return None

    def _normalize_sex(self, val):
        if not val:
            return None
        s = str(val).strip().upper()
        if s in ('M', 'MALE', 'BOY', 'MAN') or s.startswith('M'):
            return 'M'
        if s in ('F', 'FEMALE', 'GIRL', 'WOMAN') or s.startswith('F'):
            return 'F'
        return None

    @staticmethod
    def _levenshtein(s1, s2):
        if s1 == s2:
            return 0
        if len(s1) < len(s2):
            return ExcelDataManager._levenshtein(s2, s1)
        if len(s2) == 0:
            return len(s1)
        previous_row = range(len(s2) + 1)
        for i, c1 in enumerate(s1):
            current_row = [i + 1]
            for j, c2 in enumerate(s2):
                insertions = previous_row[j + 1] + 1
                deletions = current_row[j] + 1
                substitutions = previous_row[j] + (c1 != c2)
                current_row.append(min(insertions, deletions, substitutions))
            previous_row = current_row
        return previous_row[-1]

    @staticmethod
    def _stem_token(t):
        suffixes = ('bhai', 'ben', 'kumar', 'sinh', 'singh', 'lal', 'prasad', 'das')
        for suf in suffixes:
            if t.endswith(suf) and len(t) > len(suf) + 2:
                return t[:-len(suf)]
        return t

    def _tokens_match(self, t1, t2):
        if t1 == t2:
            return True
        if len(t1) == 1 and t2.startswith(t1):
            return True
        if len(t2) == 1 and t1.startswith(t2):
            return True
        st1 = self._stem_token(t1)
        st2 = self._stem_token(t2)
        if st1 == st2 or (len(st1) == 1 and st2.startswith(st1)) or (len(st2) == 1 and st1.startswith(st2)):
            return True
        if len(t1) >= 4 and len(t2) >= 4:
            if self._levenshtein(t1, t2) <= 1:
                return True
            if len(t1) >= 6 and len(t2) >= 6 and self._levenshtein(t1, t2) <= 2:
                return True
        return False

    def are_names_similar(self, name1, name2):
        """
        Smart name similarity check supporting abbreviations, initials, suffix variants,
        minor typos, and bag-of-words / inverted order (e.g. 'Parmar C Sinh' vs 'Chandan Sinh Parmar').
        """
        n1 = re.sub(r'[\(\)\.\,\-\_]', ' ', name1.lower()).strip()
        n2 = re.sub(r'[\(\)\.\,\-\_]', ' ', name2.lower()).strip()
        toks1 = [t for t in n1.split() if t and t != 'self']
        toks2 = [t for t in n2.split() if t and t != 'self']
        if not toks1 or not toks2:
            return False
        if toks1 == toks2:
            return True
        if len(toks1) == 1 and len(toks2) == 1:
            return toks1[0] == toks2[0]

        min_toks, max_toks = (toks1, toks2) if len(toks1) <= len(toks2) else (toks2, toks1)
        matched_indices = set()
        for t_min in min_toks:
            found = False
            for idx, t_max in enumerate(max_toks):
                if idx not in matched_indices and self._tokens_match(t_min, t_max):
                    matched_indices.add(idx)
                    found = True
                    break
            if not found:
                return False
        return True

    def _is_exact_full_name_match(self, name1, name2):
        """
        High-confidence 3-token name match:
        - First name exact stem match (e.g. Hetanshri == Hetanshri)
        - Last name exact stem/levenshtein match (e.g. Chudasama == Chudasama)
        - Middle name initial-to-full match (e.g. V == Vijaybhai)
        """
        n1 = re.sub(r'[\(\)\.\,\-\_]', ' ', name1.lower()).strip()
        n2 = re.sub(r'[\(\)\.\,\-\_]', ' ', name2.lower()).strip()
        t1 = [t for t in n1.split() if t and t != 'self']
        t2 = [t for t in n2.split() if t and t != 'self']
        if len(t1) == 3 and len(t2) == 3:
            f1, m1, l1 = t1
            f2, m2, l2 = t2
            f_match = (f1 == f2 or self._stem_token(f1) == self._stem_token(f2))
            l_match = (l1 == l2 or self._stem_token(l1) == self._stem_token(l2) or self._levenshtein(l1, l2) <= 1)
            m_match = (m1 == m2 or (len(m1) == 1 and m2.startswith(m1)) or (len(m2) == 1 and m1.startswith(m2)))
            return f_match and l_match and m_match
        return False

    def _cluster_customers(self, raw_customers):
        """
        Groups raw customer entries into smart merged customer clusters.
        Criteria:
        1. Age match: overlapping ages across records (handles multi-ticket rows).
        2. Sex/Gender match: overlapping sexes across records.
        3. Names must be similar / abbreviated / inverted per are_names_similar().
        4. Or exact 3-token full name match (first name + last name + middle initial) + matching sex.
        """
        names_list = list(raw_customers.keys())
        parent = {n: n for n in names_list}

        def find(i):
            path = []
            while parent[i] != i:
                path.append(i)
                i = parent[i]
            for node in path:
                parent[node] = i
            return i

        def union(i, j):
            root_i = find(i)
            root_j = find(j)
            if root_i != root_j:
                parent[root_i] = root_j

        # Pairwise comparison
        n = len(names_list)
        for i in range(n):
            name_i = names_list[i]
            data_i = raw_customers[name_i]
            ages_i = data_i.get("ages") or ({data_i.get("age")} if data_i.get("age") else set())
            sexes_i = data_i.get("sexes") or ({data_i.get("sex")} if data_i.get("sex") else set())

            for j in range(i + 1, n):
                name_j = names_list[j]
                data_j = raw_customers[name_j]
                ages_j = data_j.get("ages") or ({data_j.get("age")} if data_j.get("age") else set())
                sexes_j = data_j.get("sexes") or ({data_j.get("sex")} if data_j.get("sex") else set())

                age_match = bool(ages_i & ages_j)
                sex_match = bool(sexes_i & sexes_j)

                should_merge = False
                if self.are_names_similar(data_i["name"], data_j["name"]) and age_match and sex_match:
                    should_merge = True
                elif self._is_exact_full_name_match(data_i["name"], data_j["name"]) and sex_match:
                    should_merge = True

                if should_merge:
                    union(name_i, name_j)

        # Build clusters
        clusters = {}
        for name in names_list:
            root = find(name)
            if root not in clusters:
                clusters[root] = []
            clusters[root].append(name)

        grouped_customers = {}
        name_to_cluster = {}

        for root, group_names in clusters.items():
            # Pick canonical display name: longest, most descriptive
            all_display_names = [raw_customers[nm]["name"] for nm in group_names]
            canonical_display_name = max(all_display_names, key=lambda s: (len(s.split()), len(s)))
            canonical_norm = self.normalize_customer_name(canonical_display_name)

            combined_record_ids = []
            combined_records = []
            aliases_set = set()
            norm_aliases_set = set()
            cluster_ages = set()
            cluster_sexes = set()

            for nm in group_names:
                c_data = raw_customers[nm]
                combined_record_ids.extend(c_data["record_ids"])
                combined_records.extend(c_data["records"])
                aliases_set.add(c_data["name"])
                norm_aliases_set.add(nm)
                if c_data.get("ages"):
                    cluster_ages.update(c_data["ages"])
                elif c_data.get("age"):
                    cluster_ages.add(c_data["age"])
                if c_data.get("sexes"):
                    cluster_sexes.update(c_data["sexes"])
                elif c_data.get("sex"):
                    cluster_sexes.add(c_data["sex"])

            cluster_age = next(iter(cluster_ages)) if cluster_ages else None
            cluster_gender = next(iter(cluster_sexes)) if cluster_sexes else None

            grouped_customers[canonical_norm] = {
                "normalized_name": canonical_norm,
                "name": canonical_display_name,
                "aliases": list(aliases_set),
                "normalized_aliases": list(norm_aliases_set),
                "age": cluster_age,
                "gender": cluster_gender,
                "record_ids": combined_record_ids,
                "records": combined_records
            }

            for nm in group_names:
                name_to_cluster[nm] = canonical_norm
                disp_nm = raw_customers[nm]["name"]
                name_to_cluster[self.normalize_customer_name(disp_nm)] = canonical_norm

            name_to_cluster[canonical_norm] = canonical_norm

        return grouped_customers, name_to_cluster

    def compute_smart_grouped_customers(self, customer_rows):
        """
        Given an iterable of dicts with {"name": str, "age": any, "sex": any},
        returns the smart grouped customer clusters dict.
        """
        raw_customers = {}
        for r in customer_rows:
            name = r.get("name", "").strip()
            if not name:
                continue
            norm = self.normalize_customer_name(name)
            if not norm or norm in ('name', 'customer name', 'passenger name', 'passenger'):
                continue
            rec_age = self._normalize_age(r.get("age"))
            rec_sex = self._normalize_sex(r.get("sex"))
            if norm not in raw_customers:
                raw_customers[norm] = {
                    "normalized_name": norm,
                    "name": name,
                    "record_ids": [],
                    "records": [],
                    "ages": {rec_age} if rec_age else set(),
                    "sexes": {rec_sex} if rec_sex else set()
                }
            else:
                if rec_age:
                    raw_customers[norm]["ages"].add(rec_age)
                if rec_sex:
                    raw_customers[norm]["sexes"].add(rec_sex)

        for c_data in raw_customers.values():
            ages = c_data["ages"]
            sexes = c_data["sexes"]
            c_data["age"] = next(iter(ages)) if len(ages) == 1 else (list(ages)[0] if ages else None)
            c_data["sex"] = next(iter(sexes)) if len(sexes) == 1 else (list(sexes)[0] if sexes else None)

        grouped, _ = self._cluster_customers(raw_customers)
        return grouped

    def load_all_data(self):
        """
        Loads all customer rows across ALL files and ALL sheets into in-memory indices:
        1. records: dict of record_id -> record dict
        2. grouped_customers: dict of norm_name -> {
               "normalized_name": norm_name,
               "name": display_name,
               "record_ids": [...],
               "records": [...]
           }
        Ensures headers like 'NAME', 'TRAIN NO.' are NEVER treated as customers.
        Maintains unique customer count vs physical row count.
        """
        if self._cached_records is not None and self._cached_grouped_customers is not None:
            return self._cached_records, self._cached_grouped_customers, self._cached_sheets_meta, self._cached_metadata

        if not self.metadata_file_path.exists():
            return {}, {}, {}, {}

        try:
            with open(self.metadata_file_path, 'r', encoding='utf-8') as f:
                metadata = json.load(f)
        except Exception:
            return {}, {}, {}, {}

        records = {}
        raw_customers = {}
        sheets_meta = {}
        total_physical_rows = 0

        # Case 1: Multi-file setup
        if "files" in metadata and isinstance(metadata["files"], list):
            for file_entry in metadata["files"]:
                file_id = file_entry.get("file_id", "file_1")
                file_name = file_entry.get("filename", "")
                saved_file_rel = file_entry.get("saved_file", "")
                file_path = self.customer_dir / saved_file_rel
                if not file_path.exists():
                    continue

                try:
                    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
                except Exception as e:
                    print(f"Error loading {file_path}: {e}")
                    continue

                for sheet_meta in file_entry.get("sheets", []):
                    sheet_name = sheet_meta["name"]
                    meta_key = f"{file_id}_{sheet_name}"
                    sheet_meta_copy = dict(sheet_meta)
                    sheet_meta_copy["file_id"] = file_id
                    sheet_meta_copy["file_name"] = file_name
                    sheets_meta[meta_key] = sheet_meta_copy
                    if sheet_name not in sheets_meta:
                        sheets_meta[sheet_name] = sheet_meta_copy

                    if sheet_name not in wb.sheetnames:
                        continue
                    sheet = wb[sheet_name]

                    header_idx = sheet_meta.get("header_row_index", 0)
                    headers = sheet_meta.get("headers", [])
                    name_col = sheet_meta.get("name_column", "")
                    train_col = sheet_meta.get("train_column", "")
                    sheet_train = sheet_meta.get("sheet_train", "")
                    if sheet_train and sheet_train.upper() in ('TRAIN NO.', 'TRAIN NO', 'TRAIN'):
                        sheet_train = ""
                    sheet_route = sheet_meta.get("sheet_route", "")
                    if sheet_route and ('dep (t)' in sheet_route.lower() or 'arr (t)' in sheet_route.lower()):
                        sheet_route = ""
                    sheet_date = sheet_meta.get("sheet_date", "")
                    sheet_dep = sheet_meta.get("sheet_dep", "")
                    sheet_arr = sheet_meta.get("sheet_arr", "")

                    if self._is_passenger_block_sheet(sheet):
                        _, _, extracted_rows, _, _, _, _, _ = self._parse_passenger_block_sheet(sheet, sheet_name=sheet_name, filename=file_name)
                        for r_item in extracted_rows:
                            excel_row_num = r_item["row_number"]
                            row_data = r_item["data"]
                            customer_name = row_data.get("NAME", "").strip()
                            if not customer_name:
                                continue
                            norm_name = self.normalize_customer_name(customer_name)
                            if not norm_name:
                                continue

                            total_physical_rows += 1
                            record_id = f"{file_id}_{sheet_name}_row_{excel_row_num}"
                            row_train = row_data.get("Train No.", "").strip()
                            row_route = self._resolve_row_route(row_data, "")

                            rec_obj = {
                                "record_id": record_id,
                                "file_id": file_id,
                                "file_name": file_name,
                                "sheet_name": sheet_name,
                                "row_number": excel_row_num,
                                "name": customer_name,
                                "normalized_name": norm_name,
                                "train_no": row_train,
                                "train_route": row_route,
                                "sheet_date": row_data.get("Date of Journey", ""),
                                "sheet_dep": row_data.get("Departure", ""),
                                "sheet_arr": row_data.get("Arrival", ""),
                                "data": row_data
                            }
                            records[record_id] = rec_obj
                            rec_age = self._normalize_age(row_data.get("AGE"))
                            rec_sex = self._normalize_sex(row_data.get("SEX"))

                            if norm_name not in raw_customers:
                                raw_customers[norm_name] = {
                                    "normalized_name": norm_name,
                                    "name": customer_name,
                                    "record_ids": [record_id],
                                    "records": [rec_obj],
                                    "ages": {rec_age} if rec_age else set(),
                                    "sexes": {rec_sex} if rec_sex else set()
                                }
                            else:
                                raw_customers[norm_name]["record_ids"].append(record_id)
                                raw_customers[norm_name]["records"].append(rec_obj)
                                if rec_age:
                                    raw_customers[norm_name]["ages"].add(rec_age)
                                if rec_sex:
                                    raw_customers[norm_name]["sexes"].add(rec_sex)
                        continue

                    all_rows = sheet.iter_rows(values_only=False)
                    for row_idx, row in enumerate(all_rows):
                        excel_row_num = row_idx + 1
                        if row_idx <= header_idx and header_idx >= 0:
                            continue

                        row_data = {}
                        is_empty = True
                        for col_i, h in enumerate(headers):
                            if col_i < len(row):
                                val = self._format_cell_value(row[col_i])
                                if val.strip().upper() == h.strip().upper():
                                    val = ""
                                row_data[h] = val
                                if val != "":
                                    is_empty = False
                            else:
                                row_data[h] = ""

                        if is_empty:
                            continue

                        if self._is_invalid_or_header_row(row_data, headers, name_col):
                            continue

                        customer_name = row_data.get(name_col, "").strip()
                        if not customer_name:
                            continue

                        norm_name = self.normalize_customer_name(customer_name)
                        if not norm_name:
                            continue

                        total_physical_rows += 1
                        record_id = f"{file_id}_{sheet_name}_row_{excel_row_num}"

                        row_train = row_data.get(train_col, "").strip() or sheet_train
                        if str(row_train).strip().upper() in ('TRAIN NO.', 'TRAIN NO', 'TRAIN'):
                            row_train = ""

                        row_route = self._resolve_row_route(row_data, sheet_route)

                        rec_obj = {
                            "record_id": record_id,
                            "file_id": file_id,
                            "file_name": file_name,
                            "sheet_name": sheet_name,
                            "row_number": excel_row_num,
                            "name": customer_name,
                            "normalized_name": norm_name,
                            "train_no": row_train or sheet_train,
                            "train_route": row_route or sheet_route,
                            "sheet_date": sheet_date,
                            "sheet_dep": sheet_dep,
                            "sheet_arr": sheet_arr,
                            "data": row_data
                        }
                        records[record_id] = rec_obj

                        rec_age = None
                        rec_sex = None
                        for k, v in row_data.items():
                            k_clean = str(k).strip().upper()
                            if rec_age is None and (k_clean == 'AGE' or k_clean.startswith('AGE ')):
                                rec_age = self._normalize_age(v)
                            if rec_sex is None and (k_clean in ('SEX', 'GENDER') or k_clean.startswith('SEX ') or k_clean.startswith('GENDER ')):
                                rec_sex = self._normalize_sex(v)

                        if norm_name not in raw_customers:
                            raw_customers[norm_name] = {
                                "normalized_name": norm_name,
                                "name": customer_name,
                                "record_ids": [record_id],
                                "records": [rec_obj],
                                "ages": {rec_age} if rec_age else set(),
                                "sexes": {rec_sex} if rec_sex else set()
                            }
                        else:
                            raw_customers[norm_name]["record_ids"].append(record_id)
                            raw_customers[norm_name]["records"].append(rec_obj)
                            if rec_age:
                                raw_customers[norm_name]["ages"].add(rec_age)
                            if rec_sex:
                                raw_customers[norm_name]["sexes"].add(rec_sex)

                wb.close()

        # Case 2: Legacy single-file setup
        elif self.excel_file_path.exists():
            sheets_meta = {s["name"]: s for s in metadata.get("sheets", [])}
            try:
                wb = openpyxl.load_workbook(self.excel_file_path, read_only=True, data_only=True)
                for sheet_name in wb.sheetnames:
                    sheet = wb[sheet_name]
                    meta = sheets_meta.get(sheet_name)
                    if not meta:
                        continue

                    header_idx = meta.get("header_row_index", 0)
                    headers = meta.get("headers", [])
                    name_col = meta.get("name_column", "")
                    train_col = meta.get("train_column", "")
                    sheet_train = meta.get("sheet_train", "")
                    if sheet_train and sheet_train.upper() in ('TRAIN NO.', 'TRAIN NO', 'TRAIN'):
                        sheet_train = ""
                    sheet_route = meta.get("sheet_route", "")
                    if sheet_route and ('dep (t)' in sheet_route.lower() or 'arr (t)' in sheet_route.lower()):
                        sheet_route = ""
                    sheet_date = meta.get("sheet_date", "")
                    sheet_dep = meta.get("sheet_dep", "")
                    sheet_arr = meta.get("sheet_arr", "")

                    all_rows = sheet.iter_rows(values_only=False)
                    for row_idx, row in enumerate(all_rows):
                        excel_row_num = row_idx + 1
                        if row_idx <= header_idx and header_idx >= 0:
                            continue

                        row_data = {}
                        is_empty = True
                        for col_i, h in enumerate(headers):
                            if col_i < len(row):
                                val = self._format_cell_value(row[col_i])
                                if val.strip().upper() == h.strip().upper():
                                    val = ""
                                row_data[h] = val
                                if val != "":
                                    is_empty = False
                            else:
                                row_data[h] = ""

                        if is_empty:
                            continue

                        if self._is_invalid_or_header_row(row_data, headers, name_col):
                            continue

                        customer_name = row_data.get(name_col, "").strip()
                        if not customer_name:
                            continue

                        norm_name = self.normalize_customer_name(customer_name)
                        if not norm_name:
                            continue

                        total_physical_rows += 1
                        record_id = f"{sheet_name}_row_{excel_row_num}"

                        row_train = row_data.get(train_col, "").strip() or sheet_train
                        if str(row_train).strip().upper() in ('TRAIN NO.', 'TRAIN NO', 'TRAIN'):
                            row_train = ""

                        row_route = self._resolve_row_route(row_data, sheet_route)

                        rec_obj = {
                            "record_id": record_id,
                            "file_id": "file_1",
                            "file_name": metadata.get("filename", "customers.xlsx"),
                            "sheet_name": sheet_name,
                            "row_number": excel_row_num,
                            "name": customer_name,
                            "normalized_name": norm_name,
                            "train_no": row_train or sheet_train,
                            "train_route": row_route or sheet_route,
                            "sheet_date": sheet_date,
                            "sheet_dep": sheet_dep,
                            "sheet_arr": sheet_arr,
                            "data": row_data
                        }
                        records[record_id] = rec_obj

                        rec_age = None
                        rec_sex = None
                        for k, v in row_data.items():
                            k_clean = str(k).strip().upper()
                            if rec_age is None and (k_clean == 'AGE' or k_clean.startswith('AGE ')):
                                rec_age = self._normalize_age(v)
                            if rec_sex is None and (k_clean in ('SEX', 'GENDER') or k_clean.startswith('SEX ') or k_clean.startswith('GENDER ')):
                                rec_sex = self._normalize_sex(v)

                        if norm_name not in raw_customers:
                            raw_customers[norm_name] = {
                                "normalized_name": norm_name,
                                "name": customer_name,
                                "record_ids": [record_id],
                                "records": [rec_obj],
                                "ages": {rec_age} if rec_age else set(),
                                "sexes": {rec_sex} if rec_sex else set()
                            }
                        else:
                            raw_customers[norm_name]["record_ids"].append(record_id)
                            raw_customers[norm_name]["records"].append(rec_obj)
                            if rec_age:
                                raw_customers[norm_name]["ages"].add(rec_age)
                            if rec_sex:
                                raw_customers[norm_name]["sexes"].add(rec_sex)
                wb.close()
            except Exception as e:
                print(f"Error loading legacy workbook: {e}")

        # Resolve consensus age and sex for raw customers
        for c_data in raw_customers.values():
            ages = c_data["ages"]
            sexes = c_data["sexes"]
            c_data["age"] = next(iter(ages)) if len(ages) == 1 else (list(ages)[0] if ages else None)
            c_data["sex"] = next(iter(sexes)) if len(sexes) == 1 else (list(sexes)[0] if sexes else None)

        # Smart customer grouping based on Name similarity + Age equality + Sex equality
        grouped_customers, name_to_cluster = self._cluster_customers(raw_customers)

        metadata["customer_count"] = len(grouped_customers)
        metadata["total_records"] = total_physical_rows

        self._cached_records = records
        self._cached_grouped_customers = grouped_customers
        self._cached_name_to_cluster = name_to_cluster
        self._cached_sheets_meta = sheets_meta
        self._cached_metadata = metadata

        return self._cached_records, self._cached_grouped_customers, self._cached_sheets_meta, self._cached_metadata

    def list_customers(self, sheet_name=None, search=None, limit=None):
        """
        Lists UNIQUE customer names across all sheets and files.
        Each customer appears ONLY ONCE.
        Searches ONLY the customer NAME field.
        """
        records, grouped, sheets_meta, metadata = self.load_all_data()
        if not grouped:
            return []

        search_lower = self.normalize_customer_name(search) if search else ""

        results = []
        for canonical_norm, grp in grouped.items():
            matching_recs = grp["records"]
            if sheet_name:
                matching_recs = [r for r in grp["records"] if r["sheet_name"] == sheet_name]
                if not matching_recs:
                    continue

            if search_lower:
                if search_lower not in canonical_norm and not any(search_lower in a for a in grp.get("normalized_aliases", [])):
                    continue

            results.append({
                "name": grp["name"],
                "normalized_name": canonical_norm,
                "record_count": len(matching_recs),
                "record_ids": [r["record_id"] for r in matching_recs]
            })

            if limit is not None and len(results) >= limit:
                break

        return results

    def search_customers(self, query, sheet_name=None):
        """
        Searches customer records across ALL files and ALL sheets by customer NAME ONLY.
        Returns UNIQUE customer names with combined record_count.
        """
        records, grouped, sheets_meta, metadata = self.load_all_data()
        q = self.normalize_customer_name(query)
        if not q or not grouped:
            return {"query": query, "count": 0, "results": []}

        results = []
        for canonical_norm, grp in grouped.items():
            if q in canonical_norm or any(q in a for a in grp.get("normalized_aliases", [])):
                matching_recs = grp["records"]
                if sheet_name:
                    matching_recs = [r for r in grp["records"] if r["sheet_name"] == sheet_name]
                    if not matching_recs:
                        continue

                results.append({
                    "name": grp["name"],
                    "normalized_name": canonical_norm,
                    "record_count": len(matching_recs)
                })

        return {
            "query": query,
            "count": len(results),
            "results": results
        }

    def get_customer_records_by_name(self, name):
        """
        Fetches ALL physical Excel records belonging to that customer name across all files and sheets.
        Preserves exact Excel column order, sheets, and rows.
        Annotates source file name and sheet name.
        """
        records, grouped, sheets_meta, metadata = self.load_all_data()
        norm_name = self.normalize_customer_name(name)
        if not norm_name or not grouped:
            return None

        canonical_norm = self._cached_name_to_cluster.get(norm_name, norm_name) if self._cached_name_to_cluster else norm_name
        if canonical_norm not in grouped:
            for c_norm, c_grp in grouped.items():
                if norm_name in c_grp.get("normalized_aliases", []):
                    canonical_norm = c_norm
                    break

        grp = grouped[canonical_norm]
        records_output = []
        missing_data_list = []

        for rec in grp["records"]:
            sheet_name = rec["sheet_name"]
            file_id = rec.get("file_id", "file_1")
            file_name = rec.get("file_name", "")

            # If this record is from Missing Data sheet, collect it for missing data panel
            if "missing data" in sheet_name.lower():
                missing_fields = rec["data"].get("Missing Field(s)") or rec["data"].get("MISSING FIELD(S)") or ""
                doj_val = rec["data"].get("Date of Journey") or rec["data"].get("DATE OF JOURNEY") or ""
                reason_val = rec["data"].get("Reason") or rec["data"].get("REASON") or ""
                missing_data_list.append({
                    "record_id": rec["record_id"],
                    "sheet_name": sheet_name,
                    "row_number": rec["row_number"],
                    "missing_fields": missing_fields,
                    "date_of_journey": doj_val,
                    "reason": reason_val
                })
                continue

            # Look up sheet metadata
            meta_key = f"{file_id}_{sheet_name}"
            sheet_meta = sheets_meta.get(meta_key) or sheets_meta.get(sheet_name) or {}
            headers = sheet_meta.get("headers", list(rec["data"].keys()))

            fields = []
            for h in headers:
                raw_val = rec["data"].get(h, "")
                if str(raw_val).strip().upper() == str(h).strip().upper():
                    raw_val = ""
                disp_label = self._get_display_label(h)
                disp_val = self._format_display_value(h, raw_val)

                fields.append({
                    "key": h,
                    "label": disp_label,
                    "value": raw_val,
                    "display_value": disp_val
                })

            existing_labels = {f["label"].upper() for f in fields}
            existing_keys = {str(f["key"]).strip().upper() for f in fields}

            # Train Number
            train_val = rec.get("train_no") or sheet_meta.get("sheet_train") or ""
            if "HEAD" in str(train_val).upper() or not any(c.isdigit() for c in str(train_val)):
                train_val = ""
            if not train_val and ("19032" in file_name or "HRDW" in file_name or "19032" in metadata.get("filename", "")):
                train_val = "19032"
            if train_val and "TRAIN NUMBER" not in existing_labels and "TRAIN NO." not in existing_keys and "TRAIN NO" not in existing_keys:
                fields.append({
                    "key": "TRAIN_NO",
                    "label": "Train Number",
                    "value": str(train_val),
                    "display_value": str(train_val)
                })

            # Date of Journey
            date_val = rec.get("sheet_date") or sheet_meta.get("sheet_date") or ""
            if not date_val and ("3RD OCT" in file_name.upper() or "3RD OCT" in sheet_meta.get("sheet_route", "").upper()):
                date_val = "3rd Oct"
            if date_val and "DATE OF JOURNEY" not in existing_labels and "DOJ" not in existing_keys:
                fields.append({
                    "key": "DOJ",
                    "label": "Date of Journey",
                    "value": str(date_val),
                    "display_value": str(date_val)
                })

            # Departure
            dep_val = rec.get("sheet_dep") or sheet_meta.get("sheet_dep") or ""
            if not dep_val and ("HRDW" in file_name.upper() or "HARIDWAR" in sheet_meta.get("sheet_route", "").upper()):
                dep_val = "Haridwar"
            if dep_val and "DEPARTURE" not in existing_labels and "DEP (T)" not in existing_keys and "DEP" not in existing_keys:
                fields.append({
                    "key": "DEP",
                    "label": "Departure",
                    "value": str(dep_val),
                    "display_value": str(dep_val)
                })

            # Arrival
            arr_val = rec.get("sheet_arr") or sheet_meta.get("sheet_arr") or ""
            if not arr_val and ("AHM" in file_name.upper() or "AHMEDABAD" in sheet_meta.get("sheet_route", "").upper()):
                arr_val = "Ahmedabad"
            if arr_val and "ARRIVAL" not in existing_labels and "ARR (T)" not in existing_keys and "ARR" not in existing_keys:
                fields.append({
                    "key": "ARR",
                    "label": "Arrival",
                    "value": str(arr_val),
                    "display_value": str(arr_val)
                })

            records_output.append({
                "record_id": rec["record_id"],
                "file_id": file_id,
                "file_name": file_name,
                "sheet_name": sheet_name,
                "row_number": rec["row_number"],
                "train_no": train_val,
                "train_route": rec.get("train_route") or sheet_meta.get("sheet_route") or (f"{dep_val} → {arr_val}" if dep_val and arr_val else ""),
                "doj": date_val,
                "departure": dep_val,
                "arrival": arr_val,
                "fields": fields,
                "row_data": rec["data"]
            })

        # Fallback if only in Missing Data
        if not records_output and missing_data_list:
            for md in missing_data_list:
                records_output.append({
                    "record_id": md["record_id"],
                    "file_id": "file_1",
                    "file_name": "",
                    "sheet_name": md["sheet_name"],
                    "row_number": md["row_number"],
                    "train_no": "",
                    "train_route": "",
                    "doj": md["date_of_journey"],
                    "departure": "",
                    "arrival": "",
                    "fields": [
                        {"key": "MISSING_FIELDS", "label": "Missing Field(s)", "value": md["missing_fields"], "display_value": md["missing_fields"]},
                        {"key": "REASON", "label": "Reason", "value": md["reason"], "display_value": md["reason"]}
                    ],
                    "row_data": {}
                })

        profile = self.get_customer_profile(canonical_norm)
        return {
            "name": grp["name"],
            "normalized_name": canonical_norm,
            "record_count": len(records_output),
            "address": profile.get("address", "—"),
            "contact_no": profile.get("contact_no", "—"),
            "passport_photo": profile.get("passport_photo"),
            "records": records_output,
            "missing_data": missing_data_list
        }

    def get_record(self, record_id):
        """
        Fetches that exact Excel row across all columns with display labels.
        """
        records, grouped, sheets_meta, metadata = self.load_all_data()
        if record_id not in records:
            return None

        rec = records[record_id]
        sheet_name = rec["sheet_name"]
        file_id = rec.get("file_id", "file_1")
        file_name = rec.get("file_name", "")

        meta_key = f"{file_id}_{sheet_name}"
        sheet_meta = sheets_meta.get(meta_key) or sheets_meta.get(sheet_name) or {}
        headers = sheet_meta.get("headers", list(rec["data"].keys()))

        columns = []
        for h in headers:
            raw_val = rec["data"].get(h, "")
            if str(raw_val).strip().upper() == str(h).strip().upper():
                raw_val = ""
            display_label = self._get_display_label(h)
            display_val = self._format_display_value(h, raw_val)

            columns.append({
                "key": h,
                "label": display_label,
                "value": raw_val,
                "display_value": display_val
            })

        existing_labels = {col["label"].upper() for col in columns}
        existing_keys = {str(col["key"]).strip().upper() for col in columns}

        train_val = rec.get("train_no") or sheet_meta.get("sheet_train") or ""
        if "HEAD" in str(train_val).upper() or not any(c.isdigit() for c in str(train_val)):
            train_val = ""
        if not train_val and ("19032" in file_name or "HRDW" in file_name or "19032" in metadata.get("filename", "")):
            train_val = "19032"
        if train_val and "TRAIN NUMBER" not in existing_labels and "TRAIN NO." not in existing_keys and "TRAIN NO" not in existing_keys:
            columns.append({
                "key": "TRAIN_NO",
                "label": "Train Number",
                "value": str(train_val),
                "display_value": str(train_val)
            })

        date_val = rec.get("sheet_date") or sheet_meta.get("sheet_date") or ""
        if not date_val and ("3RD OCT" in file_name.upper() or "3RD OCT" in sheet_meta.get("sheet_route", "").upper()):
            date_val = "3rd Oct"
        if date_val and "DATE OF JOURNEY" not in existing_labels and "DOJ" not in existing_keys:
            columns.append({
                "key": "DOJ",
                "label": "Date of Journey",
                "value": str(date_val),
                "display_value": str(date_val)
            })

        dep_val = rec.get("sheet_dep") or sheet_meta.get("sheet_dep") or ""
        if not dep_val and ("HRDW" in file_name.upper() or "HARIDWAR" in sheet_meta.get("sheet_route", "").upper()):
            dep_val = "Haridwar"
        if dep_val and "DEPARTURE" not in existing_labels and "DEP (T)" not in existing_keys and "DEP" not in existing_keys:
            columns.append({
                "key": "DEP",
                "label": "Departure",
                "value": str(dep_val),
                "display_value": str(dep_val)
            })

        arr_val = rec.get("sheet_arr") or sheet_meta.get("sheet_arr") or ""
        if not arr_val and ("AHM" in file_name.upper() or "AHMEDABAD" in sheet_meta.get("sheet_route", "").upper()):
            arr_val = "Ahmedabad"
        if arr_val and "ARRIVAL" not in existing_labels and "ARR (T)" not in existing_keys and "ARR" not in existing_keys:
            columns.append({
                "key": "ARR",
                "label": "Arrival",
                "value": str(arr_val),
                "display_value": str(arr_val)
            })

        return {
            "record_id": record_id,
            "file_id": file_id,
            "file_name": file_name,
            "sheet_name": sheet_name,
            "row_number": rec["row_number"],
            "name": rec["name"],
            "train_no": train_val,
            "train_route": rec.get("train_route") or sheet_meta.get("sheet_route") or (f"{dep_val} → {arr_val}" if dep_val and arr_val else ""),
            "doj": date_val,
            "departure": dep_val,
            "arrival": arr_val,
            "total_columns": len(columns),
            "columns": columns
        }

    def delete_record(self, record_id):
        """
        Permanently deletes the exact physical row from the specific Excel file on disk.
        """
        records, grouped, sheets_meta, metadata = self.load_all_data()
        if record_id not in records:
            raise KeyError(f"Record with ID '{record_id}' not found.")

        target_rec = records[record_id]
        target_sheet_name = target_rec["sheet_name"]
        target_row_number = int(target_rec["row_number"])
        file_id = target_rec.get("file_id")

        target_file_path = None
        if "files" in metadata and isinstance(metadata["files"], list):
            for f_meta in metadata["files"]:
                if f_meta.get("file_id") == file_id:
                    target_file_path = self.customer_dir / f_meta.get("saved_file")
                    break

        if not target_file_path or not target_file_path.exists():
            target_file_path = self.excel_file_path

        if not target_file_path.exists():
            raise FileNotFoundError(f"Active workbook file not found for record '{record_id}'.")

        wb = openpyxl.load_workbook(target_file_path)
        if target_sheet_name not in wb.sheetnames:
            wb.close()
            raise KeyError(f"Sheet '{target_sheet_name}' not found in workbook.")

        sheet = wb[target_sheet_name]
        sheet.delete_rows(target_row_number, 1)
        wb.save(target_file_path)
        wb.close()

        # Re-parse affected file and update metadata
        if "files" in metadata and isinstance(metadata["files"], list):
            for f_meta in metadata["files"]:
                if f_meta.get("file_id") == file_id:
                    updated_info = self.parse_full_workbook(target_file_path, f_meta["filename"])
                    f_meta["sheet_count"] = updated_info["sheet_count"]
                    f_meta["records_count"] = updated_info["total_records"]
                    f_meta["customer_count"] = updated_info["customer_count"]
                    f_meta["sheets"] = [
                        {**s, "file_id": file_id, "file_name": f_meta["filename"]}
                        for s in updated_info["sheets"]
                    ]
                    break

            metadata["sheet_count"] = sum(f.get("sheet_count", len(f.get("sheets", []))) for f in metadata["files"])
            metadata["total_records"] = sum(f.get("records_count", 0) for f in metadata["files"])
            metadata["sheets"] = [s for f in metadata["files"] for s in f.get("sheets", [])]
        else:
            updated_info = self.parse_full_workbook(target_file_path, metadata.get("filename", "customers.xlsx"))
            metadata["sheet_count"] = updated_info["sheet_count"]
            metadata["total_records"] = updated_info["total_records"]
            metadata["customer_count"] = updated_info["customer_count"]
            metadata["sheets"] = updated_info["sheets"]

        metadata["last_updated"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with open(self.metadata_file_path, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, indent=2)

        self.invalidate_cache()
        self.load_all_data()

        return {
            "success": True,
            "deleted_record_id": record_id,
            "sheet_name": target_sheet_name,
            "deleted_row_number": target_row_number,
            "new_customer_count": len(self._cached_grouped_customers or {}),
            "metadata": metadata
        }

    def delete_customer(self, name):
        """
        Permanently deletes ALL records belonging to a customer across ALL sheets and files.
        """
        records, grouped, sheets_meta, metadata = self.load_all_data()
        norm_name = self.normalize_customer_name(name)
        if not norm_name or not grouped:
            raise KeyError(f"Customer '{name}' not found.")

        canonical_norm = self._cached_name_to_cluster.get(norm_name, norm_name) if self._cached_name_to_cluster else norm_name
        if canonical_norm not in grouped:
            for c_norm, c_grp in grouped.items():
                if norm_name in c_grp.get("normalized_aliases", []):
                    canonical_norm = c_norm
                    break

        if canonical_norm not in grouped:
            raise KeyError(f"Customer '{name}' not found.")

        target_records = grouped[canonical_norm]["records"]
        deleted_count = len(target_records)

        # Group rows by (file_id, sheet_name)
        file_sheet_rows = {}
        for rec in target_records:
            fid = rec.get("file_id", "file_1")
            sname = rec["sheet_name"]
            rnum = int(rec["row_number"])
            if fid not in file_sheet_rows:
                file_sheet_rows[fid] = {}
            if sname not in file_sheet_rows[fid]:
                file_sheet_rows[fid][sname] = []
            file_sheet_rows[fid][sname].append(rnum)

        for fid, sheets_dict in file_sheet_rows.items():
            target_path = None
            if "files" in metadata and isinstance(metadata["files"], list):
                for f_meta in metadata["files"]:
                    if f_meta.get("file_id") == fid:
                        target_path = self.customer_dir / f_meta.get("saved_file")
                        break
            if not target_path or not target_path.exists():
                target_path = self.excel_file_path

            if not target_path.exists():
                continue

            wb = openpyxl.load_workbook(target_path)
            for sname, row_nums in sheets_dict.items():
                if sname not in wb.sheetnames:
                    continue
                sheet = wb[sname]
                for r in sorted(row_nums, reverse=True):
                    sheet.delete_rows(r, 1)

            wb.save(target_path)
            wb.close()

            if "files" in metadata and isinstance(metadata["files"], list):
                for f_meta in metadata["files"]:
                    if f_meta.get("file_id") == fid:
                        updated_info = self.parse_full_workbook(target_path, f_meta["filename"])
                        f_meta["sheet_count"] = updated_info["sheet_count"]
                        f_meta["records_count"] = updated_info["total_records"]
                        f_meta["customer_count"] = updated_info["customer_count"]
                        f_meta["sheets"] = [
                            {**s, "file_id": fid, "file_name": f_meta["filename"]}
                            for s in updated_info["sheets"]
                        ]
                        break

        if "files" in metadata and isinstance(metadata["files"], list):
            metadata["sheet_count"] = sum(f.get("sheet_count", len(f.get("sheets", []))) for f in metadata["files"])
            metadata["total_records"] = sum(f.get("records_count", 0) for f in metadata["files"])
            metadata["sheets"] = [s for f in metadata["files"] for s in f.get("sheets", [])]
        else:
            updated_info = self.parse_full_workbook(self.excel_file_path, metadata.get("filename", "customers.xlsx"))
            metadata["sheet_count"] = updated_info["sheet_count"]
            metadata["total_records"] = updated_info["total_records"]
            metadata["customer_count"] = updated_info["customer_count"]
            metadata["sheets"] = updated_info["sheets"]

        metadata["last_updated"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with open(self.metadata_file_path, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, indent=2)

        self.invalidate_cache()
        self.load_all_data()

        return {
            "success": True,
            "deleted_customer": grouped[canonical_norm]["name"],
            "deleted_records_count": deleted_count,
            "new_customer_count": len(self._cached_grouped_customers or {}),
            "metadata": metadata
        }

    def delete_all_customers(self):
        """
        Permanently deletes ALL customer records across ALL sheets and files.
        Keeps sheets and header rows intact.
        """
        records, grouped, sheets_meta, metadata = self.load_all_data()

        if "files" in metadata and isinstance(metadata["files"], list):
            for f_meta in metadata["files"]:
                target_path = self.customer_dir / f_meta.get("saved_file")
                if not target_path.exists():
                    continue
                wb = openpyxl.load_workbook(target_path)
                for s_meta in f_meta.get("sheets", []):
                    sname = s_meta["name"]
                    if sname not in wb.sheetnames:
                        continue
                    sheet = wb[sname]
                    header_row_idx = s_meta.get("header_row_index", 1)
                    if header_row_idx < 0:
                        header_row_idx = 0
                    if sheet.max_row > header_row_idx:
                        sheet.delete_rows(header_row_idx + 1, sheet.max_row - header_row_idx)
                wb.save(target_path)
                wb.close()

                updated_info = self.parse_full_workbook(target_path, f_meta["filename"])
                f_meta["records_count"] = updated_info["total_records"]
                f_meta["customer_count"] = updated_info["customer_count"]
                f_meta["sheets"] = [
                    {**s, "file_id": f_meta["file_id"], "file_name": f_meta["filename"]}
                    for s in updated_info["sheets"]
                ]

            metadata["total_records"] = 0
            metadata["customer_count"] = 0
            metadata["sheets"] = [s for f in metadata["files"] for s in f.get("sheets", [])]

        elif self.excel_file_path.exists():
            wb = openpyxl.load_workbook(self.excel_file_path)
            for sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                meta = sheets_meta.get(sheet_name, {})
                header_row_idx = meta.get("header_row_index", 1)
                if header_row_idx < 0:
                    header_row_idx = 0
                if sheet.max_row > header_row_idx:
                    sheet.delete_rows(header_row_idx + 1, sheet.max_row - header_row_idx)
            wb.save(self.excel_file_path)
            wb.close()

            updated_info = self.parse_full_workbook(self.excel_file_path, metadata.get("filename", "customers.xlsx"))
            metadata["total_records"] = 0
            metadata["customer_count"] = 0
            metadata["sheets"] = updated_info["sheets"]

        metadata["last_updated"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with open(self.metadata_file_path, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, indent=2)

        self.invalidate_cache()
        self.load_all_data()

        return {
            "success": True,
            "message": "All customer records permanently deleted from Excel.",
            "new_customer_count": 0,
            "total_records": 0,
            "metadata": metadata
        }

    def clean_temp_file(self, temp_filename):
        if not temp_filename:
            return
        temp_path = self.temp_dir / temp_filename
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass

    def get_customer_profiles(self):
        """
        Parses HRDW WITH PH NO..xlsx to extract Address, Contact No., and Passport Photo.
        Maps each customer by intelligent matching to active customers.
        Caches results in memory.
        """
        if self._cached_customer_profiles is not None:
            return self._cached_customer_profiles

        profiles = {}
        hrdw_path = Path(r'f:\Train\HRDW WITH PH NO..xlsx')
        if not hrdw_path.exists():
            hrdw_path = Path(settings.BASE_DIR).parent / 'HRDW WITH PH NO..xlsx'

        if not hrdw_path.exists():
            self._cached_customer_profiles = {}
            return self._cached_customer_profiles

        try:
            wb = openpyxl.load_workbook(hrdw_path, data_only=True)
            if 'Sheet1' not in wb.sheetnames:
                wb.close()
                self._cached_customer_profiles = {}
                return self._cached_customer_profiles

            ws = wb['Sheet1']
            images_by_row = {}
            for img in getattr(ws, '_images', []):
                anchor = img.anchor
                r = anchor._from.row if hasattr(anchor, '_from') else anchor.from_.row
                images_by_row[r] = img._data()

            rows = list(ws.iter_rows(values_only=True))
            wb.close()

            self.photos_dir.mkdir(parents=True, exist_ok=True)
            records, grouped, sheets_meta, metadata = self.load_all_data()

            KNOWN_VARIATIONS = {
                "deepti r chudasama": "dipti rajeshbhai chudasama",
                "induben j chudasama": "indumatiben jayantilal chudasama",
                "sobhaben a jadav": "shobhanaben ashvinbhai jadav",
                "darshnaben p chudasama": "darshanaben paresh chavda",
                "pareshbhai k chavada": "pareshbhai karshanbhai chavda",
                "nisha a chudasama": "nishaben a chudasma",
                "pravina h chudasama": "pravina hiteshbhai chudasama",
                "rajubhai d chudasama": "rajubhai dayaljibhai chudasama",
                "dipaben r chudasama": "dipaben rajubhai chudasama"
            }

            for idx in range(1, len(rows)):
                r = rows[idx]
                if not r or not any(r):
                    continue
                cell_c = str(r[2]).strip() if len(r) > 2 and r[2] else ""
                mobile = str(r[3]).strip() if len(r) > 3 and r[3] else ""
                if mobile.endswith('.0'):
                    mobile = mobile[:-2]

                lines = [l.strip() for l in cell_c.split('\n') if l.strip()]
                if not lines:
                    continue
                raw_name = lines[0]
                address = ", ".join(lines[1:]) if len(lines) > 1 else ""

                h_norm = self.normalize_customer_name(raw_name)

                target_norm = KNOWN_VARIATIONS.get(h_norm)
                if not target_norm or target_norm not in grouped:
                    if h_norm in grouped:
                        target_norm = h_norm
                    else:
                        for c_norm, c_grp in grouped.items():
                            if h_norm in c_grp.get("normalized_aliases", []):
                                target_norm = c_norm
                                break
                        if not target_norm:
                            for c_norm, c_grp in grouped.items():
                                if self.are_names_similar(raw_name, c_grp["name"]):
                                    target_norm = c_norm
                                    break

                if target_norm:
                    img_bytes = images_by_row.get(idx)
                    photo_url = None
                    if img_bytes:
                        safe_slug = re.sub(r'[^\w\-]', '_', target_norm)
                        photo_file = self.photos_dir / f"{safe_slug}.jpg"
                        if not photo_file.exists():
                            with open(photo_file, 'wb') as pf:
                                pf.write(img_bytes)
                        photo_url = f"/media/customer_data/photos/{safe_slug}.jpg"

                    profiles[target_norm] = {
                        "address": address or "—",
                        "contact_no": mobile or "—",
                        "passport_photo": photo_url
                    }

            self._cached_customer_profiles = profiles
        except Exception as e:
            print(f"Error loading customer profiles: {e}")
            self._cached_customer_profiles = {}

        return self._cached_customer_profiles

    def get_customer_profile(self, canonical_norm):
        profiles = self.get_customer_profiles()
        return profiles.get(canonical_norm, {
            "address": "—",
            "contact_no": "—",
            "passport_photo": None
        })

    def invalidate_cache(self):
        self._cached_records = None
        self._cached_grouped_customers = None
        self._cached_name_to_cluster = None
        self._cached_sheets_meta = None
        self._cached_metadata = None
        self._cached_customer_profiles = None
