import os, django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')
django.setup()
from customers.excel_service import ExcelDataManager

em = ExcelDataManager.get_instance()
em.invalidate_cache()
recs, grp, sm, meta = em.load_all_data()

def is_blank_or_section_header_row(row_data, name_col):
    raw_name = str(row_data.get(name_col, "")).strip()
    if not raw_name:
        return True, "No name"
    
    # 1. Any cell contains (HEAD)
    for k, v in row_data.items():
        v_str = str(v).strip().upper()
        if '(HEAD)' in v_str or v_str.startswith('(HEAD)') or 'LIST OF PASSENGERS' in v_str or v_str.startswith('TOTAL'):
            return True, f"Banner keyword in {k}: '{v}'"

    # 2. Extract potential passenger booking indicators
    has_age = False
    has_sex = False
    has_seat = False
    has_tkt = False
    has_numeric_train = False
    has_doj = False
    has_fare = False
    has_sr_no = False

    for k, v in row_data.items():
        k_up = str(k).strip().upper()
        v_str = str(v).strip()
        if not v_str:
            continue
        v_up = v_str.upper()

        if ('NO.' in k_up or 'SR' in k_up) and 'TRAIN' not in k_up and 'TICKET' not in k_up and 'TCKT' not in k_up:
            if v_str.isdigit(): has_sr_no = True
        if 'AGE' in k_up and any(c.isdigit() for c in v_str):
            has_age = True
        if 'SEX' in k_up or 'GENDER' in k_up:
            if v_up in ('M', 'F', 'MALE', 'FEMALE'): has_sex = True
        if 'SEAT' in k_up or 'STATUS' in k_up or 'BERTH' in k_up:
            if v_up not in ('—', '-', '--', 'N/A', 'SEAT', 'STATUS'): has_seat = True
        if 'TICKET' in k_up or 'TKT' in k_up or 'TCKT' in k_up:
            if any(c.isdigit() for c in v_str): has_tkt = True
        if 'TRAIN' in k_up:
            if any(c.isdigit() for c in v_str): has_numeric_train = True
        if 'DOJ' in k_up or 'DATE' in k_up:
            if v_up not in ('—', '-', '--', 'N/A', 'DOJ'): has_doj = True
        if 'RS' in k_up or 'AMOUNT' in k_up or 'FARE' in k_up or 'PRICE' in k_up:
            if any(c.isdigit() for c in v_str): has_fare = True

    # A valid passenger record must have AT LEAST ONE of:
    # seat, ticket, numeric train, valid age, valid sex, fare, or valid sr_no
    has_passenger_data = has_seat or has_tkt or has_numeric_train or has_age or has_sex or has_fare or has_sr_no
    if not has_passenger_data:
        return True, "No passenger booking data (all booking columns blank)"

    return False, "Valid"

blank_count = 0
valid_count = 0
for r_id, r in recs.items():
    is_blank, reason = is_blank_or_section_header_row(r['data'], 'NAME')
    if is_blank:
        blank_count += 1
    else:
        valid_count += 1

print(f"Total current records: {len(recs)}")
print(f"Blank / Section Header rows detected: {blank_count}")
print(f"True valid passenger records remaining: {valid_count}")
