import os, django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'config.settings')
django.setup()
from customers.excel_service import ExcelDataManager

em = ExcelDataManager.get_instance()
em.invalidate_cache()
recs, grp, sm, meta = em.load_all_data()

print(f"Total current records: {len(recs)}")

suspicious = []
for r in recs.values():
    row_vals = [v for k, v in r['data'].items() if str(v).strip()]
    has_head = any('(HEAD)' in str(v).upper() or '(HEAD) -' in str(v).upper() for v in r['data'].values())
    
    # Check passenger data substance:
    # Does it have ticket, seat, age, train number, or is it just a name + route banner?
    age = None
    sex = None
    seat = None
    tkt = None
    train = None
    for k, v in r['data'].items():
        k_up = k.upper()
        if 'AGE' in k_up and v: age = v
        if ('SEX' in k_up or 'GENDER' in k_up) and v: sex = v
        if ('SEAT' in k_up or 'STATUS' in k_up) and v: seat = v
        if ('TICKET' in k_up or 'TKT' in k_up) and v: tkt = v
        if 'TRAIN' in k_up and v: train = v

    # Check if row has virtually no passenger booking data
    if has_head or (not age and not sex and not seat and not tkt):
        suspicious.append((r, has_head, age, sex, seat, tkt, train))

print(f"\nSuspicious / Blank records count: {len(suspicious)}")
for r, has_head, age, sex, seat, tkt, train in suspicious[:30]:
    print(f"Row {r['row_number']} in {r['sheet_name']} ({r['file_name']}) -> Name: '{r['name']}', HasHead: {has_head}, Age: {age}, Sex: {sex}, Seat: {seat}, Tkt: {tkt}, Train: {train}")
    print(f"   Data: {r['data']}")
