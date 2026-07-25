import os
import sys
import pandas as pd
from app import app, db, ApartmentRecord, WeeklyGrowth

def import_all(file_path):
    print(f"Reading file: {file_path}")
    
    with app.app_context():
        print("Clearing existing data...")
        db.session.query(ApartmentRecord).delete()
        db.session.commit()
        
        from app import APOB_COLUMNS, APOB_IMPORT_ALIASES
        field_for = {label: field for label, field in APOB_COLUMNS if field not in ('pt1', 'pt2', 'pt3')}
        field_for.update(APOB_IMPORT_ALIASES)
        
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        
        for sheet in sheet_names:
            if sheet == 'Install':
                continue
                
            print(f"Processing sheet: {sheet}")
            parts = sheet.split('_')
            if len(parts) != 2:
                print(f"Skipping unknown sheet format: {sheet}")
                continue
            cat, reg = parts
            
            df = pd.read_excel(file_path, sheet_name=sheet)
            df = df.where(pd.notnull(df), None)
            
            for _, row in df.iterrows():
                row_dict = row.to_dict()
                person = []
                for pt in ('PT1', 'PT2', 'PT3'):
                    v = str(row_dict.get(pt, '')).strip()
                    if v and v.lower() != 'none':
                        person.append(v)
                person_in_charge = ', '.join(person).upper() if person else None
                
                rec = ApartmentRecord(
                    category=cat,
                    region=reg,
                    person_in_charge=person_in_charge
                )
                
                for label, val in row_dict.items():
                    field = field_for.get(label)
                    if field:
                        if val is None or str(val).strip().lower() == 'nan':
                            setattr(rec, field, None)
                        elif field in ('stt', 'num_blocks', 'total_screens', 'screens_in_elevator', 'screens_outside_elevator', 'p9000', 'p6000'):
                            try:
                                setattr(rec, field, int(float(val)))
                            except (ValueError, TypeError):
                                setattr(rec, field, None)
                        else:
                            setattr(rec, field, str(val).strip().upper())
                
                db.session.add(rec)
        
        print("Committing to database...")
        db.session.commit()
        
        if 'Install' in sheet_names:
            print("Processing Install sheet...")
            df_install = pd.read_excel(file_path, sheet_name='Install')
            df_install = df_install.where(pd.notnull(df_install), None)
            
            for _, row in df_install.iterrows():
                building_name = str(row.get('Dự án (toà nhà)', '')).strip().upper()
                cat = str(row.get('Loại Hình', '')).strip().upper()
                reg = str(row.get('Khu vực', '')).strip().upper()
                
                if not building_name or building_name == 'NONE':
                    continue
                    
                rec = ApartmentRecord.query.filter_by(
                    building_name=building_name,
                    category=cat,
                    region=reg
                ).first()
                
                if rec:
                    td = row.get('Đã triển khai (MH)')
                    try:
                        rec.total_deployed = int(float(td)) if td is not None else 0
                    except (ValueError, TypeError):
                        rec.total_deployed = 0
                        
                    es = row.get('Tình trạng')
                    rec.electricity_status = str(es).strip() if es is not None else None
                    
                    no = row.get('Ghi chú')
                    rec.install_note = str(no).strip() if no is not None else None
            
            db.session.commit()
            
        print("Import completed successfully!")

if __name__ == '__main__':
    if 'DATABASE_URL' not in os.environ:
        print("Error: DATABASE_URL environment variable is not set.")
        sys.exit(1)
        
    file_path = "/Users/qweasdzxcbm/Downloads/Database_All_20260721_133250.xlsx"
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        
    if not os.path.exists(file_path):
        print(f"Error: File not found: {file_path}")
        sys.exit(1)
        
    import_all(file_path)
