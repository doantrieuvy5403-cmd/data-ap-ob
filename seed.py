"""
seed.py — Import data from Excel into SQLite database.
Run: python3 seed.py
"""
import pandas as pd
from app import app
from models import db, ApartmentRecord
import os

EXCEL_PATH = '/Users/qweasdzxcbm/Downloads/Database_Inspired Space.xlsx'

# Column names mapped by position (0-indexed)
COLUMN_NAMES = {
    0: 'stt',
    1: 'team_assignment',
    2: 'person_in_charge',
    3: 'status',
    4: 'approach_time',
    5: 'notes',
    6: 'city',
    7: 'direction',
    8: 'building_name',
    9: 'district',
    10: 'num_blocks',
    11: 'price_range',
    12: 'infrastructure',
    13: 'occupancy',
    14: 'classification',
    15: 'previous_operator',
    16: 'total_screens',
    17: 'screens_in_elevator',
    18: 'screens_outside_elevator',
    19: 'p9000',
    20: 'p6000',
    21: 'prospect',
}

INT_FIELDS = ['stt', 'num_blocks', 'total_screens', 'screens_in_elevator',
              'screens_outside_elevator', 'p9000', 'p6000']

# Data start row (0-indexed)
DATA_START_AP_MN = 11
DATA_START_AP_MB = 12


def clean_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, float) and value == int(value):
        return int(value)
    return value


def import_sheet(xls, sheet_name, region, data_start_row):
    """Import a sheet into the database."""
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None, skiprows=data_start_row)
    print(f"  Read {len(df)} rows, {len(df.columns)} columns")

    count = 0
    skipped = 0

    for _, row in df.iterrows():
        # Get building name (column 8)
        building = clean_value(row.iloc[8]) if len(row) > 8 else None
        if not building:
            skipped += 1
            continue

        record = ApartmentRecord(region=region)

        for col_idx, field_name in COLUMN_NAMES.items():
            if col_idx < len(row):
                value = clean_value(row.iloc[col_idx])
                if value is not None:
                    if field_name in INT_FIELDS:
                        try:
                            value = int(float(str(value).replace("'", "").replace(",", "")))
                        except (ValueError, TypeError):
                            value = None
                    elif field_name == 'approach_time':
                        # Handle datetime
                        if hasattr(value, 'strftime'):
                            value = value.strftime('%m/%Y')
                        else:
                            value = str(value)
                    else:
                        value = str(value)
                    setattr(record, field_name, value)

        db.session.add(record)
        count += 1

    db.session.commit()
    return count, skipped


def main():
    if not os.path.exists(EXCEL_PATH):
        print(f"ERROR: Excel file not found: {EXCEL_PATH}")
        return

    with app.app_context():
        # Drop and recreate tables
        db.drop_all()
        db.create_all()
        print("Database tables created.\n")

        xls = pd.ExcelFile(EXCEL_PATH)
        print(f"Sheets found: {xls.sheet_names}\n")

        # Import AP_MN
        print("Importing 'Databse AP_MN' (Miền Nam)...")
        count_mn, skip_mn = import_sheet(xls, 'Databse AP_MN', 'MN', DATA_START_AP_MN)
        print(f"  -> Imported: {count_mn} records (skipped {skip_mn} empty rows)\n")

        # Import AP_MB
        print("Importing 'Databse AP_MB' (Miền Bắc)...")
        count_mb, skip_mb = import_sheet(xls, 'Databse AP_MB', 'MB', DATA_START_AP_MB)
        print(f"  -> Imported: {count_mb} records (skipped {skip_mb} empty rows)\n")

        # Summary
        total_mn = ApartmentRecord.query.filter_by(region='MN').count()
        total_mb = ApartmentRecord.query.filter_by(region='MB').count()
        print(f"=== Done! ===")
        print(f"  AP_MN: {total_mn} records")
        print(f"  AP_MB: {total_mb} records")
        print(f"  Total: {total_mn + total_mb} records")


if __name__ == '__main__':
    main()
