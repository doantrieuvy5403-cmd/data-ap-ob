from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
from models import db, ApartmentRecord
from datetime import datetime
import pandas as pd
import io
import os

app = Flask(__name__)
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, 'instance', 'database.db')
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{DB_PATH}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SECRET_KEY'] = 'data-ap-ob-secret-key-2026'

# Ensure instance directory exists
os.makedirs(os.path.join(BASE_DIR, 'instance'), exist_ok=True)

db.init_app(app)

# Column mapping for seed (by position)
SEED_COLUMNS = {
    0: 'stt', 1: 'team_assignment', 2: 'person_in_charge',
    3: 'status', 4: 'approach_time', 5: 'notes',
    6: 'city', 7: 'direction', 8: 'building_name', 9: 'district',
    10: 'num_blocks', 11: 'price_range', 12: 'infrastructure',
    13: 'occupancy', 14: 'classification', 15: 'previous_operator',
    16: 'total_screens', 17: 'screens_in_elevator',
    18: 'screens_outside_elevator', 19: 'p9000', 20: 'p6000', 21: 'prospect',
}
INT_FIELDS = ['stt', 'num_blocks', 'total_screens', 'screens_in_elevator',
              'screens_outside_elevator', 'p9000', 'p6000']


def _auto_seed():
    """Seed database from data.xlsx if empty."""
    if ApartmentRecord.query.count() > 0:
        return
    data_file = os.path.join(BASE_DIR, 'data.xlsx')
    if not os.path.exists(data_file):
        print("Warning: data.xlsx not found, skipping seed")
        return

    print("Auto-seeding database from data.xlsx...")
    try:
        xls = pd.ExcelFile(data_file)
        sheets = {
            'Databse AP_MN': ('MN', 11),
            'Databse AP_MB': ('MB', 12),
        }
        for sheet_name, (region, skip) in sheets.items():
            if sheet_name not in xls.sheet_names:
                continue
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None, skiprows=skip)
            for _, row in df.iterrows():
                building = row.iloc[8] if len(row) > 8 else None
                if pd.isna(building) or not building:
                    continue
                record = ApartmentRecord(region=region)
                for col_idx, field in SEED_COLUMNS.items():
                    if col_idx >= len(row):
                        continue
                    val = row.iloc[col_idx]
                    if pd.isna(val):
                        continue
                    if field in INT_FIELDS:
                        try:
                            val = int(float(str(val).replace("'", "").replace(",", "")))
                        except (ValueError, TypeError):
                            continue
                    elif field == 'approach_time':
                        val = val.strftime('%m/%Y') if hasattr(val, 'strftime') else str(val)
                    else:
                        val = str(val)
                    setattr(record, field, val)
                db.session.add(record)
            db.session.commit()
        total = ApartmentRecord.query.count()
        print(f"Seeded {total} records successfully!")
    except Exception as e:
        db.session.rollback()
        print(f"Seed error: {e}")


# Create tables and auto-seed
with app.app_context():
    db.create_all()
    _auto_seed()


@app.route('/')
def index():
    """Landing page with summary stats"""
    total_mn = ApartmentRecord.query.filter_by(region='MN').count()
    total_mb = ApartmentRecord.query.filter_by(region='MB').count()

    # Status breakdown
    status_stats = db.session.query(
        ApartmentRecord.status,
        db.func.count(ApartmentRecord.id)
    ).group_by(ApartmentRecord.status).all()

    return render_template('index.html',
                         total_mn=total_mn,
                         total_mb=total_mb,
                         status_stats=status_stats)


@app.route('/database/<region>')
def database(region):
    """Table view for AP_MN or AP_MB with search/filter"""
    region = region.upper()
    if region not in ['MN', 'MB']:
        flash('Invalid region', 'error')
        return redirect(url_for('index'))

    # Base query
    query = ApartmentRecord.query.filter_by(region=region)

    # Search
    search = request.args.get('search', '').strip()
    if search:
        query = query.filter(
            db.or_(
                ApartmentRecord.building_name.contains(search),
                ApartmentRecord.district.contains(search),
                ApartmentRecord.person_in_charge.contains(search),
                ApartmentRecord.notes.contains(search)
            )
        )

    # Filters
    status = request.args.get('status')
    if status:
        query = query.filter_by(status=status)

    person = request.args.get('person')
    if person:
        query = query.filter_by(person_in_charge=person)

    city = request.args.get('city')
    if city:
        query = query.filter_by(city=city)

    classification = request.args.get('classification')
    if classification:
        query = query.filter_by(classification=classification)

    # Pagination
    page = request.args.get('page', 1, type=int)
    per_page = 50
    records = query.paginate(page=page, per_page=per_page, error_out=False)

    # Get filter options
    statuses = db.session.query(ApartmentRecord.status).filter_by(region=region).distinct().all()
    persons = db.session.query(ApartmentRecord.person_in_charge).filter_by(region=region).distinct().all()
    cities = db.session.query(ApartmentRecord.city).filter_by(region=region).distinct().all()
    classifications = db.session.query(ApartmentRecord.classification).filter_by(region=region).distinct().all()

    return render_template('database.html',
                         region=region,
                         records=records,
                         statuses=[s[0] for s in statuses if s[0]],
                         persons=[p[0] for p in persons if p[0]],
                         cities=[c[0] for c in cities if c[0]],
                         classifications=[c[0] for c in classifications if c[0]],
                         search=search,
                         current_status=status,
                         current_person=person,
                         current_city=city,
                         current_classification=classification)


@app.route('/database/<region>/add', methods=['GET', 'POST'])
def add_record(region):
    """Add new record"""
    region = region.upper()
    if region not in ['MN', 'MB']:
        flash('Invalid region', 'error')
        return redirect(url_for('index'))

    if request.method == 'POST':
        record = ApartmentRecord(
            region=region,
            stt=request.form.get('stt', type=int),
            team_assignment=request.form.get('team_assignment'),
            person_in_charge=request.form.get('person_in_charge'),
            status=request.form.get('status'),
            approach_time=request.form.get('approach_time'),
            notes=request.form.get('notes'),
            city=request.form.get('city'),
            direction=request.form.get('direction'),
            building_name=request.form.get('building_name'),
            district=request.form.get('district'),
            num_blocks=request.form.get('num_blocks', type=int),
            price_range=request.form.get('price_range'),
            infrastructure=request.form.get('infrastructure'),
            occupancy=request.form.get('occupancy'),
            classification=request.form.get('classification'),
            previous_operator=request.form.get('previous_operator'),
            total_screens=request.form.get('total_screens', type=int),
            screens_in_elevator=request.form.get('screens_in_elevator', type=int),
            screens_outside_elevator=request.form.get('screens_outside_elevator', type=int),
            p9000=request.form.get('p9000', type=int),
            p6000=request.form.get('p6000', type=int),
            prospect=request.form.get('prospect')
        )
        db.session.add(record)
        db.session.commit()
        flash('Record added successfully', 'success')
        return redirect(url_for('database', region=region.lower()))

    return render_template('add_edit.html', region=region, record=None)


@app.route('/record/<int:id>/edit', methods=['GET', 'POST'])
def edit_record(id):
    """Edit existing record"""
    record = ApartmentRecord.query.get_or_404(id)

    if request.method == 'POST':
        record.stt = request.form.get('stt', type=int)
        record.team_assignment = request.form.get('team_assignment')
        record.person_in_charge = request.form.get('person_in_charge')
        record.status = request.form.get('status')
        record.approach_time = request.form.get('approach_time')
        record.notes = request.form.get('notes')
        record.city = request.form.get('city')
        record.direction = request.form.get('direction')
        record.building_name = request.form.get('building_name')
        record.district = request.form.get('district')
        record.num_blocks = request.form.get('num_blocks', type=int)
        record.price_range = request.form.get('price_range')
        record.infrastructure = request.form.get('infrastructure')
        record.occupancy = request.form.get('occupancy')
        record.classification = request.form.get('classification')
        record.previous_operator = request.form.get('previous_operator')
        record.total_screens = request.form.get('total_screens', type=int)
        record.screens_in_elevator = request.form.get('screens_in_elevator', type=int)
        record.screens_outside_elevator = request.form.get('screens_outside_elevator', type=int)
        record.p9000 = request.form.get('p9000', type=int)
        record.p6000 = request.form.get('p6000', type=int)
        record.prospect = request.form.get('prospect')

        db.session.commit()
        flash('Record updated successfully', 'success')
        return redirect(url_for('database', region=record.region.lower()))

    return render_template('add_edit.html', region=record.region, record=record)


@app.route('/record/<int:id>/delete', methods=['POST'])
def delete_record(id):
    """Delete record"""
    record = ApartmentRecord.query.get_or_404(id)
    region = record.region
    db.session.delete(record)
    db.session.commit()
    flash('Record deleted successfully', 'success')
    return redirect(url_for('database', region=region.lower()))


@app.route('/dashboard')
def dashboard():
    """Dashboard with charts"""
    return render_template('dashboard.html')


@app.route('/api/stats')
def api_stats():
    """JSON endpoint for dashboard charts"""
    # Status distribution
    status_stats = db.session.query(
        ApartmentRecord.status,
        db.func.count(ApartmentRecord.id)
    ).group_by(ApartmentRecord.status).all()

    # City distribution
    city_stats = db.session.query(
        ApartmentRecord.city,
        db.func.count(ApartmentRecord.id)
    ).filter(ApartmentRecord.city.isnot(None)).group_by(ApartmentRecord.city).all()

    # Person workload
    person_stats = db.session.query(
        ApartmentRecord.person_in_charge,
        db.func.count(ApartmentRecord.id)
    ).filter(ApartmentRecord.person_in_charge.isnot(None)).group_by(ApartmentRecord.person_in_charge).all()

    # Infrastructure distribution
    infra_stats = db.session.query(
        ApartmentRecord.infrastructure,
        db.func.count(ApartmentRecord.id)
    ).filter(ApartmentRecord.infrastructure.isnot(None)).group_by(ApartmentRecord.infrastructure).all()

    # Classification distribution
    class_stats = db.session.query(
        ApartmentRecord.classification,
        db.func.count(ApartmentRecord.id)
    ).filter(ApartmentRecord.classification.isnot(None)).group_by(ApartmentRecord.classification).all()

    # MN vs MB by status
    region_status = db.session.query(
        ApartmentRecord.region,
        ApartmentRecord.status,
        db.func.count(ApartmentRecord.id)
    ).group_by(ApartmentRecord.region, ApartmentRecord.status).all()

    return jsonify({
        'status': [{'label': s[0], 'count': s[1]} for s in status_stats if s[0]],
        'city': [{'label': c[0], 'count': c[1]} for c in city_stats][:10],  # Top 10
        'person': [{'label': p[0], 'count': p[1]} for p in person_stats],
        'infrastructure': [{'label': i[0], 'count': i[1]} for i in infra_stats if i[0]],
        'classification': [{'label': c[0], 'count': c[1]} for c in class_stats if c[0]],
        'region_status': [{'region': r[0], 'status': r[1], 'count': r[2]} for r in region_status if r[1]]
    })


@app.route('/export/<region>')
def export_data(region):
    """Export data as Excel"""
    region = region.upper()
    if region not in ['MN', 'MB']:
        flash('Invalid region', 'error')
        return redirect(url_for('index'))

    records = ApartmentRecord.query.filter_by(region=region).all()

    # Convert to DataFrame
    data = [r.to_dict() for r in records]
    df = pd.DataFrame(data)

    # Rename columns to Vietnamese
    column_map = {
        'stt': 'STT',
        'team_assignment': 'NS. Phân công',
        'person_in_charge': 'NS. Phụ trách',
        'status': 'Tiến độ',
        'approach_time': 'Thời gian tiếp cận',
        'notes': 'Ghi chú',
        'city': 'TP/Tỉnh',
        'direction': 'Hướng',
        'building_name': 'Tên chung cư',
        'district': 'Quận/Khu vực',
        'num_blocks': 'Số Block',
        'price_range': 'Giá bán',
        'infrastructure': 'CSVC',
        'occupancy': 'Tỷ lệ lấp đầy',
        'classification': 'Phân loại',
        'previous_operator': 'Đơn vị cũ',
        'total_screens': 'Tổng SL màn hình',
        'screens_in_elevator': 'Số màn trong thang',
        'screens_outside_elevator': 'Số màn ngoài thang',
        'p9000': 'P9000',
        'p6000': 'P6000',
        'prospect': 'Prospect'
    }
    df = df.rename(columns=column_map)
    df = df.drop(columns=['id', 'region', 'created_at', 'updated_at'], errors='ignore')

    # Save to Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=f'AP_{region}')
    output.seek(0)

    filename = f'Database_AP_{region}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    return send_file(output,
                    as_attachment=True,
                    download_name=filename,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/import/<region>', methods=['POST'])
def import_data(region):
    """Import data from uploaded Excel/CSV"""
    region = region.upper()
    if region not in ['MN', 'MB']:
        flash('Invalid region', 'error')
        return redirect(url_for('index'))

    if 'file' not in request.files:
        flash('No file uploaded', 'error')
        return redirect(url_for('import_export'))

    file = request.files['file']
    if file.filename == '':
        flash('No file selected', 'error')
        return redirect(url_for('import_export'))

    try:
        # Read file
        if file.filename.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)

        # Map Vietnamese columns to English
        column_map = {
            'STT': 'stt',
            'NS. Phân công': 'team_assignment',
            'NS. Phụ trách': 'person_in_charge',
            'Tiến độ': 'status',
            'Thời gian tiếp cận': 'approach_time',
            'Ghi chú': 'notes',
            'TP/Tỉnh': 'city',
            'Hướng': 'direction',
            'Tên chung cư': 'building_name',
            'Quận/Khu vực': 'district',
            'Số Block': 'num_blocks',
            'Giá bán': 'price_range',
            'CSVC': 'infrastructure',
            'Tỷ lệ lấp đầy': 'occupancy',
            'Phân loại': 'classification',
            'Đơn vị cũ': 'previous_operator',
            'Tổng SL màn hình': 'total_screens',
            'Số màn trong thang': 'screens_in_elevator',
            'Số màn ngoài thang': 'screens_outside_elevator',
            'P9000': 'p9000',
            'P6000': 'p6000',
            'Prospect': 'prospect'
        }
        df = df.rename(columns=column_map)

        # Add records
        count = 0
        for _, row in df.iterrows():
            record = ApartmentRecord(region=region)
            for col in column_map.values():
                if col in row:
                    value = row[col]
                    if pd.notna(value):
                        setattr(record, col, value if not isinstance(value, (int, float)) else int(value))
            db.session.add(record)
            count += 1

        db.session.commit()
        flash(f'Successfully imported {count} records', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Error importing file: {str(e)}', 'error')

    return redirect(url_for('database', region=region.lower()))


@app.route('/import-export')
def import_export():
    """Import/Export page"""
    return render_template('import_export.html')


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)
