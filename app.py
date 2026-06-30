from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file, session
from models import db, ApartmentRecord, WeeklyGrowth, AppMeta
from datetime import datetime
import pandas as pd
import io
import os
import re
import hashlib
from functools import wraps
from werkzeug.security import check_password_hash, generate_password_hash

app = Flask(__name__)
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, 'instance', 'database.db')

# Use Postgres (Render) when DATABASE_URL is set; fall back to local SQLite.
database_url = os.environ.get('DATABASE_URL')
if database_url:
    # SQLAlchemy needs the postgresql:// scheme (Render gives postgres://)
    if database_url.startswith('postgres://'):
        database_url = database_url.replace('postgres://', 'postgresql://', 1)
    app.config['SQLALCHEMY_DATABASE_URI'] = database_url
else:
    app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{DB_PATH}'
app.config['SQLALCHEMY_ENGINE_OPTIONS'] = {'pool_pre_ping': True}
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['SECRET_KEY'] = 'data-ap-ob-secret-key-2026'
app.config['ADMIN_USERNAME'] = os.environ.get('ADMIN_USERNAME', 'admin')
app.config['ADMIN_PASSWORD_HASH'] = generate_password_hash(os.environ.get('ADMIN_PASSWORD', 'password'))

# Ensure instance directory exists
os.makedirs(os.path.join(BASE_DIR, 'instance'), exist_ok=True)

db.init_app(app)

@app.context_processor
def inject_user():
    # Public access: always treat as logged in so the UI shows the sidebar.
    return {
        'logged_in': True,
        'current_user': session.get('username') or 'Inspired Space'
    }

def login_required(view):
    # Login disabled: pass through without any authentication check.
    @wraps(view)
    def wrapped_view(*args, **kwargs):
        return view(*args, **kwargs)
    return wrapped_view

@app.route('/login', methods=['GET', 'POST'])
def login():
    if session.get('logged_in'):
        return redirect(url_for('index'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')

        if username == app.config['ADMIN_USERNAME'] and check_password_hash(app.config['ADMIN_PASSWORD_HASH'], password):
            session['logged_in'] = True
            session['username'] = username
            flash('Đăng nhập thành công', 'success')
            next_url = request.args.get('next')
            return redirect(next_url or url_for('index'))

        flash('Tên đăng nhập hoặc mật khẩu không đúng', 'error')

    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    flash('Đăng xuất thành công', 'success')
    return redirect(url_for('login'))

# Column mapping for seed (by position) — Apartment (AP) sheets
SEED_COLUMNS_AP = {
    0: 'stt', 1: 'team_assignment', 2: 'person_in_charge',
    3: 'status', 4: 'approach_time', 5: 'notes',
    6: 'city', 7: 'direction', 8: 'building_name', 9: 'district',
    10: 'num_blocks', 11: 'price_range', 12: 'infrastructure',
    13: 'occupancy', 14: 'classification', 15: 'previous_operator',
    16: 'total_screens', 17: 'screens_in_elevator',
    18: 'screens_outside_elevator', 19: 'p9000', 20: 'p6000', 21: 'must_have',
}
# Office Building (OB) sheets: extra "address" col, blocks shifted, LED/DP-LCD screens
SEED_COLUMNS_OB = {
    0: 'stt', 1: 'team_assignment', 2: 'person_in_charge',
    3: 'status', 4: 'approach_time', 5: 'notes',
    6: 'city', 7: 'direction', 8: 'building_name', 9: 'district',
    10: 'address', 11: 'num_blocks', 12: 'price_range', 13: 'infrastructure',
    14: 'occupancy', 15: 'classification', 16: 'previous_operator',
    27: 'must_have',
}
# For OB, total_screens = TỔNG SL LED (17) + TỔNG SL DP/LCD (22)
OB_SCREEN_COLS = [17, 22]

INT_FIELDS = ['stt', 'num_blocks', 'total_screens', 'screens_in_elevator',
              'screens_outside_elevator', 'p9000', 'p6000']

# Bump when the seed logic changes (forces a one-time reseed even if data.xlsx is unchanged)
SEED_VERSION = '4-mt-region-2026-07'

# (sheet_name, category, region, skiprows, column_map, screen_sum_cols)
SEED_SHEETS = [
    ('Databse AP_MN', 'AP', 'MN', 11, SEED_COLUMNS_AP, None),
    ('Databse AP_MT', 'AP', 'MT', 12, SEED_COLUMNS_AP, None),
    ('Databse AP_MB', 'AP', 'MB', 12, SEED_COLUMNS_AP, None),
    ('Databse OB_MN (A,B)', 'OB', 'MN', 13, SEED_COLUMNS_OB, OB_SCREEN_COLS),
    ('Databse OB_MT (A,B)', 'OB', 'MT', 11, SEED_COLUMNS_OB, OB_SCREEN_COLS),
    ('Databse OB_MB (A,B,C)', 'OB', 'MB', 12, SEED_COLUMNS_OB, OB_SCREEN_COLS),
]



def _derive_region(city):
    """Best-effort MB/MT/MN from a city name (install data has no region column)."""
    c = (city or '').lower()
    north = ['hà nội', 'ha noi', 'hanoi', 'hải phòng', 'hai phong', 'bắc ninh',
             'quảng ninh', 'hà nam', 'nam định', 'thái', 'vĩnh phúc', 'bắc', 'hưng yên']
    central = ['đà nẵng', 'da nang', 'huế', 'hue', 'quảng nam', 'quảng ngãi',
               'quy nhơn', 'bình định', 'nha trang', 'khánh hòa', 'phú yên',
               'quảng bình', 'quảng trị', 'hà tĩnh', 'nghệ an', 'thanh hóa',
               'đà lạt', 'lâm đồng', 'pleiku', 'gia lai', 'kon tum', 'buôn ma thuột', 'đắk lắk']
    if any(k in c for k in central):
        return 'MT'
    return 'MB' if any(k in c for k in north) else 'MN'


# Dashboard: persons in charge — (display name, matching token in data, team)
DASHBOARD_PERSONS = [
    ('Nhiên', 'NHIÊN', 'Team Nhiên'),
    ('Tân', 'TÂN', 'Team Tân'),
    ('Quỳnh Hà', 'QUỲNH HÀ', 'Team Tân'),
    ('Phú', 'PHÚ', 'Team Tân'),
    ('Phương', 'PHƯƠNG', 'Team Vy'),
    ('Dũng', 'DŨNG', 'Team Dũng'),
    ('Mai', 'MAI', 'Team Mai'),
    ('Khánh', 'KHÁNH', 'Team Mai'),
    ('Quyền', 'QUYỀN', 'Team Thuỳ'),
    ('Ngân An', 'AN', 'Team Thuỳ'),
    ('Thuỳ', 'THUỲ', 'Team Thuỳ'),
]
# Per-person screen targets (Deal + Done), split equally among all members
PERSON_TARGET_AP = 3000  # total AP screen target across the team
PERSON_TARGET_OB = 500   # total OB screen target across the team
DEAL_DONE_STAGES = ['Deal', 'Done']
# Sales funnel stages in order + completion weight for progress %
FUNNEL_STAGES = ['Research', 'Plan B', 'Plan A', 'Deal', 'Done']
FUNNEL_WEIGHT = {'Research': 0.2, 'Plan B': 0.4, 'Plan A': 0.6, 'Deal': 0.8, 'Done': 1.0}

# Screen targets per category (the chart name encodes the total target).
# "done" = sum of total_screens across these stages (Research EXCLUDED).
# The "MN" bucket combines Miền Nam + Miền Trung into one figure/chart.
SCREEN_STAGES = ['Plan B', 'Plan A', 'Deal', 'Done']
SCREEN_REGION_GROUPS = {'MB': ['MB'], 'MN': ['MN', 'MT']}
SCREEN_TARGETS = {
    'AP': {'name': 'P6000', 'MB': 2500, 'MN': 3500},
    'OB': {'name': 'P1000', 'MB': 450, 'MN': 550},
}


def _screen_progress():
    """Per-category screen progress vs target (for the P6000/P1000 donuts).

    MB = Miền Bắc; MN = Miền Nam + Miền Trung (combined).
    """
    result = {}
    for cat, cfg in SCREEN_TARGETS.items():
        sum_done = 0
        sum_target = 0
        regions = {}
        for reg in ('MB', 'MN'):
            done = db.session.query(
                db.func.coalesce(db.func.sum(ApartmentRecord.total_screens), 0)
            ).filter(
                ApartmentRecord.category == cat,
                ApartmentRecord.region.in_(SCREEN_REGION_GROUPS[reg]),
                ApartmentRecord.status.in_(SCREEN_STAGES),
            ).scalar() or 0
            done = int(done)
            target = cfg[reg]
            regions[reg] = {'done': done, 'target': target}
            sum_done += done
            sum_target += target
        result[cat] = {
            'name': cfg['name'],
            'MB': regions['MB'],
            'MN': regions['MN'],
            'SUM': {'done': sum_done, 'target': sum_target},
        }
    return result


def _compute_person_progress(category=None):
    """Per-person project counts + screen achievements vs AP/OB targets.

    Returns, per person: funnel stage counts, project-progress % (weighted
    funnel), and screen-progress % (Deal+Done screens vs per-person target,
    AP and OB weighted equally). Always computed across all categories so the
    AP/OB targets line up regardless of the dashboard category filter.
    """
    persons = DASHBOARD_PERSONS
    n = len(persons) or 1
    target_ap = PERSON_TARGET_AP / n   # per-person AP screen target
    target_ob = PERSON_TARGET_OB / n   # per-person OB screen target

    counts = {disp: {stage: 0 for stage in FUNNEL_STAGES} for disp, _, _ in persons}
    ap_screens = {disp: 0 for disp, _, _ in persons}
    ob_screens = {disp: 0 for disp, _, _ in persons}
    token_to_disp = {token: disp for disp, token, _ in persons}

    # Project counts across funnel stages
    for person, status in db.session.query(
        ApartmentRecord.person_in_charge, ApartmentRecord.status
    ).filter(
        ApartmentRecord.person_in_charge.isnot(None),
        ApartmentRecord.status.in_(FUNNEL_STAGES)
    ).all():
        for tok in str(person).split(','):
            disp = token_to_disp.get(tok.strip().upper())
            if disp:
                counts[disp][status] += 1

    # Screen achievements (Deal + Done) split by category
    for person, cat, screens in db.session.query(
        ApartmentRecord.person_in_charge, ApartmentRecord.category, ApartmentRecord.total_screens
    ).filter(
        ApartmentRecord.person_in_charge.isnot(None),
        ApartmentRecord.status.in_(DEAL_DONE_STAGES),
        ApartmentRecord.total_screens.isnot(None)
    ).all():
        if not screens:
            continue
        for tok in str(person).split(','):
            disp = token_to_disp.get(tok.strip().upper())
            if disp:
                if cat == 'AP':
                    ap_screens[disp] += int(screens)
                elif cat == 'OB':
                    ob_screens[disp] += int(screens)

    result = []
    for disp, _, team in persons:
        stage_counts = counts[disp]
        total = sum(stage_counts.values())
        weighted = sum(stage_counts[s] * FUNNEL_WEIGHT[s] for s in FUNNEL_STAGES)
        progress_project = round(weighted / total * 100, 1) if total else 0.0
        ap = ap_screens[disp]
        ob = ob_screens[disp]
        progress_ap = round(ap / target_ap * 100, 1) if target_ap else 0.0
        progress_ob = round(ob / target_ob * 100, 1) if target_ob else 0.0
        result.append({
            'person': disp,
            'team': team,
            'stages': stage_counts,
            'total': total,
            'progress': progress_project,
            'progress_ap': progress_ap,
            'progress_ob': progress_ob,
            'ap_screens': ap,
            'ob_screens': ob,
            'target_ap': round(target_ap, 1),
            'target_ob': round(target_ob, 1),
        })
    return result


def _auto_seed():
    """(Re)seed apartment data from data.xlsx when the file changes.

    Only ApartmentRecord is touched here — WeeklyGrowth history is preserved,
    so weekly growth survives across deploys when using a persistent DB.
    """
    data_file = os.path.join(BASE_DIR, 'data.xlsx')
    if not os.path.exists(data_file):
        print("Warning: data.xlsx not found, skipping seed")
        return

    try:
        with open(data_file, 'rb') as f:
            current_hash = f'{SEED_VERSION}:{hashlib.md5(f.read()).hexdigest()}'
        meta = db.session.get(AppMeta, 'data_hash')
        has_data = ApartmentRecord.query.count() > 0
        if has_data and meta and meta.value == current_hash:
            # data.xlsx & seed logic unchanged — keep all records (incl. manual
            # edits/adds) and weekly history. Only a NEW data.xlsx triggers a
            # fresh import below.
            return

        print("Seeding apartment data from data.xlsx (new/changed file)...")
        # Refresh apartment records only (weekly_growth untouched)
        ApartmentRecord.query.delete()
        db.session.commit()
        xls = pd.ExcelFile(data_file)
        for sheet_name, category, region, skip, colmap, screen_cols in SEED_SHEETS:
            if sheet_name not in xls.sheet_names:
                continue
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None, skiprows=skip)
            for _, row in df.iterrows():
                building = row.iloc[8] if len(row) > 8 else None
                if pd.isna(building) or not building:
                    continue
                record = ApartmentRecord(category=category, region=region)
                for col_idx, field in colmap.items():
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
                    elif field == 'must_have':
                        # Column holds "Must have " (trailing space) or blank
                        val = 'Must have' if 'must have' in str(val).strip().lower() else None
                        if val is None:
                            continue
                    else:
                        val = str(val)
                    setattr(record, field, val)
                if record.status == 'Deal':
                    record.total_deployed = 0
                # OB: total_screens = sum of LED + DP/LCD totals
                if screen_cols:
                    screens = 0
                    for sc in screen_cols:
                        if sc < len(row) and pd.notna(row.iloc[sc]):
                            try:
                                screens += int(float(str(row.iloc[sc]).replace("'", "").replace(",", "")))
                            except (ValueError, TypeError):
                                pass
                    record.total_screens = screens
                db.session.add(record)
            db.session.commit()
        # Record the seeded file hash so we don't reseed unchanged data
        if meta is None:
            meta = AppMeta(key='data_hash')
            db.session.add(meta)
        meta.value = current_hash
        db.session.commit()
        total = ApartmentRecord.query.count()
        print(f"Seeded {total} records successfully!")
    except Exception as e:
        db.session.rollback()
        print(f"Seed error: {e}")



def _ensure_schema():
    """Add columns introduced after initial deploy (safe for SQLite & Postgres)."""
    from sqlalchemy import inspect, text
    insp = inspect(db.engine)
    tables = insp.get_table_names()

    def add_col(table, col, ddl, backfill=None):
        if table not in tables:
            return
        existing = {c['name'] for c in insp.get_columns(table)}
        if col not in existing:
            db.session.execute(text(f"ALTER TABLE {table} ADD COLUMN {ddl}"))
            if backfill:
                db.session.execute(text(backfill))
            db.session.commit()

    add_col('apartment_record', 'category', "category VARCHAR(10)",
            "UPDATE apartment_record SET category='AP' WHERE category IS NULL")
    add_col('apartment_record', 'address', "address VARCHAR(255)")
    add_col('apartment_record', 'total_deployed', "total_deployed INTEGER DEFAULT 0")
    add_col('weekly_growth', 'category', "category VARCHAR(10)",
            "UPDATE weekly_growth SET category='AP' WHERE category IS NULL")
    for col in ('plan_b_blocks', 'plan_a_blocks', 'deal_blocks', 'done_blocks'):
        add_col('weekly_growth', col, f"{col} INTEGER DEFAULT 0")

    # Drop legacy UNIQUE(year, week) on weekly_growth — now keyed by (category, year, week)
    if 'weekly_growth' in tables and db.engine.name == 'postgresql':
        try:
            db.session.execute(text(
                "ALTER TABLE weekly_growth DROP CONSTRAINT IF EXISTS uq_weekly_year_week"))
            db.session.commit()
        except Exception:
            db.session.rollback()


# Create tables and auto-seed (never let startup crash the web process)
with app.app_context():
    try:
        db.create_all()
    except Exception as e:
        db.session.rollback()
        print(f"create_all error: {e}")
    try:
        _ensure_schema()
    except Exception as e:
        db.session.rollback()
        print(f"ensure_schema error: {e}")
    try:
        _auto_seed()
    except Exception as e:
        db.session.rollback()
        print(f"auto_seed error: {e}")



@app.route('/')
@login_required
def index():
    """Landing page with summary stats"""
    def count(category, region):
        return ApartmentRecord.query.filter_by(category=category, region=region).count()

    ap_mn = count('AP', 'MN')
    ap_mt = count('AP', 'MT')
    ap_mb = count('AP', 'MB')
    ob_mn = count('OB', 'MN')
    ob_mt = count('OB', 'MT')
    ob_mb = count('OB', 'MB')

    # Status breakdown (fixed display order: 2 rows of 5)
    STATUS_ORDER = ['Research', 'Plan B', 'Plan A', 'Deal', 'Done',
                    'Code', 'Fail', 'Pending', 'Lost', 'Reject']
    raw = dict(db.session.query(
        ApartmentRecord.status,
        db.func.count(ApartmentRecord.id)
    ).group_by(ApartmentRecord.status).all())
    status_stats = [(s, raw.get(s, 0)) for s in STATUS_ORDER]
    # Append any other statuses present in the data but not in the fixed list
    status_stats += [(s, c) for s, c in raw.items() if s and s not in STATUS_ORDER]

    return render_template('index.html',
                         ap_mn=ap_mn,
                         ap_mt=ap_mt,
                         ap_mb=ap_mb,
                         ob_mn=ob_mn,
                         ob_mt=ob_mt,
                         ob_mb=ob_mb,
                         status_stats=status_stats)


def _valid_cat_region(category, region):
    return category in ('AP', 'OB') and region in ('MN', 'MT', 'MB')


def _join_persons(form):
    """Combine PT1/PT2/PT3 inputs into a single comma-separated string."""
    parts = [form.get(f'pt{i}', '').strip() for i in (1, 2, 3)]
    joined = ', '.join(p for p in parts if p)
    return joined or None


@app.route('/database/<category>/<region>')
@login_required
def database(category, region):
    """Table view for AP/OB × MN/MB with search/filter"""
    category = category.upper()
    region = region.upper()
    if not _valid_cat_region(category, region):
        flash('Invalid category/region', 'error')
        return redirect(url_for('index'))

    # Base query
    query = ApartmentRecord.query.filter_by(category=category, region=region)

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

    district = request.args.get('district')
    if district:
        query = query.filter_by(district=district)

    must_have = request.args.get('must_have')
    if must_have == '1':
        query = query.filter(ApartmentRecord.must_have.isnot(None))
    elif must_have == '0':
        query = query.filter(ApartmentRecord.must_have.is_(None))

    # Order by real STT (nulls last), then id
    query = query.order_by(ApartmentRecord.stt.is_(None), ApartmentRecord.stt.asc(), ApartmentRecord.id.asc())

    # Pagination
    page = request.args.get('page', 1, type=int)
    per_page = 50
    records = query.paginate(page=page, per_page=per_page, error_out=False)

    # Get filter options (scoped to this category + region)
    opt = lambda col: db.session.query(col).filter_by(category=category, region=region).distinct().all()
    statuses = opt(ApartmentRecord.status)
    persons = opt(ApartmentRecord.person_in_charge)
    cities = opt(ApartmentRecord.city)
    classifications = opt(ApartmentRecord.classification)
    districts = opt(ApartmentRecord.district)
    building_names = opt(ApartmentRecord.building_name)

    return render_template('database.html',
                         category=category,
                         region=region,
                         records=records,
                         statuses=[s[0] for s in statuses if s[0]],
                         persons=[p[0] for p in persons if p[0]],
                         cities=[c[0] for c in cities if c[0]],
                         classifications=[c[0] for c in classifications if c[0]],
                         districts=sorted([d[0] for d in districts if d[0]]),
                         building_names=sorted({b[0] for b in building_names if b[0]}),
                         search=search,
                         current_status=status,
                         current_person=person,
                         current_city=city,
                         current_classification=classification,
                         current_district=district,
                         current_must_have=must_have)


@app.route('/database/<category>/<region>/add', methods=['GET', 'POST'])
@login_required
def add_record(category, region):
    """Add new record"""
    category = category.upper()
    region = region.upper()
    if not _valid_cat_region(category, region):
        flash('Invalid category/region', 'error')
        return redirect(url_for('index'))

    if request.method == 'POST':
        record = ApartmentRecord(
            category=category,
            region=region,
            stt=request.form.get('stt', type=int),
            team_assignment=request.form.get('team_assignment'),
            person_in_charge=_join_persons(request.form),
            status=request.form.get('status'),
            approach_time=request.form.get('approach_time'),
            notes=request.form.get('notes'),
            city=request.form.get('city'),
            direction=request.form.get('direction'),
            building_name=request.form.get('building_name'),
            district=request.form.get('district'),
            address=request.form.get('address'),
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
            prospect=request.form.get('prospect'),
            must_have=request.form.get('must_have') or None
        )
        db.session.add(record)
        db.session.commit()
        flash('Record added successfully', 'success')
        return redirect(url_for('database', category=category.lower(), region=region.lower()))

    # Suggest the next STT for this category + region (max existing + 1)
    max_stt = db.session.query(db.func.max(ApartmentRecord.stt)).filter_by(
        category=category, region=region).scalar()
    next_stt = (max_stt or 0) + 1
    return render_template('add_edit.html', category=category, region=region, record=None, next_stt=next_stt)


@app.route('/record/<int:id>/edit', methods=['GET', 'POST'])
@login_required
def edit_record(id):
    """Edit existing record"""
    record = ApartmentRecord.query.get_or_404(id)

    if request.method == 'POST':
        # Only the fields present on the form are updated; others (notes,
        # approach_time, direction, p9000/p6000, prospect, team_assignment,
        # address) are left untouched so they aren't wiped on edit.
        record.stt = request.form.get('stt', type=int)
        record.person_in_charge = _join_persons(request.form)
        record.status = request.form.get('status')
        if record.status == 'Deal':
            record.total_deployed = 0
        record.city = request.form.get('city')
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
        record.must_have = request.form.get('must_have') or None

        db.session.commit()
        flash('Record updated successfully', 'success')
        return redirect(url_for('database', category=record.category.lower(), region=record.region.lower()))

    return render_template('add_edit.html', category=record.category, region=record.region, record=record)


@app.route('/record/<int:id>/delete', methods=['POST'])
@login_required
def delete_record(id):
    """Delete record"""
    record = ApartmentRecord.query.get_or_404(id)
    category, region = record.category, record.region
    db.session.delete(record)
    db.session.commit()
    flash('Record deleted successfully', 'success')
    return redirect(url_for('database', category=category.lower(), region=region.lower()))


@app.route('/dashboard')
@login_required
def dashboard():
    """Dashboard with charts"""
    category = request.args.get('category', '').upper()
    category = category if category in ('AP', 'OB') else ''
    return render_template('dashboard.html', category=category)


# Conversion-rate funnel stages (columns of the two tables)
CONVERSION_STAGES = ['Plan B', 'Plan A', 'Deal', 'Done']


def _conversion_table(must_have_only=False, region=None, category=None):
    """Per-stage totals: screens, blocks, building count."""
    q = db.session.query(
        ApartmentRecord.status,
        db.func.coalesce(db.func.sum(ApartmentRecord.total_screens), 0),
        db.func.coalesce(db.func.sum(ApartmentRecord.num_blocks), 0),
        db.func.count(ApartmentRecord.id),
    ).filter(ApartmentRecord.status.in_(CONVERSION_STAGES))
    if must_have_only:
        q = q.filter(ApartmentRecord.must_have.isnot(None))
    if region:
        q = q.filter(ApartmentRecord.region == region)
    if category:
        q = q.filter(ApartmentRecord.category == category)
    q = q.group_by(ApartmentRecord.status)

    table = {s: {'screens': 0, 'blocks': 0, 'buildings': 0} for s in CONVERSION_STAGES}
    for status, screens, blocks, count in q.all():
        table[status] = {
            'screens': int(screens or 0),
            'blocks': int(blocks or 0),
            'buildings': int(count or 0),
        }
    return table


def _install_totals(region=None, category=None):
    """Totals from AP/OB for sites in Deal or Done status."""
    q = db.session.query(
        db.func.coalesce(db.func.sum(ApartmentRecord.total_screens), 0),
        db.func.coalesce(db.func.sum(ApartmentRecord.num_blocks), 0),
        db.func.count(ApartmentRecord.id),
        db.func.coalesce(db.func.sum(ApartmentRecord.total_deployed), 0),
    ).filter(ApartmentRecord.status.in_(['Deal', 'Done']))
    if region:
        q = q.filter(ApartmentRecord.region == region)
    if category:
        q = q.filter(ApartmentRecord.category == category)
    screens, blocks, buildings, deployed = q.one()
    screens = int(screens or 0)
    deployed = int(deployed or 0)
    return {
        'screens': screens,
        'blocks': int(blocks or 0),
        'buildings': int(buildings or 0),
        'deployed': deployed,
        'loading_pct': round(deployed / screens * 100, 1) if screens else 0.0,
    }


def _snapshot_current_week():
    """Upsert this ISO-week's screen totals per stage, for each category."""
    now = datetime.now()
    year, week, _ = now.isocalendar()

    for category in ('AP', 'OB'):
        rows = db.session.query(
            ApartmentRecord.status,
            db.func.coalesce(db.func.sum(ApartmentRecord.total_screens), 0),
            db.func.coalesce(db.func.sum(ApartmentRecord.num_blocks), 0),
        ).filter(
            ApartmentRecord.status.in_(CONVERSION_STAGES),
            ApartmentRecord.category == category,
        ).group_by(ApartmentRecord.status).all()
        screens = {s: 0 for s in CONVERSION_STAGES}
        blocks = {s: 0 for s in CONVERSION_STAGES}
        for status, sc, bl in rows:
            screens[status] = int(sc or 0)
            blocks[status] = int(bl or 0)

        snap = WeeklyGrowth.query.filter_by(category=category, year=year, week=week).first()
        if not snap:
            snap = WeeklyGrowth(category=category, year=year, week=week)
            db.session.add(snap)
        snap.plan_b = screens['Plan B']
        snap.plan_a = screens['Plan A']
        snap.deal = screens['Deal']
        snap.done = screens['Done']
        snap.plan_b_blocks = blocks['Plan B']
        snap.plan_a_blocks = blocks['Plan A']
        snap.deal_blocks = blocks['Deal']
        snap.done_blocks = blocks['Done']
        snap.updated_at = now
    db.session.commit()


def _weekly_growth_series(category=None):
    """Ordered weekly series for the growth line chart.

    category None -> sum of AP + OB per week; otherwise the given category.
    """
    try:
        _snapshot_current_week()
    except Exception as e:
        db.session.rollback()
        print(f"weekly snapshot error: {e}")
    q = WeeklyGrowth.query
    if category:
        q = q.filter_by(category=category)
    snaps = q.order_by(WeeklyGrowth.year, WeeklyGrowth.week).all()

    # Aggregate by (year, week) so "All" sums categories (AP + OB) into one point
    metrics = ['plan_b', 'plan_a', 'deal', 'done',
               'plan_b_blocks', 'plan_a_blocks', 'deal_blocks', 'done_blocks']
    buckets = {}
    order = []
    for s in snaps:
        key = (s.year, s.week)
        if key not in buckets:
            buckets[key] = {m: 0 for m in metrics}
            order.append(key)
        for m in metrics:
            buckets[key][m] += getattr(s, m, 0) or 0

    out = {'labels': [f'Tuần {w}' for (_, w) in order]}
    for m in metrics:
        out[m] = [buckets[k][m] for k in order]
    return out


@app.route('/conversion')
@login_required
def conversion():
    """Tỷ lệ chuyển đổi: two tables (all vs must-have), updated weekly."""
    region = request.args.get('region', '').upper()
    region = region if region in ('MN', 'MT', 'MB') else None
    category = request.args.get('category', '').upper()
    category = category if category in ('AP', 'OB') else None

    now = datetime.now()
    return render_template(
        'conversion.html',
        stages=CONVERSION_STAGES,
        table_all=_conversion_table(must_have_only=False, region=region, category=category),
        table_mh=_conversion_table(must_have_only=True, region=region, category=category),
        install=_install_totals(region=region, category=category),
        region=region,
        category=category,
        week=now.isocalendar()[1],
        updated=now.strftime('%d/%m/%Y'),
    )


@app.route('/api/stats')
@login_required
def api_stats():
    """JSON endpoint for dashboard charts"""
    category = request.args.get('category', '').upper()
    category = category if category in ('AP', 'OB') else None

    def scoped(query):
        return query.filter(ApartmentRecord.category == category) if category else query

    # Status distribution
    status_stats = scoped(db.session.query(
        ApartmentRecord.status,
        db.func.count(ApartmentRecord.id)
    )).group_by(ApartmentRecord.status).all()

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

    # Classification distribution — across ALL regions (MB + MT + MN), split by
    # AP and OB, counting only projects in the funnel (Research..Done). Labels
    # are normalized (trim) so ' B' merges into 'B'; non A/B/C values dropped.
    class_rows = db.session.query(
        ApartmentRecord.category,
        ApartmentRecord.classification,
        db.func.count(ApartmentRecord.id),
    ).filter(
        ApartmentRecord.classification.isnot(None),
        ApartmentRecord.status.in_(FUNNEL_STAGES),
    ).group_by(ApartmentRecord.category, ApartmentRecord.classification).all()

    ABC = ('A', 'B', 'C')
    cls_detail = {'AP': {}, 'OB': {}}
    cls_grouped = {'AP': {k: 0 for k in ABC}, 'OB': {k: 0 for k in ABC}}
    for cat, label, cnt in class_rows:
        if cat not in ('AP', 'OB'):
            continue
        name = str(label).strip()
        if not name:
            continue
        first = name[0].upper()
        if first not in ABC:
            continue
        cls_detail[cat][name] = cls_detail[cat].get(name, 0) + cnt
        cls_grouped[cat][first] += cnt
    detail_labels = sorted(set(cls_detail['AP']) | set(cls_detail['OB']))

    # MN vs MB by status
    region_status = db.session.query(
        ApartmentRecord.region,
        ApartmentRecord.status,
        db.func.count(ApartmentRecord.id)
    ).group_by(ApartmentRecord.region, ApartmentRecord.status).all()

    # Summary KPIs
    total = ApartmentRecord.query.count()
    total_mn = ApartmentRecord.query.filter_by(region='MN').count()
    total_mb = ApartmentRecord.query.filter_by(region='MB').count()
    status_map = {s[0]: s[1] for s in status_stats}
    funnel = {stage: status_map.get(stage, 0) for stage in FUNNEL_STAGES}
    funnel_total = sum(funnel.values())
    overall_progress = round(
        sum(funnel[s] * FUNNEL_WEIGHT[s] for s in FUNNEL_STAGES) / funnel_total * 100, 1
    ) if funnel_total else 0.0

    return jsonify({
        'summary': {
            'total': total,
            'total_mn': total_mn,
            'total_mb': total_mb,
            'done': funnel['Done'],
            'deal': funnel['Deal'],
            'funnel': funnel,
            'funnel_total': funnel_total,
            'overall_progress': overall_progress,
        },
        'person_progress': _compute_person_progress(category),
        'weekly_growth': _weekly_growth_series(category),
        'screen_progress': _screen_progress(),
        'funnel_stages': FUNNEL_STAGES,
        'status': [{'label': s[0], 'count': s[1]} for s in status_stats if s[0]],
        'city': [{'label': c[0], 'count': c[1]} for c in city_stats][:10],  # Top 10
        'person': [{'label': p[0], 'count': p[1]} for p in person_stats],
        'infrastructure': [{'label': i[0], 'count': i[1]} for i in infra_stats if i[0]],
        'classification_detail': {
            'labels': detail_labels,
            'AP': [cls_detail['AP'].get(l, 0) for l in detail_labels],
            'OB': [cls_detail['OB'].get(l, 0) for l in detail_labels],
        },
        'classification_grouped': {
            'labels': list(ABC),
            'AP': [cls_grouped['AP'][k] for k in ABC],
            'OB': [cls_grouped['OB'][k] for k in ABC],
        },
        'region_status': [{'region': r[0], 'status': r[1], 'count': r[2]} for r in region_status if r[1]]
    })


# AP/OB Excel columns — EXACT order of the database table/form. person_in_charge
# is split across 3 "Người PT" columns. Auxiliary/computed fields are excluded.
APOB_COLUMNS = [
    ('STT', 'stt'),
    ('Tên chung cư', 'building_name'),
    ('Số Block', 'num_blocks'),
    ('Quận', 'district'),
    ('Thành phố', 'city'),
    ('Trạng thái', 'status'),
    ('Người PT1', 'pt1'),
    ('Người PT2', 'pt2'),
    ('Người PT3', 'pt3'),
    ('Phân loại', 'classification'),
    ('Must have', 'must_have'),
    ('Giá bán', 'price_range'),
    ('CSVC', 'infrastructure'),
    ('Tỷ lệ lấp đầy', 'occupancy'),
    ('Đơn vị cũ', 'previous_operator'),
    ('Màn trong thang', 'screens_in_elevator'),
    ('Màn ngoài thang', 'screens_outside_elevator'),
    ('Tổng SL màn', 'total_screens'),
]
APOB_INT_FIELDS = {'stt', 'num_blocks', 'screens_in_elevator', 'screens_outside_elevator', 'total_screens'}
# Accept a few legacy/alias header names on import (old exports / variants)
APOB_IMPORT_ALIASES = {
    'Quận/Khu vực': 'district', 'TP/Tỉnh': 'city', 'Tiến độ': 'status',
    'Tổng SL màn hình': 'total_screens', 'Số màn trong thang': 'screens_in_elevator',
    'Số màn ngoài thang': 'screens_outside_elevator',
}


@app.route('/export/<category>/<region>')
@login_required
def export_data(category, region):
    """Export data as Excel matching the database table/form columns & order."""
    category = category.upper()
    region = region.upper()
    if not _valid_cat_region(category, region):
        flash('Invalid category/region', 'error')
        return redirect(url_for('index'))

    records = ApartmentRecord.query.filter_by(category=category, region=region).order_by(
        ApartmentRecord.stt.is_(None), ApartmentRecord.stt.asc(), ApartmentRecord.id.asc()).all()

    rows = []
    for r in records:
        pts = [p.strip() for p in str(r.person_in_charge or '').split(',') if p.strip()]
        pts += ['', '', '']
        row = {}
        for label, field in APOB_COLUMNS:
            if field == 'pt1':
                row[label] = pts[0]
            elif field == 'pt2':
                row[label] = pts[1]
            elif field == 'pt3':
                row[label] = pts[2]
            else:
                row[label] = getattr(r, field)
        rows.append(row)
    df = pd.DataFrame(rows, columns=[label for label, _ in APOB_COLUMNS])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=f'{category}_{region}')
    output.seek(0)

    filename = f'Database_{category}_{region}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    return send_file(output,
                    as_attachment=True,
                    download_name=filename,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/import/<category>/<region>', methods=['POST'])
@login_required
def import_data(category, region):
    """Import data from uploaded Excel/CSV"""
    category = category.upper()
    region = region.upper()
    if not _valid_cat_region(category, region):
        flash('Invalid category/region', 'error')
        return redirect(url_for('index'))

    back = redirect(url_for('database', category=category.lower(), region=region.lower()))
    if 'file' not in request.files:
        flash('Chưa chọn file', 'error')
        return back

    file = request.files['file']
    if file.filename == '':
        flash('Chưa chọn file', 'error')
        return back

    mode = request.form.get('mode', 'append')

    try:
        # Read file
        if file.filename.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)

        # Label -> field map (table columns + a few aliases). person handled below.
        field_for = {label: field for label, field in APOB_COLUMNS
                     if field not in ('pt1', 'pt2', 'pt3')}
        field_for.update(APOB_IMPORT_ALIASES)

        # Replace mode: clear existing records for this category + region first
        if mode == 'replace':
            ApartmentRecord.query.filter_by(category=category, region=region).delete()

        # Add records
        count = 0
        for _, row in df.iterrows():
            record = ApartmentRecord(category=category, region=region)
            for label in df.columns:
                field = field_for.get(str(label).strip())
                if not field:
                    continue
                value = row[label]
                if pd.isna(value):
                    continue
                if field in APOB_INT_FIELDS:
                    try:
                        value = int(float(str(value).replace("'", "").replace(",", "")))
                    except (ValueError, TypeError):
                        continue
                else:
                    value = str(value).strip()
                setattr(record, field, value)
            # Recombine Người PT1/2/3 -> person_in_charge (or legacy single column)
            pts = []
            for c in ('Người PT1', 'Người PT2', 'Người PT3'):
                if c in df.columns and pd.notna(row.get(c)):
                    t = str(row[c]).strip()
                    if t:
                        pts.append(t)
            if pts:
                record.person_in_charge = ', '.join(pts)
            elif 'NS. Phụ trách' in df.columns and pd.notna(row.get('NS. Phụ trách')):
                record.person_in_charge = str(row['NS. Phụ trách']).strip()
                
            if record.status == 'Deal':
                record.total_deployed = 0

            # Skip blank rows (no building name)
            if not record.building_name:
                continue
            db.session.add(record)
            count += 1

        db.session.commit()
        verb = 'thay thế toàn bộ bằng' if mode == 'replace' else 'thêm'
        flash(f'Đã {verb} {count} bản ghi ({category} {region})', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Lỗi import: {str(e)}', 'error')

    return back


@app.route('/import-export')
@login_required
def import_export():
    """Import/Export page"""
    return render_template('import_export.html')


@app.route('/install')
@login_required
def install():
    """DataBase Install: list of already-installed Digital Building sites."""
    category = request.args.get('category', '').upper()
    category = category if category in ('AP', 'OB') else ''
    search = request.args.get('search', '').strip()

    city = request.args.get('city', '').strip()
    status = request.args.get('status', '').strip()

    query = ApartmentRecord.query.filter(ApartmentRecord.status.in_(['Deal', 'Done']))
    if category:
        query = query.filter_by(category=category)
    if city:
        query = query.filter_by(city=city)
    if status:
        query = query.filter_by(status=status)
    if search:
        query = query.filter(db.or_(
            ApartmentRecord.building_name.contains(search),
            ApartmentRecord.district.contains(search),
            ApartmentRecord.city.contains(search),
        ))
    query = query.order_by(
        ApartmentRecord.city.is_(None), ApartmentRecord.city.asc(),
        ApartmentRecord.status.asc(),
        ApartmentRecord.building_name.is_(None), ApartmentRecord.building_name.asc(),
        ApartmentRecord.id.asc())

    page = request.args.get('page', 1, type=int)
    records = query.paginate(page=page, per_page=50, error_out=False)
    
    # Filter options for the UI
    base_q = ApartmentRecord.query.filter(ApartmentRecord.status.in_(['Deal', 'Done']))
    if category:
        base_q = base_q.filter_by(category=category)
    cities = [c[0] for c in base_q.with_entities(ApartmentRecord.city).distinct().all() if c[0]]
    building_names = [b[0] for b in base_q.with_entities(ApartmentRecord.building_name).distinct().all() if b[0]]


    totals = _install_totals(category=category or None)
    return render_template('install.html',
                           records=records,
                           category=category,
                           search=search,
                           city=city,
                           status=status,
                           cities=sorted(cities),
                           building_names=sorted(building_names),
                           totals=totals)


@app.route('/install/<int:id>/edit', methods=['GET', 'POST'])
@login_required
def install_edit(id):
    rec = ApartmentRecord.query.get_or_404(id)
    if rec.status not in ('Deal', 'Done'):
        flash('Dự án không ở trạng thái Deal/Done.', 'error')
        return redirect(url_for('install'))

    if request.method == 'POST':
        total_deployed = request.form.get('total_deployed', type=int) or 0
        if rec.status == 'Deal' and total_deployed > 0:
            flash('Dự án đang ở trạng thái Deal, vui lòng sang Database chuyển trạng thái thành Done trước khi nhập số lượng triển khai.', 'error')
        else:
            rec.total_deployed = total_deployed
            db.session.commit()
            flash('Đã cập nhật số lượng triển khai thực tế', 'success')
            return redirect(url_for('install'))
    return render_template('install_edit.html', record=rec)


@app.route('/install/export')
@login_required
def install_export():
    records = ApartmentRecord.query.filter(ApartmentRecord.status.in_(['Deal', 'Done'])).order_by(
        ApartmentRecord.city.is_(None), ApartmentRecord.city.asc(),
        ApartmentRecord.status.asc(),
        ApartmentRecord.building_name.is_(None), ApartmentRecord.building_name.asc(),
        ApartmentRecord.id.asc()).all()

    columns = [
        ('STT', 'stt'),
        ('Dự án (toà nhà)', 'building_name'),
        ('Khu vực', 'region'),
        ('Loại Hình', 'category'),
        ('Quận', 'district'),
        ('City', 'city'),
        ('Số Block', 'num_blocks'),
        ('Trạng thái', 'status'),
        ('Total (Thiết kế)', 'total_screens'),
        ('Total thực tế triển khai', 'total_deployed'),
        ('Loading %', 'loading'),
    ]

    rows = []
    for i, r in enumerate(records):
        excel_row = i + 2  # header is row 1
        row = {}
        for label, field in columns:
            if field == 'stt':
                row[label] = i + 1
            elif field == 'loading':
                # G is Total (Thiết kế), H is Total thực tế triển khai
                row[label] = f'=IF(G{excel_row}>0,ROUND(H{excel_row}/G{excel_row}*100,1),"")'
            else:
                row[label] = getattr(r, field)
        rows.append(row)
    df = pd.DataFrame(rows, columns=[label for label, _ in columns])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Install')
    output.seek(0)
    filename = f'Database_Install_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    return send_file(output, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)
