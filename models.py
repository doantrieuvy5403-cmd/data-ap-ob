from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()


class ApartmentRecord(db.Model):
    __tablename__ = "apartment_record"

    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(10), nullable=False, index=True, default='AP')  # "AP" or "OB"
    region = db.Column(db.String(10), nullable=False, index=True)  # "MN" or "MB"
    stt = db.Column(db.Integer)
    team_assignment = db.Column(db.String(255))
    person_in_charge = db.Column(db.String(255), index=True)
    status = db.Column(db.String(50), index=True)
    approach_time = db.Column(db.String(100))
    notes = db.Column(db.Text)
    city = db.Column(db.String(100), index=True)
    direction = db.Column(db.String(100))
    building_name = db.Column(db.String(255))
    district = db.Column(db.String(100))
    address = db.Column(db.String(255))  # OB street address
    num_blocks = db.Column(db.Integer)
    price_range = db.Column(db.String(100))
    infrastructure = db.Column(db.String(50))
    occupancy = db.Column(db.String(50))
    classification = db.Column(db.String(50))
    previous_operator = db.Column(db.String(100))
    total_screens = db.Column(db.Integer)
    screens_in_elevator = db.Column(db.Integer)
    screens_outside_elevator = db.Column(db.Integer)
    p9000 = db.Column(db.Integer)
    p6000 = db.Column(db.Integer)
    prospect = db.Column(db.String(100))
    must_have = db.Column(db.String(20), index=True)  # "Must have" or None
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    def to_dict(self):
        return {
            "id": self.id,
            "category": self.category,
            "region": self.region,
            "stt": self.stt,
            "team_assignment": self.team_assignment,
            "person_in_charge": self.person_in_charge,
            "status": self.status,
            "approach_time": self.approach_time,
            "notes": self.notes,
            "city": self.city,
            "direction": self.direction,
            "building_name": self.building_name,
            "district": self.district,
            "address": self.address,
            "num_blocks": self.num_blocks,
            "price_range": self.price_range,
            "infrastructure": self.infrastructure,
            "occupancy": self.occupancy,
            "classification": self.classification,
            "previous_operator": self.previous_operator,
            "total_screens": self.total_screens,
            "screens_in_elevator": self.screens_in_elevator,
            "screens_outside_elevator": self.screens_outside_elevator,
            "p9000": self.p9000,
            "p6000": self.p6000,
            "prospect": self.prospect,
            "must_have": self.must_have,
        }


class InstallRecord(db.Model):
    """Buildings already installed (Digital Building). Independent of the sales
    funnel (Research/Plan B/Plan A/Deal/Done) — these are live/installed sites."""
    __tablename__ = "install_record"

    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(10), index=True)       # "AP" or "OB" (Loại Hình)
    region = db.Column(db.String(10), index=True)         # derived "MN"/"MB" from city
    stt = db.Column(db.Integer)
    report_code = db.Column(db.String(120))
    name_of_block = db.Column(db.String(255))
    address_detail = db.Column(db.String(255))
    ward = db.Column(db.String(120))
    city = db.Column(db.String(120), index=True)
    dp_inside = db.Column(db.Integer)                      # DP — Digital Poster inside elevator
    dp_outside = db.Column(db.Integer)                     # DP/LCD — Digital Poster outside
    total = db.Column(db.Integer)                          # Total screens
    operational_status = db.Column(db.String(60), index=True)  # ON AIR / ON GOING ...
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    def to_dict(self):
        return {
            "id": self.id,
            "category": self.category,
            "region": self.region,
            "stt": self.stt,
            "report_code": self.report_code,
            "name_of_block": self.name_of_block,
            "address_detail": self.address_detail,
            "ward": self.ward,
            "city": self.city,
            "dp_inside": self.dp_inside,
            "dp_outside": self.dp_outside,
            "total": self.total,
            "operational_status": self.operational_status,
        }


class WeeklyGrowth(db.Model):
    """Weekly snapshot of screen counts per funnel stage (for growth chart)."""
    __tablename__ = "weekly_growth"

    id = db.Column(db.Integer, primary_key=True)
    category = db.Column(db.String(10), nullable=False, default='AP')  # "AP" or "OB"
    year = db.Column(db.Integer, nullable=False)
    week = db.Column(db.Integer, nullable=False)
    plan_b = db.Column(db.Integer, default=0)
    plan_a = db.Column(db.Integer, default=0)
    deal = db.Column(db.Integer, default=0)
    done = db.Column(db.Integer, default=0)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    __table_args__ = (db.UniqueConstraint('category', 'year', 'week', name='uq_weekly_cat_year_week'),)


class AppMeta(db.Model):
    """Simple key/value store for app state (e.g. seeded data.xlsx hash)."""
    __tablename__ = "app_meta"

    key = db.Column(db.String(50), primary_key=True)
    value = db.Column(db.String(255))
