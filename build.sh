#!/usr/bin/env bash
set -e

# Install dependencies
pip install -r requirements.txt

# NOTE: Database seeding is handled by app.py (_auto_seed) on startup, and ONLY
# when the database is empty. Do NOT run seed.py here — it calls db.drop_all()
# which would wipe all data (including manually-added rows) on every deploy.
