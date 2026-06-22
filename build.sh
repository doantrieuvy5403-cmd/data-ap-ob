#!/usr/bin/env bash
set -e

# Install dependencies
pip install -r requirements.txt

# Seed database if data.xlsx exists
if [ -f "data.xlsx" ]; then
    echo "Seeding database from data.xlsx..."
    python seed.py
else
    echo "Warning: data.xlsx not found, skipping database seed"
fi
