#!/usr/bin/env python
"""Production Startup Script — Petra 360 Feedback System
Start: python run_prod.py
"""
import sys, os
os.chdir(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(__file__))

from waitress import serve
from app import app

HOST = os.environ.get('HOST', '0.0.0.0')
PORT = int(os.environ.get('PORT', 5000))
THREADS = int(os.environ.get('THREADS', 4))

print(f'[Petra 360] Starting production server on {HOST}:{PORT} (threads={THREADS})')
serve(app, host=HOST, port=PORT, threads=THREADS, channel_timeout=120)
