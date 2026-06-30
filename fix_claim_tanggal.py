# fix_claim_tanggal.py
import os, sys
sys.path.append(os.path.dirname(__file__))

from app import app, db, Claim
from datetime import datetime

with app.app_context():
    claims = Claim.query.all()
    fixed = 0

    for c in claims:
        t = c.tanggal

        # 1️⃣ Jika sudah datetime, skip
        if isinstance(t, datetime):
            continue

        # 2️⃣ Jika string, coba parse
        if isinstance(t, str):
            parsed = None
            for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y-%m-%dT%H:%M:%S"):
                try:
                    parsed = datetime.strptime(t, fmt)
                    break
                except Exception:
                    pass
            if parsed:
                c.tanggal = parsed
                fixed += 1
            else:
                # format tidak dikenal, isi tanggal sekarang
                c.tanggal = datetime.utcnow()
                fixed += 1

        # 3️⃣ Kalau None atau tipe lain
        elif not t:
            c.tanggal = datetime.utcnow()
            fixed += 1

    db.session.commit()
    print(f"{fixed} baris Claim diperbaiki.")
