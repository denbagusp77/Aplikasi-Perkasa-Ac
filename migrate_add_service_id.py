from app import app, db
from sqlalchemy import text

with app.app_context():
    try:
        db.session.execute(text("""
            ALTER TABLE "transaction"
            ADD COLUMN service_id INTEGER REFERENCES service(id) ON DELETE CASCADE
        """))
        db.session.commit()
        print("✅ Kolom service_id berhasil ditambahkan.")
    except Exception as e:
        print("⚠️ Error:", e)
