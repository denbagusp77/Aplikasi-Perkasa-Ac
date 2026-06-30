from app import app, db

with app.app_context():
    with db.engine.connect() as conn:
        result = conn.execute(db.text("PRAGMA table_info('transaction');"))
        print("📋 Daftar kolom di tabel 'transaction':")
        for row in result:
            print(row)
