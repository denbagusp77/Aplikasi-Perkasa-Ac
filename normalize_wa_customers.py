from app import app, db, Customer, normalize_wa

with app.app_context():
    customers = Customer.query.all()
    seen = set()
    for c in customers:
        old = c.nomor_wa
        new = normalize_wa(old)
        # cek apakah nomor sudah dipakai customer lain
        if new in seen:
            print(f"⚠️ {c.id}: {old} -> {new} (DUPLIKAT, dilewati)")
            continue  # skip update
        # cek database langsung
        dupe = Customer.query.filter(Customer.nomor_wa == new, Customer.id != c.id).first()
        if dupe:
            print(f"⚠️ {c.id}: {old} -> {new} (DUPLIKAT di DB, dilewati)")
            continue
        c.nomor_wa = new
        seen.add(new)
        print(f"✅ {c.id}: {old} -> {new}")
    db.session.commit()

print("✅ Normalisasi selesai")
