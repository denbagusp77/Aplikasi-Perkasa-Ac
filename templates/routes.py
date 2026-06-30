@app.route("/sync_services_to_transactions")
@login_required
def sync_services_to_transactions():
    from models import Service, Transaction
    services = Service.query.all()
    created = 0

    for s in services:
        # cek apakah sudah ada transaction untuk service ini
        exists = Transaction.query.filter_by(service_id=s.id).first()
        if exists:
            continue

        subtotal = (s.harga_satuan or 0) * (s.jumlah or 1)

        new_trans = Transaction(
            tanggal=s.tanggal,
            kategori="Pendapatan",
            jenis=s.jenis_service,
            deskripsi=f"Service {s.unit.customer.nama} - {s.unit.merk}",
            jumlah=subtotal,
            pendapatan_tipe="Internal" if s.team == "Internal" else "Eksternal",
            technician_id=s.technician_id,
            technician_nama=s.teknisi,
            team=s.team,
            service_id=s.id
        )
        db.session.add(new_trans)
        created += 1

    db.session.commit()
    flash(f"{created} transaksi berhasil dibuat dari service lama.", "success")
    return redirect(url_for("dashboard"))
