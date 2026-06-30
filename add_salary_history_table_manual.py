from app import app, db
from sqlalchemy import Column, Integer, String, Float, Date, ForeignKey
from sqlalchemy.orm import relationship

# 🧩 Definisikan tabel salary_history baru
class SalaryHistory(db.Model):
    __tablename__ = 'salary_history'

    id = Column(Integer, primary_key=True)
    employee_id = Column(Integer, ForeignKey('technician.id', ondelete='CASCADE'), nullable=False)
    periode = Column(String(20), nullable=False)
    gaji_pokok = Column(Float, default=0.0)
    bonus = Column(Float, default=0.0)
    potongan = Column(Float, default=0.0)
    total_gaji = Column(Float, default=0.0)
    dibayarkan_tanggal = Column(Date, nullable=True)

    employee = relationship("Technician", backref="riwayat_gaji")

if __name__ == '__main__':
    with app.app_context():  # 🩵 tambahkan ini agar db.engine bisa diakses
        existing = db.inspect(db.engine).get_table_names()
        if 'salary_history' in existing:
            print("⚠️ Tabel salary_history sudah ada, tidak dibuat ulang.")
        else:
            print("🧱 Membuat tabel salary_history...")
            SalaryHistory.__table__.create(db.engine)
            print("✅ Tabel salary_history berhasil dibuat.")
