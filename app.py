# import click
# from flask.cli import with_appcontext
from flask import (
    Flask, render_template, redirect, url_for, request, flash,
    send_from_directory, make_response, Response,Blueprint
)
from sqlalchemy import event  # ⬅️ tambahkan ini di atas
import pdfkit
import re
from flask import make_response
from werkzeug.utils import secure_filename
import qrcode
from reportlab.lib.utils import ImageReader
from collections import defaultdict
from calendar import monthrange
from datetime import date, timedelta


from sqlalchemy import text 
from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    LoginManager, login_user, login_required,
    logout_user, current_user, UserMixin
)
from flask_migrate import Migrate
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import or_, func, extract
from sqlalchemy.orm import aliased
from datetime import date, datetime,timedelta
from dateutil.relativedelta import relativedelta
import os, io, json, requests
from urllib.parse import quote
from flask import jsonify
from sqlalchemy.ext.hybrid import hybrid_property
from flask import send_file
import requests
from sqlalchemy import or_



# Excel & PDF
from openpyxl import Workbook
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle,
    Paragraph, Spacer, Image
)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet


# =========================
# APP CONFIG
# =========================
app = Flask(__name__)
app.config['SECRET_KEY'] = 'ac-service-secret'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///ac_service.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)
migrate = Migrate(app, db)

# Login Manager
login_manager = LoginManager(app)
login_manager.login_view = 'login'

# Models
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    def set_password(self, pw):
        self.password_hash = generate_password_hash(pw)
    def check_password(self, pw):
        return check_password_hash(self.password_hash, pw)

class Customer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nama = db.Column(db.String(200), nullable=False)
    nomor_wa = db.Column(db.String(100), unique=True, nullable=True)
    alamat = db.Column(db.String(400))
    team = db.Column(db.String(50))  # Internal / Eksternal

    # Relasi
    ac_units = db.relationship(
        'ACUnit', backref='customer',
        cascade='all, delete-orphan', lazy=True
    )
    reminders = db.relationship(
        'Reminder', backref='customer',
        cascade='all, delete-orphan', lazy=True
    )
   # 🔹 event listener normalisasi dan cek duplikat
@event.listens_for(Customer, "before_insert")
@event.listens_for(Customer, "before_update")
def normalize_and_check(mapper, connection, target):
    """
    target = instance Customer yang lagi diinsert/update
    """
    # Normalisasi nomor WA
    target.nomor_wa = normalize_wa(target.nomor_wa)

    # Cek apakah nomor WA sudah ada di DB (selain dirinya sendiri)
    if target.nomor_wa:
        customer_table = Customer.__table__
        # connection.execute() → cek langsung SQL
        result = connection.execute(
            customer_table.select()
            .where(customer_table.c.nomor_wa == target.nomor_wa)
            .where(customer_table.c.id != (target.id or 0))
        ).first()

        if result:
            raise ValueError(f"Nomor WA {target.nomor_wa} sudah dipakai customer lain")

class ACUnit(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    customer_id = db.Column(db.Integer, db.ForeignKey('customer.id'), nullable=False)
    ruangan = db.Column(db.String(200))
    merk = db.Column(db.String(100))
    pk = db.Column(db.String(50))
    jumlah_unit = db.Column(db.Integer, default=1)

    services = db.relationship(
        'Service',
        back_populates="ac_unit",
        cascade='all, delete-orphan',
        lazy=True,
        order_by="desc(Service.tanggal)"
    )


class ServiceType(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nama = db.Column(db.String(200), nullable=False)
    harga = db.Column(db.Float, default=0.0)
    


class Package(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nama = db.Column(db.String(200), nullable=False)
    harga = db.Column(db.Float, default=0.0)

class Technician(db.Model):
    __tablename__ = "technician"

    id = db.Column( db.Integer, primary_key=True)
    nama = db.Column(db.String(200), nullable=False)
    team = db.Column(db.String(50))
    jabatan = db.Column(db.String(100))               
    no_hp = db.Column(db.String(20))            
    alamat = db.Column(db.Text)                       
    tanggal_masuk = db.Column(db.Date)    
    tanggal_keluar = db.Column(db.Date, nullable=True)           
    status = db.Column(db.String(50), default="Aktif") 
    gaji_pokok = db.Column(db.Float, default=0.0)     
    markup = db.Column(db.Float, default=0.28)
    catatan = db.Column(db.Text)      
    foto = db.Column(db.String(255))     # 🆕 untuk nama file foto
    dokumen = db.Column(db.String(255))  # 🆕 untuk nama file dokumen                 

    # 👇 satu relasi cukup
    services = db.relationship("Service", back_populates="technician", lazy=True)



class Attendance(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    technician_id = db.Column(db.Integer, db.ForeignKey('technician.id'), nullable=False)
    tanggal = db.Column(db.Date, default=date.today)
    jam_masuk = db.Column(db.Time)
    jam_pulang = db.Column(db.Time)
    status = db.Column(db.String(20))  # "Hadir", "Terlambat", "Izin", dll
    lokasi = db.Column(db.String(100))
    keterangan = db.Column(db.String(255))

    # ✅ Relasi ke tabel Technician
    technician = db.relationship('Technician', backref='attendances')

class Kasbon(db.Model):
    __tablename__ = 'kasbon'

    id = db.Column(db.Integer, primary_key=True)
    technician_id = db.Column(db.Integer, db.ForeignKey('technician.id'), nullable=False)
    tanggal = db.Column(db.Date, default=date.today)
    jumlah = db.Column(db.Float, nullable=False)
    keterangan = db.Column(db.String(255))
    status = db.Column(db.String(50), default='Belum Lunas')  # Belum Lunas / Lunas
    terpilih = db.Column(db.Boolean, default=False)

    # 🔹 Tambahan kolom baru untuk menyimpan hubungan ke payroll
    payroll_id = db.Column(db.Integer, db.ForeignKey('payroll.id'), nullable=True)

    # 🔹 Relasi
    technician = db.relationship('Technician', backref='kasbon_list')
    payroll = db.relationship('Payroll', backref='kasbon_list', lazy=True)

class Service(db.Model):
    __tablename__ = "service"

    id = db.Column(db.Integer, primary_key=True)

    # ✅ hanya satu FK ke ACUnit
    unit_id = db.Column(db.Integer, db.ForeignKey('ac_unit.id'), nullable=False)

    tanggal = db.Column(db.Date, default=date.today)
    tanggal_cuci_berikutnya = db.Column(db.Date, nullable=True)
    jenis_service = db.Column(db.String(200), nullable=False)
    paket_cuci = db.Column(db.String(200))

    technician_id = db.Column(db.Integer, db.ForeignKey("technician.id"), nullable=True)
    teknisi = db.Column(db.String(100))
    team = db.Column(db.String(50))

    harga_satuan = db.Column(db.Float, default=0.0)
    jumlah = db.Column(db.Integer, default=1)
    pembayaran = db.Column(db.String(100), default="Belum Lunas")
    deskripsi = db.Column(db.Text)

    # ✅ Relasi ACUnit
    ac_unit = db.relationship("ACUnit", back_populates="services", foreign_keys=[unit_id])

    # ✅ Relasi Technician
    technician = db.relationship("Technician", back_populates="services")

    @hybrid_property
    def total_harga(self):
        return (self.harga_satuan or 0) * (self.jumlah or 0)


def bisa_claim(self):
        if not self.tanggal:
            return False
        return (datetime.utcnow().date() - self.tanggal).days <= 30

class ServiceMotor(db.Model):
    __tablename__ = 'service_motor'

    id = db.Column(db.Integer, primary_key=True)
    tanggal = db.Column(db.Date, nullable=False)
    technician_id = db.Column(db.Integer, db.ForeignKey('technician.id'), nullable=False)
    persen_teknisi = db.Column(db.Float, default=50)
    persen_perusahaan = db.Column(db.Float, default=50)

    total_biaya = db.Column(db.Float, nullable=False)
    porsi_teknisi = db.Column(db.Float, nullable=True)
    porsi_perusahaan = db.Column(db.Float, nullable=True)
    keterangan = db.Column(db.Text)
    status = db.Column(db.String(20), default='Belum Dipotong')

    terpilih = db.Column(db.Boolean, default=False)  # ✅ seperti kasbon
    payroll_id = db.Column(db.Integer, db.ForeignKey('payroll.id'), nullable=True)

    technician = db.relationship('Technician', backref='service_motor')

class ServiceMotorHistory(db.Model):
    __tablename__ = "service_motor_history"

    id = db.Column(db.Integer, primary_key=True)
    service_motor_id = db.Column(
        db.Integer,
        db.ForeignKey('service_motor.id'),
        nullable=False
    )

    persen_teknisi = db.Column(db.Float, nullable=False)
    persen_perusahaan = db.Column(db.Float, nullable=False)

    porsi_teknisi = db.Column(db.Float, nullable=False)
    porsi_perusahaan = db.Column(db.Float, nullable=False)

    diubah_oleh = db.Column(db.String(50))
    diubah_pada = db.Column(db.DateTime, default=datetime.utcnow)

    service_motor = db.relationship(
        'ServiceMotor',
        backref=db.backref(
            'histories',
            cascade="all, delete-orphan",   # ✅ INI KUNCINYA
            lazy=True,
            order_by="desc(ServiceMotorHistory.diubah_pada)"
        )
    )


class Payroll(db.Model):
    __tablename__ = "payroll"

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('technician.id'), nullable=False)
    periode = db.Column(db.String(7), nullable=False)  # contoh: "2025-09"

    total_pendapatan = db.Column(db.Float, default=0)
    total_potongan = db.Column(db.Float, default=0)
    gaji_bersih = db.Column(db.Float, default=0)

    total_tidak_hadir = db.Column(db.Integer, default=0)
    tanggal_dibuat = db.Column(db.DateTime, default=datetime.utcnow)
    status = db.Column(db.String(20), default='Draft')

    # Simpan daftar ID kasbon & service motor (pakai PickleType agar mudah menyimpan list)
    kasbon_ids = db.Column(db.PickleType)
    service_motor_ids = db.Column(db.PickleType)

    # Relasi
    employee = db.relationship('Technician', backref='payrolls')
    items = db.relationship('PayrollItem', backref='payroll', cascade="all, delete-orphan")

class PayrollSetting(db.Model):
    __tablename__ = "payroll_setting"

    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(
        db.Integer,
        db.ForeignKey('technician.id', name="fk_payroll_technician_id"),  # ✅ beri nama
        nullable=False
    )
    gaji_pokok = db.Column(db.Float)
    uang_makan_harian = db.Column(db.Float)
    uang_bensin_harian = db.Column(db.Float)
    bonus_bulanan = db.Column(db.Float)
    potongan_absen_per_hari = db.Column(db.Float)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    employee = db.relationship("Technician", backref="payroll_setting", lazy=True)




class PayrollItem(db.Model):
    __tablename__ = "payroll_item"

    id = db.Column(db.Integer, primary_key=True)
    payroll_id = db.Column(db.Integer, db.ForeignKey('payroll.id', ondelete="CASCADE"), nullable=False)
    kategori = db.Column(db.String(20))  # "Pendapatan" atau "Potongan"
    deskripsi = db.Column(db.String(100))
    jumlah = db.Column(db.Float, default=0.0)

class PayrollSlip(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    payroll_id = db.Column(db.Integer, db.ForeignKey('payroll.id', ondelete="CASCADE"))
    file_path = db.Column(db.String(255))  # hasil PDF disimpan di folder static/payroll_slip/
    generated_at = db.Column(db.DateTime, default=datetime.utcnow)

    payroll = db.relationship('Payroll', backref='slip')

class PayrollLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    payroll_id = db.Column(db.Integer, db.ForeignKey('payroll.id'))
    action = db.Column(db.String(100))
    user = db.Column(db.String(50))
    timestamp = db.Column(db.DateTime, default=datetime.utcnow)

    payroll = db.relationship('Payroll', backref=db.backref('logs', lazy=True))


class Reminder(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    customer_id = db.Column(db.Integer, db.ForeignKey('customer.id'))
    service_id = db.Column(db.Integer, db.ForeignKey('service.id'), nullable=True)  # tambah ini
    pesan = db.Column(db.String(500))
    tanggal = db.Column(db.Date)
    status = db.Column(db.String(50), default='pending')
    keterangan_reminder = db.Column(db.String(50), default="Belum di-Reminder")
    reminder_sent = db.Column(db.Boolean, default=False)

    service = db.relationship('Service', backref='reminders')

class BlastLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    message = db.Column(db.Text, nullable=False)
    targets = db.Column(db.Text, nullable=False)  # disimpan sebagai string gabungan nomor
    status = db.Column(db.String(50))  # sukses / gagal
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


class Transaction(db.Model):
    __tablename__ = "transaction"

    id = db.Column(db.Integer, primary_key=True)
    tanggal = db.Column(db.Date, default=date.today, nullable=False)
    kategori = db.Column(db.String(50), nullable=False)   # Pendapatan / Beban
    jenis = db.Column(db.String(100), nullable=False)     # Service, Sparepart, Gaji, dll.
    deskripsi = db.Column(db.String(255))
    jumlah = db.Column(db.Float, default=0.0, nullable=False)

    # 🔹 kolom tambahan
    pendapatan_tipe = db.Column(db.String(20), nullable=True)  # Internal / Eksternal
    technician_id = db.Column(db.Integer, db.ForeignKey('technician.id'), nullable=True)
    technician_nama = db.Column(db.String(100), nullable=True)
    team = db.Column(db.String(50), nullable=True)

    # 🔹 relasi ke Technician
    technician = db.relationship('Technician', backref='transactions', lazy=True)

    # 🔹 foreign key ke Service
    service_id = db.Column(db.Integer, db.ForeignKey('service.id', ondelete="CASCADE"), nullable=True)
    service = db.relationship('Service', backref='transactions', lazy=True)

class Claim(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    customer_id = db.Column(db.Integer, db.ForeignKey('customer.id'), nullable=False)
    unit_id = db.Column(db.Integer, db.ForeignKey('ac_unit.id'), nullable=False)
    technician_id = db.Column(db.Integer, db.ForeignKey('technician.id'))
    tanggal = db.Column(db.DateTime, default=datetime.utcnow)
    deskripsi = db.Column(db.Text, nullable=False)
    foto = db.Column(db.String(255))  
    status = db.Column(db.String(20), default="Pending")  

    # ForeignKey ke service
    service_id = db.Column(db.Integer, db.ForeignKey("service.id"), nullable=True)

    # Relasi
    customer = db.relationship('Customer', backref=db.backref('claims', lazy=True))
    unit = db.relationship('ACUnit', backref=db.backref('claims', lazy=True))
    technician = db.relationship('Technician', backref=db.backref('claims', lazy=True))
    service = db.relationship('Service', backref=db.backref('claims', lazy=True))  # ✅ baru

    @property
    def teknisi_display(self):
        """Tampilkan nama teknisi dengan prioritas:
        1. Dari service (kalau ada)
        2. Dari klaim langsung (kalau ada)
        3. Default '-'
        """
        # Ambil dari service kalau ada
        if self.service and self.service.technician:
            return f"{self.service.technician.nama} ({self.service.technician.team})"

        # Kalau ada teknisi langsung di klaim
        if self.technician:
            return f"{self.technician.nama} ({self.technician.team})"

        return "-"

class Kwitansi(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nomor_kwitansi = db.Column(db.String(100))
    invoice_id = db.Column(db.Integer, db.ForeignKey("invoice.id"))
    nomor = db.Column(db.String(50))
    tanggal = db.Column(db.Date)
    jumlah = db.Column(db.Integer)
    terbilang = db.Column(db.Text)


class Invoice(db.Model):
    __tablename__ = "invoice"

    id = db.Column(db.Integer, primary_key=True)
    spreadsheet_link = db.Column(db.String(255))
    customer_id = db.Column(
        db.Integer,
        db.ForeignKey('customer.id', ondelete="CASCADE"),
        nullable=False
    )
    spo = db.Column(db.String(100))
    tanggal = db.Column(db.Date, default=date.today)
    total = db.Column(db.Float, default=0.0)
    template = db.Column(db.String(20), default="classic")
    bank_account_id = db.Column(db.Integer, db.ForeignKey("bank_account.id"))
    bank_account = db.relationship("BankAccount")
    customer = db.relationship(
        'Customer',
        backref=db.backref('invoices', cascade="all, delete-orphan", lazy=True)
    )

    items = db.relationship(
        'InvoiceItem',
        backref='invoice',
        cascade="all, delete-orphan",
        lazy=True
    )


class BankAccount(db.Model):
    __tablename__ = "bank_account"
    id = db.Column(db.Integer, primary_key=True)
    bank_name = db.Column(db.String(100), nullable=False)
    account_name = db.Column(db.String(100), nullable=False)
    account_number = db.Column(db.String(50), nullable=False)


class InvoiceItem(db.Model):
    __tablename__ = "invoice_item"

    id = db.Column(db.Integer, primary_key=True)
    invoice_id = db.Column(
        db.Integer,
        db.ForeignKey('invoice.id', ondelete="CASCADE"),  # biar ikut kehapus
        nullable=False
    )
    service_id = db.Column(db.Integer, db.ForeignKey('service.id'))
    deskripsi = db.Column(db.String(255))
    harga = db.Column(db.Float, default=0.0)

    service = db.relationship('Service', backref='invoice_items')

# user loader
@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# 🔹 letakkan fungsi di sini:
# di app/__init__.py atau helpers.py


def normalize_wa(raw_wa):
    if not raw_wa:
        return None

    # pastikan selalu string
    raw_wa = str(raw_wa).strip()

    # hapus semua non-digit (spasi, +, -, titik, dll.)
    cleaned = re.sub(r'\D', '', raw_wa)

    if not cleaned:
        return None

    # normalisasi prefix
    if cleaned.startswith('0'):
        cleaned = '62' + cleaned[1:]
    elif cleaned.startswith('620'):
        cleaned = '62' + cleaned[3:]
    elif not cleaned.startswith('62'):
        cleaned = '62' + cleaned  # default tambahkan 62

    return cleaned



    

# Routes - auth
@app.route('/')
def index():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method=='POST':
        u = User.query.filter_by(username=request.form['username']).first()
        if u and u.check_password(request.form['password']):
            login_user(u)
            return redirect(url_for('dashboard'))
        flash('Username atau password salah', 'danger')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

  

# @app.cli.command("sync-services")
# @with_appcontext
# def sync_services():
#     """Sinkronkan semua Service lama ke Transaction"""
#     services = Service.query.all()
#     created = 0

#     for s in services:
#         # kalau service ini belum ada transaksi
#         if not s.transactions:
#             subtotal = (s.harga_satuan or 0) * (s.jumlah or 1)

#             pendapatan_tipe = None
#             if s.team:
#                 pendapatan_tipe = "Internal" if s.team == "Internal" else "Eksternal"

#             base_desc = f"Service {s.unit.merk} ({s.unit.pk}) - {s.unit.customer.nama}"
#             if s.deskripsi:
#                 base_desc += f" | {s.deskripsi}"

#             t = Transaction(
#                 service_id=s.id,
#                 tanggal=s.tanggal,
#                 kategori="Pendapatan",
#                 jenis=s.jenis_service or "-",
#                 deskripsi=base_desc,
#                 jumlah=subtotal,
#                 pendapatan_tipe=pendapatan_tipe,
#                 technician_id=s.technician_id,
#                 technician_nama=s.teknisi,
#                 team=s.team
#             )
#             db.session.add(t)
#             created += 1

#     db.session.commit()
#     click.echo(f"✅ Sinkronisasi selesai: {created} transaksi baru dibuat.")


# Dashboard

@app.route('/dashboard')
@login_required
def dashboard():
    total_customers = Customer.query.count()
    total_units = ACUnit.query.count()
    today_services = Service.query.filter_by(tanggal=date.today()).count()

    today = date.today()
    month_start = today.replace(day=1)

    # Pendapatan bulan ini
    revenue = (
        db.session.query(func.sum(Transaction.jumlah))
        .filter(
            Transaction.tanggal.between(month_start, today),
            Transaction.kategori == "Pendapatan"
        )
        .scalar() or 0
    )

    # Total keseluruhan pendapatan
    total_all_revenue = (
        db.session.query(func.sum(Transaction.jumlah))
        .filter(Transaction.kategori == "Pendapatan")
        .scalar() or 0
    )

    # Auto Reminders
    auto_reminders = []
    services = Service.query.filter(Service.jenis_service.ilike('%cuci%')).all()
    for s in services:
        next_date = s.tanggal_cuci_berikutnya or (s.tanggal + relativedelta(months=3))
        h7_date = next_date - relativedelta(days=7)
        h1_date = next_date - relativedelta(days=1)
        if today in (h7_date, h1_date):
            auto_reminders.append({
                "nama": s.ac_unit.customer.nama,
               "unit": f"{s.ac_unit.merk} ({s.ac_unit.pk}) x{s.ac_unit.jumlah_unit}",
                "tanggal": next_date,
                "tipe": "H-7" if today == h7_date else "H-1"
            })
    reminders = sorted(auto_reminders, key=lambda x: x["tanggal"])[:5]

    # Aktivitas terakhir
    recent_services = Service.query.order_by(Service.tanggal.desc()).limit(10).all()

    # Unpaid services
    unpaid_services = (
        Service.query
        .filter(Service.pembayaran == 'Belum Lunas')  
        .order_by(Service.tanggal.desc())
        .all()
    )

    # Best Seller
    best_sellers_query = (
        db.session.query(Service.jenis_service, func.count(Service.id).label("jumlah"))
        .filter(Service.tanggal.between(month_start, today))
        .group_by(Service.jenis_service)
        .order_by(func.count(Service.id).desc())
        .limit(3)
        .all()
    )
    best_sellers = [{"jenis": row[0], "jumlah": row[1]} for row in best_sellers_query]

    # Best Paket Cuci
    paket_cuci_query = (
        db.session.query(Service.paket_cuci, func.count(Service.id).label("jumlah"))
        .filter(
            Service.tanggal.between(month_start, today),
            Service.jenis_service.ilike('%cuci%'),
            Service.paket_cuci.isnot(None)
        )
        .group_by(Service.paket_cuci)
        .order_by(func.count(Service.id).desc())
        .all()
    )
    best_cuci_packages = [{"paket": row[0], "jumlah": row[1]} for row in paket_cuci_query]

    # --- Tambahkan: hanya customer yang punya claim ---
    customers_with_claims = (
        db.session.query(Customer)
        .join(Claim)
        .group_by(Customer.id)
        .all()
    )

    return render_template(
        "dashboard_modern.html",
        total_customers=total_customers,
        total_units=total_units,
        today_services=today_services,
        revenue=revenue,
        total_all_revenue=total_all_revenue,
        auto_reminders=reminders,
        recent_services=recent_services,
        unpaid_services=unpaid_services,
        best_sellers=best_sellers,
        best_cuci_packages=best_cuci_packages,
        customers=customers_with_claims  # <-- kirim ke template
    )

# Route Blast


FONNTE_TOKEN = "9A4s3ixqbb31TqxoUiHe"

# Folder upload sementara
UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER



from io import BytesIO
import mimetypes

@app.route('/blast', methods=['GET', 'POST'], endpoint='blast') 
@login_required
def blast_page():
    if request.method == 'POST':
        pesan = request.form.get('pesan')
        selected_ids = request.form.getlist('customer_ids')
        file = request.files.get("file")  # ambil file dari form (opsional)

        if not pesan and not file:
            flash("⚠️ Harus isi pesan atau unggah file!", "warning")
            return redirect(url_for('blast'))

        if not selected_ids:
            flash("⚠️ Harus pilih customer!", "warning")
            return redirect(url_for('blast'))

        # Ambil nomor WA customer
        customers = Customer.query.filter(Customer.id.in_(selected_ids)).all()
        targets = [c.nomor_wa for c in customers if c.nomor_wa]

        if not targets:
            flash("⚠️ Tidak ada nomor WA valid!", "warning")
            return redirect(url_for('blast'))

        url = "https://api.fonnte.com/send"
        headers = {"Authorization": FONNTE_TOKEN}
        status = "failed"
        res_json = {}

        try:
            files = {}
            # Kirim file jika ada
            if file and file.filename:
                filename = secure_filename(file.filename)
                file_stream = BytesIO(file.read())
                mime_type, _ = mimetypes.guess_type(filename)
                if not mime_type:
                    mime_type = "application/octet-stream"

                # Gunakan field 'media' sesuai Fonnte
                files["media"] = (filename, file_stream, mime_type)

            # Teks dikirim sebagai field 'message'
            files["message"] = (None, pesan or "")

            # Kirim request POST
            response = requests.post(url, headers=headers, files=files)
            res_json = response.json()
            status = "sent" if response.status_code == 200 and res_json.get("status") else "failed"

        except Exception as e:
            status = "failed"
            res_json = {"error": str(e)}

        # Simpan log ke DB
        new_blast = BlastLog(
            message=pesan or "(media)",
            targets=",".join(targets),
            status=status,
            created_at=datetime.utcnow()
        )
        db.session.add(new_blast)
        db.session.commit()

        flash(f"✅ Blast dikirim ({status}) ke {len(targets)} nomor", "success")
        print("Fonnte response:", res_json)

        return redirect(url_for('blast'))

    # GET: tampilkan form + log
    customers = Customer.query.all()
    blasts = BlastLog.query.order_by(BlastLog.created_at.desc()).all()
    return render_template("blast.html", customers=customers, blasts=blasts)


@app.route('/transactions')
@login_required
def transactions():
    tahun = request.args.get('tahun', datetime.today().year, type=int)
    bulan = request.args.get('bulan', datetime.today().month, type=int)
    show_all = request.args.get('all', '0') == '1'
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')

    # 🔹 Ambil transaksi langsung dari tabel Transaction
    query = Transaction.query

    if start_date and end_date:
        try:
            sd = datetime.strptime(start_date, "%Y-%m-%d").date()
            ed = datetime.strptime(end_date, "%Y-%m-%d").date()
            query = query.filter(Transaction.tanggal.between(sd, ed))
        except:
            pass
    elif not show_all:  # default bulan berjalan
        query = query.filter(
            extract('year', Transaction.tanggal) == tahun,
            extract('month', Transaction.tanggal) == bulan
        )

    # 🔹 Ambil daftar transaksi
    trans = query.order_by(Transaction.tanggal.desc()).all()

    # 🔹 Ringkasan
    total_pendapatan = sum(t.jumlah for t in trans if t.kategori == "Pendapatan")
    total_pengeluaran = sum(t.jumlah for t in trans if t.kategori == "Beban")
    saldo = total_pendapatan - total_pengeluaran
    total_internal = sum(t.jumlah for t in trans if t.kategori == "Pendapatan" and t.pendapatan_tipe == "Internal")
    total_eksternal = sum(t.jumlah for t in trans if t.kategori == "Pendapatan" and t.pendapatan_tipe == "Eksternal")

    # ==========================================================
    # 🔹 Perhitungan Per Teknisi (mengikuti filter aktif)
    # ==========================================================
    filtered_ids = [t.id for t in trans]

    # Pendapatan
    pendapatan_teknisi = (
        db.session.query(
            Transaction.technician_nama,
            func.sum(Transaction.jumlah)
        )
        .filter(Transaction.id.in_(filtered_ids))
        .filter(Transaction.kategori == "Pendapatan")
        .group_by(Transaction.technician_nama)
        .all()
    )

    # Pengeluaran
    pengeluaran_teknisi = (
        db.session.query(
            Transaction.technician_nama,
            func.sum(Transaction.jumlah)
        )
        .filter(Transaction.id.in_(filtered_ids))
        .filter(Transaction.kategori == "Beban")
        .group_by(Transaction.technician_nama)
        .all()
    )

    # 🔹 Gabungan Pendapatan Bersih per Teknisi
    bersih_teknisi = []
    teknisi_set = set([p[0] for p in pendapatan_teknisi] + [b[0] for b in pengeluaran_teknisi])

    for nama in teknisi_set:
        pend = next((x[1] for x in pendapatan_teknisi if x[0] == nama), 0)
        beng = next((x[1] for x in pengeluaran_teknisi if x[0] == nama), 0)
        bersih_teknisi.append((nama, pend - beng))

    return render_template(
        "transactions_modern.html",
        transactions=trans,

        # ringkasan
        total_pendapatan=total_pendapatan,
        total_pengeluaran=total_pengeluaran,
        saldo=saldo,
        total_internal=total_internal,
        total_eksternal=total_eksternal,

        # tambahan baru
        pendapatan_teknisi=pendapatan_teknisi,
        pengeluaran_teknisi=pengeluaran_teknisi,
        bersih_teknisi=bersih_teknisi,

        tahun=tahun, bulan=bulan, show_all=show_all,
        start_date=start_date, end_date=end_date
    )



@app.route('/transaction/new', methods=['GET','POST'])
@login_required
def transaction_new():
    if request.method == 'POST':
        teknisi_id = request.form.get('technician_id')
        teknisi = None
        if teknisi_id:
            teknisi = Technician.query.get(int(teknisi_id))

        t = Transaction(
            tanggal=datetime.strptime(request.form.get('tanggal'), "%Y-%m-%d").date(),
            kategori=request.form.get('kategori'),   # Pendapatan / Beban
            jenis=request.form.get('jenis'),
            deskripsi=request.form.get('deskripsi'),
            jumlah=float(request.form.get('jumlah') or 0),
            technician_id=teknisi.id if teknisi else None,
            technician_nama=teknisi.nama if teknisi else None,
            team=teknisi.team if teknisi else None
        )

        db.session.add(t)
        db.session.commit()
        flash("✅ Transaksi ditambahkan", "success")
        return redirect(url_for('transactions'))

    # Ambil daftar teknisi untuk dropdown di form
    teknisi_all = Technician.query.all()
    return render_template("transaction_form_modern.html", teknisi_all=teknisi_all)


@app.route('/transaction/<int:id>/delete', methods=['POST'])
@login_required
def transaction_delete(id):
    t = Transaction.query.get_or_404(id)
    db.session.delete(t)
    db.session.commit()
    flash("🗑️ Transaksi dihapus", "success")
    return redirect(url_for('transactions'))

@app.route('/transactions/bulk_delete', methods=['POST'])
@login_required
def transaction_bulk_delete():
    ids = request.form.getlist('selected_ids')
    if ids:
        for id in ids:
            tx = Transaction.query.get(int(id))
            if tx:
                db.session.delete(tx)
        db.session.commit()
        flash(f"{len(ids)} transaksi berhasil dihapus.", "success")
    else:
        flash("Tidak ada transaksi yang dipilih.", "warning")
    return redirect(url_for('transactions'))

@app.route("/cleanup/zero", methods=["POST"])
@login_required
def cleanup_zero():
    # 🔹 Hapus service dengan harga_satuan = 0 atau subtotal = 0
    zero_services = Service.query.filter(
        (Service.harga_satuan == 0) |
        ((Service.harga_satuan * Service.jumlah) == 0)
    ).all()
    for s in zero_services:
        db.session.delete(s)

    # 🔹 Hapus transaksi dengan jumlah = 0
    zero_transactions = Transaction.query.filter(Transaction.jumlah == 0).all()
    for t in zero_transactions:
        db.session.delete(t)

    db.session.commit()

    flash(
        f"🧹 Dihapus {len(zero_services)} service & {len(zero_transactions)} transaksi bernilai 0",
        "success"
    )
    return redirect(url_for("transactions"))


@app.route('/reports/laba_rugi')
@login_required
def laba_rugi():
    tahun = request.args.get('tahun', datetime.today().year, type=int)
    bulan = request.args.get('bulan', datetime.today().month, type=int)
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')

    query = Transaction.query

    if start_date and end_date:
        try:
            sd = datetime.strptime(start_date, "%Y-%m-%d").date()
            ed = datetime.strptime(end_date, "%Y-%m-%d").date()
            query = query.filter(Transaction.tanggal.between(sd, ed))
        except ValueError:
            flash("Format tanggal tidak valid", "warning")
    else:
        # default bulan ini
        query = query.filter(
            extract('year', Transaction.tanggal) == tahun,
            extract('month', Transaction.tanggal) == bulan
        )

    trans = query.all()

    pendapatan = sum(t.jumlah for t in trans if t.kategori == "Pendapatan")
    beban = sum(t.jumlah for t in trans if t.kategori == "Beban")
    laba_bersih = pendapatan - beban

    bulan_names = [
        "Januari","Februari","Maret","April","Mei","Juni",
        "Juli","Agustus","September","Oktober","November","Desember"
    ]

    return render_template(
        "laba_rugi.html",
        tahun=tahun, bulan=bulan,
        pendapatan=pendapatan,
        beban=beban,
        laba_bersih=laba_bersih,
        trans=trans,
        start_date=start_date,
        end_date=end_date,
        format_date=format_date,
        current_year=datetime.today().year,
        current_month=datetime.today().month,
        bulan_names=bulan_names   # ✅ kirim ke template
    )

def format_date(date_str):
    return datetime.strptime(date_str, "%Y-%m-%d").strftime("%d-%m-%Y")

##claim

@app.route('/claim/<int:unit_id>/<int:service_id>', methods=['GET', 'POST'])
@login_required
def claim_service(unit_id, service_id):
    unit = ACUnit.query.get_or_404(unit_id)
    service = Service.query.get_or_404(service_id)

    # 🔒 validasi: klaim hanya bisa max 30 hari dari tanggal service
    batas_claim = service.tanggal + timedelta(days=30)
    if datetime.utcnow().date() > batas_claim:
        flash("❌ Masa klaim sudah lewat (lebih dari 30 hari).", "danger")
        return redirect(url_for('units', cid=unit.customer_id))

    if request.method == 'POST':
        deskripsi = request.form.get("deskripsi")
        foto_file = request.files.get("foto")

        foto_relpath = None
        if foto_file and foto_file.filename.strip():
            # folder upload (absolute) → static/uploads/claims
            upload_dir = os.path.join(current_app.root_path, "static", "uploads", "claims")
            os.makedirs(upload_dir, exist_ok=True)

            # buat nama file aman + timestamp agar unik
            filename = secure_filename(foto_file.filename)
            timestamp = datetime.utcnow().strftime("%Y%m%d%H%M%S")
            filename = f"{timestamp}_{filename}"

            # simpan file
            save_path = os.path.join(upload_dir, filename)
            foto_file.save(save_path)

            # path relatif untuk DB (tanpa "static/", selalu pakai "/")
            foto_relpath = f"uploads/claims/{filename}".replace("\\", "/")

        # simpan klaim ke DB
        new_claim = Claim(
            customer_id=unit.customer_id,
            unit_id=unit.id,
            technician_id=service.technician_id,
            deskripsi=deskripsi,
            foto=foto_relpath,   # contoh: uploads/claims/20250920053500_test.jpg
            status="Pending"
        )
        db.session.add(new_claim)
        db.session.commit()

        flash("✅ Klaim berhasil dicatat.", "success")
        return redirect(url_for('units', cid=unit.customer_id))

    return render_template("claim_form.html", unit=unit, service=service)


@app.route('/claims/<int:customer_id>')
@login_required
def claim_list(customer_id):
    customer = Customer.query.get_or_404(customer_id)
    claims = Claim.query.filter_by(customer_id=customer.id).order_by(Claim.tanggal.desc()).all()
    return render_template("Claim/list.html", claims=claims, customer=customer)

@app.route('/claims')
@login_required
def claim_list_all():
    claims = Claim.query.order_by(Claim.tanggal.desc()).all()
    return render_template('Claim/list.html', claims=claims, customer=None)


@app.route('/claim/edit/<int:claim_id>', methods=['GET', 'POST'])
@login_required
def claim_edit(claim_id):
    claim = Claim.query.get_or_404(claim_id)
    if request.method == 'POST':
        claim.deskripsi = request.form.get('deskripsi')
        claim.status = request.form.get('status')
        db.session.commit()
        flash('Data klaim berhasil diperbarui.', 'success')
        return redirect(url_for('claim_list', customer_id=claim.customer_id))
    
    return render_template('Claim/edit.html', claim=claim)



# -----------------------------
# Export Laba Rugi ke Excel
# -----------------------------
@app.route('/reports/laba_rugi/excel')
@login_required
def laba_rugi_excel():
    tahun = request.args.get('tahun', datetime.today().year, type=int)
    bulan = request.args.get('bulan', datetime.today().month, type=int)

    trans = Transaction.query.filter(
        extract('year', Transaction.tanggal) == tahun,
        extract('month', Transaction.tanggal) == bulan
    ).all()

    pendapatan = sum(t.jumlah for t in trans if t.kategori == "Pendapatan")
    beban = sum(t.jumlah for t in trans if t.kategori == "Beban")
    laba_bersih = pendapatan - beban

    wb = Workbook()
    ws = wb.active
    ws.title = "Laba Rugi"

    ws.append(["PERKASA AC"])
    ws.append([f"Laporan Laba Rugi - {bulan}/{tahun}"])
    ws.append([])
    ws.append(["Pendapatan", pendapatan])
    ws.append(["Beban", beban])
    ws.append(["Laba Bersih", laba_bersih])
    ws.append([])
    ws.append(["Tanggal", "Kategori", "Jenis", "Deskripsi", "Jumlah"])

    for t in trans:
        ws.append([t.tanggal.strftime("%d-%m-%Y"), t.kategori, t.jenis, t.deskripsi, t.jumlah])

    # ✅ gunakan io.BytesIO
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"laba_rugi_{bulan}_{tahun}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -----------------------------
# Export Laba Rugi ke PDF
# -----------------------------


@app.route('/reports/laba_rugi/pdf')
@login_required
def laba_rugi_pdf():
    tahun = request.args.get('tahun', datetime.today().year, type=int)
    bulan = request.args.get('bulan', datetime.today().month, type=int)

    trans = Transaction.query.filter(
        extract('year', Transaction.tanggal) == tahun,
        extract('month', Transaction.tanggal) == bulan
    ).all()

    from collections import defaultdict
    pendapatan_dict, beban_dict = defaultdict(int), defaultdict(int)

    for t in trans:
        if t.kategori == "Pendapatan":
            pendapatan_dict[t.jenis] += t.jumlah
        elif t.kategori == "Beban":
            beban_dict[t.jenis] += t.jumlah

    total_pendapatan = sum(pendapatan_dict.values())
    total_beban = sum(beban_dict.values())
    laba_bersih = total_pendapatan - total_beban

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            leftMargin=50, rightMargin=50, topMargin=60, bottomMargin=40)
    styles = getSampleStyleSheet()
    elements = []

    # 🔹 Logo & Judul
    try:
        elements.append(Image("static/logo.png", width=60, height=60))
    except:
        pass
    elements.append(Paragraph("<b>PERKASA AC</b>", styles["Title"]))
    elements.append(Paragraph(f"Laporan Laba Rugi - {bulan}/{tahun}", styles["Heading2"]))
    elements.append(Spacer(1, 20))

    # 🔹 Pendapatan
    elements.append(Paragraph("<b>Pendapatan</b>", styles["Heading3"]))
    data_pendapatan = [["Jenis", "Jumlah"]]
    for jenis, jumlah in pendapatan_dict.items():
        data_pendapatan.append([jenis, f"Rp {jumlah:,.0f}"])
    data_pendapatan.append(
        ["Total Pendapatan", f"Rp {total_pendapatan:,.0f}"]
    )
    table_p = Table(data_pendapatan, colWidths=[300, 150])
    table_p.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN", (1,0), (1,-1), "RIGHT"),
        ("FONTNAME", (0,-1), (-1,-1), "Helvetica-Bold"),
        ("BACKGROUND", (0,-1), (-1,-1), colors.whitesmoke),
    ]))
    elements.append(table_p)
    elements.append(Spacer(1, 20))

    # 🔹 Beban
    elements.append(Paragraph("<b>Beban</b>", styles["Heading3"]))
    data_beban = [["Jenis", "Jumlah"]]
    for jenis, jumlah in beban_dict.items():
        data_beban.append([jenis, f"Rp {jumlah:,.0f}"])
    data_beban.append(
        ["Total Beban", f"Rp {total_beban:,.0f}"]
    )
    table_b = Table(data_beban, colWidths=[300, 150])
    table_b.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN", (1,0), (1,-1), "RIGHT"),
        ("FONTNAME", (0,-1), (-1,-1), "Helvetica-Bold"),
        ("BACKGROUND", (0,-1), (-1,-1), colors.whitesmoke),
    ]))
    elements.append(table_b)
    elements.append(Spacer(1, 20))

    # 🔹 Laba Bersih
    data_laba = [["Laba Bersih", f"Rp {laba_bersih:,.0f}"]]
    table_l = Table(data_laba, colWidths=[300, 150])
    table_l.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("BACKGROUND", (0,0), (-1,-1), colors.lightgreen),
        ("FONTNAME", (0,0), (-1,-1), "Helvetica-Bold"),
        ("ALIGN", (1,0), (1,0), "RIGHT"),
    ]))
    elements.append(table_l)

    doc.build(elements)
    pdf = buffer.getvalue()
    buffer.close()

    response = make_response(pdf)
    response.headers["Content-Disposition"] = f"attachment; filename=laba_rugi_{bulan}_{tahun}.pdf"
    response.headers["Content-Type"] = "application/pdf"
    return response


@app.route('/reports/arus_kas')
@login_required
def arus_kas():
    tahun = request.args.get('tahun', datetime.today().year, type=int)
    bulan = request.args.get('bulan', datetime.today().month, type=int)

    trans = Transaction.query.filter(
        extract('year', Transaction.tanggal) == tahun,
        extract('month', Transaction.tanggal) == bulan
    ).all()

    kas_masuk = sum(t.jumlah for t in trans if t.kategori == "Pendapatan")
    kas_keluar = sum(t.jumlah for t in trans if t.kategori == "Beban")
    saldo = kas_masuk - kas_keluar

    return render_template(
        "arus_kas.html",
        tahun=tahun, bulan=bulan,
        kas_masuk=kas_masuk,
        kas_keluar=kas_keluar,
        saldo=saldo,
        trans=trans
    )




# Customers CRUD

@app.route('/customers')
@login_required
def customers():
    q = request.args.get('q', '').strip()
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    team = request.args.get('team')  # ✅ filter team
    
    UnitAlias = aliased(ACUnit)
    ServiceAlias = aliased(Service)

    base_query = (
    db.session.query(
        Customer,
        func.coalesce(func.sum(ACUnit.jumlah_unit), 0).label("total_units"),
        func.coalesce(func.count(Service.id), 0).label("total_services"),
        func.max(Service.tanggal).label("last_service_date")   # 🔹 ambil tanggal terakhir service
    )
    .outerjoin(ACUnit, Customer.id == ACUnit.customer_id)
    .outerjoin(Service, ACUnit.id == Service.unit_id)
    .group_by(Customer.id)
    .order_by(func.max(Service.tanggal).desc().nullslast())   # 🔹 urutkan berdasarkan tanggal service terbaru
)


    if q:
        base_query = base_query.filter(
            or_(
                Customer.nama.ilike(f"%{q}%"),
                Customer.nomor_wa.ilike(f"%{q}%")
            )
        )

    if team:
        base_query = base_query.filter(Customer.team == team)  # ✅ filter by team

    if start_date or end_date:
        base_query = (
            base_query.join(UnitAlias, Customer.id == UnitAlias.customer_id)
                      .join(ServiceAlias, UnitAlias.id == ServiceAlias.unit_id)
        )
        if start_date:
            try:
                sd = datetime.strptime(start_date, "%Y-%m-%d").date()
                base_query = base_query.filter(ServiceAlias.tanggal >= sd)
            except ValueError:
                pass
        if end_date:
            try:
                ed = datetime.strptime(end_date, "%Y-%m-%d").date()
                base_query = base_query.filter(ServiceAlias.tanggal <= ed)
            except ValueError:
                pass

    data = base_query.all()

    return render_template(
        'customers_modern.html',
        customers=data,
        q=q,
        start_date=start_date,
        end_date=end_date,
        team=team  # ✅ lempar ke template
    )




@app.route('/customers/new', methods=['GET', 'POST'])
@login_required
def customers_new():
    if request.method == 'POST':
        nama = request.form['nama']
        nomor_wa_raw = request.form.get('nomor_wa')
        alamat = request.form.get('alamat')
        team = request.form.get('team')

        nomor_wa = None

        # ✅ hanya proses jika diisi
        if nomor_wa_raw:
            nomor_wa = normalize_wa(nomor_wa_raw)

            # cek duplikat hanya jika ada nomor
            if Customer.query.filter_by(nomor_wa=nomor_wa).first():
                flash('❌ Nomor WA sudah terdaftar, gunakan nomor lain.', 'danger')
                return redirect(url_for('customers_new'))

        try:
            c = Customer(
                nama=nama,
                nomor_wa=nomor_wa,
                alamat=alamat,
                team=team
            )
            db.session.add(c)
            db.session.commit()
            flash('✅ Customer ditambahkan', 'success')
            return redirect(url_for('customers'))

        except ValueError as e:
            db.session.rollback()
            flash(str(e), 'danger')
            return redirect(url_for('customers_new'))

    return render_template('customer_form_modern.html', customer=None)



@app.route('/customers/<int:id>/edit', methods=['GET','POST'])
@login_required
def customers_edit(id):
    c = Customer.query.get_or_404(id)
    if request.method == 'POST':
        nama = request.form['nama']
        nomor_wa_raw = request.form.get('nomor_wa')
        alamat = request.form.get('alamat')
        team = request.form.get('team')

        # ✅ normalisasi nomor WA
        nomor_wa = normalize_wa(nomor_wa_raw)

        # ✅ cek duplikat nomor WA (exclude dirinya sendiri)
        existing = Customer.query.filter(
            Customer.nomor_wa == nomor_wa,
            Customer.id != id
        ).first()
        if existing:
            flash('❌ Nomor WA sudah dipakai customer lain.', 'danger')
            return redirect(url_for('customers_edit', id=id))

        try:
            # ✅ simpan perubahan
            c.nama = nama
            c.nomor_wa = nomor_wa
            c.alamat = alamat
            c.team = team
            db.session.commit()
            flash('✅ Customer disimpan', 'success')
            return redirect(url_for('customers'))
        except ValueError as e:
            db.session.rollback()
            flash(str(e), 'danger')
            return redirect(url_for('customers_edit', id=id))

    return render_template('customer_form_modern.html', customer=c)

@app.route('/customers/<int:id>/delete', methods=['POST'])
@login_required
def customers_delete(id):
    c = Customer.query.get_or_404(id)
    db.session.delete(c); db.session.commit()
    flash('Customer dihapus', 'success')
    return redirect(url_for('customers'))

from sqlalchemy.orm import joinedload
from sqlalchemy import func

@app.route('/customers/<int:cid>/units')
@login_required
def units(cid):
    c = Customer.query.get_or_404(cid)

    # ✅ Mode edit service (from URL)
    edit_service_id = request.args.get("edit_service")
    edit_service = Service.query.get(edit_service_id) if edit_service_id else None

    merk = request.args.get('merk', '').strip()
    pk = request.args.get('pk', '').strip()
    tanggal = request.args.get('tanggal', '').strip()

    units_query = (
        ACUnit.query.filter_by(customer_id=cid)
        .options(db.joinedload(ACUnit.services).joinedload(Service.technician))
    )

    if merk:
        units_query = units_query.filter(ACUnit.merk.ilike(f"%{merk}%"))
    if pk:
        units_query = units_query.filter(ACUnit.pk == pk)
    if tanggal:
        try:
            tgl = datetime.strptime(tanggal, "%Y-%m-%d").date()
            units_query = units_query.join(Service).filter(Service.tanggal == tgl)
        except ValueError:
            pass

    units = units_query.all()

    for u in units:
        u.services.sort(key=lambda s: s.tanggal or datetime.min.date(), reverse=True)

    current_date = datetime.utcnow().date()

    return render_template(
        'units_modern.html',
        customer=c,
        units=units,
        merk=merk,
        pk=pk,
        tanggal=tanggal,
        current_date=current_date,
        edit_service=edit_service  # ✅ kirim ke template
    )




@app.route('/customers/<int:cid>/units/new', methods=['POST'])
@login_required
def units_new(cid):
    ruangan = request.form.get("ruangan")
    c = Customer.query.get_or_404(cid)
    u = ACUnit(customer=c,
               merk=request.form.get('merk'),
               pk=request.form.get('pk'),
               jumlah_unit=int(request.form.get('jumlah_unit') or 1),
               ruangan=ruangan)
    db.session.add(u); db.session.commit()
    flash('Unit ditambahkan', 'success')
    return redirect(url_for('units', cid=cid))

# =======================
# Edit Unit
# =======================
@app.route('/units/<int:uid>/edit', methods=['GET', 'POST'])
@login_required
def unit_edit(uid):
    unit = ACUnit.query.get_or_404(uid)
    customer = Customer.query.get_or_404(unit.customer_id)

    if request.method == 'POST':
        unit.ruangan = request.form.get("ruangan")
        unit.merk = request.form.get("merk")
        unit.pk = request.form.get("pk")
        unit.jumlah_unit = int(request.form.get("jumlah_unit") or 1)

        db.session.commit()
        flash("✅ Unit berhasil diperbarui", "success")
        return redirect(url_for('units', cid=unit.customer_id))

    return render_template("unit_edit_form.html", unit=unit, customer=customer)


@app.route('/units/<int:uid>/delete', methods=['POST'])
@login_required
def units_delete(uid):
    u = ACUnit.query.get_or_404(uid)
    cid = u.customer_id
    db.session.delete(u); db.session.commit()
    flash('Unit dihapus', 'success')
    return redirect(url_for('units', cid=cid))


@app.route("/unit/history/all/<int:cid>")
@login_required
def unit_history_all(cid):
    customer = Customer.query.get_or_404(cid)
    units = ACUnit.query.filter_by(customer_id=cid).all()

    all_records = []

    for u in units:
        # ✅ Masukkan semua service
        for s in u.services:
            all_records.append({
                "tanggal": s.tanggal,
                "unit": u,
                "jenis": s.jenis_service,
                "teknisi": s.technician.nama if s.technician else "-",
                "team": s.technician.team if s.technician else "-",
                "deskripsi": s.deskripsi or "",
                "tipe": "Service",
                "status": None,
                "foto": None
            })

        # ✅ Masukkan semua klaim
        for c in u.claims:
            all_records.append({
                "tanggal": c.tanggal.date() if isinstance(c.tanggal, datetime) else c.tanggal,
                "unit": u,
                "jenis": "Klaim Service",
                "teknisi": "-",
                "team": "-",
                "deskripsi": c.deskripsi or "",
                "tipe": "Claim",
                "status": c.status,   # pending / on progress / selesai
                "foto": c.foto
            })

    # ✅ Helper untuk seragamkan tanggal jadi datetime
    def to_datetime_safe(tgl):
        if tgl is None:
            return datetime.min
        if isinstance(tgl, date) and not isinstance(tgl, datetime):
            return datetime.combine(tgl, datetime.min.time())
        return tgl

    # ✅ Urutkan berdasarkan tanggal terbaru
    all_records = sorted(
        all_records,
        key=lambda x: to_datetime_safe(x["tanggal"]),
        reverse=True
    )

    return render_template(
        "unit_history_all.html",
        customer=customer,
        all_services=all_records
    )


@app.route("/units/delete_all/<int:cid>", methods=["POST"])
@login_required
def units_delete_all(cid):
    units = ACUnit.query.filter_by(customer_id=cid).all()  # 🔥 fix
    for u in units:
        Service.query.filter_by(unit_id=u.id).delete()
        db.session.delete(u)
    db.session.commit()
    flash("Semua unit & servicenya berhasil dihapus", "success")
    return redirect(url_for("units", cid=cid))


# Services

@app.route('/service/new/<int:unit_id>', methods=['GET','POST'])
@login_required
def service_new(unit_id):
    unit = ACUnit.query.get_or_404(unit_id)
    service_types = ServiceType.query.all()
    packages = Package.query.all()
    technicians = Technician.query.all()

    if request.method == 'POST':
        idx = 0
        teknisi_id = request.form.get('technician_id')
        teknisi = Technician.query.get(int(teknisi_id)) if teknisi_id else None
        pembayaran = request.form.get('pembayaran')  # 🔹 ambil status pembayaran

        new_services = []

        while True:
            jenis = request.form.get(f'jenis_service_{idx}')
            if not jenis:
                break

            jumlah = int(request.form.get(f'jumlah_{idx}') or 1)
            harga_satuan = float(request.form.get(f'harga_{idx}') or 0)
            deskripsi = request.form.get(f'deskripsi_{idx}') or ""

            # 🔹 tanggal service
            tanggal_service = datetime.strptime(
                request.form.get('tanggal'), '%Y-%m-%d'
            ).date() if request.form.get('tanggal') else date.today()

            # 🔹 otomatis tanggal cuci berikutnya
            tanggal_cuci_berikutnya = tanggal_service + relativedelta(months=3)

            s = Service(
                unit_id=unit.id,
                tanggal=tanggal_service,
                tanggal_cuci_berikutnya=tanggal_cuci_berikutnya,
                jenis_service=jenis,
                paket_cuci=request.form.get(f'paket_cuci_{idx}'),
                technician_id=teknisi.id if teknisi else None,
                teknisi=teknisi.nama if teknisi else None,
                team=teknisi.team if teknisi else None,
                harga_satuan=harga_satuan,
                jumlah=jumlah,
                pembayaran=pembayaran,
                deskripsi=deskripsi
            )

            db.session.add(s)
            db.session.flush()  # agar dapat ID service sebelum commit
            new_services.append(s)

            # 🔹 Jika pembayaran "Lunas", langsung buat transaksi otomatis
            if pembayaran and pembayaran.lower() == "lunas":
                subtotal = harga_satuan * jumlah
                pendapatan_tipe = "Internal" if s.team == "Internal" else "Eksternal"
                base_desc = f"Service {unit.merk} ({unit.pk}) - {unit.customer.nama}"
                if deskripsi:
                    base_desc += f" | {deskripsi}"

                t = Transaction(
                    service_id=s.id,
                    tanggal=tanggal_service,
                    kategori="Pendapatan",
                    jenis=jenis,
                    deskripsi=base_desc,
                    jumlah=subtotal,
                    pendapatan_tipe=pendapatan_tipe,
                    technician_id=s.technician_id,
                    technician_nama=s.teknisi,
                    team=s.team
                )
                db.session.add(t)

            idx += 1

        db.session.commit()
        flash("✅ Service berhasil ditambahkan", "success")
        return redirect(url_for('units', cid=unit.customer_id))

    # GET default
    default_tcb = date.today() + relativedelta(months=3)
    return render_template(
        'service_form_modern.html',
        unit=unit,
        service=None,
        service_types=service_types,
        packages=packages,
        technicians=technicians,
        default_tcb=default_tcb
    )


# =======================
# SERVICE EDIT
# =======================
@app.route('/service/<int:id>/edit', methods=['GET', 'POST'])
@login_required
def service_edit(id):
    s = Service.query.get_or_404(id)

    # 🔹 Data tambahan untuk form
    technicians = Technician.query.all()
    service_types = ServiceType.query.all()
    packages = Package.query.all()
    unit = s.ac_unit  # ✅ FIXED: ganti dari s.unit ke s.ac_unit
    default_tcb = s.tanggal_cuci_berikutnya or (s.tanggal + relativedelta(months=3))

    # ==========================
    # 1️⃣ JIKA POST → SIMPAN PERUBAHAN
    # ==========================
    if request.method == 'POST':
        teknisi_id = request.form.get('technician_id')
        teknisi = Technician.query.get(int(teknisi_id)) if teknisi_id else None

        tanggal_raw = request.form.get('tanggal')
        s.tanggal = datetime.strptime(tanggal_raw, '%Y-%m-%d').date() if tanggal_raw else date.today()

        tcb_raw = request.form.get('tanggal_cuci_berikutnya')
        s.tanggal_cuci_berikutnya = (
            datetime.strptime(tcb_raw, '%Y-%m-%d').date()
            if tcb_raw else s.tanggal + relativedelta(months=3)
        )

        s.jenis_service = request.form.get('jenis_service') or s.jenis_service
        s.paket_cuci = request.form.get('paket_cuci') or None
        s.harga_satuan = float(request.form.get('harga_satuan') or 0)
        s.jumlah = int(request.form.get('jumlah') or 1)
        s.deskripsi = request.form.get('deskripsi')
        s.pembayaran = request.form.get('pembayaran')

        # 🟩 Teknisi
        s.technician_id = teknisi.id if teknisi else None
        s.teknisi = teknisi.nama if teknisi else None
        s.team = teknisi.team if teknisi else None

        subtotal = s.harga_satuan * s.jumlah
        pendapatan_tipe = "Internal" if s.team == "Internal" else "Eksternal"

        db.session.refresh(s.ac_unit)
        db.session.refresh(s.ac_unit.customer)

        base_desc = f"Service {s.ac_unit.merk} ({s.ac_unit.pk}) - {s.ac_unit.customer.nama}"
        if s.deskripsi:
            base_desc += f" | {s.deskripsi}"

        existing_tx = Transaction.query.filter_by(service_id=s.id).first()

        # 🔹 Update transaksi jika sudah lunas
        if s.pembayaran and s.pembayaran.lower() == "lunas":
            if existing_tx:
                existing_tx.tanggal = s.tanggal
                existing_tx.jenis = s.jenis_service or "-"
                existing_tx.deskripsi = base_desc
                existing_tx.jumlah = subtotal
                existing_tx.pendapatan_tipe = pendapatan_tipe
                existing_tx.technician_id = s.technician_id
                existing_tx.technician_nama = s.teknisi
                existing_tx.team = s.team
            else:
                new_tx = Transaction(
                    service_id=s.id,
                    tanggal=s.tanggal,
                    kategori="Pendapatan",
                    jenis=s.jenis_service or "-",
                    deskripsi=base_desc,
                    jumlah=subtotal,
                    pendapatan_tipe=pendapatan_tipe,
                    technician_id=s.technician_id,
                    technician_nama=s.teknisi,
                    team=s.team
                )
                db.session.add(new_tx)
        else:
            if existing_tx:
                db.session.delete(existing_tx)

        db.session.commit()
        flash('✅ Service berhasil diperbarui', 'success')
        return redirect(url_for('units', cid=s.ac_unit.customer_id, edit_service=s.id))  # ✅ FIXED

    # ==========================
    # 2️⃣ JIKA GET → TAMPILKAN FORM EDIT
    # ==========================
    return render_template(
        'service_form_modern.html',
        service=s,
        unit=unit,
        technicians=technicians,
        service_types=service_types,
        packages=packages,
        default_tcb=default_tcb
    )


# ✅ route delete service (dipakai di template)
@app.route('/service/<int:id>/delete', methods=['POST'])
@login_required
def service_delete(id):
    s = Service.query.get_or_404(id)
    cid = s.ac_unit.customer_id  # ✅ FIXED

    # 🔹 Hapus transaksi terkait
    Transaction.query.filter_by(service_id=s.id).delete()

    db.session.delete(s)
    db.session.commit()

    flash('✅ Service dan transaksi terkait berhasil dihapus', 'success')
    return redirect(url_for('units', cid=cid))


from sqlalchemy.orm import joinedload


@app.route('/unpaid', endpoint='unpaid_list')
@login_required
def unpaid_list():
    bulan = request.args.get('bulan', type=int)
    tahun = request.args.get('tahun', type=int)
    technician_id = request.args.get('technician_id', type=int)

    # Query dasar
    query = (
        Service.query
        .filter(Service.pembayaran == 'Belum Lunas')
        .options(
            joinedload(Service.ac_unit).joinedload(ACUnit.customer),
            joinedload(Service.technician)
        )
    )

    # Filter
    if bulan:
        query = query.filter(extract('month', Service.tanggal) == bulan)
    if tahun:
        query = query.filter(extract('year', Service.tanggal) == tahun)
    if technician_id:
        query = query.filter(Service.technician_id == technician_id)

    unpaid_services = query.order_by(Service.tanggal.desc()).all()
    technicians = Technician.query.order_by(Technician.nama.asc()).all()

    # ✅ Hitung total belum lunas dari field 'harga_total' / 'harga' / 'harga_satuan * jumlah'
    total_belum_lunas = 0
    for s in unpaid_services:
        if getattr(s, 'harga_total', None):
            total_belum_lunas += s.harga_total
        elif getattr(s, 'harga', None):
            total_belum_lunas += s.harga
        elif getattr(s, 'total', None):
            total_belum_lunas += s.total
        else:
            total_belum_lunas += (getattr(s, 'harga_satuan', 0) or 0) * (getattr(s, 'jumlah', 1) or 1)

    return render_template(
        'unpaid_list.html',
        unpaid_services=unpaid_services,
        technicians=technicians,
        bulan=bulan,
        tahun=tahun,
        technician_id=technician_id,
        total_belum_lunas=total_belum_lunas  # kirim ke template
    )

### INVOICE
# 🔹 Halaman Detail Invoice



@app.route("/invoice/<int:id>/detail")
@login_required
def invoice_detail(id):
    invoice = Invoice.query.get_or_404(id)

    # --- Fungsi aman convert angka
    def to_float(val, default=0):
        if not val:
            return default
        return float(val.replace(",", "."))
    
    # ----- Ambil input -----
    ppn = to_float(request.args.get("ppn", "0"))
    pph_raw = request.args.get("pph", "0").replace(",", ".")
    admin_fee = to_float(request.args.get("admin_fee", "0"))
    dp = to_float(request.args.get("dp", "0"))
    
    # ----- Hitung subtotal -----
    subtotal = 0
    for item in invoice.items:
        qty = item.service.jumlah if (item.service and item.service.jumlah) else 1
        harga = (
            item.harga
            or (item.service.harga_satuan if item.service and item.service.harga_satuan else 0)
        )
        subtotal += (qty or 1) * (harga or 0)

    # ----- Hitung PPH (persen / nominal) -----
    pph_amt = 0
    pph_percent = 0

    if "%" in pph_raw:
        # Input dalam persen → contoh: "2%"
        pph_percent = float(pph_raw.replace("%", "").strip())
        pph_amt = subtotal * (pph_percent / 100)
    else:
        # Bisa persen atau nominal
        pph_val = to_float(pph_raw)

        if pph_val >= 1000:
            # Nominal Rupiah
            pph_amt = pph_val
            pph_percent = (pph_amt / subtotal * 100) if subtotal > 0 else 0
        else:
            # Persen
            pph_percent = pph_val
            pph_amt = subtotal * (pph_percent / 100)

    # ----- Hitung PPN -----
    ppn_amt = subtotal * (ppn / 100)

    # ----- Total -----
    total = subtotal + ppn_amt - pph_amt - admin_fee
    sisa = total - dp
    bank_accounts = BankAccount.query.all()

    return render_template(
        "invoice_detail.html",
        invoice=invoice,
        subtotal=subtotal,
        ppn=ppn,
        ppn_amt=ppn_amt,
        pph=pph_percent,
        pph_amt=pph_amt,
        admin_fee=admin_fee,
        total=total,
        dp=dp,
        sisa=sisa,
        bank_accounts=bank_accounts,
    )

@app.route('/bank_account', endpoint='bank_account_list')
@login_required
def bank_account_list():
    banks = BankAccount.query.order_by(BankAccount.id.desc()).all()
    return render_template('bank_account_list.html', banks=banks)


@app.route('/bank_account/add', methods=['GET', 'POST'], endpoint='bank_account_add')
@login_required
def bank_account_add():
    if request.method == 'POST':
        bank_name = request.form.get('bank_name')
        account_name = request.form.get('account_name')
        account_number = request.form.get('account_number')

        new_bank = BankAccount(
            bank_name=bank_name,
            account_name=account_name,
            account_number=account_number
        )
        db.session.add(new_bank)
        db.session.commit()

        flash("Rekening bank berhasil ditambahkan!", "success")
        return redirect(url_for('bank_account_list'))

    return render_template('bank_account_add.html')


@app.route('/bank_account/<int:id>/edit', methods=['GET', 'POST'], endpoint='bank_account_edit')
@login_required
def bank_account_edit(id):
    bank = BankAccount.query.get_or_404(id)

    if request.method == 'POST':
        bank.bank_name = request.form.get('bank_name')
        bank.account_name = request.form.get('account_name')
        bank.account_number = request.form.get('account_number')

        db.session.commit()

        flash("Rekening bank berhasil diperbarui!", "success")
        return redirect(url_for('bank_account_list'))

    return render_template('bank_account_edit.html', bank=bank)


@app.route('/bank_account/<int:id>/delete', methods=['POST'], endpoint='bank_account_delete')
@login_required
def bank_account_delete(id):
    bank = BankAccount.query.get_or_404(id)
    db.session.delete(bank)
    db.session.commit()
    flash("Rekening bank berhasil dihapus!", "success")
    return redirect(url_for('bank_account_list'))




from flask import render_template, request, Response, current_app


@app.route("/invoice/<int:id>/pdf")
@login_required
def invoice_pdf(id):

    invoice = Invoice.query.get_or_404(id)

    # ==========================
    # TEMPLATE SELECTION
    # ==========================
    template_map = {
        "classic": "invoice_templates/classic.html",
        "modern": "invoice_templates/modern.html",
        "corporate": "invoice_templates/corporate.html",
        "industrial": "invoice_templates/industrial.html",
    }

    selected_template = template_map.get(
        invoice.template,
        "invoice_templates/classic.html"
    )

    # ==========================
    # HELPER FLOAT AMAN
    # ==========================
    def to_float(val, default=0.0):
        try:
            if val is None:
                return default
            if isinstance(val, (int, float)):
                return float(val)
            val = str(val).strip()
            if not val:
                return default
            val = val.replace(".", "").replace(",", ".")
            return float(val)
        except Exception:
            return default

    # ==========================
    # PARAMETER REQUEST
    # ==========================
    spo_param = request.args.get("spo", "").strip()
    spo = spo_param if spo_param else (invoice.spo or "-")

    ppn_percent = to_float(request.args.get("ppn", 0))
    pph_input = to_float(request.args.get("pph", 0))
    admin_fee = to_float(request.args.get("admin_fee", 0))
    dp = to_float(request.args.get("dp", 0))

    # ==========================
    # GROUPING RUANGAN
    # ==========================
    from collections import defaultdict
    grouped = defaultdict(list)

    for inv_item in invoice.items:

        service = inv_item.service
        if not service or not service.ac_unit:
            continue

        ruangan = service.ac_unit.ruangan or "-"

        pk = service.ac_unit.pk or "-"
        inverter_flag = ""

        if hasattr(service.ac_unit, "inverter") and service.ac_unit.inverter:
            inverter_flag = " (INVERTER)"

        # bersihkan strip dari nama service
        jenis = (service.jenis_service or "Service").replace("-", "").strip()

        deskripsi = inv_item.deskripsi or service.jenis_service or "Service"

        qty = service.jumlah or 1
        harga = inv_item.harga or service.harga_satuan or 0
        subtotal_item = qty * harga

        grouped[ruangan].append({
            "deskripsi": deskripsi,
            "qty": qty,
            "harga": harga,
            "subtotal": subtotal_item
        })

    # ==========================
    # HITUNG TOTAL (SETELAH LOOP)
    # ==========================
    subtotal = sum(
        item["subtotal"]
        for ruang in grouped.values()
        for item in ruang
    )

    # ==========================
    # PAJAK & TOTAL
    # ==========================
    ppn_amt = subtotal * (ppn_percent / 100)

    if pph_input >= 1000:
        pph_amt = pph_input
        pph_percent = (pph_amt / subtotal * 100) if subtotal > 0 else 0
    else:
        pph_percent = pph_input
        pph_amt = subtotal * (pph_percent / 100)

    total = max(0, subtotal + ppn_amt - pph_amt - admin_fee)
    sisa = max(0, total - dp)

    # ==========================
    # RENDER HTML
    # ==========================
    html = render_template(
        selected_template,
        invoice=invoice,
        grouped_items=grouped,
        subtotal=subtotal,
        ppn=ppn_percent,
        ppn_amt=ppn_amt,
        pph=pph_percent,
        pph_amt=pph_amt,
        admin_fee=admin_fee,
        total=total,
        dp=dp,
        sisa=sisa,
        bank=invoice.bank_account,
        spo=spo,
        logo_path=os.path.join(current_app.root_path, "static/logo.png")
    )

    # ==========================
    # GENERATE PDF
    # ==========================
    pdf = HTML(
        string=html,
        base_url=current_app.root_path
    ).write_pdf()

    return Response(
        pdf,
        mimetype="application/pdf",
        headers={
            "Content-Disposition": f"inline; filename=invoice_{invoice.id}.pdf"
        }
    )



from datetime import datetime, date
# =============================
# FUNCTION TERBILANG
# =============================
def terbilang(n):
    angka = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam",
             "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"]

    n = int(n)

    if n < 12:
        return angka[n]
    elif n < 20:
        return terbilang(n - 10) + " Belas"
    elif n < 100:
        return terbilang(n // 10) + " Puluh " + terbilang(n % 10)
    elif n < 200:
        return "Seratus " + terbilang(n - 100)
    elif n < 1000:
        return terbilang(n // 100) + " Ratus " + terbilang(n % 100)
    elif n < 2000:
        return "Seribu " + terbilang(n - 1000)
    elif n < 1000000:
        return terbilang(n // 1000) + " Ribu " + terbilang(n % 1000)
    elif n < 1000000000:
        return terbilang(n // 1000000) + " Juta " + terbilang(n % 1000000)
    else:
        return "Jumlah terlalu besar"
    
@app.route("/invoice/<int:id>/kwitansi", methods=["GET", "POST"])
@login_required
def kwitansi_create(id):
    invoice = Invoice.query.get_or_404(id)

    total = sum(
        (item.service.jumlah or 1) *
        (item.harga or item.service.harga_satuan)
        for item in invoice.items
    )

    if request.method == "POST":

        tanggal_input = request.form.get("tanggal")

        if tanggal_input:
            tanggal_obj = datetime.strptime(tanggal_input, "%Y-%m-%d").date()
        else:
            tanggal_obj = date.today()

        kwitansi = Kwitansi(
            invoice_id=invoice.id,
            nomor_kwitansi=request.form.get("nomor_kwitansi"),
            jumlah=total,
            terbilang=terbilang(total),
            tanggal=tanggal_obj,
        )

        db.session.add(kwitansi)
        db.session.commit()

        return redirect(url_for("kwitansi_pdf", id=kwitansi.id))

    # GET request (tampilkan form)
    return render_template(
        "kwitansi_form.html",
        invoice=invoice,
        total=total,
        total_terbilang=terbilang(total),
        date_today=date.today().strftime("%Y-%m-%d")
    )


@app.route("/kwitansi/<int:id>/pdf")
@login_required
def kwitansi_pdf(id):
    kwitansi = Kwitansi.query.get_or_404(id)

    # INI YANG WAJIB DIGANTI
    invoice = Invoice.query.get_or_404(kwitansi.invoice_id)

    html = render_template(
        "kwitansi_pdf.html",
        invoice=invoice,
        kwitansi=kwitansi,
        today=kwitansi.tanggal
    )

    pdf = pdfkit.from_string(html, False)
    response = make_response(pdf)
    response.headers["Content-Type"] = "application/pdf"
    response.headers["Content-Disposition"] = "inline; filename=kwitansi.pdf"
    return response

@app.route("/invoice/<int:id>/update_admin", methods=["POST"])
@login_required
def update_invoice_admin(id):
    invoice = Invoice.query.get_or_404(id)

    from datetime import datetime

    # ==================
    # TEMPLATE
    # ==================
    template = request.form.get("template")
    if template:
        invoice.template = template

    # ==================
    # TANGGAL
    # ==================
    tanggal = request.form.get("tanggal")
    if tanggal:
        try:
            invoice.tanggal = datetime.strptime(tanggal, "%Y-%m-%d").date()
        except ValueError:
            flash("Format tanggal tidak valid", "danger")
            return redirect(url_for("invoice_detail", id=id))

    # ==================
    # SPO
    # ==================
    invoice.spo = request.form.get("spo")

    # ==================
    # REKENING
    # ==================
    bank_id = request.form.get("bank_account_id")
    invoice.bank_account_id = int(bank_id) if bank_id else None

    db.session.commit()

    flash("Perubahan administrasi berhasil disimpan ✅", "success")
    return redirect(url_for("invoice_detail", id=id))


@app.route('/invoice/<int:id>/update_link', methods=['POST'])
def update_invoice_link(id):
    invoice = Invoice.query.get_or_404(id)
    link = request.form.get('spreadsheet_link')

    if not link:
        flash("❌ Link spreadsheet tidak boleh kosong.", "danger")
        return redirect(url_for('invoice_detail', id=id))

    invoice.spreadsheet_link = link.strip()
    db.session.commit()
    flash("✅ Link spreadsheet berhasil disimpan!", "success")
    return redirect(url_for('invoice_detail', id=id))


@app.route('/invoice/<int:id>/qr_link')
def invoice_qr_link(id):
    invoice = Invoice.query.get_or_404(id)

    if not invoice.spreadsheet_link:
        abort(404, description="Invoice belum memiliki link spreadsheet")

    # Buat QR code dari link spreadsheet
    qr = qrcode.QRCode(box_size=10, border=2)
    qr.add_data(invoice.spreadsheet_link)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")

    # Simpan ke memori
    buf = BytesIO()
    img.save(buf, format='PNG')
    buf.seek(0)

    return send_file(buf, mimetype='image/png')

# =========================
# LIST SERVICE MOTOR
# =========================
@app.route('/service_motor')
@login_required
def service_motor_list():
    services = ServiceMotor.query.order_by(ServiceMotor.tanggal.desc()).all()
    return render_template('service_motor/list.html', services=services)


# =========================
# TAMBAH SERVICE MOTOR
# =========================
@app.route('/service_motor/tambah', methods=['GET', 'POST'])
@login_required
def service_motor_tambah():
    technicians = Technician.query.filter_by(status="Aktif").all()

    if request.method == 'POST':
        tech_id = request.form['technician_id']
        total_biaya = float(request.form['total_biaya'])
        keterangan = request.form.get('keterangan', '')
        tanggal_str = request.form.get('tanggal')

        # Parsing tanggal dari form
        if tanggal_str:
            try:
                tanggal = datetime.strptime(tanggal_str, "%Y-%m-%d").date()
            except ValueError:
                flash("Format tanggal tidak valid", "danger")
                return redirect(url_for('service_motor_tambah'))
        else:
            tanggal = date.today()

        # Pembagian hasil
        persen_teknisi = 50
        persen_perusahaan = 50

        porsi_teknisi = total_biaya * (persen_teknisi / 100)
        porsi_perusahaan = total_biaya - porsi_teknisi

        service = ServiceMotor(
            technician_id=tech_id,
            tanggal=tanggal,
            total_biaya=total_biaya,
            persen_teknisi=persen_teknisi,
            persen_perusahaan=persen_perusahaan,
            porsi_teknisi=porsi_teknisi,
            porsi_perusahaan=porsi_perusahaan,
            keterangan=keterangan,
            status="Belum Dipotong"
        )

        db.session.add(service)
        db.session.commit()
        flash("✅ Service motor berhasil ditambahkan", "success")
        return redirect(url_for('service_motor_list'))

    return render_template('service_motor/form.html', technicians=technicians, date=date)

@app.route('/service_motor/<int:id>/edit_porsi', methods=['GET', 'POST'])
@login_required
def service_motor_edit_porsi(id):
    service = ServiceMotor.query.get_or_404(id)

    # ❌ Tidak boleh edit kalau sudah masuk payroll
    if service.payroll_id:
        flash("❌ Service motor sudah masuk payroll dan tidak bisa diedit", "danger")
        return redirect(url_for('service_motor_list'))

    if request.method == 'POST':
        try:
            persen_perusahaan = float(request.form['persen_perusahaan'])
        except (KeyError, ValueError):
            flash("Input persentase tidak valid", "danger")
            return redirect(request.url)

        if persen_perusahaan < 0 or persen_perusahaan > 100:
            flash("Persentase harus antara 0 - 100", "danger")
            return redirect(request.url)

        # ✅ LOGIKA TERBALIK
        persen_teknisi = 100 - persen_perusahaan

        # Simpan RIWAYAT
        history = ServiceMotorHistory(
            service_motor_id=service.id,
            persen_teknisi=persen_teknisi,
            persen_perusahaan=persen_perusahaan,
            porsi_teknisi=service.porsi_teknisi,
            porsi_perusahaan=service.porsi_perusahaan,
            diubah_oleh=current_user.username
        )
        db.session.add(history)

        # Update data utama
        service.persen_teknisi = persen_teknisi
        service.persen_perusahaan = persen_perusahaan

        service.porsi_teknisi = service.total_biaya * (persen_teknisi / 100)
        service.porsi_perusahaan = service.total_biaya * (persen_perusahaan / 100)

        db.session.commit()
        flash("✅ Porsi berhasil diperbarui (Perusahaan / Teknisi)", "success")
        return redirect(url_for('service_motor_list'))

    return render_template('service_motor/edit_porsi.html', service=service)



@app.route('/service_motor/<int:id>/riwayat')
@login_required
def service_motor_riwayat(id):
    service = ServiceMotor.query.get_or_404(id)
    return render_template(
        'service_motor/riwayat.html',
        service=service,
        histories=service.histories
    )


# =========================
# HAPUS SERVICE MOTOR
# =========================
@app.route('/service_motor/<int:id>/hapus')
@login_required
def service_motor_hapus(id):
    service = ServiceMotor.query.get_or_404(id)
    db.session.delete(service)
    db.session.commit()
    flash("🗑️ Service motor dihapus", "info")
    return redirect(url_for('service_motor_list'))


# =========================
# DETAIL SERVICE PER TEKNISI
# =========================
@app.route('/service_motor/detail/<int:technician_id>')
@login_required
def service_motor_detail(technician_id):
    tech = Technician.query.get_or_404(technician_id)
    services = ServiceMotor.query.filter_by(technician_id=technician_id).order_by(ServiceMotor.tanggal.desc()).all()
    return render_template('service_motor/detail.html', tech=tech, services=services)


# =========================
# TAMBAH SERVICE MOTOR KE PAYROLL
# =========================
@app.route('/service_motor/tambah_ke_payroll', methods=['POST'])
@login_required
def service_motor_tambah_ke_payroll():
    # Ambil id yang dicentang di form
    service_ids = request.form.getlist('service_ids[]')
    if not service_ids:
        flash('Pilih minimal satu service motor untuk ditambahkan ke payroll.', 'warning')
        return redirect(url_for('service_motor_list'))

    # Pastikan hanya ambil yang memang ada dan belum di-assign ke payroll manapun
    services = ServiceMotor.query.filter(
        ServiceMotor.id.in_(service_ids),
        ServiceMotor.payroll_id.is_(None)  # belum terhubung ke payroll
    ).all()

    if not services:
        flash('Data service motor tidak ditemukan atau sudah ditambahkan ke payroll sebelumnya.', 'danger')
        return redirect(url_for('service_motor_list'))

    for s in services:
        s.terpilih = True
        # tetap beri status "Belum Dipotong" sampai payroll di-final
        s.status = 'Belum Dipotong'
        db.session.add(s)

    db.session.commit()
    flash('✅ Service motor berhasil ditandai untuk dimasukkan ke payroll.', 'success')
    return redirect(url_for('service_motor_list'))


# =========================
# DEBUG SERVICE MOTOR
# =========================
@app.route('/debug/service_motor/<int:tech_id>')
def debug_service_motor(tech_id):
    services = ServiceMotor.query.filter_by(technician_id=tech_id).all()
    return {
        "service_motor": [
            {
                "id": s.id,
                "status": s.status,
                "terpilih": s.terpilih,
                "total_biaya": s.total_biaya,
                "keterangan": s.keterangan
            } for s in services
        ]
    }

def hitung_hari_kerja(start_date, end_date):
    """
    Menghitung hari kerja Senin-Sabtu
    Minggu tidak dihitung
    """
    total = 0

    current = start_date

    while current <= end_date:

        # Senin-Sabtu
        if current.weekday() != 6:
            total += 1

        current += timedelta(days=1)

    return total
# ==================================================
# GAJI & TUNJANGAN (PERBAIKAN)
# ==================================================

@app.route('/payroll')
@login_required
def payroll_index():
    periode = request.args.get('periode')

    query = Payroll.query
    if periode:
        query = query.filter_by(periode=periode)

    payrolls = query.order_by(Payroll.periode.desc()).all()

    technicians = Technician.query.filter_by(status="Aktif") \
                                  .order_by(Technician.nama) \
                                  .all()

    return render_template(
        'payroll/index.html',
        payrolls=payrolls,
        technicians=technicians
    )


# =========================
# DETAIL PAYROLL (DIPERBAIKI)
# =========================
@app.route('/payroll/<int:id>')
@login_required
def payroll_detail(id):
    payroll = Payroll.query.get_or_404(id)
    employee = payroll.employee
    tahun, bulan = map(int, payroll.periode.split('-'))

    hari_dalam_bulan = monthrange(tahun, bulan)[1]

    awal_bulan = date(tahun, bulan, 1)
    akhir_bulan = date(tahun, bulan, hari_dalam_bulan)

    aktif_mulai = awal_bulan

    if employee.tanggal_masuk:
        aktif_mulai = max(
            awal_bulan,
            employee.tanggal_masuk
        )

    aktif_sampai = akhir_bulan

    if hasattr(employee, "tanggal_keluar") and employee.tanggal_keluar:
        aktif_sampai = min(
            akhir_bulan,
            employee.tanggal_keluar
        )

    hari_kerja_aktif = hitung_hari_kerja(
        aktif_mulai,
        aktif_sampai
    )

    total_hadir = Attendance.query.filter(
        Attendance.technician_id == employee.id,
        Attendance.tanggal >= aktif_mulai,
        Attendance.tanggal <= aktif_sampai,
        Attendance.status == "Hadir"
    ).count()

    total_izin = Attendance.query.filter(
        Attendance.technician_id == employee.id,
        Attendance.tanggal >= aktif_mulai,
        Attendance.tanggal <= aktif_sampai,
        Attendance.status.in_(["Izin", "Sakit"])
    ).count()

    # gunakan data payroll agar sama dengan perhitungan payroll
    total_tidak_hadir = payroll.total_tidak_hadir or 0

    # ambil items yang sudah nyata ada di payroll (Pendapatan & Potongan)
    items = PayrollItem.query.filter_by(payroll_id=id).all()
    print("=" * 50)
    print("PAYROLL ID :", payroll.id)
    print("TOTAL ITEMS:", len(items))

    for i in items:
        print(
            i.kategori,
            i.deskripsi,
            i.jumlah
        )

    print("=" * 50)

    # Hitung total pendapatan & potongan dari item (PayrollItem sumber kebenaran)
    total_pendapatan = sum(i.jumlah or 0 for i in items if i.kategori == 'Pendapatan')
    total_potongan = sum(i.jumlah or 0 for i in items if i.kategori == 'Potongan')

    # Hitung potongan absen (gunakan attendance yang sudah tersimpan)
    total_tidak_hadir = payroll.total_tidak_hadir or 0

    setting = PayrollSetting.query.filter_by(employee_id=employee.id).first()
    potongan_absen_per_hari = setting.potongan_absen_per_hari if setting else 0
    potongan_absen = total_tidak_hadir * potongan_absen_per_hari

    # Ambil kasbon yang benar-benar terhubung ke payroll ini
    kasbon_list = Kasbon.query.filter_by(
        technician_id=employee.id,
        payroll_id=payroll.id
    ).all()

    total_kasbon_terpilih = (
        db.session.query(db.func.sum(Kasbon.jumlah))
        .filter_by(technician_id=employee.id, payroll_id=payroll.id)
        .scalar()
        or 0
    )

    # Ambil service motor yang benar-benar terhubung ke payroll ini
    service_motor_list = ServiceMotor.query.filter_by(
        technician_id=employee.id,
        payroll_id=payroll.id
    ).all()

    # NOTE: total potongan service motor sudah tercatat di PayrollItem (kategori Potongan),
    # jadi kita tidak menjumlahkan service_motor_list.porsi_teknisi lagi di perhitungan gaji bersih.
    # Namun kita tetap kirim service_motor_list ke template agar bisa ditampilkan detil-nya.
    # Hitung gaji bersih berdasarkan item + kasbon + potongan absen.
    gaji_bersih = (
    total_pendapatan
    - total_potongan
)

    return render_template(
        'payroll/detail.html',
        payroll=payroll,
        employee=employee,
        items=items,
        kasbon_list=kasbon_list,
        service_motor_list=service_motor_list,

        total_pendapatan=total_pendapatan,
        total_potongan=total_potongan,
        gaji_bersih=gaji_bersih,

        total_hadir=total_hadir,
        total_izin=total_izin,
        total_tidak_hadir=total_tidak_hadir,
        hari_kerja_aktif=hari_kerja_aktif,

        total_kasbon_terpilih=total_kasbon_terpilih
    )



@app.route('/payroll/settings', methods=['GET', 'POST'])
@login_required
def payroll_settings():
    employees = Technician.query.all()
    employee_id = request.args.get('employee_id') or request.form.get('employee_id')

    selected_employee = None
    setting = None

    if employee_id:
        selected_employee = Technician.query.get(int(employee_id))
        setting = PayrollSetting.query.filter_by(employee_id=employee_id).first()

        if request.method == 'POST':
            if not setting:
                setting = PayrollSetting(employee_id=employee_id)
                db.session.add(setting)

            # Ambil nilai form dengan aman
            setting.gaji_pokok = float(request.form.get('gaji_pokok', 0) or 0)
            setting.bonus_bulanan = float(request.form.get('bonus_bulanan', 0) or 0)
            setting.uang_makan_harian = float(request.form.get('uang_makan_harian', 0) or 0)
            setting.uang_bensin_harian = float(request.form.get('uang_bensin_harian', 0) or 0)
            setting.potongan_absen_per_hari = float(request.form.get('potongan_absen_per_hari', 0) or 0)
            setting.potongan_service_default = float(request.form.get('potongan_service_default', 0) or 0)

            db.session.commit()
            flash(f"✅ Pengaturan payroll untuk {selected_employee.nama} berhasil disimpan!", "success")
            return redirect(url_for('payroll_settings', employee_id=employee_id))

    return render_template(
        'payroll/settings.html',
        employees=employees,
        selected_employee=selected_employee,
        setting=setting
    )
@app.route('/payroll/generate', methods=['POST'])
@login_required
def generate_payroll():
    periode = request.form.get('periode')
    employee_ids = request.form.getlist('employee_ids[]')

    if not periode:
        flash("❌ Periode wajib diisi.", "danger")
        return redirect(url_for('payroll_index'))

    try:
        tahun, bulan = map(int, periode.split('-'))
    except ValueError:
        flash("❌ Format periode tidak valid (YYYY-MM).", "danger")
        return redirect(url_for('payroll_index'))

    if employee_ids:
        technicians = Technician.query.filter(
            Technician.id.in_(employee_ids)
        ).all()
    else:
        technicians = Technician.query.filter(
            Technician.status.in_(["Aktif", "Resign"])
        ).all()

    if not technicians:
        flash("⚠️ Tidak ada teknisi ditemukan.", "warning")
        return redirect(url_for('payroll_index'))

    try:
        total_digenerate = 0
        total_skip = 0

        for tech in technicians:

            exists = Payroll.query.filter_by(
                employee_id=tech.id,
                periode=periode
            ).first()

            if exists:
                total_skip += 1
                continue

            setting = PayrollSetting.query.filter_by(
                employee_id=tech.id
            ).first()

            gaji_pokok = float(
                setting.gaji_pokok if setting else (tech.gaji_pokok or 0)
            )

            bonus_bulanan = float(
                setting.bonus_bulanan if setting else 0
            )

            uang_makan_harian = float(
                setting.uang_makan_harian if setting else 0
            )

            uang_bensin_harian = float(
                setting.uang_bensin_harian if setting else 0
            )

            hari_dalam_bulan = monthrange(tahun, bulan)[1]

            awal_bulan = date(tahun, bulan, 1)
            akhir_bulan = date(tahun, bulan, hari_dalam_bulan)

            aktif_mulai = awal_bulan

            if tech.tanggal_masuk:
                aktif_mulai = max(
                    awal_bulan,
                    tech.tanggal_masuk
                )

            aktif_sampai = akhir_bulan

            if tech.tanggal_keluar:
                aktif_sampai = min(
                    akhir_bulan,
                    tech.tanggal_keluar
                )

            if aktif_sampai < aktif_mulai:
                continue

            hari_kerja_aktif = hitung_hari_kerja(
                aktif_mulai,
                aktif_sampai
            )

            total_hadir = Attendance.query.filter(
                Attendance.technician_id == tech.id,
                Attendance.tanggal >= aktif_mulai,
                Attendance.tanggal <= aktif_sampai,
                Attendance.status == "Hadir"
            ).count()

            total_izin = Attendance.query.filter(
                Attendance.technician_id == tech.id,
                Attendance.tanggal >= aktif_mulai,
                Attendance.tanggal <= aktif_sampai,
                Attendance.status.in_(["Izin", "Sakit"])
            ).count()

            total_alpa = max(
                hari_kerja_aktif -
                total_hadir -
                total_izin,
                0
            )

            gaji_per_hari = gaji_pokok / 26 if gaji_pokok > 0 else 0

            gaji_pokok_prorata = (
                gaji_per_hari *
                hari_kerja_aktif
            )

            potongan_absen = (
                total_alpa *
                gaji_per_hari
            )

            kasbon_list = Kasbon.query.filter(
                Kasbon.technician_id == tech.id,
                Kasbon.status == "Belum Lunas",
                Kasbon.terpilih == True,
                Kasbon.payroll_id.is_(None)
            ).all()

            potongan_kasbon = sum(
                k.jumlah or 0
                for k in kasbon_list
            )

            service_motor_list = ServiceMotor.query.filter(
                ServiceMotor.technician_id == tech.id,
                ServiceMotor.terpilih == True,
                ServiceMotor.payroll_id.is_(None)
            ).all()

            potongan_service_motor = sum(
                s.porsi_teknisi or 0
                for s in service_motor_list
            )

            total_pendapatan = (
                gaji_pokok_prorata +
                bonus_bulanan +
                (uang_makan_harian * total_hadir) +
                (uang_bensin_harian * total_hadir)
            )

            total_potongan = (
                potongan_absen +
                potongan_kasbon +
                potongan_service_motor
            )

        payroll = Payroll(
                employee_id=tech.id,
                periode=periode,
                total_pendapatan=total_pendapatan,
                total_potongan=total_potongan,
                gaji_bersih=total_pendapatan - total_potongan,
                total_tidak_hadir=total_alpa,
                status="Draft"
            )

        db.session.add(payroll)
        db.session.flush()

        items = []

            # =====================
            # PENDAPATAN
            # =====================

        if gaji_pokok_prorata > 0:
                items.append(
                    PayrollItem(
                        payroll_id=payroll.id,
                        kategori="Pendapatan",
                        deskripsi="Gaji Pokok",
                        jumlah=gaji_pokok_prorata
                    )
                )

        if bonus_bulanan > 0:
                items.append(
                    PayrollItem(
                        payroll_id=payroll.id,
                        kategori="Pendapatan",
                        deskripsi="Bonus Bulanan",
                        jumlah=bonus_bulanan
                    )
                )

        if uang_makan_harian > 0 and total_hadir > 0:
                items.append(
                    PayrollItem(
                        payroll_id=payroll.id,
                        kategori="Pendapatan",
                        deskripsi=f"Uang Makan ({total_hadir} hari)",
                        jumlah=uang_makan_harian * total_hadir
                    )
                )

        if uang_bensin_harian > 0 and total_hadir > 0:
                items.append(
                    PayrollItem(
                        payroll_id=payroll.id,
                        kategori="Pendapatan",
                        deskripsi=f"Uang Bensin ({total_hadir} hari)",
                        jumlah=uang_bensin_harian * total_hadir
                    )
                )

            # =====================
            # POTONGAN
            # =====================

        if potongan_absen > 0:
                items.append(
                    PayrollItem(
                        payroll_id=payroll.id,
                        kategori="Potongan",
                        deskripsi=f"Potongan Absen ({total_alpa} hari)",
                        jumlah=potongan_absen
                    )
                )

        if potongan_kasbon > 0:
                items.append(
                    PayrollItem(
                        payroll_id=payroll.id,
                        kategori="Potongan",
                        deskripsi="Potongan Kasbon",
                        jumlah=potongan_kasbon
                    )
                )

        if potongan_service_motor > 0:
                items.append(
                    PayrollItem(
                        payroll_id=payroll.id,
                        kategori="Potongan",
                        deskripsi="Potongan Service Motor",
                        jumlah=potongan_service_motor
                    )
                )

        db.session.add_all(items)

        for k in kasbon_list:
                k.payroll_id = payroll.id
                k.status = "Diproses Payroll"
                k.terpilih = False

        for s in service_motor_list:
                s.payroll_id = payroll.id

        print("=" * 50)
        print("PAYROLL:", payroll.id)
        print("ITEMS DIBUAT:", len(items))
        print("=" * 50)
        total_digenerate += 1

        db.session.commit()

        flash(
            f"✅ Payroll {periode} selesai. Digenerate: {total_digenerate}, Dilewati: {total_skip}",
            "success"
        )

    except SQLAlchemyError as e:
        db.session.rollback()

        flash(
            f"❌ Error generate payroll: {str(e)}",
            "danger"
        )

    return redirect(url_for('payroll_index'))

# =========================
# Hapus Payroll (DIPERBAIKI)
# =========================
@app.route('/payroll/<int:id>/delete', methods=['POST'])
@login_required
def payroll_delete(id):
    payroll = Payroll.query.get_or_404(id)

    # Reset semua kasbon terkait
    kasbons = Kasbon.query.filter_by(payroll_id=payroll.id).all()
    for k in kasbons:
        k.payroll_id = None
        k.status = 'Belum Lunas'
        k.terpilih = False
        db.session.add(k)

    # Reset semua service motor terkait (pastikan tidak double-count di periode lain)
    services = ServiceMotor.query.filter_by(payroll_id=payroll.id).all()
    for s in services:
        s.payroll_id = None
        s.status = 'Belum Dipotong'
        s.terpilih = False
        db.session.add(s)

    db.session.delete(payroll)
    db.session.commit()
    flash('🗑️ Payroll dihapus dan kasbon + service motor dikembalikan ke status awal.', 'warning')
    return redirect(url_for('payroll_index'))


@app.template_filter("rupiah")
def rupiah(value):
    try:
        value = float(value or 0)
        return f"Rp {value:,.0f}".replace(",", ".")
    except Exception:
        return "Rp 0"


# =========================
# FINALISASI PAYROLL (DIPERBAIKI)
# =========================
@app.route('/payroll/<int:id>/final', methods=['POST'])
@login_required
def payroll_final(id):
    payroll = Payroll.query.get_or_404(id)

    if payroll.status == 'Final':
        flash('⚠️ Payroll sudah difinalisasi sebelumnya.', 'info')
        return redirect(url_for('payroll_detail', id=id))

    payroll.status = 'Final'

    # Tandai semua kasbon yang terhubung dengan payroll ini menjadi Lunas
    kasbons = Kasbon.query.filter_by(payroll_id=payroll.id).all()
    for k in kasbons:
        k.status = 'Lunas'
        k.terpilih = False
        db.session.add(k)

    # Tandai semua service motor yang terhubung dengan payroll ini sebagai Sudah Dipotong
    services = ServiceMotor.query.filter_by(payroll_id=payroll.id).all()
    for s in services:
        s.status = 'Sudah Dipotong'
        s.terpilih = False
        # payroll_id tetap terisi agar jejak periodenya aman
        db.session.add(s)

    db.session.commit()
    flash('✅ Payroll berhasil difinalisasi. Kasbon ditandai lunas & service motor ditandai sudah dipotong.', 'success')
    return redirect(url_for('payroll_detail', id=id))


# =========================
# UNDO FINALISASI PAYROLL (DIPERBAIKI)
# =========================
@app.route('/payroll/<int:id>/undo_final', methods=['POST'])
@login_required
def payroll_undo_final(id):
    payroll = Payroll.query.get_or_404(id)

    if payroll.status != 'Final':
        flash('⚠️ Payroll belum difinalisasi, tidak bisa di-undo.', 'warning')
        return redirect(url_for('payroll_detail', id=id))

    payroll.status = 'Draft'

    # Kembalikan kasbon ke Belum Lunas
    kasbons = Kasbon.query.filter_by(payroll_id=payroll.id).all()
    for k in kasbons:
        k.status = 'Belum Lunas'
        k.terpilih = False
        # tetap hubungkan ke payroll sampai user memutuskan menghapus payroll
        db.session.add(k)

    # Kembalikan service motor ke Belum Dipotong, namun tetap pertahankan payroll_id
    # (agar jejak historis tetap ada) — jika Anda ingin service motor bisa dipakai lagi
    # di payroll berikutnya, Anda harus menghapus payroll terlebih dahulu.
    services = ServiceMotor.query.filter_by(payroll_id=payroll.id).all()
    for s in services:
        s.status = 'Belum Dipotong'
        s.terpilih = False
        db.session.add(s)

    db.session.commit()
    flash('⏪ Payroll dikembalikan ke Draft. Semua kasbon & service motor dikembalikan ke status awal.', 'info')
    return redirect(url_for('payroll_detail', id=id))


import pdfkit
from flask import make_response, render_template
from flask import make_response, render_template, request
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from datetime import datetime
import io


@app.route('/payroll/<int:id>/print')
@login_required
def payroll_print(id):
    payroll = Payroll.query.get_or_404(id)
    employee = payroll.employee
    items = PayrollItem.query.filter_by(payroll_id=id).all()
    print("PAYROLL:", payroll.id)
    print("ITEMS:", len(items))

    total_pendapatan = sum(i.jumlah or 0 for i in items if i.kategori == 'Pendapatan')
    total_potongan = sum(i.jumlah or 0 for i in items if i.kategori == 'Potongan')

    setting = PayrollSetting.query.filter_by(employee_id=employee.id).first()
    total_tidak_hadir = payroll.total_tidak_hadir or 0
    potongan_absen_per_hari = setting.potongan_absen_per_hari if setting else 0
    potongan_absen = total_tidak_hadir * potongan_absen_per_hari

    total_kasbon = db.session.query(db.func.sum(Kasbon.jumlah)) \
        .filter_by(payroll_id=payroll.id).scalar() or 0

    total_service_motor = db.session.query(db.func.sum(ServiceMotor.porsi_teknisi)) \
        .filter_by(payroll_id=payroll.id).scalar() or 0

    gaji_bersih = (
        total_pendapatan
        - total_potongan
        
    )

    # 👇 Jika parameter ?pdf=true → generate PDF pakai ReportLab
    if request.args.get("pdf") == "true":
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer, pagesize=A4,
            rightMargin=20*mm, leftMargin=20*mm,
            topMargin=15*mm, bottomMargin=15*mm
        )
        styles = getSampleStyleSheet()
        content = []

        # Header perusahaan
        try:
            logo = Image("static/logo.png", width=40*mm, height=20*mm)
            logo.hAlign = 'LEFT'
            content.append(logo)
        except:
            pass

        content.append(Paragraph("<b>Perkasa AC</b>", styles["Normal"]))
        content.append(Paragraph("Jl. Sukabangun 2, KM 6.5, Palembang", styles["Normal"]))
        content.append(Spacer(1, 10))
        content.append(Paragraph("<b><font size=16>Slip Gaji</font></b>", styles["Heading4"]))
        content.append(Spacer(1, 10))

        # Data karyawan
        data_karyawan = [
            ["Nama / NIK", ":", employee.nama],
            ["Jabatan", ":", employee.jabatan or "-"],
            ["Tanggal Masuk", ":", employee.tanggal_masuk.strftime("%d-%m-%Y") if employee.tanggal_masuk else "-"],
            ["Periode", ":", payroll.periode or "-"],
        ]
        table_karyawan = Table(data_karyawan, colWidths=[100, 10, 300])
        table_karyawan.setStyle(TableStyle([
            ("FONTSIZE", (0,0), (-1,-1), 10),
            ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ]))
        content.append(table_karyawan)
        content.append(Spacer(1, 10))

        # Pendapatan & potongan
        pendapatan = [["<b>Pendapatan</b>", ""]]
        potongan = [["<b>Potongan</b>", ""]]

        for i in items:
            if i.kategori == "Pendapatan":
                pendapatan.append([i.deskripsi, f"{i.jumlah:,.0f}"])
            elif i.kategori == "Potongan":
                potongan.append([i.deskripsi, f"-{i.jumlah:,.0f}"])

        pendapatan.append(["Total Pendapatan", f"{total_pendapatan:,.0f}"])
        potongan.append(["Total Potongan", f"-{total_potongan:,.0f}"])

        table_pendapatan = Table(pendapatan, colWidths=[200, 80])
        table_potongan = Table(potongan, colWidths=[200, 80])

        for t in [table_pendapatan, table_potongan]:
            t.setStyle(TableStyle([
                ("ALIGN", (1,0), (-1,-1), "RIGHT"),
                ("FONTSIZE", (0,0), (-1,-1), 10),
                ("GRID", (0,-1), (-1,-1), 0.3, colors.grey),
            ]))

        table_duo = Table([[table_pendapatan, table_potongan]], colWidths=[250, 250])
        table_duo.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP")]))
        content.append(table_duo)
        content.append(Spacer(1, 15))

        # Total gaji bersih
        content.append(Paragraph("<b>Total Diterima:</b>", styles["Heading4"]))
        content.append(Paragraph(f"<font size=14><b>Rp {gaji_bersih:,.0f}</b></font>", styles["Heading4"]))
        content.append(Spacer(1, 10))



        # Footer
        content.append(Paragraph(f"Dicetak pada: {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles["Normal"]))
        content.append(Spacer(1, 15))
        content.append(Paragraph("__________________________", styles["Normal"]))
        content.append(Paragraph("Tanda Tangan Karyawan", styles["Normal"]))

        # Bangun PDF
        doc.build(content)
        pdf = buffer.getvalue()
        buffer.close()

        response = make_response(pdf)
        response.headers["Content-Type"] = "application/pdf"
        response.headers["Content-Disposition"] = f'inline; filename=Slip_Gaji_{employee.nama}.pdf'
        return response

    # 👇 Kalau bukan PDF, tetap render HTML biasa
    html = render_template(
        'payroll/print_slip.html',
        payroll=payroll,
        employee=employee,
        items=items,
        total_pendapatan=total_pendapatan,
        total_potongan=total_potongan,
        total_tidak_hadir=total_tidak_hadir,
        potongan_absen=potongan_absen,
        total_kasbon=total_kasbon,
        total_service_motor=total_service_motor,
        gaji_bersih=gaji_bersih,
        print_date=datetime.now().strftime("%d/%m/%Y %H:%M")
    )
    return html




# ==============================
# 📘 ROUTE KASBON KARYAWAN
# ==============================
@app.route('/kasbon')
@login_required
def kasbon_list():
    bulan = request.args.get('bulan')
    tahun = request.args.get('tahun')
    teknisi_id = request.args.get('teknisi_id')

    # Mulai query dasar
    query = Kasbon.query

    # Filter berdasarkan bulan dan tahun
    if bulan and tahun:
        query = query.filter(
            db.extract('month', Kasbon.tanggal) == int(bulan),
            db.extract('year', Kasbon.tanggal) == int(tahun)
        )
    elif tahun:  # hanya tahun
        query = query.filter(db.extract('year', Kasbon.tanggal) == int(tahun))

    # Filter berdasarkan teknisi
    if teknisi_id and teknisi_id != "all":
        query = query.filter(Kasbon.technician_id == int(teknisi_id))

    kasbons = query.order_by(Kasbon.tanggal.desc()).all()

    # Ambil semua teknisi untuk dropdown
    teknisis = Technician.query.order_by(Technician.nama).all()

    return render_template(
        'karyawan/kasbon_list.html',
        kasbons=kasbons,
        teknisis=teknisis,
        bulan=bulan,
        tahun=tahun,
        teknisi_id=teknisi_id
    )



@app.route('/kasbon/tambah', methods=['GET', 'POST'])
@login_required
def kasbon_tambah():
    technicians = Technician.query.all()
    if request.method == 'POST':
        tech_id = request.form['technician_id']
        jumlah = float(request.form['jumlah'])
        keterangan = request.form.get('keterangan', '')
        tanggal_str = request.form.get('tanggal')

        # parsing tanggal dari form
        if tanggal_str:
            try:
                tanggal = datetime.strptime(tanggal_str, "%Y-%m-%d").date()
            except ValueError:
                flash("Format tanggal tidak valid", "danger")
                return redirect(url_for('kasbon_tambah'))
        else:
            tanggal = date.today()

        kasbon = Kasbon(
            technician_id=tech_id,
            tanggal=tanggal,
            jumlah=jumlah,
            keterangan=keterangan,
            status="Belum Lunas"
        )
        db.session.add(kasbon)
        db.session.commit()
        flash("✅ Kasbon berhasil ditambahkan", "success")
        return redirect(url_for('kasbon_list'))

    return render_template('karyawan/kasbon_form.html', technicians=technicians, date=date)


@app.route('/kasbon/<int:id>/hapus')
@login_required
def kasbon_hapus(id):
    kasbon = Kasbon.query.get_or_404(id)
    db.session.delete(kasbon)
    db.session.commit()
    flash("Kasbon dihapus", "info")
    return redirect(url_for('kasbon_list'))


# @app.route('/kasbon/<int:id>/lunas')
# @login_required
# def kasbon_lunas(id):
#     kasbon = Kasbon.query.get_or_404(id)
#     kasbon.status = "Lunas"
#     db.session.commit()
#     flash("Kasbon telah dilunasi", "success")
#     return redirect(url_for('kasbon_list'))

@app.route('/kasbon/detail/<int:technician_id>')
@login_required
def kasbon_detail(technician_id):
    tech = Technician.query.get_or_404(technician_id)
    kasbons = Kasbon.query.filter_by(technician_id=technician_id).order_by(Kasbon.tanggal.desc()).all()
    return render_template('karyawan/kasbon_detail.html', tech=tech, kasbons=kasbons)


@app.route('/kasbon/tambah_ke_payroll', methods=['POST'])
@login_required
def kasbon_tambah_ke_payroll():
    kasbon_ids = request.form.getlist('kasbon_ids[]')
    if not kasbon_ids:
        flash('Pilih minimal satu kasbon untuk ditambahkan ke payroll.', 'warning')
        return redirect(url_for('kasbon_list'))

    kasbons = Kasbon.query.filter(Kasbon.id.in_(kasbon_ids)).all()
    if not kasbons:
        flash('Kasbon tidak ditemukan.', 'danger')
        return redirect(url_for('kasbon_list'))

    # Tandai kasbon terpilih (tapi belum lunas)
    for k in kasbons:
        k.terpilih = True
        k.status = 'Belum Lunas'
        db.session.add(k)

    db.session.commit()
    flash('✅ Kasbon berhasil ditandai untuk dimasukkan ke payroll.', 'success')
    return redirect(url_for('kasbon_list'))

@app.route('/debug/kasbon/<int:tech_id>')
def debug_kasbon(tech_id):
    kasbons = Kasbon.query.filter_by(technician_id=tech_id).all()
    return {
        "kasbon": [
            {
                "id": k.id,
                "status": k.status,
                "terpilih": k.terpilih,
                "jumlah": k.jumlah,
                "keterangan": k.keterangan
            } for k in kasbons
        ]
    }


# ✅ Reminder otomatis untuk service cuci (H-14 s/d H-1)
@app.route('/reminders/auto')
@login_required
def reminders_auto():
    today = date.today()

    services = (
    Service.query
    .join(Service.ac_unit)
    .join(ACUnit.customer)
    .all()
)


    # -------------------------
    # KUMPULKAN DATA PER CUSTOMER
    # -------------------------
    grouped = {}   # key = customer_id, value = service info terdekat

    for s in services:

        # fallback next_date
        next_date = s.tanggal_cuci_berikutnya or (s.tanggal + relativedelta(months=3))
        delta_days = (next_date - today).days
        if not (1 <= delta_days <= 14):
            continue

        unit = s.ac_unit
        cust = unit.customer

        cust_id = cust.id

        # Jika belum ada entry untuk customer ini → masukkan
        if cust_id not in grouped:
            grouped[cust_id] = {
                "id": s.id,
                "nama": cust.nama,
                "wa": cust.nomor_wa,
                "alamat": cust.alamat,
                "team": cust.team,
                "unit": f"{unit.merk} ({unit.pk}) x{unit.jumlah_unit}",
                "tanggal": next_date,
                "delta": delta_days,
                "service": s
            }
        else:
            # Sudah ada → pilih service yg tanggalnya lebih dekat (lebih kecil)
            if next_date < grouped[cust_id]["tanggal"]:
                grouped[cust_id] = {
                    "id": s.id,
                    "nama": cust.nama,
                    "wa": cust.nomor_wa,
                    "alamat": cust.alamat,
                    "team": cust.team,
                    "unit": f"{unit.merk} ({unit.pk}) x{unit.jumlah_unit}",
                    "tanggal": next_date,
                    "delta": delta_days,
                    "service": s
                }

    reminders = []

    # PROSES DATA AKHIR MENJADI REMINDER UNIK
    for cust_id, data in grouped.items():
        s = data["service"]
        reminder = Reminder.query.filter_by(service_id=s.id).first()

        reminders.append({
            "id": s.id,
            "nama": data["nama"],
            "wa": data["wa"],
            "alamat": data["alamat"],
            "team": data["team"],
            "unit": data["unit"],
            "tanggal": data["tanggal"],
            "tipe": f"H-{data['delta']}",
            "keterangan": reminder.keterangan_reminder if reminder else "Belum di-Reminder",
            "reminder_sent": getattr(reminder, 'reminder_sent', False)
        })

    reminders = sorted(reminders, key=lambda x: x["tanggal"])
    return render_template("reminders_auto.html", reminders=reminders, current_date=today)


@app.route('/reminder/send/<int:id>', methods=['POST'])
@login_required
def reminder_send(id):
    try:
        service = Service.query.get(id)
        if not service:
            return jsonify(success=False, message="Service tidak ditemukan"), 404

        # ✅ Akses customer lewat ACUnit
        if not service.ac_unit or not service.ac_unit.customer:
            return jsonify(success=False, message="Relasi Customer tidak ditemukan"), 404

        customer = service.ac_unit.customer
        tanggal = service.tanggal_cuci_berikutnya or (service.tanggal + relativedelta(months=3))

        # ✅ Format pesan WhatsApp
        pesan = (
            f"Halo Kak *{customer.nama}*, 👋\n\n"
            "Kami dari *PERKASA AC* ingin mengingatkan bahwa jadwal cuci AC Anda adalah:\n\n"
            f"📅 *Tanggal:* {tanggal.strftime('%d/%m/%Y')}\n"
            f"📍 *Alamat:* {customer.alamat}\n"
            f"❄️ *Unit:* {service.ac_unit.merk} ({service.ac_unit.pk}) x{service.ac_unit.jumlah_unit}\n\n"
            "Apabila ada perubahan jadwal, silakan hubungi kami untuk *reschedule*.\n"
            "Terima kasih atas kepercayaannya 🙏"
        )

        # ✅ Encode pesan & buat link WA
        encoded_pesan = quote(pesan)
        wa_url = f"https://web.whatsapp.com/send?phone={customer.nomor_wa}&text={encoded_pesan}"

        # ✅ Update atau buat Reminder
        reminder = Reminder.query.filter_by(service_id=id).first()
        if reminder:
            reminder.keterangan_reminder = "Sudah di-Reminder"
            reminder.reminder_sent = True
        else:
            reminder = Reminder(
                service_id=id,
                keterangan_reminder="Sudah di-Reminder",
                reminder_sent=True
            )
            db.session.add(reminder)

        db.session.commit()

        return jsonify(success=True, message="Reminder berhasil dibuat", wa_url=wa_url)

    except Exception as e:
        print("ERROR REMINDER:", e)
        return jsonify(success=False, message=f"Terjadi kesalahan: {e}"), 500


# Upload Customers via Excel
# ======================================
# Upload Customers via Excel
# ======================================
@app.route('/upload/customers', methods=['GET','POST'])
@login_required
def upload_customers():
    import pandas as pd
    from app import normalize_wa

    if request.method == 'POST':
        f = request.files.get('file')
        if not f:
            flash('File tidak ditemukan', 'danger')
            return redirect(url_for('upload_customers'))

        df = pd.read_excel(f)

        inserted, skipped = 0, 0
        duplicates = []

        for _, row in df.iterrows():
            # ✅ Normalisasi nomor WA
            no_wa_raw = str(row.get('nomor_wa') or row.get('Nomor WA') or '').strip()
            no_wa = normalize_wa(no_wa_raw)

            nama = str(row.get('nama') or row.get('Nama') or 'Unknown').strip()

            if not no_wa:
                skipped += 1
                continue

            # ✅ Cek duplikat customer
            existing = Customer.query.filter_by(nomor_wa=no_wa).first()
            if existing:
                skipped += 1
                duplicates.append(f"{existing.nama} ({existing.nomor_wa})")
                continue

            # ✅ Simpan customer baru
            c = Customer(
                nama=nama,
                nomor_wa=no_wa,
                alamat=row.get('alamat') or row.get('Alamat') or '',
                team=row.get('team') or row.get('Team') or ''
            )
            db.session.add(c)
            db.session.flush()

            # ✅ Tambah unit jika ada
            if row.get('merk'):
                u = ACUnit(
                    customer_id=c.id,
                    merk=row.get('merk'),
                    pk=str(row.get('pk') or ''),
                    jumlah_unit=int(row.get('jumlah_unit') or 1)
                )
                db.session.add(u)

            inserted += 1

        db.session.commit()

        if skipped > 0:
            flash(
                f'Import Customers selesai. Ditambahkan: {inserted}, duplikat dilewati: {skipped}. '
                f'Duplikat: {", ".join(duplicates)}',
                'warning'
            )
        else:
            flash(f'Import Customers selesai. Ditambahkan: {inserted}, tidak ada duplikat.', 'success')

        return redirect(url_for('customers'))

    return render_template('upload_modern.html', tipe="customers")


# ======================================
# Upload Services via Excel
# ======================================
@app.route('/upload/services', methods=['GET', 'POST'])
@login_required
def upload_services():
    import pandas as pd
    from datetime import date, datetime
    errors = []

    def normalize_wa(wa):
        if not wa:
            return None
        wa_str = str(wa).replace("\u202c", "").replace("\u202d", "").strip()
        wa_str = wa_str.replace("+", "").replace(" ", "")
        if wa_str.startswith("0"):
            wa_str = "62" + wa_str[1:]
        elif not wa_str.startswith("62"):
            wa_str = "62" + wa_str
        return wa_str

    if request.method == 'POST':
        f = request.files.get('file')
        if not f:
            flash('File tidak ditemukan', 'danger')
            return redirect(url_for('upload_services'))

        try:
            df = pd.read_excel(f)
        except Exception as e:
            flash(f'Gagal membaca file Excel: {e}', 'danger')
            return redirect(url_for('upload_services'))

        inserted, skipped = 0, 0

        for idx, row in df.iterrows():
            try:
                wa = normalize_wa(row.get("customer_wa"))
                merk = str(row.get("merk") or "Unknown").strip()
                pk = str(row.get("pk") or "").strip()

                # jumlah_unit
                try:
                    jumlah_unit = int(row.get("jumlah_unit")) if pd.notna(row.get("jumlah_unit")) else 1
                except:
                    jumlah_unit = 1

                # harga
                try:
                    harga = float(row.get("harga")) if pd.notna(row.get("harga")) else 0.0
                except:
                    harga = 0.0

                # tanggal
                tgl_service = row.get("tanggal")
                if pd.notna(tgl_service):
                    try:
                        tgl_service = pd.to_datetime(tgl_service).date()
                    except:
                        raise ValueError("Tanggal tidak valid")
                else:
                    tgl_service = date.today()

                # customer
                customer = Customer.query.filter_by(nomor_wa=wa).first() if wa else None
                if not customer:
                    nomor_wa_to_use = wa or f"62{int(datetime.utcnow().timestamp())}"
                    customer = Customer(
                        nama=str(row.get("customer_nama") or "Customer Baru"),
                        nomor_wa=nomor_wa_to_use,
                        alamat=str(row.get("alamat") or ""),
                        team=str(row.get("team") or "")
                    )
                    db.session.add(customer)
                    db.session.commit()

                # unit
                unit = ACUnit.query.filter_by(customer_id=customer.id, merk=merk, pk=pk).first()
                if not unit:
                    unit = ACUnit(customer_id=customer.id, merk=merk, pk=pk, jumlah_unit=jumlah_unit)
                    db.session.add(unit)
                    db.session.commit()

                # teknisi
                teknisi_nama = (row.get("teknisi") or "").strip()
                technician_obj = None
                if teknisi_nama:
                    technician_obj = Technician.query.filter(
                        Technician.nama.ilike(f"%{teknisi_nama}%")
                    ).first()

                service = Service(
                    unit_id=unit.id,
                    tanggal=tgl_service,
                    jenis_service=str(row.get("jenis_service") or "Cuci AC"),
                    paket_cuci=str(row.get("paket_cuci") or "") if pd.notna(row.get("paket_cuci")) else None,
                    harga_satuan=harga,
                    jumlah=jumlah_unit,
                    pembayaran=str(row.get("pembayaran") or "Belum Lunas"),
                    deskripsi=str(row.get("deskripsi") or ""),
                    team=str(row.get("team") or "")
                )

                if technician_obj:
                    service.technician = technician_obj
                    service.teknisi = technician_obj.nama
                elif teknisi_nama:
                    service.teknisi = teknisi_nama

                db.session.add(service)
                db.session.commit()
                inserted += 1

            except Exception as e:
                skipped += 1
                db.session.rollback()
                errors.append(f"Row {idx+2}: {str(e)}")  # tampilkan baris Excel + error

        msg = f"Services diupload: {inserted} sukses, {skipped} gagal"
        if errors:
            msg += "<br>" + "<br>".join(errors)
        flash(msg, 'danger' if skipped else 'success')
        return redirect(url_for('upload_services'))

    return render_template('upload_modern.html')



# ✅ Route untuk download log
@app.route('/download_log/<filename>')
@login_required
def download_log(filename):
    log_dir = os.path.join(os.getcwd(), "upload_logs")
    return send_from_directory(log_dir, filename, as_attachment=True)


@app.route('/technicians/update', methods=['POST'])
@login_required
def update_technicians():
    # ambil data dari form
    tech_id = request.form.get("id")
    name = request.form.get("name")

    technician = Technician.query.get_or_404(tech_id)
    technician.name = name
    db.session.commit()

    flash("Data teknisi berhasil diperbarui", "success")
    return redirect(url_for('dashboard'))


@app.route('/template/customers')
@login_required
def download_template_customers():
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Customers"

    # Header
    ws.append(["nama", "nomor_wa", "alamat", "team", "merk", "pk", "jumlah_unit"])

    # Contoh data
    ws.append(["Agus", "628123456789", "Jl. Merdeka No. 10", "Internal", "LG", "1", "2"])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="template_customers.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route('/download_template_services')
@login_required
def download_template_services():
    wb = Workbook()
    ws = wb.active
    ws.title = "Services"

    # 🔹 Header kolom sesuai format
    headers = [
        "customer_wa", "tanggal", "jenis_service", "paket_cuci",
        "harga", "teknisi", "pembayaran", "deskripsi",
        "jumlah_unit", "merk", "pk", "tanggal_cuci_berikutnya"
    ]
    ws.append(headers)

    # 🔹 Contoh data
    ws.append([
        "628123456789", "2025-09-10", "Cuci AC", "Segar sungai musi 1/2-1PK",
        75000, "Rio", "Lunas", "Cuci rutin", 3, "LG", "1PK", "2026-03-10"
    ])
    ws.append([
        "628987654321", "2025-09-12", "Tambah Freon", "-",
        150000, "Budi", "Belum Lunas", "Isi freon R32", 1, "Sharp", "1PK", ""
    ])

    # 🔹 Simpan ke memory
    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name="template_services.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route('/download/template/teknisi')
@login_required
def download_template_teknisi():
    import pandas as pd
    from io import BytesIO
    from flask import send_file

    # contoh template teknisi
    df = pd.DataFrame(columns=["Nama", "Nomor WA", "Alamat", "Team"])
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="template_teknisi.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )



@app.route('/absen/<int:technician_id>')
@login_required
def absen(technician_id):
    # Ambil data teknisi/karyawan berdasarkan ID
    employee = db.session.execute(
        db.select(Employee).filter_by(id=technician_id)
    ).scalar_one_or_none()

    if not employee:
        flash("Karyawan tidak ditemukan.", "danger")
        return redirect(url_for('employees'))

    # Generate URL QR absensi (misalnya untuk check-in)
    qr_url = url_for('attendance_checkin', technician_id=technician_id, _external=True)

    return render_template('employee_absen.html', employee=employee, qr_url=qr_url)


# 🔹 Tambah Absen Manual
@app.route('/attendance/add', methods=['POST'])
@login_required
def attendance_add():
    try:
        technician_id = int(request.form.get('technician_id'))
        tanggal_str = request.form.get('tanggal')
        jam_masuk_str = request.form.get('jam_masuk')
        jam_pulang_str = request.form.get('jam_pulang')
        status = request.form.get('status') or 'Hadir'
        keterangan = request.form.get('keterangan') or None

        tanggal = datetime.strptime(tanggal_str, '%Y-%m-%d').date() if tanggal_str else None

        def parse_time(time_str):
            if not time_str:
                return None
            for fmt in ('%H:%M', '%H:%M:%S'):
                try:
                    return datetime.strptime(time_str, fmt).time()
                except ValueError:
                    continue
            return None

        jam_masuk = parse_time(jam_masuk_str)
        jam_pulang = parse_time(jam_pulang_str)

        rec = Attendance(
            technician_id=technician_id,
            tanggal=tanggal,
            jam_masuk=jam_masuk,
            jam_pulang=jam_pulang,
            status=status,
            keterangan=keterangan
        )
        db.session.add(rec)
        db.session.commit()
        flash('Absensi berhasil ditambahkan.', 'success')
    except Exception as e:
        db.session.rollback()
        flash(f'Gagal menambahkan absensi: {str(e)}', 'danger')

    return redirect(url_for('attendance_report'))


# 🔹 Halaman Edit Absensi
@app.route('/attendance/edit/<int:id>', methods=['POST'])
@login_required
def attendance_edit(id):
    record = Attendance.query.get_or_404(id)

    record.tanggal = datetime.strptime(request.form['tanggal'], '%Y-%m-%d').date()
    record.jam_masuk = datetime.strptime(request.form['jam_masuk'], '%H:%M').time() if request.form['jam_masuk'] else None
    record.jam_pulang = datetime.strptime(request.form['jam_pulang'], '%H:%M').time() if request.form['jam_pulang'] else None
    record.status = request.form['status']
    record.keterangan = request.form['keterangan']

    db.session.commit()
    return redirect(url_for('attendance_report'))

# 🔹 Halaman Form Edit Absensi (GET)
@app.route('/attendance/edit/<int:id>', methods=['GET'])
@login_required
def attendance_edit_form(id):
    record = Attendance.query.get_or_404(id)
    technicians = Technician.query.all()
    return render_template('attendance/edit.html', record=record, technicians=technicians)




# 🔹 Laporan Absensi
@app.route('/attendance/report')
@login_required
def attendance_report():
    from sqlalchemy import extract
    month = request.args.get('month', datetime.now().month, type=int)
    year = request.args.get('year', datetime.now().year, type=int)

    records = Attendance.query.filter(
        extract('month', Attendance.tanggal) == month,
        extract('year', Attendance.tanggal) == year
    ).order_by(Attendance.tanggal, Attendance.jam_masuk).all()

    technicians = Technician.query.order_by(Technician.nama).all()

    return render_template('attendance_report.html',
                           records=records, month=month, year=year, technicians=technicians)


# 🔹 Hapus Absen
@app.route('/attendance/delete/<int:id>', methods=['POST'])
@login_required
def delete_attendance(id):
    record = Attendance.query.get_or_404(id)
    db.session.delete(record)
    db.session.commit()
    flash('Data absensi berhasil dihapus.', 'success')
    return redirect(url_for('attendance_report'))

@app.route('/employees')
@login_required
def employees():

    q = request.args.get("q", "").strip()
    team = request.args.get("team", "")
    status = request.args.get("status", "")
    jabatan = request.args.get("jabatan", "")

    query = Technician.query

    if q:
        query = query.filter(
            or_(
                Technician.nama.ilike(f"%{q}%"),
                Technician.no_hp.ilike(f"%{q}%")
            )
        )

    if team:
        query = query.filter(Technician.team == team)

    if status:
        query = query.filter(Technician.status == status)

    if jabatan:
        query = query.filter(Technician.jabatan == jabatan)

    employees = query.order_by(Technician.nama.asc()).all()

    return render_template(
        'employees.html',
        employees=employees,
        q=q,
        team=team,
        status=status,
        jabatan=jabatan
    )

@app.route("/employees/<int:id>")
@login_required
def employee_detail(id):

    technician = Technician.query.get_or_404(id)

    services = (
        Service.query
        .filter_by(technician_id=id)
        .order_by(Service.tanggal.desc())
        .limit(20)
        .all()
    )

    attendances = (
        Attendance.query
        .filter_by(technician_id=id)
        .order_by(Attendance.tanggal.desc())
        .limit(30)
        .all()
    )

    kasbons = (
        Kasbon.query
        .filter_by(technician_id=id)
        .order_by(Kasbon.tanggal.desc())
        .all()
    )

    service_motors = (
        ServiceMotor.query
        .filter_by(technician_id=id)
        .order_by(ServiceMotor.tanggal.desc())
        .all()
    )

    payrolls = (
        Payroll.query
        .filter_by(employee_id=id)
        .order_by(Payroll.periode.desc())
        .all()
    )
    # Total Service
    total_service = Service.query.filter_by(
        technician_id=id
    ).count()

    # Total Pendapatan Service
    total_pendapatan = (
        db.session.query(
            func.sum(Service.harga_satuan * Service.jumlah)
        )
        .filter(Service.technician_id == id)
        .scalar()
    ) or 0

    # Total Kasbon Belum Lunas
    total_kasbon = (
        db.session.query(
            func.sum(Kasbon.jumlah)
        )
        .filter(
            Kasbon.technician_id == id,
            Kasbon.status == "Belum Lunas"
        )
        .scalar()
    ) or 0

    # Total Hadir Bulan Ini
    bulan_ini = datetime.now().month
    tahun_ini = datetime.now().year

    total_hadir = Attendance.query.filter(
        Attendance.technician_id == id,
        Attendance.status == "Hadir",
        extract('month', Attendance.tanggal) == bulan_ini,
        extract('year', Attendance.tanggal) == tahun_ini
    ).count()

    return render_template(
        "employee_detail.html",
        technician=technician,
        services=services,
        attendances=attendances,
        kasbons=kasbons,
        service_motors=service_motors,
        payrolls=payrolls,
        total_service=total_service,
        total_pendapatan=total_pendapatan,
        total_kasbon=total_kasbon,
        total_hadir=total_hadir
    )

@app.route('/api/employee/<int:id>')
@login_required
def api_employee(id):
    emp = Technician.query.get_or_404(id)
    return jsonify({
        'id': emp.id,
        'nama': emp.nama,
        'gaji_pokok': emp.gaji_pokok or 0
    })


from datetime import datetime

@app.route('/employees/add', methods=['GET', 'POST'])
@login_required
def add_employee():
    if request.method == 'POST':
        # 🔹 Konversi tanggal dari string ke Python date
        tanggal_str = request.form.get('tanggal_masuk')
        tanggal_masuk = None
        if tanggal_str:
            try:
                tanggal_masuk = datetime.strptime(tanggal_str, "%Y-%m-%d").date()
            except ValueError:
                flash("Format tanggal salah! Gunakan format YYYY-MM-DD.", "danger")
                return redirect(request.url)

        # 🔹 Konversi numerik aman
        try:
            gaji_pokok = float(request.form.get('gaji_pokok', 0.0))
        except ValueError:
            gaji_pokok = 0.0

        try:
            markup = float(request.form.get('markup', 0.28))
        except ValueError:
            markup = 0.28

        # 🔹 Buat objek baru
        t = Technician(
            nama=request.form['nama'],
            team=request.form.get('team'),
            jabatan=request.form.get('jabatan'),
            no_hp=request.form.get('no_hp'),
            alamat=request.form.get('alamat'),
            tanggal_masuk=tanggal_masuk,   # ✅ sudah date object, bukan string
            status=request.form.get('status'),
            gaji_pokok=gaji_pokok,
            markup=markup,
            catatan=request.form.get('catatan')
        )

        # 🔹 Upload foto & dokumen jika ada
        foto = request.files.get('foto')
        if foto and foto.filename:
            t.foto = save_file(foto)

        dokumen = request.files.get('dokumen')
        if dokumen and dokumen.filename:
            t.dokumen = save_file(dokumen)

        db.session.add(t)
        db.session.commit()
        flash('Karyawan berhasil ditambahkan', 'success')
        return redirect(url_for('employees'))

    return render_template('employee_form.html', mode='add', emp=None)




UPLOAD_FOLDER = os.path.join(app.root_path, 'static', 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def save_file(file_obj):
    if file_obj and file_obj.filename:
        filename = secure_filename(file_obj.filename)
        save_path = os.path.join(UPLOAD_FOLDER, filename)
        file_obj.save(save_path)
        return filename
    return None


# Pastikan folder upload tersedia
UPLOAD_FOLDER = os.path.join(app.root_path, 'static', 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


@app.route('/employees/edit/<int:id>', methods=['GET', 'POST'])
@login_required
def edit_employee(id):
    t = Technician.query.get_or_404(id)
    
    if request.method == 'POST':
        t.nama = request.form['nama']
        t.team = request.form.get('team')
        t.jabatan = request.form.get('jabatan')
        t.no_hp = request.form.get('no_hp')
        t.alamat = request.form.get('alamat')

        # 🩵 Perbaikan bagian tanggal
        tanggal_masuk_str = request.form.get('tanggal_masuk')
        if tanggal_masuk_str:
            try:
                t.tanggal_masuk = datetime.strptime(tanggal_masuk_str, "%Y-%m-%d").date()
            except ValueError:
                flash("Format tanggal salah! Gunakan format YYYY-MM-DD.", "danger")
                return redirect(request.url)
        else:
            t.tanggal_masuk = None

        t.status = request.form.get('status')

        # 🩵 Pastikan nilai numerik benar
        try:
            t.gaji_pokok = float(request.form.get('gaji_pokok', 0.0))
        except ValueError:
            t.gaji_pokok = 0.0

        try:
            t.markup = float(request.form.get('markup', 0.28))
        except ValueError:
            t.markup = 0.28

        t.catatan = request.form.get('catatan')

        # 🔹 Upload Foto (jika ada)
        foto = request.files.get('foto')
        if foto and foto.filename:
            filename_foto = secure_filename(foto.filename)
            foto.save(os.path.join(app.config['UPLOAD_FOLDER'], filename_foto))
            t.foto = filename_foto  # pastikan ada kolom `foto` di model Technician

        # 🔹 Upload Dokumen (jika ada)
        dokumen = request.files.get('dokumen')
        if dokumen and dokumen.filename:
            filename_dok = secure_filename(dokumen.filename)
            dokumen.save(os.path.join(app.config['UPLOAD_FOLDER'], filename_dok))
            t.dokumen = filename_dok  # pastikan ada kolom `dokumen` di model Technician

        db.session.commit()
        flash('✅ Data karyawan berhasil diperbarui.', 'success')
        return redirect(url_for('employees'))

    return render_template('employee_form.html', mode='edit', emp=t)

@app.route('/employees/delete/<int:id>')
@login_required
def delete_employee(id):

    t = Technician.query.get_or_404(id)

    if Attendance.query.filter_by(technician_id=id).count():
        flash('Teknisi masih memiliki data absensi', 'danger')
        return redirect(url_for('employees'))

    if Payroll.query.filter_by(employee_id=id).count():
        flash('Teknisi masih memiliki data payroll', 'danger')
        return redirect(url_for('employees'))

    db.session.delete(t)
    db.session.commit()

    flash('Teknisi berhasil dihapus', 'success')
    return redirect(url_for('employees'))




# Settings CRUD
# ---------------------------
# Settings CRUD (ServiceType, Package, Technician)
# ---------------------------

@app.route('/settings')
@login_required
def settings():
    kategori = request.args.get("kategori", "")
    q = request.args.get("q", "")

    service_types = ServiceType.query.all()
    packages = Package.query.all()

    # gabungkan
    pricelist = []
    for s in service_types:
        pricelist.append({
            "id": s.id,
            "nama": s.nama,
            "harga": s.harga,
            "kategori": "Service"
        })
    for p in packages:
        pricelist.append({
            "id": p.id,
            "nama": p.nama,
            "harga": p.harga,
            "kategori": "Package"
        })

    # filter kategori
    if kategori and kategori != "Semua":
        pricelist = [i for i in pricelist if i["kategori"] == kategori]

    # filter pencarian nama
    if q:
        pricelist = [i for i in pricelist if q.lower() in i["nama"].lower()]

    technicians = Technician.query.all()

    return render_template(
        "settings_modern.html",
        pricelist=pricelist,
        technicians=technicians,
        kategori=kategori,
        q=q
    )



# ---------------------------
# ServiceType CRUD
# ---------------------------
@app.route('/settings/service_type/new', methods=['POST'])
@login_required
def service_type_new():
    st = ServiceType(
        nama=request.form.get('nama'),
        harga=float(request.form.get('harga') or 0)
    )
    db.session.add(st)
    db.session.commit()
    flash('✅ Service type ditambahkan', 'success')
    return redirect(url_for('settings'))

@app.route('/settings/service_type/<int:id>/update', methods=['POST'])
@login_required
def service_type_update(id):
    s = ServiceType.query.get_or_404(id)
    s.nama = request.form['nama']
    s.harga = float(request.form['harga'] or 0)
    db.session.commit()
    flash("✅ Service type diperbarui", "success")
    return redirect(url_for('settings'))

@app.route('/settings/service_type/<int:id>/delete', methods=['POST'])
@login_required
def service_type_delete(id):
    s = ServiceType.query.get_or_404(id)
    db.session.delete(s)
    db.session.commit()
    flash('🗑️ Service type dihapus', 'success')
    return redirect(url_for('settings'))

# ---------------------------
# Package CRUD
# ---------------------------
@app.route('/settings/package/new', methods=['POST'])
@login_required
def package_new():
    p = Package(
        nama=request.form.get('nama'),
        harga=float(request.form.get('harga') or 0)
    )
    db.session.add(p)
    db.session.commit()
    flash('✅ Package ditambahkan', 'success')
    return redirect(url_for('settings'))

@app.route('/settings/package/<int:id>/update', methods=['POST'])
@login_required
def package_update(id):
    p = Package.query.get_or_404(id)
    p.nama = request.form['nama']
    p.harga = float(request.form['harga'] or 0)
    db.session.commit()
    flash("✅ Package diperbarui", "success")
    return redirect(url_for('settings'))

@app.route('/settings/package/<int:id>/delete', methods=['POST'])
@login_required
def package_delete(id):
    p = Package.query.get_or_404(id)
    db.session.delete(p)
    db.session.commit()
    flash('🗑️ Package dihapus', 'success')
    return redirect(url_for('settings'))

# ========================
# Pricelist (Service + Package)
# ========================

@app.route("/settings/pricelist/new", methods=["POST"])
@login_required
def pricelist_new():
    nama = request.form.get("nama")
    kategori = request.form.get("kategori")  # "Service" atau "Package"
    harga = float(request.form.get("harga") or 0)

    if kategori == "Service":
        s = ServiceType(nama=nama, harga=harga)
        db.session.add(s)
    elif kategori == "Package":
        p = Package(nama=nama, harga=harga)
        db.session.add(p)

    db.session.commit()
    flash("✅ Item baru ditambahkan", "success")
    return redirect(url_for("settings"))


@app.route('/settings/pricelist/update/<kategori>/<int:id>', methods=['POST'])
@login_required
def pricelist_update(kategori, id):
    nama = request.form.get('nama')
    harga = request.form.get('harga')

    if kategori == "Service":
        item = ServiceType.query.get_or_404(id)
    else:
        item = Package.query.get_or_404(id)

    item.nama = nama
    item.harga = float(harga or 0)

    db.session.commit()
    flash(f"{kategori} berhasil diperbarui", "success")
    return redirect(url_for('settings'))



@app.route('/settings/pricelist/delete/<string:kategori>/<int:id>', methods=['POST'])
@login_required
def pricelist_delete(kategori, id):
    if kategori == "Service":
        item = ServiceType.query.get_or_404(id)
    else:
        item = Package.query.get_or_404(id)

    db.session.delete(item)
    db.session.commit()
    flash(f"{kategori} dihapus", "success")
    return redirect(url_for('settings'))




# ---------------------------
# Technician CRUD
# ---------------------------
@app.route('/settings/technician/new', methods=['POST'])
@login_required
def tech_new():
    t = Technician(
        nama=request.form.get('nama'),
        team=request.form.get('team'),
        markup=float(request.form.get('markup') or 0.0)
    )
    db.session.add(t)
    db.session.commit()
    flash('✅ Teknisi ditambahkan', 'success')
    return redirect(url_for('settings'))

@app.route('/settings/technician/<int:id>/update', methods=['POST'])
@login_required
def tech_update(id):
    t = Technician.query.get_or_404(id)
    t.nama = request.form['nama']
    t.team = request.form['team']
    t.markup = float(request.form.get('markup') or 0.0)
    db.session.commit()
    flash('✅ Teknisi diperbarui', 'success')
    return redirect(url_for('settings'))

@app.route('/settings/technician/<int:id>/delete', methods=['POST'])
@login_required
def tech_delete(id):
    t = Technician.query.get_or_404(id)
    db.session.delete(t)
    db.session.commit()
    flash('🗑️ Teknisi dihapus', 'success')
    return redirect(url_for('settings'))




##invoice
@app.route('/invoice/generate/<int:cid>', methods=['POST'])
@login_required
def invoice_generate(cid):
    customer = Customer.query.get_or_404(cid)
    
    # Ambil daftar ID service yang dicentang
    service_ids = request.form.getlist("services")
    if not service_ids:
        flash("❌ Pilih minimal 1 service untuk membuat invoice", "danger")
        return redirect(url_for('units', cid=cid))

    services = Service.query.filter(Service.id.in_(service_ids)).all()

    # Buat invoice baru
    invoice = Invoice(
        customer_id=customer.id,
        tanggal=date.today(),
        total=sum(s.harga_satuan for s in services)
    )
    db.session.add(invoice)
    db.session.flush()  # untuk dapatkan invoice.id

    # Tambahkan detail invoice
    for s in services:
        item = InvoiceItem(
            invoice_id=invoice.id,
            service_id=s.id,
            deskripsi=f"{s.jenis_service} {s.paket_cuci or ''}".strip(),
            harga=s.harga_satuan
        )
        db.session.add(item)

    db.session.commit()
    flash("✅ Invoice berhasil dibuat", "success")

    return redirect(url_for('invoice_detail', id=invoice.id))

@app.route('/service/invoice/<int:id>')
@login_required
def service_invoice(id):
    service = Service.query.get_or_404(id)
    return render_template('invoice/service_invoice.html', service=service)

@app.template_filter("percent")
def percent(value):
    try:
        return f"{int(float(value))}%"
    except Exception:
        return "0%"

from collections import defaultdict
import numpy as np

@app.route("/analytics/cuci")
@login_required
def analytics_cuci():

    services = (
        Service.query
        .filter(Service.jenis_service.ilike("%cuci%"))
        .order_by(Service.tanggal)
        .all()
    )

    if not services:
        return render_template("analytics_cuci.html")

    # ======================
    # HARI
    # ======================
    nama_hari = {
        "Monday": "Senin",
        "Tuesday": "Selasa",
        "Wednesday": "Rabu",
        "Thursday": "Kamis",
        "Friday": "Jumat",
        "Saturday": "Sabtu",
        "Sunday": "Minggu"
    }

    hari_map = defaultdict(int)

    for s in services:
        if not s.tanggal:
            continue
        hari_en = s.tanggal.strftime("%A")
        hari_id = nama_hari.get(hari_en, hari_en)
        hari_map[hari_id] += int(s.jumlah or 1)

    hari_data = sorted(hari_map.items(), key=lambda x: x[1], reverse=True)

    # ======================
    # BULAN
    # ======================
    bulan_map = defaultdict(int)

    for s in services:
        if not s.tanggal:
            continue
        key = s.tanggal.strftime("%Y-%m")
        bulan_map[key] += int(s.jumlah or 1)

    bulan_data = sorted(bulan_map.items())

    # ======================
    # TAHUN
    # ======================
    tahun_map = defaultdict(int)

    for s in services:
        if not s.tanggal:
            continue
        tahun_map[s.tanggal.year] += int(s.jumlah or 1)

    tahun_data = sorted(tahun_map.items())

    # ======================
    # SUMMARY
    # ======================
    total_unit = int(sum(x[1] for x in bulan_data))
    avg_bulan = int(round(total_unit / len(bulan_data))) if bulan_data else 0

    peak_month = max(bulan_data, key=lambda x: x[1]) if bulan_data else ("-", 0)
    low_month = min(bulan_data, key=lambda x: x[1]) if bulan_data else ("-", 0)

    # ======================
    # GROWTH
    # ======================
    growth_data = []

    for i in range(1, len(bulan_data)):
        prev = bulan_data[i-1][1]
        curr = bulan_data[i][1]
        growth = round(((curr - prev) / prev) * 100, 1) if prev > 0 else 0
        growth_data.append((bulan_data[i][0], float(growth)))

    # ======================
    # TREND
    # ======================
    trend = "Stabil"

    if len(bulan_data) >= 2:
        if bulan_data[-1][1] > bulan_data[-2][1]:
            trend = "Naik"
        elif bulan_data[-1][1] < bulan_data[-2][1]:
            trend = "Turun"

    # ======================
    # FORECAST 6 BULAN (AMAN JSON)
    # ======================
    forecast_values = []

    if len(bulan_data) >= 3:
        y = np.array([x[1] for x in bulan_data], dtype=float)
        x = np.arange(len(y))

        slope, intercept = np.polyfit(x, y, 1)

        future_x = np.arange(len(y), len(y) + 6)
        raw_forecast = slope * future_x + intercept

        forecast_values = [int(round(v)) for v in raw_forecast]

    # ======================
    # INTERNAL VS EKSTERNAL (ANTI ERROR RELASI NULL)
    # ======================
    internal = 0
    external = 0

    for s in services:
        jumlah = int(s.jumlah or 1)

        if (
            hasattr(s, "ac_unit")
            and s.ac_unit
            and hasattr(s.ac_unit, "customer")
            and s.ac_unit.customer
            and s.ac_unit.customer.team == "Internal"
        ):
            internal += jumlah
        else:
            external += jumlah

    # ======================
    # OMSET & PROFIT
    # ======================
    total_omset = int(sum(float(s.total_harga or 0) for s in services))
    estimasi_profit = int(round(total_omset * 0.3))

    # ======================
    # INSIGHT
    # ======================
    insight = []

    insight.append(f"Total service cuci: {total_unit} unit.")

    if trend == "Naik":
        insight.append("Trend sedang meningkat.")
    elif trend == "Turun":
        insight.append("Trend sedang menurun.")
    else:
        insight.append("Trend relatif stabil.")

    insight.append(f"Bulan teramai: {peak_month[0]} ({peak_month[1]} unit).")
    insight.append(f"Bulan tersepi: {low_month[0]} ({low_month[1]} unit).")

    return render_template(
        "analytics_cuci.html",
        hari_data=hari_data,
        bulan_data=bulan_data,
        tahun_data=tahun_data,
        avg_bulan=avg_bulan,
        peak_month=peak_month,
        low_month=low_month,
        growth_data=growth_data,
        forecast_values=forecast_values,
        total_unit=total_unit,
        trend=trend,
        internal=internal,
        external=external,
        total_omset=total_omset,
        estimasi_profit=estimasi_profit,
        insight=insight
    )


# static download helper
@app.route('/download/<path:filename>')
@login_required
def download_file(filename):
    return send_from_directory(os.path.abspath('.'), filename, as_attachment=True)

# Config helpers
def load_config():
    try:
        with open('config.json', 'r', encoding='utf-8') as cf:
            return json.load(cf)
    except Exception:
        return {}

def save_config(data):
    with open('config.json', 'w', encoding='utf-8') as cf:
        json.dump(data, cf, indent=2)

# Seed data
def seed_defaults():
    if not ServiceType.query.first():
        db.session.add_all([
            ServiceType(nama='Cuci AC', harga=75000),
            ServiceType(nama='Service Besar', harga=200000),
            ServiceType(nama='Isi Freon', harga=150000),
            ServiceType(nama='Ganti Sparepart', harga=0)
        ])
    if not Package.query.first():
        db.session.add_all([
            Package(nama='Paket Hemat', harga=65000),
            Package(nama='Paket Standar', harga=90000),
            Package(nama='Paket Premium', harga=120000)
        ])
    if not Technician.query.first():
        db.session.add_all([
            Technician(nama='Budi', team='Internal'),
            Technician(nama='Andi', team='Eksternal'),
            Technician(nama='Sari', team='Internal')
        ])
    db.session.commit()

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        if not User.query.filter_by(username='admin').first():
            u = User(username='admin'); u.set_password('admin'); db.session.add(u)
        seed_defaults()
        print(app.url_map)

    app.run(debug=True)