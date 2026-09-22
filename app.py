from flask import (
    Flask, render_template, redirect, url_for, request, flash,
    send_from_directory, make_response, Response, jsonify, send_file
)
from flask_login import (
    LoginManager, login_user, login_required,
    logout_user, current_user, UserMixin
)
from flask_migrate import Migrate
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from sqlalchemy import event, text, or_, func, extract
from sqlalchemy.exc import SQLAlchemyError
from sqlalchemy.orm import aliased, joinedload
from sqlalchemy.ext.hybrid import hybrid_property
from dateutil.relativedelta import relativedelta
from collections import defaultdict
from calendar import monthrange
from datetime import date, datetime, timedelta
from urllib.parse import quote
import io
import json
import os
import re
import requests
import pdfkit

# Excel & PDF
from openpyxl import Workbook, load_workbook
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle,
    Paragraph, Spacer, Image
)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm


# =========================
# APP CONFIG
# =========================
app = Flask(__name__)
app.config['SECRET_KEY'] = 'ac-service-secret'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///ac_service.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

COMPANY_PROFILES = {
    'perkasa': {
        'name': 'PERKASA AC',
        'tagline': 'HVAC Sales, Installation & Service',
        'address': 'Jl. Sukabangun 2, KM 6,5\nPalembang, Indonesia',
        'logo': 'logo.png',
    },
    'andi_jaya': {
        'name': 'CV ANDI JAYA',
        'tagline': 'HVAC Sales, Installation & Service',
        'address': 'Jl. Sukabangun 2, KM 6,5\nPalembang, Indonesia',
        'logo': 'logo.png',
    },
}


def get_company_profile(company_key=None):
    key = (company_key or 'perkasa').strip().lower()
    profile = COMPANY_PROFILES.get(key, COMPANY_PROFILES['perkasa'])
    return {
        'name': profile['name'],
        'tagline': profile['tagline'],
        'address': profile['address'],
        'logo': profile['logo'],
        'short_name': profile['name'],
    }


@app.context_processor
def inject_company_profile():
    company_key = request.args.get('company', 'perkasa')
    return {'company_profile': get_company_profile(company_key)}


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


class UploadedDocument(db.Model):
    __tablename__ = 'uploaded_document'

    id = db.Column(db.Integer, primary_key=True)
    nama = db.Column(db.String(200), nullable=False)
    nomor = db.Column(db.String(100))
    jenis = db.Column(db.String(100), nullable=False, default='Laporan')
    terkait_dengan = db.Column(db.String(200))
    tanggal = db.Column(db.Date, nullable=False, default=date.today)
    periode = db.Column(db.String(100))
    deskripsi = db.Column(db.Text)
    status = db.Column(db.String(50), nullable=False, default='Final')
    dibuat_oleh = db.Column(db.String(100), default='admin')
    akses = db.Column(db.String(50), nullable=False, default='Internal')
    penting = db.Column(db.Boolean, default=False)
    tag = db.Column(db.String(200))
    file_name = db.Column(db.String(255), nullable=False)
    original_name = db.Column(db.String(255), nullable=False)
    mime_type = db.Column(db.String(100))
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)


# Master Pricelist AC
class ACBrand(db.Model):
    __tablename__ = 'ac_brand'

    id = db.Column(db.Integer, primary_key=True)
    nama = db.Column(db.String(100), unique=True, nullable=False)
    logo = db.Column(db.String(255))
    negara_asal = db.Column(db.String(100))
    website = db.Column(db.String(255))
    garansi = db.Column(db.String(255))
    status = db.Column(db.String(20), nullable=False, default='Aktif')
    products = db.relationship('ACProduct', back_populates='brand', lazy=True)


class ACProduct(db.Model):
    __tablename__ = 'ac_product'

    id = db.Column(db.Integer, primary_key=True)
    brand_id = db.Column(db.Integer, db.ForeignKey('ac_brand.id'), nullable=False)
    seri = db.Column(db.String(100))
    model = db.Column(db.String(150), unique=True, nullable=False)
    pk = db.Column(db.Float)
    jenis = db.Column(db.String(50), nullable=False, default='Standard')
    refrigerant = db.Column(db.String(50))
    tegangan_nominal = db.Column(db.String(50))
    frekuensi_hz = db.Column(db.Float)
    kapasitas_btu = db.Column(db.Float)
    daya_watt = db.Column(db.Float)
    arus_ampere = db.Column(db.Float)
    eer = db.Column(db.Float)
    cop = db.Column(db.Float)
    pipa_liquid = db.Column(db.String(30))
    pipa_gas = db.Column(db.String(30))
    panjang_pipa_maksimum = db.Column(db.Float)
    beda_tinggi_maksimum = db.Column(db.Float)
    berat_indoor = db.Column(db.Float)
    berat_outdoor = db.Column(db.Float)
    dimensi_indoor = db.Column(db.String(100))
    dimensi_outdoor = db.Column(db.String(100))
    warna = db.Column(db.String(50))
    made_in = db.Column(db.String(100))
    garansi_kompresor = db.Column(db.String(100))
    garansi_sparepart = db.Column(db.String(100))
    harga_modal = db.Column(db.Float, nullable=False, default=0.0)
    harga_distributor = db.Column(db.Float, nullable=False, default=0.0)
    harga_dealer = db.Column(db.Float, nullable=False, default=0.0)
    harga_jual = db.Column(db.Float, nullable=False, default=0.0)
    status = db.Column(db.String(20), nullable=False, default='Aktif')
    brosur_pdf = db.Column(db.String(255))
    manual_book = db.Column(db.String(255))
    foto_indoor = db.Column(db.String(255))
    foto_outdoor = db.Column(db.String(255))
    foto_produk = db.Column(db.String(255))
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow, onupdate=datetime.utcnow)
    brand = db.relationship('ACBrand', back_populates='products')

    @property
    def margin_persen(self):
        if not self.harga_jual:
            return 0.0
        return round(((self.harga_jual - self.harga_modal) / self.harga_jual) * 100, 2)


class CatalogService(db.Model):
    __tablename__ = 'catalog_service'

    id = db.Column(db.Integer, primary_key=True)
    kode = db.Column(db.String(50), unique=True, nullable=False)
    nama = db.Column(db.String(150), nullable=False)
    harga_modal = db.Column(db.Float, nullable=False, default=0.0)
    harga_jual = db.Column(db.Float, nullable=False, default=0.0)
    komisi_teknisi = db.Column(db.Float, nullable=False, default=0.0)
    estimasi_waktu_menit = db.Column(db.Integer)
    garansi = db.Column(db.String(100))
    status = db.Column(db.String(20), nullable=False, default='Aktif')

    @property
    def margin_persen(self):
        if not self.harga_jual:
            return 0.0
        return round(((self.harga_jual - self.harga_modal) / self.harga_jual) * 100, 2)


class Material(db.Model):
    __tablename__ = 'material'

    id = db.Column(db.Integer, primary_key=True)
    nama = db.Column(db.String(150), nullable=False)
    kategori = db.Column(db.String(100), nullable=False)
    merk = db.Column(db.String(100))
    satuan = db.Column(db.String(30), nullable=False, default='pcs')
    harga_modal = db.Column(db.Float, nullable=False, default=0.0)
    harga_jual = db.Column(db.Float, nullable=False, default=0.0)
    supplier = db.Column(db.String(150))
    stok = db.Column(db.Float, nullable=False, default=0.0)
    minimal_stok = db.Column(db.Float, nullable=False, default=0.0)
    lokasi_gudang = db.Column(db.String(100))
    barcode = db.Column(db.String(100), unique=True)
    status = db.Column(db.String(20), nullable=False, default='Aktif')

    @property
    def margin_persen(self):
        if not self.harga_jual:
            return 0.0
        return round(((self.harga_jual - self.harga_modal) / self.harga_jual) * 100, 2)


class InstallationPackage(db.Model):
    __tablename__ = 'installation_package'

    id = db.Column(db.Integer, primary_key=True)
    kode = db.Column(db.String(50), unique=True, nullable=False)
    nama = db.Column(db.String(150), nullable=False)
    deskripsi = db.Column(db.Text)
    garansi = db.Column(db.String(100))
    status = db.Column(db.String(20), nullable=False, default='Aktif')
    items = db.relationship('InstallationPackageItem', back_populates='package', cascade='all, delete-orphan', lazy=True)

    @property
    def total_modal(self):
        return sum(item.quantity * ((item.material.harga_modal if item.material else item.service.harga_modal if item.service else 0) or 0) for item in self.items)

    @property
    def total_jual(self):
        return sum(item.quantity * ((item.material.harga_jual if item.material else item.service.harga_jual if item.service else 0) or 0) for item in self.items)

    @property
    def margin_persen(self):
        return round(((self.total_jual - self.total_modal) / self.total_jual) * 100, 2) if self.total_jual else 0.0


class InstallationPackageItem(db.Model):
    __tablename__ = 'installation_package_item'

    id = db.Column(db.Integer, primary_key=True)
    package_id = db.Column(db.Integer, db.ForeignKey('installation_package.id'), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey('ac_product.id'))
    material_id = db.Column(db.Integer, db.ForeignKey('material.id'))
    service_id = db.Column(db.Integer, db.ForeignKey('catalog_service.id'))
    quantity = db.Column(db.Float, nullable=False, default=1.0)
    package = db.relationship('InstallationPackage', back_populates='items')
    product = db.relationship('ACProduct')
    material = db.relationship('Material')
    service = db.relationship('CatalogService')


class Quotation(db.Model):
    __tablename__ = 'quotation'

    id = db.Column(db.Integer, primary_key=True)
    nomor = db.Column(db.String(50), unique=True, nullable=False)
    customer_id = db.Column(db.Integer, db.ForeignKey('customer.id'), nullable=False)
    alamat = db.Column(db.Text)
    pic = db.Column(db.String(150))
    tanggal = db.Column(db.Date, nullable=False, default=date.today)
    berlaku_sampai = db.Column(db.Date)
    diskon = db.Column(db.Float, nullable=False, default=0.0)
    ppn_persen = db.Column(db.Float, nullable=False, default=0.0)
    catatan = db.Column(db.Text)
    syarat_pembayaran = db.Column(db.Text)
    garansi = db.Column(db.Text)
    tanda_tangan = db.Column(db.String(255))
    status = db.Column(db.String(20), nullable=False, default='Draft')
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    customer = db.relationship('Customer', backref=db.backref('quotations', lazy=True))
    items = db.relationship('QuotationItem', back_populates='quotation', cascade='all, delete-orphan', lazy=True)

    @property
    def subtotal(self):
        return sum(item.quantity * item.harga_satuan for item in self.items)

    @property
    def total(self):
        after_discount = max(self.subtotal - self.diskon, 0)
        return after_discount + (after_discount * self.ppn_persen / 100)


class QuotationItem(db.Model):
    __tablename__ = 'quotation_item'

    id = db.Column(db.Integer, primary_key=True)
    quotation_id = db.Column(db.Integer, db.ForeignKey('quotation.id'), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey('ac_product.id'))
    material_id = db.Column(db.Integer, db.ForeignKey('material.id'))
    service_id = db.Column(db.Integer, db.ForeignKey('catalog_service.id'))
    package_id = db.Column(db.Integer, db.ForeignKey('installation_package.id'))
    jenis_item = db.Column(db.String(20), nullable=False)
    deskripsi = db.Column(db.String(255), nullable=False)
    quantity = db.Column(db.Float, nullable=False, default=1.0)
    harga_modal_snapshot = db.Column(db.Float, nullable=False, default=0.0)
    harga_satuan = db.Column(db.Float, nullable=False, default=0.0)
    quotation = db.relationship('Quotation', back_populates='items')
    product = db.relationship('ACProduct')
    material = db.relationship('Material')
    service = db.relationship('CatalogService')
    package = db.relationship('InstallationPackage')

    @property
    def subtotal(self):
        return self.quantity * self.harga_satuan

    @property
    def profit(self):
        return self.quantity * (self.harga_satuan - self.harga_modal_snapshot)


class MaterialEstimate(db.Model):
    __tablename__ = 'material_estimate'

    id = db.Column(db.Integer, primary_key=True)
    nomor = db.Column(db.String(50), unique=True, nullable=False)
    customer_id = db.Column(db.Integer, db.ForeignKey('customer.id'))
    quotation_id = db.Column(db.Integer, db.ForeignKey('quotation.id'))
    panjang_pipa = db.Column(db.Float, nullable=False, default=0.0)
    jumlah_unit = db.Column(db.Integer, nullable=False, default=1)
    biaya_tambahan = db.Column(db.Float, nullable=False, default=0.0)
    catatan = db.Column(db.Text)
    created_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    customer = db.relationship('Customer')
    quotation = db.relationship('Quotation')
    items = db.relationship('MaterialEstimateItem', back_populates='estimate', cascade='all, delete-orphan', lazy=True)

    @property
    def biaya_material(self):
        return sum(item.quantity * item.harga_modal_snapshot for item in self.items if item.jenis_item == 'Material')

    @property
    def biaya_jasa(self):
        return sum(item.quantity * item.harga_modal_snapshot for item in self.items if item.jenis_item == 'Jasa')

    @property
    def harga_jual(self):
        return sum(item.quantity * item.harga_jual_snapshot for item in self.items) + self.biaya_tambahan

    @property
    def profit(self):
        return self.harga_jual - self.biaya_material - self.biaya_jasa


class MaterialEstimateItem(db.Model):
    __tablename__ = 'material_estimate_item'

    id = db.Column(db.Integer, primary_key=True)
    estimate_id = db.Column(db.Integer, db.ForeignKey('material_estimate.id'), nullable=False)
    material_id = db.Column(db.Integer, db.ForeignKey('material.id'))
    service_id = db.Column(db.Integer, db.ForeignKey('catalog_service.id'))
    jenis_item = db.Column(db.String(20), nullable=False)
    deskripsi = db.Column(db.String(255), nullable=False)
    quantity = db.Column(db.Float, nullable=False, default=1.0)
    harga_modal_snapshot = db.Column(db.Float, nullable=False, default=0.0)
    harga_jual_snapshot = db.Column(db.Float, nullable=False, default=0.0)
    estimate = db.relationship('MaterialEstimate', back_populates='items')
    material = db.relationship('Material')
    service = db.relationship('CatalogService')

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


class Project(db.Model):
    __tablename__ = 'project'

    id = db.Column(db.Integer, primary_key=True)
    kode = db.Column(db.String(50), unique=True, nullable=False)
    nama = db.Column(db.String(200), nullable=False)
    tipe = db.Column(db.String(30), nullable=False, default='Customer')
    customer_id = db.Column(db.Integer, db.ForeignKey('customer.id'), nullable=True)
    vendor = db.Column(db.String(200))
    periode_mulai = db.Column(db.Date)
    periode_selesai = db.Column(db.Date)
    status = db.Column(db.String(30), nullable=False, default='Berjalan')
    catatan = db.Column(db.Text)
    customer = db.relationship('Customer', backref=db.backref('projects', lazy=True))
    transactions = db.relationship('Transaction', back_populates='project', lazy=True)


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
    project_id = db.Column(db.Integer, db.ForeignKey('project.id'), nullable=True)
    project = db.relationship('Project', back_populates='transactions')

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


REMINDER_SEND_LIMIT = 3


def reminder_send_count(reminder):
    if not reminder:
        return 0
    match = re.search(r'\((\d+)/3\)', reminder.keterangan_reminder or '')
    if match:
        return min(int(match.group(1)), REMINDER_SEND_LIMIT)
    return 1 if reminder.reminder_sent else 0



    

# Routes - auth
@app.route('/')
def index():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET','POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
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
    pending_claim_total = Claim.query.filter(func.lower(Claim.status) == 'pending').count()
    unpaid_total = len(unpaid_services)
    unpaid_amount = sum(
        (service.harga_satuan or 0) * (service.jumlah or 1)
        for service in unpaid_services
    )

    priority_items = []

    for item in reminders:
        priority_items.append({
            "tab": "reminder",
            "name": item["nama"],
            "meta": item["unit"],
            "label": item["tipe"],
            "date": item["tanggal"].strftime("%d-%m-%Y") if hasattr(item["tanggal"], "strftime") else item["tanggal"],
            "url": url_for('reminders_auto', tab='reminder'),
        })

    for service in unpaid_services[:5]:
        customer = service.ac_unit.customer if service.ac_unit and service.ac_unit.customer else None
        unit = service.ac_unit if service.ac_unit else None
        if customer:
            priority_items.append({
                "tab": "unpaid",
                "name": customer.nama,
                "meta": f"{unit.merk} ({unit.pk}) x{unit.jumlah_unit}" if unit else '-',
                "label": "Belum Lunas",
                "date": service.tanggal.strftime("%d-%m-%Y") if service.tanggal else '-',
                "url": url_for('unpaid_list', tab='unpaid'),
            })

    for customer in customers_with_claims:
        pending_claims = [c for c in customer.claims if str(c.status).lower() == 'pending']
        if pending_claims:
            priority_items.append({
                "tab": "claim",
                "name": customer.nama,
                "meta": f"{len(pending_claims)} claim pending",
                "label": "Klaim",
                "date": max((c.tanggal for c in pending_claims), default=date.today()).strftime("%d-%m-%Y") if isinstance(max((c.tanggal for c in pending_claims), default=date.today()), datetime) else date.today().strftime("%d-%m-%Y"),
                "url": url_for('claim_list_all', tab='claim'),
            })

# ===========================================
# DASHBOARD KPI
# ===========================================

    today = date.today()

    bulan = today.month
    tahun = today.year

    # Customer
    total_customer = Customer.query.count()

    # Unit
    total_unit = ACUnit.query.count()

    # Teknisi
    total_teknisi = Technician.query.count()

    # Service Hari Ini
    today_service = Service.query.filter(
        Service.tanggal == today
    ).count()

    # Pendapatan Hari Ini
    today_income = db.session.query(
        func.coalesce(func.sum(Transaction.jumlah),0)
    ).filter(
        Transaction.kategori=="Pendapatan",
        Transaction.tanggal==today
    ).scalar()

    # Beban Hari Ini
    today_expense = db.session.query(
        func.coalesce(func.sum(Transaction.jumlah),0)
    ).filter(
        Transaction.kategori=="Beban",
        Transaction.tanggal==today
    ).scalar()

    # Pendapatan Bulan Ini
    monthly_income = db.session.query(
        func.coalesce(func.sum(Transaction.jumlah),0)
    ).filter(
        extract("month",Transaction.tanggal)==bulan,
        extract("year",Transaction.tanggal)==tahun,
        Transaction.kategori=="Pendapatan"
    ).scalar()

    # Beban Bulan Ini
    monthly_expense = db.session.query(
        func.coalesce(func.sum(Transaction.jumlah),0)
    ).filter(
        extract("month",Transaction.tanggal)==bulan,
        extract("year",Transaction.tanggal)==tahun,
        Transaction.kategori=="Beban"
    ).scalar()

    profit = monthly_income-monthly_expense

    chart_income = []
    for m in range(1, 13):
        total = db.session.query(
            func.coalesce(func.sum(Transaction.jumlah), 0)
        ).filter(
            Transaction.kategori == "Pendapatan",
            extract("month", Transaction.tanggal) == m,
            extract("year", Transaction.tanggal) == tahun
        ).scalar()
        chart_income.append(float(total))


    top_teknisi = (
    db.session.query(

        Technician.nama,

        func.coalesce(func.sum(Transaction.jumlah),0).label("omzet")

    )

    .join(Transaction)

    .filter(

        Transaction.kategori=="Pendapatan",

        extract("month",Transaction.tanggal)==bulan,

        extract("year",Transaction.tanggal)==tahun

    )

    .group_by(Technician.id)

    .order_by(func.sum(Transaction.jumlah).desc())

    .limit(5)

    .all()

)
    
    top_service = (

    db.session.query(

        Transaction.jenis,

        func.count(Transaction.id)

    )

    .filter(

        Transaction.kategori=="Pendapatan"

    )

    .group_by(Transaction.jenis)

    .all()

)
    # ==========================
# TOTAL TEKNISI
# ==========================

    total_teknisi = Technician.query.count()

    total_teknisi_aktif = Technician.query.filter_by(
        status="Aktif"
    ).count()


  # ==========================
#Pendapatan Hari ini
# ==========================
    today_revenue = db.session.query(
    func.coalesce(func.sum(Transaction.jumlah), 0)
).filter(
    Transaction.kategori == "Pendapatan",
    Transaction.tanggal == date.today()
).scalar()
    
 # ==========================
#Pendapatan Bulan ini
# ==========================
    bulan = datetime.now().month
    tahun = datetime.now().year

    monthly_revenue = db.session.query(
        func.coalesce(func.sum(Transaction.jumlah),0)
    ).filter(

        extract("month", Transaction.tanggal)==bulan,
        extract("year",Transaction.tanggal)==tahun,
        Transaction.kategori=="Pendapatan"

    ).scalar()


 # ==========================
#Service Hari ini
# ==========================
    today_services = Service.query.filter(
        Service.tanggal == date.today()
    ).count()

# ==========================
#Customer Baru Bulan ini
# ==========================
    total_customer = Customer.query.count()

    # ==========================
#Grafik Omzet 12 Bulan
# ==========================
    monthly_chart = []

    for m in range(1,13):

        total = db.session.query(

            func.coalesce(func.sum(Transaction.jumlah),0)

        ).filter(

            extract("month",Transaction.tanggal)==m,
            extract("year",Transaction.tanggal)==tahun,
            Transaction.kategori=="Pendapatan"

        ).scalar()

        monthly_chart.append(float(total))
    
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
        customers=customers_with_claims,
        priority_items=priority_items,
        total_teknisi=total_teknisi,
        total_teknisi_aktif=total_teknisi_aktif,
        today_revenue=today_revenue,
        monthly_revenue=monthly_revenue,
        monthly_chart=monthly_chart,
        chart_income=chart_income,
        top_teknisi=top_teknisi,
        top_service=top_service
        ,pending_claim_total=pending_claim_total
        ,unpaid_total=unpaid_total
        ,unpaid_amount=unpaid_amount
    )


@app.route('/documents', methods=['GET', 'POST'])
@login_required
def documents():
    document_types = ['Laporan', 'Invoice', 'Kontrak', 'SPK', 'Foto', 'Lainnya']
    related_options = ['Proyek', 'Customer', 'Vendor', 'Teknisi', 'Internal']

    if request.method == 'POST':
        required_fields = ['nama_dokumen', 'jenis_dokumen']
        if not all(request.form.get(field, '').strip() for field in required_fields):
            flash('Nama dokumen dan jenis dokumen harus diisi.', 'danger')
            return redirect(url_for('documents'))

        uploaded = request.files.get('file')
        if not uploaded or not uploaded.filename:
            flash('File dokumen wajib diunggah.', 'warning')
            return redirect(url_for('documents'))

        safe_name = secure_filename(uploaded.filename)
        if not safe_name:
            flash('Nama file tidak valid.', 'danger')
            return redirect(url_for('documents'))

        doc_dir = os.path.join(app.config['UPLOAD_FOLDER'], 'documents')
        os.makedirs(doc_dir, exist_ok=True)
        file_ext = os.path.splitext(safe_name)[1].lower()
        unique_name = f"doc_{datetime.utcnow().strftime('%Y%m%d%H%M%S%f')}{file_ext}"
        file_path = os.path.join(doc_dir, unique_name)
        uploaded.save(file_path)

        tanggal_value = request.form.get('tanggal_dokumen') or date.today().isoformat()
        try:
            tanggal = datetime.strptime(tanggal_value, '%Y-%m-%d').date()
        except ValueError:
            tanggal = date.today()

        doc = UploadedDocument(
            nama=request.form.get('nama_dokumen', '').strip(),
            nomor=(request.form.get('nomor_dokumen') or '').strip() or None,
            jenis=request.form.get('jenis_dokumen') or 'Laporan',
            terkait_dengan=request.form.get('terkait_dengan') or 'Proyek',
            tanggal=tanggal,
            periode=request.form.get('periode') or '',
            deskripsi=request.form.get('deskripsi') or '',
            status=request.form.get('status') or 'Final',
            dibuat_oleh=request.form.get('dibuat_oleh') or current_user.username,
            akses=request.form.get('akses') or 'Internal',
            penting=bool(request.form.get('penting')),
            tag=request.form.get('tag') or '',
            file_name=unique_name,
            original_name=safe_name,
            mime_type=uploaded.mimetype or 'application/octet-stream'
        )
        db.session.add(doc)
        db.session.commit()
        flash('Dokumen berhasil disimpan.', 'success')
        return redirect(url_for('documents'))

    documents_list = UploadedDocument.query.order_by(UploadedDocument.created_at.desc()).all()
    return render_template(
        'documents.html',
        documents=documents_list,
        document_types=document_types,
        related_options=related_options,
        today_date=date.today().isoformat()
    )


@app.route('/documents/history')
@login_required
def documents_history():
    docs = UploadedDocument.query.order_by(UploadedDocument.tanggal.desc(), UploadedDocument.created_at.desc()).all()
    return render_template('documents_history.html', documents=docs, active='history')


@app.route('/documents/templates')
@login_required
def document_templates():
    templates = [
        {'name': 'Template Laporan Service', 'format': 'PDF', 'updated_at': '2026-09-20'},
        {'name': 'Template Invoice', 'format': 'DOCX', 'updated_at': '2026-09-20'},
        {'name': 'Template Kontrak', 'format': 'PDF', 'updated_at': '2026-09-20'},
        {'name': 'Template Rutin Cuci', 'format': 'XLSX', 'updated_at': '2026-09-20'},
    ]
    return render_template('documents_templates.html', templates=templates, active='templates')


@app.route('/documents/archive')
@login_required
def documents_archive():
    docs = UploadedDocument.query.filter(UploadedDocument.status.in_(['Draft', 'Review'])).order_by(UploadedDocument.tanggal.desc(), UploadedDocument.created_at.desc()).all()
    if not docs:
        docs = UploadedDocument.query.order_by(UploadedDocument.tanggal.desc(), UploadedDocument.created_at.desc()).limit(5).all()
    return render_template('documents_archive.html', documents=docs, active='archive')


@app.route('/documents/<int:doc_id>/download')
@login_required
def document_download(doc_id):
    doc = UploadedDocument.query.get_or_404(doc_id)
    path = os.path.join(app.config['UPLOAD_FOLDER'], 'documents', doc.file_name)
    if not os.path.exists(path):
        flash('File dokumen tidak ditemukan.', 'warning')
        return redirect(url_for('documents'))
    return send_file(path, as_attachment=True, download_name=doc.original_name, mimetype=doc.mime_type)

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
    project_id = request.args.get('project_id', type=int)
    if project_id:
        query = query.filter(Transaction.project_id == project_id)

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
        ,projects=Project.query.order_by(Project.nama).all(), project_id=project_id
    )



@app.route('/transaction/new', methods=['GET','POST'])
@login_required
def transaction_new():
    def clean_money_value(value):
        if value is None:
            return 0.0
        try:
            if isinstance(value, (int, float)):
                return float(value)
            text = str(value).strip()
            if not text:
                return 0.0
            text = re.sub(r'[^0-9,\.]', '', text)
            if not text:
                return 0.0

            if ',' in text and '.' in text:
                if text.rfind(',') > text.rfind('.'):
                    text = text.replace('.', '').replace(',', '.')
                else:
                    text = text.replace(',', '')
            elif ',' in text:
                parts = text.split(',')
                if len(parts) > 1 and len(parts[-1]) == 3 and all(part for part in parts[:-1]):
                    text = ''.join(parts)
                else:
                    text = text.replace(',', '.')
            elif '.' in text:
                parts = text.split('.')
                if len(parts) > 1 and len(parts[-1]) == 3 and all(part for part in parts[:-1]):
                    text = ''.join(parts)

            return float(text)
        except (TypeError, ValueError):
            return 0.0

    if request.method == 'POST':
        teknisi_id = request.form.get('technician_id')
        teknisi = None
        if teknisi_id:
            teknisi = Technician.query.get(int(teknisi_id))

        deskripsi = (request.form.get('deskripsi') or '').strip()

        t = Transaction(
            tanggal=datetime.strptime(request.form.get('tanggal'), "%Y-%m-%d").date(),
            kategori=request.form.get('kategori'),   # Pendapatan / Beban
            jenis=request.form.get('jenis'),
            deskripsi=deskripsi,
            jumlah=clean_money_value(request.form.get('jumlah')),
            project_id=request.form.get('project_id', type=int) or None,
            technician_id=teknisi.id if teknisi else None,
            technician_nama=teknisi.nama if teknisi else None,
            team=teknisi.team if teknisi else None
        )

        db.session.add(t)
        db.session.commit()
        flash("✅ Transaksi ditambahkan", "success")
        if t.project_id:
            return redirect(url_for('projects', project_id=t.project_id))
        return redirect(url_for('transactions'))

    # Ambil daftar teknisi yang aktif saja
    teknisi_all = Technician.query.filter_by(status='Aktif').order_by(Technician.nama).all()
    selected_project_id = request.args.get('project_id', type=int)
    return render_template(
        "transaction_form_modern.html",
        teknisi_all=teknisi_all,
        projects=Project.query.filter_by(status='Berjalan').order_by(Project.nama).all(),
        selected_project_id=selected_project_id,
    )


@app.route('/projects', methods=['GET', 'POST'])
@login_required
def projects():
    if request.method == 'POST':
        mulai = request.form.get('periode_mulai') or None
        selesai = request.form.get('periode_selesai') or None
        tipe = request.form.get('tipe') or 'Customer'
        year = datetime.strptime(mulai, '%Y-%m-%d').year if mulai else datetime.utcnow().year
        code_prefix = 'VND' if tipe.lower() == 'vendor' else 'CUS'
        code_pattern = f'{code_prefix}-{year}-'
        existing_codes = Project.query.filter(Project.kode.like(f'{code_pattern}%')).all()
        sequence = 1
        for existing in existing_codes:
            try:
                sequence = max(sequence, int(existing.kode.rsplit('-', 1)[-1]) + 1)
            except (ValueError, AttributeError):
                continue
        generated_code = f'{code_pattern}{sequence:03d}'
        project = Project(
            kode=generated_code,
            nama=request.form.get('nama', '').strip(),
            tipe=tipe,
            customer_id=request.form.get('customer_id', type=int) or None,
            vendor=request.form.get('vendor'),
            periode_mulai=datetime.strptime(mulai, '%Y-%m-%d').date() if mulai else None,
            periode_selesai=datetime.strptime(selesai, '%Y-%m-%d').date() if selesai else None,
            status=request.form.get('status') or 'Berjalan',
            catatan=request.form.get('catatan')
        )
        if not project.nama or Project.query.filter_by(kode=project.kode).first():
            flash('Nama proyek wajib diisi dan kode otomatis harus unik.', 'danger')
        else:
            db.session.add(project)
            db.session.commit()
            flash('Proyek berhasil dibuat.', 'success')
        return redirect(url_for('projects'))

    status_filter = request.args.get('status', '').strip()
    selected_id = request.args.get('project_id', type=int)
    expense_category = request.args.get('expense_category', '').strip()
    expense_query = request.args.get('expense_query', '').strip()
    expense_start = request.args.get('expense_start', '').strip()
    expense_end = request.args.get('expense_end', '').strip()
    try:
        expense_start_date = datetime.strptime(expense_start, '%Y-%m-%d').date() if expense_start else None
    except ValueError:
        expense_start_date = None
    try:
        expense_end_date = datetime.strptime(expense_end, '%Y-%m-%d').date() if expense_end else None
    except ValueError:
        expense_end_date = None

    all_projects = Project.query.order_by(Project.id.desc()).all()
    visible_projects = [project for project in all_projects if not status_filter or project.status == status_filter]
    rows = []
    for project in visible_projects:
        transactions = sorted(
            project.transactions,
            key=lambda transaction: (transaction.tanggal or date.min, transaction.id or 0),
            reverse=True,
        )
        if project.id == selected_id and any((expense_category, expense_query, expense_start_date, expense_end_date)):
            transactions = [
                transaction for transaction in transactions
                if (not expense_category or transaction.kategori == expense_category)
                and (not expense_start_date or (transaction.tanggal and transaction.tanggal >= expense_start_date))
                and (not expense_end_date or (transaction.tanggal and transaction.tanggal <= expense_end_date))
                and (not expense_query or expense_query.lower() in ' '.join([
                    transaction.jenis or '', transaction.deskripsi or ''
                ]).lower())
            ]
        income = sum(t.jumlah for t in transactions if t.kategori == 'Pendapatan')
        expense = sum(t.jumlah for t in transactions if t.kategori == 'Beban')
        paid = sum(t.jumlah for t in transactions if t.kategori == 'Pendapatan' and getattr(t, 'status', None) == 'Dibayar')
        receivable = max(income - paid, 0)
        chart_income = [0, 0, 0, 0]
        chart_expense = [0, 0, 0, 0]
        expense_categories = defaultdict(float)
        for transaction in transactions:
            bucket = min(max(((transaction.tanggal.day if transaction.tanggal else 1) - 1) // 7, 0), 3)
            if transaction.kategori == 'Pendapatan':
                chart_income[bucket] += transaction.jumlah or 0
            elif transaction.kategori == 'Beban':
                chart_expense[bucket] += transaction.jumlah or 0
                expense_categories[transaction.jenis or 'Lainnya'] += transaction.jumlah or 0
        expense_breakdown = []
        category_colors = ['#2d8cf0', '#2bb673', '#f2a900', '#8b5cf6', '#ef5b67', '#aab7c9']
        category_start = 0
        if expense:
            for index, (category, amount) in enumerate(sorted(expense_categories.items(), key=lambda item: item[1], reverse=True)):
                category_end = 100 if index == len(expense_categories) - 1 else category_start + (amount / expense * 100)
                expense_breakdown.append({
                    'label': category,
                    'amount': amount,
                    'percent': amount / expense * 100,
                    'start': category_start,
                    'end': category_end,
                    'color': category_colors[index % len(category_colors)],
                })
                category_start = category_end
        rows.append({
            'project': project,
            'income': income,
            'expense': expense,
            'profit': income - expense,
            'paid': paid,
            'receivable': receivable,
            'transactions': transactions,
            'chart_income': chart_income,
            'chart_expense': chart_expense,
            'expense_categories': sorted(expense_categories.items(), key=lambda item: item[1], reverse=True),
            'expense_breakdown': expense_breakdown,
        })

    selected = None
    if selected_id:
        selected = next((row for row in rows if row['project'].id == selected_id), rows[0] if rows else None)
    elif rows:
        selected = rows[0]

    for row in rows:
        row['selected'] = row is selected

    return render_template(
        'projects.html',
        projects=rows,
        selected_project=selected,
        status_filter=status_filter,
        expense_category=expense_category,
        expense_query=expense_query,
        expense_start=expense_start,
        expense_end=expense_end,
        customers=Customer.query.order_by(Customer.nama).all(),
    )


@app.route('/projects/<int:id>')
@login_required
def project_detail(id):
    project = Project.query.get_or_404(id)
    transactions = Transaction.query.filter_by(project_id=id).order_by(Transaction.tanggal.desc()).all()
    income = sum(t.jumlah for t in transactions if t.kategori == 'Pendapatan')
    expense = sum(t.jumlah for t in transactions if t.kategori == 'Beban')
    return render_template('project_detail.html', project=project, transactions=transactions, income=income, expense=expense, profit=income-expense)


@app.route('/projects/<int:id>/edit', methods=['GET', 'POST'])
@login_required
def project_edit(id):
    project = Project.query.get_or_404(id)

    if request.method == 'POST':
        mulai = request.form.get('periode_mulai') or None
        selesai = request.form.get('periode_selesai') or None
        project.nama = request.form.get('nama', '').strip()
        project.tipe = request.form.get('tipe') or 'Customer'
        project.customer_id = request.form.get('customer_id', type=int) or None
        project.vendor = request.form.get('vendor', '').strip() or None
        project.periode_mulai = datetime.strptime(mulai, '%Y-%m-%d').date() if mulai else None
        project.periode_selesai = datetime.strptime(selesai, '%Y-%m-%d').date() if selesai else None
        project.status = request.form.get('status') or 'Berjalan'
        project.catatan = request.form.get('catatan', '').strip() or None

        if not project.nama:
            flash('Nama proyek wajib diisi.', 'danger')
        else:
            db.session.commit()
            flash('Proyek berhasil diperbarui.', 'success')
            return redirect(url_for('projects', project_id=project.id))

    return render_template(
        'project_edit.html',
        project=project,
        customers=Customer.query.order_by(Customer.nama).all(),
    )


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
    project_id = request.args.get('project_id', type=int)
    if project_id:
        query = query.filter(Transaction.project_id == project_id)

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
        bulan_names=bulan_names,
        projects=Project.query.order_by(Project.nama).all(), project_id=project_id
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
    active_tab = request.args.get('tab', 'claim')
    return render_template('Claim/list.html', claims=claims, customer=None, active_tab=active_tab)


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

    query = Transaction.query.filter(
        extract('year', Transaction.tanggal) == tahun,
        extract('month', Transaction.tanggal) == bulan
    )
    project_id = request.args.get('project_id', type=int)
    if project_id:
        query = query.filter(Transaction.project_id == project_id)
    trans = query.all()

    kas_masuk = sum(t.jumlah for t in trans if t.kategori == "Pendapatan")
    kas_keluar = sum(t.jumlah for t in trans if t.kategori == "Beban")
    saldo = kas_masuk - kas_keluar

    return render_template(
        "arus_kas.html",
        tahun=tahun, bulan=bulan,
        kas_masuk=kas_masuk,
        kas_keluar=kas_keluar,
        saldo=saldo,
        trans=trans, projects=Project.query.order_by(Project.nama).all(), project_id=project_id
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

        # Setelah service cuci dilakukan, jadwal berikutnya dihitung ulang
        # dari tanggal service agar tanggal reminder lama tidak terpakai.
        if (
            s.jenis_service and 'cuci' in s.jenis_service.lower()
            and s.tanggal_cuci_berikutnya <= s.tanggal
        ):
            s.tanggal_cuci_berikutnya = s.tanggal + relativedelta(months=3)

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


@app.route('/service/<int:id>/unpaid-reminder', methods=['POST'])
@login_required
def unpaid_service_reminder(id):
    service = Service.query.get_or_404(id)
    customer = service.ac_unit.customer if service.ac_unit and service.ac_unit.customer else None

    if not customer:
        return jsonify(success=False, message='Customer tidak ditemukan untuk service ini.'), 404

    nomor_wa = customer.nomor_wa
    if not nomor_wa:
        return jsonify(success=False, message='Nomor WA customer belum terdaftar.'), 400

    total = (
        service.harga_total
        if getattr(service, 'harga_total', None) is not None else
        service.harga
        if getattr(service, 'harga', None) is not None else
        (service.harga_satuan or 0) * (service.jumlah or 1)
    )

    pesan = (
        f"Halo Kak *{customer.nama}*, 👋\n\n"
        "Kami dari *PERKASA AC* ingin mengingatkan bahwa status service Anda masih *Belum Lunas*.\n\n"
        f"📌 Service: *{service.jenis_service or '-'}*\n"
        f"💰 Total Tagihan: *Rp {total:,.0f}*\n"
        f"📅 Tanggal Service: *{service.tanggal.strftime('%d-%m-%Y') if service.tanggal else '-'}*\n\n"
        "Silakan segera melakukan pembayaran agar proses administrasi dapat kami lanjutkan.\n"
        "Terima kasih atas kepercayaannya 🙏"
    )

    encoded_pesan = quote(pesan)
    wa_url = f"https://web.whatsapp.com/send?phone={nomor_wa}&text={encoded_pesan}"
    return jsonify(success=True, message='Reminder WhatsApp dibuat', wa_url=wa_url)


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

    active_tab = request.args.get('tab', 'unpaid')

    return render_template(
        'unpaid_list.html',
        unpaid_services=unpaid_services,
        technicians=technicians,
        bulan=bulan,
        tahun=tahun,
        technician_id=technician_id,
        total_belum_lunas=total_belum_lunas,
        active_tab=active_tab
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

    technicians = Technician.query.filter_by(status="Aktif", team="Internal") \
                                  .order_by(Technician.nama) \
                                  .all()

    # Hitung periode valid (bulan berjalan + 3 bulan sebelumnya)
    today = date.today()
    valid_periods = []
    
    for i in range(4):  # Bulan ini + 3 bulan sebelumnya
        target_date = today - relativedelta(months=i)
        periode_str = target_date.strftime('%Y-%m')
        periode_display = target_date.strftime('%B %Y')
        valid_periods.append({
            'value': periode_str,
            'display': periode_display
        })
    
    valid_periods.reverse()  # Urutkan dari yang paling lama

    return render_template(
        'payroll/index.html',
        payrolls=payrolls,
        technicians=technicians,
        valid_periods=valid_periods,
        current_periode=today.strftime('%Y-%m')
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

    # Validasi periode hanya bulan berjalan + 3 bulan sebelumnya
    today = date.today()
    valid_periods = []
    
    for i in range(4):  # Bulan ini + 3 bulan sebelumnya
        target_date = today - relativedelta(months=i)
        periode_str = target_date.strftime('%Y-%m')
        valid_periods.append(periode_str)
    
    if periode not in valid_periods:
        last_valid = valid_periods[0]
        first_valid = valid_periods[-1]
        flash(f"❌ Periode hanya bisa antara {first_valid} hingga {last_valid}.", "danger")
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
            rightMargin=12*mm, leftMargin=12*mm,
            topMargin=12*mm, bottomMargin=12*mm
        )
        styles = getSampleStyleSheet()
        content = []

        navy = colors.HexColor('#0f172a')
        blue = colors.HexColor('#0f4c81')
        blue_soft = colors.HexColor('#dbeafe')
        light = colors.HexColor('#f8fafc')
        line = colors.HexColor('#cbd5e1')
        text = colors.HexColor('#1f2937')
        muted = colors.HexColor('#475569')
        green = colors.HexColor('#15803d')
        red = colors.HexColor('#dc2626')

        title_style = ParagraphStyle(
            'TitleStyle', parent=styles['Title'],
            fontName='Helvetica-Bold', fontSize=18, leading=20,
            textColor=navy, alignment=1, spaceAfter=8
        )
        subtitle_style = ParagraphStyle(
            'SubTitleStyle', parent=styles['Normal'],
            fontName='Helvetica', fontSize=9.5, leading=12,
            textColor=muted
        )
        label_style = ParagraphStyle(
            'LabelStyle', parent=styles['BodyText'],
            fontName='Helvetica', fontSize=9, leading=12,
            textColor=muted
        )
        value_style = ParagraphStyle(
            'ValueStyle', parent=styles['BodyText'],
            fontName='Helvetica-Bold', fontSize=9, leading=12,
            textColor=navy
        )
        section_style = ParagraphStyle(
            'SectionStyle', parent=styles['Heading2'],
            fontName='Helvetica-Bold', fontSize=10.5, leading=14,
            textColor=navy, spaceBefore=2, spaceAfter=5
        )
        normal_cell_style = ParagraphStyle(
            'NormalCellStyle', parent=styles['BodyText'],
            fontName='Helvetica', fontSize=9, leading=11,
            textColor=text
        )
        amount_style = ParagraphStyle(
            'AmountStyle', parent=styles['BodyText'],
            fontName='Helvetica-Bold', fontSize=9, leading=11,
            textColor=navy, alignment=2
        )
        negative_amount_style = ParagraphStyle(
            'NegativeAmountStyle', parent=styles['BodyText'],
            fontName='Helvetica-Bold', fontSize=9, leading=11,
            textColor=red, alignment=2
        )
        total_label_style = ParagraphStyle(
            'TotalLabelStyle', parent=styles['BodyText'],
            fontName='Helvetica-Bold', fontSize=9, leading=11,
            textColor=navy
        )
        total_value_style = ParagraphStyle(
            'TotalValueStyle', parent=styles['BodyText'],
            fontName='Helvetica-Bold', fontSize=9, leading=11,
            textColor=navy, alignment=2
        )
        grand_total_style = ParagraphStyle(
            'GrandTotalStyle', parent=styles['Heading2'],
            fontName='Helvetica-Bold', fontSize=17, leading=18,
            textColor=green, alignment=1
        )

        # modern header + brand block
        try:
            logo = Image("static/logo.png", width=20*mm, height=20*mm)
            logo_cell = logo
        except Exception:
            logo_cell = Paragraph("", subtitle_style)

        brand_header = Table([
            [logo_cell, Paragraph("<b>PERKASA AC</b>", title_style), Paragraph("<b>Slip Gaji</b>", title_style)]
        ], colWidths=[26*mm, 70*mm, 65*mm], hAlign='LEFT')
        brand_header.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), blue_soft),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('LEFTPADDING', (0,0), (-1,-1), 8),
            ('RIGHTPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
            ('GRID', (0,0), (-1,-1), 0.5, blue_soft),
        ]))
        content.append(brand_header)
        content.append(Spacer(1, 10))

        employee_rows = [
            [Paragraph("Nama / NIK", label_style), Paragraph(":", label_style), Paragraph(employee.nama or "-", value_style)],
            [Paragraph("Jabatan", label_style), Paragraph(":", label_style), Paragraph(employee.jabatan or "-", value_style)],
            [Paragraph("Tanggal Masuk", label_style), Paragraph(":", label_style), Paragraph(employee.tanggal_masuk.strftime("%d-%m-%Y") if employee.tanggal_masuk else "-", value_style)],
            [Paragraph("Periode", label_style), Paragraph(":", label_style), Paragraph(payroll.periode or "-", value_style)],
        ]
        employee_table = Table(employee_rows, colWidths=[34*mm, 7*mm, 118*mm])
        employee_table.setStyle(TableStyle([
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ]))
        content.append(employee_table)
        content.append(Spacer(1, 12))

        pendapatan_rows = [[Paragraph("<b>Pendapatan</b>", section_style), Paragraph("", section_style)]]
        for item in items:
            if item.kategori == "Pendapatan":
                pendapatan_rows.append([
                    Paragraph(item.deskripsi or "-", normal_cell_style),
                    Paragraph(f"Rp {item.jumlah:,.0f}", amount_style),
                ])
        pendapatan_rows.append([
            Paragraph("<b>Total Pendapatan</b>", total_label_style),
            Paragraph(f"Rp {total_pendapatan:,.0f}", total_value_style),
        ])

        potongan_rows = [[Paragraph("<b>Potongan</b>", section_style), Paragraph("", section_style)]]
        for item in items:
            if item.kategori == "Potongan":
                potongan_rows.append([
                    Paragraph(item.deskripsi or "-", normal_cell_style),
                    Paragraph(f"-Rp {item.jumlah:,.0f}", negative_amount_style),
                ])
        potongan_rows.append([
            Paragraph("<b>Total Potongan</b>", total_label_style),
            Paragraph(f"-Rp {total_potongan:,.0f}", negative_amount_style),
        ])

        table_pendapatan = Table(pendapatan_rows, colWidths=[62*mm, 26*mm])
        table_potongan = Table(potongan_rows, colWidths=[62*mm, 26*mm])

        for table_item in [table_pendapatan, table_potongan]:
            table_item.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), blue_soft),
                ('GRID', (0,0), (-1,-1), 0.8, line),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('ALIGN', (1,1), (-1,-1), 'RIGHT'),
                ('LEFTPADDING', (0,0), (-1,-1), 7),
                ('RIGHTPADDING', (0,0), (-1,-1), 7),
                ('TOPPADDING', (0,0), (-1,-1), 6),
                ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, light]),
            ]))

        summary_table = Table([[table_pendapatan, table_potongan]], colWidths=[88*mm, 88*mm])
        summary_table.setStyle(TableStyle([
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('SPACING', (0,0), (-1,-1), 0),
        ]))
        content.append(summary_table)
        content.append(Spacer(1, 18))

        grand_total_table = Table([
            [Paragraph("<b>Total Diterima</b>", total_label_style), Paragraph(f"<b>Rp {gaji_bersih:,.0f}</b>", grand_total_style)]
        ], colWidths=[78*mm, 76*mm])
        grand_total_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#ecfdf5')),
            ('GRID', (0,0), (-1,-1), 1.2, line),
            ('ALIGN', (1,0), (1,0), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
            ('RIGHTPADDING', (0,0), (-1,-1), 10),
            ('TOPPADDING', (0,0), (-1,-1), 10),
            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ]))
        content.append(grand_total_table)
        content.append(Spacer(1, 16))

        content.append(Paragraph(f"Dicetak pada: {datetime.now().strftime('%d/%m/%Y %H:%M')}", subtitle_style))
        content.append(Spacer(1, 10))
        content.append(Paragraph("Tanda Tangan Karyawan", subtitle_style))
        content.append(Spacer(1, 8))
        content.append(Paragraph("__________________________", subtitle_style))

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
    active_tab = request.args.get('tab', 'reminder')

    services = (
    Service.query
    .join(Service.ac_unit)
    .join(ACUnit.customer)
    .all()
)


    # Ambil hanya service cuci terakhir untuk setiap customer. Dengan begitu,
    # service lama tidak dapat menampilkan reminder setelah customer cuci lagi.
    latest_cuci = {}
    for service in services:
        if not service.jenis_service or 'cuci' not in service.jenis_service.lower():
            continue
        customer_id = service.ac_unit.customer_id
        previous = latest_cuci.get(customer_id)
        if not previous or service.tanggal > previous.tanggal:
            latest_cuci[customer_id] = service

    grouped = {}
    for s in latest_cuci.values():
        next_date = s.tanggal_cuci_berikutnya or (s.tanggal + relativedelta(months=3))
        if next_date <= s.tanggal:
            next_date = s.tanggal + relativedelta(months=3)

        delta_days = (next_date - today).days
        # Tampilkan H-14 s/d H-1 dan jadwal terlewat maksimal satu bulan.
        if delta_days > 14 or next_date < today - relativedelta(months=1):
            continue

        unit = s.ac_unit
        cust = unit.customer
        grouped[cust.id] = {
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
        if data["delta"] < 0:
            status_reminder = f"Lewat {abs(data['delta'])} hari"
        elif data["delta"] == 0:
            status_reminder = "Hari Ini"
        else:
            status_reminder = f"H-{data['delta']}"

        send_count = reminder_send_count(reminder)

        reminders.append({
            "id": s.id,
            "nama": data["nama"],
            "wa": data["wa"],
            "alamat": data["alamat"],
            "team": data["team"],
            "unit": data["unit"],
            "tanggal": data["tanggal"],
            "tipe": status_reminder,
            "keterangan": reminder.keterangan_reminder if reminder else "Belum di-Reminder",
            "reminder_sent": send_count > 0,
            "send_count": send_count
        })

    reminders = sorted(reminders, key=lambda x: x["tanggal"])
    return render_template("reminders_auto.html", reminders=reminders, current_date=today, active_tab=active_tab)


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
        send_count = reminder_send_count(reminder)
        if send_count >= REMINDER_SEND_LIMIT:
            return jsonify(
                success=False,
                message=f"Reminder sudah dikirim maksimal {REMINDER_SEND_LIMIT} kali."
            ), 429

        send_count += 1
        if reminder:
            reminder.keterangan_reminder = f"Sudah di-Reminder ({send_count}/3)"
            reminder.reminder_sent = True
        else:
            reminder = Reminder(
                service_id=id,
                keterangan_reminder=f"Sudah di-Reminder ({send_count}/3)",
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

    technicians = Technician.query.filter_by(team='Internal', status='Aktif').order_by(Technician.nama).all()

    status_order = ['Hadir', 'Izin', 'Sakit', 'Alpa']
    summary_by_technician = {}

    for tech in technicians:
        summary_by_technician[tech.id] = {
            'technician': tech,
            'Hadir': 0,
            'Izin': 0,
            'Sakit': 0,
            'Alpa': 0,
            'Total': 0,
        }

    for record in records:
        tech_id = record.technician_id
        if tech_id not in summary_by_technician:
            continue

        status_name = record.status or 'Alpa'
        if status_name not in summary_by_technician[tech_id]:
            summary_by_technician[tech_id][status_name] = 0

        summary_by_technician[tech_id][status_name] += 1
        summary_by_technician[tech_id]['Total'] += 1

    attendance_summary = []
    for tech in technicians:
        summary = summary_by_technician.get(tech.id, {
            'technician': tech,
            'Hadir': 0,
            'Izin': 0,
            'Sakit': 0,
            'Alpa': 0,
            'Total': 0,
        })
        attendance_summary.append(summary)

    return render_template(
        'attendance_report.html',
        records=records,
        month=month,
        year=year,
        technicians=technicians,
        attendance_summary=attendance_summary,
        status_order=status_order,
    )


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


def money_value(raw_value, default=0.0):
    try:
        cleaned = str(raw_value or '').replace('Rp', '').replace('.', '').replace(',', '.').strip()
        return max(float(cleaned), 0.0)
    except (TypeError, ValueError):
        return default


def selling_price_from_margin(cost, margin_percent):
    cost = float(cost or 0) if isinstance(cost, (int, float)) else money_value(cost)
    margin_percent = float(margin_percent or 0) if isinstance(margin_percent, (int, float)) else money_value(margin_percent)
    if not cost or margin_percent >= 100:
        return 0.0
    return round(cost / (1 - (margin_percent / 100)), -3)


def next_document_number(prefix, model):
    today = date.today().strftime('%Y%m')
    count = model.query.filter(model.nomor.like(f'{prefix}-{today}-%')).count() + 1
    return f'{prefix}-{today}-{count:04d}'


def seed_master_pricelist():
    brands = [
        ('Daikin', 'Jepang', 'https://www.daikin.co.id', '1 tahun sparepart, 3 tahun kompresor'),
        ('Panasonic', 'Jepang', 'https://www.panasonic.com/id', '1 tahun sparepart, 3 tahun kompresor'),
        ('LG', 'Korea Selatan', 'https://www.lg.com/id', '1 tahun sparepart, 10 tahun kompresor'),
        ('Samsung', 'Korea Selatan', 'https://www.samsung.com/id', '1 tahun sparepart, 10 tahun kompresor'),
        ('Sharp', 'Jepang', 'https://id.sharp', '1 tahun sparepart, 3 tahun kompresor'),
        ('Gree', 'Tiongkok', 'https://www.gree.id', '1 tahun sparepart, 5 tahun kompresor'),
        ('Midea', 'Tiongkok', 'https://www.midea.com/id', '1 tahun sparepart, 5 tahun kompresor'),
        ('AUX', 'Tiongkok', 'https://auxair.id', '1 tahun sparepart, 5 tahun kompresor'),
        ('TCL', 'Tiongkok', 'https://www.tcl.com/id', '1 tahun sparepart, 5 tahun kompresor'),
        ('Polytron', 'Indonesia', 'https://polytron.co.id', '1 tahun sparepart, 5 tahun kompresor'),
        ('Aqua', 'Tiongkok', 'https://aquaelektronik.com', '1 tahun sparepart, 5 tahun kompresor'),
        ('Mitsubishi Electric', 'Jepang', 'https://mitsubishielectric.co.id', '1 tahun sparepart, 3 tahun kompresor'),
        ('Mitsubishi Heavy Industries', 'Jepang', 'https://www.mhi.com', '1 tahun sparepart, 3 tahun kompresor'),
        ('Toshiba', 'Jepang', 'https://www.toshiba-lifestyle.com/id', '1 tahun sparepart, 3 tahun kompresor'),
        ('Hitachi', 'Jepang', 'https://www.hitachi-homeappliances.com/id', '1 tahun sparepart, 3 tahun kompresor'),
        ('Hisense', 'Tiongkok', 'https://www.hisense.id', '1 tahun sparepart, 5 tahun kompresor'),
        ('Changhong', 'Tiongkok', 'https://www.changhong.com', '1 tahun sparepart, 5 tahun kompresor'),
        ('Akari', 'Indonesia', 'https://akari.co.id', '1 tahun sparepart, 3 tahun kompresor'),
        ('Sanken', 'Indonesia', 'https://sanken.co.id', '1 tahun sparepart, 5 tahun kompresor'),
        ('Haier', 'Tiongkok', 'https://www.haier.com/id', '1 tahun sparepart, 10 tahun kompresor'),
        ('York', 'Amerika Serikat', 'https://www.york.com', '1 tahun sparepart, 5 tahun kompresor'),
        ('Carrier', 'Amerika Serikat', 'https://www.carrier.com', '1 tahun sparepart, 5 tahun kompresor'),
        ('Trane', 'Amerika Serikat', 'https://www.trane.com', '1 tahun sparepart, 5 tahun kompresor'),
        ('Blue Star', 'India', 'https://www.bluestarindia.com', '1 tahun sparepart, 5 tahun kompresor'),
        ('General', 'Jepang', 'https://www.general-hvac.com', '1 tahun sparepart, 3 tahun kompresor')
    ]
    for nama, negara_asal, website, garansi in brands:
        brand = ACBrand.query.filter_by(nama=nama).first()
        if not brand:
            brand = ACBrand(nama=nama)
            db.session.add(brand)
        brand.negara_asal = brand.negara_asal or negara_asal
        brand.website = brand.website or website
        brand.logo = brand.logo or f"https://logo.clearbit.com/{website.replace('https://www.', '').replace('https://', '')}"
        brand.garansi = brand.garansi or garansi
        brand.status = 'Aktif'
    db.session.flush()

    services = [
        ('JSA-PASANG', 'Pasang AC Baru Split', 175000, 300000, 75000, 180, '30 hari jasa'),
        ('JSA-BONGKAR', 'Bongkar AC Split', 50000, 100000, 30000, 60, 'Tidak ada'),
        ('JSA-BP', 'Bongkar Pasang AC Split', 225000, 400000, 100000, 240, '30 hari jasa'),
        ('JSA-RELOKASI', 'Relokasi AC Split', 275000, 500000, 125000, 300, '30 hari jasa'),
        ('JSA-VAKUM', 'Vakum Sistem AC', 40000, 100000, 30000, 45, 'Tidak ada'),
        ('JSA-FREON', 'Isi Freon R32 per PK', 110000, 225000, 60000, 90, '14 hari'),
        ('JSA-LAS', 'Las Pipa Tembaga per Titik', 45000, 100000, 30000, 45, '30 hari jasa'),
        ('JSA-BOCOR', 'Perbaikan Kebocoran AC', 175000, 350000, 100000, 180, '30 hari jasa'),
        ('JSA-CUCI', 'Cuci AC Split', 35000, 75000, 30000, 45, '7 hari'),
        ('JSA-SB', 'Service Besar AC', 125000, 275000, 75000, 150, '30 hari jasa'),
        ('JSA-OH', 'Overhaul Unit AC', 300000, 650000, 175000, 360, '30 hari jasa'),
        ('JSA-KAP', 'Penggantian Kapasitor', 40000, 100000, 30000, 45, '30 hari jasa'),
        ('JSA-FAN', 'Penggantian Fan Motor', 125000, 275000, 75000, 120, '30 hari jasa'),
        ('JSA-PCB', 'Penggantian PCB', 150000, 350000, 100000, 120, '30 hari jasa'),
        ('JSA-KOMP', 'Penggantian Kompresor', 400000, 850000, 250000, 360, '30 hari jasa'),
        ('JSA-CAS', 'Instalasi Cassette', 500000, 900000, 250000, 420, '30 hari jasa'),
        ('JSA-FS', 'Instalasi Floor Standing', 450000, 850000, 225000, 360, '30 hari jasa'),
        ('JSA-VRF', 'Instalasi VRF / VRV per Indoor', 750000, 1500000, 400000, 480, '30 hari jasa'),
        ('JSA-TEST', 'Testing dan Commissioning', 100000, 250000, 75000, 90, '7 hari')
    ]
    for kode, nama, harga_modal, harga_jual, komisi, waktu, garansi in services:
        service = CatalogService.query.filter_by(kode=kode).first()
        if not service:
            service = CatalogService(kode=kode, nama=nama)
            db.session.add(service)
        service.harga_modal = service.harga_modal or harga_modal
        service.harga_jual = service.harga_jual or harga_jual
        service.komisi_teknisi = service.komisi_teknisi or komisi
        service.estimasi_waktu_menit = service.estimasi_waktu_menit or waktu
        service.garansi = service.garansi or garansi
        service.status = 'Aktif'

    materials = [
        ('Pipa Tembaga 1/4 inch', 'Pipa Tembaga', 'Kembla', 'meter', 36000, 50000, 30), ('Pipa Tembaga 3/8 inch', 'Pipa Tembaga', 'Kembla', 'meter', 52000, 70000, 30),
        ('Pipa Tembaga 1/2 inch', 'Pipa Tembaga', 'Kembla', 'meter', 78000, 100000, 20), ('Pipa Tembaga 5/8 inch', 'Pipa Tembaga', 'Kembla', 'meter', 105000, 135000, 15),
        ('Pipa Tembaga 3/4 inch', 'Pipa Tembaga', 'Kembla', 'meter', 145000, 185000, 10), ('Pipa Insulasi 1/4 inch', 'Isolasi', 'Thermaflex', 'meter', 7000, 12000, 50),
        ('Pipa Insulasi 3/8 inch', 'Isolasi', 'Thermaflex', 'meter', 9000, 15000, 50), ('Pipa Insulasi 1/2 inch', 'Isolasi', 'Thermaflex', 'meter', 12000, 18000, 40),
        ('Kabel NYM 2 x 1.5 mm', 'Kabel', 'Supreme', 'meter', 9500, 14500, 100), ('Kabel NYM 2 x 2.5 mm', 'Kabel', 'Supreme', 'meter', 16000, 23500, 100),
        ('Kabel NYM 2 x 4 mm', 'Kabel', 'Supreme', 'meter', 25500, 36000, 50), ('Kabel NYM 3 x 1.5 mm', 'Kabel', 'Supreme', 'meter', 14000, 20000, 100),
        ('Kabel NYM 3 x 2.5 mm', 'Kabel', 'Supreme', 'meter', 23000, 32000, 100), ('Kabel NYM 3 x 4 mm', 'Kabel', 'Supreme', 'meter', 36000, 48000, 50),
        ('Kabel Interkoneksi 2 x 1.5 mm', 'Kabel Interkoneksi', 'Supreme', 'meter', 10000, 16000, 100), ('Kabel Interkoneksi 2 x 2.5 mm', 'Kabel Interkoneksi', 'Supreme', 'meter', 18500, 27000, 50),
        ('Kabel Interkoneksi 3 x 1.5 mm', 'Kabel Interkoneksi', 'Supreme', 'meter', 14500, 21500, 100), ('Kabel Interkoneksi 3 x 2.5 mm', 'Kabel Interkoneksi', 'Supreme', 'meter', 24000, 34000, 50),
        ('Kabel Interkoneksi 4 x 1.5 mm', 'Kabel Interkoneksi', 'Supreme', 'meter', 17000, 25000, 100), ('Kabel Interkoneksi 4 x 2.5 mm', 'Kabel Interkoneksi', 'Supreme', 'meter', 28500, 39000, 50),
        ('Bracket Outdoor 0.5-1 PK', 'Bracket', 'Local', 'set', 65000, 95000, 10),
        ('Bracket Outdoor 1.5-2 PK', 'Bracket', 'Local', 'set', 100000, 140000, 10), ('Bracket Cassette', 'Bracket', 'Local', 'set', 225000, 300000, 5),
        ('Pipa Drain PVC 3/4 inch', 'Drain', 'Wavin', 'meter', 7000, 12000, 50), ('Selang Drain Fleksibel', 'Drain', 'Local', 'meter', 4500, 8000, 50),
        ('Duct PVC 65 x 65 mm', 'Duct', 'Eterna', 'meter', 30000, 45000, 30), ('Duct PVC 100 x 100 mm', 'Duct', 'Eterna', 'meter', 50000, 70000, 20),
        ('MCB 6A 1P', 'MCB', 'Schneider', 'pcs', 45000, 65000, 10), ('MCB 10A 1P', 'MCB', 'Schneider', 'pcs', 50000, 75000, 10),
        ('MCB 16A 1P', 'MCB', 'Schneider', 'pcs', 65000, 90000, 10), ('Stop Kontak AC 16A', 'Stop Kontak', 'Panasonic', 'pcs', 45000, 70000, 10),
        ('Steker AC 16A', 'Steker', 'Panasonic', 'pcs', 25000, 40000, 10), ('Freon R32', 'Freon', 'Daikin', 'kg', 105000, 150000, 10),
        ('Freon R410A', 'Freon', 'Honeywell', 'kg', 140000, 190000, 10), ('Freon R22', 'Freon', 'Chemours', 'kg', 95000, 140000, 10),
        ('Conduit PVC 20 mm', 'Conduit', 'Eterna', 'meter', 8500, 14000, 50), ('Clamp Pipa 3/4 inch', 'Clamp', 'Local', 'pcs', 1000, 2000, 100),
        ('Dynabolt M8', 'Dynabolt', 'Fischer', 'pcs', 3500, 6000, 50), ('Screw dan Fischer S8', 'Screw', 'Local', 'set', 1200, 2500, 100),
        ('Duct Tape Aluminium', 'Duct Tape', 'Nitto', 'roll', 18000, 30000, 10), ('Cable Tray 100 mm', 'Cable Tray', 'Local', 'meter', 85000, 120000, 10),
        ('Kabel Ties 30 cm', 'Aksesoris', 'Local', 'pack', 12000, 20000, 10), ('Baut Roof Rack', 'Aksesoris', 'Local', 'pcs', 4000, 7000, 50),
        ('PVC Elbow 3/4 inch', 'PVC', 'Wavin', 'pcs', 3500, 6000, 50), ('PVC Socket 3/4 inch', 'PVC', 'Wavin', 'pcs', 2500, 4500, 50),
        ('Lem PVC', 'PVC', 'Rucika', 'tube', 11000, 18000, 10), ('Peredam Getar Outdoor', 'Aksesoris', 'Local', 'set', 20000, 35000, 10),
        ('Pompa Kondensat', 'Drain', 'Aspen', 'pcs', 450000, 600000, 2)
    ]
    for nama, kategori, merk, satuan, harga_modal, harga_jual, minimal_stok in materials:
        material = Material.query.filter_by(nama=nama).first()
        if not material:
            material = Material(nama=nama, kategori=kategori)
            db.session.add(material)
        material.kategori = kategori
        material.merk = merk
        material.satuan = satuan
        material.harga_modal = material.harga_modal or harga_modal
        material.harga_jual = material.harga_jual or harga_jual
        material.minimal_stok = material.minimal_stok or minimal_stok
        material.lokasi_gudang = material.lokasi_gudang or 'Gudang Utama'
        material.barcode = material.barcode or f"PAC-MAT-{re.sub(r'[^A-Z0-9]', '', nama.upper())[:70]}"
        material.status = 'Aktif'

    products = [
        ('Daikin', 'FTC', 'FTC15NV14', 0.5, 'Standard', 'R32', 5000, 370, 1.8, 4200000), ('Panasonic', 'PU', 'PU9XKH', 1.0, 'Standard', 'R32', 9000, 790, 3.7, 4700000),
        ('LG', 'DualCool', 'T10EV5', 1.0, 'Inverter', 'R32', 9000, 650, 3.1, 5400000), ('Samsung', 'WindFree', 'AR09CYFAAWKNSE', 1.0, 'Premium Inverter', 'R32', 9000, 700, 3.4, 5900000),
        ('Sharp', 'AH-A', 'AH-A9BEY', 1.0, 'Low Watt', 'R32', 9000, 690, 3.3, 4400000), ('Gree', 'F5S', 'GWC-05F5S', 0.5, 'Standard', 'R32', 5000, 350, 1.6, 3000000),
        ('Midea', 'M-Smart', 'MSAF-09CRN2', 1.0, 'Standard', 'R32', 9000, 800, 3.8, 3900000), ('AUX', 'A-Series', 'ASW-09A4', 1.0, 'Standard', 'R32', 9000, 780, 3.7, 3600000),
        ('TCL', 'Elite', 'TAC-09CSD', 1.0, 'Inverter', 'R32', 9000, 690, 3.3, 4100000), ('Polytron', 'Neuva Pro', 'PAC-09VZ', 1.0, 'Low Watt', 'R32', 9000, 660, 3.1, 4000000),
        ('Aqua', 'Turbo Cool', 'AQA-KCR9AHP', 1.0, 'Standard', 'R32', 9000, 780, 3.7, 3800000), ('Mitsubishi Electric', 'MS-JP', 'MS-JP09VF', 1.0, 'Inverter', 'R32', 9000, 710, 3.4, 7000000),
        ('Mitsubishi Heavy Industries', 'SRK', 'SRK10CRS', 1.0, 'Inverter', 'R32', 9000, 720, 3.5, 7200000), ('Toshiba', 'U2KSG', 'RAS-10U2KSG', 1.0, 'Inverter', 'R32', 9000, 690, 3.3, 5600000),
        ('Hitachi', 'Mokai', 'RAK-DJ10PH', 1.0, 'Inverter', 'R32', 9000, 700, 3.4, 6200000), ('Hisense', 'AN', 'AN09CEG', 1.0, 'Standard', 'R32', 9000, 780, 3.7, 3500000),
        ('Changhong', 'CSC', 'CSC-09NVB', 1.0, 'Standard', 'R32', 9000, 800, 3.8, 3300000), ('Akari', 'AC', 'AC-09D3LW', 1.0, 'Low Watt', 'R32', 9000, 680, 3.2, 3400000),
        ('Sanken', 'SAC', 'SAC-09DN', 1.0, 'Low Watt', 'R32', 9000, 670, 3.2, 3700000), ('Haier', 'CleanCool', 'HSU-09VQD03', 1.0, 'Inverter', 'R32', 9000, 680, 3.2, 4800000),
        ('York', 'YWM', 'YWM10J', 1.0, 'Standard', 'R32', 9000, 800, 3.8, 5000000), ('Carrier', 'XPower', '42CVUR010', 1.0, 'Inverter', 'R32', 9000, 700, 3.4, 6500000),
        ('Trane', 'TVR', 'TVR-S 1.5HP', 1.5, 'VRF', 'R410A', 12000, 1100, 5.2, 12000000), ('Blue Star', 'IC', 'IC312YATU', 1.5, 'Inverter', 'R32', 12000, 1050, 5.0, 6500000),
        ('General', 'ASHG', 'ASHG09KPCA', 1.0, 'Inverter', 'R32', 9000, 680, 3.2, 6800000)
    ]
    for brand_name, seri, model, pk, jenis, refrigerant, btu, watt, ampere, harga_modal in products:
        brand = ACBrand.query.filter_by(nama=brand_name).first()
        pipa_gas = '3/8 inch' if pk <= 1 else '1/2 inch'
        harga_jual = round(harga_modal * 1.22, -3)
        product = ACProduct.query.filter_by(model=model).first()
        if not product:
            product = ACProduct(brand_id=brand.id, model=model)
            db.session.add(product)
        product.seri = product.seri or seri
        product.pk = product.pk or pk
        product.jenis = product.jenis if product.jenis not in (None, '', 'Standard') or jenis == 'Standard' else jenis
        product.refrigerant = product.refrigerant or refrigerant
        product.tegangan_nominal = product.tegangan_nominal or '220-240 Volt'
        product.frekuensi_hz = product.frekuensi_hz or 50
        product.kapasitas_btu = product.kapasitas_btu or btu
        product.daya_watt = product.daya_watt or watt
        product.arus_ampere = product.arus_ampere or ampere
        product.eer = product.eer or round(btu / watt, 2)
        product.cop = product.cop or round((btu * 0.293071) / watt, 2)
        product.pipa_liquid = product.pipa_liquid or '1/4 inch'
        product.pipa_gas = product.pipa_gas or pipa_gas
        product.panjang_pipa_maksimum = product.panjang_pipa_maksimum or (15 if pk <= 1 else 20)
        product.beda_tinggi_maksimum = product.beda_tinggi_maksimum or (7 if pk <= 1 else 10)
        product.berat_indoor = product.berat_indoor or round(7.5 + pk * 1.5, 1)
        product.berat_outdoor = product.berat_outdoor or round(20 + pk * 5, 1)
        product.dimensi_indoor = product.dimensi_indoor or '790 x 200 x 275 mm'
        product.dimensi_outdoor = product.dimensi_outdoor or '720 x 270 x 495 mm'
        product.warna = product.warna or 'Putih'
        product.made_in = product.made_in or 'Indonesia / Tiongkok'
        product.garansi_kompresor = product.garansi_kompresor or '5 tahun'
        product.garansi_sparepart = product.garansi_sparepart or '1 tahun'
        product.harga_modal = product.harga_modal or harga_modal
        product.harga_distributor = product.harga_distributor or round(harga_modal * 1.07, -3)
        product.harga_dealer = product.harga_dealer or round(harga_modal * 1.13, -3)
        product.harga_jual = product.harga_jual or harga_jual
        product.status = 'Aktif'
    db.session.commit()


@app.route('/master-pricelist')
@login_required
def master_pricelist():
    products = ACProduct.query.order_by(ACProduct.updated_at.desc()).all()
    materials = Material.query.order_by(Material.kategori, Material.nama).all()
    services = CatalogService.query.order_by(CatalogService.nama).all()
    packages = InstallationPackage.query.order_by(InstallationPackage.nama).all()
    quotations = Quotation.query.order_by(Quotation.created_at.desc()).limit(6).all()
    brand_sales = (
        db.session.query(ACBrand.nama, func.count(QuotationItem.id))
        .select_from(ACBrand)
        .join(ACProduct, ACProduct.brand_id == ACBrand.id)
        .join(QuotationItem, QuotationItem.product_id == ACProduct.id)
        .group_by(ACBrand.nama)
        .all()
    )
    return render_template(
        'master_pricelist.html', products=products, materials=materials, services=services, packages=packages,
        brands=ACBrand.query.order_by(ACBrand.nama).all(), quotations=quotations,
        metrics={
            'products': len(products), 'brands': ACBrand.query.count(), 'materials': len(materials),
            'quotation_total': sum(quotation.total for quotation in Quotation.query.all()),
            'profit': sum(item.profit for quotation in Quotation.query.all() for item in quotation.items)
        },
        brand_sales_labels=[item[0] for item in brand_sales], brand_sales_values=[item[1] for item in brand_sales]
    )


@app.route('/master-pricelist/seed', methods=['POST'])
@login_required
def master_pricelist_seed():
    seed_master_pricelist()
    flash('Master merk, produk, jasa, dan material lengkap telah disiapkan atau dilengkapi.', 'success')
    return redirect(url_for('master_pricelist'))


@app.route('/master-pricelist/products', methods=['POST'])
@login_required
def product_create():
    model = request.form.get('model', '').strip()
    brand_id = request.form.get('brand_id', type=int)
    if not model or not brand_id:
        flash('Merk dan model produk wajib diisi.', 'danger')
        return redirect(url_for('master_pricelist'))
    if ACProduct.query.filter_by(model=model).first():
        flash('Model produk sudah terdaftar.', 'danger')
        return redirect(url_for('master_pricelist'))
    harga_modal = money_value(request.form.get('harga_modal'))
    harga_jual = money_value(request.form.get('harga_jual'))
    if request.form.get('margin_persen') and harga_modal:
        harga_jual = selling_price_from_margin(harga_modal, request.form.get('margin_persen'))
    product = ACProduct(
        brand_id=brand_id, model=model, seri=request.form.get('seri'), pk=money_value(request.form.get('pk')),
        jenis=request.form.get('jenis') or 'Standard', refrigerant=request.form.get('refrigerant'),
        tegangan_nominal=request.form.get('tegangan_nominal'), frekuensi_hz=money_value(request.form.get('frekuensi_hz')),
        kapasitas_btu=money_value(request.form.get('kapasitas_btu')), daya_watt=money_value(request.form.get('daya_watt')),
        arus_ampere=money_value(request.form.get('arus_ampere')), pipa_liquid=request.form.get('pipa_liquid'),
        pipa_gas=request.form.get('pipa_gas'), harga_modal=harga_modal,
        harga_distributor=money_value(request.form.get('harga_distributor')), harga_dealer=money_value(request.form.get('harga_dealer')),
        harga_jual=harga_jual, status=request.form.get('status') or 'Aktif'
    )
    db.session.add(product)
    db.session.commit()
    flash('Produk AC berhasil ditambahkan.', 'success')
    return redirect(url_for('master_pricelist'))


@app.route('/master-pricelist/materials', methods=['POST'])
@login_required
def material_create():
    nama = request.form.get('nama', '').strip()
    kategori = request.form.get('kategori', '').strip()
    if not nama or not kategori:
        flash('Nama dan kategori material wajib diisi.', 'danger')
        return redirect(url_for('master_pricelist'))
    barcode = request.form.get('barcode', '').strip() or None
    if barcode and Material.query.filter_by(barcode=barcode).first():
        flash('Barcode material sudah digunakan.', 'danger')
        return redirect(url_for('master_pricelist'))
    db.session.add(Material(
        nama=nama, kategori=kategori, merk=request.form.get('merk'), satuan=request.form.get('satuan') or 'pcs',
        harga_modal=money_value(request.form.get('harga_modal')), harga_jual=money_value(request.form.get('harga_jual')),
        supplier=request.form.get('supplier'), stok=money_value(request.form.get('stok')), minimal_stok=money_value(request.form.get('minimal_stok')),
        lokasi_gudang=request.form.get('lokasi_gudang'), barcode=barcode, status=request.form.get('status') or 'Aktif'
    ))
    db.session.commit()
    flash('Material berhasil ditambahkan.', 'success')
    return redirect(url_for('master_pricelist'))


@app.route('/master-pricelist/services', methods=['POST'])
@login_required
def catalog_service_create():
    kode = request.form.get('kode', '').strip().upper()
    nama = request.form.get('nama', '').strip()
    if not kode or not nama or CatalogService.query.filter_by(kode=kode).first():
        flash('Kode jasa wajib unik dan nama wajib diisi.', 'danger')
        return redirect(url_for('master_pricelist'))
    db.session.add(CatalogService(
        kode=kode, nama=nama, harga_modal=money_value(request.form.get('harga_modal')),
        harga_jual=money_value(request.form.get('harga_jual')), komisi_teknisi=money_value(request.form.get('komisi_teknisi')),
        estimasi_waktu_menit=request.form.get('estimasi_waktu_menit', type=int), garansi=request.form.get('garansi'), status='Aktif'
    ))
    db.session.commit()
    flash('Jasa berhasil ditambahkan.', 'success')
    return redirect(url_for('master_pricelist'))


@app.route('/api/master-pricelist/<string:resource>', methods=['GET'])
@login_required
def master_pricelist_api(resource):
    resources = {
        'products': ACProduct.query.filter_by(status='Aktif').all(),
        'materials': Material.query.filter_by(status='Aktif').all(),
        'services': CatalogService.query.filter_by(status='Aktif').all(),
        'packages': InstallationPackage.query.filter_by(status='Aktif').all()
    }
    if resource not in resources:
        return jsonify({'error': 'Resource tidak ditemukan'}), 404
    result = []
    for item in resources[resource]:
        result.append({
            'id': item.id, 'name': getattr(item, 'model', None) or item.nama,
            'price': item.total_jual if resource == 'packages' else getattr(item, 'harga_jual', 0),
            'cost': item.total_modal if resource == 'packages' else getattr(item, 'harga_modal', 0),
            'description': getattr(item, 'model', None) or item.nama,
            'power_watt': getattr(item, 'daya_watt', None), 'current_ampere': getattr(item, 'arus_ampere', None),
            'brand_id': item.brand_id if resource == 'products' else None,
            'brand_name': item.brand.nama if resource == 'products' else None,
            'pk': item.pk if resource == 'products' else None,
            'ac_type': item.jenis if resource == 'products' else None,
            'refrigerant': item.refrigerant if resource == 'products' else None,
            'category': item.kategori if resource == 'materials' else None
        })
    return jsonify(result)


@app.route('/master-pricelist/materials/import', methods=['POST'])
@login_required
def material_import():
    uploaded_file = request.files.get('file')
    if not uploaded_file or not uploaded_file.filename.endswith('.xlsx'):
        flash('Pilih berkas Excel .xlsx.', 'danger')
        return redirect(url_for('master_pricelist'))
    workbook = load_workbook(uploaded_file, read_only=True, data_only=True)
    worksheet = workbook.active
    headers = [str(cell.value or '').strip().lower() for cell in next(worksheet.iter_rows(min_row=1, max_row=1))]
    required_headers = {'nama', 'kategori', 'satuan'}
    if not required_headers.issubset(headers):
        flash('Header Excel minimal: nama, kategori, satuan.', 'danger')
        return redirect(url_for('master_pricelist'))
    created = 0
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        values = dict(zip(headers, row))
        nama = str(values.get('nama') or '').strip()
        kategori = str(values.get('kategori') or '').strip()
        if nama and kategori and not Material.query.filter_by(nama=nama).first():
            db.session.add(Material(nama=nama, kategori=kategori, satuan=str(values.get('satuan') or 'pcs'),
                                    harga_modal=money_value(values.get('harga_modal')), harga_jual=money_value(values.get('harga_jual'))))
            created += 1
    db.session.commit()
    flash(f'{created} material baru diimpor.', 'success')
    return redirect(url_for('master_pricelist'))


@app.route('/master-pricelist/materials/export')
@login_required
def material_export():
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = 'Material'
    worksheet.append(['Nama', 'Kategori', 'Merk', 'Satuan', 'Harga Modal', 'Harga Jual', 'Supplier', 'Stok', 'Minimal Stok', 'Lokasi Gudang', 'Barcode', 'Status'])
    for material in Material.query.order_by(Material.nama).all():
        worksheet.append([material.nama, material.kategori, material.merk, material.satuan, material.harga_modal, material.harga_jual,
                          material.supplier, material.stok, material.minimal_stok, material.lokasi_gudang, material.barcode, material.status])
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='master-material-perkasa-ac.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/master-pricelist/materials/<int:id>/qr')
@login_required
def material_qr(id):
    material = Material.query.get_or_404(id)
    payload = f'MATERIAL:{material.id}|{material.barcode or material.nama}'
    image = qrcode.make(payload)
    output = io.BytesIO()
    image.save(output, 'PNG')
    output.seek(0)
    return send_file(output, mimetype='image/png')


def electrical_load(product, quantity, voltage, operating_hours, tariff):
    total_power = (product.daya_watt or 0) * quantity
    current = total_power / voltage if voltage else 0
    recommended_mcb = next((rating for rating in [6, 10, 16, 20, 25, 32, 40, 50, 63] if rating >= current * 1.25), 63)
    cable_size = '1.5 mm2' if current <= 10 else '2.5 mm2' if current <= 16 else '4 mm2' if current <= 25 else '6 mm2'
    daily_kwh = (total_power / 1000) * operating_hours
    monthly_kwh = daily_kwh * 30
    load_percentage = (current / recommended_mcb * 100) if recommended_mcb else 0
    advice = 'Gunakan mode eco dan set suhu 24-26 C.' if product.jenis == 'Inverter' else 'Pertimbangkan AC inverter dan jadwalkan pembersihan filter rutin.'
    return {'total_power_watt': total_power, 'current_ampere': round(current, 2), 'recommended_mcb_ampere': recommended_mcb,
            'cable_size': cable_size, 'daily_kwh': round(daily_kwh, 2), 'monthly_kwh': round(monthly_kwh, 2),
            'monthly_cost': round(monthly_kwh * tariff, 2), 'annual_cost': round(monthly_kwh * tariff * 12, 2),
            'load_percentage': round(load_percentage, 1), 'advice': advice}


@app.route('/master-pricelist/load-calculator', methods=['GET', 'POST'])
@login_required
def load_calculator():
    calculation = None
    product_id = request.values.get('product_id', type=int)
    if product_id:
        product = ACProduct.query.get_or_404(product_id)
        if request.method == 'POST':
            calculation = electrical_load(product, request.form.get('quantity', type=int) or 1,
                                           money_value(request.form.get('voltage')) or 220,
                                           money_value(request.form.get('operating_hours')) or 8,
                                           money_value(request.form.get('tariff')) or 1444.7)
    return render_template('load_calculator.html', products=ACProduct.query.filter_by(status='Aktif').order_by(ACProduct.model).all(),
                           product_id=product_id, calculation=calculation)


@app.route('/api/master-pricelist/load-calculator', methods=['POST'])
@login_required
def load_calculator_api():
    data = request.get_json(silent=True) or {}
    product = ACProduct.query.get_or_404(data.get('product_id'))
    return jsonify(electrical_load(product, int(data.get('quantity', 1)), float(data.get('voltage', 220)),
                                  float(data.get('operating_hours', 8)), float(data.get('tariff', 1444.7))))


@app.route('/quotations', methods=['GET', 'POST'])
@login_required
def quotations():
    if request.method == 'POST':
        customer_id = request.form.get('customer_id', type=int)
        selected_items = request.form.getlist('catalog_item')
        selected_products = request.form.getlist('product_selection')
        selected_items.extend(f'products:{product_id}:{quantity}' for product_id, quantity in
                              (selection.split(':', 1) for selection in selected_products))
        if not customer_id or not selected_items:
            flash('Customer dan minimal satu item wajib dipilih.', 'danger')
            return redirect(url_for('quotations'))
        customer = Customer.query.get_or_404(customer_id)
        quotation = Quotation(nomor=next_document_number('SPH', Quotation), customer_id=customer.id,
                              alamat=request.form.get('alamat') or customer.alamat, pic=request.form.get('pic'),
                              berlaku_sampai=datetime.strptime(request.form['berlaku_sampai'], '%Y-%m-%d').date() if request.form.get('berlaku_sampai') else None,
                              diskon=money_value(request.form.get('diskon')), ppn_persen=money_value(request.form.get('ppn_persen')),
                              catatan=request.form.get('catatan'), syarat_pembayaran=request.form.get('syarat_pembayaran'), garansi=request.form.get('garansi'))
        db.session.add(quotation)
        for token in selected_items:
            try:
                resource, item_id, quantity = token.split(':')
                quantity = max(float(quantity), 1)
                item_id = int(item_id)
            except (TypeError, ValueError):
                continue
            model_map = {
                'products': (ACProduct, 'Produk'),
                'materials': (Material, 'Material'),
                'services': (CatalogService, 'Jasa'),
                'packages': (InstallationPackage, 'Paket')
            }
            if resource not in model_map:
                continue
            model, jenis_item = model_map[resource]
            item = db.session.get(model, item_id)
            if not item or item.status != 'Aktif':
                continue
            description = getattr(item, 'model', None) or item.nama
            if resource == 'products':
                description = f'{item.brand.nama} {item.model} - {item.pk:g} PK {item.jenis}'
            quotation.items.append(QuotationItem(
                jenis_item=jenis_item, deskripsi=description, quantity=quantity,
                product_id=item.id if resource == 'products' else None, material_id=item.id if resource == 'materials' else None,
                service_id=item.id if resource == 'services' else None,
                package_id=item.id if resource == 'packages' else None,
                harga_modal_snapshot=item.total_modal if resource == 'packages' else item.harga_modal,
                harga_satuan=item.total_jual if resource == 'packages' else item.harga_jual
            ))
        db.session.commit()
        return redirect(url_for('quotation_detail', id=quotation.id))
    return render_template('quotation_form.html', quotations=Quotation.query.order_by(Quotation.created_at.desc()).all(),
                           customers=Customer.query.order_by(Customer.nama).all())


@app.route('/quotations/<int:id>')
@login_required
def quotation_detail(id):
    company_key = request.args.get('company', 'perkasa')
    company_profile = get_company_profile(company_key)
    return render_template('quotation_print.html', quotation=Quotation.query.get_or_404(id), print_date=date.today(), company_profile=company_profile)


@app.route('/quotations/<int:id>/excel')
@login_required
def quotation_excel(id):
    quotation = Quotation.query.get_or_404(id)
    company_profile = get_company_profile(request.args.get('company', 'perkasa'))
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = 'Penawaran'
    worksheet.append([company_profile['name'], 'Surat Penawaran', quotation.nomor])
    worksheet.append(['Customer', quotation.customer.nama])
    worksheet.append([])
    worksheet.append(['Jenis', 'Deskripsi', 'Qty', 'Harga Satuan', 'Subtotal'])
    for item in quotation.items:
        worksheet.append([item.jenis_item, item.deskripsi, item.quantity, item.harga_satuan, item.subtotal])
    worksheet.append(['', '', '', 'Grand Total', quotation.total])
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    safe_company = company_profile['name'].replace(' ', '_')
    return send_file(output, as_attachment=True, download_name=f'{quotation.nomor}_{safe_company}.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/quotations/<int:id>/pdf')
@login_required
def quotation_pdf(id):
    quotation = Quotation.query.get_or_404(id)
    company_profile = get_company_profile(request.args.get('company', 'perkasa'))
    output = io.BytesIO()
    styles = getSampleStyleSheet()
    document = SimpleDocTemplate(output, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    rows = [['No', 'Jenis', 'Deskripsi', 'Qty', 'Harga', 'Jumlah']]
    for number, item in enumerate(quotation.items, start=1):
        rows.append([str(number), item.jenis_item, item.deskripsi, str(item.quantity),
                     f'Rp {item.harga_satuan:,.0f}', f'Rp {item.subtotal:,.0f}'])
    rows.extend([['', '', '', '', 'Subtotal', f'Rp {quotation.subtotal:,.0f}'],
                 ['', '', '', '', 'Diskon', f'Rp {quotation.diskon:,.0f}'],
                 ['', '', '', '', 'Grand Total', f'Rp {quotation.total:,.0f}']])
    table = Table(rows, colWidths=[28, 55, 175, 35, 80, 85])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1769aa')), ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('GRID', (0, 0), (-1, -1), 0.25, colors.HexColor('#cbd5e1')),
        ('ALIGN', (3, 1), (-1, -1), 'RIGHT'), ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#e8f1f8')), ('FONTNAME', (4, -1), (-1, -1), 'Helvetica-Bold')
    ]))
    story = [
        Paragraph(company_profile['name'], styles['Title']),
        Paragraph('SURAT PENAWARAN', styles['Heading2']),
        Paragraph(f'<b>{quotation.nomor}</b><br/>Customer: {quotation.customer.nama}<br/>Alamat: {quotation.alamat or quotation.customer.alamat or "-"}', styles['BodyText']),
        Spacer(1, 16), table, Spacer(1, 16),
        Paragraph(f'<b>Syarat Pembayaran:</b> {quotation.syarat_pembayaran or "-"}', styles['BodyText']),
        Paragraph(f'<b>Garansi:</b> {quotation.garansi or "-"}', styles['BodyText'])
    ]
    document.build(story)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f'{quotation.nomor}.pdf', mimetype='application/pdf')


@app.route('/quotations/<int:id>/word')
@login_required
def quotation_word(id):
    quotation = Quotation.query.get_or_404(id)
    company_profile = get_company_profile(request.args.get('company', 'perkasa'))
    rows = ''.join(f'<tr><td>{index}</td><td>{item.jenis_item}</td><td>{item.deskripsi}</td><td>{item.quantity}</td><td>Rp {item.harga_satuan:,.0f}</td><td>Rp {item.subtotal:,.0f}</td></tr>' for index, item in enumerate(quotation.items, start=1))
    document = f'''<html><head><meta charset="utf-8"><style>body{{font-family:Arial;color:#1f2937}}table{{border-collapse:collapse;width:100%}}th{{background:#1769aa;color:white}}th,td{{border:1px solid #cbd5e1;padding:7px;text-align:left}}</style></head><body><h1 style="color:#1769aa">{company_profile['name']}</h1><h2>SURAT PENAWARAN</h2><p><b>{quotation.nomor}</b><br>Customer: {quotation.customer.nama}<br>Alamat: {quotation.alamat or quotation.customer.alamat or '-'}</p><table><tr><th>No</th><th>Jenis</th><th>Deskripsi</th><th>Qty</th><th>Harga</th><th>Jumlah</th></tr>{rows}<tr><td colspan="5"><b>Grand Total</b></td><td><b>Rp {quotation.total:,.0f}</b></td></tr></table><p><b>Syarat Pembayaran:</b> {quotation.syarat_pembayaran or '-'}</p><p><b>Garansi:</b> {quotation.garansi or '-'}</p></body></html>'''
    return Response(document, mimetype='application/msword', headers={'Content-Disposition': f'attachment; filename={quotation.nomor}.doc'})


@app.route('/quotations/<int:id>/to-invoice', methods=['POST'])
@login_required
def quotation_to_invoice(id):
    quotation = Quotation.query.get_or_404(id)
    invoice = Invoice(customer_id=quotation.customer_id, tanggal=date.today(), total=quotation.total, spo=quotation.nomor)
    db.session.add(invoice)
    db.session.flush()
    for item in quotation.items:
        db.session.add(InvoiceItem(invoice_id=invoice.id, deskripsi=item.deskripsi, harga=item.harga_satuan))
    quotation.status = 'Dikonversi'
    db.session.commit()
    flash('Penawaran berhasil dikonversi menjadi invoice dengan harga snapshot.', 'success')
    return redirect(url_for('invoice_detail', id=invoice.id))


@app.route('/master-pricelist/packages', methods=['POST'])
@login_required
def installation_package_create():
    kode = request.form.get('kode', '').strip().upper()
    nama = request.form.get('nama', '').strip()
    if not kode or not nama or InstallationPackage.query.filter_by(kode=kode).first():
        flash('Kode paket wajib unik dan nama wajib diisi.', 'danger')
        return redirect(url_for('master_pricelist'))
    package = InstallationPackage(kode=kode, nama=nama, deskripsi=request.form.get('deskripsi'),
                                  garansi=request.form.get('garansi'), status='Aktif')
    db.session.add(package)
    for token in request.form.getlist('package_component'):
        try:
            resource, item_id, quantity = token.split(':')
            quantity = max(float(quantity), 1)
            item_id = int(item_id)
        except (TypeError, ValueError):
            continue
        if resource == 'materials':
            item = db.session.get(Material, item_id)
            if item and item.status == 'Aktif':
                package.items.append(InstallationPackageItem(material_id=item.id, quantity=quantity))
        elif resource == 'services':
            item = db.session.get(CatalogService, item_id)
            if item and item.status == 'Aktif':
                package.items.append(InstallationPackageItem(service_id=item.id, quantity=quantity))
    db.session.commit()
    flash('Paket instalasi berhasil ditambahkan.', 'success')
    return redirect(url_for('master_pricelist'))


@app.route('/api/master-pricelist/<string:resource>', methods=['POST'])
@login_required
def master_pricelist_api_create(resource):
    payload = request.get_json(silent=True) or {}
    models = {'products': ACProduct, 'materials': Material, 'services': CatalogService, 'packages': InstallationPackage}
    model = models.get(resource)
    if not model:
        return jsonify({'error': 'Resource tidak ditemukan'}), 404
    required = {'products': ('brand_id', 'model'), 'materials': ('nama', 'kategori'), 'services': ('kode', 'nama'), 'packages': ('kode', 'nama')}[resource]
    if any(not payload.get(field) for field in required):
        return jsonify({'error': f"Field wajib: {', '.join(required)}"}), 400
    try:
        item = model(**{key: value for key, value in payload.items() if hasattr(model, key)})
        db.session.add(item)
        db.session.commit()
        return jsonify({'id': item.id, 'message': 'Data dibuat'}), 201
    except Exception as error:
        db.session.rollback()
        return jsonify({'error': str(error)}), 400


@app.route('/api/master-pricelist/<string:resource>/<int:id>', methods=['PATCH', 'DELETE'])
@login_required
def master_pricelist_api_detail(resource, id):
    models = {'products': ACProduct, 'materials': Material, 'services': CatalogService, 'packages': InstallationPackage}
    model = models.get(resource)
    if not model:
        return jsonify({'error': 'Resource tidak ditemukan'}), 404
    item = db.session.get(model, id)
    if not item:
        return jsonify({'error': 'Data tidak ditemukan'}), 404
    if request.method == 'DELETE':
        item.status = 'Nonaktif'
        db.session.commit()
        return jsonify({'message': 'Data dinonaktifkan'})
    payload = request.get_json(silent=True) or {}
    if resource == 'products' and 'margin_persen' in payload:
        cost = payload.get('harga_modal', item.harga_modal)
        payload['harga_jual'] = selling_price_from_margin(cost, payload.pop('margin_persen'))
    editable = {
        'products': {'seri', 'model', 'pk', 'jenis', 'refrigerant', 'tegangan_nominal', 'kapasitas_btu', 'daya_watt', 'arus_ampere', 'harga_modal', 'harga_distributor', 'harga_dealer', 'harga_jual', 'status'},
        'materials': {'nama', 'kategori', 'merk', 'satuan', 'harga_modal', 'harga_jual', 'supplier', 'stok', 'minimal_stok', 'lokasi_gudang', 'barcode', 'status'},
        'services': {'nama', 'harga_modal', 'harga_jual', 'komisi_teknisi', 'estimasi_waktu_menit', 'garansi', 'status'},
        'packages': {'nama', 'deskripsi', 'garansi', 'status'}
    }[resource]
    for field, value in payload.items():
        if field in editable:
            setattr(item, field, value)
    db.session.commit()
    return jsonify({'id': item.id, 'message': 'Data diperbarui'})


@app.route('/material-estimates', methods=['GET', 'POST'])
@login_required
def material_estimates():
    if request.method == 'POST':
        estimate = MaterialEstimate(nomor=next_document_number('EST', MaterialEstimate),
                                    customer_id=request.form.get('customer_id', type=int),
                                    panjang_pipa=money_value(request.form.get('panjang_pipa')),
                                    jumlah_unit=request.form.get('jumlah_unit', type=int) or 1,
                                    biaya_tambahan=money_value(request.form.get('biaya_tambahan')), catatan=request.form.get('catatan'))
        db.session.add(estimate)
        for token in request.form.getlist('estimate_item'):
            resource, item_id, quantity = token.split(':')
            item = db.session.get(Material if resource == 'materials' else CatalogService, int(item_id))
            estimate.items.append(MaterialEstimateItem(
                jenis_item='Material' if resource == 'materials' else 'Jasa', deskripsi=item.nama, quantity=max(float(quantity), 1),
                harga_modal_snapshot=item.harga_modal, harga_jual_snapshot=item.harga_jual,
                material_id=item.id if resource == 'materials' else None, service_id=item.id if resource == 'services' else None
            ))
        db.session.commit()
        flash(f'Estimasi {estimate.nomor} berhasil disimpan.', 'success')
        return redirect(url_for('material_estimates'))
    return render_template('material_estimates.html', estimates=MaterialEstimate.query.order_by(MaterialEstimate.created_at.desc()).all(),
                           customers=Customer.query.order_by(Customer.nama).all())

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        if not User.query.filter_by(username='admin').first():
            u = User(username='admin'); u.set_password('admin'); db.session.add(u)
        seed_defaults()
        print(app.url_map)

    app.run(debug=True)