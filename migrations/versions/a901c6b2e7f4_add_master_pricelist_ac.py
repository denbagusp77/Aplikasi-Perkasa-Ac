"""add master pricelist AC domain

Revision ID: a901c6b2e7f4
Revises: f4e59549813c
Create Date: 2026-07-28 00:00:00.000000
"""
from alembic import op
import sqlalchemy as sa


revision = 'a901c6b2e7f4'
down_revision = 'f4e59549813c'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table('ac_brand',
        sa.Column('id', sa.Integer(), primary_key=True), sa.Column('nama', sa.String(100), nullable=False, unique=True),
        sa.Column('logo', sa.String(255)), sa.Column('negara_asal', sa.String(100)), sa.Column('website', sa.String(255)),
        sa.Column('garansi', sa.String(255)), sa.Column('status', sa.String(20), nullable=False, server_default='Aktif'))
    op.create_table('ac_product',
        sa.Column('id', sa.Integer(), primary_key=True), sa.Column('brand_id', sa.Integer(), sa.ForeignKey('ac_brand.id'), nullable=False),
        sa.Column('seri', sa.String(100)), sa.Column('model', sa.String(150), nullable=False, unique=True), sa.Column('pk', sa.Float()),
        sa.Column('jenis', sa.String(50), nullable=False, server_default='Standard'), sa.Column('refrigerant', sa.String(50)),
        sa.Column('tegangan_nominal', sa.String(50)), sa.Column('frekuensi_hz', sa.Float()), sa.Column('kapasitas_btu', sa.Float()),
        sa.Column('daya_watt', sa.Float()), sa.Column('arus_ampere', sa.Float()), sa.Column('eer', sa.Float()), sa.Column('cop', sa.Float()),
        sa.Column('pipa_liquid', sa.String(30)), sa.Column('pipa_gas', sa.String(30)), sa.Column('panjang_pipa_maksimum', sa.Float()),
        sa.Column('beda_tinggi_maksimum', sa.Float()), sa.Column('berat_indoor', sa.Float()), sa.Column('berat_outdoor', sa.Float()),
        sa.Column('dimensi_indoor', sa.String(100)), sa.Column('dimensi_outdoor', sa.String(100)), sa.Column('warna', sa.String(50)),
        sa.Column('made_in', sa.String(100)), sa.Column('garansi_kompresor', sa.String(100)), sa.Column('garansi_sparepart', sa.String(100)),
        sa.Column('harga_modal', sa.Float(), nullable=False, server_default='0'), sa.Column('harga_distributor', sa.Float(), nullable=False, server_default='0'),
        sa.Column('harga_dealer', sa.Float(), nullable=False, server_default='0'), sa.Column('harga_jual', sa.Float(), nullable=False, server_default='0'),
        sa.Column('status', sa.String(20), nullable=False, server_default='Aktif'), sa.Column('brosur_pdf', sa.String(255)),
        sa.Column('manual_book', sa.String(255)), sa.Column('foto_indoor', sa.String(255)), sa.Column('foto_outdoor', sa.String(255)),
        sa.Column('foto_produk', sa.String(255)), sa.Column('created_at', sa.DateTime(), nullable=False), sa.Column('updated_at', sa.DateTime(), nullable=False))
    op.create_table('catalog_service',
        sa.Column('id', sa.Integer(), primary_key=True), sa.Column('kode', sa.String(50), nullable=False, unique=True),
        sa.Column('nama', sa.String(150), nullable=False), sa.Column('harga_modal', sa.Float(), nullable=False, server_default='0'),
        sa.Column('harga_jual', sa.Float(), nullable=False, server_default='0'), sa.Column('komisi_teknisi', sa.Float(), nullable=False, server_default='0'),
        sa.Column('estimasi_waktu_menit', sa.Integer()), sa.Column('garansi', sa.String(100)), sa.Column('status', sa.String(20), nullable=False, server_default='Aktif'))
    op.create_table('material',
        sa.Column('id', sa.Integer(), primary_key=True), sa.Column('nama', sa.String(150), nullable=False), sa.Column('kategori', sa.String(100), nullable=False),
        sa.Column('merk', sa.String(100)), sa.Column('satuan', sa.String(30), nullable=False, server_default='pcs'),
        sa.Column('harga_modal', sa.Float(), nullable=False, server_default='0'), sa.Column('harga_jual', sa.Float(), nullable=False, server_default='0'),
        sa.Column('supplier', sa.String(150)), sa.Column('stok', sa.Float(), nullable=False, server_default='0'), sa.Column('minimal_stok', sa.Float(), nullable=False, server_default='0'),
        sa.Column('lokasi_gudang', sa.String(100)), sa.Column('barcode', sa.String(100), unique=True), sa.Column('status', sa.String(20), nullable=False, server_default='Aktif'))
    op.create_table('installation_package',
        sa.Column('id', sa.Integer(), primary_key=True), sa.Column('kode', sa.String(50), nullable=False, unique=True),
        sa.Column('nama', sa.String(150), nullable=False), sa.Column('deskripsi', sa.Text()), sa.Column('garansi', sa.String(100)), sa.Column('status', sa.String(20), nullable=False, server_default='Aktif'))
    op.create_table('installation_package_item',
        sa.Column('id', sa.Integer(), primary_key=True), sa.Column('package_id', sa.Integer(), sa.ForeignKey('installation_package.id'), nullable=False),
        sa.Column('product_id', sa.Integer(), sa.ForeignKey('ac_product.id')), sa.Column('material_id', sa.Integer(), sa.ForeignKey('material.id')),
        sa.Column('service_id', sa.Integer(), sa.ForeignKey('catalog_service.id')), sa.Column('quantity', sa.Float(), nullable=False, server_default='1'))
    op.create_table('quotation',
        sa.Column('id', sa.Integer(), primary_key=True), sa.Column('nomor', sa.String(50), nullable=False, unique=True),
        sa.Column('customer_id', sa.Integer(), sa.ForeignKey('customer.id'), nullable=False), sa.Column('alamat', sa.Text()), sa.Column('pic', sa.String(150)),
        sa.Column('tanggal', sa.Date(), nullable=False), sa.Column('berlaku_sampai', sa.Date()), sa.Column('diskon', sa.Float(), nullable=False, server_default='0'),
        sa.Column('ppn_persen', sa.Float(), nullable=False, server_default='0'), sa.Column('catatan', sa.Text()), sa.Column('syarat_pembayaran', sa.Text()),
        sa.Column('garansi', sa.Text()), sa.Column('tanda_tangan', sa.String(255)), sa.Column('status', sa.String(20), nullable=False, server_default='Draft'), sa.Column('created_at', sa.DateTime(), nullable=False))
    op.create_table('quotation_item',
        sa.Column('id', sa.Integer(), primary_key=True), sa.Column('quotation_id', sa.Integer(), sa.ForeignKey('quotation.id'), nullable=False),
        sa.Column('product_id', sa.Integer(), sa.ForeignKey('ac_product.id')), sa.Column('material_id', sa.Integer(), sa.ForeignKey('material.id')),
        sa.Column('service_id', sa.Integer(), sa.ForeignKey('catalog_service.id')), sa.Column('package_id', sa.Integer(), sa.ForeignKey('installation_package.id')),
        sa.Column('jenis_item', sa.String(20), nullable=False), sa.Column('deskripsi', sa.String(255), nullable=False), sa.Column('quantity', sa.Float(), nullable=False, server_default='1'),
        sa.Column('harga_modal_snapshot', sa.Float(), nullable=False, server_default='0'), sa.Column('harga_satuan', sa.Float(), nullable=False, server_default='0'))
    op.create_table('material_estimate',
        sa.Column('id', sa.Integer(), primary_key=True), sa.Column('nomor', sa.String(50), nullable=False, unique=True),
        sa.Column('customer_id', sa.Integer(), sa.ForeignKey('customer.id')), sa.Column('quotation_id', sa.Integer(), sa.ForeignKey('quotation.id')),
        sa.Column('panjang_pipa', sa.Float(), nullable=False, server_default='0'), sa.Column('jumlah_unit', sa.Integer(), nullable=False, server_default='1'),
        sa.Column('biaya_tambahan', sa.Float(), nullable=False, server_default='0'), sa.Column('catatan', sa.Text()), sa.Column('created_at', sa.DateTime(), nullable=False))
    op.create_table('material_estimate_item',
        sa.Column('id', sa.Integer(), primary_key=True), sa.Column('estimate_id', sa.Integer(), sa.ForeignKey('material_estimate.id'), nullable=False),
        sa.Column('material_id', sa.Integer(), sa.ForeignKey('material.id')), sa.Column('service_id', sa.Integer(), sa.ForeignKey('catalog_service.id')),
        sa.Column('jenis_item', sa.String(20), nullable=False), sa.Column('deskripsi', sa.String(255), nullable=False), sa.Column('quantity', sa.Float(), nullable=False, server_default='1'),
        sa.Column('harga_modal_snapshot', sa.Float(), nullable=False, server_default='0'), sa.Column('harga_jual_snapshot', sa.Float(), nullable=False, server_default='0'))


def downgrade():
    op.drop_table('material_estimate_item')
    op.drop_table('material_estimate')
    op.drop_table('quotation_item')
    op.drop_table('quotation')
    op.drop_table('installation_package_item')
    op.drop_table('installation_package')
    op.drop_table('material')
    op.drop_table('catalog_service')
    op.drop_table('ac_product')
    op.drop_table('ac_brand')