"""add kwitansi table

Revision ID: f4e59549813c
Revises: 8c2f615cbd3d
Create Date: 2026-01-21 13:23:00.920095
"""
from alembic import op
import sqlalchemy as sa

revision = 'f4e59549813c'
down_revision = '8c2f615cbd3d'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'kwitansi',
        sa.Column('id', sa.Integer(), primary_key=True),
        sa.Column('invoice_id', sa.Integer(), nullable=False),
        sa.Column('nomor_kwitansi', sa.String(100)),
        sa.Column('diterima_dari', sa.String(200)),
        sa.Column('jumlah', sa.Integer()),
        sa.Column('terbilang', sa.Text()),
        sa.Column('untuk_pembayaran', sa.String(200)),
        sa.Column('tanggal', sa.Date()),
        sa.Column('dibuat_oleh', sa.String(100)),
    )


def downgrade():
    op.drop_table('kwitansi')
