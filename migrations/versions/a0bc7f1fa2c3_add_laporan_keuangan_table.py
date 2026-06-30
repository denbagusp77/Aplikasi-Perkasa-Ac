"""add laporan_keuangan table

Revision ID: a0bc7f1fa2c3
Revises: 08c8629f8a38
Create Date: 2025-09-05 23:53:53.870727

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'a0bc7f1fa2c3'
down_revision = '08c8629f8a38'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'laporan_keuangan',
        sa.Column('id', sa.Integer(), primary_key=True),
        sa.Column('bulan', sa.String(20), nullable=False),
        sa.Column('tahun', sa.Integer(), nullable=False),
        sa.Column('pendapatan_internal', sa.Integer(), default=0),
        sa.Column('pendapatan_eksternal', sa.Integer(), default=0),
        sa.Column('bonus', sa.Integer(), default=0),
        sa.Column('kasbon', sa.Integer(), default=0),
        sa.Column('created_at', sa.DateTime(), server_default=sa.func.now())
    )
    # ### end Alembic commands ###


def downgrade():
    op.drop_table('laporan_keuangan')
    # ### end Alembic commands ###
