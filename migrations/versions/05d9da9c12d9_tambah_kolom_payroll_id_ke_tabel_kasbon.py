"""Tambah kolom payroll_id ke tabel kasbon

Revision ID: 05d9da9c12d9
Revises: 6278fb216c0d
Create Date: 2025-10-26 22:57:35.912374
"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '05d9da9c12d9'
down_revision = '6278fb216c0d'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('kasbon', schema=None) as batch_op:
        batch_op.add_column(sa.Column('payroll_id', sa.Integer(), nullable=True))
        batch_op.create_foreign_key(
            'fk_kasbon_payroll_id',  # ✅ beri nama constraint
            'payroll',               # tabel referensi
            ['payroll_id'],          # kolom di tabel kasbon
            ['id'],                  # kolom di tabel payroll
            ondelete='SET NULL'      # opsional, bisa disesuaikan
        )


def downgrade():
    with op.batch_alter_table('kasbon', schema=None) as batch_op:
        batch_op.drop_constraint('fk_kasbon_payroll_id', type_='foreignkey')
        batch_op.drop_column('payroll_id')
