"""Add bank_account_id to invoice

Revision ID: 6aee9b5d7739
Revises: caa8be899049
Create Date: 2025-12-10 11:11:49.056142

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '6aee9b5d7739'
down_revision = 'caa8be899049'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('invoice', schema=None) as batch_op:
        # Tambah kolom
        batch_op.add_column(sa.Column('bank_account_id', sa.Integer(), nullable=True))

        # Tambah foreign key dengan nama (WAJIB di SQLite)
        batch_op.create_foreign_key(
            'fk_invoice_bank_account',     # ← beri nama
            'bank_account',                # tabel referensi
            ['bank_account_id'],           # kolom lokal
            ['id']                         # kolom referensi
        )


def downgrade():
    with op.batch_alter_table('invoice', schema=None) as batch_op:
        # Hapus constraint dengan nama yang sama
        batch_op.drop_constraint('fk_invoice_bank_account', type_='foreignkey')
        batch_op.drop_column('bank_account_id')
