"""add keterangan_reminder and service_id to reminder

Revision ID: <new_revision_id>  # Ganti dengan revision ID baru, misal: '1a2b3c4d5e6f'
Revises: 0ff3767b5947
Create Date: 2025-09-09 23:29:36.030134
"""

from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = '<new_revision_id>'  # Ganti ini dengan ID migrasi baru
down_revision = '0ff3767b5947'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('reminder', schema=None) as batch_op:
        # Tambah kolom keterangan_reminder jika belum ada
        batch_op.add_column(
            sa.Column('keterangan_reminder', sa.String(length=50), nullable=False, server_default='Belum di-Reminder')
        )
        # Tambah kolom service_id
        batch_op.add_column(sa.Column('service_id', sa.Integer, nullable=True))
        # Buat foreign key dengan nama constraint eksplisit
        batch_op.create_foreign_key(
            'fk_reminder_service_id',  # Nama constraint FK
            'service',                # Tabel referensi
            ['service_id'],           # Kolom di reminder
            ['id']                   # Kolom di service
        )


def downgrade():
    with op.batch_alter_table('reminder', schema=None) as batch_op:
        # Drop FK constraint dulu
        batch_op.drop_constraint('fk_reminder_service_id', type_='foreignkey')
        # Baru drop kolom
        batch_op.drop_column('service_id')
        batch_op.drop_column('keterangan_reminder')
