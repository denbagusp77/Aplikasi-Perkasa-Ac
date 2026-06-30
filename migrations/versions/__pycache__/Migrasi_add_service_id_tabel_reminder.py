"""add keterangan_reminder and service_id to reminder

Revision ID: <new_revision_id>
Revises: 0ff3767b5947
Create Date: 2025-09-09 23:29:36.030134
"""

from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = '<new_revision_id>'  # ganti dengan id baru, misal: '1a2b3c4d5e6f'
down_revision = '0ff3767b5947'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('reminder', schema=None) as batch_op:
        # Kolom keterangan_reminder (kalau belum ada)
        batch_op.add_column(
            sa.Column('keterangan_reminder', sa.String(length=50), nullable=False, server_default='Belum di-Reminder')
        )
        # Kolom service_id baru, nullable karena mungkin tidak selalu ada
        batch_op.add_column(
            sa.Column('service_id', sa.Integer, sa.ForeignKey('service.id'), nullable=True)
        )


def downgrade():
    with op.batch_alter_table('reminder', schema=None) as batch_op:
        batch_op.drop_column('service_id')
        batch_op.drop_column('keterangan_reminder')
