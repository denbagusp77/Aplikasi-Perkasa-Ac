"""add keterangan_reminder

Revision ID: 0ff3767b5947
Revises: e7be9a1ed971
Create Date: 2025-09-09 23:29:36.030134
"""

from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = '0ff3767b5947'
down_revision = 'e7be9a1ed971'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('reminder', schema=None) as batch_op:
        batch_op.add_column(
            sa.Column('keterangan_reminder', sa.String(length=50), nullable=False, server_default='Belum di-Reminder')
        )


def downgrade():
    with op.batch_alter_table('reminder', schema=None) as batch_op:
        batch_op.drop_column('keterangan_reminder')
