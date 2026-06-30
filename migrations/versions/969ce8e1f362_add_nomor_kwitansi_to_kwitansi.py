"""add nomor_kwitansi to kwitansi

Revision ID: 969ce8e1f362
Revises: f4e59549813c
Create Date: 2026-01-21 13:33:50.212463
"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '969ce8e1f362'
down_revision = 'f4e59549813c'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('kwitansi', schema=None) as batch_op:
        batch_op.add_column(
            sa.Column('nomor_kwitansi', sa.String(length=100), nullable=True)
        )


def downgrade():
    with op.batch_alter_table('kwitansi', schema=None) as batch_op:
        batch_op.drop_column('nomor_kwitansi')
