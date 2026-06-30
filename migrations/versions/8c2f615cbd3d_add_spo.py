"""add spo

Revision ID: 8c2f615cbd3d
Revises: 9fe1139897f1
Create Date: 2026-01-21 11:19:57.219031
"""
from alembic import op
import sqlalchemy as sa

revision = '8c2f615cbd3d'
down_revision = '9fe1139897f1'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('invoice', schema=None) as batch_op:
        batch_op.add_column(sa.Column('spo', sa.String(length=100), nullable=True))


def downgrade():
    with op.batch_alter_table('invoice', schema=None) as batch_op:
        batch_op.drop_column('spo')
