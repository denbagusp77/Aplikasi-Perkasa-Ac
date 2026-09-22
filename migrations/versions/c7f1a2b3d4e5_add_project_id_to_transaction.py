"""add project relationship to transaction

Revision ID: c7f1a2b3d4e5
Revises: a1f4c3e796e3
Create Date: 2026-09-19

"""
from alembic import op
import sqlalchemy as sa


revision = 'c7f1a2b3d4e5'
down_revision = 'a1f4c3e796e3'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('transaction', schema=None) as batch_op:
        batch_op.add_column(sa.Column('project_id', sa.Integer(), nullable=True))
        batch_op.create_foreign_key(
            'fk_transaction_project_id', 'project', ['project_id'], ['id']
        )


def downgrade():
    with op.batch_alter_table('transaction', schema=None) as batch_op:
        batch_op.drop_constraint('fk_transaction_project_id', type_='foreignkey')
        batch_op.drop_column('project_id')