"""Add ac_unit_id to service

Revision ID: 7e07cc190359
Revises: 465a2e15558b
Create Date: 2025-10-31 22:55:15.549418

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '7e07cc190359'
down_revision = '465a2e15558b'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('service') as batch_op:
        batch_op.add_column(sa.Column('ac_unit_id', sa.Integer(), nullable=True))
        batch_op.create_foreign_key(
            'fk_service_ac_unit',  # 🔥 beri nama constraint
            'ac_unit',
            ['ac_unit_id'],
            ['id']
        )


def downgrade():
    with op.batch_alter_table('service') as batch_op:
        batch_op.drop_constraint('fk_service_ac_unit', type_='foreignkey')
        batch_op.drop_column('ac_unit_id')
