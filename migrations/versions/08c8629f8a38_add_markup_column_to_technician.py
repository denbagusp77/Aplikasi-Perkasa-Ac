"""Add markup column to technician

Revision ID: 08c8629f8a38
Revises: 
Create Date: 2025-08-29 00:51:12.118766

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '08c8629f8a38'
down_revision = None
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('technician', schema=None) as batch_op:
        batch_op.add_column(sa.Column('markup', sa.Float(), nullable=True, server_default="0"))

    # ### end Alembic commands ###


def downgrade():
    with op.batch_alter_table('technician', schema=None) as batch_op:
        batch_op.add_column(sa.Column('markup', sa.Float(), nullable=True))

