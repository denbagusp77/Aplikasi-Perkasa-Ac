"""add claim model

Revision ID: 7672766910da
Revises: e17366faccc9
Create Date: 2025-09-19 15:36:48.224650

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '7672766910da'
down_revision = 'e17366faccc9'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'claim',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('customer_id', sa.Integer(), nullable=False),
        sa.Column('unit_id', sa.Integer(), nullable=True),
        sa.Column('tanggal', sa.Date(), nullable=True),
        sa.Column('deskripsi', sa.Text(), nullable=False),
        sa.Column('foto', sa.String(length=255), nullable=True),
        sa.Column('status', sa.String(length=20), nullable=True, server_default='Pending'),
        sa.ForeignKeyConstraint(['customer_id'], ['customer.id'], ondelete='CASCADE'),
        sa.ForeignKeyConstraint(['unit_id'], ['ac_unit.id'], ondelete='SET NULL'),
        sa.PrimaryKeyConstraint('id')
    )


def downgrade():
    op.drop_table('claim')
