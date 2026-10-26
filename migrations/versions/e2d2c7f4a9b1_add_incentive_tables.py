"""Add incentive tables

Revision ID: e2d2c7f4a9b1
Revises: 01214b71508c
Create Date: 2026-09-30 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'e2d2c7f4a9b1'
down_revision = '01214b71508c'
branch_labels = None
depends_on = None


def upgrade():
    existing_tables = set(sa.inspect(op.get_bind()).get_table_names())
    if 'incentive_setting' not in existing_tables:
        op.create_table(
            'incentive_setting',
            sa.Column('id', sa.Integer(), nullable=False),
            sa.Column('technician_id', sa.Integer(), nullable=False),
            sa.Column('service_type', sa.String(length=100), nullable=True),
            sa.Column('percentage', sa.Float(), nullable=False),
            sa.Column('effective_from', sa.Date(), nullable=False),
            sa.Column('effective_to', sa.Date(), nullable=True),
            sa.Column('is_active', sa.Boolean(), nullable=False),
            sa.Column('created_at', sa.DateTime(), nullable=True),
            sa.Column('updated_at', sa.DateTime(), nullable=True),
            sa.ForeignKeyConstraint(['technician_id'], ['technician.id']),
            sa.PrimaryKeyConstraint('id')
        )

    existing_tables = set(sa.inspect(op.get_bind()).get_table_names())
    if 'incentive' not in existing_tables:
        op.create_table(
            'incentive',
            sa.Column('id', sa.Integer(), nullable=False),
            sa.Column('technician_id', sa.Integer(), nullable=False),
            sa.Column('service_id', sa.Integer(), nullable=True),
            sa.Column('invoice_id', sa.Integer(), nullable=True),
            sa.Column('period', sa.String(length=7), nullable=False),
            sa.Column('service_value', sa.Float(), nullable=True),
            sa.Column('percentage', sa.Float(), nullable=True),
            sa.Column('incentive_amount', sa.Float(), nullable=True),
            sa.Column('status', sa.String(length=30), nullable=True),
            sa.Column('payroll_id', sa.Integer(), nullable=True),
            sa.Column('approved_by', sa.String(length=100), nullable=True),
            sa.Column('approved_at', sa.DateTime(), nullable=True),
            sa.Column('rejection_reason', sa.Text(), nullable=True),
            sa.Column('created_at', sa.DateTime(), nullable=True),
            sa.Column('updated_at', sa.DateTime(), nullable=True),
            sa.ForeignKeyConstraint(['invoice_id'], ['invoice.id']),
            sa.ForeignKeyConstraint(['payroll_id'], ['payroll.id']),
            sa.ForeignKeyConstraint(['service_id'], ['service.id']),
            sa.ForeignKeyConstraint(['technician_id'], ['technician.id']),
            sa.PrimaryKeyConstraint('id'),
            sa.UniqueConstraint('service_id', name='uq_incentive_service')
        )

    existing_tables = set(sa.inspect(op.get_bind()).get_table_names())
    if 'incentive_audit_log' not in existing_tables:
        op.create_table(
            'incentive_audit_log',
            sa.Column('id', sa.Integer(), nullable=False),
            sa.Column('entity_type', sa.String(length=50), nullable=False),
            sa.Column('entity_id', sa.Integer(), nullable=False),
            sa.Column('action', sa.String(length=100), nullable=False),
            sa.Column('user', sa.String(length=100), nullable=True),
            sa.Column('before_value', sa.Text(), nullable=True),
            sa.Column('after_value', sa.Text(), nullable=True),
            sa.Column('notes', sa.Text(), nullable=True),
            sa.Column('created_at', sa.DateTime(), nullable=True),
            sa.PrimaryKeyConstraint('id')
        )


def downgrade():
    existing_tables = set(sa.inspect(op.get_bind()).get_table_names())
    if 'incentive_audit_log' in existing_tables:
        op.drop_table('incentive_audit_log')
    if 'incentive' in existing_tables:
        op.drop_table('incentive')
    if 'incentive_setting' in existing_tables:
        op.drop_table('incentive_setting')
