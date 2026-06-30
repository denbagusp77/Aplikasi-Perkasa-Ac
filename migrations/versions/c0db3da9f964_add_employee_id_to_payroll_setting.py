"""add employee_id to payroll_setting

Revision ID: c0db3da9f964
Revises: 84d258acc22d
Create Date: 2025-10-10 00:16:39.719709
"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'c0db3da9f964'
down_revision = '84d258acc22d'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table('payroll_setting', schema=None) as batch_op:
        batch_op.add_column(sa.Column('employee_id', sa.Integer(), nullable=True))
        batch_op.add_column(sa.Column('gaji_pokok', sa.Float(), nullable=True))
        batch_op.add_column(sa.Column('bonus_bulanan', sa.Float(), nullable=True))
        batch_op.create_foreign_key(
            'fk_payroll_setting_employee_id', 'technician', ['employee_id'], ['id']
        )

        # hapus kolom lama jika ada
        for col in ['potongan_kasbon_default', 'bonus_bulanan_default', 'gaji_pokok_default', 'potongan_service_default']:
            try:
                batch_op.drop_column(col)
            except Exception:
                pass



def downgrade():
    with op.batch_alter_table('payroll_setting', schema=None) as batch_op:
        # tambahkan kembali kolom lama jika rollback
        batch_op.add_column(sa.Column('potongan_service_default', sa.FLOAT(), nullable=True))
        batch_op.add_column(sa.Column('gaji_pokok_default', sa.FLOAT(), nullable=True))
        batch_op.add_column(sa.Column('bonus_bulanan_default', sa.FLOAT(), nullable=True))
        batch_op.add_column(sa.Column('potongan_kasbon_default', sa.FLOAT(), nullable=True))
        # hapus foreign key dengan nama yang sama
        batch_op.drop_constraint('fk_payroll_setting_employee_id', type_='foreignkey')
        batch_op.drop_column('bonus_bulanan')
        batch_op.drop_column('gaji_pokok')
        batch_op.drop_column('employee_id')
