"""add technician_id to claim

Revision ID: 6765eb0bc425
Revises: 7672766910da
Create Date: 2025-09-20 10:37:17.245695

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = '6765eb0bc425'
down_revision = '7672766910da'
branch_labels = None
depends_on = None


def upgrade():
    # Gunakan batch_alter_table biar kompatibel SQLite
    with op.batch_alter_table("claim", schema=None) as batch_op:
        batch_op.add_column(sa.Column('technician_id', sa.Integer(), nullable=True))
        batch_op.create_foreign_key(
            'fk_claim_technician',   # nama constraint
            'technician',            # tabel tujuan
            ['technician_id'],       # kolom asal
            ['id']                   # kolom tujuan
        )


def downgrade():
    with op.batch_alter_table("claim", schema=None) as batch_op:
        batch_op.drop_constraint('fk_claim_technician', type_='foreignkey')
        batch_op.drop_column('technician_id')
