"""add technician_id to Service

Revision ID: dec8ecba41a1
Revises: ab06add7e145
Create Date: 2025-09-07 22:25:10.411661

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'dec8ecba41a1'
down_revision = 'ab06add7e145'
branch_labels = None
depends_on = None


def upgrade():
    # Tambahkan kolom technician_id ke tabel service
    with op.batch_alter_table("service", schema=None) as batch_op:
        batch_op.add_column(sa.Column("technician_id", sa.Integer(), nullable=True))
        batch_op.create_foreign_key(
            "fk_service_technician_id",  # nama constraint
            "technician",               # tabel tujuan
            ["technician_id"],          # kolom di service
            ["id"]                      # kolom di technician
        )


def downgrade():
    # Hapus kolom technician_id dari tabel service
    with op.batch_alter_table("service", schema=None) as batch_op:
        batch_op.drop_constraint("fk_service_technician_id", type_="foreignkey")
        batch_op.drop_column("technician_id")
