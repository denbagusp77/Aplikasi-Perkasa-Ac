"""add service_id to claim

Revision ID: 614153718903
Revises: 6765eb0bc425
Create Date: 2025-09-20 12:00:00
"""

from alembic import op
import sqlalchemy as sa

# revision identifiers, used by Alembic.
revision = '614153718903'
down_revision = '6765eb0bc425'
branch_labels = None
depends_on = None


def upgrade():
    with op.batch_alter_table("claim", schema=None) as batch_op:
        batch_op.add_column(sa.Column("service_id", sa.Integer(), nullable=True))
        batch_op.create_foreign_key(
            "fk_claim_service",   # 🔑 kasih nama constraint
            "service",            # referensi ke tabel
            ["service_id"],       # kolom lokal
            ["id"],               # kolom referensi
            ondelete="SET NULL"
        )


def downgrade():
    with op.batch_alter_table("claim", schema=None) as batch_op:
        batch_op.drop_constraint("fk_claim_service", type_="foreignkey")  # 🔑 pakai nama yg sama
        batch_op.drop_column("service_id")
