"""add blast log

Revision ID: e17366faccc9
Revises: <new_revision_id>
Create Date: 2025-09-18 22:54:05.881467

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'e17366faccc9'
down_revision = '<new_revision_id>'
branch_labels = None
depends_on = None


def upgrade():
    # ✅ hanya tambah tabel blast_log
    op.create_table(
        'blast_log',
        sa.Column('id', sa.Integer(), primary_key=True),
        sa.Column('message', sa.Text(), nullable=False),
        sa.Column('targets', sa.Text(), nullable=False),
        sa.Column('status', sa.String(length=50), nullable=True),
        sa.Column('created_at', sa.DateTime(), nullable=True),
    )


def downgrade():
    # ✅ rollback: hapus tabel blast_log
    op.drop_table('blast_log')
