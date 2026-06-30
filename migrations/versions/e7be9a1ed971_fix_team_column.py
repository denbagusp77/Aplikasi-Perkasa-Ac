"""fix team column

Revision ID: e7be9a1ed971
Revises: dec8ecba41a1
Create Date: 2025-09-07 22:31:00.545691

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'e7be9a1ed971'
down_revision = 'dec8ecba41a1'
branch_labels = None
depends_on = None


def upgrade():
    # Tambah kolom team ke tabel service
    with op.batch_alter_table('service', schema=None) as batch_op:
        batch_op.add_column(sa.Column('team', sa.String(length=50), nullable=True))


def downgrade():
    # Hapus kolom team dari tabel service
    with op.batch_alter_table('service', schema=None) as batch_op:
        batch_op.drop_column('team')
