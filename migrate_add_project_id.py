"""Add the project relationship column to the existing SQLite database."""

from datetime import datetime
from pathlib import Path
import shutil
import sqlite3


BASE_DIR = Path(__file__).resolve().parent
DATABASE = BASE_DIR / 'instance' / 'ac_service.db'


def main():
    if not DATABASE.exists():
        raise SystemExit(f'Database tidak ditemukan: {DATABASE}')

    backup = DATABASE.with_name(
        f'{DATABASE.stem}_before_project_id_{datetime.now():%Y%m%d_%H%M%S}{DATABASE.suffix}'
    )
    shutil.copy2(DATABASE, backup)

    with sqlite3.connect(DATABASE) as connection:
        columns = {row[1] for row in connection.execute("PRAGMA table_info('transaction')")}
        if 'project_id' not in columns:
            connection.execute('ALTER TABLE "transaction" ADD COLUMN project_id INTEGER')
            connection.commit()
            print('Kolom transaction.project_id berhasil ditambahkan.')
        else:
            print('Kolom transaction.project_id sudah tersedia.')

    print(f'Backup database: {backup}')


if __name__ == '__main__':
    main()