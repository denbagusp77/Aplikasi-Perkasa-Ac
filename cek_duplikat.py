import sqlite3

# Path ke database kamu
db_path = r"instance/ac_service.db"  # sesuaikan kalau beda

conn = sqlite3.connect(db_path)
cur = conn.cursor()

cur.execute("""
SELECT nomor_wa, COUNT(*) AS jumlah
FROM customer
GROUP BY nomor_wa
HAVING COUNT(*) > 1;
""")

print("Nomor WA yang duplikat:")
print("-" * 40)
rows = cur.fetchall()
if not rows:
    print("✅ Tidak ada duplikat")
else:
    for nomor_wa, jumlah in rows:
        print(f"{nomor_wa} → {jumlah} data")

conn.close()
