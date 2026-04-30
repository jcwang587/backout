import os
import re
import sqlite3
import hashlib
from datetime import datetime

BASE_DIR = os.path.expanduser("~/email_backup")
DB_PATH  = os.path.join(BASE_DIR, "backup.db")


def init_db():
    os.makedirs(BASE_DIR, exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.execute("""
        CREATE TABLE IF NOT EXISTS emails (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            hash         TEXT UNIQUE,
            date         TEXT,
            sender       TEXT,
            subject      TEXT,
            filepath     TEXT,
            backed_up_at TEXT
        )
    """)
    con.commit()
    return con


def email_hash(date_str, sender, subject):
    key = f"{date_str}|{sender}|{subject}"
    return hashlib.sha256(key.encode()).hexdigest()


def is_backed_up(con, hash_val):
    row = con.execute("SELECT 1 FROM emails WHERE hash = ?", (hash_val,)).fetchone()
    return row is not None


def record_backup(con, hash_val, date_str, sender, subject, filepath):
    con.execute(
        "INSERT INTO emails (hash, date, sender, subject, filepath, backed_up_at) VALUES (?,?,?,?,?,?)",
        (hash_val, date_str, sender, subject, filepath, datetime.now().isoformat())
    )
    con.commit()


def get_stats(year=None, month=None):
    con = sqlite3.connect(DB_PATH)
    filters, params = [], []

    if year:
        filters.append("strftime('%Y', date) = ?")
        params.append(str(year))
    if month:
        filters.append("strftime('%m', date) = ?")
        params.append(f"{int(month):02d}")

    where = f"WHERE {' AND '.join(filters)}" if filters else ""

    total = con.execute(f"SELECT COUNT(*) FROM emails {where}", params).fetchone()[0]

    by_month = con.execute(f"""
        SELECT strftime('%Y-%m', date) as month, COUNT(*) as count
        FROM emails {where}
        GROUP BY month ORDER BY month DESC
    """, params).fetchall()

    top_senders = con.execute(f"""
        SELECT sender, COUNT(*) as count
        FROM emails {where}
        GROUP BY sender ORDER BY count DESC LIMIT 10
    """, params).fetchall()

    con.close()

    print(f"\n{'─' * 50}")
    print(f"  Backup statistics" + (f" — {year or ''}/{month or ''}".rstrip("/") if year or month else ""))
    print(f"{'─' * 50}")
    print(f"  Total emails backed up: {total}")

    print(f"\n  By month:")
    for m, count in by_month:
        print(f"    {m}  {count} emails")

    print(f"\n  Top senders:")
    for sender, count in top_senders:
        print(f"    {count:>4}  {sender}")
    print(f"{'─' * 50}\n")


def safe_filename(text, max_length=60):
    text = re.sub(r'[\\/*?:"<>|\n\r]', "", text)
    return text.strip()[:max_length]


def parse_date(date_str):
    for fmt in [
        "%A, %B %d, %Y at %I:%M:%S %p",
        "%A, %B %d, %Y at %I:%M %p",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
    ]:
        try:
            return datetime.strptime(date_str.strip(), fmt)
        except ValueError:
            continue
    return datetime.now()


def save_email(date_str, sender, subject, body, con=None):
    close_after = con is None
    if con is None:
        con = init_db()

    date_key = date_str if isinstance(date_str, str) else date_str.strftime("%Y-%m-%d %H:%M")
    hash_val = email_hash(date_key, sender, subject)

    if is_backed_up(con, hash_val):
        if close_after:
            con.close()
        return False  # already exists, skipped

    dt = parse_date(date_key) if isinstance(date_key, str) else date_key
    folder = os.path.join(BASE_DIR, dt.strftime("%Y"), dt.strftime("%m"))
    os.makedirs(folder, exist_ok=True)

    base = f"{dt.strftime('%Y-%m-%d_%H%M')}_{safe_filename(subject)}"
    filepath = os.path.join(folder, f"{base}.txt")
    counter = 2
    while os.path.exists(filepath):
        filepath = os.path.join(folder, f"{base}_{counter}.txt")
        counter += 1

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(f"DATE:    {date_key}\n")
        f.write(f"FROM:    {sender.strip()}\n")
        f.write(f"SUBJECT: {subject.strip()}\n")
        f.write(f"{'─' * 60}\n\n")
        f.write(body.strip())
        f.write("\n")

    record_backup(con, hash_val, dt.strftime("%Y-%m-%d %H:%M"), sender, subject, filepath)

    if close_after:
        con.close()
    return True  # newly saved
