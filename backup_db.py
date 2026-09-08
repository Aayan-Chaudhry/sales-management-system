from pathlib import Path
from datetime import datetime
import sqlite3
import sys

DB_FILE = Path("sales.db")
BACKUP_DIR = Path("backups")
KEEP_LATEST = 50

def main():
    if not DB_FILE.exists():
        print("No sales.db found yet — skipping database backup.")
        return 0

    BACKUP_DIR.mkdir(exist_ok=True)

    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup_path = BACKUP_DIR / f"sales_backup_{timestamp}.db"

    try:
        # SQLite backup API makes a consistent backup, even if the database uses WAL mode.
        source = sqlite3.connect(str(DB_FILE))
        destination = sqlite3.connect(str(backup_path))
        with destination:
            source.backup(destination)
        source.close()
        destination.close()

        # Quick integrity check on the backup.
        check_conn = sqlite3.connect(str(backup_path))
        result = check_conn.execute("PRAGMA integrity_check").fetchone()[0]
        check_conn.close()

        if result != "ok":
            backup_path.unlink(missing_ok=True)
            print(f"Backup failed integrity check: {result}")
            return 1

        print(f"Database backup created: {backup_path}")

        # Keep only the newest backups.
        backups = sorted(
            BACKUP_DIR.glob("sales_backup_*.db"),
            key=lambda p: p.stat().st_mtime,
            reverse=True
        )

        for old_backup in backups[KEEP_LATEST:]:
            try:
                old_backup.unlink()
            except OSError:
                pass

        if len(backups) > KEEP_LATEST:
            print(f"Old backups cleaned. Keeping newest {KEEP_LATEST} backups.")

        return 0

    except Exception as e:
        print(f"Database backup failed: {e}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
