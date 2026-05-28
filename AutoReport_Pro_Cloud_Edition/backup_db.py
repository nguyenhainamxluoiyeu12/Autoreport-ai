"""Tạo backup local cho SQLite/outputs/uploads.
Chạy: python backup_db.py
"""
from pathlib import Path
from datetime import datetime
import zipfile

BASE_DIR = Path(__file__).resolve().parent
stamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
backup_dir = BASE_DIR / "backups"
backup_dir.mkdir(exist_ok=True)
zip_path = backup_dir / f"autoreport_backup_{stamp}.zip"

with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
    for fp in [BASE_DIR / "autoreport_cloud.db", BASE_DIR / ".env"]:
        if fp.exists():
            z.write(fp, fp.name)
    for folder_name in ["outputs", "static/uploads"]:
        folder = BASE_DIR / folder_name
        if folder.exists():
            for file in folder.rglob("*"):
                if file.is_file():
                    z.write(file, str(file.relative_to(BASE_DIR)))
print(f"Backup xong: {zip_path}")
