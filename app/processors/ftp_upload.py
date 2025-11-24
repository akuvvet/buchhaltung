from configparser import ConfigParser
from io import BytesIO, StringIO
import os
import ftplib
import csv


def _load_config():
    """
    Liest FTP-Konfiguration aus config.ini oder Umgebungsvariablen:
    - ENV: FTP_HOST, FTP_USER, FTP_PASS, FTP_DIR
    - INI: [FTP] host, user, password, directory
    """
    cfg = ConfigParser()
    cfg.read("config.ini")
    host = os.getenv("FTP_HOST") or (cfg.get("FTP", "host", fallback=None) if cfg.has_section("FTP") else None)
    user = os.getenv("FTP_USER") or (cfg.get("FTP", "user", fallback=None) if cfg.has_section("FTP") else None)
    password = os.getenv("FTP_PASS") or (cfg.get("FTP", "password", fallback=None) if cfg.has_section("FTP") else None)
    directory = os.getenv("FTP_DIR") or (cfg.get("FTP", "directory", fallback=None) if cfg.has_section("FTP") else None)
    if not all([host, user, password, directory]):
        raise RuntimeError(
            "FTP-Konfiguration fehlt. Legen Sie eine config.ini mit [FTP] host/user/password/directory an "
            "oder setzen Sie die Umgebungsvariablen FTP_HOST, FTP_USER, FTP_PASS, FTP_DIR."
        )
    return host, user, password, directory


def _connect():
    host, user, password, directory = _load_config()
    ftp = ftplib.FTP(host)
    ftp.login(user, password)
    ftp.cwd(directory)
    return ftp


def upload_images_and_generate_csv(artikelnummer: str, files: list[tuple[str, bytes]]) -> tuple[bytes, str]:
    """
    Lädt Bilder zu FTP hoch und erzeugt shopimage CSV.
    Dateien werden als {artikelnummer}-{laufende Nummer}.jpg gespeichert.
    """
    if not artikelnummer or not files:
        return b"", f"shopimage-{artikelnummer or 'unknown'}.csv"

    ftp = _connect()
    uploaded = []
    try:
        for idx, (_name, content) in enumerate(files, start=1):
            filename = f"{artikelnummer}-{idx}.jpg"
            bio = BytesIO(content)
            ftp.storbinary(f"STOR {filename}", bio)
            uploaded.append(filename)
    finally:
        try:
            ftp.quit()
        except Exception:
            pass

    # CSV bauen
    url_base = "https://www.okaycomputer.de/media/templates/produktbilder/"
    # csv.writer erwartet einen Text-Stream → StringIO nutzen und anschließend nach UTF-8 kodieren
    output = StringIO()
    writer = csv.writer(output, delimiter=";", quotechar='"', quoting=csv.QUOTE_MINIMAL, lineterminator="\n")
    writer.writerow(["ordernumber", "image", "main", "description", "position", "width", "height", "relations"])
    for idx, _fname in enumerate(uploaded, start=1):
        image_url = f"{url_base}{artikelnummer}-{idx}.jpg"
        writer.writerow([artikelnummer, image_url, 1 if idx == 1 else 0, "", idx, 0, 0, ""])
    csv_text = output.getvalue()
    return csv_text.encode("utf-8"), f"shopimage-{artikelnummer}.csv"


