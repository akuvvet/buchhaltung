from io import BytesIO
from datetime import datetime
import re
import csv as csv_std
import pandas as pd
import chardet


def _last_month_slug() -> str:
    now = datetime.now()
    year = now.year
    month = now.month - 1
    if month == 0:
        month = 12
        year -= 1
    return f"{year:04d}-{month:02d}"


def _extract_filename_part(file_bytes: bytes) -> str:
    if not file_bytes:
        return "unbekannt"
    detection = chardet.detect(file_bytes[:4096])
    enc = detection.get("encoding") or "utf-8"
    try:
        text = file_bytes.decode(enc, errors="replace")
    except Exception:
        text = file_bytes.decode("utf-8", errors="replace")
    # 9. Zeile (i==8) durchsuchen
    lines = text.splitlines()
    if len(lines) > 8:
        line = lines[8]
        # Nach "Verkäufer" suchen und rechts davon nehmen
        if "Verkäufer" in line:
            part = line.split("Verkäufer")[-1].strip().strip('"')
        else:
            part = line.strip().strip('"')
        # Unerlaubte Zeichen für Dateinamen entfernen
        part = re.sub(r'[\\/*?:"<>|;]', "", part)
        return part or "unbekannt"
    return "unbekannt"


def process_ebay_csv(file_bytes: bytes, original_filename: str) -> tuple[bytes, str]:
    if not file_bytes:
        return b"", f"ebay-trans-unbekannt-{_last_month_slug()}.csv"

    detection = chardet.detect(file_bytes[:8192])
    enc = detection.get("encoding") or "utf-8"
    file_like = BytesIO(file_bytes)

    df = pd.read_csv(
        file_like,
        delimiter=";",
        encoding=enc,
        skiprows=11,
        quoting=csv_std.QUOTE_NONE,
        engine="python",
    )
    # Anführungszeichen aus Spalten und Zellen entfernen
    df.columns = [c.replace('"', "") for c in df.columns]
    for col in df.columns:
        if pd.api.types.is_string_dtype(df[col]):
            df[col] = df[col].str.replace('"', "", regex=False)

    columns_to_keep = [
        "Datum der Transaktionserstellung",
        "Typ",
        "Name des Käufers",
        "Versandziel - Land",
        "Betrag abzügl. Kosten",
        "Vom Verkäufer angegebener MwSt.-Satz",
        "Fixer Anteil der Verkaufsprovision",
        "Variabler Anteil der Verkaufsprovision",
        "Gebühr für sehr hohe Quote an „nicht wie beschriebenen Artikeln“",
        "Gebühr für unterdurchschnittlichen Servicestatus",
        "Transaktionsbetrag (inkl. Kosten)",
    ]
    df = df[[c for c in columns_to_keep if c in df.columns]]

    out_csv = df.to_csv(index=False, sep=";", encoding="utf-8")
    name_part = _extract_filename_part(file_bytes)
    filename = f"ebay-trans-{name_part}-{_last_month_slug()}.csv"
    return out_csv.encode("utf-8"), filename


