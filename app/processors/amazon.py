from io import BytesIO
from datetime import datetime
import os
import pandas as pd


def _last_month_slug() -> str:
    now = datetime.now()
    year = now.year
    month = now.month - 1
    if month == 0:
        month = 12
        year -= 1
    return f"{year:04d}-{month:02d}"


def _suggest_filename(original_filename: str) -> str:
    lower = (original_filename or "").lower()
    date_part = _last_month_slug()
    if "okay" in lower:
        return f"amazon-trans-okay-{date_part}.csv"
    if "zone" in lower:
        return f"amazon-trans-zone-{date_part}.csv"
    return f"amazon-trans-{date_part}.csv"


def process_amazon_csv(file_bytes: bytes, original_filename: str) -> tuple[bytes, str]:
    """
    Erwartet eine Amazon CSV (Original: delimiter ',', header=7).
    Gibt CSV mit ';' als Trennzeichen zurück und vorgeschlagenen Dateinamen.
    """
    if not file_bytes:
        return b"", _suggest_filename(original_filename)

    bio = BytesIO(file_bytes)
    df = pd.read_csv(bio, delimiter=",", encoding="utf-8", header=7)

    # Datum/Uhrzeit -> nur Datum
    if "Datum/Uhrzeit" in df.columns:
        df["Datum/Uhrzeit"] = pd.to_datetime(
            df["Datum/Uhrzeit"], format="%d.%m.%Y %H:%M:%S %Z", errors="coerce"
        ).dt.strftime("%d.%m.%Y")

    columns_to_keep = [
        "Datum/Uhrzeit",
        "Typ",
        "Menge",
        "Marketplace",
        "Umsätze",
        "Verkaufsgebühren",
        "Andere Transaktionsgebühren",
        "Andere",
        "Gesamt",
    ]
    df = df[[c for c in columns_to_keep if c in df.columns]]

    out = df.to_csv(index=False, sep=";", encoding="utf-8")
    return out.encode("utf-8"), _suggest_filename(original_filename)


