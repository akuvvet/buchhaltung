from io import BytesIO
from datetime import datetime
import os
import pandas as pd


TARGET_HEADERS = [
    "Umsatz in Euro",
    "Steuerschlüssel",
    "Gegenkonto",
    "Beleg1",
    "Beleg2",
    "Datum",
    "Konto",
    "Kost1",
    "Kost2",
    "Skonto in Euro",
    "Buchungstext",
    "Umsatzsteuer-ID",
    "Zusatzart",
    "Zusatzinformation",
]

SOURCE_TO_TARGET = {
    "Betrag MW": "Umsatz in Euro",
    "Beleg-Nr.": "Beleg1",
    "Datum": "Datum",
    "Name 2": "Buchungstext",
}


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
        return f"eingang-mention-okay-{date_part}.csv"
    if "zone" in lower:
        return f"eingang-mention-zone-{date_part}.csv"
    base = os.path.splitext(os.path.basename(original_filename or ""))[0]
    return f"{base}-{date_part}.csv" if base else f"eingang-mention-{date_part}.csv"


def process_mention_eingang_excel(file_bytes: bytes, original_filename: str) -> tuple[bytes, str]:
    if not file_bytes:
        return b"", _suggest_filename(original_filename)

    df = pd.read_excel(BytesIO(file_bytes), dtype=object)
    out_df = pd.DataFrame(columns=TARGET_HEADERS)

    for src, tgt in SOURCE_TO_TARGET.items():
        if src in df.columns:
            out_df[tgt] = df[src]
        else:
            out_df[tgt] = None

    if "Datum" in out_df.columns:
        def _fmt_date(v):
            if pd.isna(v):
                return None
            try:
                return pd.to_datetime(v).strftime("%d.%m.%Y")
            except Exception:
                return None
        out_df["Datum"] = out_df["Datum"].map(_fmt_date)

    # Fixe Felder
    out_df["Gegenkonto"] = "3400"
    out_df["Konto"] = "1000"

    for col in TARGET_HEADERS:
        if col not in out_df.columns:
            out_df[col] = None

    out_df = out_df[TARGET_HEADERS]
    out_csv = out_df.to_csv(index=False, sep=";", encoding="utf-8")
    return out_csv.encode("utf-8"), _suggest_filename(original_filename)


