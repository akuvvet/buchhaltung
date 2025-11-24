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
    "Brutto Gesamt": "Umsatz in Euro",
    "Beleg-Nr.": "Beleg1",
    "Datum": "Datum",
    "Name 2": "Buchungstext",
    "Umsatzsteuer-ID": "Umsatzsteuer-ID",
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
    base = os.path.splitext(os.path.basename(original_filename or ""))[0].lower()
    date_part = _last_month_slug()
    if "okay" in base:
        return f"ausgang-mention-okay-{date_part}.csv"
    if "zone" in base:
        return f"ausgang-mention-zone-{date_part}.csv"
    return f"{base}-{date_part}.csv" if base else f"ausgang-mention-{date_part}.csv"


def process_mention_ausgang_excel(file_bytes: bytes, original_filename: str) -> tuple[bytes, str]:
    if not file_bytes:
        return b"", _suggest_filename(original_filename)

    # pandas kann xls und xlsx lesen (openpyxl/xlrd als Engines installiert)
    df = pd.read_excel(BytesIO(file_bytes), dtype=object)
    out_df = pd.DataFrame(columns=TARGET_HEADERS)

    for src, tgt in SOURCE_TO_TARGET.items():
        if src in df.columns:
            out_df[tgt] = df[src]
        else:
            out_df[tgt] = None

    # Datum formatieren
    if "Datum" in out_df.columns:
        def _fmt_date(v):
            if pd.isna(v):
                return None
            # Excel-Datum (als Timestamp/Serial) oder String
            try:
                return pd.to_datetime(v).strftime("%d.%m.%Y")
            except Exception:
                return None
        out_df["Datum"] = out_df["Datum"].map(_fmt_date)

    # Belegnummer ganzzahlig ohne .0
    if "Beleg1" in out_df.columns:
        def _fmt_beleg(v):
            if pd.isna(v):
                return None
            try:
                as_int = int(float(v))
                return str(as_int)
            except Exception:
                return str(v)
        out_df["Beleg1"] = out_df["Beleg1"].map(_fmt_beleg)

    # Fixe Felder wie im Original
    out_df["Gegenkonto"] = "3400"
    out_df["Konto"] = "1000"

    # Fehlende Zielspalten sicherstellen
    for col in TARGET_HEADERS:
        if col not in out_df.columns:
            out_df[col] = None

    out_df = out_df[TARGET_HEADERS]
    out_csv = out_df.to_csv(index=False, sep=";", encoding="utf-8")
    return out_csv.encode("utf-8"), _suggest_filename(original_filename)


