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
    "Gesamtbruttobetrag (inkl. Versand)": "Umsatz in Euro",
    "Dok.-Nr.": "Beleg1",
    "Dok. Datum": "Datum",
    "Kundenname": "Buchungstext",
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
    base = os.path.splitext(os.path.basename(original_filename or ""))[0]
    return f"{base}-{_last_month_slug()}.csv" if base else f"sale-ausgang-{_last_month_slug()}.csv"


def process_sale_ausgang_csv(file_bytes: bytes, original_filename: str) -> tuple[bytes, str]:
    if not file_bytes:
        return b"", _suggest_filename(original_filename)
    # Eingabe ist cp1252, Trennzeichen ';'
    df = pd.read_csv(BytesIO(file_bytes), sep=";", encoding="cp1252", dtype=object)

    out_df = pd.DataFrame(columns=TARGET_HEADERS)
    for src, tgt in SOURCE_TO_TARGET.items():
        if src in df.columns:
            out_df[tgt] = df[src]
        else:
            out_df[tgt] = None

    # Datum formatieren (Dok. Datum, falls vorhanden)
    if "Datum" in out_df.columns:
        def _fmt_date(v):
            if pd.isna(v):
                return None
            s = str(v)
            for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
                try:
                    return datetime.strptime(s, fmt).strftime("%d.%m.%Y")
                except Exception:
                    continue
            return s
        out_df["Datum"] = out_df["Datum"].map(_fmt_date)

    # Fixe Felder wie im Original-Skript
    out_df["Gegenkonto"] = "8400"
    out_df["Konto"] = "1000"

    for col in TARGET_HEADERS:
        if col not in out_df.columns:
            out_df[col] = None

    out_df = out_df[TARGET_HEADERS]
    # Ausgabe weiterhin in cp1252
    out_csv = out_df.to_csv(index=False, sep=";", encoding="cp1252")
    return out_csv.encode("cp1252", errors="replace"), _suggest_filename(original_filename)


