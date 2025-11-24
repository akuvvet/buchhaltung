from io import BytesIO, StringIO
from datetime import datetime
import os
import pandas as pd


TARGET_HEADERS = [
    "booking_date",
    "booking_text",
    "amount",
    "balance",
    "price_gross",
    "shipping_charges_gross",
    "sum_price_gross",
    "fee_%",
    "fee_net",
    "fee_vat_%",
    "fee_gross",
    "shipping.country",
]

SOURCE_TO_TARGET = {
    "booking_date": "booking_date",
    "booking_text": "booking_text",
    "amount": "amount",
    "balance": "balance",
    "price_gross": "price_gross",
    "shipping_charges_gross": "shipping_charges_gross",
    "sum_price_gross": "sum_price_gross",
    "fee_%": "fee_%",
    "fee_net": "fee_net",
    "fee_vat_%": "fee_vat_%",
    "fee_gross": "fee_gross",
    "shipping.country": "shipping.country",
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
    return f"{base}-{_last_month_slug()}.csv" if base else f"kaufland-{_last_month_slug()}.csv"


def process_kaufland_csv(file_bytes: bytes, original_filename: str) -> tuple[bytes, str]:
    if not file_bytes:
        return b"", _suggest_filename(original_filename)

    # CSV ist UTF-8 mit ';'
    df = pd.read_csv(BytesIO(file_bytes), sep=";", encoding="utf-8", dtype=str)
    # Ziel-DataFrame mit allen Zielspalten
    out_df = pd.DataFrame(columns=TARGET_HEADERS)
    for src, tgt in SOURCE_TO_TARGET.items():
        if src in df.columns:
            out_df[tgt] = df[src].astype(str).str.strip()
        else:
            out_df[tgt] = None

    # booking_date umwandeln: '%d.%m.%Y %H:%M' -> '%d.%m.%Y'
    if "booking_date" in out_df.columns:
        def _fmt(d: str) -> str | None:
            if not d or d.lower() == "nan":
                return None
            for fmt in ("%d.%m.%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%d.%m.%Y"):
                try:
                    return datetime.strptime(d, fmt).strftime("%d.%m.%Y")
                except Exception:
                    continue
            return None

        out_df["booking_date"] = out_df["booking_date"].map(_fmt)

    out_csv = out_df.to_csv(index=False, sep=";", encoding="utf-8")
    return out_csv.encode("utf-8"), _suggest_filename(original_filename)


