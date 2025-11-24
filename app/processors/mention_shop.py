from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
import os
import pandas as pd
from datetime import datetime


def _sanitize_ean(value):
    if pd.isna(value):
        return ""
    try:
        return str(int(value))
    except Exception:
        s = str(value).strip()
        return s if s and s.lower() != "nan" else ""


def _last_month_slug() -> str:
    now = datetime.now()
    year = now.year
    month = now.month - 1
    if month == 0:
        month = 12
        year -= 1
    return f"{year:04d}-{month:02d}"


def process_shop_files(excel_bytes: bytes, excel_filename: str, bestand_bytes: bytes, bestand_filename: str) -> tuple[bytes, str]:
    """
    Repliziert die Logik aus shopandmention.process_files:
    - Liest eine Mention-Excel (xls/xlsx) und eine Kosatec-Bestand (txt/csv, tab-separiert)
    - Erzeugt 4 CSVs: main.csv, shopimage.csv, bestand.csv, pricegroup.csv
    - Liefert sie als ZIP zurück
    """
    if not excel_bytes or not bestand_bytes:
        return b"", "mention-shop-files.zip"

    df = pd.read_excel(BytesIO(excel_bytes))
    bestand_df = pd.read_csv(BytesIO(bestand_bytes), sep="\t", on_bad_lines="skip")

    out_df = pd.DataFrame()
    out_df["ordernumber"] = df["Nummer"]
    out_df["mainnumber"] = df["Nummer"]
    out_df["name"] = df["Bezeichnung"]
    out_df["additionalText"] = df["Bezeichnung"]
    out_df["supplier"] = df["Hersteller-Name"]
    out_df["tax"] = 19
    out_df["price_EK"] = df["VK-Preis 1 Brutto"]
    out_df["pseudoprice_EK"] = 1
    out_df["baseprice_EK"] = ""
    out_df["from_EK"] = 1
    out_df["to_EK"] = "beliebig"
    out_df["price_H"] = df["VK-Preis 3"]
    out_df["pseudoprice_H"] = 0
    out_df["baseprice_H"] = ""
    out_df["from_H"] = 1
    out_df["to_H"] = "beliebig"
    out_df["active"] = 1

    bestand_map = dict(zip(bestand_df["herstnr"], bestand_df["menge"]))

    def choose_stock(row):
        herst = row.get("Hersteller-Nr.")
        fallback = row.get("Bestand")
        qty = bestand_map.get(herst)
        if qty is None or qty == 0:
            return fallback
        return qty

    out_df["instock"] = df.apply(choose_stock, axis=1)
    out_df["instock"] = out_df["instock"].fillna(0).astype(int)

    out_df["stockmin"] = 0
    out_df["description"] = ""
    out_df["description_long"] = df.get("Erweiterte Benennung", "")
    out_df["shippingtime"] = ""
    out_df["added"] = ""
    out_df["changed"] = ""
    out_df["releasedate"] = ""
    out_df["shippingfree"] = 0
    out_df["topseller"] = 0
    out_df["keywords"] = df.get("Keywords Online", "")
    out_df["minpurchase"] = 1
    out_df["purchasesteps"] = ""
    out_df["maxpurchase"] = ""
    out_df["purchaseunit"] = ""
    out_df["referenceunit"] = ""
    out_df["packunit"] = ""
    out_df["unitID"] = ""
    out_df["pricegroupID"] = 1
    out_df["pricegroupActive"] = 1
    out_df["laststock"] = 0
    out_df["suppliernumber"] = ""
    out_df["weight"] = ""
    out_df["width"] = ""
    out_df["height"] = ""
    out_df["length"] = ""
    out_df["ean"] = df.get("EAN-Code", "").map(_sanitize_ean)
    out_df["similar"] = ""
    out_df["configuratorsetID"] = ""
    out_df["configuratortype"] = ""
    out_df["configuratorOptions"] = ""
    out_df["categories"] = df.get("Kategorie", "")
    out_df["propertyGroupName"] = ""
    out_df["propertyValueName"] = ""
    out_df["accessory"] = ""
    out_df["imageUrl"] = ["https://www.okaycomputer.de/media/templates/produktbilder/" + str(nr) + "-1.jpg" for nr in df["Nummer"]]
    out_df["main"] = ""
    out_df["attr1"] = ""
    out_df["attr2"] = ""
    out_df["attr3"] = ""
    out_df["purchasePrice"] = ""
    out_df["metatitle"] = df.get("Meta Title Online", "")
    out_df["description"] = df.get("Meta Description Online", "")

    # shopimage.csv basierend auf Abbildung 1..3
    def _img_row(row, image_number: int):
        col = f"Abbildung {image_number}"
        if col in row and pd.notna(row[col]):
            nummer = row["Nummer"]
            return {
                "ordernumber": nummer,
                "image": f"https://www.okaycomputer.de/media/templates/produktbilder/{nummer}-{image_number}.jpg",
                "main": image_number,
                "description": "",
                "position": image_number,
                "width": 0,
                "height": 0,
                "relations": "",
            }
        return None

    image_rows = []
    for _, r in df.iterrows():
        for img_num in (1, 2, 3):
            entry = _img_row(r, img_num)
            if entry:
                image_rows.append(entry)
    shopimage_df = pd.DataFrame(image_rows)

    bestand_out_df = out_df[["ordernumber", "instock"]].copy()

    pricegroup_df = pd.DataFrame(
        {
            "ordernumber": pd.concat([out_df["ordernumber"], out_df["ordernumber"]]),
            "price": pd.concat([out_df["price_EK"], out_df["price_H"]]),
            "pricegroup": ["EK"] * len(out_df) + ["H"] * len(out_df),
        }
    )

    # ZIP erstellen
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, "w", compression=ZIP_DEFLATED) as zf:
        zf.writestr("main.csv", out_df.to_csv(index=False, sep=";", encoding="utf-8"))
        zf.writestr("shopimage.csv", shopimage_df.to_csv(index=False, sep=";", encoding="utf-8"))
        zf.writestr("bestand.csv", bestand_out_df.to_csv(index=False, sep=";", encoding="utf-8"))
        zf.writestr("pricegroup.csv", pricegroup_df.to_csv(index=False, sep=";", encoding="utf-8"))
    zip_buffer.seek(0)

    base = os.path.splitext(os.path.basename(excel_filename or "mention"))[0]
    suggested = f"{base}-{_last_month_slug()}-shopfiles.zip"
    return zip_buffer.getvalue(), suggested


