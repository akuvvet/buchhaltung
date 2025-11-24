import pandas as pd
import ftplib
import os
from datetime import datetime
from tkinter import Tk, simpledialog, messagebox, Button, Label, Toplevel, filedialog
import tkinter as tk  # Importiere tkinter als tk
from PIL import ImageGrab, Image
import csv
from configparser import ConfigParser
import logging

# ======================== Logging ========================
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# ======================== Konfiguration laden ========================
config = ConfigParser()
config.read('config.ini')

FTP_HOST = config.get('FTP', 'host')
FTP_USER = config.get('FTP', 'user')
FTP_PASS = config.get('FTP', 'password')
FTP_DIR = config.get('FTP', 'directory')

LOCAL_DIR = config.get('Local', 'directory')

# ======================== FTP-Verbindungen ========================
def connect_ftp():
    try:
        ftp = ftplib.FTP(FTP_HOST)
        ftp.login(FTP_USER, FTP_PASS)
        ftp.cwd(FTP_DIR)
        logging.info(f"Verbunden mit FTP-Server: {FTP_HOST}")
        return ftp
    except ftplib.all_errors as e:
        logging.error(f"FTP-Verbindungsfehler: {e}")
        messagebox.showerror("FTP-Verbindungsfehler", str(e))
        return None

def get_local_files(local_dir, date):
    files = []
    for root, _, filenames in os.walk(local_dir):
        for filename in filenames:
            if filename.endswith(".jpg"):
                file_path = os.path.join(root, filename)
                file_date = datetime.fromtimestamp(os.path.getmtime(file_path))
                if file_date >= date:
                    files.append(file_path)
    return files

def get_ftp_files(ftp):
    try:
        ftp_files = ftp.nlst()
        return ftp_files
    except ftplib.error_perm as e:
        if str(e) == "550 No files found":
            logging.info("Keine Dateien im Verzeichnis gefunden.")
        else:
            logging.error(f"Fehler beim Abrufen der FTP-Dateien: {e}")
        return []

def upload_file(ftp, local_file):
    try:
        with open(local_file, 'rb') as file:
            ftp.storbinary(f'STOR {os.path.basename(local_file)}', file)
        logging.info(f"Datei hochgeladen: {local_file}")
    except ftplib.all_errors as e:
        logging.error(f"Fehler beim Hochladen der Datei {local_file}: {e}")

def upload_multiple_files():
    date_str = simpledialog.askstring("Datum eingeben", "Bitte geben Sie das Datum im Format YYYY-MM-DD ein:")
    if not date_str:
        logging.info("Kein Datum eingegeben. Programm wird beendet.")
        return

    try:
        date = datetime.strptime(date_str, "%Y-%m-%d")
    except ValueError:
        logging.error("Ungültiges Datum. Programm wird beendet.")
        return

    ftp = connect_ftp()
    if ftp is None:
        return

    local_files = get_local_files(LOCAL_DIR, date)
    ftp_files = get_ftp_files(ftp)
    
    uploaded_files = []

    for local_file in local_files:
        filename = os.path.basename(local_file)
        if filename in ftp_files:
            logging.info(f"Datei {filename} existiert bereits auf dem Server und wird übersprungen.")
        else:
            logging.info(f"Datei {filename} wird auf den Server hochgeladen.")
            upload_file(ftp, local_file)
            uploaded_files.append(filename)

    ftp.quit()
    
    if uploaded_files:
        messagebox.showinfo("Übertragungsbericht", f"Die folgenden Dateien wurden hochgeladen:\n\n" + "\n".join(uploaded_files))
    else:
        messagebox.showinfo("Übertragungsbericht", "Es wurden keine neuen Dateien hochgeladen.")

    logging.info("Alle Dateien wurden überprüft und gegebenenfalls hochgeladen.")

def upload_single_file():
    artikelnummer = simpledialog.askstring("Artikelnummer eingeben", "Bitte geben Sie die Artikelnummer ein:")
    if not artikelnummer:
        logging.info("Keine Artikelnummer eingegeben. Programm wird beendet.")
        return

    local_files = [f for f in os.listdir(LOCAL_DIR) if f.startswith(artikelnummer) and f.endswith(".jpg")]
    if not local_files:
        messagebox.showerror("Fehler", f"Es wurden keine Dateien mit der Artikelnummer {artikelnummer} im lokalen Verzeichnis gefunden.")
        return

    ftp = connect_ftp()
    if ftp is None:
        return

    ftp_files = get_ftp_files(ftp)
    uploaded_files = []

    for local_file in local_files:
        local_file_path = os.path.join(LOCAL_DIR, local_file)
        if local_file in ftp_files:
            logging.info(f"Datei {local_file} existiert bereits auf dem Server und wird überschrieben.")
        else:
            logging.info(f"Datei {local_file} wird auf den Server hochgeladen.")
        upload_file(ftp, local_file_path)
        uploaded_files.append(local_file)

    ftp.quit()

    if uploaded_files:
        messagebox.showinfo("Übertragungsbericht", f"Die folgenden Dateien wurden hochgeladen:\n\n" + "\n".join(uploaded_files))
        create_csv_file(artikelnummer, uploaded_files)
    else:
        messagebox.showinfo("Übertragungsbericht", "Es wurden keine neuen Dateien hochgeladen.")

def create_csv_file(artikelnummer, uploaded_files):
    csv_filename = f"shopimage-{artikelnummer}.csv"
    save_path = filedialog.asksaveasfilename(
        initialfile=csv_filename,
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if not save_path:
        logging.info("Speichern abgebrochen.")
        return

    csv_url_base = "https://www.okaycomputer.de/media/templates/produktbilder/"
    
    try:
        with open(save_path, mode='w', newline='', encoding='utf-8') as csvfile:
            csv_writer = csv.writer(csvfile, delimiter=';', quotechar='"', quoting=csv.QUOTE_MINIMAL)
            csv_writer.writerow(['ordernumber', 'image', 'main', 'description', 'position', 'width', 'height', 'relations'])
            
            for idx, file in enumerate(uploaded_files, start=1):
                image_url = f"{csv_url_base}{artikelnummer}-{idx}.jpg"
                csv_writer.writerow([artikelnummer, image_url, 1 if idx == 1 else 0, '', idx, 0, 0, ''])
        
        logging.info(f"CSV-Datei wurde erstellt: {save_path}")
    except Exception as e:
        logging.error(f"Fehler beim Erstellen der CSV-Datei: {e}")

def waehle_datei_aus(title, filetypes=[("Alle Dateien", "*.*")]):
    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=filetypes
    )
    return file_path

def process_files():
    excel_dateipfad = waehle_datei_aus("Wählen Sie die aus der dsInfo erstellte xls Datei aus", [("Excel-Dateien", "*.xls *.xlsx")])
    bestand_dateipfad = waehle_datei_aus("Bitte wählen Sie die Kosatec txt Datei aus!", [("Textdateien", "*.txt"), ("CSV-Dateien", "*.csv")])

    if excel_dateipfad and bestand_dateipfad:
        try:
            df = pd.read_excel(excel_dateipfad)
            bestand_df = pd.read_csv(bestand_dateipfad, sep='\t', on_bad_lines='skip')
            
            neuer_df = pd.DataFrame()
            neuer_df['ordernumber'] = df['Nummer']
            neuer_df['mainnumber'] = df['Nummer']
            neuer_df['name'] = df['Bezeichnung']
            neuer_df['additionalText'] = df['Bezeichnung']
            neuer_df['supplier'] = df['Hersteller-Name']
            neuer_df['tax'] = 19
            neuer_df['price_EK'] = df['VK-Preis 1 Brutto']
            neuer_df['pseudoprice_EK'] = 1
            neuer_df['baseprice_EK'] = ''
            neuer_df['from_EK'] = 1
            neuer_df['to_EK'] = 'beliebig'
            neuer_df['price_H'] = df['VK-Preis 3']
            neuer_df['pseudoprice_H'] = 0
            neuer_df['baseprice_H'] = ''
            neuer_df['from_H'] = 1
            neuer_df['to_H'] = 'beliebig'
            neuer_df['active'] = 1
            
            bestand_dict = dict(zip(bestand_df['herstnr'], bestand_df['menge']))
            
            def waehle_bestand(hersteller_nr, bestand):
                menge = bestand_dict.get(hersteller_nr)
                if menge is None or menge == 0:
                    return bestand
                else:
                    return menge
            
            neuer_df['instock'] = df.apply(lambda x: waehle_bestand(x['Hersteller-Nr.'], x['Bestand']), axis=1)
            neuer_df['instock'] = neuer_df['instock'].fillna(0).astype(int)

            neuer_df['stockmin'] = 0
            neuer_df['description'] = ''
            neuer_df['description_long'] = df['Erweiterte Benennung']
            neuer_df['shippingtime'] = ''
            neuer_df['added'] = ''
            neuer_df['changed'] = ''
            neuer_df['releasedate'] = ''
            neuer_df['shippingfree'] = 0
            neuer_df['topseller'] = 0
            neuer_df['keywords'] = df['Keywords Online']
            neuer_df['minpurchase'] = 1
            neuer_df['purchasesteps'] = ''
            neuer_df['maxpurchase'] = ''
            neuer_df['purchaseunit'] = ''
            neuer_df['referenceunit'] = ''
            neuer_df['packunit'] = ''
            neuer_df['unitID'] = ''
            neuer_df['pricegroupID'] = 1
            neuer_df['pricegroupActive'] = 1
            neuer_df['laststock'] = 0
            neuer_df['suppliernumber'] = ''
            neuer_df['weight'] = ''
            neuer_df['width'] = ''
            neuer_df['height'] = ''
            neuer_df['length'] = ''
            neuer_df['ean'] = df['EAN-Code'].apply(lambda x: str(int(x)) if pd.notnull(x) else '')
            neuer_df['similar'] = ''
            neuer_df['configuratorsetID'] = ''
            neuer_df['configuratortype'] = ''
            neuer_df['configuratorOptions'] = ''
            neuer_df['categories'] = df['Kategorie']
            neuer_df['propertyGroupName'] = ''
            neuer_df['propertyValueName'] = ''
            neuer_df['accessory'] = ''
            neuer_df['imageUrl'] = ["https://www.okaycomputer.de/media/templates/produktbilder/" + str(nummer) + "-1.jpg" for nummer in df['Nummer']]
            neuer_df['main'] = ''
            neuer_df['attr1'] = ''
            neuer_df['attr2'] = ''
            neuer_df['attr3'] = ''
            neuer_df['purchasePrice'] = ''
            neuer_df['metatitle'] = df['Meta Title Online']
            neuer_df['description'] = df['Meta Description Online']

            csv_dateipfad = excel_dateipfad.replace('.xls', '.csv').replace('.xlsx', '.csv')
            neuer_df.to_csv(csv_dateipfad, index=False, sep=';')
            logging.info(f"Die CSV-Datei wurde gespeichert unter: {csv_dateipfad}")

            def create_image_rows(row, image_number):
                if pd.notnull(row[f'Abbildung {image_number}']):
                    return {
                        'ordernumber': row['Nummer'],
                        'image': f"https://www.okaycomputer.de/media/templates/produktbilder/{row['Nummer']}-{image_number}.jpg",
                        'main': image_number,
                        'description': '',
                        'position': image_number,
                        'width': 0,
                        'height': 0,
                        'relations': ''
                    }
                return None

            image_rows = []
            for _, row in df.iterrows():
                for image_number in [1, 2, 3]:
                    image_row = create_image_rows(row, image_number)
                    if image_row:
                        image_rows.append(image_row)

            shopimage_df = pd.DataFrame(image_rows)

            shopimage_csv_dateipfad = excel_dateipfad.replace('.xls', '_shopimage.csv').replace('.xlsx', '_shopimage.csv')
            shopimage_df.to_csv(shopimage_csv_dateipfad, index=False, sep=';')
            logging.info(f"Die shopimage.csv Datei wurde gespeichert unter: {shopimage_csv_dateipfad}")

            bestand_df = neuer_df[['ordernumber', 'instock']]
            bestand_csv_dateipfad = excel_dateipfad.replace('.xls', '_bestand.csv').replace('.xlsx', '_bestand.csv')
            bestand_df.to_csv(bestand_csv_dateipfad, index=False, sep=';')
            logging.info(f"Die Bestand CSV-Datei wurde gespeichert unter: {bestand_csv_dateipfad}")

            pricegroup_df = pd.DataFrame({
                'ordernumber': pd.concat([neuer_df['ordernumber'], neuer_df['ordernumber']]),
                'price': pd.concat([neuer_df['price_EK'], neuer_df['price_H']]),
                'pricegroup': ['EK'] * len(neuer_df) + ['H'] * len(neuer_df)
            })
            pricegroup_csv_dateipfad = excel_dateipfad.replace('.xls', '_pricegroup.csv').replace('.xlsx', '_pricegroup.csv')
            pricegroup_df.to_csv(pricegroup_csv_dateipfad, index=False, sep=';')
            logging.info(f"Die pricegroup CSV-Datei wurde gespeichert unter: {pricegroup_csv_dateipfad}")

        except Exception as e:
            logging.error(f"Fehler beim Verarbeiten der Dateien: {e}")
    else:
        logging.info("Eine oder beide Dateien wurden nicht ausgewählt.")

def zeige_benutzerhinweis():
    messagebox.showinfo("Hinweis", "1. Screenshot erstellen\nKlicke OK, um fortzufahren.")
    erstelle_screenshot()

def erstelle_screenshot():
    zielordner = r'M:\mention\bilder'
    max_breite = 800

    screenshot = ImageGrab.grabclipboard()

    if screenshot:
        breite, hoehe = screenshot.size

        if breite > max_breite:
            neue_breite = max_breite
            neue_hoehe = int(hoehe * (max_breite / breite))

            screenshot = screenshot.resize((neue_breite, neue_hoehe), Image.LANCZOS)

        screenshot = screenshot.convert("RGB")

        dateien = os.listdir(zielordner)
        max_nummer = 0

        for datei in dateien:
            if datei.endswith(".jpg") and "-" in datei:
                try:
                    nummer = int(datei.split("-")[0])
                    max_nummer = max(max_nummer, nummer)
                except ValueError:
                    pass

        neue_nummer = max_nummer + 1
        dateiname = f"{neue_nummer}-1.jpg"
        bild_pfad = os.path.join(zielordner, dateiname)

        screenshot.save(bild_pfad, "JPEG")
        logging.info(f"Bild erfolgreich als '{dateiname}' gespeichert.")

        zweites_bild = messagebox.askyesno("Frage", "Erstelle 2. Screenshot")

        if zweites_bild:
            screenshot = ImageGrab.grabclipboard()
            breite, hoehe = screenshot.size

            if breite > max_breite:
                neue_breite = max_breite
                neue_hoehe = int(hoehe * (max_breite / breite))

                screenshot = screenshot.resize((neue_breite, neue_hoehe), Image.LANCZOS)

            screenshot = screenshot.convert("RGB")
            dateiname2 = f"{neue_nummer}-2.jpg"
            bild_pfad2 = os.path.join(zielordner, dateiname2)
            screenshot.save(bild_pfad2, "JPEG")
            logging.info(f"Zweites Bild erfolgreich als '{dateiname2}' gespeichert.")

            drittes_bild = messagebox.askyesno("Frage", "Erstelle 3. Screenshot?")

            if drittes_bild:
                screenshot = ImageGrab.grabclipboard()
                breite, hoehe = screenshot.size

                if breite > max_breite:
                    neue_breite = max_breite
                    neue_hoehe = int(hoehe * (max_breite / breite))

                screenshot = screenshot.resize((neue_breite, neue_hoehe), Image.LANCZOS)

                screenshot = screenshot.convert("RGB")
                dateiname3 = f"{neue_nummer}-3.jpg"
                bild_pfad3 = os.path.join(zielordner, dateiname3)
                screenshot.save(bild_pfad3, "JPEG")
                logging.info(f"Drittes Bild erfolgreich als '{dateiname3}' gespeichert.")
            else:
                logging.info("Das Skript wurde beendet, da kein drittes Bild existiert.")
        else:
            logging.info("Das Skript wurde beendet, da kein zweites Bild existiert.")
    else:
        logging.info("Kein Bild in der Zwischenablage gefunden.")

class ToolTip(object):
    def __init__(self, widget, text, delay=500):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tip_window = None
        self.id = None
        self.x = self.y = 0
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)

    def enter(self, event=None):
        self.schedule()

    def leave(self, event=None):
        self.unschedule()
        self.hidetip()

    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.delay, self.showtip)

    def unschedule(self):
        id = self.id
        self.id = None
        if id:
            self.widget.after_cancel(id)

    def showtip(self, event=None):
        x, y, _cx, _cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 25
        y = y + self.widget.winfo_rooty() + 25
        self.tip_window = tw = Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry(f"+{x}+{y}")
        label = Label(tw, text=self.text, justify=tk.LEFT,
                      background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                      font=("tahoma", "10", "normal"))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tip_window
        self.tip_window = None
        if tw:
            tw.destroy()

def show_choice_window():
    choice_window = Toplevel(root)
    choice_window.title("Bildübertragung")

    def on_multiple():
        choice_window.destroy()
        upload_multiple_files()
        show_choice_window()

    def on_single():
        choice_window.destroy()
        upload_single_file()
        show_choice_window()

    def on_process_files():
        choice_window.destroy()
        process_files()
        show_choice_window()

    def on_mention_bilder():
        choice_window.destroy()
        zeige_benutzerhinweis()
        show_choice_window()

    def on_exit():
        choice_window.destroy()
        root.quit()

    Label(choice_window, text="Treffen Sie eine Auswahl um Shopware-Artikel zu pflegen").pack(pady=20)

    btn_multiple = Button(choice_window, text=" - Mehrere Bilder hochladen - ", command=on_multiple, width=20)
    btn_multiple.pack(pady=10)
    ToolTip(btn_multiple, "Lädt mehrere Bilder basierend\nauf einem angegebenen Datum zum Shop hoch.")

    btn_single = Button(choice_window, text=" - Einzelne Bild hochladen -", command=on_single, width=20)
    btn_single.pack(pady=10)
    ToolTip(btn_single, "Lädt ein einzelnes Bild basierend auf\neiner Artikelnummer zum Shop hoch.")

    btn_process_files = Button(choice_window, text=" - Shopdateien erstellen (.csv) - ", command=on_process_files, width=20)
    btn_process_files.pack(pady=10)
    ToolTip(btn_process_files, "Verarbeitet Exceldatei die mit Mention erstellt ist\nund Textdateien von Kosatec heruntergeladen ist,\num CSV-Dateien zu erstellen.")

    btn_mention_bilder = Button(choice_window, text=" - Mention Bilder - ", command=on_mention_bilder, width=20)
    btn_mention_bilder.pack(pady=10)
    ToolTip(btn_mention_bilder, "Erstellte Screenshots werden unter\nMention \ bilder gespeichert.")

    btn_exit = Button(choice_window, text="Beenden", command=on_exit, width=20)
    btn_exit.pack(pady=10)
    ToolTip(btn_exit, "Schließt das Programm.")

    choice_window.mainloop()

def main():
    global root
    root = Tk()
    root.withdraw()
    show_choice_window()
    root.mainloop()

if __name__ == "__main__":
    main()
