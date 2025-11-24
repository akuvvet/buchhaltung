import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import os
import csv
import re
import chardet
import xlrd

def create_tooltip(widget, text):
    toolTip = ToolTip(widget)
    def enter(event):
        toolTip.showtip(text)
    def leave(event):
        toolTip.hidetip()
    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)

class ToolTip(object):
    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0

    def showtip(self, text):
        "Display text in tooltip window"
        self.text = text
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 27
        y = y + cy + self.widget.winfo_rooty() + 27
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                         background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                         font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

# ---------------- Amazon Button Function ----------------
def amazon_function():
    def choose_file():
        file_path = filedialog.askopenfilename()
        return file_path

    def process_csv(file_path):
        df = pd.read_csv(file_path, delimiter=',', encoding='utf-8', header=7)
        df['Datum/Uhrzeit'] = pd.to_datetime(df['Datum/Uhrzeit'], format='%d.%m.%Y %H:%M:%S %Z').dt.strftime('%d.%m.%Y')

        columns_to_keep = [
            'Datum/Uhrzeit', 'Typ', 'Menge', 'Marketplace', 'Umsätze', 
            'Verkaufsgebühren', 'Andere Transaktionsgebühren', 'Andere', 'Gesamt'
        ]
        df = df[columns_to_keep]
        return df

    def generate_file_name(original_file_path):
        last_month_date = datetime.now() - relativedelta(months=1)
        date_part = last_month_date.strftime('%Y-%m')
        
        if "okay" in original_file_path.lower():
            return f"amazon-trans-okay-{date_part}.csv"
        elif "zone" in original_file_path.lower():
            return f"amazon-trans-zone-{date_part}.csv"
        else:
            return f"amazon-trans-{date_part}.csv"

    def save_csv(df, file_name):
        save_path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile=file_name, filetypes=[("CSV-Dateien", "*.csv")])  
        if save_path:
            df.to_csv(save_path, sep=';', index=False, encoding='utf-8')
            messagebox.showinfo("Erfolg", f"Datei gespeichert unter: {save_path}")
        else:
            messagebox.showwarning("Abbruch", "Speichern abgebrochen.")

    input_file_path = choose_file()
    if input_file_path:
        processed_df = process_csv(input_file_path)
        new_file_name = generate_file_name(os.path.basename(input_file_path).lower())
        save_csv(processed_df, new_file_name)
    else:
        messagebox.showwarning("Abbruch", "Keine Datei ausgewählt.")

# ---------------- eBay Button Function ----------------
def ebay_function():
    def choose_file():
        file_path = filedialog.askopenfilename()
        return file_path

    def extract_filename_part(file_path):
        with open(file_path, 'rb') as rawdata:
            result = chardet.detect(rawdata.read())
            encoding = result['encoding']

        with open(file_path, 'r', encoding=encoding) as file:
            for i, line in enumerate(file):
                if i == 8:  
                    return re.sub(r'[\\/*?:"<>|;]', '', line.split("Verkäufer")[-1].strip().strip('"'))

    def read_and_process_csv(file_path):
        try:
            df = pd.read_csv(file_path, delimiter=';', encoding='utf-8', skiprows=11, quoting=csv.QUOTE_NONE)
            df.columns = [col.replace('"', '') for col in df.columns]
            df = df.apply(lambda x: x.str.replace('"', '') if x.dtype == "object" else x)

            columns_to_keep = [
                'Datum der Transaktionserstellung',
                'Typ',
                'Name des Käufers',
                'Versandziel - Land',
                'Betrag abzügl. Kosten',
                'Vom Verkäufer angegebener MwSt.-Satz',
                'Fixer Anteil der Verkaufsprovision',
                'Variabler Anteil der Verkaufsprovision',
                'Gebühr für sehr hohe Quote an „nicht wie beschriebenen Artikeln“',
                'Gebühr für unterdurchschnittlichen Servicestatus',
                'Transaktionsbetrag (inkl. Kosten)'
            ]
            df = df[columns_to_keep]
            return df
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Lesen der Datei: {e}")
            return pd.DataFrame()

    def save_csv_file(df, file_name_part):
        last_month = datetime.now() - relativedelta(months=1)
        new_file_name = f"ebay-trans-{file_name_part}-{last_month.strftime('%Y-%m')}"

        save_path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile=new_file_name)
        
        if save_path:
            df.to_csv(save_path, sep=';', index=False, encoding='utf-8')
            messagebox.showinfo("Erfolg", f"Datei gespeichert unter: {save_path}")
        else:
            messagebox.showwarning("Abbruch", "Speichern abgebrochen.")

    file_path = choose_file()
    if file_path:
        file_name_part = extract_filename_part(file_path)
        df_processed = read_and_process_csv(file_path)
        save_csv_file(df_processed, file_name_part)
    else:
        messagebox.showwarning("Abbruch", "Keine Datei ausgewählt.")

# ---------------- Kaufland Button Function ----------------
def kaufland_function():
    def kopiere_daten():
        dateiname = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not dateiname:
            return

        ueberschriften = ["booking_date", "booking_text", "amount", "balance",
                          "price_gross", "shipping_charges_gross", "sum_price_gross", "fee_%", "fee_net", "fee_vat_%", "fee_gross", "shipping.country"]

        ziel_daten = [ueberschriften]

        spalten_map = {
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
            "shipping.country": "shipping.country"
        }

        with open(dateiname, newline='', encoding='UTF-8') as csvfile:
            quell_tabelle = csv.reader(csvfile, delimiter=';')
            ueberschriften_csv = next(quell_tabelle)
            
            for row in quell_tabelle:
                new_row = [None] * len(ueberschriften)
                for spalte_csv, ziel_spalte in spalten_map.items():
                    if spalte_csv in ueberschriften_csv:
                        index = ueberschriften_csv.index(spalte_csv)
                        ziel_index = ueberschriften.index(ziel_spalte)
                        cell_value = row[index].strip()
        
                        if spalte_csv == "booking_date":
                            try:
                                datetime_obj = datetime.strptime(cell_value, '%d.%m.%Y %H:%M')
                                cell_value = datetime_obj.strftime('%d.%m.%Y')
                            except ValueError:
                                cell_value = None 

                        new_row[ziel_index] = cell_value

                ziel_daten.append(new_row)

        heute = datetime.now()
        erster_dieses_monats = heute.replace(day=1)
        letzter_vorheriger_monat = erster_dieses_monats - timedelta(days=1)
        vorschlag_dateiname = letzter_vorheriger_monat.strftime("%Y-%m")

        speicherpfad = filedialog.asksaveasfilename(defaultextension=".csv",
                                                    filetypes=[("CSV files", "*.csv")],
                                                    initialfile=os.path.splitext(os.path.basename(dateiname))[0] + "-" + vorschlag_dateiname)
        if not speicherpfad:
            return

        try:
            with open(speicherpfad, 'w', newline='', encoding='UTF-8') as csvfile:
                csvwriter = csv.writer(csvfile, delimiter=';')
                csvwriter.writerows(ziel_daten)
            messagebox.showinfo("Erfolg", f"Datei gespeichert als {speicherpfad}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten: {e}")

    kopiere_daten()

# ---------------- Mention Ausgang Button Function ----------------
def mention_ausgang_function():
    def kopiere_daten():
        dateiname = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls"), ("Excel files", "*.xlsx")])
        if not dateiname:
            return

        quelle = xlrd.open_workbook(dateiname)
        quell_tabelle = quelle.sheet_by_index(0)

        ueberschriften = ["Umsatz in Euro", "Steuerschlüssel", "Gegenkonto", "Beleg1",
                          "Beleg2", "Datum", "Konto", "Kost1", "Kost2", "Skonto in Euro",
                          "Buchungstext", "Umsatzsteuer-ID", "Zusatzart", "Zusatzinformation"]

        spalten_map = {"Brutto Gesamt": "Umsatz in Euro", "Beleg-Nr.": "Beleg1",
                       "Datum": "Datum", "Name 2": "Buchungstext", "Umsatzsteuer-ID": "Umsatzsteuer-ID"}

        daten = [ueberschriften]

        for row_idx in range(1, quell_tabelle.nrows):
            row = quell_tabelle.row_values(row_idx)
            new_row = [None] * len(ueberschriften)
            for spalte, ziel_spalte in spalten_map.items():
                if spalte in quell_tabelle.row_values(0):
                    index = quell_tabelle.row_values(0).index(spalte)
                    ziel_index = ueberschriften.index(ziel_spalte)
                    cell_value = row[index]

                    if spalte == "Datum" and isinstance(cell_value, float):
                        date_tuple = xlrd.xldate_as_tuple(cell_value, quelle.datemode)
                        cell_value = datetime(*date_tuple).strftime("%d.%m.%Y")

                    if spalte == "Beleg-Nr." and isinstance(cell_value, float):
                        cell_value = int(cell_value)

                    new_row[ziel_index] = cell_value

            new_row[2] = "3400"
            new_row[6] = "1000"
            daten.append(new_row)

        heute = datetime.now()
        erster_dieses_monats = heute.replace(day=1)
        letzter_vorheriger_monat = erster_dieses_monats - timedelta(days=1)
        vorschlag_dateiname = letzter_vorheriger_monat.strftime("%Y-%m")

        basisname = os.path.splitext(os.path.basename(dateiname))[0].lower()
        if "okay" in basisname:
            vorschlag_dateiname = f"ausgang-mention-okay-{vorschlag_dateiname}.csv"
        elif "zone" in basisname:
            vorschlag_dateiname = f"ausgang-mention-zone-{vorschlag_dateiname}.csv"
        else:
            vorschlag_dateiname = basisname + "-" + vorschlag_dateiname + ".csv"

        speicherpfad = filedialog.asksaveasfilename(defaultextension=".csv",
                                                    filetypes=[("CSV files", "*.csv")],
                                                    initialfile=vorschlag_dateiname)
        if not speicherpfad:
            return

        with open(speicherpfad, 'w', newline='', encoding='utf-8') as csvfile:
            csvwriter = csv.writer(csvfile, delimiter=';')
            csvwriter.writerows(daten)

        messagebox.showinfo("Erfolg", f"Datei gespeichert als {speicherpfad}")

    kopiere_daten()

# ---------------- Mention Eingang Button Function ----------------
def mention_eingang_function():
    def kopiere_daten():
        dateiname = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls"), ("Excel files", "*.xlsx")])
        if not dateiname:
            return

        quelle = xlrd.open_workbook(dateiname)
        quell_tabelle = quelle.sheet_by_index(0)

        ueberschriften = ["Umsatz in Euro", "Steuerschlüssel", "Gegenkonto", "Beleg1",
                          "Beleg2", "Datum", "Konto", "Kost1", "Kost2", "Skonto in Euro",
                          "Buchungstext", "Umsatzsteuer-ID", "Zusatzart", "Zusatzinformation"]

        spalten_map = {"Betrag MW": "Umsatz in Euro", "Beleg-Nr.": "Beleg1",
                       "Datum": "Datum", "Name 2": "Buchungstext"}

        daten = [ueberschriften]

        for row_idx in range(1, quell_tabelle.nrows):
            row = quell_tabelle.row_values(row_idx)
            new_row = [None] * len(ueberschriften)
            for spalte, ziel_spalte in spalten_map.items():
                if spalte in quell_tabelle.row_values(0):
                    index = quell_tabelle.row_values(0).index(spalte)
                    ziel_index = ueberschriften.index(ziel_spalte)
                    cell_value = row[index]

                    if spalte == "Datum" and isinstance(cell_value, float):
                        date_tuple = xlrd.xldate_as_tuple(cell_value, quelle.datemode)
                        cell_value = datetime(*date_tuple).strftime("%d.%m.%Y")

                    new_row[ziel_index] = cell_value

            new_row[2] = "3400"
            new_row[6] = "1000"
            daten.append(new_row)

        heute = datetime.now()
        erster_dieses_monats = heute.replace(day=1)
        letzter_vorheriger_monat = erster_dieses_monats - timedelta(days=1)
        vorschlag_dateiname = letzter_vorheriger_monat.strftime("%Y-%m")

        if "okay" in dateiname.lower():
            vorschlag_dateiname = f"eingang-mention-okay-{vorschlag_dateiname}.csv"
        elif "zone" in dateiname.lower():
            vorschlag_dateiname = f"eingang-mention-zone-{vorschlag_dateiname}.csv"
        else:
            vorschlag_dateiname = os.path.splitext(os.path.basename(dateiname))[0] + "-" + vorschlag_dateiname + ".csv"

        speicherpfad = filedialog.asksaveasfilename(defaultextension=".csv",
                                                    filetypes=[("CSV files", "*.csv")],
                                                    initialfile=vorschlag_dateiname)
        if not speicherpfad:
            return

        with open(speicherpfad, 'w', newline='', encoding='utf-8') as csvfile:
            csvwriter = csv.writer(csvfile, delimiter=';')
            csvwriter.writerows(daten)

        messagebox.showinfo("Erfolg", f"Datei gespeichert als {speicherpfad}")

    kopiere_daten()

# ---------------- Sale Ausgang Button Function ----------------
def sale_ausgang_function():
    def kopiere_daten():
        dateiname = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not dateiname:
            return

        ueberschriften = ["Umsatz in Euro", "Steuerschlüssel", "Gegenkonto", "Beleg1",
                          "Beleg2", "Datum", "Konto", "Kost1", "Kost2", "Skonto in Euro",
                          "Buchungstext", "Umsatzsteuer-ID", "Zusatzart", "Zusatzinformation"]

        ziel_daten = [ueberschriften]

        spalten_map = {
            "Gesamtbruttobetrag (inkl. Versand)": "Umsatz in Euro",
            "Dok.-Nr.": "Beleg1",
            "Dok. Datum": "Datum",
            "Kundenname": "Buchungstext"
        }

        with open(dateiname, newline='', encoding='cp1252') as csvfile:
            quell_tabelle = csv.reader(csvfile, delimiter=';')
            ueberschriften_csv = next(quell_tabelle)
        
            for row in quell_tabelle:
                new_row = [None] * len(ueberschriften)
                for spalte_csv, ziel_spalte in spalten_map.items():
                    if spalte_csv in ueberschriften_csv:
                        index = ueberschriften_csv.index(spalte_csv)
                        ziel_index = ueberschriften.index(ziel_spalte)
                        cell_value = row[index].strip()
        
                        if spalte_csv == "Rg-Datum":
                            try:
                                cell_value = datetime.strptime(cell_value, "%d.%m.%Y").strftime("%d.%m.%Y")
                            except ValueError:
                                cell_value = None 

                        new_row[ziel_index] = cell_value

                new_row[2] = "8400"
                new_row[6] = "1000"
                ziel_daten.append(new_row)

        heute = datetime.now()
        erster_dieses_monats = heute.replace(day=1)
        letzter_vorheriger_monat = erster_dieses_monats - timedelta(days=1)
        vorschlag_dateiname = letzter_vorheriger_monat.strftime("%Y-%m")

        speicherpfad = filedialog.asksaveasfilename(defaultextension=".csv",
                                                    filetypes=[("CSV files", "*.csv")],
                                                    initialfile=os.path.splitext(os.path.basename(dateiname))[0] + "-" + vorschlag_dateiname)
        if not speicherpfad:
            return

        try:
            with open(speicherpfad, 'w', newline='', encoding='cp1252') as csvfile:
                csvwriter = csv.writer(csvfile, delimiter=';')
                csvwriter.writerows(ziel_daten)
            messagebox.showinfo("Erfolg", f"Datei gespeichert als {speicherpfad}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten: {e}")

    kopiere_daten()

# Hauptfenster
root = tk.Tk()
root.title("Datenverarbeitung")

# Amazon Button
amazon_button = tk.Button(root, text="Amazon-Transaction-CSV", command=amazon_function)
amazon_button.pack(pady=20)
create_tooltip(amazon_button, "Verarbeitet und speichert Amazon CSV-Dateien")

# eBay Button
ebay_button = tk.Button(root, text="eBay-Transaction-CSV", command=ebay_function)
ebay_button.pack(pady=20)
create_tooltip(ebay_button, "Verarbeitet und speichert eBay CSV-Dateien")

# Kaufland Button
kaufland_button = tk.Button(root, text="Kaufland-Transaction-CSV", command=kaufland_function)
kaufland_button.pack(pady=20)
create_tooltip(kaufland_button, "Verarbeitet und speichert Kaufland CSV-Dateien")

# Mention Ausgang Button
mention_ausgang_button = tk.Button(root, text="Mention Ausgangsrechnungen", command=mention_ausgang_function)
mention_ausgang_button.pack(pady=20)
create_tooltip(mention_ausgang_button, "Verarbeitet und speichert Mention Ausgang Excel-Dateien")

# Mention Eingang Button
mention_eingang_button = tk.Button(root, text="Mention Eingangsrechnungen", command=mention_eingang_function)
mention_eingang_button.pack(pady=20)
create_tooltip(mention_eingang_button, "Verarbeitet und speichert Mention Eingang Excel-Dateien")

# Sale Ausgang Button
sale_ausgang_button = tk.Button(root, text="Sale Ausgangsrechnungen", command=sale_ausgang_function)
sale_ausgang_button.pack(pady=20)
create_tooltip(sale_ausgang_button, "Verarbeitet und speichert Sale Ausgang CSV-Dateien")

# Beenden Button
exit_button = tk.Button(root, text="Beenden", command=root.quit)
exit_button.pack(pady=20)
create_tooltip(exit_button, "Beendet das Programm")

root.mainloop()
