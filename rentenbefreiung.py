import sys
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple

try:
	from tkinter import Tk, messagebox
	from tkinter.filedialog import askopenfilename, asksaveasfilename
	from tkinter import simpledialog
	from tkinter import Toplevel, Label, Entry, Button, Checkbutton, StringVar, BooleanVar
except Exception:
	print("Fehler beim Laden von tkinter. Bitte stellen Sie sicher, dass tkinter installiert/verfügbar ist.")
	raise

try:
	import win32com.client as win32
except ImportError:
	print("Das Paket 'pywin32' wird benötigt. Bitte installieren mit: pip install pywin32")
	sys.exit(1)


DATE_FMT = "%d.%m.%Y"


def parse_date_ddmmyyyy(value: str, *, allow_today_default: bool = False) -> str:
	"""
	Validiert ein Datums-String im Format DD.MM.JJJJ und gibt exakt dieses Format zurück.
	Bei allow_today_default=True wird ein leerer String als 'heute' interpretiert.
	"""
	value = (value or "").strip()
	if not value and allow_today_default:
		return datetime.today().strftime(DATE_FMT)
	try:
		dt = datetime.strptime(value, DATE_FMT)
		return dt.strftime(DATE_FMT)
	except ValueError:
		raise ValueError("Bitte Datum im Format TT.MM.JJJJ angeben (z. B. 25.11.2025).")


def ask_user_inputs() -> Tuple[str, str, str, str, str, str]:
	"""
	Fragt die benötigten Werte per Windows-Popup ab.
	Gibt zurück:
	- familienname
	- vorname
	- rvnr
	- ort
	- aktuelles_datum (DD.MM.JJJJ)
	- beginn_befreiung (DD.MM.JJJJ)
	"""
	root = Tk()
	root.withdraw()
	root.wm_attributes("-topmost", 1)

	# Arbeitnehmer
	familienname = simpledialog.askstring("Eingabe", "Familienname:", parent=root)
	if not familienname:
		raise ValueError("Familienname darf nicht leer sein.")

	vorname = simpledialog.askstring("Eingabe", "Vorname:", parent=root)
	if not vorname:
		raise ValueError("Vorname darf nicht leer sein.")

	rvnr = simpledialog.askstring("Eingabe", "Rentenversicherungsnr:", parent=root)
	if not rvnr:
		raise ValueError("Rentenversicherungsnr darf nicht leer sein.")

	# Arbeitgeber
	ort = simpledialog.askstring("Eingabe", "Ort:", initialvalue="Solingen", parent=root)
	if not ort:
		raise ValueError("Ort darf nicht leer sein.")

	aktuelles_datum_raw = simpledialog.askstring(
		"Eingabe", "Aktuelles Datum (TT.MM.JJJJ) – leer = heute:", parent=root
	)
	aktuelles_datum = parse_date_ddmmyyyy(aktuelles_datum_raw or "", allow_today_default=True)

	beginn_befreiung_raw = simpledialog.askstring(
		"Eingabe", "Befreiung beginnt am (TT.MM.JJJJ):", parent=root
	)
	if not beginn_befreiung_raw:
		raise ValueError("Befreiung beginnt am darf nicht leer sein.")
	beginn_befreiung = parse_date_ddmmyyyy(beginn_befreiung_raw)

	return familienname.strip(), vorname.strip(), rvnr.strip(), ort.strip(), aktuelles_datum, beginn_befreiung


def ask_excel_path() -> Path:
	root = Tk()
	root.withdraw()
	root.wm_attributes("-topmost", 1)
	filename = askopenfilename(
		title="Excel-Datei auswählen",
		filetypes=[("Excel-Dateien", "*.xlsx *.xlsm *.xls"), ("Alle Dateien", "*.*")],
	)
	if not filename:
		raise ValueError("Es wurde keine Excel-Datei ausgewählt.")
	return Path(filename)


def ask_all_inputs_single_form() -> Tuple[str, str, str, str, str, str, Path, Path, Optional[Path]]:
	"""
	Ein einziges Popup-Formular für alle Eingaben inkl. Datei-Auswahl.
	Rückgabe:
	(familienname, vorname, rvnr, ort, aktuelles_datum, beginn_befreiung, excel_path, pdf_path, signature_path_or_none)
	"""
	root = Tk()
	root.withdraw()
	root.wm_attributes("-topmost", 1)

	top = Toplevel(root)
	top.title("Eingaben")
	top.wm_attributes("-topmost", 1)

	# String-/Bool-Variablen
	var_familienname = StringVar()
	var_vorname = StringVar()
	var_rvnr = StringVar()
	var_ort = StringVar(value="Solingen")
	var_datum = StringVar()  # leer = heute
	var_beginn = StringVar()
	var_excel = StringVar()
	var_pdf = StringVar()
	var_sig_enabled = BooleanVar(value=False)
	var_sig_path = StringVar()

	# Layout
	row = 0
	Label(top, text="Familienname:").grid(row=row, column=0, sticky="e", padx=6, pady=4)
	Entry(top, textvariable=var_familienname, width=40).grid(row=row, column=1, columnspan=2, sticky="we", padx=6, pady=4)
	row += 1

	Label(top, text="Vorname:").grid(row=row, column=0, sticky="e", padx=6, pady=4)
	Entry(top, textvariable=var_vorname, width=40).grid(row=row, column=1, columnspan=2, sticky="we", padx=6, pady=4)
	row += 1

	Label(top, text="Rentenversicherungsnr:").grid(row=row, column=0, sticky="e", padx=6, pady=4)
	Entry(top, textvariable=var_rvnr, width=40).grid(row=row, column=1, columnspan=2, sticky="we", padx=6, pady=4)
	row += 1

	Label(top, text="Ort:").grid(row=row, column=0, sticky="e", padx=6, pady=4)
	Entry(top, textvariable=var_ort, width=40).grid(row=row, column=1, columnspan=2, sticky="we", padx=6, pady=4)
	row += 1

	Label(top, text="Aktuelles Datum (TT.MM.JJJJ, leer = heute):").grid(row=row, column=0, sticky="e", padx=6, pady=4)
	Entry(top, textvariable=var_datum, width=40).grid(row=row, column=1, columnspan=2, sticky="we", padx=6, pady=4)
	row += 1

	Label(top, text="Befreiung beginnt am (TT.MM.JJJJ):").grid(row=row, column=0, sticky="e", padx=6, pady=4)
	Entry(top, textvariable=var_beginn, width=40).grid(row=row, column=1, columnspan=2, sticky="we", padx=6, pady=4)
	row += 1

	Label(top, text="Excel-Datei:").grid(row=row, column=0, sticky="e", padx=6, pady=4)
	entry_excel = Entry(top, textvariable=var_excel, width=34)
	entry_excel.grid(row=row, column=1, sticky="we", padx=6, pady=4)
	def on_browse_excel():
		fn = askopenfilename(title="Excel-Datei auswählen", filetypes=[("Excel-Dateien", "*.xlsx *.xlsm *.xls"), ("Alle Dateien", "*.*")])
		if fn:
			var_excel.set(fn)
	Button(top, text="Durchsuchen…", command=on_browse_excel).grid(row=row, column=2, sticky="w", padx=6, pady=4)
	row += 1

	Label(top, text="PDF speichern unter:").grid(row=row, column=0, sticky="e", padx=6, pady=4)
	entry_pdf = Entry(top, textvariable=var_pdf, width=34)
	entry_pdf.grid(row=row, column=1, sticky="we", padx=6, pady=4)
	def on_browse_pdf():
		default = Path(var_excel.get()).with_suffix(".pdf").name if var_excel.get() else "Ausgabe.pdf"
		fn = asksaveasfilename(title="PDF-Ziel speichern unter", defaultextension=".pdf", initialfile=default, filetypes=[("PDF-Datei", "*.pdf")])
		if fn:
			var_pdf.set(fn)
	Button(top, text="Durchsuchen…", command=on_browse_pdf).grid(row=row, column=2, sticky="w", padx=6, pady=4)
	row += 1

	cb = Checkbutton(top, text="Unterschriftsbild einfügen", variable=var_sig_enabled)
	cb.grid(row=row, column=0, columnspan=3, sticky="w", padx=6, pady=(10,4))
	row += 1

	Label(top, text="Unterschrift-Datei:").grid(row=row, column=0, sticky="e", padx=6, pady=4)
	entry_sig = Entry(top, textvariable=var_sig_path, width=34, state="disabled")
	entry_sig.grid(row=row, column=1, sticky="we", padx=6, pady=4)
	def on_browse_sig():
		fn = askopenfilename(title="Unterschriftsbild auswählen", filetypes=[("Bilddateien", "*.png *.jpg *.jpeg *.bmp *.gif"), ("Alle Dateien", "*.*")])
		if fn:
			var_sig_path.set(fn)
	def on_sig_toggle(*_args):
		entry_sig.configure(state="normal" if var_sig_enabled.get() else "disabled")
	Button(top, text="Durchsuchen…", command=on_browse_sig).grid(row=row, column=2, sticky="w", padx=6, pady=4)
	row += 1
	var_sig_enabled.trace_add("write", on_sig_toggle)

	# Buttons
	result = {"ok": False}

	def on_ok():
		try:
			fam = (var_familienname.get() or "").strip()
			vor = (var_vorname.get() or "").strip()
			rv = (var_rvnr.get() or "").strip()
			ort = (var_ort.get() or "").strip()
			akt_raw = (var_datum.get() or "").strip()
			beg_raw = (var_beginn.get() or "").strip()
			excel_p = (var_excel.get() or "").strip()
			pdf_p = (var_pdf.get() or "").strip()

			if not fam:
				raise ValueError("Familienname darf nicht leer sein.")
			if not vor:
				raise ValueError("Vorname darf nicht leer sein.")
			if not rv:
				raise ValueError("Rentenversicherungsnr darf nicht leer sein.")
			if not ort:
				raise ValueError("Ort darf nicht leer sein.")
			if not beg_raw:
				raise ValueError("Befreiung beginnt am darf nicht leer sein.")
			if not excel_p:
				raise ValueError("Bitte Excel-Datei auswählen.")
			if not pdf_p:
				raise ValueError("Bitte PDF-Ziel auswählen.")

			aktuelles_datum = parse_date_ddmmyyyy(akt_raw or "", allow_today_default=True)
			beginn_befreiung = parse_date_ddmmyyyy(beg_raw)

			signature_path = None
			if var_sig_enabled.get():
				sig_p = (var_sig_path.get() or "").strip()
				if not sig_p:
					raise ValueError("Bitte Unterschriftsbild auswählen oder die Option deaktivieren.")
				signature_path = Path(sig_p)

			result.update({
				"ok": True,
				"familienname": fam,
				"vorname": vor,
				"rvnr": rv,
				"ort": ort,
				"aktuelles_datum": aktuelles_datum,
				"beginn_befreiung": beginn_befreiung,
				"excel_path": Path(excel_p),
				"pdf_path": Path(pdf_p),
				"signature_path": signature_path
			})
			top.destroy()
		except Exception as err:
			messagebox.showerror("Eingabefehler", str(err), parent=top)

	def on_cancel():
		top.destroy()

	btn_frame_row = row
	Button(top, text="OK", width=12, command=on_ok).grid(row=btn_frame_row, column=1, sticky="e", padx=6, pady=(10,8))
	Button(top, text="Abbrechen", width=12, command=on_cancel).grid(row=btn_frame_row, column=2, sticky="w", padx=6, pady=(10,8))

	# Fokus
	top.grab_set()
	top.protocol("WM_DELETE_WINDOW", on_cancel)
	top.resizable(False, False)

	root.wait_window(top)

	if not result.get("ok"):
		raise ValueError("Abgebrochen.")

	return (
		result["familienname"],
		result["vorname"],
		result["rvnr"],
		result["ort"],
		result["aktuelles_datum"],
		result["beginn_befreiung"],
		result["excel_path"],
		result["pdf_path"],
		result["signature_path"],
	)

def ask_pdf_path(default_from_excel: Path) -> Path:
	root = Tk()
	root.withdraw()
	root.wm_attributes("-topmost", 1)
	default_pdf_name = default_from_excel.with_suffix(".pdf").name
	filename = asksaveasfilename(
		title="PDF-Ziel speichern unter",
		defaultextension=".pdf",
		initialfile=default_pdf_name,
		filetypes=[("PDF-Datei", "*.pdf")],
	)
	if not filename:
		raise ValueError("Es wurde kein Speicherort für die PDF-Datei gewählt.")
	return Path(filename)


def ask_signature_path_optional() -> Optional[Path]:
	"""
	Fragt optional nach einem Unterschriftsbild. Bei Abbruch wird None zurückgegeben.
	"""
	root = Tk()
	root.withdraw()
	root.wm_attributes("-topmost", 1)

	if not messagebox.askyesno("Unterschrift", "Möchten Sie ein Unterschriftsbild einfügen?"):
		return None

	filename = askopenfilename(
		title="Unterschriftsbild auswählen",
		filetypes=[
			("Bilddateien", "*.png *.jpg *.jpeg *.bmp *.gif"),
			("Alle Dateien", "*.*"),
		],
	)
	if not filename:
		return None
	return Path(filename)


def fill_excel_and_export_pdf(
	excel_path: Path,
	pdf_path: Path,
	familienname: str,
	vorname: str,
	rvnr: str,
	ort: str,
	aktuelles_datum: str,
	beginn_befreiung: str,
	signature_path: Optional[Path] = None,
) -> None:
	"""
	Öffnet Excel per COM, trägt Werte ein und exportiert als PDF.
	Excel-Formatierungen werden nicht verändert; es werden Strings gesetzt.
	Die Arbeitsmappe wird NICHT gespeichert.
	"""
	excel = None
	wb = None
	try:
		excel = win32.gencache.EnsureDispatch("Excel.Application")
		excel.Visible = False
		excel.DisplayAlerts = False

		wb = excel.Workbooks.Open(str(excel_path))
		ws = wb.Worksheets(1)  # Erstes Arbeitsblatt

		# Einträge in die vorgegebenen Zellen
		ws.Range("C10").Value = familienname
		ws.Range("C10").Font.Bold = True
		ws.Range("C12").Value = vorname
		ws.Range("C12").Font.Bold = True
		ws.Range("C14").Value = rvnr
		ws.Range("C14").Font.Bold = True

		ort_komma_datum = f"{ort}, {aktuelles_datum}"
		ws.Range("B26").Value = ort_komma_datum
		ws.Range("B26").Font.Bold = True
		ws.Range("B40").Value = ort_komma_datum
		ws.Range("B40").Font.Bold = True

		ws.Range("E36").Value = aktuelles_datum
		ws.Range("E36").Font.Bold = True
		ws.Range("D38").Value = beginn_befreiung
		ws.Range("D38").Font.Bold = True

		# Unterschriftsbild optional oberhalb von D27 platzieren
		if signature_path is not None and signature_path.exists():
			try:
				target_cell = ws.Range("D27")
				cell_left = target_cell.Left
				cell_top = target_cell.Top

				# Bild als schwebendes Objekt einfügen (keine Skalierung)
				pic = ws.Pictures().Insert(str(signature_path))

				# Links an Spalte D ausrichten, oberhalb der Zelle D27 positionieren
				pic.Left = cell_left
				new_top = cell_top - pic.Height
				pic.Top = new_top if new_top > 0 else 0
			except Exception:
				# Falls das Einfügen fehlschlägt, Vorgang ohne Unterschrift fortsetzen
				pass

		# Export als PDF (nur aktives/erstes Blatt)
		# Type=0 -> xlTypePDF
		ws.ExportAsFixedFormat(
			Type=0,
			Filename=str(pdf_path),
			Quality=0,  # xlQualityStandard
			IncludeDocProperties=True,
			IgnorePrintAreas=False,
			OpenAfterPublish=False,
		)
	finally:
		# Workbooks ohne Speichern schließen
		if wb is not None:
			try:
				wb.Close(SaveChanges=False)
			except Exception:
				pass
		if excel is not None:
			try:
				excel.Quit()
			except Exception:
				pass


def main() -> None:
	try:
		# Alle Eingaben in einem einzigen Formular abfragen
		(
			familienname,
			vorname,
			rvnr,
			ort,
			aktuelles_datum,
			beginn_befreiung,
			excel_path,
			pdf_path,
			signature_path,
		) = ask_all_inputs_single_form()

		# Sicherstellen, dass Zielverzeichnis existiert
		pdf_path.parent.mkdir(parents=True, exist_ok=True)

		fill_excel_and_export_pdf(
			excel_path=excel_path,
			pdf_path=pdf_path,
			familienname=familienname,
			vorname=vorname,
			rvnr=rvnr,
			ort=ort,
			aktuelles_datum=aktuelles_datum,
			beginn_befreiung=beginn_befreiung,
			signature_path=signature_path,
		)

		root = Tk()
		root.withdraw()
		root.wm_attributes("-topmost", 1)
		messagebox.showinfo("Fertig", f"PDF wurde erstellt:\n{pdf_path}")
	except Exception as e:
		root = Tk()
		root.withdraw()
		root.wm_attributes("-topmost", 1)
		messagebox.showerror("Fehler", str(e))
		# Zusätzlich auf stdout für Logs:
		print(f"Fehler: {e}", file=sys.stderr)
		sys.exit(1)


if __name__ == "__main__":
	main()


