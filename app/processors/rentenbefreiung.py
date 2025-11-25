from __future__ import annotations

from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Optional, Tuple

try:
	import win32com.client as win32  # type: ignore
	import pythoncom  # type: ignore
except Exception as e:  # pragma: no cover
	# Auf Windows ist pywin32 erforderlich, um Excel zu automatisieren.
	# Die Route wird andernfalls mit einer klaren Fehlermeldung abbrechen.
	raise RuntimeError("pywin32 (win32com) ist erforderlich, um Excel zu automatisieren.") from e


DATE_FMT = "%d.%m.%Y"


def parse_date_ddmmyyyy(value: str, *, allow_today_default: bool = False) -> str:
	value = (value or "").strip()
	if not value and allow_today_default:
		return datetime.today().strftime(DATE_FMT)
	try:
		dt = datetime.strptime(value, DATE_FMT)
		return dt.strftime(DATE_FMT)
	except ValueError:
		raise ValueError("Bitte Datum im Format TT.MM.JJJJ angeben (z. B. 25.11.2025).")


def _fill_excel_and_export_pdf(
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
	excel = None
	wb = None
	try:
		# COM in diesem (Request-)Thread initialisieren
		pythoncom.CoInitialize()

		excel = win32.gencache.EnsureDispatch("Excel.Application")
		excel.Visible = False
		excel.DisplayAlerts = False
		wb = excel.Workbooks.Open(str(excel_path))
		ws = wb.Worksheets(1)

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

		if signature_path is not None and signature_path.exists():
			try:
				target_cell = ws.Range("D27")
				cell_left = target_cell.Left
				cell_top = target_cell.Top
				pic = ws.Pictures().Insert(str(signature_path))
				# Maximalmaße: 5 cm Breite, 2 cm Höhe (in Punkte umrechnen und proportional skalieren)
				POINTS_PER_INCH = 72.0
				CM_PER_INCH = 2.54
				max_width_pt = (5.0 / CM_PER_INCH) * POINTS_PER_INCH
				max_height_pt = (2.0 / CM_PER_INCH) * POINTS_PER_INCH
				try:
					cur_w = float(pic.Width)
					cur_h = float(pic.Height)
					if cur_w > 0.0 and cur_h > 0.0:
						scale = min(max_width_pt / cur_w, max_height_pt / cur_h, 1.0)
						if scale < 1.0:
							pic.Width = cur_w * scale
							pic.Height = cur_h * scale
				except Exception:
					# Falls Größenanpassung fehlschlägt, ohne Skalierung fortfahren
					pass
				pic.Left = cell_left
				new_top = cell_top - pic.Height
				pic.Top = new_top if new_top > 0 else 0
			except Exception:
				# Wenn Einfügen scheitert, einfach ohne Unterschrift fortfahren
				pass

		ws.ExportAsFixedFormat(
			Type=0,  # xlTypePDF
			Filename=str(pdf_path),
			Quality=0,  # xlQualityStandard
			IncludeDocProperties=True,
			IgnorePrintAreas=False,
			OpenAfterPublish=False,
		)
	finally:
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
		# COM wieder freigeben
		try:
			pythoncom.CoUninitialize()
		except Exception:
			pass


def export_rentenbefreiung_pdf(
	excel_bytes: bytes,
	familienname: str,
	vorname: str,
	rvnr: str,
	ort: str,
	aktuelles_datum: str,
	beginn_befreiung: str,
	signature_bytes: Optional[bytes] = None,
) -> Tuple[bytes, str]:
	"""
	Nimmt das Excel-Template als Bytes entgegen, befüllt es per COM und liefert PDF-Bytes zurück.
	"""
	familienname_clean = (familienname or "").strip()
	vorname_clean = (vorname or "").strip()
	if not familienname_clean:
		raise ValueError("Familienname darf nicht leer sein.")
	if not vorname_clean:
		raise ValueError("Vorname darf nicht leer sein.")
	if not (rvnr or "").strip():
		raise ValueError("Rentenversicherungsnr darf nicht leer sein.")
	ort_clean = (ort or "").strip() or "Solingen"
	aktuelles_datum_norm = parse_date_ddmmyyyy(aktuelles_datum, allow_today_default=True)
	beginn_befreiung_norm = parse_date_ddmmyyyy(beginn_befreiung, allow_today_default=False)

	with TemporaryDirectory(prefix="rentenbefreiung_") as tmpdir:
		tmp_dir = Path(tmpdir)
		excel_path = tmp_dir / "vorlage.xlsx"
		pdf_path = tmp_dir / "ausgabe.pdf"
		excel_path.write_bytes(excel_bytes)
		signature_path: Optional[Path] = None
		if signature_bytes:
			signature_path = tmp_dir / "signature.png"
			signature_path.write_bytes(signature_bytes)

		_fill_excel_and_export_pdf(
			excel_path=excel_path,
			pdf_path=pdf_path,
			familienname=familienname_clean,
			vorname=vorname_clean,
			rvnr=(rvnr or "").strip(),
			ort=ort_clean,
			aktuelles_datum=aktuelles_datum_norm,
			beginn_befreiung=beginn_befreiung_norm,
			signature_path=signature_path,
		)

		pdf_bytes = pdf_path.read_bytes()

		filename = f"Rentenbefreiung_{familienname_clean}_{vorname_clean}.pdf"
		return pdf_bytes, filename


