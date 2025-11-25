from flask import Flask, render_template, request, send_file, abort, redirect, url_for, session
from io import BytesIO
from datetime import datetime, timedelta
import os
from functools import wraps
from werkzeug.security import check_password_hash, generate_password_hash
from time import time

from app.processors.amazon import process_amazon_csv
from app.processors.ebay import process_ebay_csv
from app.processors.kaufland import process_kaufland_csv
from app.processors.mention_ausgang import process_mention_ausgang_excel
from app.processors.mention_eingang import process_mention_eingang_excel
from app.processors.sale_ausgang import process_sale_ausgang_csv
from app.processors.mention_shop import process_shop_files
from app.processors.ftp_upload import upload_images_and_generate_csv
from app.processors.apetito_web import process_apetito_excel


def create_app() -> Flask:
    app = Flask(__name__)
    # Session Secret (für Login). In PROD per ENV setzen.
    app.secret_key = os.environ.get("SECRET_KEY", "change-me-in-prod")
    # 30 Minuten Inaktivität
    app.permanent_session_lifetime = timedelta(
        seconds=int(os.environ.get("SESSION_IDLE_TIMEOUT_SECONDS", "1800"))
    )

    # Einfache Auth: fester Benutzer und gehashter Passwort-String
    ALLOWED_EMAIL = os.environ.get("OKAYTOOL_USER", "info@okaycomputer.de").lower()
    # Hinweis: Im Produktivbetrieb OKAYTOOL_PASSWORD_HASH als ENV setzen und
    # den Klartext nicht im Code hinterlegen.
    DEFAULT_HASH = generate_password_hash("OKAY1993!")  # fallback
    PASSWORD_HASH = os.environ.get("OKAYTOOL_PASSWORD_HASH", DEFAULT_HASH)

    def login_required(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            if not session.get("user_email"):
                return redirect(url_for("login", next=request.path))
            return fn(*args, **kwargs)
        return wrapper
    @app.before_request
    def _session_timeout_guard():
        if not session.get("user_email"):
            return
        last_seen = session.get("last_seen")
        now = int(time())
        timeout_seconds = int(os.environ.get("SESSION_IDLE_TIMEOUT_SECONDS", "1800"))
        if last_seen and now - int(last_seen) > timeout_seconds:
            session.clear()
            return redirect(url_for("login", expired=1))
        session["last_seen"] = now

    def _last_month_folder() -> str:
        # YYYYMM des Vormonats, z.B. 202510 für Oktober 2025 wenn aktueller Monat Nov 2025 ist
        now = datetime.now()
        year = now.year
        month = now.month - 1
        if month == 0:
            month = 12
            year -= 1
        return f"{year:04d}{month:02d}"

    def _network_target_dir(original_filename: str | None) -> str | None:
        name = (original_filename or "").lower()
        base_dir = None
        if "okay" in name:
            base_dir = r"O:\buchhaltung\berichte"
        elif "zone" in name or "best" in name:
            base_dir = r"Z:\buchhaltung\berichte"
        if not base_dir:
            return None
        return os.path.join(base_dir, _last_month_folder(), "sonstiges")

    def _maybe_save_to_network(result_bytes: bytes, filename: str, original_filename: str | None):
        target_dir = _network_target_dir(original_filename)
        if not target_dir:
            return
        try:
            os.makedirs(target_dir, exist_ok=True)
            target_path = os.path.join(target_dir, filename or "export.csv")
            with open(target_path, "wb") as fh:
                fh.write(result_bytes)
        except Exception:
            # Netzwerkpfad nicht verfügbar oder keine Berechtigung -> ignorieren
            pass

    @app.get("/")
    def index():
        if not session.get("user_email"):
            return redirect(url_for("login"))
        return redirect(url_for("transaktion_page"))

    @app.get("/login")
    def login():
        msg = "Sitzung abgelaufen. Bitte erneut anmelden." if request.args.get("expired") else None
        return render_template("login.html", error=msg)

    @app.post("/login")
    def login_post():
        email = (request.form.get("email") or "").strip().lower()
        password = request.form.get("password") or ""
        if email != ALLOWED_EMAIL or not check_password_hash(PASSWORD_HASH, password):
            return render_template("login.html", error="Ungültige Zugangsdaten.")
        session["user_email"] = email
        session["last_seen"] = int(time())
        session.permanent = True
        target = request.args.get("next") or url_for("transaktion_page")
        return redirect(target)

    @app.get("/logout")
    def logout():
        session.clear()
        return redirect(url_for("login"))

    @app.get("/transaktion")
    @login_required
    def transaktion_page():
        return render_template("transaktion.html")

    @app.get("/mention")
    @login_required
    def mention_page():
        return render_template("mention.html")

    @app.get("/telematik")
    @login_required
    def telematik_page():
        return render_template("telematik.html")

    @app.get("/rentenbefreiung")
    @login_required
    def rentenbefreiung_page():
        return render_template("rentenbefreiung.html")

    @app.post("/rentenbefreiung/process")
    @login_required
    def rentenbefreiung_process():
        # Lokale Datumsvalidierung, um Importprobleme auf Linux zu vermeiden
        def _parse_date_ddmmyyyy(value: str, allow_today_default: bool = False) -> str:
            val = (value or "").strip()
            if not val and allow_today_default:
                return datetime.today().strftime("%d.%m.%Y")
            try:
                dt = datetime.strptime(val, "%d.%m.%Y")
                return dt.strftime("%d.%m.%Y")
            except ValueError:
                abort(400, "Bitte Datum im Format TT.MM.JJJJ angeben (z. B. 25.11.2025).")

        familienname = (request.form.get("familienname") or "").strip()
        vorname = (request.form.get("vorname") or "").strip()
        rvnr = (request.form.get("rvnr") or "").strip()
        ort = (request.form.get("ort") or "").strip() or "Solingen"
        aktuelles_datum_raw = (request.form.get("aktuelles_datum") or "").strip()
        beginn_befreiung_raw = (request.form.get("beginn_befreiung") or "").strip()

        excel_file = request.files.get("excel")
        if not excel_file or not (excel_file.filename or "").lower().endswith((".xlsx", ".xlsm", ".xls")):
            abort(400, "Bitte eine Excel-Vorlage (.xlsx/.xlsm/.xls) hochladen.")
        excel_bytes = excel_file.read()

        signature_file = request.files.get("signature")
        signature_bytes = signature_file.read() if signature_file and signature_file.filename else None

        # Validierung/Normalisierung der Datumsangaben
        aktuelles_datum = _parse_date_ddmmyyyy(aktuelles_datum_raw, allow_today_default=True)
        beginn_befreiung = _parse_date_ddmmyyyy(beginn_befreiung_raw, allow_today_default=False)

        try:
            # Lazy-Import (keine OS-spezifische Abhängigkeit mehr)
            from app.processors.rentenbefreiung import export_rentenbefreiung_xlsx  # type: ignore

            xlsx_bytes, filename = export_rentenbefreiung_xlsx(
                excel_bytes=excel_bytes,
                familienname=familienname,
                vorname=vorname,
                rvnr=rvnr,
                ort=ort,
                aktuelles_datum=aktuelles_datum,
                beginn_befreiung=beginn_befreiung,
                signature_bytes=signature_bytes,
            )
        except ValueError as e:
            abort(400, str(e))

        bio = BytesIO(xlsx_bytes)
        bio.seek(0)
        return send_file(
            bio,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename
        )

    # RoutenCalc entfernt

    @app.post("/telematik/process")
    @login_required
    def telematik_process():
        if "file" not in request.files:
            abort(400, "Keine Datei hochgeladen.")
        f = request.files["file"]
        if not f or not (f.filename or "").lower().endswith(".xlsx"):
            abort(400, "Bitte eine .xlsx-Datei hochladen.")
        xlsx_bytes = f.read()
        out_xlsx, xlsx_name, clip_bytes, clip_name = process_apetito_excel(xlsx_bytes)

        # Serverseitig zusätzlich nach O:\apetito\apetitotelematik\YYYYMMDD.xlsx speichern
        try:
            target_dir = r"O:\apetito\apetitotelematik"
            os.makedirs(target_dir, exist_ok=True)
            server_target = os.path.join(target_dir, xlsx_name)
            with open(server_target, "wb") as fh:
                fh.write(out_xlsx)
        except Exception:
            pass

        # Primär das Excel als Download liefern
        return _send_result(out_xlsx, xlsx_name, encoding_hint="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    def _send_result(result_bytes: bytes, filename: str, encoding_hint: str | None = None):
        if not result_bytes:
            abort(400, "Keine Ausgabedatei erzeugt.")
        bio = BytesIO(result_bytes)
        bio.seek(0)
        as_attachment_filename = filename or f"export-{datetime.now().strftime('%Y%m%d-%H%M%S')}.csv"
        # Force CSV download
        return send_file(
            bio,
            mimetype="text/csv" if not encoding_hint else f"text/csv; charset={encoding_hint}",
            as_attachment=True,
            download_name=as_attachment_filename
        )

    @app.post("/process/amazon")
    @login_required
    def process_amazon():
        if "file" not in request.files:
            abort(400, "Keine Datei hochgeladen.")
        f = request.files["file"]
        if not (f and (f.filename or "").lower().__contains__("amazon")):
            abort(400, "Ungültige Datei: Dateiname muss 'amazon' enthalten.")
        file_bytes = f.read()
        result_bytes, suggested_filename = process_amazon_csv(file_bytes, f.filename or "")
        _maybe_save_to_network(result_bytes, suggested_filename, f.filename)
        return _send_result(result_bytes, suggested_filename, encoding_hint="utf-8")

    @app.post("/process/ebay")
    @login_required
    def process_ebay():
        if "file" not in request.files:
            abort(400, "Keine Datei hochgeladen.")
        f = request.files["file"]
        if not (f and (f.filename or "").lower().__contains__("ebay")):
            abort(400, "Ungültige Datei: Dateiname muss 'ebay' enthalten.")
        file_bytes = f.read()
        result_bytes, suggested_filename = process_ebay_csv(file_bytes, f.filename or "")
        _maybe_save_to_network(result_bytes, suggested_filename, f.filename)
        return _send_result(result_bytes, suggested_filename, encoding_hint="utf-8")

    @app.post("/process/kaufland")
    @login_required
    def process_kaufland():
        if "file" not in request.files:
            abort(400, "Keine Datei hochgeladen.")
        f = request.files["file"]
        if not (f and (f.filename or "").lower().__contains__("kaufland")):
            abort(400, "Ungültige Datei: Dateiname muss 'kaufland' enthalten.")
        file_bytes = f.read()
        result_bytes, suggested_filename = process_kaufland_csv(file_bytes, f.filename or "")
        _maybe_save_to_network(result_bytes, suggested_filename, f.filename)
        return _send_result(result_bytes, suggested_filename, encoding_hint="utf-8")

    @app.post("/process/mention-ausgang")
    @login_required
    def process_mention_ausgang():
        if "file" not in request.files:
            abort(400, "Keine Datei hochgeladen.")
        f = request.files["file"]
        if not (f and (f.filename or "").lower().__contains__("mention")):
            abort(400, "Ungültige Datei: Dateiname muss 'mention' enthalten.")
        file_bytes = f.read()
        result_bytes, suggested_filename = process_mention_ausgang_excel(file_bytes, f.filename or "")
        _maybe_save_to_network(result_bytes, suggested_filename, f.filename)
        return _send_result(result_bytes, suggested_filename, encoding_hint="utf-8")

    @app.post("/process/mention-eingang")
    @login_required
    def process_mention_eingang():
        if "file" not in request.files:
            abort(400, "Keine Datei hochgeladen.")
        f = request.files["file"]
        if not (f and (f.filename or "").lower().__contains__("mention")):
            abort(400, "Ungültige Datei: Dateiname muss 'mention' enthalten.")
        file_bytes = f.read()
        result_bytes, suggested_filename = process_mention_eingang_excel(file_bytes, f.filename or "")
        _maybe_save_to_network(result_bytes, suggested_filename, f.filename)
        return _send_result(result_bytes, suggested_filename, encoding_hint="utf-8")

    @app.post("/process/sale-ausgang")
    @login_required
    def process_sale_ausgang():
        if "file" not in request.files:
            abort(400, "Keine Datei hochgeladen.")
        f = request.files["file"]
        if not (f and (f.filename or "").lower().__contains__("sale")):
            abort(400, "Ungültige Datei: Dateiname muss 'sale' enthalten.")
        file_bytes = f.read()
        result_bytes, suggested_filename = process_sale_ausgang_csv(file_bytes, f.filename or "")
        # cp1252 für Sale-Ausgang wie bisher
        _maybe_save_to_network(result_bytes, suggested_filename, f.filename)
        return _send_result(result_bytes, suggested_filename, encoding_hint="windows-1252")

    # Mention: Shopdateien erzeugen (Excel + Kosatec Bestand) -> ZIP
    @app.post("/mention/shop-files")
    @login_required
    def mention_shop_files():
        excel = request.files.get("excel")
        bestand = request.files.get("bestand")
        if not excel or not bestand:
            abort(400, "Bitte Excel- und Bestand-Datei hochladen.")
        res_bytes, fname = process_shop_files(excel.read(), excel.filename or "mention.xlsx", bestand.read(), bestand.filename or "bestand.txt")
        return _send_result(res_bytes, fname, encoding_hint="utf-8")

    # Mention: Bilder zu FTP hochladen und shopimage CSV zurückgeben
    @app.post("/mention/upload-images")
    @login_required
    def mention_upload_images():
        artikelnummer = request.form.get("artikelnummer", "").strip()
        if not artikelnummer:
            abort(400, "Bitte eine Artikelnummer angeben.")
        files = request.files.getlist("images")
        if not files:
            abort(400, "Bitte mindestens ein Bild hochladen.")
        content = [(f.filename, f.read()) for f in files if f and f.filename]
        try:
            res_bytes, fname = upload_images_and_generate_csv(artikelnummer, content)
        except RuntimeError as e:
            abort(400, str(e))
        return _send_result(res_bytes, fname, encoding_hint="utf-8")

    return app


if __name__ == "__main__":
    # Dev-Server starten (für Produktion siehe README: waitress)
    flask_app = create_app()
    flask_app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5003)), debug=True)


