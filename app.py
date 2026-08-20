import os
import io
import csv
import math
import calendar
import smtplib
import time
import uuid
import json
from email.message import EmailMessage
from functools import wraps
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from urllib.parse import quote

import openpyxl

WIB = ZoneInfo("Asia/Jakarta")


def now_wib():
    return datetime.now(WIB).replace(tzinfo=None)


def today_wib():
    return now_wib().date()

from io import BytesIO

from flask import Flask, render_template, redirect, url_for, request, flash, jsonify, Response
from flask_login import (
    LoginManager,
    login_user,
    logout_user,
    current_user,
)
from werkzeug.utils import secure_filename
from xhtml2pdf import pisa

from models import (
    db, User, Employee, Attendance, Settings, Payroll, PengajuanIzin,
    LaporanPekerjaan, PengajuanLembur, IklanMarketplace, ProdukIklan,
)

BULAN_NAMA = [
    "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]

STATUS_PEGAWAI_BULANAN = ("Karyawan Tetap", "Probation")

JABATAN_LIST = [
    "Admin Ops",
    "Admin Sales & Marketplace",
    "Admin Affiliator",
    "Admin Gudang",
    "Host Live",
    "Manager",
    "Human Resource",
    "Supervisor",
]

MARKETPLACE_LIST = ["Shopee", "Tokopedia", "TikTok Shop", "Lazada", "Blibli"]

KOLOM_TARGET_IKLAN = [
    ("tanggal", "Tanggal", True, ["tanggal", "date", "tgl", "periode"]),
    ("biaya", "Biaya Iklan", True, ["biaya", "cost", "spend", "pengeluaran", "biaya iklan"]),
    ("impresi", "Impresi/Dilihat", False, ["impresi", "impression", "dilihat", "views", "tayangan", "reach"]),
    ("klik", "Klik", False, ["klik", "click", "jumlah klik"]),
    ("pesanan", "Pesanan/Konversi", False, ["pesanan", "konversi", "conversion", "order", "dibeli", "produk terjual", "checkout", "terjual"]),
    ("omzet", "Omzet Penjualan", False, ["omzet", "omset", "penjualan", "revenue", "gmv", "nilai penjualan", "sales"]),
]


KOLOM_TARGET_PRODUK = [
    ("nama_produk", "Nama Produk", True, ["nama produk", "produk", "product", "nama barang", "item name", "item"]),
    ("tanggal", "Tanggal", True, ["tanggal", "date", "tgl", "periode"]),
    ("biaya", "Biaya Iklan", True, ["biaya", "cost", "spend", "pengeluaran", "biaya iklan"]),
    ("impresi", "Impresi/Dilihat", False, ["impresi", "impression", "dilihat", "views", "tayangan", "reach"]),
    ("klik", "Klik", False, ["klik", "click", "jumlah klik"]),
    ("pesanan", "Pesanan/Konversi", False, ["pesanan", "konversi", "conversion", "order", "dibeli", "produk terjual", "checkout", "terjual"]),
    ("omzet", "Omzet Penjualan", False, ["omzet", "omset", "penjualan", "revenue", "gmv", "nilai penjualan", "sales"]),
]


def tebak_kolom(headers, target_list=None):
    target_list = target_list or KOLOM_TARGET_IKLAN
    tebakan = {}
    headers_lower = [str(h).strip().lower() for h in headers]
    for key, _label, _wajib, kata_kunci in target_list:
        pilihan = ""
        for idx, h in enumerate(headers_lower):
            if any(kk in h for kk in kata_kunci):
                pilihan = str(idx)
                break
        tebakan[key] = pilihan
    return tebakan


def baca_file_iklan(file_storage):
    filename = file_storage.filename or ""
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    raw = file_storage.read()

    if ext == "csv":
        text = raw.decode("utf-8-sig", errors="ignore")
        sample = text[:2048]
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel
        reader = csv.reader(io.StringIO(text), dialect)
        semua = [row for row in reader if any(str(cell).strip() for cell in row)]
    elif ext in ("xlsx", "xlsm"):
        try:
            wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
        except Exception:
            return None, None, "File Excel tidak bisa dibaca. Pastikan file tidak rusak."
        ws = wb.active
        semua = []
        for row in ws.iter_rows(values_only=True):
            if any(cell not in (None, "") for cell in row):
                semua.append(["" if c is None else c for c in row])
    else:
        return None, None, "Format file tidak didukung. Gunakan file CSV atau XLSX (export laporan iklan dari marketplace)."

    if len(semua) < 2:
        return None, None, "File kosong atau tidak ada baris data."

    headers = [str(h).strip() for h in semua[0]]
    rows = semua[1:]
    return headers, rows, None


def parse_angka_iklan(value):
    if value is None:
        return 0
    if isinstance(value, (int, float)):
        return value
    s = str(value).strip()
    if not s:
        return 0
    s = s.replace("Rp", "").replace("rp", "").replace("%", "").strip()
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        kanan = s.split(",")[-1]
        s = s.replace(",", ".") if len(kanan) <= 2 else s.replace(",", "")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    elif "." in s:
        kepala, ekor = s.split(".")
        if len(ekor) == 3 and len(kepala) <= 3:
            s = s.replace(".", "")
    cleaned = "".join(ch for ch in s if ch.isdigit() or ch in "-.")
    try:
        return float(cleaned) if cleaned not in ("", "-", ".") else 0
    except ValueError:
        return 0


def parse_tanggal_iklan(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    if not s:
        return None
    if "T" in s and len(s) >= 10:
        s = s.split("T")[0]

    parts = s.split()
    if len(parts) == 3:
        try:
            hari = int(parts[0])
            tahun = int(parts[2])
            bulan_txt = parts[1].strip(",").lower()
            for idx, nama in enumerate(BULAN_NAMA):
                if idx == 0:
                    continue
                if bulan_txt == nama.lower() or bulan_txt == nama.lower()[:3]:
                    return date(tahun, idx, hari)
        except (ValueError, IndexError):
            pass

    formats = [
        "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y",
        "%m/%d/%Y", "%Y/%m/%d",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


_SATUAN = [
    "", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan",
    "Sepuluh", "Sebelas",
]


def _terbilang(n):
    n = int(n)
    if n < 12:
        return _SATUAN[n]
    if n < 20:
        return _terbilang(n - 10) + " Belas"
    if n < 100:
        return _terbilang(n // 10) + " Puluh " + _terbilang(n % 10)
    if n < 200:
        return "Seratus " + _terbilang(n - 100)
    if n < 1000:
        return _terbilang(n // 100) + " Ratus " + _terbilang(n % 100)
    if n < 2000:
        return "Seribu " + _terbilang(n - 1000)
    if n < 1000000:
        return _terbilang(n // 1000) + " Ribu " + _terbilang(n % 1000)
    if n < 1000000000:
        return _terbilang(n // 1000000) + " Juta " + _terbilang(n % 1000000)
    return _terbilang(n // 1000000000) + " Miliar " + _terbilang(n % 1000000000)


def terbilang_rupiah(n):
    n = int(n or 0)
    if n == 0:
        return "Nol Rupiah"
    words = _terbilang(n) + " Rupiah"
    return " ".join(words.split())


def haversine_meter(lat1, lng1, lat2, lng2):
    r = 6371000
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lng2 - lng1)
    a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def create_app():
    app = Flask(__name__)
    app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "ganti-secret-key-ini")

    db_url = os.environ.get("DATABASE_URL", "sqlite:///absensi.db")
    if db_url.startswith("postgres://"):
        db_url = db_url.replace("postgres://", "postgresql://", 1)
    app.config["SQLALCHEMY_DATABASE_URI"] = db_url
    app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

    upload_folder = os.path.join(app.static_folder, "uploads")
    os.makedirs(upload_folder, exist_ok=True)
    app.config["UPLOAD_FOLDER"] = upload_folder

    izin_upload_folder = os.path.join(upload_folder, "izin")
    os.makedirs(izin_upload_folder, exist_ok=True)
    app.config["IZIN_UPLOAD_FOLDER"] = izin_upload_folder

    EKSTENSI_DOKUMEN_IZIN = {"pdf", "doc", "docx", "jpg", "jpeg", "png"}

    tmp_iklan_folder = os.path.join(app.instance_path, "tmp_iklan")
    os.makedirs(tmp_iklan_folder, exist_ok=True)
    app.config["TMP_IKLAN_FOLDER"] = tmp_iklan_folder

    db.init_app(app)

    login_manager = LoginManager()
    login_manager.login_view = "login"
    login_manager.login_message = "Silakan login terlebih dahulu."
    login_manager.init_app(app)

    @login_manager.user_loader
    def load_user(user_id):
        kind, _, raw_id = user_id.partition(":")
        if kind == "admin":
            return db.session.get(User, int(raw_id))
        if kind == "emp":
            return db.session.get(Employee, int(raw_id))
        return None

    def admin_required(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            if not current_user.is_authenticated:
                return redirect(url_for("login"))
            if getattr(current_user, "role", None) != "admin":
                return redirect(url_for("pegawai_dashboard"))
            return f(*args, **kwargs)
        return wrapper

    def pegawai_required(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            if not current_user.is_authenticated:
                return redirect(url_for("pegawai_login"))
            if getattr(current_user, "role", None) != "pegawai":
                return redirect(url_for("dashboard"))
            return f(*args, **kwargs)
        return wrapper

    def marketing_required(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            if not current_user.is_authenticated:
                return redirect(url_for("login"))
            role = getattr(current_user, "role", None)
            if role == "admin":
                return f(*args, **kwargs)
            if role == "pegawai" and current_user.akses_marketing:
                return f(*args, **kwargs)
            if role == "pegawai":
                return redirect(url_for("pegawai_dashboard"))
            return redirect(url_for("login"))
        return wrapper

    def rupiah(value):
        try:
            value = int(value or 0)
        except (ValueError, TypeError):
            value = 0
        return f"Rp {value:,.0f}".replace(",", ".")

    app.jinja_env.filters["rupiah"] = rupiah
    app.jinja_env.filters["terbilang"] = terbilang_rupiah

    def buat_pesan_wa_slip(p, settings):
        label_total = "GAJI BRUTO" if p.tipe_pegawai == "Freelance" else "PENERIMAAN BERSIH"
        baris = [
            f"Halo {p.employee.nama},",
            "",
            f"Berikut slip gaji Anda periode {BULAN_NAMA[p.bulan]} {p.tahun} dari {settings.nama_perusahaan} "
            "(rincian lengkap ada di file PDF terlampir).",
            "",
            f"{label_total}: {rupiah(p.gaji_bersih)}",
            f"({terbilang_rupiah(p.gaji_bersih)})",
            "",
            "Status: SUDAH DIBAYAR",
            "",
            "Terima kasih atas kerja keras Anda selama ini. Mohon dicek kembali, jika ada pertanyaan silakan hubungi admin.",
        ]
        return "\n".join(baris)

    def normalisasi_no_hp_wa(no_hp):
        digit = "".join(ch for ch in (no_hp or "") if ch.isdigit())
        if not digit:
            return ""
        if digit.startswith("0"):
            digit = "62" + digit[1:]
        elif not digit.startswith("62"):
            digit = "62" + digit
        return digit

    def buat_wa_link(p, settings):
        if p.status != "Dibayar":
            return None
        nomor = normalisasi_no_hp_wa(p.employee.no_hp)
        if not nomor:
            return None
        pesan = buat_pesan_wa_slip(p, settings)
        return f"https://wa.me/{nomor}?text={quote(pesan)}"

    def format_tanggal_panjang(d):
        return f"{d.day:02d} {BULAN_NAMA[d.month]} {d.year}"

    app.jinja_env.filters["tanggal_panjang"] = format_tanggal_panjang

    def get_settings():
        settings = Settings.query.first()
        if not settings:
            settings = Settings()
            db.session.add(settings)
            db.session.commit()
        return settings

    def parse_hhmm(value):
        try:
            parts = value.split(":")
            return int(parts[0]) * 60 + int(parts[1])
        except (ValueError, AttributeError, IndexError, TypeError):
            return None

    def standar_jam_pegawai(settings, tipe_pegawai):
        if tipe_pegawai == "Freelance":
            return settings.jam_masuk_standar_freelance, settings.jam_pulang_standar_freelance
        return settings.jam_masuk_standar, settings.jam_pulang_standar

    def hitung_telat_lembur(jam_masuk, jam_pulang, settings, tipe_pegawai="Karyawan Tetap"):
        telat = 0
        lembur = 0
        jam_masuk_standar, jam_pulang_standar = standar_jam_pegawai(settings, tipe_pegawai)
        standar_masuk = parse_hhmm(jam_masuk_standar)
        standar_pulang = parse_hhmm(jam_pulang_standar)
        actual_masuk = parse_hhmm(jam_masuk)
        actual_pulang = parse_hhmm(jam_pulang)

        if standar_masuk is not None and actual_masuk is not None:
            selisih = actual_masuk - standar_masuk - (settings.toleransi_telat_menit or 0)
            if selisih > 0:
                telat = selisih

        if standar_pulang is not None and actual_pulang is not None:
            selisih = actual_pulang - standar_pulang
            if selisih >= (settings.toleransi_lembur_menit or 0):
                lembur = selisih

        return telat, lembur

    def cek_pulang_cepat(jam_pulang, settings, tipe_pegawai="Karyawan Tetap"):
        _, jam_pulang_standar = standar_jam_pegawai(settings, tipe_pegawai)
        standar_pulang = parse_hhmm(jam_pulang_standar)
        actual_pulang = parse_hhmm(jam_pulang)
        if standar_pulang is None or actual_pulang is None:
            return 0
        selisih = standar_pulang - actual_pulang
        return selisih if selisih > 0 else 0

    with app.app_context():
        db.create_all()
        kolom_employee = {c["name"] for c in db.inspect(db.engine).get_columns("employee")}
        if "akses_marketing" not in kolom_employee:
            db.session.execute(db.text(
                "ALTER TABLE employee ADD COLUMN akses_marketing BOOLEAN NOT NULL DEFAULT 0"
            ))
            db.session.commit()
        if not User.query.first():
            admin = User(username="admin", nama="Administrator")
            admin.set_password("admin123")
            db.session.add(admin)
            db.session.commit()
        get_settings()

    # ---------- AUTH ADMIN ----------
    @app.route("/login", methods=["GET", "POST"])
    def login():
        if current_user.is_authenticated:
            if getattr(current_user, "role", None) == "admin":
                return redirect(url_for("dashboard"))
            logout_user()
        if request.method == "POST":
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "")
            user = User.query.filter_by(username=username).first()
            if user and user.check_password(password):
                login_user(user)
                return redirect(url_for("dashboard"))
            flash("Username atau password salah.", "danger")
        return render_template("login.html")

    @app.route("/logout")
    def logout():
        logout_user()
        return redirect(url_for("login"))

    # ---------- AUTH PEGAWAI ----------
    @app.route("/pegawai/login", methods=["GET", "POST"])
    def pegawai_login():
        if current_user.is_authenticated:
            if getattr(current_user, "role", None) == "pegawai":
                return redirect(url_for("pegawai_dashboard"))
            logout_user()
        if request.method == "POST":
            no_hp = request.form.get("no_hp", "").strip()
            password = request.form.get("password", "")
            emp = Employee.query.filter_by(no_hp=no_hp, status="Aktif").first()
            if emp and emp.check_password(password):
                login_user(emp)
                return redirect(url_for("pegawai_dashboard"))
            flash("No. HP atau password salah.", "danger")
        return render_template("pegawai/login.html")

    @app.route("/pegawai/logout")
    def pegawai_logout():
        logout_user()
        return redirect(url_for("pegawai_login"))

    # ---------- DASHBOARD ----------
    @app.route("/")
    @admin_required
    def dashboard():
        total_karyawan = Employee.query.filter_by(status="Aktif").count()
        hari_ini = today_wib()
        absen_hari_ini = Attendance.query.filter_by(tanggal=hari_ini).count()
        hadir_hari_ini = Attendance.query.filter_by(tanggal=hari_ini, status="Hadir").count()

        bulan_ini, tahun_ini = hari_ini.month, hari_ini.year
        payrolls_bulan_ini = Payroll.query.filter_by(bulan=bulan_ini, tahun=tahun_ini).all()
        total_gaji_bulan_ini = sum(p.gaji_bersih for p in payrolls_bulan_ini)

        return render_template(
            "dashboard.html",
            total_karyawan=total_karyawan,
            absen_hari_ini=absen_hari_ini,
            hadir_hari_ini=hadir_hari_ini,
            total_gaji_bulan_ini=total_gaji_bulan_ini,
            bulan_nama=BULAN_NAMA[bulan_ini],
            tahun_ini=tahun_ini,
        )

    # ---------- KARYAWAN ----------
    @app.route("/karyawan")
    @admin_required
    def karyawan_list():
        karyawan = Employee.query.order_by(Employee.nama).all()
        return render_template("employees_list.html", karyawan=karyawan)

    @app.route("/karyawan/tambah", methods=["GET", "POST"])
    @admin_required
    def karyawan_tambah():
        if request.method == "POST":
            no_hp = request.form.get("no_hp", "").strip()
            if no_hp and Employee.query.filter_by(no_hp=no_hp).first():
                flash("No. HP sudah dipakai karyawan lain. Gunakan nomor lain agar login tidak tertukar.", "danger")
                return render_template("employee_form.html", karyawan=None, jabatan_list=JABATAN_LIST)

            emp = Employee(
                nama=request.form["nama"].strip(),
                jabatan=request.form.get("jabatan", "").strip(),
                email=request.form.get("email", "").strip(),
                no_hp=no_hp,
                alamat=request.form.get("alamat", "").strip(),
                nomor_rekening=request.form.get("nomor_rekening", "").strip(),
                gaji_pokok=int(request.form.get("gaji_pokok") or 0),
                tunjangan_makan=int(request.form.get("tunjangan_makan") or 0),
                tunjangan_transport=int(request.form.get("tunjangan_transport") or 0),
                tipe_pegawai=request.form.get("tipe_pegawai", "Karyawan Tetap"),
                target_tercapai=request.form.get("target_tercapai", "Tidak Tercapai"),
                bpjs_jkk=int(request.form.get("bpjs_jkk") or 0),
                bpjs_jkm=int(request.form.get("bpjs_jkm") or 0),
                bpjs_jht=int(request.form.get("bpjs_jht") or 0),
                bpjs_kesehatan=int(request.form.get("bpjs_kesehatan") or 0),
                tarif_unit_freelance=int(request.form.get("tarif_unit_freelance") or 0),
                co_host_bulan_ini=request.form.get("co_host_bulan_ini", "Tidak"),
                status=request.form.get("status", "Aktif"),
                akses_marketing=request.form.get("akses_marketing") == "on",
            )
            password_baru = request.form.get("password_baru", "")
            if password_baru:
                emp.set_password(password_baru)
            db.session.add(emp)
            db.session.commit()
            flash(f"Karyawan {emp.nama} berhasil ditambahkan.", "success")
            return redirect(url_for("karyawan_list"))
        return render_template("employee_form.html", karyawan=None, jabatan_list=JABATAN_LIST)

    @app.route("/karyawan/<int:emp_id>/edit", methods=["GET", "POST"])
    @admin_required
    def karyawan_edit(emp_id):
        emp = db.session.get(Employee, emp_id) or abort_404()
        hari_ini = today_wib()
        hari_hadir_bulan_ini = Attendance.query.filter(
            Attendance.employee_id == emp.id,
            Attendance.status == "Hadir",
            db.extract("month", Attendance.tanggal) == hari_ini.month,
            db.extract("year", Attendance.tanggal) == hari_ini.year,
        ).count()
        if request.method == "POST":
            no_hp = request.form.get("no_hp", "").strip()
            if no_hp and Employee.query.filter(Employee.no_hp == no_hp, Employee.id != emp.id).first():
                flash("No. HP sudah dipakai karyawan lain. Gunakan nomor lain agar login tidak tertukar.", "danger")
                return render_template(
                    "employee_form.html", karyawan=emp, jabatan_list=JABATAN_LIST,
                    hari_hadir_bulan_ini=hari_hadir_bulan_ini,
                )

            emp.nama = request.form["nama"].strip()
            emp.jabatan = request.form.get("jabatan", "").strip()
            emp.email = request.form.get("email", "").strip()
            emp.no_hp = no_hp
            emp.alamat = request.form.get("alamat", "").strip()
            emp.nomor_rekening = request.form.get("nomor_rekening", "").strip()
            emp.gaji_pokok = int(request.form.get("gaji_pokok") or 0)
            emp.tunjangan_makan = int(request.form.get("tunjangan_makan") or 0)
            emp.tunjangan_transport = int(request.form.get("tunjangan_transport") or 0)
            emp.tipe_pegawai = request.form.get("tipe_pegawai", "Karyawan Tetap")
            emp.target_tercapai = request.form.get("target_tercapai", "Tidak Tercapai")
            emp.bpjs_jkk = int(request.form.get("bpjs_jkk") or 0)
            emp.bpjs_jkm = int(request.form.get("bpjs_jkm") or 0)
            emp.bpjs_jht = int(request.form.get("bpjs_jht") or 0)
            emp.bpjs_kesehatan = int(request.form.get("bpjs_kesehatan") or 0)
            emp.tarif_unit_freelance = int(request.form.get("tarif_unit_freelance") or 0)
            emp.co_host_bulan_ini = request.form.get("co_host_bulan_ini", "Tidak")
            emp.status = request.form.get("status", "Aktif")
            emp.akses_marketing = request.form.get("akses_marketing") == "on"
            password_baru = request.form.get("password_baru", "")
            if password_baru:
                emp.set_password(password_baru)
            db.session.commit()
            flash(f"Data {emp.nama} berhasil diperbarui.", "success")
            return redirect(url_for("karyawan_list"))
        return render_template(
            "employee_form.html", karyawan=emp, jabatan_list=JABATAN_LIST,
            hari_hadir_bulan_ini=hari_hadir_bulan_ini,
        )

    @app.route("/karyawan/<int:emp_id>/hapus", methods=["POST"])
    @admin_required
    def karyawan_hapus(emp_id):
        emp = db.session.get(Employee, emp_id) or abort_404()
        nama = emp.nama
        db.session.delete(emp)
        db.session.commit()
        flash(f"Karyawan {nama} beserta riwayatnya dihapus.", "info")
        return redirect(url_for("karyawan_list"))

    # ---------- ABSENSI ----------
    @app.route("/absensi")
    @admin_required
    def absensi_list():
        tanggal_str = request.args.get("tanggal", today_wib().isoformat())
        try:
            tanggal = datetime.strptime(tanggal_str, "%Y-%m-%d").date()
        except ValueError:
            tanggal = today_wib()

        data = (
            Attendance.query.filter_by(tanggal=tanggal)
            .join(Employee)
            .order_by(Employee.nama)
            .all()
        )
        sudah_absen_ids = {a.employee_id for a in data}
        belum_absen = Employee.query.filter(
            Employee.status == "Aktif", ~Employee.id.in_(sudah_absen_ids) if sudah_absen_ids else True
        ).order_by(Employee.nama).all()

        return render_template(
            "attendance_list.html",
            data=data,
            tanggal=tanggal,
            belum_absen=belum_absen,
        )

    @app.route("/absensi/koreksi", methods=["GET", "POST"], defaults={"att_id": None})
    @app.route("/absensi/<int:att_id>/koreksi", methods=["GET", "POST"])
    @admin_required
    def absensi_koreksi(att_id):
        settings = get_settings()
        att = db.session.get(Attendance, att_id) if att_id else None

        if request.method == "POST":
            employee_id = int(request.form["employee_id"])
            tanggal = datetime.strptime(request.form["tanggal"], "%Y-%m-%d").date()
            status = request.form["status"]
            jam_masuk = request.form.get("jam_masuk") or None
            jam_pulang = request.form.get("jam_pulang") or None
            catatan = request.form.get("catatan", "").strip()

            telat, lembur = (0, 0)
            if status == "Hadir":
                emp_koreksi = db.session.get(Employee, employee_id)
                tipe_koreksi = emp_koreksi.tipe_pegawai if emp_koreksi else "Karyawan Tetap"
                telat, lembur = hitung_telat_lembur(jam_masuk, jam_pulang, settings, tipe_koreksi)
            else:
                jam_masuk, jam_pulang = None, None

            target = att or Attendance.query.filter_by(
                employee_id=employee_id, tanggal=tanggal
            ).first()
            if not target:
                target = Attendance(employee_id=employee_id, tanggal=tanggal)
                db.session.add(target)

            target.employee_id = employee_id
            target.tanggal = tanggal
            target.status = status
            target.jam_masuk = jam_masuk
            target.jam_pulang = jam_pulang
            target.telat_menit = telat
            target.lembur_menit = lembur
            target.catatan = catatan
            db.session.commit()
            flash("Data absensi berhasil disimpan.", "success")
            return redirect(url_for("absensi_list", tanggal=tanggal.isoformat()))

        employees = Employee.query.filter_by(status="Aktif").order_by(Employee.nama).all()
        tanggal_default = request.args.get("tanggal", today_wib().isoformat())
        return render_template(
            "attendance_koreksi.html",
            att=att,
            employees=employees,
            tanggal_default=tanggal_default,
        )

    @app.route("/absensi/<int:att_id>/hapus", methods=["POST"])
    @admin_required
    def absensi_hapus(att_id):
        att = db.session.get(Attendance, att_id) or abort_404()
        tanggal = att.tanggal
        db.session.delete(att)
        db.session.commit()
        flash("Data absensi dihapus.", "info")
        return redirect(url_for("absensi_list", tanggal=tanggal.isoformat()))

    # ---------- PENGAJUAN IZIN (ADMIN) ----------
    @app.route("/pengajuan-izin")
    @admin_required
    def pengajuan_izin_list():
        status_filter = request.args.get("status", "Menunggu")
        q = PengajuanIzin.query.join(Employee)
        if status_filter in ("Menunggu", "Disetujui", "Ditolak"):
            q = q.filter(PengajuanIzin.status == status_filter)
        pengajuan = q.order_by(PengajuanIzin.tanggal_diajukan.desc()).all()
        jumlah_menunggu = PengajuanIzin.query.filter_by(status="Menunggu").count()
        return render_template(
            "pengajuan_izin_list.html",
            pengajuan=pengajuan,
            status_filter=status_filter,
            jumlah_menunggu=jumlah_menunggu,
        )

    @app.route("/pengajuan-izin/<int:pid>/setujui", methods=["POST"])
    @admin_required
    def pengajuan_izin_setujui(pid):
        p = db.session.get(PengajuanIzin, pid) or abort_404()
        p.status = "Disetujui"
        p.tanggal_diproses = now_wib()

        att = Attendance.query.filter_by(employee_id=p.employee_id, tanggal=p.tanggal).first()
        if not att:
            att = Attendance(employee_id=p.employee_id, tanggal=p.tanggal)
            db.session.add(att)
        att.status = p.jenis
        att.jam_masuk = None
        att.jam_pulang = None
        att.telat_menit = 0
        att.lembur_menit = 0
        att.catatan = p.alasan

        db.session.commit()
        flash(f"Pengajuan {p.employee.nama} disetujui & tercatat di absensi.", "success")
        return redirect(url_for("pengajuan_izin_list"))

    @app.route("/pengajuan-izin/<int:pid>/tolak", methods=["POST"])
    @admin_required
    def pengajuan_izin_tolak(pid):
        p = db.session.get(PengajuanIzin, pid) or abort_404()
        p.status = "Ditolak"
        p.catatan_admin = request.form.get("catatan_admin", "").strip()
        p.tanggal_diproses = now_wib()
        db.session.commit()
        flash(f"Pengajuan {p.employee.nama} ditolak.", "info")
        return redirect(url_for("pengajuan_izin_list"))

    # ---------- AREA PEGAWAI ----------
    @app.route("/pegawai")
    @pegawai_required
    def pegawai_dashboard():
        hari_ini = today_wib()
        absen = Attendance.query.filter_by(employee_id=current_user.id, tanggal=hari_ini).first()
        settings = get_settings()
        return render_template("pegawai/dashboard.html", absen=absen, settings=settings, hari_ini=hari_ini)

    @app.route("/pegawai/absen", methods=["POST"])
    @pegawai_required
    def pegawai_absen():
        data = request.get_json(silent=True) or {}
        lat = data.get("lat")
        lng = data.get("lng")
        if lat is None or lng is None:
            return jsonify(success=False, message="Lokasi tidak terdeteksi. Aktifkan GPS/izin lokasi di HP Anda."), 400

        settings = get_settings()
        if settings.kantor_lat is not None and settings.kantor_lng is not None:
            jarak = haversine_meter(float(lat), float(lng), settings.kantor_lat, settings.kantor_lng)
            radius = settings.radius_kantor_meter or 100
            if jarak > radius:
                return jsonify(
                    success=False,
                    message=f"Anda berjarak {int(jarak)} meter dari kantor (maksimal {radius} meter). Absen ditolak.",
                ), 400

        hari_ini = today_wib()
        jam_sekarang = now_wib().strftime("%H:%M:%S")
        tipe_pegawai = current_user.tipe_pegawai
        att = Attendance.query.filter_by(employee_id=current_user.id, tanggal=hari_ini).first()

        if not att or not att.jam_masuk:
            telat, _ = hitung_telat_lembur(jam_sekarang, None, settings, tipe_pegawai)
            if not att:
                att = Attendance(employee_id=current_user.id, tanggal=hari_ini)
                db.session.add(att)
            att.status = "Hadir"
            att.jam_masuk = jam_sekarang
            att.telat_menit = telat
            att.lokasi_masuk_lat = lat
            att.lokasi_masuk_lng = lng
            db.session.commit()
            pesan = f"Absen masuk berhasil pukul {jam_sekarang}."
            if telat > 0:
                pesan += f" Anda terlambat {telat} menit."
            return jsonify(success=True, aksi="masuk", message=pesan, telat=telat)

        if not att.jam_pulang:
            _, lembur = hitung_telat_lembur(att.jam_masuk, jam_sekarang, settings, tipe_pegawai)
            pulang_cepat = cek_pulang_cepat(jam_sekarang, settings, tipe_pegawai)
            att.jam_pulang = jam_sekarang
            att.lembur_menit = lembur
            att.lokasi_pulang_lat = lat
            att.lokasi_pulang_lng = lng
            db.session.commit()
            pesan = f"Absen pulang berhasil pukul {jam_sekarang}."
            if pulang_cepat > 0:
                pesan += f" ⚠️ Anda pulang {pulang_cepat} menit lebih awal dari jam standar."
            return jsonify(success=True, aksi="pulang", message=pesan, pulang_cepat=pulang_cepat)

        return jsonify(success=False, message="Anda sudah absen masuk & pulang hari ini."), 400

    @app.route("/pegawai/riwayat")
    @pegawai_required
    def pegawai_riwayat():
        riwayat = (
            Attendance.query.filter_by(employee_id=current_user.id)
            .order_by(Attendance.tanggal.desc())
            .limit(31)
            .all()
        )
        return render_template("pegawai/riwayat.html", riwayat=riwayat)

    @app.route("/pegawai/izin", methods=["GET", "POST"])
    @pegawai_required
    def pegawai_izin():
        if request.method == "POST":
            jenis = request.form.get("jenis")
            try:
                tanggal = datetime.strptime(request.form.get("tanggal", ""), "%Y-%m-%d").date()
            except ValueError:
                tanggal = None

            dokumen_file = request.files.get("dokumen")
            dokumen_ada = bool(dokumen_file and dokumen_file.filename)

            if jenis not in ("Sakit", "Izin", "Cuti") or not tanggal:
                flash("Lengkapi tanggal dan jenis pengajuan dengan benar.", "danger")
            elif jenis in ("Sakit", "Cuti") and not dokumen_ada:
                flash(
                    "Dokumen wajib dilampirkan untuk pengajuan Sakit atau Cuti "
                    "(surat izin cuti / surat keterangan dokter).",
                    "danger",
                )
            elif dokumen_ada and dokumen_file.filename.rsplit(".", 1)[-1].lower() not in EKSTENSI_DOKUMEN_IZIN:
                flash("Format dokumen harus PDF, DOC, DOCX, JPG, atau PNG.", "danger")
            else:
                sudah_ada = PengajuanIzin.query.filter_by(
                    employee_id=current_user.id, tanggal=tanggal, status="Menunggu"
                ).first()
                if sudah_ada:
                    flash("Sudah ada pengajuan untuk tanggal ini yang masih menunggu persetujuan.", "warning")
                else:
                    dokumen_filename = None
                    if dokumen_ada:
                        ext = dokumen_file.filename.rsplit(".", 1)[-1].lower()
                        dokumen_filename = secure_filename(
                            f"izin_{current_user.id}_{tanggal.isoformat()}_{now_wib().strftime('%H%M%S')}.{ext}"
                        )
                        dokumen_file.save(os.path.join(app.config["IZIN_UPLOAD_FOLDER"], dokumen_filename))

                    db.session.add(
                        PengajuanIzin(
                            employee_id=current_user.id,
                            tanggal=tanggal,
                            jenis=jenis,
                            alasan=request.form.get("alasan", "").strip(),
                            dokumen_filename=dokumen_filename,
                        )
                    )
                    db.session.commit()
                    flash("Pengajuan berhasil dikirim, menunggu persetujuan admin.", "success")
            return redirect(url_for("pegawai_izin"))

        riwayat = (
            PengajuanIzin.query.filter_by(employee_id=current_user.id)
            .order_by(PengajuanIzin.tanggal_diajukan.desc())
            .limit(20)
            .all()
        )
        return render_template("pegawai/izin.html", riwayat=riwayat)

    @app.route("/pegawai/akun", methods=["GET", "POST"])
    @pegawai_required
    def pegawai_akun():
        if request.method == "POST":
            password_baru = request.form.get("password_baru", "")
            password_konfirmasi = request.form.get("password_konfirmasi", "")
            if len(password_baru) < 6:
                flash("Password baru minimal 6 karakter.", "danger")
            elif password_baru != password_konfirmasi:
                flash("Konfirmasi password tidak cocok.", "danger")
            else:
                current_user.set_password(password_baru)
                db.session.commit()
                flash("Password berhasil diubah.", "success")
            return redirect(url_for("pegawai_akun"))
        return render_template("pegawai/akun.html")

    @app.route("/pegawai/laporan", methods=["GET", "POST"])
    @pegawai_required
    def pegawai_laporan():
        if request.method == "POST":
            try:
                tanggal = datetime.strptime(request.form.get("tanggal", ""), "%Y-%m-%d").date()
            except ValueError:
                tanggal = None
            isi_laporan = request.form.get("isi_laporan", "").strip()

            if not tanggal or not isi_laporan:
                flash("Lengkapi tanggal dan isi laporan.", "danger")
            else:
                db.session.add(
                    LaporanPekerjaan(employee_id=current_user.id, tanggal=tanggal, isi_laporan=isi_laporan)
                )
                db.session.commit()
                flash("Laporan pekerjaan berhasil dikirim.", "success")
            return redirect(url_for("pegawai_laporan"))

        riwayat = (
            LaporanPekerjaan.query.filter_by(employee_id=current_user.id)
            .order_by(LaporanPekerjaan.tanggal.desc())
            .limit(30)
            .all()
        )
        return render_template("pegawai/laporan.html", riwayat=riwayat, tanggal_default=today_wib().isoformat())

    @app.route("/pegawai/lembur", methods=["GET", "POST"])
    @pegawai_required
    def pegawai_lembur():
        if request.method == "POST":
            try:
                tanggal = datetime.strptime(request.form.get("tanggal", ""), "%Y-%m-%d").date()
            except ValueError:
                tanggal = None
            jam_mulai = request.form.get("jam_mulai") or None
            jam_selesai = request.form.get("jam_selesai") or None
            alasan = request.form.get("alasan", "").strip()

            durasi_valid = False
            if jam_mulai and jam_selesai:
                try:
                    hm, mm = jam_mulai.split(":")
                    hs, ms = jam_selesai.split(":")
                    durasi_valid = (int(hs) * 60 + int(ms)) > (int(hm) * 60 + int(mm))
                except ValueError:
                    durasi_valid = False

            if not tanggal or not jam_mulai or not jam_selesai:
                flash("Lengkapi tanggal, jam mulai, dan jam selesai.", "danger")
            elif not durasi_valid:
                flash("Jam selesai harus lebih besar dari jam mulai.", "danger")
            else:
                db.session.add(
                    PengajuanLembur(
                        employee_id=current_user.id,
                        tanggal=tanggal,
                        jam_mulai=jam_mulai,
                        jam_selesai=jam_selesai,
                        alasan=alasan,
                    )
                )
                db.session.commit()
                flash("Pengajuan lembur berhasil dikirim, menunggu persetujuan admin.", "success")
            return redirect(url_for("pegawai_lembur"))

        riwayat = (
            PengajuanLembur.query.filter_by(employee_id=current_user.id)
            .order_by(PengajuanLembur.tanggal_diajukan.desc())
            .limit(20)
            .all()
        )
        return render_template("pegawai/lembur.html", riwayat=riwayat)

    # ---------- LAPORAN PEKERJAAN (ADMIN) ----------
    @app.route("/laporan-pekerjaan")
    @admin_required
    def laporan_pekerjaan_list():
        employee_id = request.args.get("employee_id", type=int)
        query = LaporanPekerjaan.query.join(Employee)
        if employee_id:
            query = query.filter(LaporanPekerjaan.employee_id == employee_id)
        laporan = query.order_by(LaporanPekerjaan.tanggal.desc()).limit(200).all()
        employees = Employee.query.order_by(Employee.nama).all()
        return render_template(
            "laporan_pekerjaan_list.html", laporan=laporan, employees=employees, employee_id=employee_id
        )

    # ---------- PENGAJUAN LEMBUR (ADMIN) ----------
    @app.route("/pengajuan-lembur")
    @admin_required
    def pengajuan_lembur_list():
        status_filter = request.args.get("status", "Menunggu")
        query = PengajuanLembur.query.join(Employee)
        if status_filter in ("Menunggu", "Disetujui", "Ditolak"):
            query = query.filter(PengajuanLembur.status == status_filter)
        pengajuan = query.order_by(PengajuanLembur.tanggal_diajukan.desc()).all()
        jumlah_menunggu = PengajuanLembur.query.filter_by(status="Menunggu").count()
        return render_template(
            "pengajuan_lembur_list.html",
            pengajuan=pengajuan,
            status_filter=status_filter,
            jumlah_menunggu=jumlah_menunggu,
        )

    @app.route("/pengajuan-lembur/<int:pid>/setujui", methods=["POST"])
    @admin_required
    def pengajuan_lembur_setujui(pid):
        p = db.session.get(PengajuanLembur, pid) or abort_404()
        p.status = "Disetujui"
        p.tanggal_diproses = now_wib()

        hm, mm = p.jam_mulai.split(":")
        hs, ms = p.jam_selesai.split(":")
        durasi_menit = (int(hs) * 60 + int(ms)) - (int(hm) * 60 + int(mm))

        att = Attendance.query.filter_by(employee_id=p.employee_id, tanggal=p.tanggal).first()
        if not att:
            att = Attendance(employee_id=p.employee_id, tanggal=p.tanggal, status="Hadir")
            db.session.add(att)
        att.lembur_menit = (att.lembur_menit or 0) + durasi_menit

        db.session.commit()
        flash(f"Pengajuan lembur {p.employee.nama} disetujui & ditambahkan ke absensi.", "success")
        return redirect(url_for("pengajuan_lembur_list"))

    @app.route("/pengajuan-lembur/<int:pid>/tolak", methods=["POST"])
    @admin_required
    def pengajuan_lembur_tolak(pid):
        p = db.session.get(PengajuanLembur, pid) or abort_404()
        p.status = "Ditolak"
        p.catatan_admin = request.form.get("catatan_admin", "").strip()
        p.tanggal_diproses = now_wib()
        db.session.commit()
        flash(f"Pengajuan lembur {p.employee.nama} ditolak.", "info")
        return redirect(url_for("pengajuan_lembur_list"))

    # ---------- PENGGAJIAN ----------
    @app.route("/penggajian")
    @admin_required
    def penggajian_list():
        bulan = int(request.args.get("bulan", today_wib().month))
        tahun = int(request.args.get("tahun", today_wib().year))
        payrolls = (
            Payroll.query.filter_by(bulan=bulan, tahun=tahun)
            .join(Employee)
            .order_by(Employee.nama)
            .all()
        )
        settings = get_settings()
        for p in payrolls:
            p.wa_link = buat_wa_link(p, settings)
        total_gaji = sum(p.gaji_bersih for p in payrolls)
        return render_template(
            "payroll_list.html",
            payrolls=payrolls,
            bulan=bulan,
            tahun=tahun,
            bulan_nama=BULAN_NAMA,
            total_gaji=total_gaji,
        )

    @app.route("/penggajian/generate", methods=["POST"])
    @admin_required
    def penggajian_generate():
        bulan = int(request.form["bulan"])
        tahun = int(request.form["tahun"])
        settings = get_settings()

        awal = date(tahun, bulan, 1)
        akhir = date(tahun, bulan, calendar.monthrange(tahun, bulan)[1])

        employees = Employee.query.filter(
            Employee.status == "Aktif", Employee.tipe_pegawai != "Freelance"
        ).all()
        for emp in employees:
            absensi = Attendance.query.filter(
                Attendance.employee_id == emp.id,
                Attendance.tanggal >= awal,
                Attendance.tanggal <= akhir,
            ).all()

            total_hadir = sum(1 for a in absensi if a.status == "Hadir")
            total_sakit = sum(1 for a in absensi if a.status == "Sakit")
            total_izin = sum(1 for a in absensi if a.status == "Izin")
            total_cuti = sum(1 for a in absensi if a.status == "Cuti")
            total_alpha = sum(1 for a in absensi if a.status == "Alpha")
            total_telat_menit = sum(a.telat_menit or 0 for a in absensi)
            total_lembur_menit = sum(a.lembur_menit or 0 for a in absensi)

            bonus_target = (
                settings.bonus_target_tercapai or 0
            ) if emp.target_tercapai == "Tercapai" else 0

            co_host_fee = 0
            upah_freelance = 0

            total_pokok = emp.gaji_pokok + emp.tunjangan_makan + emp.tunjangan_transport
            hari_kerja = settings.hari_kerja_per_bulan or 22
            gaji_harian = total_pokok / hari_kerja if hari_kerja else 0

            potongan_alpha = round(total_alpha * gaji_harian)
            potongan_telat = total_telat_menit * (settings.denda_telat_per_menit or 0)
            uang_lembur = round((total_lembur_menit / 60) * (settings.upah_lembur_per_jam or 0))

            bpjs_jkk = emp.bpjs_jkk or 0
            bpjs_jkm = emp.bpjs_jkm or 0
            bpjs_jht = emp.bpjs_jht or 0
            bpjs_kesehatan = emp.bpjs_kesehatan or 0

            gaji_bersih = (
                total_pokok
                - potongan_alpha
                - potongan_telat
                - bpjs_jkk - bpjs_jkm - bpjs_jht - bpjs_kesehatan
                + uang_lembur
                + bonus_target
            )

            gaji_bersih = max(gaji_bersih, 0)

            payroll = Payroll.query.filter_by(
                employee_id=emp.id, bulan=bulan, tahun=tahun
            ).first()
            if not payroll:
                payroll = Payroll(employee_id=emp.id, bulan=bulan, tahun=tahun)
                db.session.add(payroll)

            if payroll.status == "Dibayar":
                continue  # jangan timpa slip yang sudah dibayar

            payroll.gaji_pokok = emp.gaji_pokok
            payroll.tunjangan_makan = emp.tunjangan_makan
            payroll.tunjangan_transport = emp.tunjangan_transport
            payroll.tipe_pegawai = emp.tipe_pegawai
            payroll.total_hadir = total_hadir
            payroll.total_sakit = total_sakit
            payroll.total_izin = total_izin
            payroll.total_cuti = total_cuti
            payroll.total_alpha = total_alpha
            payroll.total_telat_menit = total_telat_menit
            payroll.total_lembur_menit = total_lembur_menit
            payroll.potongan_alpha = potongan_alpha
            payroll.potongan_telat = potongan_telat
            payroll.uang_lembur = uang_lembur
            payroll.upah_freelance = upah_freelance
            payroll.bonus_target = bonus_target
            payroll.co_host_fee = co_host_fee
            payroll.bpjs_jkk = bpjs_jkk
            payroll.bpjs_jkm = bpjs_jkm
            payroll.bpjs_jht = bpjs_jht
            payroll.bpjs_kesehatan = bpjs_kesehatan
            payroll.gaji_bersih = gaji_bersih

        db.session.commit()
        flash(f"Slip gaji Karyawan Tetap/Probation {BULAN_NAMA[bulan]} {tahun} berhasil dibuat/diperbarui.", "success")
        return redirect(url_for("penggajian_list", bulan=bulan, tahun=tahun))

    @app.route("/penggajian/freelance-review")
    @admin_required
    def penggajian_freelance_review():
        bulan = int(request.args.get("bulan", today_wib().month))
        tahun = int(request.args.get("tahun", today_wib().year))
        settings = get_settings()

        awal = date(tahun, bulan, 1)
        akhir = date(tahun, bulan, calendar.monthrange(tahun, bulan)[1])

        freelancer_rows = []
        employees = Employee.query.filter_by(status="Aktif", tipe_pegawai="Freelance").order_by(Employee.nama).all()
        for emp in employees:
            absensi = Attendance.query.filter(
                Attendance.employee_id == emp.id,
                Attendance.tanggal >= awal,
                Attendance.tanggal <= akhir,
            ).all()
            total_hadir = sum(1 for a in absensi if a.status == "Hadir")
            existing_payroll = Payroll.query.filter_by(employee_id=emp.id, bulan=bulan, tahun=tahun).first()
            tarif_unit = emp.tarif_unit_freelance or settings.tarif_harian_freelance or 0
            freelancer_rows.append({
                "employee": emp,
                "hari_kerja_absensi": total_hadir,
                "tarif_unit": tarif_unit,
                "sudah_dibayar": bool(existing_payroll and existing_payroll.status == "Dibayar"),
            })

        return render_template(
            "payroll_freelance_review.html",
            rows=freelancer_rows,
            bulan=bulan,
            tahun=tahun,
            bulan_nama=BULAN_NAMA,
        )

    @app.route("/penggajian/freelance-generate", methods=["POST"])
    @admin_required
    def penggajian_freelance_generate():
        bulan = int(request.form["bulan"])
        tahun = int(request.form["tahun"])
        settings = get_settings()

        awal = date(tahun, bulan, 1)
        akhir = date(tahun, bulan, calendar.monthrange(tahun, bulan)[1])

        employees = Employee.query.filter_by(status="Aktif", tipe_pegawai="Freelance").all()
        for emp in employees:
            absensi = Attendance.query.filter(
                Attendance.employee_id == emp.id,
                Attendance.tanggal >= awal,
                Attendance.tanggal <= akhir,
            ).all()
            total_hadir = sum(1 for a in absensi if a.status == "Hadir")
            hari_kerja = total_hadir
            total_sakit = sum(1 for a in absensi if a.status == "Sakit")
            total_izin = sum(1 for a in absensi if a.status == "Izin")
            total_cuti = sum(1 for a in absensi if a.status == "Cuti")
            total_alpha = sum(1 for a in absensi if a.status == "Alpha")
            total_telat_menit = sum(a.telat_menit or 0 for a in absensi)
            total_lembur_menit = sum(a.lembur_menit or 0 for a in absensi)

            bonus_target = (
                settings.bonus_target_tercapai or 0
            ) if emp.target_tercapai == "Tercapai" else 0

            tarif_unit = emp.tarif_unit_freelance or settings.tarif_harian_freelance or 0
            upah_freelance = hari_kerja * tarif_unit
            uang_lembur = round(
                (total_lembur_menit / 60) * (settings.upah_lembur_freelance_per_jam or 0)
            )
            co_host_fee = (settings.tarif_co_host or 0) if emp.co_host_bulan_ini == "Ya" else 0

            gaji_bersih = max(upah_freelance + uang_lembur + bonus_target + co_host_fee, 0)

            payroll = Payroll.query.filter_by(
                employee_id=emp.id, bulan=bulan, tahun=tahun
            ).first()
            if not payroll:
                payroll = Payroll(employee_id=emp.id, bulan=bulan, tahun=tahun)
                db.session.add(payroll)

            if payroll.status == "Dibayar":
                continue  # jangan timpa slip yang sudah dibayar

            payroll.gaji_pokok = 0
            payroll.tunjangan_makan = 0
            payroll.tunjangan_transport = 0
            payroll.tipe_pegawai = emp.tipe_pegawai
            payroll.total_hadir = hari_kerja
            payroll.total_sakit = total_sakit
            payroll.total_izin = total_izin
            payroll.total_cuti = total_cuti
            payroll.total_alpha = total_alpha
            payroll.total_telat_menit = total_telat_menit
            payroll.total_lembur_menit = total_lembur_menit
            payroll.potongan_alpha = 0
            payroll.potongan_telat = 0
            payroll.uang_lembur = uang_lembur
            payroll.upah_freelance = upah_freelance
            payroll.bonus_target = bonus_target
            payroll.co_host_fee = co_host_fee
            payroll.bpjs_jkk = 0
            payroll.bpjs_jkm = 0
            payroll.bpjs_jht = 0
            payroll.bpjs_kesehatan = 0
            payroll.gaji_bersih = gaji_bersih

        db.session.commit()
        flash(f"Slip gaji Freelance {BULAN_NAMA[bulan]} {tahun} berhasil dibuat/diperbarui.", "success")
        return redirect(url_for("penggajian_list", bulan=bulan, tahun=tahun))

    @app.route("/penggajian/<int:payroll_id>")
    @admin_required
    def penggajian_detail(payroll_id):
        payroll = db.session.get(Payroll, payroll_id) or abort_404()
        settings = get_settings()
        periode_awal = date(payroll.tahun, payroll.bulan, 1)
        periode_akhir = date(payroll.tahun, payroll.bulan, calendar.monthrange(payroll.tahun, payroll.bulan)[1])
        tarif_unit_efektif = (
            payroll.employee.tarif_unit_freelance or settings.tarif_harian_freelance or 0
        )
        return render_template(
            "payroll_detail.html",
            p=payroll,
            bulan_nama=BULAN_NAMA,
            settings=settings,
            periode_awal=periode_awal,
            periode_akhir=periode_akhir,
            tarif_unit_efektif=tarif_unit_efektif,
            wa_link=buat_wa_link(payroll, settings),
        )

    def pdf_link_callback(uri, rel):
        if uri.startswith("/static/"):
            return os.path.join(app.root_path, uri.lstrip("/"))
        return uri

    def nama_file_slip(payroll):
        return f"Slip_Gaji_{payroll.employee.nama.replace(' ', '_')}_{BULAN_NAMA[payroll.bulan]}_{payroll.tahun}.pdf"

    def buat_slip_pdf_bytes(payroll, settings):
        periode_awal = date(payroll.tahun, payroll.bulan, 1)
        periode_akhir = date(payroll.tahun, payroll.bulan, calendar.monthrange(payroll.tahun, payroll.bulan)[1])
        tarif_unit_efektif = (
            payroll.employee.tarif_unit_freelance or settings.tarif_harian_freelance or 0
        )
        logo_path = None
        if settings.logo_filename:
            candidate = os.path.join(app.config["UPLOAD_FOLDER"], settings.logo_filename)
            if os.path.exists(candidate):
                logo_path = candidate

        html = render_template(
            "payroll_slip_pdf.html",
            p=payroll,
            settings=settings,
            periode_awal=periode_awal,
            periode_akhir=periode_akhir,
            tarif_unit_efektif=tarif_unit_efektif,
            logo_path=logo_path,
        )

        buffer = BytesIO()
        pisa.CreatePDF(html, dest=buffer, link_callback=pdf_link_callback)
        return buffer.getvalue()

    def kirim_email_slip(payroll, settings):
        email_tujuan = (payroll.employee.email or "").strip()
        if not email_tujuan:
            return False, "Karyawan belum punya alamat email terdaftar."

        smtp_email = os.environ.get("SMTP_EMAIL")
        smtp_password = os.environ.get("SMTP_APP_PASSWORD")
        if not smtp_email or not smtp_password:
            return False, "Pengiriman email belum dikonfigurasi di server (SMTP_EMAIL/SMTP_APP_PASSWORD)."

        pdf_bytes = buat_slip_pdf_bytes(payroll, settings)

        msg = EmailMessage()
        msg["Subject"] = f"Slip Gaji {BULAN_NAMA[payroll.bulan]} {payroll.tahun} - {settings.nama_perusahaan}"
        msg["From"] = smtp_email
        msg["To"] = email_tujuan
        msg.set_content(
            f"Halo {payroll.employee.nama},\n\n"
            f"Berikut slip gaji Anda periode {BULAN_NAMA[payroll.bulan]} {payroll.tahun} dari "
            f"{settings.nama_perusahaan} (rincian lengkap ada di file PDF terlampir).\n\n"
            "Status: SUDAH DIBAYAR\n\n"
            "Terima kasih atas kerja keras Anda selama ini. Jika ada pertanyaan silakan hubungi admin."
        )
        msg.add_attachment(
            pdf_bytes, maintype="application", subtype="pdf", filename=nama_file_slip(payroll)
        )

        try:
            with smtplib.SMTP("smtp.gmail.com", 587) as server:
                server.starttls()
                server.login(smtp_email, smtp_password)
                server.send_message(msg)
            return True, None
        except Exception as e:
            return False, str(e)

    @app.route("/penggajian/<int:payroll_id>/pdf")
    @admin_required
    def penggajian_pdf(payroll_id):
        payroll = db.session.get(Payroll, payroll_id) or abort_404()
        settings = get_settings()
        pdf_bytes = buat_slip_pdf_bytes(payroll, settings)
        return Response(
            pdf_bytes,
            mimetype="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{nama_file_slip(payroll)}"'},
        )

    @app.route("/penggajian/<int:payroll_id>/kirim-email", methods=["POST"])
    @admin_required
    def penggajian_kirim_email(payroll_id):
        payroll = db.session.get(Payroll, payroll_id) or abort_404()
        settings = get_settings()
        berhasil, pesan_error = kirim_email_slip(payroll, settings)
        if berhasil:
            flash(f"Slip gaji berhasil dikirim ke email {payroll.employee.email}.", "success")
        else:
            flash(f"Gagal mengirim email: {pesan_error}", "danger")
        return redirect(url_for("penggajian_detail", payroll_id=payroll.id))

    @app.route("/penggajian/<int:payroll_id>/bayar", methods=["POST"])
    @admin_required
    def penggajian_bayar(payroll_id):
        payroll = db.session.get(Payroll, payroll_id) or abort_404()
        payroll.status = "Dibayar"
        payroll.tanggal_dibayar = now_wib()
        db.session.commit()

        settings = get_settings()
        berhasil, pesan_error = kirim_email_slip(payroll, settings)
        if berhasil:
            flash(
                f"Gaji {payroll.employee.nama} ditandai sudah dibayar. "
                f"Slip terkirim otomatis ke email {payroll.employee.email}.",
                "success",
            )
        else:
            flash(
                f"Gaji {payroll.employee.nama} ditandai sudah dibayar. "
                f"Email slip tidak terkirim otomatis: {pesan_error}",
                "warning",
            )
        return redirect(url_for("penggajian_list", bulan=payroll.bulan, tahun=payroll.tahun))

    # ---------- PENGATURAN ----------
    @app.route("/pengaturan", methods=["GET", "POST"])
    @admin_required
    def pengaturan():
        settings = get_settings()
        if request.method == "POST":
            settings.nama_perusahaan = request.form.get("nama_perusahaan", "").strip()
            settings.jam_masuk_standar = request.form.get("jam_masuk_standar")
            settings.jam_pulang_standar = request.form.get("jam_pulang_standar")
            settings.jam_masuk_standar_freelance = request.form.get("jam_masuk_standar_freelance")
            settings.jam_pulang_standar_freelance = request.form.get("jam_pulang_standar_freelance")
            settings.hari_kerja_per_bulan = int(request.form.get("hari_kerja_per_bulan") or 22)
            settings.toleransi_telat_menit = int(request.form.get("toleransi_telat_menit") or 0)
            settings.toleransi_lembur_menit = int(request.form.get("toleransi_lembur_menit") or 30)
            settings.denda_telat_per_menit = int(request.form.get("denda_telat_per_menit") or 0)
            settings.upah_lembur_per_jam = int(request.form.get("upah_lembur_per_jam") or 0)
            settings.tarif_harian_freelance = int(request.form.get("tarif_harian_freelance") or 0)
            settings.upah_lembur_freelance_per_jam = int(request.form.get("upah_lembur_freelance_per_jam") or 0)
            settings.tarif_co_host = int(request.form.get("tarif_co_host") or 0)
            settings.bonus_target_tercapai = int(request.form.get("bonus_target_tercapai") or 0)
            settings.alamat_perusahaan = request.form.get("alamat_perusahaan", "").strip()
            settings.kontak_perusahaan = request.form.get("kontak_perusahaan", "").strip()
            settings.instagram_perusahaan = request.form.get("instagram_perusahaan", "").strip()
            settings.kota_perusahaan = request.form.get("kota_perusahaan", "").strip()

            kantor_lat = request.form.get("kantor_lat", "").strip()
            kantor_lng = request.form.get("kantor_lng", "").strip()
            settings.kantor_lat = float(kantor_lat) if kantor_lat else None
            settings.kantor_lng = float(kantor_lng) if kantor_lng else None
            settings.radius_kantor_meter = int(request.form.get("radius_kantor_meter") or 100)

            logo_file = request.files.get("logo")
            if logo_file and logo_file.filename:
                ext = logo_file.filename.rsplit(".", 1)[-1].lower()
                if ext in ("png", "jpg", "jpeg", "webp"):
                    filename = secure_filename(f"logo.{ext}")
                    logo_file.save(os.path.join(app.config["UPLOAD_FOLDER"], filename))
                    settings.logo_filename = filename
                else:
                    flash("Format logo harus PNG, JPG, atau WEBP.", "danger")

            db.session.commit()
            flash("Pengaturan berhasil disimpan.", "success")
            return redirect(url_for("pengaturan"))
        return render_template("settings.html", settings=settings)

    # ---------- MARKETING: ANALISA IKLAN PER MARKETPLACE ----------
    def bersihkan_tmp_iklan_lama():
        batas = time.time() - 24 * 3600
        try:
            for nama in os.listdir(app.config["TMP_IKLAN_FOLDER"]):
                path_file = os.path.join(app.config["TMP_IKLAN_FOLDER"], nama)
                if os.path.isfile(path_file) and os.path.getmtime(path_file) < batas:
                    os.remove(path_file)
        except OSError:
            pass

    @app.route("/marketing/iklan")
    @marketing_required
    def marketing_iklan_dashboard():
        marketplace_filter = request.args.get("marketplace", "Semua")
        hari_ini = today_wib()
        default_dari = hari_ini.replace(day=1)
        try:
            dari = datetime.strptime(request.args.get("dari", ""), "%Y-%m-%d").date()
        except ValueError:
            dari = default_dari
        try:
            sampai = datetime.strptime(request.args.get("sampai", ""), "%Y-%m-%d").date()
        except ValueError:
            sampai = hari_ini

        query = IklanMarketplace.query.filter(
            IklanMarketplace.tanggal >= dari, IklanMarketplace.tanggal <= sampai
        )
        semua_data = query.order_by(IklanMarketplace.tanggal).all()

        if marketplace_filter != "Semua":
            data_terfilter = [d for d in semua_data if d.marketplace == marketplace_filter]
        else:
            data_terfilter = semua_data

        def totalkan(items):
            return {
                "biaya": sum(d.biaya or 0 for d in items),
                "impresi": sum(d.impresi or 0 for d in items),
                "klik": sum(d.klik or 0 for d in items),
                "pesanan": sum(d.pesanan or 0 for d in items),
                "omzet": sum(d.omzet or 0 for d in items),
            }

        total = totalkan(data_terfilter)
        ringkasan = {
            **total,
            "roas": (total["omzet"] / total["biaya"]) if total["biaya"] else 0,
            "ctr": (total["klik"] / total["impresi"] * 100) if total["impresi"] else 0,
            "cpc": (total["biaya"] / total["klik"]) if total["klik"] else 0,
            "cpa": (total["biaya"] / total["pesanan"]) if total["pesanan"] else 0,
            "konversi_rate": (total["pesanan"] / total["klik"] * 100) if total["klik"] else 0,
        }

        breakdown = []
        for mp in MARKETPLACE_LIST:
            item_mp = [d for d in semua_data if d.marketplace == mp]
            if not item_mp:
                continue
            t = totalkan(item_mp)
            breakdown.append({
                "marketplace": mp,
                **t,
                "roas": (t["omzet"] / t["biaya"]) if t["biaya"] else 0,
                "cpa": (t["biaya"] / t["pesanan"]) if t["pesanan"] else 0,
            })

        tren_map = {}
        for d in data_terfilter:
            key = d.tanggal.isoformat()
            if key not in tren_map:
                tren_map[key] = {"biaya": 0, "omzet": 0, "klik": 0}
            tren_map[key]["biaya"] += d.biaya or 0
            tren_map[key]["omzet"] += d.omzet or 0
            tren_map[key]["klik"] += d.klik or 0
        tren_tanggal = sorted(tren_map.keys())
        tren_biaya = [tren_map[k]["biaya"] for k in tren_tanggal]
        tren_omzet = [tren_map[k]["omzet"] for k in tren_tanggal]

        return render_template(
            "marketing/iklan_dashboard.html",
            marketplace_list=MARKETPLACE_LIST,
            marketplace_filter=marketplace_filter,
            dari=dari,
            sampai=sampai,
            ringkasan=ringkasan,
            breakdown=breakdown,
            data_list=list(reversed(data_terfilter)),
            tren_tanggal=tren_tanggal,
            tren_biaya=tren_biaya,
            tren_omzet=tren_omzet,
        )

    def simpan_tmp_upload(prefix, marketplace, file):
        headers, rows, error = baca_file_iklan(file)
        if error:
            return None, None, None, error
        rows_bersih = []
        for row in rows:
            row_lengkap = list(row) + [""] * (len(headers) - len(row))
            row_serial = [
                c.isoformat() if isinstance(c, (datetime, date)) else c
                for c in row_lengkap[: len(headers)]
            ]
            rows_bersih.append(row_serial)
        token = uuid.uuid4().hex
        path_tmp = os.path.join(app.config["TMP_IKLAN_FOLDER"], f"{prefix}_{token}.json")
        with open(path_tmp, "w", encoding="utf-8") as f:
            json.dump({
                "marketplace": marketplace,
                "headers": headers,
                "rows": rows_bersih,
                "sumber_file": secure_filename(file.filename),
            }, f)
        return token, headers, rows_bersih, None

    @app.route("/marketing/iklan/upload", methods=["GET", "POST"])
    @marketing_required
    def marketing_iklan_upload():
        bersihkan_tmp_iklan_lama()
        if request.method == "POST":
            marketplace = request.form.get("marketplace", "")
            file = request.files.get("file")
            if marketplace not in MARKETPLACE_LIST:
                flash("Pilih marketplace terlebih dahulu.", "danger")
                return redirect(url_for("marketing_iklan_upload"))
            if not file or not file.filename:
                flash("Pilih file laporan iklan (CSV/XLSX) terlebih dahulu.", "danger")
                return redirect(url_for("marketing_iklan_upload"))

            token, headers, rows_bersih, error = simpan_tmp_upload("iklan", marketplace, file)
            if error:
                flash(error, "danger")
                return redirect(url_for("marketing_iklan_upload"))

            tebakan = tebak_kolom(headers, KOLOM_TARGET_IKLAN)
            return render_template(
                "marketing/iklan_mapping.html",
                token=token,
                marketplace=marketplace,
                headers=headers,
                preview_rows=rows_bersih[:8],
                kolom_target=KOLOM_TARGET_IKLAN,
                tebakan=tebakan,
                jumlah_baris=len(rows_bersih),
                judul="Cocokkan Kolom",
                konfirmasi_url=url_for("marketing_iklan_konfirmasi"),
                upload_url=url_for("marketing_iklan_upload"),
            )

        return render_template("marketing/iklan_upload.html", marketplace_list=MARKETPLACE_LIST)

    @app.route("/marketing/iklan/konfirmasi", methods=["POST"])
    @marketing_required
    def marketing_iklan_konfirmasi():
        token = request.form.get("token", "")
        path_tmp = os.path.join(app.config["TMP_IKLAN_FOLDER"], f"iklan_{token}.json")
        if not os.path.isfile(path_tmp):
            flash("Sesi upload sudah kedaluwarsa, silakan upload ulang file.", "danger")
            return redirect(url_for("marketing_iklan_upload"))

        with open(path_tmp, "r", encoding="utf-8") as f:
            data_tmp = json.load(f)

        mapping = {}
        for key, label, wajib, _kk in KOLOM_TARGET_IKLAN:
            nilai = request.form.get(f"map_{key}", "")
            mapping[key] = int(nilai) if nilai != "" else None
            if wajib and mapping[key] is None:
                flash(f"Kolom '{label}' wajib dipilih.", "danger")
                os.remove(path_tmp)
                return redirect(url_for("marketing_iklan_upload"))

        agregat = {}
        dilewati = 0
        for row in data_tmp["rows"]:
            def ambil(key):
                idx = mapping.get(key)
                if idx is None or idx >= len(row):
                    return None
                return row[idx]

            tanggal = parse_tanggal_iklan(ambil("tanggal"))
            if not tanggal:
                dilewati += 1
                continue

            if tanggal not in agregat:
                agregat[tanggal] = {"biaya": 0, "impresi": 0, "klik": 0, "pesanan": 0, "omzet": 0}
            for key in ("biaya", "impresi", "klik", "pesanan", "omzet"):
                nilai_mentah = ambil(key)
                if nilai_mentah is not None:
                    agregat[tanggal][key] += parse_angka_iklan(nilai_mentah)

        marketplace = data_tmp["marketplace"]
        sumber_file = data_tmp.get("sumber_file", "")
        for tanggal, nilai in agregat.items():
            existing = IklanMarketplace.query.filter_by(marketplace=marketplace, tanggal=tanggal).first()
            if not existing:
                existing = IklanMarketplace(marketplace=marketplace, tanggal=tanggal)
                db.session.add(existing)
            existing.biaya = round(nilai["biaya"])
            existing.impresi = round(nilai["impresi"])
            existing.klik = round(nilai["klik"])
            existing.pesanan = round(nilai["pesanan"])
            existing.omzet = round(nilai["omzet"])
            existing.sumber_file = sumber_file
            existing.dibuat_pada = now_wib()
        db.session.commit()
        os.remove(path_tmp)

        if agregat:
            tgl_min = min(agregat.keys())
            tgl_max = max(agregat.keys())
            pesan = f"Berhasil impor {len(agregat)} hari data iklan {marketplace} ({tgl_min.strftime('%d/%m/%Y')} - {tgl_max.strftime('%d/%m/%Y')})."
            if dilewati:
                pesan += f" {dilewati} baris dilewati karena tanggal tidak terbaca."
            flash(pesan, "success")
        else:
            flash("Tidak ada baris data yang berhasil diproses (format tanggal tidak dikenali).", "danger")

        return redirect(url_for("marketing_iklan_dashboard", marketplace=marketplace))

    @app.route("/marketing/iklan/manual", methods=["POST"])
    @marketing_required
    def marketing_iklan_manual():
        marketplace = request.form.get("marketplace", "")
        try:
            tanggal = datetime.strptime(request.form.get("tanggal", ""), "%Y-%m-%d").date()
        except ValueError:
            tanggal = None

        if marketplace not in MARKETPLACE_LIST or not tanggal:
            flash("Marketplace dan tanggal wajib diisi dengan benar.", "danger")
            return redirect(url_for("marketing_iklan_dashboard"))

        existing = IklanMarketplace.query.filter_by(marketplace=marketplace, tanggal=tanggal).first()
        if not existing:
            existing = IklanMarketplace(marketplace=marketplace, tanggal=tanggal)
            db.session.add(existing)
        existing.biaya = round(parse_angka_iklan(request.form.get("biaya", "0")))
        existing.impresi = round(parse_angka_iklan(request.form.get("impresi", "0")))
        existing.klik = round(parse_angka_iklan(request.form.get("klik", "0")))
        existing.pesanan = round(parse_angka_iklan(request.form.get("pesanan", "0")))
        existing.omzet = round(parse_angka_iklan(request.form.get("omzet", "0")))
        existing.sumber_file = "Input manual"
        existing.dibuat_pada = now_wib()
        db.session.commit()
        flash(f"Data iklan {marketplace} tanggal {tanggal.strftime('%d/%m/%Y')} berhasil disimpan.", "success")
        return redirect(url_for("marketing_iklan_dashboard", marketplace=marketplace))

    @app.route("/marketing/iklan/hapus/<int:iklan_id>", methods=["POST"])
    @marketing_required
    def marketing_iklan_hapus(iklan_id):
        data = db.session.get(IklanMarketplace, iklan_id)
        if data:
            marketplace = data.marketplace
            db.session.delete(data)
            db.session.commit()
            flash("Data iklan berhasil dihapus.", "success")
            return redirect(url_for("marketing_iklan_dashboard", marketplace=marketplace))
        flash("Data tidak ditemukan.", "danger")
        return redirect(url_for("marketing_iklan_dashboard"))

    @app.route("/marketing/produk/upload", methods=["GET", "POST"])
    @marketing_required
    def marketing_produk_upload():
        bersihkan_tmp_iklan_lama()
        if request.method == "POST":
            marketplace = request.form.get("marketplace", "")
            file = request.files.get("file")
            if marketplace not in MARKETPLACE_LIST:
                flash("Pilih marketplace terlebih dahulu.", "danger")
                return redirect(url_for("marketing_produk_upload"))
            if not file or not file.filename:
                flash("Pilih file laporan iklan per produk (CSV/XLSX) terlebih dahulu.", "danger")
                return redirect(url_for("marketing_produk_upload"))

            token, headers, rows_bersih, error = simpan_tmp_upload("produk", marketplace, file)
            if error:
                flash(error, "danger")
                return redirect(url_for("marketing_produk_upload"))

            tebakan = tebak_kolom(headers, KOLOM_TARGET_PRODUK)
            return render_template(
                "marketing/iklan_mapping.html",
                token=token,
                marketplace=marketplace,
                headers=headers,
                preview_rows=rows_bersih[:8],
                kolom_target=KOLOM_TARGET_PRODUK,
                tebakan=tebakan,
                jumlah_baris=len(rows_bersih),
                judul="Cocokkan Kolom — Laporan Per Produk",
                konfirmasi_url=url_for("marketing_produk_konfirmasi"),
                upload_url=url_for("marketing_produk_upload"),
            )

        return render_template("marketing/produk_upload.html", marketplace_list=MARKETPLACE_LIST)

    @app.route("/marketing/produk/konfirmasi", methods=["POST"])
    @marketing_required
    def marketing_produk_konfirmasi():
        token = request.form.get("token", "")
        path_tmp = os.path.join(app.config["TMP_IKLAN_FOLDER"], f"produk_{token}.json")
        if not os.path.isfile(path_tmp):
            flash("Sesi upload sudah kedaluwarsa, silakan upload ulang file.", "danger")
            return redirect(url_for("marketing_produk_upload"))

        with open(path_tmp, "r", encoding="utf-8") as f:
            data_tmp = json.load(f)

        mapping = {}
        for key, label, wajib, _kk in KOLOM_TARGET_PRODUK:
            nilai = request.form.get(f"map_{key}", "")
            mapping[key] = int(nilai) if nilai != "" else None
            if wajib and mapping[key] is None:
                flash(f"Kolom '{label}' wajib dipilih.", "danger")
                os.remove(path_tmp)
                return redirect(url_for("marketing_produk_upload"))

        agregat = {}
        dilewati = 0
        for row in data_tmp["rows"]:
            def ambil(key):
                idx = mapping.get(key)
                if idx is None or idx >= len(row):
                    return None
                return row[idx]

            nama_produk_raw = ambil("nama_produk")
            nama_produk = str(nama_produk_raw).strip() if nama_produk_raw is not None else ""
            tanggal = parse_tanggal_iklan(ambil("tanggal"))
            if not nama_produk or not tanggal:
                dilewati += 1
                continue

            kunci = (nama_produk, tanggal)
            if kunci not in agregat:
                agregat[kunci] = {"biaya": 0, "impresi": 0, "klik": 0, "pesanan": 0, "omzet": 0}
            for k in ("biaya", "impresi", "klik", "pesanan", "omzet"):
                nilai_mentah = ambil(k)
                if nilai_mentah is not None:
                    agregat[kunci][k] += parse_angka_iklan(nilai_mentah)

        marketplace = data_tmp["marketplace"]
        sumber_file = data_tmp.get("sumber_file", "")
        for (nama_produk, tanggal), nilai in agregat.items():
            existing = ProdukIklan.query.filter_by(
                marketplace=marketplace, nama_produk=nama_produk, tanggal=tanggal
            ).first()
            if not existing:
                existing = ProdukIklan(marketplace=marketplace, nama_produk=nama_produk, tanggal=tanggal)
                db.session.add(existing)
            existing.biaya = round(nilai["biaya"])
            existing.impresi = round(nilai["impresi"])
            existing.klik = round(nilai["klik"])
            existing.pesanan = round(nilai["pesanan"])
            existing.omzet = round(nilai["omzet"])
            existing.sumber_file = sumber_file
            existing.dibuat_pada = now_wib()
        db.session.commit()
        os.remove(path_tmp)

        if agregat:
            jumlah_produk = len(set(k[0] for k in agregat.keys()))
            pesan = f"Berhasil impor {len(agregat)} baris data ({jumlah_produk} produk) untuk {marketplace}."
            if dilewati:
                pesan += f" {dilewati} baris dilewati karena nama produk/tanggal tidak terbaca."
            flash(pesan, "success")
        else:
            flash("Tidak ada baris data yang berhasil diproses.", "danger")

        return redirect(url_for("marketing_produk_dashboard", marketplace=marketplace))

    @app.route("/marketing/produk")
    @marketing_required
    def marketing_produk_dashboard():
        hari_ini = today_wib()
        marketplace_filter = request.args.get("marketplace", "Semua")
        try:
            periode_hari = max(int(request.args.get("periode", 30)), 1)
        except ValueError:
            periode_hari = 30
        try:
            threshold_ctr = float(request.args.get("ctr_min", 1))
        except ValueError:
            threshold_ctr = 1.0
        try:
            threshold_roas = float(request.args.get("roas_min", 3))
        except ValueError:
            threshold_roas = 3.0
        try:
            threshold_pesanan14 = int(request.args.get("pesanan14_min", 3))
        except ValueError:
            threshold_pesanan14 = 3

        dari_periode = hari_ini - timedelta(days=periode_hari - 1)
        dari_14 = hari_ini - timedelta(days=13)

        query_periode = ProdukIklan.query.filter(
            ProdukIklan.tanggal >= dari_periode, ProdukIklan.tanggal <= hari_ini
        )
        query_14 = ProdukIklan.query.filter(
            ProdukIklan.tanggal >= dari_14, ProdukIklan.tanggal <= hari_ini
        )
        if marketplace_filter != "Semua":
            query_periode = query_periode.filter(ProdukIklan.marketplace == marketplace_filter)
            query_14 = query_14.filter(ProdukIklan.marketplace == marketplace_filter)

        pesanan14_map = {}
        for d in query_14.all():
            kunci = (d.marketplace, d.nama_produk)
            pesanan14_map[kunci] = pesanan14_map.get(kunci, 0) + (d.pesanan or 0)

        produk_map = {}
        for d in query_periode.all():
            kunci = (d.marketplace, d.nama_produk)
            if kunci not in produk_map:
                produk_map[kunci] = {
                    "marketplace": d.marketplace, "nama_produk": d.nama_produk,
                    "biaya": 0, "impresi": 0, "klik": 0, "pesanan": 0, "omzet": 0,
                }
            produk_map[kunci]["biaya"] += d.biaya or 0
            produk_map[kunci]["impresi"] += d.impresi or 0
            produk_map[kunci]["klik"] += d.klik or 0
            produk_map[kunci]["pesanan"] += d.pesanan or 0
            produk_map[kunci]["omzet"] += d.omzet or 0

        tidak_perform = []
        perform = []
        for kunci, p in produk_map.items():
            if p["biaya"] <= 0:
                continue
            ctr = (p["klik"] / p["impresi"] * 100) if p["impresi"] else 0
            roas = (p["omzet"] / p["biaya"]) if p["biaya"] else 0
            pesanan14 = pesanan14_map.get(kunci, 0)

            alasan = []
            rekomendasi = []
            if p["impresi"] > 0 and ctr < threshold_ctr:
                alasan.append(f"CTR rendah ({ctr:.2f}% < {threshold_ctr:.2f}%)")
                rekomendasi.append("Ganti foto utama & judul produk, uji materi iklan baru agar lebih menarik diklik.")
            if roas < threshold_roas:
                alasan.append(f"ROAS rendah ({roas:.2f}x < {threshold_roas:.2f}x)")
                rekomendasi.append("Klik sudah ada tapi konversi kurang — cek harga, deskripsi, ulasan & foto produk, atau turunkan bid iklan.")
            if pesanan14 < threshold_pesanan14:
                alasan.append(f"Penjualan 14 hari terakhir rendah ({pesanan14} < {threshold_pesanan14})")
                rekomendasi.append("Penjualan masih minim — naikkan budget bertahap sambil dipantau, atau alihkan budget bila 1-2 minggu tidak membaik.")

            item = {**p, "ctr": ctr, "roas": roas, "pesanan14": pesanan14, "alasan": alasan, "rekomendasi": rekomendasi}
            if alasan:
                tidak_perform.append(item)
            else:
                if roas >= threshold_roas * 1.5:
                    item["rekomendasi"] = ["Performa baik, pertimbangkan naikkan budget iklan bertahap (+20-30%) untuk maksimalkan momentum penjualan."]
                else:
                    item["rekomendasi"] = ["Performa sudah sesuai target, pertahankan strategi saat ini."]
                perform.append(item)

        tidak_perform.sort(key=lambda x: x["roas"])
        perform.sort(key=lambda x: -x["roas"])

        return render_template(
            "marketing/produk_dashboard.html",
            marketplace_list=MARKETPLACE_LIST,
            marketplace_filter=marketplace_filter,
            periode_hari=periode_hari,
            threshold_ctr=threshold_ctr,
            threshold_roas=threshold_roas,
            threshold_pesanan14=threshold_pesanan14,
            tidak_perform=tidak_perform,
            perform=perform,
            total_produk=len(tidak_perform) + len(perform),
        )

    @app.route("/marketing/proyeksi")
    @marketing_required
    def marketing_proyeksi():
        hari_ini = today_wib()
        periode_hari = 30
        dari = hari_ini - timedelta(days=periode_hari - 1)
        data = IklanMarketplace.query.filter(
            IklanMarketplace.tanggal >= dari, IklanMarketplace.tanggal <= hari_ini
        ).all()

        bulan_depan = hari_ini.month + 1
        tahun_depan = hari_ini.year
        if bulan_depan > 12:
            bulan_depan = 1
            tahun_depan += 1
        jumlah_hari_bulan_depan = calendar.monthrange(tahun_depan, bulan_depan)[1]

        proyeksi = []
        for mp in MARKETPLACE_LIST:
            item = [d for d in data if d.marketplace == mp]
            if not item:
                continue
            jumlah_hari_data = len(set(d.tanggal for d in item))
            total_biaya = sum(d.biaya or 0 for d in item)
            total_omzet = sum(d.omzet or 0 for d in item)
            avg_biaya_harian = (total_biaya / jumlah_hari_data) if jumlah_hari_data else 0
            roas_historis = (total_omzet / total_biaya) if total_biaya else 0
            proyeksi_budget = avg_biaya_harian * jumlah_hari_bulan_depan
            proyeksi.append({
                "marketplace": mp,
                "avg_biaya_harian": avg_biaya_harian,
                "roas_historis": roas_historis,
                "proyeksi_budget": proyeksi_budget,
                "proyeksi_omzet": proyeksi_budget * roas_historis,
            })

        target_marketplace = request.args.get("target_marketplace", "")
        try:
            target_omzet = float(request.args.get("target_omzet", "") or 0)
        except ValueError:
            target_omzet = 0

        hasil_target = None
        if target_marketplace and target_omzet > 0:
            if target_marketplace == "Semua":
                item = data
            else:
                item = [d for d in data if d.marketplace == target_marketplace]
            total_biaya_t = sum(d.biaya or 0 for d in item)
            total_omzet_t = sum(d.omzet or 0 for d in item)
            roas_pakai = (total_omzet_t / total_biaya_t) if total_biaya_t else 0
            if roas_pakai > 0:
                hasil_target = {
                    "marketplace": target_marketplace,
                    "target_omzet": target_omzet,
                    "roas_pakai": roas_pakai,
                    "budget_dibutuhkan": target_omzet / roas_pakai,
                }
            else:
                hasil_target = {"marketplace": target_marketplace, "error": True}

        return render_template(
            "marketing/proyeksi.html",
            marketplace_list=MARKETPLACE_LIST,
            proyeksi=proyeksi,
            periode_hari=periode_hari,
            bulan_depan_label=f"{BULAN_NAMA[bulan_depan]} {tahun_depan}",
            target_marketplace=target_marketplace,
            target_omzet=target_omzet,
            hasil_target=hasil_target,
        )

    @app.route("/akun", methods=["GET", "POST"])
    @admin_required
    def akun():
        if request.method == "POST":
            password_baru = request.form.get("password_baru", "")
            password_konfirmasi = request.form.get("password_konfirmasi", "")
            if len(password_baru) < 6:
                flash("Password baru minimal 6 karakter.", "danger")
            elif password_baru != password_konfirmasi:
                flash("Konfirmasi password tidak cocok.", "danger")
            else:
                current_user.set_password(password_baru)
                db.session.commit()
                flash("Password berhasil diubah.", "success")
            return redirect(url_for("akun"))
        return render_template("account.html")

    def abort_404():
        from flask import abort
        abort(404)

    @app.context_processor
    def inject_globals():
        pending_izin = 0
        pending_lembur = 0
        if current_user.is_authenticated and getattr(current_user, "role", None) == "admin":
            pending_izin = PengajuanIzin.query.filter_by(status="Menunggu").count()
            pending_lembur = PengajuanLembur.query.filter_by(status="Menunggu").count()
        return {
            "bulan_nama_list": BULAN_NAMA,
            "pending_izin_count": pending_izin,
            "pending_lembur_count": pending_lembur,
            "site_settings": get_settings(),
        }

    return app


app = create_app()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
