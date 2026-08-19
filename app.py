import os
import math
import calendar
from functools import wraps
from datetime import datetime, date
from urllib.parse import quote

from flask import Flask, render_template, redirect, url_for, request, flash, jsonify
from flask_login import (
    LoginManager,
    login_user,
    logout_user,
    current_user,
)
from werkzeug.utils import secure_filename

from models import db, User, Employee, Attendance, Settings, Payroll, PengajuanIzin

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

    def rupiah(value):
        try:
            value = int(value or 0)
        except (ValueError, TypeError):
            value = 0
        return f"Rp {value:,.0f}".replace(",", ".")

    app.jinja_env.filters["rupiah"] = rupiah
    app.jinja_env.filters["terbilang"] = terbilang_rupiah

    def buat_pesan_wa_slip(p, settings, tarif_unit_efektif):
        baris = [
            f"Halo {p.employee.nama},",
            "",
            f"Berikut slip gaji Anda periode {BULAN_NAMA[p.bulan]} {p.tahun} dari {settings.nama_perusahaan}:",
            "",
        ]
        if p.tipe_pegawai == "Freelance":
            baris += [
                f"Upah Kerja ({p.total_hadir} hari x {rupiah(tarif_unit_efektif)}): {rupiah(p.upah_freelance)}",
                f"Lembur/Over Time: {rupiah(p.uang_lembur)}",
                f"Achieve Target: {rupiah(p.bonus_target)}",
                f"Co-Host: {rupiah(p.co_host_fee)}",
            ]
            label_total = "GAJI BRUTO"
        else:
            baris += [
                f"Gaji Pokok: {rupiah(p.gaji_pokok)}",
                f"Tj. Makan & Transport: {rupiah(p.tunjangan_makan + p.tunjangan_transport)}",
                f"Lembur Harian: {rupiah(p.uang_lembur)}",
                f"Bonus: {rupiah(p.bonus_target)}",
                f"Potongan Alpha: -{rupiah(p.potongan_alpha)}",
                f"Potongan Keterlambatan: -{rupiah(p.potongan_telat)}",
                f"BPJS (JKK+JKM+JHT+Kesehatan): -{rupiah(p.bpjs_jkk + p.bpjs_jkm + p.bpjs_jht + p.bpjs_kesehatan)}",
            ]
            label_total = "PENERIMAAN BERSIH"
        baris += [
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
        tarif_unit_efektif = p.employee.tarif_unit_freelance or settings.tarif_harian_freelance or 0
        pesan = buat_pesan_wa_slip(p, settings, tarif_unit_efektif)
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
        hari_ini = date.today()
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
        hari_ini = date.today()
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
        tanggal_str = request.args.get("tanggal", date.today().isoformat())
        try:
            tanggal = datetime.strptime(tanggal_str, "%Y-%m-%d").date()
        except ValueError:
            tanggal = date.today()

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
        tanggal_default = request.args.get("tanggal", date.today().isoformat())
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
        p.tanggal_diproses = datetime.utcnow()

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
        p.tanggal_diproses = datetime.utcnow()
        db.session.commit()
        flash(f"Pengajuan {p.employee.nama} ditolak.", "info")
        return redirect(url_for("pengajuan_izin_list"))

    # ---------- AREA PEGAWAI ----------
    @app.route("/pegawai")
    @pegawai_required
    def pegawai_dashboard():
        hari_ini = date.today()
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

        hari_ini = date.today()
        jam_sekarang = datetime.now().strftime("%H:%M:%S")
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

            if jenis not in ("Sakit", "Izin", "Cuti") or not tanggal:
                flash("Lengkapi tanggal dan jenis pengajuan dengan benar.", "danger")
            else:
                sudah_ada = PengajuanIzin.query.filter_by(
                    employee_id=current_user.id, tanggal=tanggal, status="Menunggu"
                ).first()
                if sudah_ada:
                    flash("Sudah ada pengajuan untuk tanggal ini yang masih menunggu persetujuan.", "warning")
                else:
                    db.session.add(
                        PengajuanIzin(
                            employee_id=current_user.id,
                            tanggal=tanggal,
                            jenis=jenis,
                            alasan=request.form.get("alasan", "").strip(),
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

    # ---------- PENGGAJIAN ----------
    @app.route("/penggajian")
    @admin_required
    def penggajian_list():
        bulan = int(request.args.get("bulan", date.today().month))
        tahun = int(request.args.get("tahun", date.today().year))
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
        bulan = int(request.args.get("bulan", date.today().month))
        tahun = int(request.args.get("tahun", date.today().year))
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

    @app.route("/penggajian/<int:payroll_id>/bayar", methods=["POST"])
    @admin_required
    def penggajian_bayar(payroll_id):
        payroll = db.session.get(Payroll, payroll_id) or abort_404()
        payroll.status = "Dibayar"
        payroll.tanggal_dibayar = datetime.utcnow()
        db.session.commit()
        flash(f"Gaji {payroll.employee.nama} ditandai sudah dibayar.", "success")
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
        pending = 0
        if current_user.is_authenticated and getattr(current_user, "role", None) == "admin":
            pending = PengajuanIzin.query.filter_by(status="Menunggu").count()
        return {
            "bulan_nama_list": BULAN_NAMA,
            "pending_izin_count": pending,
            "site_settings": get_settings(),
        }

    return app


app = create_app()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
