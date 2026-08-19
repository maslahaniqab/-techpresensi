from datetime import datetime
from zoneinfo import ZoneInfo
from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

db = SQLAlchemy()

WIB = ZoneInfo("Asia/Jakarta")


def now_wib():
    return datetime.now(WIB).replace(tzinfo=None)


class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(64), unique=True, nullable=False)
    nama = db.Column(db.String(128), nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    role = "admin"

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def get_id(self):
        return f"admin:{self.id}"


class Settings(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    jam_masuk_standar = db.Column(db.String(5), default="08:00")
    jam_pulang_standar = db.Column(db.String(5), default="17:00")
    jam_masuk_standar_freelance = db.Column(db.String(5), default="08:00")
    jam_pulang_standar_freelance = db.Column(db.String(5), default="17:00")
    hari_kerja_per_bulan = db.Column(db.Integer, default=22)
    toleransi_telat_menit = db.Column(db.Integer, default=0)
    toleransi_lembur_menit = db.Column(db.Integer, default=30)
    denda_telat_per_menit = db.Column(db.Integer, default=1000)
    upah_lembur_per_jam = db.Column(db.Integer, default=20000)
    tarif_harian_freelance = db.Column(db.Integer, default=50000)
    upah_lembur_freelance_per_jam = db.Column(db.Integer, default=15000)
    tarif_co_host = db.Column(db.Integer, default=25000)
    bonus_target_tercapai = db.Column(db.Integer, default=100000)
    nama_perusahaan = db.Column(db.String(128), default="Perusahaan Saya")
    alamat_perusahaan = db.Column(db.String(256), default="")
    kontak_perusahaan = db.Column(db.String(64), default="")
    instagram_perusahaan = db.Column(db.String(64), default="")
    kota_perusahaan = db.Column(db.String(64), default="Bandung")
    logo_filename = db.Column(db.String(128), default="")
    kantor_lat = db.Column(db.Float)
    kantor_lng = db.Column(db.Float)
    radius_kantor_meter = db.Column(db.Integer, default=100)


class Employee(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nama = db.Column(db.String(128), nullable=False)
    jabatan = db.Column(db.String(64))
    email = db.Column(db.String(128))
    no_hp = db.Column(db.String(32))
    password_hash = db.Column(db.String(256))
    role = "pegawai"
    alamat = db.Column(db.String(256))
    nomor_rekening = db.Column(db.String(64))
    gaji_pokok = db.Column(db.Integer, nullable=False, default=0)
    tunjangan_makan = db.Column(db.Integer, nullable=False, default=0)
    tunjangan_transport = db.Column(db.Integer, nullable=False, default=0)
    tipe_pegawai = db.Column(db.String(20), default="Karyawan Tetap")  # Karyawan Tetap / Probation / Freelance
    target_tercapai = db.Column(db.String(16), default="Tidak Tercapai")  # Tercapai / Tidak Tercapai
    bpjs_jkk = db.Column(db.Integer, nullable=False, default=0)
    bpjs_jkm = db.Column(db.Integer, nullable=False, default=0)
    bpjs_jht = db.Column(db.Integer, nullable=False, default=0)
    bpjs_kesehatan = db.Column(db.Integer, nullable=False, default=0)
    tarif_unit_freelance = db.Column(db.Integer, nullable=False, default=0)  # 0 = pakai default di Pengaturan
    co_host_bulan_ini = db.Column(db.String(8), default="Tidak")  # Ya / Tidak
    status = db.Column(db.String(16), default="Aktif")  # Aktif / Nonaktif

    attendances = db.relationship(
        "Attendance", backref="employee", cascade="all, delete-orphan"
    )
    payrolls = db.relationship(
        "Payroll", backref="employee", cascade="all, delete-orphan"
    )

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        if not self.password_hash:
            return False
        return check_password_hash(self.password_hash, password)

    def get_id(self):
        return f"emp:{self.id}"


class Attendance(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employee.id"), nullable=False)
    tanggal = db.Column(db.Date, nullable=False)
    status = db.Column(db.String(16), nullable=False)  # Hadir/Sakit/Izin/Cuti/Alpha
    jam_masuk = db.Column(db.String(8))
    jam_pulang = db.Column(db.String(8))
    telat_menit = db.Column(db.Integer, default=0)
    lembur_menit = db.Column(db.Integer, default=0)
    catatan = db.Column(db.String(256))
    lokasi_masuk_lat = db.Column(db.Float)
    lokasi_masuk_lng = db.Column(db.Float)
    lokasi_pulang_lat = db.Column(db.Float)
    lokasi_pulang_lng = db.Column(db.Float)

    __table_args__ = (
        db.UniqueConstraint("employee_id", "tanggal", name="uq_employee_tanggal"),
    )


class Payroll(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employee.id"), nullable=False)
    bulan = db.Column(db.Integer, nullable=False)
    tahun = db.Column(db.Integer, nullable=False)

    gaji_pokok = db.Column(db.Integer, default=0)
    tunjangan_makan = db.Column(db.Integer, default=0)
    tunjangan_transport = db.Column(db.Integer, default=0)
    tipe_pegawai = db.Column(db.String(20), default="Karyawan Tetap")
    total_hadir = db.Column(db.Integer, default=0)
    total_sakit = db.Column(db.Integer, default=0)
    total_izin = db.Column(db.Integer, default=0)
    total_cuti = db.Column(db.Integer, default=0)
    total_alpha = db.Column(db.Integer, default=0)
    total_telat_menit = db.Column(db.Integer, default=0)
    total_lembur_menit = db.Column(db.Integer, default=0)

    potongan_alpha = db.Column(db.Integer, default=0)
    potongan_telat = db.Column(db.Integer, default=0)
    uang_lembur = db.Column(db.Integer, default=0)
    upah_freelance = db.Column(db.Integer, default=0)
    bonus_target = db.Column(db.Integer, default=0)
    co_host_fee = db.Column(db.Integer, default=0)
    bpjs_jkk = db.Column(db.Integer, default=0)
    bpjs_jkm = db.Column(db.Integer, default=0)
    bpjs_jht = db.Column(db.Integer, default=0)
    bpjs_kesehatan = db.Column(db.Integer, default=0)
    gaji_bersih = db.Column(db.Integer, default=0)

    status = db.Column(db.String(16), default="Draft")  # Draft / Dibayar
    tanggal_dibuat = db.Column(db.DateTime, default=now_wib)
    tanggal_dibayar = db.Column(db.DateTime)

    __table_args__ = (
        db.UniqueConstraint("employee_id", "bulan", "tahun", name="uq_employee_periode"),
    )


class PengajuanIzin(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employee.id"), nullable=False)
    tanggal = db.Column(db.Date, nullable=False)
    jenis = db.Column(db.String(16), nullable=False)  # Sakit / Izin / Cuti
    alasan = db.Column(db.String(512))
    dokumen_filename = db.Column(db.String(256))
    status = db.Column(db.String(16), default="Menunggu")  # Menunggu / Disetujui / Ditolak
    catatan_admin = db.Column(db.String(256))
    tanggal_diajukan = db.Column(db.DateTime, default=now_wib)
    tanggal_diproses = db.Column(db.DateTime)

    employee = db.relationship("Employee", backref="pengajuan_izin")
