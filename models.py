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
    no_hp_perusahaan = db.Column(db.String(32), default="")
    email_perusahaan = db.Column(db.String(128), default="")
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
    bpjs_kesehatan_terdaftar = db.Column(db.Boolean, nullable=False, default=False)
    tarif_unit_freelance = db.Column(db.Integer, nullable=False, default=0)  # 0 = pakai default di Pengaturan
    co_host_bulan_ini = db.Column(db.String(8), default="Tidak")  # Ya / Tidak
    status = db.Column(db.String(16), default="Aktif")  # Aktif / Nonaktif
    akses_marketing = db.Column(db.Boolean, default=False, nullable=False)

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
    jumlah_lembur = db.Column(db.Integer, default=0)  # dihitung harian (hari), diisi manual oleh admin
    tarif_lembur = db.Column(db.Integer, default=0)  # Rp per hari lembur, diisi manual oleh admin
    uang_lembur = db.Column(db.Integer, default=0)  # = jumlah_lembur x tarif_lembur
    upah_freelance = db.Column(db.Integer, default=0)
    bonus_target = db.Column(db.Integer, default=0)
    co_host_fee = db.Column(db.Integer, default=0)
    bpjs_jkk = db.Column(db.Integer, default=0)
    bpjs_jkm = db.Column(db.Integer, default=0)
    bpjs_jht = db.Column(db.Integer, default=0)
    bpjs_kesehatan = db.Column(db.Integer, default=0)
    bpjs_kesehatan_perusahaan = db.Column(db.Integer, default=0)
    gaji_bersih = db.Column(db.Integer, default=0)

    status = db.Column(db.String(16), default="Draft")  # Draft / Dibayar
    tanggal_dibuat = db.Column(db.DateTime, default=now_wib)
    tanggal_dibayar = db.Column(db.DateTime)

    __table_args__ = (
        db.UniqueConstraint("employee_id", "bulan", "tahun", name="uq_employee_periode"),
    )


class HariLibur(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    tanggal = db.Column(db.Date, nullable=False, unique=True)
    keterangan = db.Column(db.String(128), nullable=False)


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


class LaporanPekerjaan(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employee.id"), nullable=False)
    tanggal = db.Column(db.Date, nullable=False)
    isi_laporan = db.Column(db.Text, nullable=False)
    lampiran_filename = db.Column(db.String(256))
    lampiran_nama_asli = db.Column(db.String(256))
    tanggal_dibuat = db.Column(db.DateTime, default=now_wib)

    employee = db.relationship("Employee", backref="laporan_pekerjaan")


class Produk(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nama_produk = db.Column(db.String(128), nullable=False)
    modal = db.Column(db.Integer, nullable=False, default=0)
    hpp = db.Column(db.Integer, nullable=False, default=0)
    harga_dasar = db.Column(db.Integer, default=0)
    harga_normal = db.Column(db.Integer, default=0)
    harga_flash_sale = db.Column(db.Integer, default=0)
    harga_big_campaign = db.Column(db.Integer, default=0)
    dibuat_pada = db.Column(db.DateTime, default=now_wib)
    diperbarui_pada = db.Column(db.DateTime, default=now_wib, onupdate=now_wib)


class BahanBaku(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nama_bahan = db.Column(db.String(128), nullable=False)
    satuan = db.Column(db.String(16), nullable=False, default="Yard")
    stok_saat_ini = db.Column(db.Float, nullable=False, default=0)
    harga_per_yard = db.Column(db.Integer, default=0)  # Rp, dari input Masuk terakhir
    warna = db.Column(db.String(64))  # dari input Masuk terakhir
    tinggi_meter = db.Column(db.Float)  # lebar/tinggi kain (meter), dari input Masuk terakhir
    suplier = db.Column(db.String(128))  # suplier langganan/terakhir
    catatan = db.Column(db.String(256))
    dibuat_pada = db.Column(db.DateTime, default=now_wib)
    diperbarui_pada = db.Column(db.DateTime, default=now_wib, onupdate=now_wib)

    kebutuhan_list = db.relationship(
        "BahanBakuKebutuhan", backref="bahan_baku", cascade="all, delete-orphan",
        order_by="BahanBakuKebutuhan.id",
    )
    transaksi_list = db.relationship(
        "BahanBakuTransaksi", backref="bahan_baku", cascade="all, delete-orphan",
        order_by="BahanBakuTransaksi.tanggal.desc()",
    )


class BahanBakuKebutuhan(db.Model):
    """Berapa yard bahan baku ini dibutuhkan untuk membuat 1 unit produk tertentu."""
    id = db.Column(db.Integer, primary_key=True)
    bahan_baku_id = db.Column(db.Integer, db.ForeignKey("bahan_baku.id"), nullable=False)
    produk_id = db.Column(db.Integer, db.ForeignKey("produk.id"), nullable=False)
    jumlah_yard = db.Column(db.Float, nullable=False, default=0)

    produk = db.relationship("Produk")

    __table_args__ = (
        db.UniqueConstraint("bahan_baku_id", "produk_id", name="uq_bahan_produk"),
    )


class BahanBakuTransaksi(db.Model):
    """Kartu stok: pencatatan bahan baku masuk (pembelian) & keluar (terpakai produksi/cutting)."""
    id = db.Column(db.Integer, primary_key=True)
    bahan_baku_id = db.Column(db.Integer, db.ForeignKey("bahan_baku.id"), nullable=False)
    tanggal = db.Column(db.Date, nullable=False)
    jenis = db.Column(db.String(8), nullable=False)  # Masuk / Keluar
    jumlah_yard = db.Column(db.Float, nullable=False)
    produk_id = db.Column(db.Integer, db.ForeignKey("produk.id"))  # tujuan/peruntukan produk
    keterangan = db.Column(db.String(256))
    dibuat_pada = db.Column(db.DateTime, default=now_wib)

    # -- khusus Masuk (halaman Input Barang) --
    warna = db.Column(db.String(64))
    suplier = db.Column(db.String(128))
    tinggi_meter = db.Column(db.Float)  # lebar/tinggi kain (meter)
    harga_per_yard = db.Column(db.Integer)
    total_dibayar = db.Column(db.Integer)  # jumlah_yard x harga_per_yard batch ini

    # -- khusus Keluar (halaman Cutting & Produksi) --
    vendor = db.Column(db.String(128))  # vendor/penjahit yg mengerjakan cutting
    lebar_kain = db.Column(db.Float)  # lebar kain dalam meter yg dipakai batch ini
    produk_jadi_pcs = db.Column(db.Integer)  # realisasi/estimasi jumlah produk jadi

    # -- kalau baris ini dibuat otomatis dari sebuah Purchase Order --
    purchase_order_id = db.Column(db.Integer, db.ForeignKey("purchase_order.id"))

    produk = db.relationship("Produk")
    purchase_order = db.relationship("PurchaseOrder")


class ProdukSpekUkuran(db.Model):
    """Spek ukuran badan per size utk 1 pcs produk -- ditampilkan sbg acuan cutting
    di halaman Barang Cutting & Produksi saat produk tsb dipilih.

    5 kolom ukuran (lingkar_dada s/d pergelangan) dipakai sbg "slot" generik yg
    label & maknanya beda-beda tergantung `kategori` (lihat KATEGORI_SPEK_FIELDS
    di app.py) -- misal utk kategori Cadar, kolom lingkar_dada dipakai buat nyimpen
    "Lingkar Kepala", bukan lingkar dada beneran. kategori kosong (baris lama sblm
    fitur ini ada) fallback ke label generik LD/PA/Lingkar Pinggang/LD Lengan/Pergelangan."""
    id = db.Column(db.Integer, primary_key=True)
    produk_id = db.Column(db.Integer, db.ForeignKey("produk.id"), nullable=False)
    kategori = db.Column(db.String(24))  # Cadar/Khimar/Pashmina/Handshock/Celamis/Abaya/Oneset/Gamis
    size = db.Column(db.String(32), nullable=False, default="All Size")
    lingkar_dada = db.Column(db.Float)
    panjang_atas = db.Column(db.Float)
    lingkar_pinggang = db.Column(db.Float)
    ld_lengan = db.Column(db.Float)
    pergelangan = db.Column(db.Float)

    produk = db.relationship("Produk", backref=db.backref("spek_ukuran_list", cascade="all, delete-orphan"))


class Vendor(db.Model):
    """Master data vendor -- mencakup penjahit/produksi (dipakai di Cutting &
    Produksi) maupun suplier kain (dipakai di Pengadaan Barang), jadi satu daftar."""
    id = db.Column(db.Integer, primary_key=True)
    nama_vendor = db.Column(db.String(128), nullable=False)
    jenis = db.Column(db.String(24), nullable=False, default="Keduanya")  # Penjahit / Suplier Kain / Keduanya
    kontak = db.Column(db.String(64))
    alamat = db.Column(db.String(256))
    catatan = db.Column(db.String(256))
    dibuat_pada = db.Column(db.DateTime, default=now_wib)


class Gudang(db.Model):
    """Master data alamat/lokasi gudang -- referensi saja, belum dipakai buat
    memisah stok per lokasi."""
    id = db.Column(db.Integer, primary_key=True)
    nama_gudang = db.Column(db.String(128), nullable=False)
    alamat = db.Column(db.String(256))
    pic = db.Column(db.String(64))
    kontak = db.Column(db.String(64))
    catatan = db.Column(db.String(256))
    dibuat_pada = db.Column(db.DateTime, default=now_wib)


class AkunPembayaran(db.Model):
    """Master data rekening bank/e-wallet/kas tunai milik toko -- referensi saja,
    belum dipakai/dikaitkan ke transaksi manapun."""
    id = db.Column(db.Integer, primary_key=True)
    nama_akun = db.Column(db.String(128), nullable=False)
    jenis = db.Column(db.String(24), nullable=False, default="Bank")  # Bank / E-wallet / Tunai
    nomor_rekening = db.Column(db.String(64))
    catatan = db.Column(db.String(256))
    dibuat_pada = db.Column(db.DateTime, default=now_wib)


class PurchaseOrder(db.Model):
    """PO produksi ke vendor/penjahit -- satu PO bisa punya banyak Item Produk (yg
    dipesan/diproduksi) & banyak Pemakaian Bahan (yg dipotong dari stok saat disimpan)."""
    id = db.Column(db.Integer, primary_key=True)
    nomor_po = db.Column(db.String(64), nullable=False, unique=True)
    vendor_id = db.Column(db.Integer, db.ForeignKey("vendor.id"), nullable=False)
    tanggal_order = db.Column(db.Date, nullable=False)
    estimasi_selesai = db.Column(db.Date)
    total_biaya = db.Column(db.Integer, nullable=False, default=0)
    # Menunggu Produksi -> Diproses -> Selesai Produksi, atau Dibatalkan kapan saja
    status = db.Column(db.String(24), nullable=False, default="Menunggu Produksi")
    dibuat_pada = db.Column(db.DateTime, default=now_wib)

    vendor = db.relationship("Vendor")
    item_produk_list = db.relationship(
        "PurchaseOrderItemProduk", backref="po", cascade="all, delete-orphan",
        order_by="PurchaseOrderItemProduk.id",
    )
    bahan_pakai_list = db.relationship(
        "PurchaseOrderBahanPakai", backref="po", cascade="all, delete-orphan",
        order_by="PurchaseOrderBahanPakai.id",
    )
    pembayaran_list = db.relationship(
        "PurchaseOrderPembayaran", backref="po", cascade="all, delete-orphan",
        order_by="PurchaseOrderPembayaran.tanggal.desc()",
    )

    @property
    def total_dibayar(self):
        return sum(p.jumlah for p in self.pembayaran_list)

    @property
    def sisa_bayar(self):
        return max(self.total_biaya - self.total_dibayar, 0)

    @property
    def status_pembayaran(self):
        if self.total_dibayar <= 0:
            return "Belum Lunas"
        if self.total_dibayar < self.total_biaya:
            return "Cicilan"
        return "Lunas"

    @property
    def total_qty_item(self):
        return sum(ip.qty for ip in self.item_produk_list)

    @property
    def total_jahit_selesai(self):
        return sum(ip.jahit_selesai for ip in self.item_produk_list)

    @property
    def total_finish_selesai(self):
        return sum(ip.finish_selesai for ip in self.item_produk_list)

    @property
    def persen_progress(self):
        """Global progress % -- dihitung dari rata2 Jahit & Finish thd total qty item."""
        total = self.total_qty_item
        if not total:
            return 0
        capaian = (self.total_jahit_selesai + self.total_finish_selesai) / 2
        return round(min(capaian, total) / total * 100)

    @property
    def status_produksi(self):
        """Status produksi (Menunggu/Diproses/Selesai) OTOMATIS ngikutin progress Jahit &
        Finish per item -- admin cuma bisa nge-override manual ke 'Dibatalkan' (field
        `status` dipakai khusus utk itu, gak dipakai lagi utk 3 status lainnya)."""
        if self.status == "Dibatalkan":
            return "Dibatalkan"
        total = self.total_qty_item
        if not total:
            return "Menunggu Produksi"
        if self.total_finish_selesai >= total:
            return "Selesai Produksi"
        if self.total_jahit_selesai > 0 or self.total_finish_selesai > 0:
            return "Diproses"
        return "Menunggu Produksi"


class PurchaseOrderItemProduk(db.Model):
    """Baris produk yg dipesan/diproduksi dalam 1 Purchase Order."""
    id = db.Column(db.Integer, primary_key=True)
    purchase_order_id = db.Column(db.Integer, db.ForeignKey("purchase_order.id"), nullable=False)
    produk_id = db.Column(db.Integer, db.ForeignKey("produk.id"), nullable=False)
    warna = db.Column(db.String(64))
    size = db.Column(db.String(32))
    qty = db.Column(db.Integer, nullable=False, default=0)
    total = db.Column(db.Integer, nullable=False, default=0)  # snapshot qty x modal produk saat itu
    jahit_selesai = db.Column(db.Integer, nullable=False, default=0)  # progress: sdh dijahit sebanyak ini
    finish_selesai = db.Column(db.Integer, nullable=False, default=0)  # progress: sdh finishing sebanyak ini

    produk = db.relationship("Produk")


class PurchaseOrderBahanPakai(db.Model):
    """Baris bahan baku yg dipakai (dikurangi dari stok) dalam 1 Purchase Order."""
    id = db.Column(db.Integer, primary_key=True)
    purchase_order_id = db.Column(db.Integer, db.ForeignKey("purchase_order.id"), nullable=False)
    bahan_baku_id = db.Column(db.Integer, db.ForeignKey("bahan_baku.id"), nullable=False)
    qty_pakai = db.Column(db.Float, nullable=False, default=0)
    satuan = db.Column(db.String(16))  # snapshot satuan bahan saat itu

    bahan_baku = db.relationship("BahanBaku")


class PurchaseOrderPembayaran(db.Model):
    """Catatan pembayaran (bisa dicicil) ke vendor utk 1 Purchase Order -- status
    Belum Lunas/Cicilan/Lunas di PurchaseOrder dihitung otomatis dari total baris ini."""
    id = db.Column(db.Integer, primary_key=True)
    purchase_order_id = db.Column(db.Integer, db.ForeignKey("purchase_order.id"), nullable=False)
    tanggal = db.Column(db.Date, nullable=False)
    jumlah = db.Column(db.Integer, nullable=False, default=0)
    # Kasbon = uang muka/DP ke vendor sblm produksi kelar, Cicilan = bayar bertahap,
    # Pelunasan = pembayaran yg nutup sisa tagihan
    metode = db.Column(db.String(16), nullable=False, default="Cicilan")
    akun_pembayaran_id = db.Column(db.Integer, db.ForeignKey("akun_pembayaran.id"))
    catatan = db.Column(db.String(256))
    dibuat_pada = db.Column(db.DateTime, default=now_wib)

    akun_pembayaran = db.relationship("AkunPembayaran")


class IklanMarketplace(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    marketplace = db.Column(db.String(32), nullable=False)  # Shopee/Tokopedia/TikTok Shop/Lazada/Blibli
    tanggal = db.Column(db.Date, nullable=False)
    biaya = db.Column(db.Integer, default=0)
    impresi = db.Column(db.Integer, default=0)
    klik = db.Column(db.Integer, default=0)
    pesanan = db.Column(db.Integer, default=0)
    omzet = db.Column(db.Integer, default=0)
    sumber_file = db.Column(db.String(256))
    dibuat_pada = db.Column(db.DateTime, default=now_wib)

    __table_args__ = (
        db.UniqueConstraint("marketplace", "tanggal", name="uq_marketplace_tanggal"),
    )


class IklanMeta(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    tanggal = db.Column(db.Date, nullable=False, unique=True)
    biaya = db.Column(db.Integer, default=0)
    pajak = db.Column(db.Integer, default=0)
    impresi = db.Column(db.Integer, default=0)
    klik = db.Column(db.Integer, default=0)
    pesanan = db.Column(db.Integer, default=0)
    omzet = db.Column(db.Integer, default=0)
    sumber_file = db.Column(db.String(256))
    dibuat_pada = db.Column(db.DateTime, default=now_wib)


class ProdukIklan(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    marketplace = db.Column(db.String(32), nullable=False)
    nama_produk = db.Column(db.String(256), nullable=False)
    tanggal = db.Column(db.Date, nullable=False)
    biaya = db.Column(db.Integer, default=0)
    impresi = db.Column(db.Integer, default=0)
    klik = db.Column(db.Integer, default=0)
    pesanan = db.Column(db.Integer, default=0)
    omzet = db.Column(db.Integer, default=0)
    sumber_file = db.Column(db.String(256))
    dibuat_pada = db.Column(db.DateTime, default=now_wib)

    __table_args__ = (
        db.UniqueConstraint("marketplace", "nama_produk", "tanggal", name="uq_marketplace_produk_tanggal"),
    )


class ItemLabaRugi(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    bulan = db.Column(db.Integer, nullable=False)
    tahun = db.Column(db.Integer, nullable=False)
    kelompok = db.Column(db.String(32), nullable=False)
    # Pendapatan / Beban Pokok Penjualan / Pendapatan Non Operasional / Beban Non Operasional
    deskripsi = db.Column(db.String(128), nullable=False)
    jumlah = db.Column(db.Integer, nullable=False, default=0)
    dibuat_pada = db.Column(db.DateTime, default=now_wib)


class PengeluaranOperasional(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    tanggal = db.Column(db.Date, nullable=False)
    kategori = db.Column(db.String(64), nullable=False)
    keterangan = db.Column(db.String(256))
    jumlah = db.Column(db.Integer, nullable=False, default=0)
    dibuat_pada = db.Column(db.DateTime, default=now_wib)


class PenjualanMarketplace(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    marketplace = db.Column(db.String(32), nullable=False)
    tanggal = db.Column(db.Date, nullable=False)
    jumlah_pesanan = db.Column(db.Integer, default=0)
    total_penjualan = db.Column(db.Integer, default=0)
    total_diskon = db.Column(db.Integer, default=0)
    sumber_file = db.Column(db.String(256))
    dibuat_pada = db.Column(db.DateTime, default=now_wib)

    __table_args__ = (
        db.UniqueConstraint("marketplace", "tanggal", name="uq_penjualan_marketplace_tanggal"),
    )


class PesananMarketplace(db.Model):
    """Baris per item produk dari laporan Order marketplace (Shopee dkk)."""
    id = db.Column(db.Integer, primary_key=True)
    marketplace = db.Column(db.String(32), nullable=False)
    no_pesanan = db.Column(db.String(64), nullable=False)
    tanggal_pesanan = db.Column(db.Date, nullable=False)
    status_pesanan = db.Column(db.String(32), default="")
    nama_produk = db.Column(db.String(256), default="")
    sku = db.Column(db.String(128), default="")
    jumlah = db.Column(db.Integer, default=0)
    subtotal = db.Column(db.Integer, default=0)  # nilai jual baris ini (sebelum dipotong biaya platform)
    sumber_file = db.Column(db.String(256))
    dibuat_pada = db.Column(db.DateTime, default=now_wib)

    __table_args__ = (
        db.UniqueConstraint("marketplace", "no_pesanan", "sku", "nama_produk", name="uq_pesanan_item"),
    )


class PendapatanPesanan(db.Model):
    """Ringkasan dana yang benar-benar diterima per pesanan, dari laporan Income marketplace."""
    id = db.Column(db.Integer, primary_key=True)
    marketplace = db.Column(db.String(32), nullable=False)
    no_pesanan = db.Column(db.String(64), nullable=False)
    tanggal_dana_dilepas = db.Column(db.Date)
    total_penghasilan = db.Column(db.Integer, default=0)  # dana bersih diterima, sudah dipotong semua biaya
    biaya_admin = db.Column(db.Integer, default=0)
    biaya_layanan = db.Column(db.Integer, default=0)
    biaya_lainnya = db.Column(db.Integer, default=0)  # gabungan biaya transaksi/kampanye/komisi ads/proses/dll
    sumber_file = db.Column(db.String(256))
    dibuat_pada = db.Column(db.DateTime, default=now_wib)

    __table_args__ = (
        db.UniqueConstraint("marketplace", "no_pesanan", name="uq_pendapatan_pesanan"),
    )


class PengajuanLembur(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey("employee.id"), nullable=False)
    tanggal = db.Column(db.Date, nullable=False)
    jam_mulai = db.Column(db.String(5), nullable=False)
    jam_selesai = db.Column(db.String(5), nullable=False)
    alasan = db.Column(db.String(512))
    status = db.Column(db.String(16), default="Menunggu")  # Menunggu / Disetujui / Ditolak
    catatan_admin = db.Column(db.String(256))
    tanggal_diajukan = db.Column(db.DateTime, default=now_wib)
    tanggal_diproses = db.Column(db.DateTime)

    employee = db.relationship("Employee", backref="pengajuan_lembur")
