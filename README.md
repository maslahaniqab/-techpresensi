# MN Portal

Aplikasi web sederhana untuk mencatat kehadiran karyawan dan menghitung gaji otomatis (gaji pokok, potongan telat/alpha, uang lembur).

## Fitur
- Login admin
- Manajemen data karyawan (CRUD)
- Input absensi harian (Hadir/Sakit/Izin/Cuti/Alpha) beserta jam masuk & pulang
- Perhitungan otomatis keterlambatan & lembur berdasarkan jam kerja standar
- Generate slip gaji bulanan otomatis dari data absensi
- Slip gaji bisa dicetak/print, dan ditandai "Sudah Dibayar"
- Pengaturan jam kerja, denda telat, tarif lembur, hari kerja per bulan

## Menjalankan di Komputer Sendiri (Lokal)

1. Buat virtual environment & install dependency:
   ```bash
   python -m venv venv
   venv\Scripts\activate
   pip install -r requirements.txt
   ```
2. Jalankan aplikasi:
   ```bash
   python app.py
   ```
3. Buka browser ke `http://localhost:5000`
4. Login default: **username `admin`, password `admin123`** — segera ganti password di menu "Akun Saya" setelah login pertama.

## Cara Pakai Singkat

1. **Pengaturan** → atur jam kerja standar, hari kerja/bulan, denda telat, tarif lembur.
2. **Karyawan** → tambahkan data karyawan beserta gaji pokok.
3. **Absensi** → tiap hari, klik "Input / Edit Absensi", pilih status tiap karyawan (Hadir/Sakit/Izin/Cuti/Alpha), isi jam masuk & pulang jika Hadir.
4. **Penggajian** → pilih bulan & tahun, klik "Generate Slip Gaji". Sistem otomatis menghitung gaji bersih dari rekap absensi bulan tersebut. Buka slip untuk melihat rincian atau mencetak, lalu tandai "Sudah Dibayar" setelah ditransfer.

## Deploy Online (agar bisa diakses dari HP/dari mana saja)

Aplikasi ini sudah siap di-deploy ke layanan hosting gratis/berbayar seperti **Render** atau **Railway**.

### Deploy ke Render.com (gratis, disarankan)
1. Push folder project ini ke repository GitHub.
2. Buat akun di [render.com](https://render.com), klik **New +** → **Web Service**, hubungkan ke repo GitHub Anda.
3. Isi konfigurasi:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app`
4. Tambahkan environment variable `SECRET_KEY` dengan nilai acak/rahasia.
5. **Penting (penyimpanan data):** Render's free tier menghapus file lokal (termasuk database SQLite) setiap kali di-restart/redeploy. Agar data absensi & karyawan tidak hilang, tambahkan **Render Persistent Disk** (mount ke folder project) dan set env var `DATABASE_URL` menunjuk ke file di disk tsb, atau pakai **Render PostgreSQL** (gratis untuk skala kecil) — cukup tambahkan env var `DATABASE_URL` dari database Postgres yang dibuat, aplikasi ini otomatis mendukungnya.
6. Setelah deploy selesai, Anda akan mendapat URL seperti `https://nama-app.onrender.com` yang bisa dibuka dari HP.

### Deploy ke Railway.app
1. Push ke GitHub, lalu import project di [railway.app](https://railway.app).
2. Railway otomatis mendeteksi `Procfile` dan `requirements.txt`.
3. Tambahkan Railway **PostgreSQL plugin**, lalu set environment variable `DATABASE_URL` (Railway biasanya mengisi otomatis) dan `SECRET_KEY`.
4. Deploy, lalu buka URL yang diberikan Railway.

## Struktur Proyek
```
app.py                 # Routing & logika aplikasi
models.py               # Model database (User, Employee, Attendance, Payroll, Settings)
templates/               # Halaman HTML
requirements.txt
Procfile                 # Untuk deploy ke Render/Railway
```

## Catatan Keamanan
- Segera ganti password admin default setelah login pertama kali.
- Saat deploy online, pastikan set `SECRET_KEY` yang unik lewat environment variable, jangan gunakan nilai default di kode.
