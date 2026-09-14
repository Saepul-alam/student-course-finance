# 📚 LMS & Buku Tabungan - Google Apps Script

Sistem terintegrasi untuk **Lembaga Kursus/les** dan **Buku Tabungan Murid** menggunakan Google Apps Script dan Google Sheets.

## ✨ Fitur

### 🏫 Manajemen LMS
- ✅ **Manajemen Murid** - Tambah, edit, daftar murid lengkap
- ✅ **Manajemen Kelas** - Buat kelas, tentukan guru, kapasitas, biaya
- ✅ **Absensi** - Catat kehadiran harian (Hadir, Sakit, Izin, Alpha)
- ✅ **Progres** - Input nilai, topik, dan perkembangan murid

### 💰 Buku Tabungan
- ✅ **Rekening Otomatis** - Setiap murid dapat ID Tabungan unik
- ✅ **Transaksi Setoran** - Tambah saldo (tabungan)
- ✅ **Transaksi Penarikan** - Kurangi saldo dengan validasi
- ✅ **Saldo Otomatis** - Update real-time setelah setiap transaksi
- ✅ **Riwayat Lengkap** - Tampilkan semua transaksi per murid

### 🚀 Fitur Lanjutan (Web App)
- ✅ **Web App Responsif** - Buka via URL, mobile-friendly
- ✅ **Export Laporan PDF** - Absensi, Tabungan, Per Murid, Per Kelas
- ✅ **Notifikasi Email** - Kirim info absensi & transaksi ke email orang tua
- ✅ **Notifikasi WhatsApp** - Via Fonnte API, format pesan rapi + emoji
- ✅ **Login Gmail** - Akses terbatas via sheet **Users** + seed `ALLOWED_USERS`
- ✅ **Manajemen User** - Tambah/edit/hapus user login dari web app (khusus Admin)
- ✅ **Laporan Otomatis Bulanan** - PDF terkirim otomatis tiap tanggal 1 + arsip Drive
- ✅ **Grafik Tren Saldo** - Line chart 7/30/90 hari di dashboard (SVG, tanpa library)

### 🔗 Integrasi
- Semua data terhubung melalui **ID Murid**
- Dashboard lengkap dengan ringkasan
- Warna sheet yang berbeda untuk setiap modul

---

## 📁 File yang Ada

| File | Keterangan |
|------|------------|
| `Code.gs` | Backend logic semua fitur |
| `index.html` | UI Web App lengkap + halaman login dalam satu file (responsif, via URL) |
| `README.md` | Dokumentasi ini |

> 📝 **Catatan:** Project ini tidak menyertakan file `appsscript.json`. Apps Script
> membuatnya otomatis. Setelah paste kode, cukup set **timezone ke Jakarta** lewat
> Project Settings (lihat Langkah 7).

---

## 📖 Panduan Instalasi Lengkap

> ⏱️ Estimasi waktu: **10–15 menit**. Ikuti langkah berurutan, jangan ada yang dilewati.

### 📌 Bagian 1 — Siapkan Spreadsheet & Script Editor

**Langkah 1: Buat Spreadsheet**
1. Buka [sheets.google.com](https://sheets.google.com) dan login dengan akun Google
2. Klik **+ Blank** (spreadsheet kosong)
3. Beri nama, misal: `LMS & Buku Tabungan`

**Langkah 2: Buka Apps Script Editor**
1. Di spreadsheet, klik menu **Extensions** → **Apps Script**
2. Tab baru terbuka berisi editor kode
3. Beri nama project, misal: `LMS & Tabungan System` (kotak nama di kiri atas)

### 📌 Bagian 2 — Copy Semua File Kode

**Langkah 3: Copy `Code.gs`**
1. Di editor, file `Code.gs` sudah ada secara default
2. Hapus semua isi default-nya (`function myFunction() {}`)
3. Copy **seluruh isi** file `Code.gs` dari project ini → paste ke sana
4. Tekan `Ctrl+S` (Windows) / `Cmd+S` (Mac) untuk simpan

**Langkah 4: Buat File `index.html`**
1. Klik **+** (Add file) di sisi kiri editor → pilih **HTML**
2. Beri nama persis: `index` (Apps Script otomatis menambahkan `.html`)
3. Hapus isi default, paste **seluruh isi** file `index.html` dari project ini

> ⚠️ **Penting:** Cukup **satu file HTML** bernama persis `index` (huruf kecil, tanpa
> spasi) karena dipanggil dari kode. Halaman login sudah otomatis menyatu di dalam
> file ini — tampil sebagai overlay jika email belum terdaftar di `ALLOWED_USERS`.

**Langkah 5: Set Timezone Jakarta (Pengganti appsscript.json)**
1. Di editor, klik ikon **⚙️ Project Settings** (roda gigi, sidebar kiri)
2. Scroll ke bagian **General settings**
3. Pada dropdown **Time zone**, pilih **(GMT+07:00) Jakarta**
4. Selesai — tidak perlu edit file manifest

> 📄 Tidak ada `appsscript.json` di project ini — Apps Script membuat dan mengelola
> manifest-nya sendiri. Daftar izin (OAuth scopes) juga **dideteksi otomatis** dari
> API yang dipakai kode (`SpreadsheetApp`, `MailApp`, `DriveApp`, dll) saat pertama
> kali otorisasi.
>
> ⏰ **Kenapa timezone penting:** trigger laporan otomatis ("setiap tanggal 1 jam
> 08:00") mengikuti timezone project. Kalau tidak di-set, jam eksekusi bisa tidak
> sesuai waktu Indonesia.

### 📌 Bagian 3 — Konfigurasi Awal

**Langkah 6: Daftarkan Email Admin**
1. Buka `Code.gs`, cari baris berikut (ada di bagian `AUTHENTICATION`):
   ```javascript
   const ALLOWED_USERS = ['email@sekolah.com', 'admin@sekolah.com'];
   ```
2. Ganti dengan email Gmail yang boleh akses web app:
   ```javascript
   const ALLOWED_USERS = ['budi@gmail.com', 'guru.matematika@gmail.com'];
   ```
3. Simpan

> 💡 `ALLOWED_USERS` hanya **seed awal** — saat `setupSheets` dijalankan, email di dalamnya
> otomatis masuk sheet **Users** berperan **Admin**. Setelah itu, kelola user (tambah guru,
> nonaktifkan, ganti peran) langsung dari web app → menu **👥 Users**, tanpa edit kode lagi.
> Jika sheet Users kosong, sistem fallback ke `ALLOWED_USERS` agar tidak pernah terkunci.

### 📌 Bagian 4 — Jalankan Setup & Beri Izin

**Langkah 7: Run Pertama Kali (Otorisasi)**
1. Di editor, pilih fungsi `setupSheets` dari dropdown (sebelah tombol ▶️ **Run**)
2. Klik **▶️ Run**
3. Popup **Authorization required** muncul → klik **Review permissions**
4. Pilih akun Google kamu
5. Muncul peringatan **Google hasn't verified this app** → klik **Advanced** → **Go to LMS & Tabungan System (unsafe)**
   > ⚠️ Ini normal! Peringatan muncul karena script buatan sendiri belum diverifikasi Google.
   > Script ini hanya mengakses spreadsheet milikmu sendiri.
6. Centang semua izin yang diminta → klik **Allow**
7. Run selesai → cek spreadsheet: **8 sheet** baru otomatis dibuat

**Langkah 8: Verifikasi di Spreadsheet**
1. Kembali ke tab spreadsheet, refresh halaman (F5)
2. Menu **📚 LMS & Tabungan** muncul di toolbar (paling kanan)
3. Sheet yang dibuat otomatis: `Murid`, `Kelas`, `Absensi`, `Progres`, `Tabungan`, `Transaksi`, `Settings`, `Users`
   > Sheet **Users** otomatis terisi email dari `ALLOWED_USERS` sebagai Admin — ini pintu masuk pertamamu.

### 📌 Bagian 5 — Deploy Web App (Akses via URL/HP)

**Langkah 9: Deploy**
1. Di Apps Script editor, klik kanan atas **Deploy** → **New deployment**
2. Klik ikon ⚙️ di sebelah "Select type" → pilih **Web app**
3. Isi konfigurasi:
   | Field | Nilai |
   |-------|-------|
   | Description | `LMS & Tabungan v1` |
   | Execute as | **Me** (akun kamu) |
   | Who has access | **Anyone** atau **Anyone with Google account** |
4. Klik **Deploy**
5. Copy **Web app URL** yang muncul (format: `https://script.google.com/macros/s/AKfycb.../exec`)

> 🔐 Meski akses "Anyone", hanya email berstatus **Aktif** di sheet **Users** yang bisa login ke aplikasi.

**Langkah 10: Buka Web App**
1. Paste URL di browser (atau kirim ke HP)
2. Bila email Google kamu terdaftar di sheet **Users** (status Aktif), dashboard langsung tampil.
   Jika belum, layar login muncul lebih dulu — masukkan email yang terdaftar.
3. Selesai! 🎉

> 🔄 **Penting untuk update kode di masa depan:** setiap kali mengubah kode,
> lakukan **Deploy → Manage deployments → ✏️ Edit → Version: New version → Deploy**.
> Kalau tidak, URL lama masih menjalankan kode versi lama.

### ✅ Checklist Instalasi

- [ ] Spreadsheet baru dibuat
- [ ] `Code.gs` di-paste
- [ ] File HTML dibuat: `index` (login sudah menyatu di dalamnya)
- [ ] `ALLOWED_USERS` diganti dengan email kamu (seed Admin di sheet Users)
- [ ] Timezone di-set ke **(GMT+07:00) Jakarta** di Project Settings
- [ ] `setupSheets` di-Run + otorisasi berhasil
- [ ] Menu **LMS & Tabungan** muncul di spreadsheet
- [ ] Web app di-deploy & URL bisa dibuka

### ⚙️ Konfigurasi Opsional (Setelah Instalasi)

| Fitur | Cara Aktifkan |
|-------|---------------|
| 💬 **Notifikasi WhatsApp** | Web app → menu **WhatsApp** → daftar [fonnte.com](https://fonnte.com), tempel token, aktifkan, test kirim |
| ⏰ **Laporan Otomatis Bulanan** | Web app → menu **Laporan** → bagian bawah → isi email penerima → **Simpan & Aktifkan** |
| 📧 **Notifikasi Email Orang Tua** | Otomatis tersedia — cukup isi kolom **Email Orang Tua** di data murid, lalu centang notifikasi saat input |

---

## 📋 Sheet yang Dibuat

| Sheet | Fungsi |
|-------|--------|
| **Murid** | Data murid + ID Tabungan |
| **Kelas** | Daftar kelas dan guru |
| **Absensi** | Riwayat kehadiran |
| **Progres** | Nilai dan topik pembelajaran |
| **Tabungan** | Saldo per murid |
| **Transaksi** | Riwayat setoran & penarikan |
| **Settings** | Konfigurasi sekolah + email penerima laporan |
| **Log Laporan** | Riwayat eksekusi laporan otomatis (dibuat otomatis) |

---

## 📊 Cara Menggunakan

### Menu Utama
Setelah setup, klik menu **LMS & Tabungan** di toolbar spreadsheet.

### Tambah Murid
1. Pilih **Dashboard** → **Tambah Murid**
2. Isi form nama, kelas, dll
3. Murid otomatis dapat rekening tabungan!

### Absensi
1. Pilih **Absensi Hari Ini**
2. Pilih murid dari dropdown
3. Pilih status (Hadir/Sakit/Izin/Alpha)
4. Klik **Simpan**

### Transaksi Tabungan
1. Pilih **T Tambah Transaksi**
2. Pilih murid → ID tabungan muncul otomatis
3. Pilih Setoran atau Penarikan
4. Masukkan jumlah
5. Lihat saldo saat ini di card info
6. Klik **Proses Transaksi**

---

## 🔐 Hak Akses & Izin

Saat pertama kali Run, Google otomatis mendeteksi dan meminta izin berikut
(sesuai API yang dipakai kode):

| Izin | Dipakai Untuk |
|------|---------------|
| `spreadsheets` | Baca/tulis data murid, absensi, tabungan |
| `script.send_mail` | Kirim notifikasi email & laporan bulanan |
| `drive` | Arsip PDF laporan ke folder Drive |
| `documents` | Generate file PDF laporan |
| `script.scriptapp` | Trigger laporan otomatis tiap tanggal 1 |
| `script.external_request` | Kirim WhatsApp via Fonnte API |

Data yang diakses hanyalah spreadsheet milikmu sendiri. Token WhatsApp disimpan di
**Script Properties** (bukan di kode), jadi tidak tersebar di source code.

---

## 💡 Tips

1. **Order penting**: Buat kelas dulu sebelum tambah murid
2. **Backup**: Export spreadsheet secara berkala (File → Download → Microsoft Excel)
3. **Integrasi**: Gunakan ID Murid sebagai referensi utama
4. **Mobile friendly**: Web app bisa dibuka dari browser HP dengan URL deploy
5. **Update kode**: Setelah edit kode, selalu **Deploy → New version** agar web app pakai kode terbaru

---

## ⏰ Laporan Otomatis Bulanan

Sistem bisa mengirim **PDF laporan bulanan** (Tabungan + Absensi) secara otomatis:

1. Buka web app → menu **Laporan** → bagian **Laporan Otomatis Bulanan**
2. Isi **Email Penerima** (bisa lebih dari satu, pisahkan koma)
3. Klik **Simpan & Aktifkan**
4. Setiap **tanggal 1 jam 08:00**, laporan bulan sebelumnya terkirim otomatis

Bonus:
- **Test Kirim Sekarang** — kirim email test dengan data bulan berjalan
- **Arsip Drive** — semua PDF tersimpan di folder `LMS & Tabungan - Arsip Laporan`
- **Log Laporan** — sheet berisi riwayat eksekusi (berhasil/gagal)
- Bisa juga diaktifkan dari menu spreadsheet: **LMS & Tabungan → Laporan Otomatis**

## 💬 Notifikasi WhatsApp (Fonnte)

Sistem bisa mengirim notifikasi WhatsApp ke orang tua saat absensi & transaksi tabungan:

1. Buka web app → menu **💬 WhatsApp**
2. Daftar di [fonnte.com](https://fonnte.com) (gratis), salin **Token API**
3. Scan QR dengan WhatsApp yang akan jadi pengirim
4. Tempel token → centang **Aktifkan** → **Simpan**
5. Klik **Test Kirim** dengan nomor HP kamu untuk memastikan berfungsi

Saat input absensi/transaksi, centang **💬 Kirim notifikasi WhatsApp** dan pesan otomatis terkirim ke No HP orang tua (format 08xx dikonversi otomatis ke 628xx).

## 🛠 Troubleshooting

### Saat Instalasi

| Masalah | Solusi |
|---------|--------|
| Menu **LMS & Tabungan** tidak muncul | Refresh spreadsheet (F5), tunggu 10 detik. Kalau masih belum, pastikan `Code.gs` tersimpan (✓ oranye di editor berarti belum disimpan) |
| Error `Script function not found: setupSheets` | Pastikan `Code.gs` ter-paste utuh, lalu jalankan `setupSheets` dari dropdown editor |
| Peringatan "unverified app / unsafe" | Normal untuk script sendiri — klik **Advanced** → **Go to project (unsafe)** → **Allow** |
| Timezone belum di-set / laporan otomatis jam-nya salah | Buka **Project Settings → General settings → Time zone** → pilih **(GMT+07:00) Jakarta** |
| Web app tampil halaman kosong / error | Pastikan file HTML bernama persis `index` (lihat sidebar editor) |
| Web app masih kode lama setelah edit | **Deploy → Manage deployments → Edit → New version** |
| Login ditolak "email tidak terdaftar" | Buka sheet **Users** → tambah baris email tsb (Status: Aktif), atau masukkan ke `ALLOWED_USERS` lalu jalankan ulang `setupSheets` — tunggu ±2 menit (cache) |
| WhatsApp tidak terkirim | Cek token Fonnte masih aktif (QR tidak logout), dan status **Aktif** di menu WhatsApp web app |
| Laporan otomatis tidak jalan | Cek sheet **Log Laporan** untuk alasan gagal; pastikan email penerima terisi & trigger aktif |

### Pesan Error Umum

| Pesan | Arti & Solusi |
|-------|---------------|
| `Rekening tabungan tidak ditemukan!` | ID Tabungan di sheet Tabungan tidak cocok — cek kolom ID di sheet Murid & Tabungan |
| `Saldo tidak cukup!` | Penarikan melebihi saldo saat ini |
| `Absensi sudah tercatat hari ini!` | Satu murid hanya boleh 1 absensi per hari |
| `Email orang tua tidak tersedia!` | Isi kolom **Email Orang Tua** di sheet Murid dulu |
| `No HP orang tua tidak tersedia!` | Isi kolom **No HP** di sheet Murid (format 08xx) |

---

## 📄 Lisensi

Project ini gratis untuk digunakan. Silakan dimodifikasi sesuai kebutuhan.

---

## 🆘 Bantuan

Jika ada pertanyaan atau perlu fitur tambahan, buka **Bantuan** dari menu atau edit kode langsung di Apps Script editor.

---

**Made with ❤️ using Google Apps Script**
