# Panduan Pemasangan Koperasi App

Tutorial membangun aplikasi koperasi ini **dari repositori**, untuk:

- **memasang ulang** koperasi yang sudah memakai aplikasi ini (misalnya pindah akun atau lupa langkahnya), atau
- **memasang untuk koperasi lain** yang pembukuannya serupa (simpan pinjam anggota, potong gaji).

Panduan pemakaian harian untuk admin dan petugas ada di [`Panduan Aplikasi Koperasi.pdf`](Panduan%20Aplikasi%20Koperasi.pdf).

---

## Daftar isi

1. [Gambaran singkat](#1-gambaran-singkat)
2. [Yang perlu disiapkan](#2-yang-perlu-disiapkan)
3. [Memasang dari nol](#3-memasang-dari-nol)
4. [Pengaturan awal di aplikasi](#4-pengaturan-awal-di-aplikasi)
5. [Mengisi data awal](#5-mengisi-data-awal)
6. [Memakai untuk koperasi lain](#6-memakai-untuk-koperasi-lain)
7. [Memperbarui kode](#7-memperbarui-kode)
8. [Pekerjaan rutin & tahunan](#8-pekerjaan-rutin--tahunan)
9. [Mengatasi masalah](#9-mengatasi-masalah)
10. [Lampiran: struktur data](#10-lampiran-struktur-data)

---

## 1. Gambaran singkat

```
Google Sheets (basis data)  ←→  Apps Script (kode di repo ini)  ←→  Web App (browser/HP)
        │                                   ▲
        │                                   │ clasp push
        └── satu spreadsheet per koperasi   └── folder repo ini di laptop
```

- **Spreadsheet** menyimpan semua data: anggota, transaksi, kas, jurnal, pengaturan, pengguna aplikasi.
- **Apps Script** adalah proyek skrip yang *terikat* (bound) ke spreadsheet itu. Kode dari repo ini dikirim ke sana dengan `clasp`.
- **Web App** adalah hasil *deploy* proyek skrip. Semua orang memakai link `/exec` dan masuk dengan **akun aplikasi** (email + sandi yang dibuat admin), bukan akun Google.

---

## 2. Yang perlu disiapkan

| Kebutuhan | Keterangan |
|---|---|
| Akun Google pemilik | Pemilik spreadsheet & skrip. Semua data & cadangan tersimpan di Drive akun ini. Sebaiknya akun lembaga, bukan akun pribadi. |
| Laptop dengan **Node.js** | Untuk menjalankan `clasp`. Unduh dari https://nodejs.org (versi LTS). |
| **Git** | Untuk mengambil kode dari GitHub. https://git-scm.com |
| **Python 3** + `openpyxl` | Hanya bila memindahkan data dari Excel pembukuan (bagian 5A). |

Pasang `clasp` (sekali saja):

```bash
npm install -g @google/clasp
clasp login
```

`clasp login` membuka browser. Masuk dengan **akun Google pemilik**. Lalu aktifkan Apps Script API di
https://script.google.com/home/usersettings → **Google Apps Script API: On**.

---

## 3. Memasang dari nol

### 3.1 Ambil kode

```bash
git clone https://github.com/asep12/koperasi-app.git
cd koperasi-app
```

Repo ini privat. Untuk koperasi lain, salin (fork) atau minta akses dulu.

### 3.2 Buat spreadsheet & proyek skrip

1. Buka https://sheets.google.com → buat **spreadsheet kosong** → beri nama, misalnya *Koperasi Sekolah*.
2. Di spreadsheet: **Ekstensi → Apps Script**. Tab proyek skrip terbuka.
3. Di Apps Script: **⚙️ Setelan proyek** (Project Settings) → salin **ID skrip** (Script ID).

### 3.3 Hubungkan folder repo ke proyek skrip

```bash
cp .clasp.json.contoh .clasp.json                    # Windows: copy .clasp.json.contoh .clasp.json
cp identitas_lokal.contoh.js identitas_lokal.js      # Windows: copy ...
```

- Buka `.clasp.json`, ganti `ISI_DENGAN_SCRIPT_ID_...` dengan ID skrip tadi.
- Buka `identitas_lokal.js`, isi identitas koperasi (nama, jenis, alamat, kota, nomor badan hukum, sekolah/instansi,
  nama folder cadangan) dan nama ketua, sekretaris, bendahara, anggota pengurus, serta pengawas.

> Kedua file ini sengaja **tidak** masuk Git (`.gitignore`), karena berisi ID proyek, identitas lembaga, dan nama orang.
> Kode di repo ini netral: layar login, cetakan, dan laporan RAT mengambil nama dari identitas tersebut.

### 3.4 Kirim kode

```bash
clasp push
```

Jika ditanya *"Manifest file has been updated. Do you want to push and overwrite?"* jawab **y**.
Yang terkirim hanya file `*.js`, `*.html`, dan `appsscript.json` (lihat `.claspignore`); file Excel, dokumen,
contoh, dan data migrasi tidak ikut.

### 3.5 Buat sheet-sheet

1. Kembali ke spreadsheet, **muat ulang** halamannya. Muncul menu **🔧 SETUP** dan **📋 KOPERASI**.
2. **🔧 SETUP → ▶️ Jalankan Setup Lengkap**.
3. Pertama kali, Google meminta izin: **Tinjau izin → pilih akun pemilik → Lanjutan → Buka (tidak aman) → Izinkan**.
   (Peringatan "tidak aman" muncul karena skrip buatan sendiri, bukan dari Google.)
4. Jalankan lagi **Setup Lengkap** bila perlu. Hasilnya: sheet `anggota`, `setting_rat`, `coa_akun`,
   `transaksi_*`, `jurnal_umum`, `saldo_awal`, `log_aktivitas`, dan lainnya. Setup aman diulang: sheet yang sudah ada tidak disentuh.

### 3.6 Buat admin pertama

Kata sandi tidak ditulis di kode. Di Apps Script:

1. **⚙️ Setelan proyek → Properti skrip → Tambahkan properti skrip**:
   - `ADMIN_EMAIL` = email admin (dipakai untuk login aplikasi)
   - `ADMIN_NAMA` = nama admin
   - `ADMIN_SANDI` = sandi awal (minimal 6 karakter)
2. **Simpan**. Buka editor, pilih file `sesi.js`, pilih fungsi **`buatAdminPertama`** → **Jalankan**.
3. Log eksekusi: *"Admin pertama dibuat"*. `ADMIN_SANDI` otomatis dihapus dari properti.

Cara yang sama dipakai bila **lupa sandi admin**: isi lagi ketiga properti, jalankan `buatAdminPertama`.

### 3.7 Deploy sebagai Web App

1. Apps Script: **Terapkan (Deploy) → Deployment baru**.
2. Jenis: **Aplikasi web**.
   - **Jalankan sebagai:** *Saya* (akun pemilik)
   - **Yang memiliki akses:** *Siapa saja* (Anyone)
3. **Terapkan** → salin **URL aplikasi web** (berakhiran `/exec`). Itulah link yang dibagikan.

> "Siapa saja" aman karena aplikasi punya login sendiri, pembatasan peran, dan penguncian setelah 5 kali salah sandi.
> Akun Google Workspace sekolah (mis. belajar.id) kadang **tidak menyediakan** pilihan "Siapa saja"; bila begitu,
> pasang di akun Gmail biasa, atau semua pengguna harus login dengan akun domain yang sama.

### 3.8 Tugas otomatis (cadangan & pengingat)

Di editor, file `log_email_backup.js` → fungsi **`pasangTugasHarian`** → **Jalankan**. Setiap hari pukul 06.00–07.00:

- tanggal 1: salinan spreadsheet ke folder Drive *Backup Koperasi …* (12 terakhir disimpan);
- tanggal 10: email pengingat angsuran ke admin;
- tanggal 25: email pengingat jasa sukarela.

---

## 4. Pengaturan awal di aplikasi

Buka link `/exec`, masuk sebagai admin, lalu:

1. **Pengaturan → Koperasi & Tahun Buku → Identitas Koperasi**: nama, jenis, alamat, kota, nomor badan hukum,
   telepon, email, logo (PNG/JPG), nama pengurus & pengawas. Semua cetakan dan laporan RAT memakai isian ini.
2. **Tahun Buku & RAT**:
   - Tahun buku aktif bawaan **2026**. Bila memulai di tahun lain, ubah angka tahun pada baris pertama sheet
     `setting_rat` langsung di spreadsheet (kolom `tahun`, `status_aktif` = TRUE).
   - Jasa pinjaman (% per bulan, menurun), **rumus jasa sukarela** (saldo terendah × %/bulan atau aturan lama
     %/tahun × 30/365), alokasi SHU 7 pos + pajak, porsi SHU simpanan/jasa, nominal simpanan pokok, wajib,
     sukarela minimal.
3. **Pengguna**: buat akun untuk bendahara & petugas. Peran **Admin** (semua fitur) atau **Operator**
   (petugas: input & cetak; tidak bisa membukukan potong gaji, memposting jasa, atau membatalkan transaksi).

---

## 5. Mengisi data awal

Pilih salah satu:

### 5A. Dari Excel pembukuan (template "PEMBUKUAN KOPERASI")

Cocok bila koperasi memakai Excel *PEMBUKUAN KOPERASI \<tahun\>.xlsx* (sheet `DATA_Anggota`, `Isim(n)`, `Ang (n)`,
`KAS (n)`, `Simp_TH_Sblmnya`, dst.). Seluruh riwayat tahun berjalan ikut dipindahkan dan dicocokkan per bulan.

```bash
pip install openpyxl
python tools/ekstrak_excel.py "referensi/PEMBUKUAN KOPERASI 2026.xlsx"
```

Opsi yang sering dipakai (lihat `python tools/ekstrak_excel.py -h`):

| Opsi | Guna |
|---|---|
| `--jasa-pinjaman 1.5` | Persen jasa pinjaman per bulan (menurun) |
| `--jasa-sukarela 1` `--metode-jasa terendah` | Aturan jasa sukarela tahun berjalan |
| `--pokok 10000 --wajib 50000` | Nominal simpanan |
| `--bulan 9` | Batasi sampai bulan tertentu |
| `--koreksi koreksi.json` | Membetulkan label pos kas yang salah di Excel |

Skrip menghitung ulang **saldo tiap anggota, piutang, dan kas setiap bulan** dan membandingkannya dengan Excel.
Ada selisih → file **tidak** dibuat dan daftar selisih ditampilkan. Tanda **⚠** adalah hal yang perlu dicek
bendahara (saldo sukarela minus, jasa tanpa pinjaman, pengeluaran dana melebihi saldonya).

Contoh `koreksi.json` (angka ilustrasi) — bila Excel menaruh pembayaran SHU anggota di baris "Dana Pengurus":

```json
[{"bulan": 1, "akun_dari": "212", "akun_ke": "211", "nominal": 50000000,
  "ket": "Pembagian SHU tahun lalu kepada anggota (dana anggota)"}]
```

`nominal` harus sama persis dengan angka di sheet KAS; skrip menolak koreksi yang tidak cocok dengan tepat satu pos.

Kode akun dana: 211 anggota · 212 pengurus · 213 kesejahteraan · 214 pendidikan · 215 sosial · 216 pembangunan daerah.

Lalu kirim data & jalankan migrasi:

1. Buka `.claspignore`, beri tanda `#` di depan baris `migrasi_data.js` (sementara) → `clasp push`.
2. Aplikasi → **Pengaturan → Koperasi & Tahun Buku → Import Data Awal → Migrasi dari Excel pembukuan →
   Jalankan migrasi** → ketik `MIGRASI`.
   Urutan otomatis: salinan cadangan → kosongkan data lama → anggota & saldo 31 Des → bulan demi bulan → status.
   Setiap langkah atomik; bila gagal, klik **Lanjutkan migrasi** setelah masalah dibereskan.
3. Setelah selesai: hapus lagi tanda `#` di `.claspignore` → `clasp push` (file data terhapus dari Apps Script).
   **Jangan** commit `migrasi_data.js` (sudah di `.gitignore`).

Hasil migrasi: anggota diberi ID `AGT-0001…` (PNS dulu, lalu Non PNS); saldo 31 Desember tahun lalu tersimpan
sebagai saldo awal; SHU tahun lalu dibagi di bulan Januari sesuai persentase RAT; tiap transaksi punya kas & jurnal.

### 5B. Manual (koperasi baru / Excel berbeda)

1. **Anggota**:
   - satu per satu di aplikasi (**Anggota → Tambah Anggota**), atau
   - tempel daftar ke sheet `anggota` (kolom `nama`, `jabatan`, `status` = `aktif`, `tanggal_masuk`),
     lalu **🔧 SETUP → Generate ID Anggota Kosong**.
2. **Saldo awal** (posisi saat mulai memakai aplikasi): **Import Data Awal → Import Saldo Awal**, satu baris per
   saldo: `id_anggota atau nama ; jenis ; nominal`. Jenis per anggota: `pokok`, `wajib`, `sukarela`, `piutang`.
   Milik koperasi (kolom pertama `-`): `kas`, `inventaris`, `penyusutan`, `cadangan`, `dana_anggota`, `dana_pengurus`,
   `dana_kesejahteraan`, `dana_pendidikan`, `dana_sosial`, `dana_pembangunan`, `shu`. Angka boleh `4.370.495,82`
   atau `4370495.82`. Aplikasi menampilkan cek keseimbangan aktiva = pasiva.
3. **Pinjaman berjalan**: **Import Pinjaman Migrasi**, satu baris per pinjaman:
   `nama ; sisa_pinjaman ; tenor_sisa ; bulan_asli ; tahun_asli ; angsuran_pokok`.
   Jumlah sisa pinjaman harus sama dengan baris `piutang` di saldo awal.

---

## 6. Memakai untuk koperasi lain

Kode repo ini tidak menyebut nama koperasi atau sekolah mana pun. Semua identitas diambil dari:

1. **`identitas_lokal.js`** (bagian 3.3): nilai bawaan saat pertama dipasang, termasuk nama folder cadangan di Drive.
2. **Pengaturan → Identitas Koperasi** di aplikasi: nama, jenis, alamat, kota, badan hukum, sekolah/instansi, logo,
   pengurus & pengawas. Isian yang disimpan di sini selalu diutamakan.

Layar login (nama & inisial koperasi), kop semua cetakan, slip potong gaji, dan laporan RAT
("Sambutan Kepala …", tempat rapat, bendahara gaji …, dinas kota …) otomatis mengikuti identitas tersebut.

Yang mungkin masih perlu disesuaikan di kode:

| File | Yang diubah |
|---|---|
| `rat.js` | Kalimat baku laporan RAT (sambutan, evaluasi program, saran pengawas, bantuan kesejahteraan) bila format RAT koperasi berbeda. Bisa juga cukup disunting di file Word hasilnya. |
| `setup.js` → `setting_rat.dataAwal` | Tahun buku awal & aturan bawaan (jasa, alokasi SHU, nominal). |
| `profil_koperasi.js` → `LOGO_KOPERASI_INDONESIA` | Logo bawaan (lambang Koperasi Indonesia); logo sendiri cukup diunggah lewat Pengaturan. |

Aturan yang **tidak perlu** mengubah kode (cukup di Pengaturan): persen jasa pinjaman & sukarela, rumus jasa
sukarela, alokasi SHU, pajak, nominal simpanan, pengguna & peran.

---

## 7. Memperbarui kode

```bash
# ubah kode …
clasp push                 # kirim ke Apps Script → langsung terlihat di link /dev (khusus pemilik)
```

1. Uji lewat link **`/dev`** (Apps Script → Terapkan → Uji deployment). Link ini hanya bisa dibuka pemilik/editor.
2. Bila sudah benar, perbarui link yang dipakai semua orang: **Terapkan → Kelola deployment → ✏️ Edit →
   Versi: Versi baru → Terapkan**. Link `/exec` tetap sama.
3. Simpan ke GitHub:

   ```bash
   git status                   # pastikan tidak ada file data (Excel, migrasi_data.js, identitas_lokal.js)
   git add -A
   git commit -m "keterangan perubahan"
   git push
   ```

---

## 8. Pekerjaan rutin & tahunan

| Kapan | Pekerjaan | Di aplikasi |
|---|---|---|
| Awal bulan | Tagihan & slip potong gaji | Bulanan → Tampilkan → Cetak |
| Setelah gaji dibayar | Bukukan potong gaji (admin) | Bulanan → Proses Potong Gaji |
| Tgl 25–31 | Hitung & posting jasa sukarela | Bulanan → Jasa simpanan sukarela |
| Akhir tahun | Cek & tutup buku (jurnal penutup, alokasi SHU, tahun baru aktif) | Pengaturan → Tutup Buku |
| RAT | Laporan RAT (Word, F4) | Keuangan → Laporan & SHU → Laporan RAT |
| Setelah RAT | Perbarui aturan tahun baru bila berubah | Pengaturan → Tahun Buku & RAT |

Rincian langkah untuk bendahara/petugas: [`Panduan Aplikasi Koperasi.pdf`](Panduan%20Aplikasi%20Koperasi.pdf).

---

## 9. Mengatasi masalah

| Gejala | Penyebab & jalan keluar |
|---|---|
| *"Sorry, unable to open the file at this time"* | Browser login lebih dari satu akun Google. Buka link di jendela **Samaran/Incognito** (Ctrl+Shift+N) atau profil Chrome berisi satu akun. |
| Link `/dev` tidak bisa dibuka orang lain | Memang khusus pemilik/editor. Bagikan link `/exec`. |
| Perubahan kode tidak terlihat di `/exec` | Belum membuat **versi baru** deployment (bagian 7). |
| Error *"Pilih: … / … / …"* saat menyimpan | Aturan dropdown lama di sheet. Di editor jalankan fungsi **`perbaruiValidasiSheet`** (file `setup.js`). |
| Lupa sandi admin | Bagian 3.6. Pengguna lain: admin → Pengaturan → Pengguna → atur ulang sandi. |
| Akun terkunci | 5 kali salah sandi → tunggu 15 menit. |
| *"Sistem sedang memproses transaksi lain"* | Dua penyimpanan bersamaan; coba lagi beberapa detik kemudian. |
| Data terhapus/rusak | Buka folder Drive *Backup Koperasi …*, buat salinan cadangan terakhir, lalu hubungkan skrip ke salinan itu (atau salin sheet yang diperlukan). |
| Menu 🔧 SETUP tidak muncul | Muat ulang spreadsheet; pastikan `clasp push` berhasil. |

---

## 10. Lampiran: struktur data

| Sheet | Isi |
|---|---|
| `anggota` | ID, NIP/NIS, nama, jabatan, alamat, telepon, tanggal masuk, status (aktif/nonaktif/keluar), sukarela rutin |
| `setting_rat` | Satu baris per tahun buku: jasa, rumus jasa sukarela, alokasi SHU, pajak, nominal, tanggal RAT, status aktif |
| `saldo_awal` | Saldo pembuka per anggota & milik koperasi |
| `transaksi_simpanan` / `transaksi_pengambilan` | Setoran & penarikan (penarikan melalui persetujuan admin) |
| `transaksi_pinjaman` / `transaksi_angsuran` | Pinjaman (induk & tambahan) dan angsuran pokok + jasa |
| `transaksi_kas` | Semua kas masuk/keluar |
| `jurnal_umum` | Jurnal berpasangan (debit = kredit) untuk neraca & rugi laba |
| `rekap_jasa_sukarela` | Jasa sukarela per anggota per bulan (DRAFT → POSTED) |
| `coa_akun` | Bagan akun (101 Kas, 103 Piutang, 2xx simpanan & dana, 3xx modal, 4xx pendapatan, 5xx biaya) |
| `profil_koperasi` | Identitas koperasi (kunci–nilai) |
| `pengguna` | Akun aplikasi, peran, sandi ter-hash |
| `log_aktivitas` | Jejak semua perubahan |

Setiap transaksi yang dibatalkan tidak dihapus: statusnya menjadi VOID dan jurnalnya dibalik, sehingga
riwayat selalu bisa ditelusuri.
