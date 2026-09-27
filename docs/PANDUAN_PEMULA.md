# Panduan Pemula: Pasang Tanpa Laptop Khusus (Salin-Tempel)

Cara paling mudah memasang Koperasi App. **Tidak perlu** Node.js, Git, `clasp`, atau perintah terminal.
Cukup browser (Chrome disarankan) dan akun Google. Bisa dikerjakan dari laptop atau komputer mana saja.

Waktu yang dibutuhkan: sekitar 20–30 menit.

> **Izin pemakaian.** Aplikasi ini dilindungi hak cipta (lihat [LICENSE](../LICENSE)). Sebelum memasang untuk
> koperasi Anda, minta izin dulu ke [@asep94 di Threads](https://www.threads.com/@asep94). Untuk kustomisasi fitur
> juga bisa menghubungi akun tersebut.

Butuh cara lengkap untuk pengembang (dengan `clasp`)? Lihat [Panduan Pemasangan](PANDUAN_PEMASANGAN.md).

---

## Yang disalin

Semua kode sudah disiapkan di folder **[`paket_manual/`](../paket_manual)**:

| File di `paket_manual/` | Ditempel ke Apps Script sebagai | Jenis file |
|---|---|---|
| `appsscript.json` | `appsscript.json` (sudah ada, isinya diganti) | Manifes |
| `Kode_1.gs` | `Kode_1` | Skrip |
| `Kode_2.gs` | `Kode_2` | Skrip |
| `Kode_3.gs` | `Kode_3` | Skrip |
| `index.html` | `index` | HTML |

> Jumlah file `Kode_N.gs` bisa bertambah atau berkurang di versi mendatang. Salin **semua** file `Kode_…` yang ada
> di folder itu.

### Cara menyalin satu file dari GitHub

1. Buka folder [`paket_manual/`](../paket_manual), klik nama file (misalnya `Kode_1.gs`).
2. Di kanan atas isi file, klik tombol **Copy raw file** (ikon dua kotak bertumpuk). Seluruh isi file tersalin.
3. Di Apps Script, klik di dalam file tujuan, tekan **Ctrl+A** (pilih semua), lalu **Ctrl+V** (tempel).

Jangan menyalin dengan menyeret/memblok teks di halaman GitHub: file ini panjang, dan bagian yang terlewat
membuat aplikasi error.

---

## Langkah 1: Buat spreadsheet & buka Apps Script

1. Buka https://sheets.google.com dengan **akun Google pemilik** (sebaiknya akun koperasi/lembaga, bukan pribadi).
2. Buat **Spreadsheet kosong**, beri nama, misalnya *Koperasi Sekolah*.
3. Menu **Ekstensi → Apps Script**. Tab baru terbuka berisi editor skrip.
4. Klik judul *Proyek tanpa judul* di kiri atas, ganti misalnya menjadi *Koperasi App*.

## Langkah 2: Tempel manifes (`appsscript.json`)

1. Di editor, klik **⚙️ Setelan proyek** (ikon roda gigi di kiri).
2. Centang **Tampilkan file manifes "appsscript.json" di editor**.
3. Kembali ke **< > Editor**. Klik file **`appsscript.json`**.
4. Salin isi [`paket_manual/appsscript.json`](../paket_manual/appsscript.json), tempel menggantikan seluruh isinya.
5. Tekan **Ctrl+S** untuk menyimpan.

## Langkah 3: Tempel kode (`Kode_1`, `Kode_2`, `Kode_3`)

1. Di daftar file sudah ada **`Kode.gs`** (atau `Code.gs`). Klik titik tiga di sebelahnya → **Ganti nama** →
   ketik `Kode_1`.
2. Buka `Kode_1`, hapus seluruh isinya, tempel isi [`paket_manual/Kode_1.gs`](../paket_manual/Kode_1.gs).
   Tekan **Ctrl+S**.
3. Klik **＋** (di samping tulisan *File*) → **Skrip** → ketik nama `Kode_2` → Enter.
   Hapus isi bawaannya (`function myFunction() {}`), tempel isi `paket_manual/Kode_2.gs`, **Ctrl+S**.
4. Ulangi untuk `Kode_3` (dan `Kode_4` dst. bila ada).

## Langkah 4: Tempel tampilan (`index`)

1. Klik **＋** → **HTML** → ketik nama **`index`** (huruf kecil semua, tanpa `.html`; Apps Script menambahkannya
   sendiri) → Enter.
2. Hapus isi bawaannya, tempel isi [`paket_manual/index.html`](../paket_manual/index.html), **Ctrl+S**.

> Nama harus persis `index`. Bila namanya lain, aplikasi web menampilkan error saat dibuka.

Setelah langkah ini, daftar file berisi: `appsscript.json`, `Kode_1.gs`, `Kode_2.gs`, `Kode_3.gs`, `index.html`.

## Langkah 5: Buat sheet-sheet basis data

1. Kembali ke tab **spreadsheet**, tekan **F5** (muat ulang). Tunggu beberapa detik sampai muncul menu
   **🔧 SETUP** dan **📋 KOPERASI** di bagian atas.
2. Klik **🔧 SETUP → ▶️ Jalankan Setup Lengkap**.
3. Pertama kali, Google meminta izin:
   **Lanjutkan → pilih akun pemilik → Lanjutan (Advanced) → Buka Koperasi App (tidak aman) → Izinkan**.
   Peringatan "tidak aman" muncul karena skrip ini buatan sendiri, bukan aplikasi resmi Google.
4. Setelah izin diberikan, klik lagi **🔧 SETUP → ▶️ Jalankan Setup Lengkap**. Sheet `anggota`, `setting_rat`,
   `coa_akun`, `transaksi_…`, dan lainnya akan dibuat.

## Langkah 6: Buat akun admin pertama

1. Di Apps Script: **⚙️ Setelan proyek** → gulir ke **Properti skrip** → **Tambahkan properti skrip**, isi tiga baris:

   | Properti | Nilai |
   |---|---|
   | `ADMIN_EMAIL` | email untuk login aplikasi, misalnya `bendahara@gmail.com` |
   | `ADMIN_NAMA` | nama admin |
   | `ADMIN_SANDI` | sandi awal, minimal 6 karakter |

   Klik **Simpan properti skrip**.
2. Kembali ke **< > Editor**, buka file **`Kode_2`**. Di bilah atas, pada daftar fungsi (di sebelah tombol
   **Debug**) pilih **`buatAdminPertama`** → klik **▷ Jalankan**.
3. Di *Log eksekusi* muncul *"Admin pertama dibuat"*. Properti `ADMIN_SANDI` otomatis dihapus demi keamanan.

> Daftar fungsi hanya menampilkan fungsi dari file yang sedang dibuka. Bila `buatAdminPertama` tidak ada di
> `Kode_2`, coba buka `Kode_1` atau `Kode_3`. Baris *"Berisi: …"* di bagian atas tiap file menunjukkan isinya
> (fungsi ini berasal dari `sesi.js`).

## Langkah 7: Terbitkan sebagai aplikasi web

1. Di Apps Script, klik tombol biru **Terapkan (Deploy) → Deployment baru**.
2. Klik ikon ⚙️ di samping *Pilih jenis* → **Aplikasi web**.
3. Isi:
   - **Deskripsi:** misalnya `Versi 1`
   - **Jalankan sebagai:** **Saya** (akun pemilik)
   - **Yang memiliki akses:** **Siapa saja** (Anyone)
4. Klik **Terapkan**, lalu salin **URL aplikasi web** (berakhiran `/exec`). Link inilah yang dibagikan ke
   bendahara dan petugas, dan bisa dibuka di HP.

> "Siapa saja" aman karena aplikasi punya login sendiri, pembatasan peran, dan penguncian setelah 5 kali salah sandi.
> Akun Google sekolah (misalnya belajar.id) kadang tidak menyediakan pilihan "Siapa saja". Bila begitu, pasang
> dengan akun Gmail biasa.

## Langkah 8 (disarankan): Cadangan & pengingat otomatis

Buka file **`Kode_1`**, pilih fungsi **`pasangTugasHarian`** → **▷ Jalankan** (izinkan bila diminta).
Setiap hari pukul 06.00–07.00 aplikasi akan:

- tanggal 1: menyalin spreadsheet ke folder Drive *Backup Koperasi …* (12 salinan terakhir disimpan);
- tanggal 10: mengirim email pengingat angsuran ke admin;
- tanggal 25: mengirim email pengingat jasa sukarela.

(Fungsi ini berasal dari `log_email_backup.js`; bila tidak ada di `Kode_1`, cari di file `Kode_…` lain.)

## Langkah 9: Mulai memakai aplikasi

1. Buka link `/exec`, masuk dengan email dan sandi admin dari Langkah 6.
2. **Pengaturan → Koperasi & Tahun Buku → Identitas Koperasi**: isi nama koperasi, alamat, kota, badan hukum,
   sekolah/instansi, logo, serta nama pengurus dan pengawas. Semua cetakan dan laporan RAT memakai isian ini.
3. Atur tahun buku, jasa, dan pengguna: lihat [Panduan Pemasangan bagian 4](PANDUAN_PEMASANGAN.md#4-pengaturan-awal-di-aplikasi).
4. Isi anggota dan saldo awal: lihat [bagian 5B (manual)](PANDUAN_PEMASANGAN.md#5b-manual-koperasi-baru--excel-berbeda).

Panduan pemakaian harian untuk bendahara dan petugas: [`Panduan Aplikasi Koperasi.pdf`](Panduan%20Aplikasi%20Koperasi.pdf).

---

## Memperbarui ke versi baru

1. Di Apps Script, buka tiap file (`appsscript.json`, `Kode_1`, `Kode_2`, …, `index`), **Ctrl+A**, lalu tempel isi
   terbaru dari `paket_manual/`. **Ctrl+S**.
2. Bila versi baru punya lebih banyak file `Kode_N`, tambahkan file baru (Langkah 3). Bila lebih sedikit, hapus
   file `Kode_N` yang tidak ada lagi di `paket_manual/` (titik tiga → **Hapus**).
3. **Terapkan → Kelola deployment → ✏️ Edit → Versi: Versi baru → Terapkan**. Link `/exec` tetap sama.

Data koperasi di spreadsheet **tidak** terhapus saat kode diperbarui.

---

## Mengatasi masalah

| Gejala | Penyebab & jalan keluar |
|---|---|
| Menu 🔧 SETUP tidak muncul | Kode belum tersimpan (tanda titik oranye di nama file) atau belum muat ulang spreadsheet. Simpan semua file (Ctrl+S), lalu tekan F5 di spreadsheet. |
| *"ReferenceError: … is not defined"* | Ada file `Kode_N` yang belum ditempel atau tertempel tidak lengkap. Tempel ulang dengan tombol **Copy raw file**. |
| *"SyntaxError"* saat menyimpan | Isi file terpotong atau tercampur kode bawaan. Ctrl+A di file itu, tempel ulang. |
| Aplikasi web error: file HTML `index` tidak ditemukan (*"No HTML file named index was found"*) | File HTML belum dibuat atau namanya bukan `index` (Langkah 4). |
| *"Sorry, unable to open the file at this time"* | Browser login lebih dari satu akun Google. Buka link di jendela **Samaran/Incognito** (Ctrl+Shift+N). |
| Perubahan kode tidak terlihat di link `/exec` | Belum membuat **Versi baru** deployment (bagian *Memperbarui*, langkah 3). |
| Lupa sandi admin | Ulangi Langkah 6. |
| Akun terkunci | 5 kali salah sandi → tunggu 15 menit. |

Masalah lain: lihat [Panduan Pemasangan bagian 9](PANDUAN_PEMASANGAN.md#9-mengatasi-masalah), atau hubungi
[@asep94 di Threads](https://www.threads.com/@asep94).
