# Kebijakan Keamanan

Aplikasi ini mengelola data pribadi anggota dan data keuangan koperasi, jadi laporan keamanan
kami tangani dengan serius.

## Melaporkan celah keamanan

**Jangan laporkan celah keamanan lewat Issue publik.** Issue bisa dibaca siapa saja, termasuk
pihak yang ingin menyalahgunakan celah tersebut sebelum diperbaiki.

Laporkan secara pribadi dengan salah satu cara berikut:

1. **GitHub Private Vulnerability Reporting**: tab **Security** → **Report a vulnerability**
   (jika fitur ini aktif di repositori).
2. **Pesan langsung (DM) di Threads**: [@asep94](https://www.threads.com/@asep94).

Sertakan dalam laporan:

- Bagian yang terdampak (file atau fungsi, misalnya `sesi.js`, `webapp.js`).
- Langkah untuk mereproduksi masalah.
- Dampak yang mungkin terjadi (misalnya melewati login, melihat data anggota lain, mengubah saldo).
- Saran perbaikan, jika ada.

## Yang termasuk lingkup

- Melewati login, sesi, atau pembatasan peran (admin/petugas).
- Akses atau perubahan data anggota dan transaksi tanpa hak.
- Kebocoran data pribadi (nama, NIK, saldo, pinjaman) atau kredensial.
- Injeksi (misalnya HTML/skrip di `index.html`, rumus di Google Sheets).
- Kelemahan penyimpanan sandi atau batas percobaan login.

## Yang di luar lingkup

- Celah pada layanan Google (Apps Script, Sheets, Drive) itu sendiri. Laporkan ke
  [Google](https://bughunters.google.com/).
- Pemasangan yang tidak mengikuti [Panduan Pemasangan](docs/PANDUAN_PEMASANGAN.md), misalnya
  membagikan Spreadsheet ke publik.
- Serangan yang membutuhkan akses ke akun Google pemilik.

## Setelah laporan diterima

- Laporan akan ditanggapi sesegera mungkin.
- Rincian celah tidak akan dipublikasikan sebelum perbaikan tersedia.
- Pelapor akan diberi kredit atas temuannya, kecuali ia meminta untuk tidak disebutkan.

## Tips keamanan untuk pengelola koperasi

- Jangan membagikan Spreadsheet basis data ke publik ("Siapa saja yang memiliki link").
- Jangan meng-commit `identitas_lokal.js`, `.clasp.json`, `migrasi_data.js`, atau berkas
  pembukuan Excel/Word ke repositori (sudah dikecualikan lewat `.gitignore`).
- Pakai sandi yang kuat untuk akun admin, dan hapus akun petugas yang sudah tidak aktif.
- Data anggota termasuk data pribadi yang dilindungi Undang-Undang Nomor 27 Tahun 2022 tentang
  Pelindungan Data Pribadi.
