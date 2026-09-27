# Koperasi App

Aplikasi pembukuan koperasi simpan pinjam (anggota guru/pegawai, potong gaji), berbasis
Google Apps Script dengan Google Sheets sebagai basis data.

## Fitur

- **Anggota**: profil, saldo simpanan (pokok, wajib, sukarela), riwayat, anggota keluar.
- **Transaksi**: setoran, penarikan dengan persetujuan admin, pinjaman (jasa 1,5%/bulan menurun, pinjaman tambahan digabung), angsuran, pembatalan dengan jejak.
- **Bulanan**: tagihan & potong gaji (slip 6 per lembar), jasa simpanan sukarela (rumus diatur per tahun buku).
- **Keuangan**: kas & pengeluaran, Buku Kas, Rugi Laba, Neraca Komparatif, Penjelasan Neraca, SHU per anggota (7 pos alokasi + pajak).
- **Laporan RAT**: file Word (kertas F4) berisi undangan, tata tertib, laporan pengurus & pengawas beserta rasio keuangan, rencana kerja.
- **Tahun buku**: tutup buku dengan jurnal penutup, saldo awal, migrasi dari Excel.
- **Keamanan**: login aplikasi sendiri (sandi di-hash + garam), peran admin/petugas, batas percobaan login, semua perubahan tercatat di log.

## Struktur

| File | Isi |
|---|---|
| `index.html` | Seluruh tampilan (satu halaman) |
| `webapp.js`, `sesi.js` | Pintu masuk web, login, sesi, pembatasan akses |
| `database.js`, `config.js`, `setup.js` | Akses sheet, konstanta, pembuatan sheet |
| `simpanan.js`, `pinjaman.js`, `angsuran.js`, `bayar_angsuran.js`, `approval.js` | Transaksi |
| `tagihan.js`, `jasa.js` | Potong gaji & jasa sukarela |
| `akuntansi.js`, `laporan*.js`, `tutup_buku.js`, `rat.js` | Jurnal, laporan, tutup buku, laporan RAT |
| `migrasi.js` | Pemindahan data dari Excel pembukuan |
| `identitas_lokal.contoh.js` | Contoh nama pengurus (salin menjadi `identitas_lokal.js`) |
| `paket_manual/` | Kode siap salin-tempel untuk [Panduan Pemula](docs/PANDUAN_PEMULA.md); dibuat oleh `tools/buat_paket_manual.py` |

## Data pribadi

Repositori ini **tidak** berisi data anggota maupun keuangan. File berikut sengaja dikecualikan (`.gitignore`):
pembukuan Excel/Word, folder `referensi/`, `migrasi_data.js`, `identitas_lokal.js`, dan `.clasp.json`.
Seluruh data koperasi hanya tersimpan di Google Sheets milik koperasi.

## Panduan

- **[Panduan Pemula](docs/PANDUAN_PEMULA.md)**: cara termudah, cukup lewat browser. Salin-tempel beberapa file
  dari folder [`paket_manual/`](paket_manual) ke editor Apps Script, tanpa Node.js, Git, atau clasp.
- **[Panduan Pemasangan](docs/PANDUAN_PEMASANGAN.md)**: memasang dari nol, memasang ulang, memindahkan data
  dari Excel pembukuan, memakai untuk koperasi lain, memperbarui kode, dan mengatasi masalah.
- **`tools/ekstrak_excel.py`**: mengubah Excel pembukuan menjadi data migrasi yang dicocokkan per bulan.

Ringkas pemasangan: `clasp login` → salin `.clasp.json.contoh` & `identitas_lokal.contoh.js` → `clasp push` →
menu 🔧 SETUP → Jalankan Setup Lengkap → `buatAdminPertama` → Deploy Web App (Execute as: Me, Anyone).

Panduan pemakaian untuk admin dan petugas: [`docs/Panduan Aplikasi Koperasi.pdf`](docs/Panduan%20Aplikasi%20Koperasi.pdf).

## Kredit

Dikembangkan oleh **Asep Hanuryana**, Threads: [@asep94](https://www.threads.com/@asep94).

Jika aplikasi ini bermanfaat atau Anda membagikannya, mohon cantumkan kredit dan mention
[@asep94](https://www.threads.com/@asep94) di Threads.

## Lisensi, komersialisasi & kustomisasi

Hak Cipta © 2026 Asep Hanuryana. **Seluruh hak dilindungi.** Lihat [LICENSE](LICENSE).

Repositori ini boleh dilihat publik, tetapi **memasang, memakai, menyalin, mengubah, menyebarkan,
atau menjual** aplikasi ini memerlukan izin tertulis dari pemilik.

Hubungi [@asep94 di Threads](https://www.threads.com/@asep94) untuk:

- izin memakai aplikasi di koperasi Anda,
- komersialisasi atau kerja sama,
- kustomisasi fitur sesuai kebutuhan koperasi Anda.

## Keamanan

Menemukan celah keamanan? **Jangan** membuat Issue publik. Ikuti petunjuk di [SECURITY.md](SECURITY.md).
