# Paket salin-tempel

File di folder ini untuk memasang Koperasi App **tanpa clasp**, dengan menyalin-tempel langsung ke editor
Apps Script. Langkah lengkapnya: **[Panduan Pemula](../docs/PANDUAN_PEMULA.md)**.

- `Kode_N.gs`: gabungan semua file kode aplikasi (`*.js` di folder utama), dibagi agar tiap file tidak terlalu besar.
- `index.html`, `appsscript.json`: salinan apa adanya dari folder utama.

Jangan menyunting file di sini. Semuanya dibuat otomatis oleh `tools/buat_paket_manual.py`; setelah mengubah
kode, jalankan `python tools/buat_paket_manual.py` lalu commit hasilnya.

Hak Cipta © Asep Hanuryana ([@asep94](https://www.threads.com/@asep94)). Seluruh hak dilindungi; lihat [LICENSE](../LICENSE).
