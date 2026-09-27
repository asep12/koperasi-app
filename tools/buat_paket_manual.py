#!/usr/bin/env python3
"""
Membuat folder paket_manual/ untuk pemasangan TANPA laptop/clasp (salin-tempel di editor Apps Script).

Semua file kode aplikasi (*.js yang ikut `clasp push`) digabung menjadi beberapa file Kode_N.gs,
supaya pemula cukup menyalin beberapa file, bukan puluhan. index.html dan appsscript.json disalin apa adanya.

Pemakaian (jalankan dari folder repo setiap kali kode berubah, lalu commit hasilnya):
    python tools/buat_paket_manual.py          # buat ulang paket_manual/
    python tools/buat_paket_manual.py --cek    # hanya periksa apakah paket_manual/ sudah sesuai kode
"""
import argparse
import fnmatch
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / 'paket_manual'

# Sama dengan pengecualian di .claspignore, ditambah identitas_lokal.js (berisi nama orang, tidak boleh
# masuk paket publik; tanpa file ini aplikasi memakai isian Pengaturan → Identitas Koperasi).
KECUALI = ['tes_*.js', 'test*.js', 'fix.js', 'migrasi_data.js',
           'identitas_lokal.contoh.js', 'identitas_lokal.js']

# Batas ukuran per file hasil gabungan. index.html (±200 KB) terbukti lancar di editor Apps Script.
BATAS_BYTE = 200_000

KEPALA = """\
/**
 * {nama} — bagian {ke} dari {total} kode Koperasi App.
 * Hak Cipta (c) Asep Hanuryana (Threads: @asep94). Seluruh hak dilindungi; lihat LICENSE.
 *
 * FILE INI DIBUAT OTOMATIS oleh tools/buat_paket_manual.py. Jangan disunting di sini:
 * ubah file aslinya di repositori, lalu jalankan ulang alat tersebut.
 *
 * Berisi: {daftar}
 */
"""


def file_kode():
    return sorted(p for p in ROOT.glob('*.js')
                  if not any(fnmatch.fnmatch(p.name, pola) for pola in KECUALI))


def bagi(files):
    """Kelompokkan file (urutan tetap) agar tiap kelompok tidak melebihi BATAS_BYTE."""
    kelompok, kini, ukuran = [], [], 0
    for f in files:
        n = len(f.read_bytes())
        if kini and ukuran + n > BATAS_BYTE:
            kelompok.append(kini)
            kini, ukuran = [], 0
        kini.append(f)
        ukuran += n
    if kini:
        kelompok.append(kini)
    return kelompok


def susun():
    """Kembalikan {nama_file: isi} untuk seluruh isi paket_manual/."""
    hasil = {}
    kelompok = bagi(file_kode())
    for i, grup in enumerate(kelompok, 1):
        nama = f'Kode_{i}.gs'
        isi = KEPALA.format(nama=nama, ke=i, total=len(kelompok),
                            daftar=', '.join(f.name for f in grup))
        for f in grup:
            isi += f'\n// ===== {f.name} =====\n' + f.read_text(encoding='utf-8').rstrip() + '\n'
        hasil[nama] = isi
    hasil['index.html'] = (ROOT / 'index.html').read_text(encoding='utf-8')
    hasil['appsscript.json'] = (ROOT / 'appsscript.json').read_text(encoding='utf-8')
    return hasil


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--cek', action='store_true', help='periksa saja, jangan menulis')
    args = ap.parse_args()

    paket = susun()
    lama = {p.name for p in OUT.glob('*') if p.name != 'BACA_SAYA.md'} if OUT.exists() else set()

    if args.cek:
        beda = [n for n, isi in paket.items()
                if not (OUT / n).exists() or (OUT / n).read_text(encoding='utf-8') != isi]
        beda += sorted(lama - set(paket))
        if beda:
            print('paket_manual/ belum diperbarui: ' + ', '.join(beda))
            print('Jalankan: python tools/buat_paket_manual.py')
            sys.exit(1)
        print('paket_manual/ sudah sesuai kode.')
        return

    OUT.mkdir(exist_ok=True)
    for n in lama - set(paket):
        (OUT / n).unlink()
    for n, isi in paket.items():
        (OUT / n).write_text(isi, encoding='utf-8', newline='\n')
        print(f'{n:18} {len(isi.encode("utf-8")):>8,} byte')


if __name__ == '__main__':
    main()
