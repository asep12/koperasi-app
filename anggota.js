/**
 * ============================================================
 * ANGGOTA.GS — Generate ID Massal untuk Data Anggota
 * ============================================================
 * Kasus pakai: Anda sudah isi kolom "nama" (dan mungkin
 * "jabatan"/"status") untuk banyak baris, tapi "id_anggota"
 * masih kosong. Fungsi ini otomatis mengisi id_anggota untuk
 * SEMUA baris yang: nama-nya sudah ada, id_anggota-nya kosong.
 *
 * Baris yang id_anggota-nya SUDAH terisi tidak akan disentuh/
 * ditimpa — aman dijalankan berulang kali.
 *
 * TIDAK PERLU kolom lain (nip_nis, alamat, telepon,
 * tanggal_masuk) terisi dulu — boleh disusulkan belakangan.
 *
 * CARA PAKAI:
 * Menu 🔧 SETUP → 🆔 Generate ID Anggota Kosong
 * (atau jalankan generateIdAnggotaMassal dari editor)
 * ============================================================
 */

function generateIdAnggotaMassal() {
  const sheet = getSheet(SHEET.ANGGOTA);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    tampilkanPesan('Belum ada data anggota sama sekali. Isi kolom "nama" dulu.');
    return;
  }

  const header = getHeader(SHEET.ANGGOTA);
  const idxId = header.indexOf('id_anggota');
  const idxNama = header.indexOf('nama');

  if (idxId === -1 || idxNama === -1) {
    tampilkanPesan('❌ Kolom id_anggota atau nama tidak ditemukan di sheet anggota.');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, header.length).getValues();

  // Cari nomor urut terbesar yang SUDAH ada, supaya lanjut dari situ
  let nomorTerbesar = 0;
  data.forEach(function(row) {
    const id = String(row[idxId]);
    if (id.indexOf('AGT-') === 0) {
      const nomor = parseInt(id.substring(4), 10);
      if (!isNaN(nomor) && nomor > nomorTerbesar) nomorTerbesar = nomor;
    }
  });

  let terisi = 0;
  let dilewatiKosong = 0;
  let dilewatiSudahAda = 0;

  const updatedRows = []; // {rowIndex(1-based di sheet), id}

  data.forEach(function(row, i) {
    const idSaatIni = String(row[idxId]).trim();
    const namaSaatIni = String(row[idxNama]).trim();

    if (idSaatIni !== '') {
      dilewatiSudahAda++;
      return; // sudah punya ID, jangan disentuh
    }
    if (namaSaatIni === '') {
      dilewatiKosong++;
      return; // baris kosong total, lewati
    }

    nomorTerbesar++;
    const idBaru = 'AGT-' + String(nomorTerbesar).padStart(4, '0');
    updatedRows.push({ rowNumber: i + 2, id: idBaru }); // +2 karena header + 0-index
    terisi++;
  });

  // Tulis sekaligus (lebih cepat daripada satu-satu)
  updatedRows.forEach(function(item) {
    sheet.getRange(item.rowNumber, idxId + 1).setValue(item.id);
  });

  if (terisi > 0) {
    logAktivitas('INSERT', SHEET.ANGGOTA, terisi + ' baris',
      null, { keterangan: 'Generate ID massal: ' + terisi + ' anggota baru diberi ID' });
  }

  tampilkanPesan(
    '✅ SELESAI\n\n' +
    'ID baru dibuat: ' + terisi + '\n' +
    'Dilewati (sudah punya ID): ' + dilewatiSudahAda + '\n' +
    'Dilewati (baris kosong/nama kosong): ' + dilewatiKosong + '\n\n' +
    (terisi > 0
      ? 'ID terakhir yang dibuat: AGT-' + String(nomorTerbesar).padStart(4, '0')
      : 'Tidak ada baris baru yang perlu diisi.')
  );
}