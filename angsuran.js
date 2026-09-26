/**
 * ============================================================
 * ANGSURAN.GS — Generate Angsuran Bulanan
 * Koperasi App
 * Sesuai dokumen: REKAP_Koperasi_v2.md (BAGIAN 3, 12, 15, 17, 18)
 * ============================================================
 *
 * ALUR (sesuai checklist BAGIAN 15, pertengahan bulan tgl 10-15):
 * 1. Admin klik "Generate Angsuran Bulan Ini"
 * 2. Sistem cari SEMUA pinjaman berstatus aktif
 *    (termasuk pinjaman hasil migrasi — satu jalur logika,
 *    sesuai keputusan v2)
 * 3. Untuk tiap pinjaman yang BELUM digenerate bulan ini:
 *    - Hitung porsi pokok & jasa bulan ini
 *    - Kalau ini pembayaran TERAKHIR (sisa <= angsuran normal),
 *      rapikan supaya sisa jadi PERSIS 0 (bukan minus/nyisa recehan)
 *    - Langsung PAID (tidak ada status pending — sesuai keputusan
 *      "semua angsuran via potong gaji")
 *    - Kas masuk + jurnal otomatis
 *    - Update status pinjaman jadi 'lunas' jika sisa = 0
 * 4. Anggota yang SUDAH digenerate bulan ini otomatis dilewati
 *    (aman diklik berkali-kali, tidak akan dobel)
 * ============================================================
 */

/**
 * Fungsi utama: generate angsuran untuk SEMUA pinjaman aktif,
 * untuk bulan & tahun tertentu.
 *
 * @param {number} [tahun]  default: tahun aktif
 * @param {number} [bulan]  default: bulan berjalan sekarang
 * @return {Object} ringkasan hasil generate
 */
function generateAngsuranBulanan(tahun, bulan) {
  const sekarang = new Date();
  tahun = tahun || getTahunAktif();
  bulan = bulan || (sekarang.getMonth() + 1);

  validateTahunAktif(tahun);

  const semuaPinjamanAktif = getRowsByFilter(SHEET.PINJAMAN, function(row) {
    return row.status === 'aktif' && row.status_lock !== STATUS_LOCK.VOID;
  });

  if (semuaPinjamanAktif.length === 0) {
    tampilkanPesan('Tidak ada pinjaman aktif. Tidak ada yang digenerate.\n\n' +
      'Cek sheet transaksi_pinjaman — pastikan ada baris dengan status = "aktif". ' +
      'Jika Anda baru menguji modul pinjaman.gs, pastikan ID anggota di test_pinjaman.gs ' +
      'sudah diganti sesuai ID yang benar-benar ada, dan tesnya berhasil (bukan ❌).');
    return { berhasil: [], dilewati: [], gagal: [],
      pesan: 'Tidak ada pinjaman aktif. Tidak ada yang digenerate.' };
  }

  const hasil = { berhasil: [], dilewati: [], gagal: [] };

  semuaPinjamanAktif.forEach(function(pinjaman) {
    try {
      // Lewati jika bulan ini SUDAH pernah digenerate untuk pinjaman ini
      const sudahAda = getRowsByFilter(SHEET.ANGSURAN, function(row) {
        return row.id_pinjaman === pinjaman.id_pinjaman &&
               Number(row.tahun) === Number(tahun) &&
               Number(row.bulan) === Number(bulan) &&
               row.status_lock !== STATUS_LOCK.VOID;
      });
      if (sudahAda.length > 0) {
        hasil.dilewati.push(pinjaman.id_pinjaman + ' (sudah digenerate bulan ini)');
        return;
      }
      // Jasa menurun: pinjaman yang baru cair bulan ini mulai ditagih bulan depan
      if (isJasaMenurun(pinjaman) &&
          new Date(pinjaman.tanggal) >= new Date(tahun, bulan - 1, 1)) {
        hasil.dilewati.push(pinjaman.id_pinjaman + ' (baru cair bulan ini — ditagih mulai bulan depan)');
        return;
      }

      const idAngsuran = generateSatuAngsuran(pinjaman, tahun, bulan);
      hasil.berhasil.push(pinjaman.id_pinjaman + ' → ' + idAngsuran);

    } catch (e) {
      hasil.gagal.push(pinjaman.id_pinjaman + ': ' + e.message);
    }
  });

  const pesan =
    '✅ GENERATE ANGSURAN — Tahun ' + tahun + ' Bulan ' + bulan + '\n\n' +
    'Berhasil: ' + hasil.berhasil.length + '\n' +
    (hasil.berhasil.length ? hasil.berhasil.join('\n') + '\n\n' : '\n') +
    'Dilewati (sudah ada): ' + hasil.dilewati.length + '\n' +
    (hasil.dilewati.length ? hasil.dilewati.join('\n') + '\n\n' : '\n') +
    (hasil.gagal.length ? 'GAGAL: ' + hasil.gagal.length + '\n' + hasil.gagal.join('\n') : '');

  tampilkanPesan(pesan);
  return hasil;
}

/**
 * Generate SATU baris angsuran untuk satu pinjaman.
 * Dipisah dari fungsi utama supaya bisa dites/dipanggil sendiri.
 */
function generateSatuAngsuran(pinjaman, tahun, bulan) {
  const anggota = getAnggota(pinjaman.id_anggota);
  const namaAnggota = anggota ? anggota.nama : pinjaman.id_anggota;

  const sisaSebelumBayar = getSisaPinjaman(pinjaman.id_pinjaman);
  if (sisaSebelumBayar <= 0) {
    // Sudah lunas tapi statusnya belum ke-update — rapikan saja
    perbaruiStatusPinjamanJikaLunas(pinjaman.id_pinjaman);
    throw new Error('Pinjaman ini sisa-nya sudah 0 (lunas), tidak perlu angsuran lagi.');
  }

  let angsuranPokok, jasaBulanIni, totalBayar;

  if (isJasaMenurun(pinjaman)) {
    // ---- JASA MENURUN: pokok tetap, jasa = persen × sisa pokok awal bulan ----
    angsuranPokok = Math.min(Number(pinjaman.angsuran_perbulan), sisaSebelumBayar);
    jasaBulanIni = hitungJasaMenurun(pinjaman, tahun, bulan);
    totalBayar = angsuranPokok + jasaBulanIni;
  } else {
    // ---- DATA LAMA (jasa flat): porsi pokok & jasa proporsional ----
    const rasioPokok = Number(pinjaman.nominal) / Number(pinjaman.total_tagihan);
    const angsuranNormal = Number(pinjaman.angsuran_perbulan);

    // Pembayaran terakhir dirapikan supaya sisa persis 0
    totalBayar = sisaSebelumBayar <= angsuranNormal ? sisaSebelumBayar : angsuranNormal;
    angsuranPokok = Math.round(totalBayar * rasioPokok);
    jasaBulanIni = totalBayar - angsuranPokok;
  }

  const tanggalGenerate = new Date(tahun, bulan - 1, 10); // tanggal 10, sesuai batas bayar

  const idAngsuran = generateId(SHEET.ANGSURAN, ID_PREFIX.ANGSURAN, 1, tahun, null);

  const rowAngsuran = {
    id_angsuran: idAngsuran,
    tanggal: tanggalGenerate,
    tahun: tahun,
    bulan: bulan,
    id_pinjaman: pinjaman.id_pinjaman,
    id_anggota: pinjaman.id_anggota,
    angsuran_pokok: angsuranPokok,
    jasa: jasaBulanIni,
    denda: 0,  // default 0 — hanya diisi manual untuk kasus khusus (keputusan v2)
    total_bayar: totalBayar,
    keterangan: 'Generate otomatis — potong gaji',
    user_input: getUserEmail(),
    timestamp: new Date(),
    status_lock: STATUS_LOCK.OPEN
  };
  appendRowFromObject(SHEET.ANGSURAN, rowAngsuran);

  // ---------- KAS MASUK ----------
  catatKasMasuk({
    tanggal: tanggalGenerate,
    tahun: tahun,
    bulan: bulan,
    kategori: 'angsuran',
    referensi: idAngsuran,
    keterangan: 'Angsuran pinjaman — ' + namaAnggota,
    nominal: totalBayar
  });

  // ---------- JURNAL (BAGIAN 12 #5) ----------
  // Debit Kas (total_bayar) = Kredit Piutang (pokok) + Kredit Pendapatan Jasa (jasa)
  catatJurnal(tanggalGenerate, tahun, bulan, idAngsuran,
    'Angsuran pinjaman — ' + namaAnggota,
    [
      { kode_akun: AKUN.KAS, debit: totalBayar, kredit: 0 },
      { kode_akun: AKUN.PIUTANG_PINJAMAN, debit: 0, kredit: angsuranPokok },
      { kode_akun: AKUN.PENDAPATAN_JASA_PINJAMAN, debit: 0, kredit: jasaBulanIni }
    ]
  );

  logAktivitas('INSERT', SHEET.ANGSURAN, idAngsuran, null, rowAngsuran);

  // ---------- UPDATE STATUS PINJAMAN JIKA LUNAS ----------
  perbaruiStatusPinjamanJikaLunas(pinjaman.id_pinjaman);

  return idAngsuran;
}

// ============================================================
// VALIDASI (BAGIAN 17)
// ============================================================

function validateGenerateAngsuran(tahun, bulan) {
  validateTahunAktif(tahun);

  const pinjamanAktif = getRowsByFilter(SHEET.PINJAMAN, function(row) {
    return row.status === 'aktif';
  });
  if (pinjamanAktif.length === 0) {
    throw new Error('Tidak ada pinjaman aktif untuk digenerate.');
  }
}

// ============================================================
// CETAK STRUK POTONGAN GAJI (BAGIAN 7, 14, 15)
// ============================================================

/**
 * Tulis rekap angsuran bulan tertentu ke sheet struk_potongan,
 * siap dicetak/diserahkan ke bendahara gaji.
 */
function cetakStrukPotongan(tahun, bulan) {
  tahun = tahun || getTahunAktif();
  bulan = bulan || (new Date().getMonth() + 1);

  const angsuranBulanIni = getRowsByFilter(SHEET.ANGSURAN, function(row) {
    return Number(row.tahun) === Number(tahun) && Number(row.bulan) === Number(bulan) &&
           row.status_lock !== STATUS_LOCK.VOID;
  });

  const sheet = getSheet(SHEET.STRUK_POTONGAN);
  sheet.clear();

  sheet.getRange('A1').setValue(
    '📋 STRUK POTONGAN GAJI — Koperasi ' + getProfilKoperasi().nama_koperasi.replace(/"/g, '') + ' — Bulan ' + bulan + '/' + tahun
  ).setFontWeight('bold').setFontSize(13);

  const header = ['No', 'Nama', 'NIP/NIS', 'Jabatan', 'Angsuran Pokok', 'Jasa', 'Denda', 'Total Potongan'];
  sheet.getRange(3, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');

  if (angsuranBulanIni.length === 0) {
    sheet.getRange(4, 1).setValue('(Belum ada angsuran untuk bulan ini)').setFontStyle('italic');
    return { jumlahAnggota: 0, totalPotongan: 0 };
  }

  const baris = angsuranBulanIni.map(function(row, i) {
    const anggota = getAnggota(row.id_anggota);
    return [
      i + 1,
      anggota ? anggota.nama : row.id_anggota,
      anggota ? anggota.nip_nis : '',
      anggota ? anggota.jabatan : '',
      row.angsuran_pokok,
      row.jasa,
      row.denda,
      row.total_bayar
    ];
  });
  sheet.getRange(4, 1, baris.length, header.length).setValues(amanBaris_(baris));
  sheet.getRange(4, 5, baris.length, 4).setNumberFormat('Rp #,##0');

  const totalRow = 4 + baris.length;
  const totalPotongan = baris.reduce(function(s, r) { return s + r[7]; }, 0);
  sheet.getRange(totalRow, 4).setValue('TOTAL').setFontWeight('bold');
  sheet.getRange(totalRow, 8).setValue(totalPotongan).setFontWeight('bold')
    .setNumberFormat('Rp #,##0');

  for (let c = 1; c <= header.length; c++) sheet.autoResizeColumn(c);

  tampilkanPesan('✅ Struk potongan bulan ' + bulan + '/' + tahun + ' siap.\n' +
    'Jumlah anggota: ' + baris.length + '\n' +
    'Total potongan: ' + formatRupiah(totalPotongan) + '\n\n' +
    'Cek sheet struk_potongan untuk cetak/export.');

  return { jumlahAnggota: baris.length, totalPotongan: totalPotongan };
}

// ============================================================
// PEMBUNGKUS UNTUK MENU (tanpa parameter — pakai tahun/bulan berjalan)
// ============================================================

function menuGenerateAngsuranBulanIni() {
  generateAngsuranBulanan(); // default: tahun aktif, bulan sekarang
}

function menuCetakStrukPotongan() {
  cetakStrukPotongan(); // default: tahun aktif, bulan sekarang
}