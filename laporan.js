/**
 * ============================================================
 * LAPORAN.GS — Laba Rugi, Neraca, SHU, dan Rekap
 * Koperasi App
 * Sesuai dokumen: REKAP_Koperasi_v2.md (BAGIAN 2B, 7, 18)
 * ============================================================
 *
 * SUMBER DATA SEMUA LAPORAN = jurnal_umum.
 * Karena setiap transaksi selalu membuat jurnal balance,
 * laporan di sini tinggal menjumlahkan per akun.
 *
 * Semua laporan menerima parameter tahun (default: tahun aktif)
 * sesuai keputusan v2 #6 (multi-tahun).
 * ============================================================
 */

// ============================================================
// FONDASI: saldo per akun, alokasi SHU, laba rugi & neraca
// ada di akuntansi.js (getSaldoPerAkun, getTotalSHUTahun,
// hitungAlokasiDana, hitungLabaRugi, hitungNeraca).
// ============================================================

// ============================================================
// SHU PER ANGGOTA (basis: DANA ANGGOTA dari alokasi SHU)
// ============================================================

/**
 * Hitung SHU satu anggota untuk satu tahun buku.
 * Sesuai Excel: Dana Anggota (mis. 50% SHU) dibagi menjadi porsi
 * simpanan (pokok + wajib per 31 Desember) dan porsi jasa (jasa
 * pinjaman yang dibayar anggota selama tahun itu). Sukarela tidak ikut.
 */
function hitungSHU(idAnggota, tahun) {
  const basis = getBasisSHU(tahun);
  const setting = basis.setting;
  const danaAnggota = basis.alokasi.danaAnggota;

  const simpananAnggota = basis.simpananPerAnggota[String(idAnggota)] || 0;
  const shuSimpanan = basis.totalSimpananAll > 0
    ? Math.round((Number(setting.shu_simpanan) / 100) * danaAnggota *
        simpananAnggota / basis.totalSimpananAll)
    : 0;

  const jasaAnggota = basis.jasaPerAnggota[String(idAnggota)] || 0;
  const shuJasa = basis.totalJasaAll > 0
    ? Math.round((Number(setting.shu_jasa) / 100) * danaAnggota * jasaAnggota / basis.totalJasaAll)
    : 0;

  return {
    shuSimpanan: shuSimpanan,
    shuJasa: shuJasa,
    total: shuSimpanan + shuJasa,
    simpananAnggota: simpananAnggota,
    jasaAnggota: jasaAnggota
  };
}

/** Angka dasar SHU satu tahun (dipakai bersama oleh semua anggota). */
function getBasisSHU(tahun) {
  return cacheTurunan('basisSHU_' + tahun, function() {
    const setting = getSettingRAT(tahun);
    if (!setting) throw new Error('Setting RAT tahun ' + tahun + ' tidak ditemukan.');

    const peta = getPetaSaldoSimpanan(tahun);   // posisi 31 Desember tahun itu
    const simpananPerAnggota = {};
    let totalSimpananAll = 0;
    Object.keys(peta).forEach(function(id) {
      const n = Math.max(0, peta[id].pokok) + Math.max(0, peta[id].wajib);
      simpananPerAnggota[id] = n;
      totalSimpananAll += n;
    });

    const jasaPerAnggota = {};
    let totalJasaAll = 0;
    sheetToObjects(SHEET.ANGSURAN).forEach(function(row) {
      if (row.status_lock === STATUS_LOCK.VOID) return;
      if (Number(row.tahun) !== Number(tahun)) return;
      const jasa = Number(row.jasa) || 0;
      const id = String(row.id_anggota);
      jasaPerAnggota[id] = (jasaPerAnggota[id] || 0) + jasa;
      totalJasaAll += jasa;
    });

    return {
      setting: setting,
      alokasi: hitungAlokasiDana(tahun),
      simpananPerAnggota: simpananPerAnggota,
      totalSimpananAll: totalSimpananAll,
      totalJasaAll: totalJasaAll,
      jasaPerAnggota: jasaPerAnggota
    };
  });
}

/** Daftar SHU semua anggota yang punya simpanan atau membayar jasa pada tahun itu. */
function daftarSHUAnggota(tahun) {
  const basis = getBasisSHU(tahun);
  return sheetToObjects(SHEET.ANGGOTA).map(function(a) {
    const s = hitungSHU(a.id_anggota, tahun);
    return { id: a.id_anggota, nama: a.nama, status: a.status,
      simpanan: s.simpananAnggota, jasa: s.jasaAnggota,
      shuSimpanan: s.shuSimpanan, shuJasa: s.shuJasa, total: s.total };
  }).filter(function(r) {
    return r.simpanan > 0 || r.jasa > 0 || basis.simpananPerAnggota[String(r.id)] > 0;
  });
}

// ============================================================
// PENULIS SHEET LAPORAN (arsip di spreadsheet)
// ============================================================

/**
 * Tulis baris laporan sekaligus. Setiap baris: [label, nilai, gaya]
 * gaya: 'judul' | 'sub' | 'total' | '' (nilai kosong = baris teks).
 */
function tulisSheetLaporan(namaSheet, judul, baris) {
  const sheet = getSheet(namaSheet);
  sheet.clear();
  const isi = [[judul, ''], ['Digenerate: ' + Utilities.formatDate(new Date(),
    Session.getScriptTimeZone(), 'dd MMMM yyyy, HH:mm'), ''], ['', '']]
    .concat(baris.map(function(b) { return [b[0], b[1] === undefined ? '' : b[1]]; }));
  sheet.getRange(1, 1, isi.length, 2).setValues(amanBaris_(isi));
  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  sheet.getRange(2, 1).setFontStyle('italic').setFontColor('#666666');
  sheet.getRange(4, 2, Math.max(1, baris.length), 1).setNumberFormat('Rp #,##0');
  baris.forEach(function(b, i) {
    if (b[2] === 'judul' || b[2] === 'total') sheet.getRange(4 + i, 1, 1, 2).setFontWeight('bold');
  });
  sheet.setColumnWidth(1, 360); sheet.setColumnWidth(2, 170);
  return sheet;
}

// ============================================================
// LAPORAN 1: LABA RUGI → sheet lap_labarugi
// ============================================================

function generateLaporanLabaRugi(tahun) {
  tahun = tahun || getTahunAktif();
  const lr = hitungLabaRugi(tahun);
  const baris = [['I. PENDAPATAN', '', 'judul']];
  lr.pendapatan.forEach(function(p) { baris.push(['   ' + p.label, p.nilai]); });
  baris.push(['   JUMLAH PENDAPATAN', lr.totalPendapatan, 'total'], ['', '']);
  baris.push(['II. BIAYA-BIAYA', '', 'judul']);
  lr.biaya.forEach(function(b) {
    baris.push(['   ' + b.label, b.nilai]);
    (b.rincian || []).forEach(function(r) { baris.push(['        - ' + r.label, r.nilai]); });
  });
  baris.push(['   JUMLAH BIAYA', lr.totalBiaya, 'total'], ['', '']);
  baris.push(['SHU TAHUN BERJALAN BELUM PAJAK ' + tahun, lr.shu, 'total']);
  baris.push(['Pajak ' + lr.pajakPersen + '% dari pendapatan', lr.pajak]);
  baris.push(['SHU TAHUN BERJALAN SUDAH PAJAK', lr.shuSetelahPajak, 'total'], ['', '']);
  if (lr.alokasi.length) {
    baris.push(['Yang akan dialokasikan untuk:', '', 'judul']);
    lr.alokasi.forEach(function(p) { baris.push(['   - ' + p.label + ' ' + p.persen + '%', p.nominal]); });
  }
  tulisSheetLaporan(SHEET.LAP_LABARUGI, 'DAFTAR RUGI LABA — PER 31 DESEMBER ' + tahun, baris);

  tampilkanPesan('✅ Laporan Laba Rugi tahun ' + tahun + ' selesai.\n' +
    'SHU belum pajak: ' + formatRupiah(lr.shu) + '\n\nCek sheet lap_labarugi.');
  return lr.shu;
}

// ============================================================
// LAPORAN 2: NERACA → sheet lap_neraca
// ============================================================

function generateLaporanNeraca(tahun) {
  tahun = tahun || getTahunAktif();
  const n = hitungNeraca(tahun);
  const kelompok = function(judul, daftar) {
    return [[judul, '', 'judul']].concat(daftar.map(function(r) { return ['   ' + r.label, r.nilai]; }));
  };
  const baris = []
    .concat(kelompok('I. AKTIVA LANCAR', n.aktivaLancar))
    .concat(kelompok('II. AKTIVA TETAP', n.aktivaTetap))
    .concat([['JUMLAH AKTIVA', n.totalAktiva, 'total'], ['', '']])
    .concat(kelompok('III. HUTANG LANCAR', n.hutangLancar))
    .concat(kelompok('   Dana-dana', n.danaDana))
    .concat(kelompok('IV. MODAL SENDIRI', n.modalSendiri))
    .concat([['V. SHU TAHUN BERJALAN', n.shuBerjalan, 'judul'],
             ['JUMLAH PASIVA', n.totalPasiva, 'total'], ['', ''],
             [n.balance ? '✅ BALANCE' : '❌ SELISIH', n.selisih]]);
  const sheet = tulisSheetLaporan(SHEET.LAP_NERACA, 'NERACA — PER 31 DESEMBER ' + tahun, baris);
  sheet.getRange(3 + baris.length, 1).setFontColor(n.balance ? '#38761d' : '#cc0000');

  tampilkanPesan('✅ Neraca tahun ' + tahun + ' selesai.\n' +
    'Aktiva: ' + formatRupiah(n.totalAktiva) + '\n' +
    'Pasiva: ' + formatRupiah(n.totalPasiva) + '\n' +
    (n.balance ? '✅ BALANCE' : '❌ SELISIH ' + formatRupiah(n.selisih)) +
    '\n\nCek sheet lap_neraca.');
  return { totalAktiva: n.totalAktiva, totalPasiva: n.totalPasiva, balance: n.balance };
}

// ============================================================
// LAPORAN 3: SHU PER ANGGOTA → sheet lap_shu
// ============================================================

function generateLaporanSHU(tahun) {
  tahun = tahun || getTahunAktif();
  const setting = getSettingRAT(tahun);
  const alokasi = hitungAlokasiDana(tahun);
  const daftar = daftarSHUAnggota(tahun);

  const sheet = getSheet(SHEET.LAP_SHU);
  sheet.clear();
  sheet.getRange('A1').setValue('DAFTAR SISA HASIL USAHA (SHU) — Tahun Buku ' + tahun)
    .setFontWeight('bold').setFontSize(13);

  const ringkasan = [
    ['SHU tahun berjalan (belum pajak)', alokasi.shuKotor],
    ['Dana Anggota ' + (alokasi.pos[1] ? alokasi.pos[1].persen : '') + '%', alokasi.danaAnggota],
    ['   Porsi SHU simpanan ' + setting.shu_simpanan + '%',
      Math.round(alokasi.danaAnggota * setting.shu_simpanan / 100)],
    ['   Porsi SHU jasa pinjaman ' + setting.shu_jasa + '%',
      Math.round(alokasi.danaAnggota * setting.shu_jasa / 100)]
  ];
  sheet.getRange(3, 1, ringkasan.length, 2).setValues(amanBaris_(ringkasan));
  sheet.getRange(3, 2, ringkasan.length, 1).setNumberFormat('Rp #,##0');

  const baseRow = 3 + ringkasan.length + 2;
  const header = ['No', 'Nama Anggota', 'Simpanan (pokok+wajib)', 'Jasa Pinjaman',
                  'SHU Simpanan', 'SHU Jasa Pinjaman', 'Jumlah SHU'];
  sheet.getRange(baseRow, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#0E4D3C').setFontColor('#ffffff');

  let totalSemuaSHU = 0;
  const baris = daftar.map(function(r, i) {
    totalSemuaSHU += r.total;
    return [i + 1, r.nama, r.simpanan, r.jasa, r.shuSimpanan, r.shuJasa, r.total];
  });
  if (baris.length > 0) {
    sheet.getRange(baseRow + 1, 1, baris.length, header.length).setValues(amanBaris_(baris));
    sheet.getRange(baseRow + 1, 3, baris.length, 5).setNumberFormat('Rp #,##0');
    const totalRow = baseRow + baris.length + 1;
    sheet.getRange(totalRow, 2).setValue('JUMLAH').setFontWeight('bold');
    sheet.getRange(totalRow, 7).setValue(totalSemuaSHU).setFontWeight('bold')
      .setNumberFormat('Rp #,##0');
  }
  for (let c = 1; c <= header.length; c++) sheet.autoResizeColumn(c);

  tampilkanPesan('✅ Laporan SHU tahun ' + tahun + ' selesai.\n' +
    'Dana anggota: ' + formatRupiah(alokasi.danaAnggota) + '\n' +
    'Total terbagi ke ' + baris.length + ' anggota: ' + formatRupiah(totalSemuaSHU) +
    '\n(Selisih kecil adalah pembulatan)\n\nCek sheet lap_shu.');
  return { alokasi: alokasi, totalTerbagi: totalSemuaSHU };
}

// ============================================================
// LAPORAN 4: REKAP SIMPANAN → sheet rekap_simpanan
// ============================================================

function generateRekapSimpanan(tahun) {
  tahun = tahun || getTahunAktif();
  const anggotaList = sheetToObjects(SHEET.ANGGOTA);

  const sheet = getSheet(SHEET.REKAP_SIMPANAN);
  sheet.clear();
  sheet.getRange('A1').setValue('💰 REKAP SIMPANAN — Tahun ' + tahun)
    .setFontWeight('bold').setFontSize(13);

  const header = ['id_anggota', 'nama', 'status', 'Pokok', 'Wajib', 'Sukarela', 'TOTAL'];
  sheet.getRange(3, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');

  const baris = anggotaList.map(function(a) {
    const pokok = getSaldoSimpanan(a.id_anggota, 'pokok');
    const wajib = getSaldoSimpanan(a.id_anggota, 'wajib');
    const sukarela = getSaldoSimpanan(a.id_anggota, 'sukarela');
    return [a.id_anggota, a.nama, a.status, pokok, wajib, sukarela,
            pokok + wajib + sukarela];
  });

  if (baris.length > 0) {
    sheet.getRange(4, 1, baris.length, header.length).setValues(amanBaris_(baris));
    sheet.getRange(4, 4, baris.length, 4).setNumberFormat('Rp #,##0');
  }
  for (let c = 1; c <= header.length; c++) sheet.autoResizeColumn(c);

  tampilkanPesan('✅ Rekap simpanan selesai — ' + baris.length + ' anggota. Cek sheet rekap_simpanan.');
}

// ============================================================
// LAPORAN 5: REKAP PINJAMAN → sheet rekap_pinjaman
// ============================================================

function generateRekapPinjaman(tahun) {
  tahun = tahun || getTahunAktif();
  const pinjamanList = getRowsByFilter(SHEET.PINJAMAN, function(row) {
    return row.status_lock !== STATUS_LOCK.VOID;
  });

  const sheet = getSheet(SHEET.REKAP_PINJAMAN);
  sheet.clear();
  sheet.getRange('A1').setValue('🏦 REKAP PINJAMAN — per ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMMM yyyy'))
    .setFontWeight('bold').setFontSize(13);

  const header = ['id_pinjaman', 'nama', 'tanggal', 'nominal', 'tenor',
                  'total_tagihan', 'angsuran/bln', 'SISA', 'status'];
  sheet.getRange(3, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');

  const baris = pinjamanList.map(function(p) {
    const anggota = getAnggota(p.id_anggota);
    return [p.id_pinjaman, anggota ? anggota.nama : p.id_anggota, p.tanggal,
            p.nominal, p.tenor, p.total_tagihan, p.angsuran_perbulan,
            getSisaPinjaman(p.id_pinjaman), p.status];
  });

  if (baris.length > 0) {
    sheet.getRange(4, 1, baris.length, header.length).setValues(amanBaris_(baris));
    sheet.getRange(4, 4, baris.length, 1).setNumberFormat('Rp #,##0');
    sheet.getRange(4, 6, baris.length, 3).setNumberFormat('Rp #,##0');
  }
  for (let c = 1; c <= header.length; c++) sheet.autoResizeColumn(c);

  tampilkanPesan('✅ Rekap pinjaman selesai — ' + baris.length + ' pinjaman. Cek sheet rekap_pinjaman.');
}

// ============================================================
// PEMBUNGKUS UNTUK MENU
// ============================================================

function menuLaporanLabaRugi() { generateLaporanLabaRugi(); }
function menuLaporanNeraca()   { generateLaporanNeraca(); }
function menuLaporanSHU()      { generateLaporanSHU(); }
function menuRekapSimpanan()   { generateRekapSimpanan(); }
function menuRekapPinjaman()   { generateRekapPinjaman(); }

/** Generate SEMUA laporan sekaligus (praktis untuk persiapan RAT). */
function menuGenerateSemuaLaporan() {
  MODE_SENYAP = true;  // supaya tidak muncul 5 popup berturut-turut
  generateLaporanLabaRugi();
  generateLaporanNeraca();
  generateLaporanSHU();
  generateRekapSimpanan();
  generateRekapPinjaman();
  MODE_SENYAP = false;
  tampilkanPesan('✅ SEMUA laporan selesai digenerate:\n' +
    '• lap_labarugi\n• lap_neraca\n• lap_shu\n• rekap_simpanan\n• rekap_pinjaman');
}