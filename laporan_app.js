/**
 * ============================================================
 * LAPORAN_APP.GS — Data Laporan untuk Ditampilkan & Dicetak di Aplikasi
 * Koperasi App
 * ============================================================
 *
 * Mengikuti format Excel pembukuan:
 *   - Buku Kas Bulanan   ("KAS (bulan)")
 *   - Daftar Rugi Laba   ("Rugi Laba")
 *   - Neraca Komparatif  ("NERACA KOMPARATIF": tahun lalu vs tahun ini)
 *   - Penjelasan Neraca  ("Penjelasan Neraca")
 *   - Daftar SHU         ("SHU")
 * Semua hanya MEMBACA data. Kop & tanda tangan diambil dari
 * Identitas Koperasi (profil_koperasi.js).
 * ============================================================
 */

function tglTtd(tanggal) {
  const t = new Date(tanggal);
  return t.getDate() + ' ' + NAMA_BULAN_PANJANG[t.getMonth()] + ' ' + t.getFullYear();
}

// ============================================================
// BUKU KAS BULANAN
// ============================================================

function hitungBukuKas(tahun, bulan) {
  tahun = Number(tahun); bulan = Number(bulan);
  const awal = new Date(tahun, bulan - 1, 1);
  const akhir = new Date(tahun, bulan, 0, 23, 59, 59, 999);
  const saldoAwal = getSaldoKas(new Date(awal.getTime() - 1));

  const cari = function(sheet, kolom) {
    const peta = {};
    sheetToObjects(sheet).forEach(function(r) { peta[String(r[kolom])] = r; });
    return peta;
  };
  const simpanan = cari(SHEET.SIMPANAN, 'id_transaksi');
  const ambil = cari(SHEET.PENGAMBILAN, 'id_pengambilan');
  const angsuran = cari(SHEET.ANGSURAN, 'id_angsuran');
  const pinjaman = cari(SHEET.PINJAMAN, 'id_pinjaman');
  const kategoriBeban = {};
  AKUN_BEBAN.forEach(function(a) { kategoriBeban[a.kode] = a; });

  // bagian → { label → { masuk, keluar } } (urutan mengikuti Excel)
  const bagian = [
    ['setor', 'Setoran Simpanan', ['Simpanan Pokok', 'Simpanan Wajib', 'Simpanan Sukarela']],
    ['ambil', 'Pengambilan Simpanan', ['Simpanan Pokok', 'Simpanan Wajib', 'Simpanan Sukarela']],
    ['pinjam', 'Pinjaman/Piutang', ['Pemberian Pinjaman/Piutang', 'Setoran Angsuran Piutang']],
    ['jasa', 'Jasa Simpan Pinjam', []],
    ['lainMasuk', 'Penghasilan Lain-lain', []],
    ['dana', 'Pengeluaran Dana-Dana', AKUN_BEBAN.filter(function(a) { return /^21/.test(a.kode); })
      .map(function(a) { return a.label; })],
    ['honor', 'Pembayaran Honor Pengurus', []],
    ['thr', 'Pembayaran THR', []],
    ['inventaris', 'Inventaris', []],
    ['umum', 'Pengeluaran Biaya Umum', AKUN_BEBAN.filter(function(a) { return a.grup === 'Biaya-biaya Umum'; })
      .map(function(a) { return a.label; })],
    ['lain', 'Lain-lain', []]
  ];
  const isi = {};
  bagian.forEach(function(b) {
    isi[b[0]] = {};
    b[2].forEach(function(l) { isi[b[0]][l] = { masuk: 0, keluar: 0 }; });
  });
  const tambah = function(kunci, label, masuk, keluar) {
    const g = isi[kunci];
    if (!g[label]) g[label] = { masuk: 0, keluar: 0 };
    g[label].masuk += masuk; g[label].keluar += keluar;
  };
  const labelJenis = function(j) {
    return j === 'pokok' ? 'Simpanan Pokok' : j === 'wajib' ? 'Simpanan Wajib' : 'Simpanan Sukarela';
  };

  sheetToObjects(SHEET.KAS).forEach(function(k) {
    if (k.status_lock === STATUS_LOCK.VOID) return;
    const t = new Date(k.tanggal);
    if (!(t >= awal && t <= akhir)) return;
    const masuk = Number(k.masuk) || 0, keluar = Number(k.keluar) || 0;
    const ref = String(k.referensi || ''), kat = String(k.kategori || '');

    if (kat === 'setoran' && simpanan[ref]) {
      tambah('setor', labelJenis(simpanan[ref].jenis_simpanan), masuk, keluar);
    } else if (kat === 'pengambilan' && ambil[ref]) {
      tambah('ambil', labelJenis(ambil[ref].jenis_simpanan || 'sukarela'), masuk, keluar);
    } else if (kat === 'pinjaman' || pinjaman[ref]) {
      tambah('pinjam', 'Pemberian Pinjaman/Piutang', masuk, keluar);
    } else if (kat === 'angsuran' && angsuran[ref]) {
      const a = angsuran[ref];
      const pokok = Math.min(masuk, Number(a.angsuran_pokok) || 0);
      tambah('pinjam', 'Setoran Angsuran Piutang', pokok, 0);
      tambah('jasa', 'Jasa Simpan Pinjam', masuk - pokok, 0);
    } else if (kat === 'pendapatan_lain') {
      tambah('lainMasuk', 'Penghasilan Lain-lain', masuk, keluar);
    } else if (kat === 'operasional' && kategoriBeban[ref]) {
      const a = kategoriBeban[ref];
      const kunci = /^21/.test(a.kode) ? 'dana' : a.kode === '509' ? 'honor' : a.kode === '506' ? 'thr'
        : a.kode === '111' ? 'inventaris' : 'umum';
      tambah(kunci, kunci === 'honor' || kunci === 'thr' || kunci === 'inventaris' ? a.grup : a.label, masuk, keluar);
    } else if (kat === 'operasional') {
      tambah('umum', namaAkun(ref) || 'Biaya lainnya', masuk, keluar);  // akun beban versi lama
    } else {
      tambah('lain', String(k.keterangan || kat || 'Lain-lain').slice(0, 60), masuk, keluar);
    }
  });

  let totMasuk = 0, totKeluar = 0;
  const hasilBagian = bagian.map(function(b, i) {
    const items = Object.keys(isi[b[0]]).map(function(l) {
      const v = isi[b[0]][l];
      totMasuk += v.masuk; totKeluar += v.keluar;
      return { label: l, masuk: v.masuk, keluar: v.keluar };
    });
    // Bagian tanpa rincian: satu baris dengan label bagian
    const tunggal = b[2].length === 0;
    const jumlah = items.reduce(function(s, x) { return { masuk: s.masuk + x.masuk, keluar: s.keluar + x.keluar }; },
      { masuk: 0, keluar: 0 });
    return { no: i + 2, judul: b[1], items: tunggal ? [] : items,
      masuk: tunggal ? jumlah.masuk : null, keluar: tunggal ? jumlah.keluar : null };
  });

  const saldoAkhir = saldoAwal + totMasuk - totKeluar;
  return {
    tahun: tahun, bulan: bulan, namaBulan: NAMA_BULAN_PANJANG[bulan - 1],
    saldoAwal: saldoAwal, bagian: hasilBagian,
    totMasuk: totMasuk, totKeluar: totKeluar, saldoAkhir: saldoAkhir,
    cocok: Math.abs(saldoAkhir - getSaldoKas(akhir)) < 1,
    tanggalTtd: tglTtd(new Date(tahun, bulan, 0))
  };
}

// ============================================================
// PENJELASAN NERACA
// ============================================================

function hitungPenjelasanNeraca(tahun) {
  tahun = Number(tahun);
  const n = hitungNeraca(tahun), lalu = hitungNeraca(tahun - 1);
  const akhirLalu = new Date(tahun - 1, 11, 31, 23, 59, 59, 999);
  const akhirIni = new Date(tahun, 11, 31, 23, 59, 59, 999);

  let penerimaan = 0, pengeluaran = 0;
  sheetToObjects(SHEET.KAS).forEach(function(k) {
    if (k.status_lock === STATUS_LOCK.VOID) return;
    const t = new Date(k.tanggal);
    if (t > akhirLalu && t <= akhirIni) {
      penerimaan += Number(k.masuk) || 0; pengeluaran += Number(k.keluar) || 0;
    }
  });
  let pemberian = 0, angsuranPokok = 0;
  sheetToObjects(SHEET.JURNAL).forEach(function(j) {
    if (Number(j.tahun) !== tahun || String(j.kode_akun) !== AKUN.PIUTANG_PINJAMAN || isJurnalPenutup(j)) return;
    pemberian += Number(j.debit) || 0; angsuranPokok += Number(j.kredit) || 0;
  });
  const nilai = function(daftar, kode) {
    const r = daftar.find(function(x) { return x.kode === kode; });
    return r ? r.nilai : 0;
  };
  return {
    tahun: tahun,
    kas: { saldoAwal: getSaldoKas(akhirLalu), penerimaan: penerimaan, pengeluaran: pengeluaran,
      sisa: nilai(n.aktivaLancar, '101') + nilai(n.aktivaLancar, '102') },
    piutang: { tahunLalu: nilai(lalu.aktivaLancar, '103'), pemberian: pemberian,
      angsuran: angsuranPokok, sisa: nilai(n.aktivaLancar, '103') },
    inventaris: n.aktivaTetap,
    sukarela: nilai(n.hutangLancar, '203'),
    danaDana: n.danaDana,
    pokok: nilai(n.modalSendiri, '201'),
    wajib: nilai(n.modalSendiri, '202'),
    cadangan: nilai(n.modalSendiri, '304'),
    shu: n.shuBerjalan,
    neraca: n,
    tanggalTtd: tglTtd(new Date(tahun, 11, 31))
  };
}

// ============================================================
// API
// ============================================================

function tahunLaporan(f) {
  return Number(f && f.tahun) || getTahunAktif() || new Date().getFullYear();
}

function apiLapBukuKas(f) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const h = hitungBukuKas(tahunLaporan(f), Number(f && f.bulan) || new Date().getMonth() + 1);
  h.profil = getProfilKoperasi();
  return h;
}

function apiLapLabaRugi(f) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const tahun = tahunLaporan(f);
  const h = hitungLabaRugi(tahun);
  h.profil = getProfilKoperasi();
  h.tanggalTtd = tglTtd(new Date(tahun, 11, 31));
  return h;
}

function apiLapNeraca(f) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const tahun = tahunLaporan(f);
  return { tahun: tahun, ini: hitungNeraca(tahun), lalu: hitungNeraca(tahun - 1),
    profil: getProfilKoperasi(), tanggalTtd: tglTtd(new Date(tahun, 11, 31)) };
}

function apiLapPenjelasan(f) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const h = hitungPenjelasanNeraca(tahunLaporan(f));
  h.profil = getProfilKoperasi();
  return h;
}

function apiLapSHU(f) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const tahun = tahunLaporan(f);
  const setting = getSettingRAT(tahun);
  if (!setting) throw new Error('Setting tahun ' + tahun + ' tidak ditemukan.');
  return { tahun: tahun, alokasi: hitungAlokasiDana(tahun), baris: daftarSHUAnggota(tahun),
    shuSimpananPersen: Number(setting.shu_simpanan), shuJasaPersen: Number(setting.shu_jasa),
    profil: getProfilKoperasi(), tanggalTtd: tglTtd(new Date(tahun, 11, 31)) };
}
