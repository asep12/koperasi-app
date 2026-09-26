/**
 * ============================================================
 * AKUNTANSI.GS — Bagan Akun, Alokasi SHU, Laba Rugi & Neraca
 * Koperasi App
 * ============================================================
 *
 * Mengikuti susunan Excel "Pembukuan Koperasi" (Rugi Laba,
 * Penjelasan Neraca, Neraca Komparatif):
 *
 * LABA RUGI
 *   Pendapatan : Jasa Simpan Pinjam (401), Penghasilan Lain-lain (403)
 *   Biaya      : Biaya-biaya Umum (5xx, termasuk jasa simpanan
 *                sukarela 510), Honor Pengurus (509), THR (506)
 *   SHU belum pajak → pajak (informasi) → alokasi 7 pos dari SHU
 *   belum pajak (Cadangan, Dana Anggota, Pengurus, Kesejahteraan
 *   Pegawai, Pendidikan, Sosial, Pembangunan Daerah Kerja).
 *
 * NERACA (kumulatif dari jurnal lintas tahun + saldo awal migrasi)
 *   Aktiva Lancar : Kas, Piutang
 *   Aktiva Tetap  : Inventaris, Akumulasi Penyusutan
 *   Hutang Lancar : Simpanan Sukarela, Dana-dana
 *   Modal Sendiri : Simpanan Pokok, Simpanan Wajib, Cadangan, ...
 *   SHU tahun berjalan
 *
 * TUTUP BUKU membuat jurnal penutup (referensi TUTUP-BUKU-tahun):
 * akun 4xx/5xx dinolkan, SHU dipindah ke Cadangan (304) & dana-dana
 * (211–216) sesuai persen alokasi. Tidak ada lagi salinan saldo
 * anggota ke saldo_awal (dulu membuat simpanan terhitung dobel).
 * ============================================================
 */

// [kunci setting, label, kode akun tujuan, persen bawaan (Excel)]
const POS_ALOKASI = [
  ['dana_cadangan',      'Cadangan',                      '304', 25],
  ['dana_anggota',       'Dana Anggota',                  '211', 50],
  ['dana_pengurus',      'Dana Pengurus',                 '212', 10],
  ['dana_kesejahteraan', 'Dana Kesejahteraan Pegawai',    '213', 5],
  ['dana_pendidikan',    'Dana Pendidikan',               '214', 5],
  ['dana_sosial',        'Dana Sosial',                   '215', 2.5],
  ['dana_pembangunan',   'Dana Pembangunan Daerah Kerja', '216', 2.5]
];
const PAJAK_PERSEN_BAWAAN = 0.5;
const PREFIX_TUTUP_BUKU = 'TUTUP-BUKU-';

// Akun yang wajib ada (ditambahkan otomatis ke coa_akun bila belum ada)
const COA_WAJIB = [
  ['101', 'Kas', 'Aktiva'],
  ['102', 'Bank', 'Aktiva'],
  ['103', 'Piutang Pinjaman', 'Aktiva'],
  ['111', 'Inventaris', 'Aktiva'],
  ['112', 'Akumulasi Penyusutan Inventaris', 'Aktiva'],
  ['201', 'Simpanan Pokok', 'Kewajiban'],
  ['202', 'Simpanan Wajib', 'Kewajiban'],
  ['203', 'Simpanan Sukarela', 'Kewajiban'],
  ['211', 'Dana Anggota', 'Kewajiban'],
  ['212', 'Dana Pengurus', 'Kewajiban'],
  ['213', 'Dana Kesejahteraan Pegawai', 'Kewajiban'],
  ['214', 'Dana Pendidikan', 'Kewajiban'],
  ['215', 'Dana Sosial', 'Kewajiban'],
  ['216', 'Dana Pembangunan Daerah Kerja', 'Kewajiban'],
  ['301', 'Modal Koperasi', 'Ekuitas'],
  ['302', 'SHU Tahun Berjalan', 'Ekuitas'],
  ['303', 'SHU Ditahan', 'Ekuitas'],
  ['304', 'Cadangan', 'Ekuitas'],
  ['401', 'Jasa Simpan Pinjam', 'Pendapatan'],
  ['402', 'Pendapatan Jasa Sukarela (lama)', 'Pendapatan'],
  ['403', 'Penghasilan Lain-lain', 'Pendapatan'],
  ['501', 'Biaya Operasional', 'Beban'],
  ['505', 'Transport Kesejahteraan', 'Beban'],
  ['506', 'THR', 'Beban'],
  ['507', 'Uang Duduk', 'Beban'],
  ['508', 'Souvenir RAT', 'Beban'],
  ['509', 'Honor Pengurus', 'Beban'],
  ['510', 'Jasa Simpanan Sukarela', 'Beban'],
  ['511', 'Brosur RAT', 'Beban'],
  ['512', 'Administrasi/ATK', 'Beban'],
  ['513', 'Konsumsi RAT', 'Beban'],
  ['514', 'Pengolahan Neraca', 'Beban'],
  ['515', 'Rapat Pengurus', 'Beban'],
  ['516', 'Door Prize', 'Beban'],
  ['517', 'Transport Belanja', 'Beban'],
  ['518', 'Honor BP & Pembina', 'Beban'],
  ['519', 'Beban Penyusutan Inventaris', 'Beban'],
  ['520', 'Biaya Lain-lain', 'Beban']
];

/** Tambahkan akun yang belum ada di coa_akun (akun lama tidak diubah). */
function pastikanCOA() {
  const ada = {};
  sheetToObjects(SHEET.COA_AKUN).forEach(function(a) { ada[String(a.kode)] = true; });
  const kurang = COA_WAJIB.filter(function(a) { return !ada[a[0]]; });
  if (kurang.length === 0) return 0;
  appendRowsFromObjects(SHEET.COA_AKUN, kurang.map(function(a) {
    return { kode: a[0], nama_akun: a[1], jenis: a[2] };
  }));
  return kurang.length;
}

function namaAkun(kode) {
  const a = sheetToObjects(SHEET.COA_AKUN).find(function(x) { return String(x.kode) === String(kode); });
  if (a) return String(a.nama_akun);
  const w = COA_WAJIB.find(function(x) { return x[0] === String(kode); });
  return w ? w[1] : 'Akun ' + kode;
}

// ============================================================
// SETTING ALOKASI (kolom baru di setting_rat)
// ============================================================

const KOLOM_ALOKASI_BARU = ['dana_anggota', 'dana_kesejahteraan', 'dana_pendidikan',
  'dana_pembangunan', 'pajak_persen'];

/**
 * Sheet lama (3 pos dana) → 7 pos seperti Excel. Saat kolom baru
 * ditambahkan, SEMUA persen alokasi diisi nilai Excel (25/50/10/5/5/2,5/2,5)
 * dan pajak 0,5% — model lama (dana dipotong lalu sisa ke anggota)
 * tidak sebanding. Periksa kembali di halaman Tahun Buku.
 * @return {boolean} true jika baru dimigrasi
 */
function pastikanKolomAlokasi() {
  const sheet = getSheet(SHEET.SETTING_RAT);
  const header = getHeader(SHEET.SETTING_RAT).map(function(h) { return String(h).trim(); });
  const kurang = KOLOM_ALOKASI_BARU.filter(function(k) { return header.indexOf(k) === -1; });
  if (kurang.length === 0) return false;

  sheet.getRange(1, header.length + 1, 1, kurang.length).setValues([kurang]).setFontWeight('bold');
  hapusCache(SHEET.SETTING_RAT);
  sheetToObjects(SHEET.SETTING_RAT).forEach(function(row) {
    const isi = { pajak_persen: PAJAK_PERSEN_BAWAAN };
    POS_ALOKASI.forEach(function(p) { isi[p[0]] = p[3]; });
    updateRowByRowNumber(SHEET.SETTING_RAT, row.__row, isi);
  });
  logAktivitas('EDIT', SHEET.SETTING_RAT, 'alokasi-7-pos', null,
    { keterangan: 'Kolom alokasi 7 pos + pajak ditambahkan, diisi nilai Excel' });
  return true;
}

// ============================================================
// SALDO PER AKUN
// ============================================================

function isJurnalPenutup(row) {
  return String(row.referensi || '').indexOf(PREFIX_TUTUP_BUKU) === 0;
}

/**
 * Total debit & kredit per akun untuk SATU tahun buku.
 * Jurnal penutup (tutup buku) tidak ikut — supaya laporan tahun yang
 * sudah ditutup tetap memperlihatkan pendapatan & biaya aslinya.
 * @return {Object} kode → { debit, kredit }
 */
function getSaldoPerAkun(tahun) {
  return cacheTurunan('saldoAkun_' + tahun, function() {
    const peta = {};
    sheetToObjects(SHEET.JURNAL).forEach(function(row) {
      if (Number(row.tahun) !== Number(tahun) || isJurnalPenutup(row)) return;
      const kode = String(row.kode_akun);
      if (!peta[kode]) peta[kode] = { debit: 0, kredit: 0 };
      peta[kode].debit += Number(row.debit) || 0;
      peta[kode].kredit += Number(row.kredit) || 0;
    });
    return peta;
  });
}

/** Pendapatan − Biaya satu tahun (SHU belum pajak). */
function getTotalSHUTahun(tahun) {
  const saldo = getSaldoPerAkun(tahun);
  let shu = 0;
  Object.keys(saldo).forEach(function(kode) {
    if (/^[45]/.test(kode)) shu += saldo[kode].kredit - saldo[kode].debit;
  });
  return shu;
}

// ============================================================
// ALOKASI SHU (7 pos + pajak)
// ============================================================

/**
 * @return {Object} { pendapatan, biaya, shuKotor, pajakPersen, pajak, shuSetelahPajak,
 *   pos: [{kunci,label,akun,persen,nominal}], danaAnggota, shuNetto (= dana anggota),
 *   danaCadangan, danaPengurus, danaSosial }
 */
function hitungAlokasiDana(tahun) {
  const setting = getSettingRAT(tahun);
  if (!setting) throw new Error('Setting RAT tahun ' + tahun + ' tidak ditemukan.');
  validateAlokasiSHU(setting);

  const saldo = getSaldoPerAkun(tahun);
  let pendapatan = 0, biaya = 0;
  Object.keys(saldo).forEach(function(kode) {
    const bersih = saldo[kode].kredit - saldo[kode].debit;
    if (kode.charAt(0) === '4') pendapatan += bersih;
    if (kode.charAt(0) === '5') biaya -= bersih;
  });
  const shu = pendapatan - biaya;
  const pajakPersen = Number(setting.pajak_persen) || 0;
  // Sesuai Excel: pajak dihitung dari pendapatan & ditampilkan;
  // alokasi dibagi dari SHU BELUM pajak.
  const pajak = Math.round(pendapatan * pajakPersen / 100);

  let sisa = shu > 0 ? shu : 0;
  const pos = POS_ALOKASI.map(function(p, i) {
    const persen = Number(setting[p[0]]) || 0;
    const nominal = shu <= 0 ? 0
      : i === POS_ALOKASI.length - 1 ? sisa           // sisa pembulatan ke pos terakhir
      : Math.round(shu * persen / 100);
    sisa -= nominal;
    return { kunci: p[0], label: p[1], akun: p[2], persen: persen, nominal: nominal };
  });
  const cari = function(k) { return pos.find(function(p) { return p.kunci === k; }).nominal; };

  return {
    pendapatan: pendapatan, biaya: biaya,
    shuKotor: shu, pajakPersen: pajakPersen, pajak: pajak, shuSetelahPajak: shu - pajak,
    pos: pos,
    danaAnggota: cari('dana_anggota'),
    shuNetto: cari('dana_anggota'),           // basis SHU per anggota
    danaCadangan: cari('dana_cadangan'),
    danaPengurus: cari('dana_pengurus'),
    danaSosial: cari('dana_sosial')
  };
}

// ============================================================
// LABA RUGI (susunan Excel "Daftar Rugi Laba")
// ============================================================

function hitungLabaRugi(tahun) {
  const saldo = getSaldoPerAkun(tahun);
  const bersih = function(kode) {
    const s = saldo[kode];
    return s ? s.kredit - s.debit : 0;
  };

  const pendapatan = [
    { kode: '401', label: 'Jasa Simpan Pinjam', nilai: bersih('401') }
  ];
  let lain = 0;
  Object.keys(saldo).forEach(function(k) {
    if (k.charAt(0) === '4' && k !== '401' && k !== '402') lain += bersih(k);
  });
  pendapatan.push({ kode: '403', label: 'Penghasilan Lain-lain', nilai: lain });

  // Biaya-biaya umum: semua 5xx selain Honor Pengurus & THR,
  // + akun 402 lama (dulu jasa sukarela didebet ke 402).
  const rincianUmum = [];
  Object.keys(saldo).sort().forEach(function(k) {
    if (k.charAt(0) !== '5' || k === '509' || k === '506') return;
    const nilai = -bersih(k);
    if (nilai !== 0) rincianUmum.push({ kode: k, label: namaAkun(k), nilai: nilai });
  });
  if (bersih('402') < 0) {
    rincianUmum.push({ kode: '402', label: 'Jasa Simpanan Sukarela (akun lama)', nilai: -bersih('402') });
  } else if (bersih('402') > 0) {
    pendapatan[1].nilai += bersih('402');
  }
  const totalUmum = rincianUmum.reduce(function(s, r) { return s + r.nilai; }, 0);
  const biaya = [
    { label: 'Biaya-biaya Umum', nilai: totalUmum, rincian: rincianUmum },
    { kode: '509', label: 'Honor Pengurus', nilai: -bersih('509') },
    { kode: '506', label: 'THR', nilai: -bersih('506') }
  ];

  const totalPendapatan = pendapatan.reduce(function(s, r) { return s + r.nilai; }, 0);
  const totalBiaya = biaya.reduce(function(s, r) { return s + r.nilai; }, 0);
  const setting = getSettingRAT(tahun);
  const alokasi = setting ? hitungAlokasiDana(tahun) : null;

  return {
    tahun: Number(tahun),
    pendapatan: pendapatan, totalPendapatan: totalPendapatan,
    biaya: biaya, totalBiaya: totalBiaya,
    shu: totalPendapatan - totalBiaya,
    pajakPersen: alokasi ? alokasi.pajakPersen : 0,
    pajak: alokasi ? alokasi.pajak : 0,
    shuSetelahPajak: totalPendapatan - totalBiaya - (alokasi ? alokasi.pajak : 0),
    alokasi: alokasi ? alokasi.pos : []
  };
}

// ============================================================
// NERACA (kumulatif, susunan Excel "Neraca Komparatif")
// ============================================================

/** Baris saldo_awal hasil tutup buku versi lama (salinan saldo) — diabaikan. */
function isSaldoAwalSalinan(row) {
  return String(row.keterangan || '').indexOf('Saldo akhir tutup buku') === 0;
}

// jenis saldo_awal → [kode akun, +1 = sisi aktiva / −1 = sisi pasiva]
const AKUN_SALDO_AWAL = {
  kas: ['101', 1], piutang: ['103', 1],
  inventaris: ['111', 1], penyusutan: ['112', -1],
  simpanan_pokok: ['201', -1], simpanan_wajib: ['202', -1],
  simpanan_sukarela: ['203', -1],
  dana_anggota: ['211', -1], dana_pengurus: ['212', -1], dana_kesejahteraan: ['213', -1],
  dana_pendidikan: ['214', -1], dana_sosial: ['215', -1], dana_pembangunan: ['216', -1],
  modal: ['301', -1], shu: ['303', -1], cadangan: ['304', -1]
};

/**
 * Cek keseimbangan saldo awal migrasi: sisi aktiva harus = sisi pasiva.
 * @return {Object} { aktiva, pasiva, selisih, balance }
 */
function cekSaldoAwal() {
  let aktiva = 0, pasiva = 0;
  sheetToObjects(SHEET.SALDO_AWAL).forEach(function(r) {
    if (isSaldoAwalSalinan(r)) return;
    const map = AKUN_SALDO_AWAL[String(r.jenis)];
    if (!map) return;
    const n = Number(r.nominal) || 0;
    if (map[1] > 0) aktiva += n; else pasiva += n;
  });
  return { aktiva: aktiva, pasiva: pasiva, selisih: aktiva - pasiva,
    balance: Math.abs(aktiva - pasiva) < 1 };
}

/**
 * Posisi per 31 Desember tahun tertentu.
 * saldo akun riil (1xx–3xx) = saldo awal migrasi + jurnal s/d tahun itu;
 * SHU tahun berjalan = pendapatan − biaya tahun itu (sebelum jurnal penutup).
 */
function hitungNeraca(tahun) {
  tahun = Number(tahun);
  return cacheTurunan('neraca_' + tahun, function() {
    const d = {};         // kode → debit − kredit (akun riil, kumulatif)
    let shuBerjalan = 0;  // tahun ini
    let shuLalu = 0;      // tahun-tahun lalu yang belum ditutup (data lama)
    const tambah = function(kode, nilai) { d[kode] = (d[kode] || 0) + nilai; };

    sheetToObjects(SHEET.SALDO_AWAL).forEach(function(r) {
      if (isSaldoAwalSalinan(r) || Number(r.tahun) > tahun) return;
      const map = AKUN_SALDO_AWAL[String(r.jenis)];
      if (map) tambah(map[0], map[1] * (Number(r.nominal) || 0));
    });

    sheetToObjects(SHEET.JURNAL).forEach(function(r) {
      const t = Number(r.tahun);
      if (t > tahun) return;
      if (t === tahun && isJurnalPenutup(r)) return;
      const kode = String(r.kode_akun), dk = (Number(r.debit) || 0) - (Number(r.kredit) || 0);
      if (/^[123]/.test(kode)) tambah(kode, dk);
      else if (t === tahun) shuBerjalan -= dk;
      else shuLalu -= dk;
    });

    const aktiva = function(k) { return d[k] || 0; };
    const pasiva = function(k) { return 0 - (d[k] || 0); };   // 0 − x: hindari −0
    const dipakai = {};
    const baris = function(kode, label, nilai) { dipakai[kode] = true; return { kode: kode, label: label, nilai: nilai }; };

    const aktivaLancar = [baris('101', 'Kas', aktiva('101'))];
    if (aktiva('102')) aktivaLancar.push(baris('102', 'Bank', aktiva('102')));
    else dipakai['102'] = true;
    aktivaLancar.push(baris('103', 'Piutang Simpan Pinjam', aktiva('103')));
    const aktivaTetap = [
      baris('111', 'Inventaris', aktiva('111')),
      baris('112', 'Penyusutan Inventaris', aktiva('112'))
    ];
    const danaDana = POS_ALOKASI.filter(function(p) { return p[2] !== '304'; })
      .map(function(p) { return baris(p[2], p[1], pasiva(p[2])); });
    const hutangLancar = [baris('203', 'Simpanan Sukarela', pasiva('203'))];
    const modalSendiri = [
      baris('201', 'Simpanan Pokok', pasiva('201')),
      baris('202', 'Simpanan Wajib', pasiva('202')),
      baris('304', 'Cadangan', pasiva('304'))
    ];
    const modalLain = pasiva('301') + pasiva('302') + pasiva('303') + shuLalu;
    dipakai['301'] = dipakai['302'] = dipakai['303'] = true;
    if (modalLain) modalSendiri.push({ kode: '303', label: 'Modal & SHU Ditahan', nilai: modalLain });

    // Akun riil lain yang tidak dikenal tetap masuk supaya neraca seimbang
    Object.keys(d).forEach(function(k) {
      if (dipakai[k] || !d[k]) return;
      if (k.charAt(0) === '1') aktivaLancar.push(baris(k, namaAkun(k), aktiva(k)));
      else if (k.charAt(0) === '2') hutangLancar.push(baris(k, namaAkun(k), pasiva(k)));
      else modalSendiri.push(baris(k, namaAkun(k), pasiva(k)));
    });

    const jumlah = function(a) { return a.reduce(function(s, r) { return s + r.nilai; }, 0); };
    const totalAktiva = jumlah(aktivaLancar) + jumlah(aktivaTetap);
    const totalPasiva = jumlah(hutangLancar) + jumlah(danaDana) + jumlah(modalSendiri) + shuBerjalan;
    return {
      tahun: tahun,
      aktivaLancar: aktivaLancar, aktivaTetap: aktivaTetap,
      hutangLancar: hutangLancar, danaDana: danaDana, modalSendiri: modalSendiri,
      shuBerjalan: shuBerjalan,
      totalAktiva: totalAktiva, totalPasiva: totalPasiva,
      selisih: totalAktiva - totalPasiva,
      balance: Math.abs(totalAktiva - totalPasiva) < 1
    };
  });
}

// ============================================================
// JURNAL PENUTUP (dipanggil tutup buku)
// ============================================================

/**
 * Nolkan akun pendapatan & biaya tahun lama; SHU dipindah ke
 * Cadangan & dana-dana sesuai alokasi. Rugi → mengurangi Cadangan.
 * @return {Object} alokasi yang dijurnal
 */
function buatJurnalPenutup(tahunLama) {
  pastikanCOA();
  const sudah = getRowsByFilter(SHEET.JURNAL, function(r) {
    return String(r.referensi) === PREFIX_TUTUP_BUKU + tahunLama;
  });
  if (sudah.length > 0) throw new Error('Jurnal penutup tahun ' + tahunLama + ' sudah ada.');

  const saldo = getSaldoPerAkun(tahunLama);
  const alokasi = hitungAlokasiDana(tahunLama);
  const baris = [];
  Object.keys(saldo).forEach(function(kode) {
    if (!/^[45]/.test(kode)) return;
    const dk = saldo[kode].debit - saldo[kode].kredit;
    if (dk > 0) baris.push({ kode_akun: kode, debit: 0, kredit: dk });
    if (dk < 0) baris.push({ kode_akun: kode, debit: -dk, kredit: 0 });
  });
  if (alokasi.shuKotor > 0) {
    alokasi.pos.forEach(function(p) {
      if (p.nominal > 0) baris.push({ kode_akun: p.akun, debit: 0, kredit: p.nominal });
    });
  } else if (alokasi.shuKotor < 0) {
    baris.push({ kode_akun: AKUN.CADANGAN, debit: -alokasi.shuKotor, kredit: 0 });
  }
  if (baris.length > 0) {
    catatJurnal(new Date(tahunLama, 11, 31), tahunLama, 12, PREFIX_TUTUP_BUKU + tahunLama,
      'Jurnal penutup & alokasi SHU tahun ' + tahunLama, baris);
  }
  return alokasi;
}
