/**
 * Kode_1.gs — bagian 1 dari 3 kode Koperasi App.
 * Hak Cipta (c) Asep Hanuryana (Threads: @asep94). Seluruh hak dilindungi; lihat LICENSE.
 *
 * FILE INI DIBUAT OTOMATIS oleh tools/buat_paket_manual.py. Jangan disunting di sini:
 * ubah file aslinya di repositori, lalu jalankan ulang alat tersebut.
 *
 * Berisi: akuntansi.js, anggota.js, angsuran.js, approval.js, bayar_angsuran.js, config.js, dashboard.js, database.js, fitur_lanjutan.js, jasa.js, kas_beban.js, keluar_anggota.js, koreksi_data.js, laporan.js, laporan_app.js, log_email_backup.js
 */

// ===== akuntansi.js =====
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

// ===== anggota.js =====
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

// ===== angsuran.js =====
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

// ===== approval.js =====
/**
 * ============================================================
 * APPROVAL.GS — Approve/Reject Pengambilan Simpanan Sukarela
 * Koperasi App
 * Sesuai dokumen: REKAP_Koperasi_v2.md (BAGIAN 3, 12, 17 — keputusan v2 #7)
 * ============================================================
 *
 * KEPUTUSAN v2: status_approval TIDAK BOLEH diubah manual di sel.
 * Harus lewat menu ini, supaya:
 * 1. Tercatat di log_aktivitas SIAPA yang approve/reject + kapan
 * 2. Saldo divalidasi ULANG saat approve (bukan cuma saat pengajuan
 *    — saldo bisa berubah di antara pengajuan dan approval)
 * 3. Kas keluar + jurnal hanya dibuat saat benar-benar APPROVED
 *
 * ALUR:
 * PENDING (dari inputPengajuanPengambilan di simpanan.gs)
 *    ├── APPROVED → kas keluar + jurnal + log APPROVE
 *    └── REJECTED → hanya ubah status + log REJECT (tanpa kas/jurnal)
 * ============================================================
 */

// ============================================================
// FUNGSI INTI (bisa dipanggil dari test/editor dengan parameter)
// ============================================================

/**
 * APPROVE satu pengajuan pengambilan.
 * @param {string} idPengambilan  contoh: 'AMB-2026-0001'
 * @return {Object} { pesan }
 */
function approvePengambilan(idPengambilan) {
  const row = getRowByField(SHEET.PENGAMBILAN, 'id_pengambilan', idPengambilan);
  if (!row) throw new Error('Pengajuan ' + idPengambilan + ' tidak ditemukan.');

  if (row.status_approval !== 'PENDING') {
    throw new Error('Pengajuan ' + idPengambilan + ' berstatus ' + row.status_approval +
      ' — hanya yang PENDING yang bisa di-approve.');
  }
  if (row.status_lock === STATUS_LOCK.VOID) {
    throw new Error('Pengajuan ini sudah di-VOID.');
  }

  validateTahunAktif(Number(row.tahun));
  const anggota = pastikanAnggotaAktif(row.id_anggota);

  // VALIDASI ULANG SALDO — saldo bisa berubah sejak pengajuan dibuat
  const saldoSukarela = getSaldoSimpanan(row.id_anggota, 'sukarela');
  if (Number(row.jumlah) > saldoSukarela) {
    throw new Error('Saldo sukarela ' + anggota.nama + ' saat ini (' +
      formatRupiah(saldoSukarela) + ') tidak cukup untuk pengambilan ' +
      formatRupiah(row.jumlah) + '. Pengajuan TIDAK di-approve.');
  }

  // VALIDASI KAS — kas koperasi harus cukup untuk bayar tunai
  const saldoKas = getSaldoKas();
  if (Number(row.jumlah) > saldoKas) {
    throw new Error('Saldo kas koperasi (' + formatRupiah(saldoKas) +
      ') tidak cukup untuk pengambilan ' + formatRupiah(row.jumlah) + '.');
  }

  // ---------- UBAH STATUS ----------
  updateRowByField(SHEET.PENGAMBILAN, 'id_pengambilan', idPengambilan,
    { status_approval: 'APPROVED' });

  // ---------- KAS KELUAR ----------
  const tanggal = new Date();
  catatKasKeluar({
    tanggal: tanggal,
    tahun: Number(row.tahun),
    bulan: Number(row.bulan),
    kategori: 'pengambilan',
    referensi: idPengambilan,
    keterangan: 'Pengambilan sukarela — ' + anggota.nama,
    nominal: Number(row.jumlah)
  });

  // ---------- JURNAL (BAGIAN 12 #6) ----------
  // Debit Simpanan Sukarela, Kredit Kas
  catatJurnal(tanggal, Number(row.tahun), Number(row.bulan), idPengambilan,
    'Pengambilan simpanan sukarela — ' + anggota.nama,
    [
      { kode_akun: AKUN.SIMPANAN_SUKARELA, debit: Number(row.jumlah), kredit: 0 },
      { kode_akun: AKUN.KAS, debit: 0, kredit: Number(row.jumlah) }
    ]
  );

  // ---------- LOG APPROVE ----------
  logAktivitas('APPROVE', SHEET.PENGAMBILAN, idPengambilan,
    { status_approval: 'PENDING' },
    { status_approval: 'APPROVED', approver: getUserEmail() });

  return {
    pesan: '✅ APPROVED — Pengambilan ' + formatRupiah(row.jumlah) + ' untuk ' +
      anggota.nama + ' disetujui.\nKas keluar & jurnal sudah dibuat otomatis.\n' +
      'Sisa saldo sukarela anggota: ' +
      formatRupiah(getSaldoSimpanan(row.id_anggota, 'sukarela'))
  };
}

/**
 * REJECT satu pengajuan pengambilan.
 * Tidak ada kas/jurnal — hanya ubah status + catat log.
 * @param {string} idPengambilan
 * @param {string} [alasan]
 */
function rejectPengambilan(idPengambilan, alasan) {
  const row = getRowByField(SHEET.PENGAMBILAN, 'id_pengambilan', idPengambilan);
  if (!row) throw new Error('Pengajuan ' + idPengambilan + ' tidak ditemukan.');

  if (row.status_approval !== 'PENDING') {
    throw new Error('Pengajuan ' + idPengambilan + ' berstatus ' + row.status_approval +
      ' — hanya yang PENDING yang bisa di-reject.');
  }

  updateRowByField(SHEET.PENGAMBILAN, 'id_pengambilan', idPengambilan, {
    status_approval: 'REJECTED',
    keterangan: (row.keterangan ? row.keterangan + ' | ' : '') +
      'DITOLAK: ' + (alasan || 'tanpa alasan tertulis')
  });

  logAktivitas('REJECT', SHEET.PENGAMBILAN, idPengambilan,
    { status_approval: 'PENDING' },
    { status_approval: 'REJECTED', approver: getUserEmail(), alasan: alasan || '' });

  const anggota = getAnggota(row.id_anggota);
  return {
    pesan: '🚫 REJECTED — Pengajuan ' + formatRupiah(row.jumlah) + ' untuk ' +
      (anggota ? anggota.nama : row.id_anggota) + ' ditolak.'
  };
}

/** Ambil daftar semua pengajuan yang masih PENDING. */
function getDaftarPending() {
  return getRowsByFilter(SHEET.PENGAMBILAN, function(row) {
    return row.status_approval === 'PENDING' && row.status_lock !== STATUS_LOCK.VOID;
  });
}

// ============================================================
// MENU INTERAKTIF (dipanggil dari menu spreadsheet)
// ============================================================

/**
 * Menu: tampilkan daftar PENDING, minta admin pilih ID,
 * lalu pilih Approve / Reject.
 */
function menuApprovalPengambilan() {
  const ui = SpreadsheetApp.getUi();
  const pendingList = getDaftarPending();

  if (pendingList.length === 0) {
    ui.alert('Tidak ada pengajuan pengambilan yang berstatus PENDING. ✅');
    return;
  }

  // Susun daftar untuk ditampilkan
  const daftar = pendingList.map(function(row, i) {
    const anggota = getAnggota(row.id_anggota);
    return (i + 1) + '. ' + row.id_pengambilan + ' — ' +
      (anggota ? anggota.nama : row.id_anggota) + ' — ' +
      formatRupiah(row.jumlah) +
      (row.keterangan ? ' (' + row.keterangan + ')' : '');
  }).join('\n');

  const respId = ui.prompt(
    '📋 Pengajuan PENDING (' + pendingList.length + ')',
    daftar + '\n\nKetik ID pengajuan yang mau diproses (contoh: ' +
    pendingList[0].id_pengambilan + '):',
    ui.ButtonSet.OK_CANCEL
  );
  if (respId.getSelectedButton() !== ui.Button.OK) return;

  const idPilihan = respId.getResponseText().trim();
  const target = pendingList.find(function(r) { return r.id_pengambilan === idPilihan; });
  if (!target) {
    ui.alert('ID "' + idPilihan + '" tidak ada di daftar PENDING. Proses dibatalkan.');
    return;
  }

  const anggota = getAnggota(target.id_anggota);
  const respAksi = ui.alert(
    'Proses ' + idPilihan,
    'Pengambilan ' + formatRupiah(target.jumlah) + ' oleh ' +
    (anggota ? anggota.nama : target.id_anggota) + '\n\n' +
    'YES = APPROVE (setujui, kas keluar)\n' +
    'NO  = REJECT (tolak)\n' +
    'CANCEL = batal, tidak melakukan apa-apa',
    ui.ButtonSet.YES_NO_CANCEL
  );

  try {
    if (respAksi === ui.Button.YES) {
      const hasil = approvePengambilan(idPilihan);
      ui.alert(hasil.pesan);
    } else if (respAksi === ui.Button.NO) {
      const respAlasan = ui.prompt('Alasan penolakan (opsional):', ui.ButtonSet.OK);
      const hasil = rejectPengambilan(idPilihan, respAlasan.getResponseText());
      ui.alert(hasil.pesan);
    }
    // CANCEL: tidak melakukan apa-apa
  } catch (e) {
    ui.alert('❌ GAGAL: ' + e.message);
  }
}

// ===== bayar_angsuran.js =====
/**
 * ============================================================
 * BAYAR_ANGSURAN.GS — Pembayaran Angsuran Manual / Lump-Sum
 * Koperasi App
 * ============================================================
 *
 * LATAR BELAKANG:
 * Kebijakan koperasi: pinjaman migrasi memakai tenor maksimal
 * agar cicilan wajib (potong gaji) kecil. Anggota yang mampu
 * boleh membayar beberapa tenor sekaligus — pembayaran ekstra
 * itu dicatat lewat fitur ini.
 *
 * CARA KERJA:
 * - Pembayaran manual = baris transaksi_angsuran biasa dengan
 *   keterangan "PEMBAYARAN MANUAL" (bukan hasil Generate).
 * - Kas masuk + jurnal #5 otomatis, sama seperti angsuran lain.
 * - Sisa pinjaman dihitung real-time (total_tagihan − SUM bayar),
 *   jadi Generate Angsuran bulan berikutnya OTOMATIS menyesuaikan:
 *   kalau sisa tinggal sedikit, tagihan terakhir dirapikan; kalau
 *   sudah 0, pinjaman jadi lunas dan tak ditagih lagi.
 * - Untuk pinjaman biasa (ada jasa), porsi pokok/jasa dibagi
 *   proporsional seperti Generate. Untuk pinjaman migrasi
 *   (jasa 0), seluruhnya potong pokok.
 *
 * PASANG: buat file baru "bayar_angsuran" di Apps Script,
 * tempel seluruh isi file ini, simpan, lalu deploy New version.
 * ============================================================
 */

/**
 * Daftar pinjaman AKTIF milik satu anggota + sisa masing-masing.
 * Dipakai UI untuk menampilkan pilihan pinjaman yang mau dibayar.
 */
function apiPinjamanAnggota(idAnggota) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;

  const pinjaman = getRowsByFilter(SHEET.PINJAMAN, function(row) {
    return row.id_anggota === idAnggota && row.status === 'aktif' &&
           row.status_lock !== STATUS_LOCK.VOID;
  });

  const kini = new Date();
  const tahun = getTahunAktif() || kini.getFullYear();
  const bulan = kini.getMonth() + 1;
  return pinjaman.map(function(p) {
    const menurun = isJasaMenurun(p);
    // Jasa bulan ini sudah ditagih (ada angsuran bulan ini yang memuat jasa)?
    const sudahDitagih = getRowsByFilter(SHEET.ANGSURAN, function(a) {
      return String(a.id_pinjaman) === String(p.id_pinjaman) &&
             a.status_lock !== STATUS_LOCK.VOID && Number(a.tahun) === tahun &&
             Number(a.bulan) === bulan && Number(a.jasa) > 0;
    }).length > 0;
    return {
      id: p.id_pinjaman,
      keterangan: String(p.keterangan || ''),
      totalTagihan: Number(p.total_tagihan),
      angsuranPerbulan: Number(p.angsuran_perbulan),
      tenor: Number(p.tenor),
      sisa: getSisaPinjaman(p.id_pinjaman),
      metode: menurun ? 'menurun' : 'flat',
      jasaPersen: Number(p.jasa_persen) || 0,
      jasaBulanIni: menurun && !sudahDitagih ? hitungJasaMenurun(p, tahun, bulan) : 0,
      jasaSudahDitagih: sudahDitagih
    };
  });
}

/**
 * Catat satu pembayaran angsuran manual.
 * @param {Object} d  { id_pinjaman, nominal, keterangan }
 * @return {string} pesan sukses
 */
function apiBayarAngsuran(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const hasil = inputAngsuranManual({
    id_pinjaman: d.id_pinjaman,
    nominal: Number(d.nominal),
    pokok: d.pokok,
    jasa: d.jasa,
    keterangan: d.keterangan || ''
  });
  return { pesan: hasil.pesan, struk: dataStruk('angsuran', hasil.id_angsuran) };
}

/**
 * Mesin inti pembayaran manual — bisa juga dipanggil dari editor.
 */
function inputAngsuranManual(data) {
  const tanggal = data.tanggal ? new Date(data.tanggal) : new Date();
  const tahun = getTahunAktif();
  const bulan = tanggal.getMonth() + 1;

  validateTahunAktif(tahun);

  // ---------- VALIDASI PINJAMAN ----------
  const pinjamanList = getRowsByFilter(SHEET.PINJAMAN, function(row) {
    return row.id_pinjaman === data.id_pinjaman;
  });
  if (pinjamanList.length === 0) {
    throw new Error('Pinjaman ' + data.id_pinjaman + ' tidak ditemukan.');
  }
  const pinjaman = pinjamanList[0];
  if (pinjaman.status_lock === STATUS_LOCK.VOID) {
    throw new Error('Pinjaman ini sudah di-VOID.');
  }
  if (pinjaman.status !== 'aktif') {
    throw new Error('Pinjaman ini sudah lunas.');
  }

  const anggota = getAnggota(pinjaman.id_anggota);
  const namaAnggota = anggota ? anggota.nama : pinjaman.id_anggota;

  // ---------- VALIDASI NOMINAL & SPLIT POKOK / JASA ----------
  const sisa = getSisaPinjaman(pinjaman.id_pinjaman);
  let angsuranPokok, jasaBagian, nominal;

  if (isJasaMenurun(pinjaman)) {
    // Jasa menurun: pokok & jasa diisi terpisah. Tanpa rincian → semua pokok
    // (mis. pelunasan dipercepat / anggota keluar).
    angsuranPokok = Math.round(Number(data.pokok !== undefined && data.pokok !== '' &&
      data.pokok !== null ? data.pokok : data.nominal) || 0);
    jasaBagian = Math.round(Number(data.jasa) || 0);
    if (angsuranPokok < 0 || jasaBagian < 0) throw new Error('Pokok dan jasa tidak boleh minus.');
    nominal = angsuranPokok + jasaBagian;
    if (!(nominal > 0)) throw new Error('Nominal pembayaran harus lebih dari 0.');
    if (angsuranPokok > sisa) {
      throw new Error('Pokok (' + formatRupiah(angsuranPokok) + ') melebihi sisa pokok (' +
        formatRupiah(sisa) + '). Untuk pelunasan penuh, isi pokok persis ' +
        formatRupiah(sisa) + '.');
    }
  } else {
    nominal = Math.round(Number(data.nominal));
    if (!(nominal > 0)) {
      throw new Error('Nominal pembayaran harus lebih dari 0.');
    }
    if (nominal > sisa) {
      throw new Error('Nominal (' + formatRupiah(nominal) +
        ') melebihi sisa pinjaman (' + formatRupiah(sisa) + '). ' +
        'Untuk pelunasan penuh, isi persis ' + formatRupiah(sisa) + '.');
    }
    // Data lama (flat): split proporsional, konsisten dengan Generate
    const rasioPokok = Number(pinjaman.total_tagihan) > 0
      ? Number(pinjaman.nominal) / Number(pinjaman.total_tagihan) : 1;
    angsuranPokok = Math.round(nominal * rasioPokok);
    jasaBagian = nominal - angsuranPokok;
  }

  // ---------- SIMPAN transaksi_angsuran ----------
  const idAngsuran = generateId(SHEET.ANGSURAN, ID_PREFIX.ANGSURAN, 1, tahun, null);

  const rowAngsuran = {
    id_angsuran: idAngsuran,
    tanggal: tanggal,
    tahun: tahun,
    bulan: bulan,
    id_pinjaman: pinjaman.id_pinjaman,
    id_anggota: pinjaman.id_anggota,
    angsuran_pokok: angsuranPokok,
    jasa: jasaBagian,
    denda: 0,
    total_bayar: nominal,
    keterangan: 'PEMBAYARAN MANUAL' +
      (data.keterangan ? ' — ' + data.keterangan : ''),
    user_input: getUserEmail(),
    timestamp: new Date(),
    status_lock: STATUS_LOCK.OPEN
  };
  appendRowFromObject(SHEET.ANGSURAN, rowAngsuran);

  // ---------- KAS MASUK ----------
  catatKasMasuk({
    tanggal: tanggal,
    tahun: tahun,
    bulan: bulan,
    kategori: 'angsuran',
    referensi: idAngsuran,
    keterangan: 'Pembayaran angsuran manual — ' + namaAnggota,
    nominal: nominal
  });

  // ---------- JURNAL #5 ----------
  catatJurnal(tanggal, tahun, bulan, idAngsuran,
    'Pembayaran angsuran manual — ' + namaAnggota,
    [
      { kode_akun: AKUN.KAS, debit: nominal, kredit: 0 },
      { kode_akun: AKUN.PIUTANG_PINJAMAN, debit: 0, kredit: angsuranPokok },
      { kode_akun: AKUN.PENDAPATAN_JASA_PINJAMAN, debit: 0, kredit: jasaBagian }
    ]
  );

  logAktivitas('INSERT', SHEET.ANGSURAN, idAngsuran, null, rowAngsuran);

  // ---------- CEK LUNAS ----------
  const lunas = perbaruiStatusPinjamanJikaLunas(pinjaman.id_pinjaman);
  const sisaBaru = getSisaPinjaman(pinjaman.id_pinjaman);

  return {
    id_angsuran: idAngsuran,
    sisa_baru: sisaBaru,
    lunas: lunas,
    pesan: '✅ Pembayaran ' + formatRupiah(nominal) + ' dari ' + namaAnggota +
      ' tercatat (' + idAngsuran + ').\n' +
      (lunas
        ? '🎉 Pinjaman ' + pinjaman.id_pinjaman + ' LUNAS!'
        : 'Sisa pinjaman sekarang: ' + formatRupiah(sisaBaru))
  };
}

// ===== config.js =====
/**
 * ============================================================
 * CONFIG.GS — Konstanta & Helper Inti
 * Koperasi App
 * Sesuai dokumen: REKAP_Koperasi_v2.md (BAGIAN 6, 11, 17)
 * ============================================================
 *
 * File ini TIDAK menampilkan menu apapun — murni kumpulan
 * konstanta & fungsi helper yang dipanggil dari file lain
 * (simpanan.gs, pinjaman.gs, angsuran.gs, dst).
 *
 * Urutan file yang dibaca Apps Script tidak masalah — semua
 * fungsi global bisa saling panggil antar file .gs.
 * ============================================================
 */

// ============================================================
// NAMA SHEET (satu sumber kebenaran, jangan ketik nama sheet
// langsung di file lain — selalu pakai konstanta ini)
// ============================================================

const SHEET = {
  DASHBOARD:        'DASHBOARD',
  ANGGOTA:          'anggota',
  SETTING_RAT:      'setting_rat',
  COA_AKUN:         'coa_akun',
  SALDO_AWAL:       'saldo_awal',
  SIMPANAN:         'transaksi_simpanan',
  PENGAMBILAN:      'transaksi_pengambilan',
  PINJAMAN:         'transaksi_pinjaman',
  ANGSURAN:         'transaksi_angsuran',
  KAS:              'transaksi_kas',
  JURNAL:           'jurnal_umum',
  JASA_SUKARELA:    'rekap_jasa_sukarela',
  REALISASI_SHU:    'realisasi_shu',
  LOG:              'log_aktivitas',
  LAP_SHU:          'lap_shu',
  LAP_NERACA:       'lap_neraca',
  LAP_LABARUGI:     'lap_labarugi',
  REKAP_PINJAMAN:   'rekap_pinjaman',
  REKAP_ANGSURAN:   'rekap_angsuran',
  REKAP_SIMPANAN:   'rekap_simpanan',
  STRUK_POTONGAN:   'struk_potongan'
};

// ============================================================
// PREFIX ID per sheet (sesuai BAGIAN 11 dokumen rekap)
// ============================================================

const ID_PREFIX = {
  ANGGOTA:       'AGT',
  SALDO_AWAL:    'SAL',
  SIMPANAN:      'SMP',
  PENGAMBILAN:   'AMB',
  PINJAMAN:      'PJM',
  ANGSURAN:      'ANG',
  KAS:           'KAS',
  JURNAL:        'JRN',
  JASA_SUKARELA: 'JAS',
  REALISASI_SHU: 'SHU',
  LOG:           'LOG'
};

// Kode akun COA (sesuai BAGIAN 4 — biar tidak salah ketik kode di file lain)
const AKUN = {
  KAS:                 '101',
  BANK:                '102',
  PIUTANG_PINJAMAN:    '103',
  SIMPANAN_POKOK:      '201',
  SIMPANAN_WAJIB:      '202',
  SIMPANAN_SUKARELA:   '203',
  MODAL_KOPERASI:      '301',
  SHU_BERJALAN:        '302',
  SHU_DITAHAN:         '303',
  PENDAPATAN_JASA_PINJAMAN:  '401',
  PENDAPATAN_JASA_SUKARELA:  '402',
  BEBAN_OPERASIONAL:   '501',
  DANA_CADANGAN:       '502',
  DANA_SOSIAL:         '503',
  DANA_PENGURUS:       '504',
  TRANSPORT:           '505',
  THR_LEBARAN:         '506',
  UANG_DUDUK_RAT:      '507',
  SOUVENIR_RAT:        '508',
  INSENTIF_PENGURUS:   '509',   // label baru: Honor Pengurus
  // Akun tambahan (mengikuti Excel pembukuan — lihat akuntansi.js)
  INVENTARIS:          '111',
  AKUM_PENYUSUTAN:     '112',
  CADANGAN:            '304',
  PENDAPATAN_LAIN:     '403',
  BEBAN_JASA_SUKARELA: '510',
  BEBAN_PENYUSUTAN:    '519'
};

const STATUS_LOCK = { OPEN: 'OPEN', LOCKED: 'LOCKED', VOID: 'VOID' };

// ============================================================
// GENERATOR ID OTOMATIS
// Format: PREFIX-TAHUN-NOMOR  atau  PREFIX-NOMOR (anggota, log)
// atau    PREFIX-TAHUN-BULAN-NOMOR (rekap_jasa_sukarela)
// ============================================================

/**
 * Membuat ID baru untuk sebuah sheet.
 * @param {string} namaSheet   nama sheet (pakai konstanta SHEET.xxx)
 * @param {string} prefix      prefix ID (pakai konstanta ID_PREFIX.xxx)
 * @param {number} kolomId     nomor kolom (1-based) tempat ID berada
 * @param {number|null} tahun  tahun buku, null jika sheet tidak per-tahun
 * @param {number|null} bulan  bulan (1-12), hanya untuk rekap_jasa_sukarela
 * @return {string} ID baru, contoh: "SMP-2026-0001"
 */
function generateId(namaSheet, prefix, kolomId, tahun, bulan) {
  let awalan = prefix;
  if (tahun) awalan += '-' + tahun;
  if (bulan) awalan += '-' + String(bulan).padStart(2, '0');
  awalan += '-';

  if (typeof __TULIS_MASSAL !== 'undefined' && __TULIS_MASSAL) {
    // Mode tulis massal: baca kolom ID sekali (termasuk baris tampungan), lalu lanjutkan di memori
    const b = bufferSheet_(namaSheet);
    if (!b.nomor.hasOwnProperty(awalan)) {
      const kolom = b.header[kolomId - 1];
      let maks = 0;
      __CACHE_SHEET[namaSheet].forEach(function(r) {
        const id = String(r[kolom]);
        if (id.indexOf(awalan) === 0) {
          const n = parseInt(id.substring(awalan.length), 10);
          if (!isNaN(n) && n > maks) maks = n;
        }
      });
      b.nomor[awalan] = maks;
    }
    b.nomor[awalan]++;
    return awalan + String(b.nomor[awalan]).padStart(4, '0');
  }

  const sheet = getSheet(namaSheet);
  const lastRow = sheet.getLastRow();

  let nomorTerbesar = 0;

  if (lastRow > 1) {
    const idList = sheet.getRange(2, kolomId, lastRow - 1, 1).getValues();
    idList.forEach(function(row) {
      const id = String(row[0]);
      if (id.indexOf(awalan) === 0) {
        const nomor = parseInt(id.substring(awalan.length), 10);
        if (!isNaN(nomor) && nomor > nomorTerbesar) nomorTerbesar = nomor;
      }
    });
  }

  const nomorBaru = String(nomorTerbesar + 1).padStart(4, '0');
  return awalan + nomorBaru;
}

/**
 * Buat beberapa ID berurutan sekaligus (mis. 3 baris jurnal satu
 * transaksi) — sheet cukup dibaca sekali, bukan sekali per ID.
 * @return {string[]}
 */
function generateIdBeruntun(namaSheet, prefix, kolomId, tahun, bulan, jumlah) {
  if (typeof __TULIS_MASSAL !== 'undefined' && __TULIS_MASSAL) {
    const semua = [];
    for (let i = 0; i < jumlah; i++) semua.push(generateId(namaSheet, prefix, kolomId, tahun, bulan));
    return semua;
  }
  const pertama = generateId(namaSheet, prefix, kolomId, tahun, bulan);
  const awalan = pertama.substring(0, pertama.lastIndexOf('-') + 1);
  const nomorAwal = parseInt(pertama.substring(awalan.length), 10);
  const hasil = [];
  for (let i = 0; i < jumlah; i++) {
    hasil.push(awalan + String(nomorAwal + i).padStart(4, '0'));
  }
  return hasil;
}

/** Khusus log_aktivitas: format LOG-000001 (6 digit, tanpa tahun) */
function generateIdLog() {
  const sheet = getSheet(SHEET.LOG);
  const lastRow = sheet.getLastRow();
  let nomorTerbesar = 0;

  // Jalur cepat: log selalu ditambah di bawah, jadi ID terbesar ada di
  // baris terakhir. Sheet log bisa ribuan baris — tidak perlu dibaca semua.
  if (lastRow > 1) {
    const idTerakhir = String(sheet.getRange(lastRow, 1).getValue());
    const nomor = parseInt(idTerakhir.substring(4), 10);
    if (idTerakhir.indexOf('LOG-') === 0 && !isNaN(nomor)) {
      return 'LOG-' + String(nomor + 1).padStart(6, '0');
    }
  }

  if (lastRow > 1) {
    const idList = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    idList.forEach(function(row) {
      const id = String(row[0]);
      if (id.indexOf('LOG-') === 0) {
        const nomor = parseInt(id.substring(4), 10);
        if (!isNaN(nomor) && nomor > nomorTerbesar) nomorTerbesar = nomor;
      }
    });
  }
  return 'LOG-' + String(nomorTerbesar + 1).padStart(6, '0');
}

/** Khusus anggota: format AGT-0001 (4 digit, tanpa tahun) */
function generateIdAnggota() {
  const sheet = getSheet(SHEET.ANGGOTA);
  const lastRow = sheet.getLastRow();
  let nomorTerbesar = 0;

  if (lastRow > 1) {
    const idList = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    idList.forEach(function(row) {
      const id = String(row[0]);
      if (id.indexOf('AGT-') === 0) {
        const nomor = parseInt(id.substring(4), 10);
        if (!isNaN(nomor) && nomor > nomorTerbesar) nomorTerbesar = nomor;
      }
    });
  }
  return 'AGT-' + String(nomorTerbesar + 1).padStart(4, '0');
}

// ============================================================
// SETTING RAT & TAHUN AKTIF
// ============================================================

/**
 * Ambil baris setting_rat untuk tahun tertentu.
 * @return {Object|null} objek setting, atau null jika tidak ada
 */
function getSettingRAT(tahun) {
  const data = sheetToObjects(SHEET.SETTING_RAT);
  const found = data.find(function(row) { return Number(row.tahun) === Number(tahun); });
  return found || null;
}

/** Ambil tahun yang sedang status_aktif = TRUE. @return {number|null} */
function getTahunAktif() {
  const data = sheetToObjects(SHEET.SETTING_RAT);
  const aktif = data.find(function(row) { return row.status_aktif === true; });
  return aktif ? Number(aktif.tahun) : null;
}

// ============================================================
// NOMINAL SIMPANAN (diatur per tahun buku di halaman Pengaturan)
// ============================================================

const DEFAULT_NOMINAL_SIMPANAN = {
  nominal_pokok: 10000,      // simpanan pokok (sekali seumur keanggotaan)
  nominal_wajib: 50000,      // simpanan wajib per bulan
  minimal_sukarela: 10000    // setoran sukarela minimal
};

/**
 * Tambahkan kolom nominal simpanan ke setting_rat bila belum ada
 * (sheet lama dibuat sebelum fitur ini). Baris yang sudah ada diisi
 * nilai bawaan supaya perilaku tidak berubah.
 */
function pastikanKolomNominalSimpanan() {
  const sheet = getSheet(SHEET.SETTING_RAT);
  const header = getHeader(SHEET.SETTING_RAT).map(function(h) { return String(h).trim(); });
  const kurang = Object.keys(DEFAULT_NOMINAL_SIMPANAN)
    .filter(function(k) { return header.indexOf(k) === -1; });
  if (kurang.length === 0) return;

  const kolomAwal = header.length + 1;
  sheet.getRange(1, kolomAwal, 1, kurang.length).setValues([kurang])
    .setFontWeight('bold');
  const jumlahBaris = sheet.getLastRow() - 1;
  if (jumlahBaris > 0) {
    const isi = [];
    for (let i = 0; i < jumlahBaris; i++) {
      isi.push(kurang.map(function(k) { return DEFAULT_NOMINAL_SIMPANAN[k]; }));
    }
    sheet.getRange(2, kolomAwal, jumlahBaris, kurang.length).setValues(isi);
  }
  hapusCache(SHEET.SETTING_RAT);
}

/**
 * Nominal simpanan untuk satu tahun buku. Kolom kosong / belum ada
 * → pakai nilai bawaan.
 * @return {{nominal_pokok:number, nominal_wajib:number, minimal_sukarela:number}}
 */
function getNominalSimpanan(tahun) {
  const setting = getSettingRAT(tahun) || {};
  const hasil = {};
  Object.keys(DEFAULT_NOMINAL_SIMPANAN).forEach(function(k) {
    const nilai = Number(setting[k]);
    hasil[k] = nilai > 0 ? nilai : DEFAULT_NOMINAL_SIMPANAN[k];
  });
  return hasil;
}

/**
 * Validasi wajib sebelum simpan transaksi apapun (BAGIAN 17 v2).
 * Melempar error (throw) jika tahun tidak valid / tidak aktif.
 */
function validateTahunAktif(tahun) {
  const setting = getSettingRAT(tahun);
  if (!setting) {
    throw new Error('Setting RAT untuk tahun ' + tahun + ' belum ada. ' +
      'Isi dulu di sheet setting_rat.');
  }
  if (setting.status_aktif !== true) {
    throw new Error('Tahun ' + tahun + ' sudah ditutup (LOCKED). ' +
      'Input transaksi baru ditolak.');
  }
  return setting;
}

/**
 * Validasi total persen alokasi SHU (BAGIAN 17 v2).
 * Dipanggil sebelum proses tutup buku / hitung SHU.
 */
function validateAlokasiSHU(setting) {
  // 7 pos alokasi SHU (lihat POS_ALOKASI di akuntansi.js) harus berjumlah 100%
  let totalPos = 0;
  POS_ALOKASI.forEach(function(p) {
    const n = Number(setting[p[0]]);
    if (isNaN(n) || n < 0) throw new Error('Persen ' + p[1] + ' tidak valid.');
    totalPos += n;
  });
  if (Math.abs(totalPos - 100) > 0.001) {
    throw new Error('Jumlah persen alokasi SHU (cadangan, dana anggota, pengurus, ' +
      'kesejahteraan, pendidikan, sosial, pembangunan) harus = 100%, saat ini = ' +
      Math.round(totalPos * 1000) / 1000 + '%.');
  }
  const pajak = Number(setting.pajak_persen || 0);
  if (isNaN(pajak) || pajak < 0 || pajak > 100) throw new Error('Persen pajak tidak valid.');
  const totalSplit = Number(setting.shu_simpanan) + Number(setting.shu_jasa);
  if (totalSplit !== 100) {
    throw new Error('shu_simpanan + shu_jasa harus = 100, saat ini = ' + totalSplit);
  }
  return true;
}

// ============================================================
// LOG AKTIVITAS
// Dipanggil setiap ada INSERT / EDIT / DELETE / VOID / APPROVE / REJECT
// ============================================================

/**
 * Catat satu baris ke log_aktivitas.
 * @param {string} aksi         INSERT / EDIT / DELETE / VOID / APPROVE / REJECT
 * @param {string} sheetTarget  nama sheet yang diubah (pakai SHEET.xxx)
 * @param {string} idReferensi  ID baris yang diubah
 * @param {Object|null} dataLama  nilai sebelum perubahan (akan di-JSON.stringify)
 * @param {Object|null} dataBaru  nilai sesudah perubahan (akan di-JSON.stringify)
 */
function logAktivitas(aksi, sheetTarget, idReferensi, dataLama, dataBaru) {
  // Mode tulis massal (migrasi): log per baris dilewati — pemanggil menulis satu log ringkasan
  if (typeof __TULIS_MASSAL !== 'undefined' && __TULIS_MASSAL) return;
  const sheet = getSheet(SHEET.LOG);
  const id = generateIdLog();
  const user = getUserEmail();

  sheet.appendRow(amanBaris_([
    id,
    new Date(),
    user,
    aksi,
    sheetTarget,
    idReferensi,
    dataLama ? JSON.stringify(dataLama) : '',
    dataBaru ? JSON.stringify(dataBaru) : ''
  ]));
}

// ============================================================
// HELPER UMUM
// ============================================================

/** Email operator yang sedang login (dipakai di kolom user_input) */
function getUserEmail() {
  // Web app memakai login mandiri (sesi.js): Session.getActiveUser()
  // selalu kosong untuk pengunjung, jadi utamakan email dari sesi.
  const dariSesi = typeof __SESI_AKTIF !== 'undefined' && __SESI_AKTIF;
  return dariSesi || Session.getActiveUser().getEmail() || 'unknown';
}

/** Format angka jadi Rupiah untuk ditampilkan di dialog/laporan. */
function formatRupiah(angka) {
  return 'Rp ' + (Math.round(Number(angka) || 0) || 0).toLocaleString('id-ID');   // rupiah penuh, tanpa "-0"
}

/**
 * MODE_SENYAP: kalau true, semua pesan hanya ke Logger (tanpa popup).
 * Dipakai saat menjalankan test dari editor supaya eksekusi tidak
 * berhenti menunggu klik OK di jendela spreadsheet. Default false
 * di setiap eksekusi baru, jadi menu spreadsheet tetap normal.
 */
let MODE_SENYAP = false;

/**
 * Tampilkan pesan ke pengguna: pakai popup jika bisa (dijalankan
 * dari menu spreadsheet), kalau tidak bisa (dijalankan langsung
 * dari editor Apps Script, tanpa konteks UI) otomatis dicatat
 * ke Logger saja tanpa error.
 *
 * SELALU pakai fungsi ini untuk menampilkan pesan ke pengguna —
 * jangan panggil SpreadsheetApp.getUi().alert() langsung di file
 * lain, supaya semua fungsi aman dijalankan dari editor maupun
 * dari menu spreadsheet.
 */
function tampilkanPesan(pesan) {
  if (MODE_SENYAP) {
    Logger.log(pesan);
    return;
  }
  try {
    SpreadsheetApp.getUi().alert(pesan);
  } catch (e) {
    Logger.log(pesan);
  }
}

// ===== dashboard.js =====
/**
 * ============================================================
 * DASHBOARD.GS — Ringkasan Saldo Per Anggota
 * Koperasi App
 * ============================================================
 *
 * TUJUAN:
 * Sheet-sheet seperti saldo_awal, transaksi_simpanan, dll
 * sengaja dibuat format PANJANG (satu anggota = banyak baris)
 * supaya gampang diproses script. Tapi format ini agak
 * merepotkan kalau mau dibaca manusia langsung.
 *
 * Script ini membuat RINGKASAN di sheet DASHBOARD — satu baris
 * per anggota, kolom ke samping untuk tiap jenis saldo — mirip
 * Excel lama, tapi ANGKANYA DIHITUNG OTOMATIS dari data asli
 * (bukan diketik manual, jadi tidak akan pernah beda sendiri).
 *
 * CARA PAKAI:
 * Menu 🔧 SETUP → 📊 Refresh Dashboard
 * (atau jalankan fungsi refreshDashboard dari editor)
 *
 * Jalankan ulang kapan saja setelah ada transaksi baru untuk
 * memperbarui angkanya.
 * ============================================================
 */

const HEADER_DASHBOARD = [
  'id_anggota', 'nama', 'jabatan', 'status',
  'simpanan_pokok', 'simpanan_wajib', 'simpanan_sukarela',
  'total_simpanan', 'sisa_pinjaman', 'saldo_kas_koperasi'
];

function refreshDashboard() {
  const sheet = getSheet(SHEET.DASHBOARD);
  const anggotaList = sheetToObjects(SHEET.ANGGOTA);

  // Bersihkan isi lama (kecuali baris 1 label lama, akan ditimpa header baru)
  sheet.clear();

  // Judul & waktu update
  sheet.getRange('A1').setValue('📊 RINGKASAN SALDO ANGGOTA — Koperasi ' + getProfilKoperasi().nama_koperasi.replace(/"/g, ''))
    .setFontWeight('bold').setFontSize(14);
  sheet.getRange('A2').setValue('Diperbarui: ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMMM yyyy, HH:mm'))
    .setFontStyle('italic').setFontColor('#666666');

  // Saldo kas koperasi (satu angka untuk semua, bukan per-anggota)
  const saldoKas = getSaldoKas();
  sheet.getRange('A3').setValue('Saldo Kas Koperasi saat ini:')
    .setFontWeight('bold');
  sheet.getRange('B3').setValue(saldoKas)
    .setNumberFormat('Rp #,##0').setFontWeight('bold').setFontColor('#38761d');

  // Header tabel mulai baris 5
  const baseRow = 5;
  const headerRange = sheet.getRange(baseRow, 1, 1, HEADER_DASHBOARD.length);
  headerRange.setValues([HEADER_DASHBOARD]);
  headerRange.setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#4a86e8').setHorizontalAlignment('center');
  sheet.setFrozenRows(baseRow);

  if (anggotaList.length === 0) {
    sheet.getRange(baseRow + 1, 1).setValue('(Belum ada data anggota)')
      .setFontStyle('italic').setFontColor('#999999');
    formatKolom(sheet, baseRow);
    return;
  }

  // Hitung baris per anggota
  const baris = anggotaList.map(function(anggota) {
    const pokok    = getSaldoSimpanan(anggota.id_anggota, 'pokok');
    const wajib    = getSaldoSimpanan(anggota.id_anggota, 'wajib');
    const sukarela = getSaldoSimpanan(anggota.id_anggota, 'sukarela');
    const totalSimpanan = pokok + wajib + sukarela;
    const sisaPinjaman = getSisaPinjamanAnggota(anggota.id_anggota);

    return [
      anggota.id_anggota,
      anggota.nama,
      anggota.jabatan,
      anggota.status,
      pokok,
      wajib,
      sukarela,
      totalSimpanan,
      sisaPinjaman,
      '' // saldo_kas_koperasi hanya diisi di baris ringkasan atas, kosongkan per baris
    ];
  });

  sheet.getRange(baseRow + 1, 1, baris.length, HEADER_DASHBOARD.length).setValues(amanBaris_(baris));

  // Baris TOTAL di paling bawah
  const totalRow = baseRow + baris.length + 1;
  sheet.getRange(totalRow, 2).setValue('TOTAL').setFontWeight('bold');
  ['E', 'F', 'G', 'H', 'I'].forEach(function(kolom, i) {
    const kolomIdx = 5 + i; // E=5 (simpanan_pokok) s/d I=9 (sisa_pinjaman)
    const range = sheet.getRange(baseRow + 1, kolomIdx, baris.length, 1);
    sheet.getRange(totalRow, kolomIdx).setFormula(
      '=SUM(' + range.getA1Notation() + ')'
    ).setFontWeight('bold');
  });

  formatKolom(sheet, baseRow);

  SpreadsheetApp.flush();
  tampilkanPesan('✅ Dashboard diperbarui. ' + anggotaList.length + ' anggota ditampilkan.');
}

/** Rapikan lebar kolom & format angka Rupiah */
function formatKolom(sheet, baseRow) {
  sheet.setColumnWidth(1, 100);  // id_anggota
  sheet.setColumnWidth(2, 160);  // nama
  sheet.setColumnWidth(3, 90);   // jabatan
  sheet.setColumnWidth(4, 80);   // status
  for (let c = 5; c <= 9; c++) sheet.setColumnWidth(c, 140);

  const lastRow = sheet.getLastRow();
  if (lastRow > baseRow) {
    sheet.getRange(baseRow + 1, 5, lastRow - baseRow, 5)
      .setNumberFormat('Rp #,##0');
  }
}

/**
 * Total sisa pinjaman AKTIF milik satu anggota
 * (bisa lebih dari satu pinjaman aktif secara teori,
 * meski aturan bisnis membatasi 1 pinjaman aktif per anggota).
 */
function getSisaPinjamanAnggota(idAnggota) {
  const pinjamanList = getRowsByFilter(SHEET.PINJAMAN, function(row) {
    return row.id_anggota === idAnggota && row.status === 'aktif';
  });
  return pinjamanList.reduce(function(total, p) {
    return total + getSisaPinjaman(p.id_pinjaman);
  }, 0);
}

// ===== database.js =====
/**
 * ============================================================
 * DATABASE.GS — Fungsi Baca/Tulis Generik ke Sheet
 * Koperasi App
 * Sesuai dokumen: REKAP_Koperasi_v2.md (BAGIAN 6)
 * ============================================================
 *
 * File ini adalah lapisan "akses data" — semua modul lain
 * (simpanan.gs, pinjaman.gs, dst) memakai fungsi di sini
 * untuk baca/tulis sheet, TIDAK langsung memanggil
 * SpreadsheetApp / getRange di file masing-masing.
 *
 * Keuntungan: kalau nanti ada perubahan struktur sheet,
 * cukup diperbaiki di satu tempat ini.
 * ============================================================
 */

// ============================================================
// AKSES SHEET DASAR
// ============================================================

/** Ambil objek Sheet dari nama. Error jelas jika sheet tidak ada. */
function getSheet(namaSheet) {
  pastikanAksesData(); // lihat sesi.js — cegah panggilan langsung dari browser
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet);
  if (!sheet) {
    throw new Error('Sheet "' + namaSheet + '" tidak ditemukan. ' +
      'Jalankan menu 🔧 SETUP → Jalankan Setup Lengkap dulu.');
  }
  return sheet;
}

/** Ambil array header (baris 1) dari sebuah sheet. */
function getHeader(namaSheet) {
  if (__TULIS_MASSAL && __TULIS_MASSAL.sheet[namaSheet]) return __TULIS_MASSAL.sheet[namaSheet].header.slice();
  const sheet = getSheet(namaSheet);
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

// ============================================================
// BACA DATA — sheet → array of objects
// ============================================================

// ============================================================
// CACHE — sheet hanya dibaca SEKALI per eksekusi
// Tanpa ini, fungsi seperti hitungSHU bisa membaca sheet yang
// sama ratusan kali dan berjalan bermenit-menit.
// Cache otomatis dihapus saat ada penulisan ke sheet tersebut.
// ============================================================

let __CACHE_SHEET = {};

// Cache turunan (peta saldo, sisa pinjaman, data SHU) — dihitung dari
// beberapa sheet sekaligus, jadi ikut dibuang setiap ada penulisan.
let __CACHE_TURUNAN = {};

/** Hapus cache satu sheet (dipanggil otomatis setiap ada penulisan). */
function hapusCache(namaSheet) {
  if (!(__TULIS_MASSAL && __TULIS_MASSAL.sheet[namaSheet])) delete __CACHE_SHEET[namaSheet];
  __CACHE_TURUNAN = {};
}

/** Hapus semua cache (jarang perlu — hanya jika edit sheet manual di tengah eksekusi). */
function hapusSemuaCache() {
  __CACHE_SHEET = {};
  __CACHE_TURUNAN = {};
}

/** Ambil nilai cache turunan; kalau belum ada, hitung sekali lalu simpan. */
function cacheTurunan(kunci, hitungFn) {
  if (!__CACHE_TURUNAN.hasOwnProperty(kunci)) __CACHE_TURUNAN[kunci] = hitungFn();
  return __CACHE_TURUNAN[kunci];
}

/**
 * Baca seluruh isi sheet (kecuali header) jadi array of objects.
 * Setiap objek juga punya properti tersembunyi __row (nomor baris asli
 * di spreadsheet, 1-based) supaya bisa dipakai untuk update/hapus.
 *
 * HASIL DI-CACHE per eksekusi — pembacaan kedua dst dari sheet yang
 * sama langsung dari memori (jauh lebih cepat).
 *
 * Contoh hasil: [{ id_anggota: 'AGT-0001', nama: 'Budi', ..., __row: 2 }, ...]
 */
function sheetToObjects(namaSheet) {
  if (__CACHE_SHEET[namaSheet]) return __CACHE_SHEET[namaSheet];

  const sheet = getSheet(namaSheet);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol === 0) {
    __CACHE_SHEET[namaSheet] = [];
    return __CACHE_SHEET[namaSheet];
  }

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const hasil = data.map(function(row, i) {
    const obj = {};
    header.forEach(function(kolom, j) { obj[kolom] = row[j]; });
    obj.__row = i + 2; // baris asli di sheet (1-based, header = baris 1)
    return obj;
  });

  __CACHE_SHEET[namaSheet] = hasil;
  return hasil;
}

/**
 * Cari SATU baris berdasarkan nilai kolom tertentu.
 * @return {Object|null}
 */
function getRowByField(namaSheet, namaKolom, nilai) {
  const data = sheetToObjects(namaSheet);
  const found = data.find(function(row) {
    return String(row[namaKolom]) === String(nilai);
  });
  return found || null;
}

/**
 * Cari BANYAK baris yang cocok dengan fungsi filter.
 * @param {function(Object): boolean} filterFn
 * @return {Object[]}
 */
function getRowsByFilter(namaSheet, filterFn) {
  return sheetToObjects(namaSheet).filter(filterFn);
}

// ============================================================
// TULIS DATA — objects → sheet
// ============================================================

/**
 * Teks yang diawali = + - @ akan dibaca Google Sheets sebagai RUMUS
 * (mis. keterangan "=IMPORTXML(...)" bisa mengirim data ke luar).
 * Beri tanda petik di depan agar disimpan sebagai teks biasa;
 * saat dibaca kembali, Sheets mengembalikan teks aslinya tanpa petik.
 */
function amanSel_(v) {
  return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v;
}
function amanBaris_(baris) {
  return baris.map(function(r) { return Array.isArray(r) ? r.map(amanSel_) : amanSel_(r); });
}

// ============================================================
// MODE TULIS MASSAL (impor/migrasi ribuan baris dalam satu eksekusi)
// ------------------------------------------------------------
// Selama aktif, baris baru & perubahan ditampung di memori (dan langsung
// terlihat oleh sheetToObjects/generateId), lalu ditulis sekaligus oleh
// selesaiTulisMassal(). Gagal di tengah → batalTulisMassal(): tidak ada
// yang tertulis sama sekali. Fungsi bisnis (setoran, angsuran, dst.)
// tidak perlu diubah.
// ============================================================

let __TULIS_MASSAL = null;

/** Selama mode ini, logAktivitas per baris dilewati (tulis satu log ringkasan sesudahnya). */
function mulaiTulisMassal() {
  if (__TULIS_MASSAL) throw new Error('Mode tulis massal sudah aktif.');
  __TULIS_MASSAL = { sheet: {} };
}

function bufferSheet_(namaSheet) {
  let b = __TULIS_MASSAL.sheet[namaSheet];
  if (b) return b;
  const sheet = getSheet(namaSheet);
  const lastCol = sheet.getLastColumn();
  const header = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  sheetToObjects(namaSheet);                         // muat isi lama ke cache
  b = __TULIS_MASSAL.sheet[namaSheet] = {
    header: header, baris: [], ubah: [], barisAwal: Math.max(sheet.getLastRow(), 1) + 1, nomor: {}
  };
  return b;
}

function tampungBaris_(namaSheet, dataObjek) {
  const b = bufferSheet_(namaSheet);
  const obj = {};
  b.baris.push(b.header.map(function(kolom) {
    const ada = dataObjek.hasOwnProperty(kolom);
    obj[kolom] = ada ? dataObjek[kolom] : '';
    return ada ? amanSel_(dataObjek[kolom]) : '';
  }));
  obj.__row = b.barisAwal + b.baris.length - 1;
  __CACHE_SHEET[namaSheet].push(obj);
  __CACHE_TURUNAN = {};
  return obj.__row;
}

/** Tulis semua tampungan ke sheet. @return {Object} jumlah baris per sheet */
function selesaiTulisMassal() {
  const m = __TULIS_MASSAL;
  if (!m) return {};
  const ringkas = {};
  Object.keys(m.sheet).forEach(function(nama) {
    const b = m.sheet[nama];
    const sheet = getSheet(nama);
    if (b.baris.length) {
      pastikanKapasitasBaris_(sheet, b.barisAwal + b.baris.length - 1);
      sheet.getRange(b.barisAwal, 1, b.baris.length, b.header.length).setValues(b.baris);
      ringkas[nama] = b.baris.length;
    }
    b.ubah.forEach(function(u) { sheet.getRange(u[0], u[1]).setValue(u[2]); });
  });
  __TULIS_MASSAL = null;
  // Apps Script menunda penulisan ke sheet; paksa sekarang supaya penolakan
  // (mis. aturan validasi sel) muncul DI SINI, bukan setelah langkah dianggap berhasil.
  SpreadsheetApp.flush();
  hapusSemuaCache();
  return ringkas;
}

/**
 * getRange().setValues() TIDAK menambah baris sendiri (beda dengan appendRow);
 * sheet baru hanya punya 1.000 baris. Tambah baris kosong bila perlu.
 */
function pastikanKapasitasBaris_(sheet, barisTerakhir) {
  const maks = sheet.getMaxRows();
  if (typeof maks === 'number' && barisTerakhir > maks) sheet.insertRowsAfter(maks, barisTerakhir - maks + 100);
}

/** Buang semua tampungan (dipakai saat terjadi error). */
function batalTulisMassal() {
  __TULIS_MASSAL = null;
  hapusSemuaCache();
}

/**
 * Tambah satu baris baru ke sheet berdasarkan objek.
 * Kolom yang tidak ada di objek akan dikosongkan.
 * Urutan diambil otomatis dari header sheet (bukan urutan objek),
 * jadi aman meskipun properti objek ditulis acak.
 *
 * @param {string} namaSheet
 * @param {Object} dataObjek  contoh: { id_anggota: 'AGT-0001', nama: 'Budi' }
 * @return {number} nomor baris tempat data baru disimpan
 */
function appendRowFromObject(namaSheet, dataObjek) {
  if (__TULIS_MASSAL) return tampungBaris_(namaSheet, dataObjek);
  const sheet = getSheet(namaSheet);
  const header = getHeader(namaSheet);

  const baris = header.map(function(kolom) {
    return dataObjek.hasOwnProperty(kolom) ? amanSel_(dataObjek[kolom]) : '';
  });

  sheet.appendRow(baris);
  hapusCache(namaSheet); // data berubah → cache tidak berlaku lagi
  return sheet.getLastRow();
}

/**
 * Tambah BANYAK baris sekaligus (satu kali tulis, jauh lebih cepat
 * daripada appendRow berulang). Aturan kolom sama dengan appendRowFromObject.
 * @param {string} namaSheet
 * @param {Object[]} daftarObjek
 */
function appendRowsFromObjects(namaSheet, daftarObjek) {
  if (!daftarObjek || daftarObjek.length === 0) return;
  if (__TULIS_MASSAL) { daftarObjek.forEach(function(o) { tampungBaris_(namaSheet, o); }); return; }
  const sheet = getSheet(namaSheet);
  const header = getHeader(namaSheet);

  const baris = daftarObjek.map(function(obj) {
    return header.map(function(kolom) {
      return obj.hasOwnProperty(kolom) ? amanSel_(obj[kolom]) : '';
    });
  });

  pastikanKapasitasBaris_(sheet, sheet.getLastRow() + baris.length);
  sheet.getRange(sheet.getLastRow() + 1, 1, baris.length, header.length).setValues(baris);
  hapusCache(namaSheet);
}

/**
 * Update satu baris yang sudah ada, berdasarkan nomor baris (__row).
 * Hanya kolom yang ada di dataBaru yang diubah — kolom lain tidak disentuh.
 *
 * @param {string} namaSheet
 * @param {number} nomorBaris  ambil dari properti __row hasil sheetToObjects
 * @param {Object} dataBaru    contoh: { status: 'nonaktif' }
 */
function updateRowByRowNumber(namaSheet, nomorBaris, dataBaru) {
  if (__TULIS_MASSAL) {
    const b = bufferSheet_(namaSheet);
    const obj = __CACHE_SHEET[namaSheet].find(function(r) { return r.__row === nomorBaris; });
    Object.keys(dataBaru).forEach(function(kolom) {
      const idx = b.header.indexOf(kolom);
      if (idx === -1) throw new Error('Kolom "' + kolom + '" tidak ada di sheet "' + namaSheet + '"');
      if (obj) obj[kolom] = dataBaru[kolom];
      if (nomorBaris >= b.barisAwal) b.baris[nomorBaris - b.barisAwal][idx] = amanSel_(dataBaru[kolom]);
      else b.ubah.push([nomorBaris, idx + 1, amanSel_(dataBaru[kolom])]);
    });
    __CACHE_TURUNAN = {};
    return;
  }
  const sheet = getSheet(namaSheet);
  const header = getHeader(namaSheet);

  Object.keys(dataBaru).forEach(function(kolom) {
    const idx = header.indexOf(kolom);
    if (idx === -1) {
      throw new Error('Kolom "' + kolom + '" tidak ada di sheet "' + namaSheet + '"');
    }
    sheet.getRange(nomorBaris, idx + 1).setValue(amanSel_(dataBaru[kolom]));
  });
  hapusCache(namaSheet); // data berubah → cache tidak berlaku lagi
}

/**
 * Cari satu baris berdasarkan kolom tertentu, lalu update.
 * Gabungan getRowByField + updateRowByRowNumber, lebih ringkas dipakai.
 *
 * @return {boolean} true jika baris ditemukan & diupdate, false jika tidak ketemu
 */
function updateRowByField(namaSheet, namaKolomCari, nilaiCari, dataBaru) {
  const row = getRowByField(namaSheet, namaKolomCari, nilaiCari);
  if (!row) return false;
  updateRowByRowNumber(namaSheet, row.__row, dataBaru);
  return true;
}

// ============================================================
// FUNGSI SUM / AGREGAT — dipakai untuk hitung saldo, dsb
// ============================================================

/**
 * Jumlahkan nilai satu kolom numerik dari baris yang lolos filter.
 * Baris berstatus VOID otomatis diabaikan (jika sheet punya status_lock).
 *
 * @param {string} namaSheet
 * @param {string} kolomAngka   nama kolom yang mau dijumlah
 * @param {function(Object): boolean} filterFn  filter tambahan (opsional)
 * @return {number}
 */
function sumKolom(namaSheet, kolomAngka, filterFn) {
  const data = sheetToObjects(namaSheet);
  return data.reduce(function(total, row) {
    if (row.status_lock === STATUS_LOCK.VOID) return total; // skip VOID
    if (filterFn && !filterFn(row)) return total;
    const nilai = Number(row[kolomAngka]) || 0;
    return total + nilai;
  }, 0);
}

// ============================================================
// LOOKUP UMUM (dipakai lintas modul)
// ============================================================

/** Ambil data satu anggota berdasarkan id_anggota. @return {Object|null} */
function getAnggota(idAnggota) {
  return getRowByField(SHEET.ANGGOTA, 'id_anggota', idAnggota);
}

/** Cek apakah anggota berstatus aktif. Error jika anggota tidak ditemukan. */
function pastikanAnggotaAktif(idAnggota) {
  const anggota = getAnggota(idAnggota);
  if (!anggota) throw new Error('Anggota ' + idAnggota + ' tidak ditemukan.');
  if (anggota.status !== 'aktif') {
    throw new Error('Anggota ' + idAnggota + ' (' + anggota.nama + ') berstatus ' +
      (anggota.status || 'nonaktif') + '.');
  }
  return anggota;
}

/**
 * Saldo kas real-time (BAGIAN 18).
 * @param {Date|null} sampaiTanggal  jika diisi, hanya hitung s/d tanggal ini
 */
function getSaldoKas(sampaiTanggal) {
  const filterFn = sampaiTanggal
    ? function(row) { return new Date(row.tanggal) <= sampaiTanggal; }
    : null;
  const masuk = sumKolom(SHEET.KAS, 'masuk', filterFn);
  const keluar = sumKolom(SHEET.KAS, 'keluar', filterFn);

  // Saldo awal kas hasil migrasi (saldo_awal jenis 'kas').
  // Selalu ikut dihitung karena mendahului semua transaksi.
  // Aman dari dobel: tutup buku TIDAK menulis baris kas baru
  // (transaksi_kas bersifat kumulatif lintas tahun).
  const saldoAwalKas = sumKolom(SHEET.SALDO_AWAL, 'nominal', function(row) {
    return row.jenis === 'kas' && !isSaldoAwalSalinan(row);
  });

  return saldoAwalKas + masuk - keluar;
}

/**
 * Peta total angsuran terbayar per pinjaman: id_pinjaman → rupiah.
 * Dihitung SEKALI per eksekusi (bukan sekali per pinjaman).
 */
function getPetaTotalAngsuran() {
  return cacheTurunan('totalAngsuran', function() {
    const peta = {};
    sheetToObjects(SHEET.ANGSURAN).forEach(function(row) {
      if (row.status_lock === STATUS_LOCK.VOID) return;
      const id = String(row.id_pinjaman);
      peta[id] = (peta[id] || 0) + (Number(row.total_bayar) || 0);
    });
    return peta;
  });
}

/** Peta total POKOK angsuran terbayar per pinjaman (tidak VOID). */
function getPetaAngsuranPokok() {
  return cacheTurunan('angsuranPokok', function() {
    const peta = {};
    sheetToObjects(SHEET.ANGSURAN).forEach(function(row) {
      if (row.status_lock === STATUS_LOCK.VOID) return;
      const id = String(row.id_pinjaman);
      peta[id] = (peta[id] || 0) + (Number(row.angsuran_pokok) || 0);
    });
    return peta;
  });
}

/** Pinjaman memakai jasa menurun (persen × sisa pokok tiap bulan)? */
function isJasaMenurun(pinjaman) {
  return String(pinjaman && pinjaman.metode_jasa || '').trim().toLowerCase() === 'menurun';
}

/**
 * Baris pokok milik satu pinjaman induk: induk + semua pinjaman
 * tambahan (id_induk), yang tidak VOID.
 */
function getBarisPokokPinjaman(idInduk) {
  const id = String(idInduk);
  return getRowsByFilter(SHEET.PINJAMAN, function(r) {
    return r.status_lock !== STATUS_LOCK.VOID &&
           (String(r.id_pinjaman) === id || String(r.id_induk || '') === id);
  });
}

/**
 * Sisa pinjaman real-time (BAGIAN 18).
 * - Jasa MENURUN: sisa POKOK = pokok dicairkan (induk + tambahan)
 *                 − SUM(angsuran_pokok). Jasa ditagih terpisah tiap bulan.
 * - Jasa FLAT (data lama): sisa = total_tagihan − SUM(total_bayar).
 * - Baris pinjaman TAMBAHAN selalu 0 (sisanya ikut induk).
 */
function getSisaPinjaman(idPinjaman) {
  const pinjaman = getRowByField(SHEET.PINJAMAN, 'id_pinjaman', idPinjaman);
  if (!pinjaman) throw new Error('Pinjaman ' + idPinjaman + ' tidak ditemukan.');

  if (isJasaMenurun(pinjaman)) {
    if (String(pinjaman.id_induk || '')) return 0;
    const pokok = getBarisPokokPinjaman(idPinjaman).reduce(function(s, r) {
      return s + (Number(r.nominal) || 0);
    }, 0);
    return pokok - (getPetaAngsuranPokok()[String(idPinjaman)] || 0);
  }

  const totalAngsuran = getPetaTotalAngsuran()[String(idPinjaman)] || 0;
  return Number(pinjaman.total_tagihan) - totalAngsuran;
}

/**
 * Sisa pokok pada AWAL bulan (tahun, bulan) — dasar hitung jasa menurun.
 * Pokok yang dicairkan dalam bulan itu & angsuran bulan itu belum dihitung
 * (sesuai Excel: jasa Januari = 1,5% × sisa per 1 Januari).
 */
function getSisaPokokAwalBulan(idInduk, tahun, bulan) {
  const awal = new Date(tahun, bulan - 1, 1);
  const pokok = getBarisPokokPinjaman(idInduk).reduce(function(s, r) {
    return new Date(r.tanggal) < awal ? s + (Number(r.nominal) || 0) : s;
  }, 0);
  const periode = tahun * 12 + bulan;
  const dibayar = getRowsByFilter(SHEET.ANGSURAN, function(a) {
    return String(a.id_pinjaman) === String(idInduk) && a.status_lock !== STATUS_LOCK.VOID &&
           Number(a.tahun) * 12 + Number(a.bulan) < periode;
  }).reduce(function(s, a) { return s + (Number(a.angsuran_pokok) || 0); }, 0);
  return pokok - dibayar;
}

/** Jasa menurun bulan (tahun, bulan) untuk satu pinjaman induk. */
function hitungJasaMenurun(pinjaman, tahun, bulan) {
  const dasar = Math.max(0, getSisaPokokAwalBulan(pinjaman.id_pinjaman, tahun, bulan));
  return Math.round(dasar * Number(pinjaman.jasa_persen || 0) / 100);
}

/** Angsuran pokok per bulan: dibulatkan KE ATAS ke ribuan (angsuran terakhir lebih kecil). */
function hitungAngsuranPokok(total, tenor) {
  return Math.ceil(Number(total) / Number(tenor) / 1000) * 1000;
}

/**
 * Tambahkan kolom baru ke sheet bila belum ada (sheet lama dibuat
 * sebelum fitur tertentu). Aman dipanggil berulang.
 */
function pastikanKolomTambahan(namaSheet, daftarKolom) {
  const header = getHeader(namaSheet).map(function(h) { return String(h).trim(); });
  const kurang = daftarKolom.filter(function(k) { return header.indexOf(k) === -1; });
  if (kurang.length === 0) return false;
  getSheet(namaSheet).getRange(1, header.length + 1, 1, kurang.length)
    .setValues([kurang]).setFontWeight('bold');
  hapusCache(namaSheet);
  return true;
}

const KOLOM_PINJAMAN_MENURUN = ['metode_jasa', 'id_induk', 'angsuran_sebelumnya'];
function pastikanKolomPinjaman() {
  return pastikanKolomTambahan(SHEET.PINJAMAN, KOLOM_PINJAMAN_MENURUN);
}

const JENIS_SALDO_AWAL_SIMPANAN = {
  simpanan_pokok: 'pokok', simpanan_wajib: 'wajib', simpanan_sukarela: 'sukarela'
};

/**
 * Peta saldo simpanan SEMUA anggota: id_anggota → {pokok, wajib, sukarela}.
 * Setiap sheet dibaca satu kali saja, lalu dipakai berulang oleh
 * dashboard, SHU, rekap, dll. (Dulu setiap panggilan getSaldoSimpanan
 * menyisir ulang 4 sheet → dashboard & SHU sangat lambat.)
 */
function getPetaSaldoSimpanan(sampaiTahun) {
  // sampaiTahun (opsional): saldo per 31 Desember tahun itu (untuk SHU tahun lalu)
  const batas = Number(sampaiTahun) || 0;
  const ikut = function(row) { return !batas || Number(row.tahun) <= batas; };
  return cacheTurunan('saldoSimpanan_' + batas, function() {
    const peta = {};
    const tambah = function(id, jenis, nilai) {
      id = String(id);
      if (!peta[id]) peta[id] = { pokok: 0, wajib: 0, sukarela: 0 };
      if (peta[id].hasOwnProperty(jenis)) peta[id][jenis] += Number(nilai) || 0;
    };

    sheetToObjects(SHEET.SIMPANAN).forEach(function(row) {
      if (row.status_lock === STATUS_LOCK.VOID || !ikut(row)) return;
      tambah(row.id_anggota, row.jenis_simpanan, row.jumlah_setoran);
    });
    sheetToObjects(SHEET.PENGAMBILAN).forEach(function(row) {
      if (row.status_lock === STATUS_LOCK.VOID || !ikut(row)) return;
      if (row.status_approval !== 'APPROVED') return;
      tambah(row.id_anggota, row.jenis_simpanan, -(Number(row.jumlah) || 0));
    });
    sheetToObjects(SHEET.JASA_SUKARELA).forEach(function(row) {
      if (row.status_lock === STATUS_LOCK.VOID || !ikut(row)) return;
      if (row.status_posting !== 'POSTED') return;
      tambah(row.id_anggota, 'sukarela', row.nominal_jasa);
    });
    sheetToObjects(SHEET.SALDO_AWAL).forEach(function(row) {
      if (row.status_lock === STATUS_LOCK.VOID || isSaldoAwalSalinan(row) || !ikut(row)) return;
      const jenis = JENIS_SALDO_AWAL_SIMPANAN[row.jenis];
      if (jenis) tambah(row.id_anggota, jenis, row.nominal);
    });
    return peta;
  });
}

/**
 * Saldo simpanan per anggota per jenis (BAGIAN 18).
 * @param {string} idAnggota
 * @param {string} jenisSimpanan  'pokok' / 'wajib' / 'sukarela'
 */
function getSaldoSimpanan(idAnggota, jenisSimpanan) {
  const saldo = getPetaSaldoSimpanan()[String(idAnggota)];
  return saldo ? (saldo[jenisSimpanan] || 0) : 0;
}

/**
 * Saldo simpanan SUKARELA per anggota, dihitung "sampai akhir bulan
 * sebelum" (tahun, bulan) tertentu — dipakai sebagai basis hitung
 * jasa sukarela (BAGIAN 2A: "basis = saldo akhir bulan sebelumnya").
 *
 * Contoh: getSaldoSukarelaAkhirBulan('AGT-0001', 2026, 3) menghitung
 * saldo sukarela per akhir Februari 2026 (SEBELUM bulan Maret),
 * dipakai untuk hitung jasa sukarela bulan Maret.
 *
 * Saldo migrasi (saldo_awal) selalu dihitung ikut (dianggap sebelum
 * semua periode transaksi sistem berjalan).
 */
function getSaldoSukarelaAkhirBulan(idAnggota, tahun, bulan) {
  const sebelumPeriode = function(row) {
    const t = Number(row.tahun), b = Number(row.bulan);
    return (t < tahun) || (t === tahun && b < bulan);
  };

  const totalSetor = sumKolom(SHEET.SIMPANAN, 'jumlah_setoran', function(row) {
    return row.id_anggota === idAnggota && row.jenis_simpanan === 'sukarela' &&
           sebelumPeriode(row);
  });

  const totalAmbil = sumKolom(SHEET.PENGAMBILAN, 'jumlah', function(row) {
    return row.id_anggota === idAnggota && row.jenis_simpanan === 'sukarela' &&
           row.status_approval === 'APPROVED' && sebelumPeriode(row);
  });

  const totalJasa = sumKolom(SHEET.JASA_SUKARELA, 'nominal_jasa', function(row) {
    return row.id_anggota === idAnggota && row.status_posting === 'POSTED' &&
           sebelumPeriode(row);
  });

  const saldoAwal = sumKolom(SHEET.SALDO_AWAL, 'nominal', function(row) {
    return row.id_anggota === idAnggota && row.jenis === 'simpanan_sukarela' &&
           !isSaldoAwalSalinan(row);
  });

  return saldoAwal + totalSetor + totalJasa - totalAmbil;
}

// ===== fitur_lanjutan.js =====
/**
 * ============================================================
 * FITUR_LANJUTAN.GS — Kuitansi, Tahun Buku, Grafik, Import
 * Koperasi App
 * ============================================================
 * Endpoint untuk 4 fitur:
 * 5. Kuitansi/bukti setoran (data untuk dicetak dari aplikasi)
 * 6. Pengaturan RAT + Cek Kesiapan + Tutup Buku dari aplikasi
 * 7. Grafik tren bulanan untuk Dashboard
 * 8. Import massal saldo_awal dari CSV
 * ============================================================
 */

// ============================================================
// 5. SETORAN + DATA KUITANSI
// ============================================================

/** Seperti apiSubmitSetoran, tapi mengembalikan data bukti untuk dicetak. */
function apiSetoranDenganBukti(d) {
  const profil = requireRole(['admin', 'operator']);
  MODE_SENYAP = true;

  const hasil = inputSetoranSimpanan({
    id_anggota: d.id_anggota, jenis_simpanan: d.jenis_simpanan,
    jumlah_setoran: Number(d.jumlah), keterangan: d.keterangan || ''
  });

  return { pesan: hasil.pesan, struk: dataStruk('setoran', hasil.id_transaksi) };
}

// ============================================================
// 6. PENGATURAN TAHUN BUKU (setting_rat) + TUTUP BUKU
// ============================================================

function apiSettingList() {
  requireRole(['admin']);
  MODE_SENYAP = true;
  return sheetToObjects(SHEET.SETTING_RAT).map(function(s) {
    const nominal = getNominalSimpanan(Number(s.tahun));
    return {
      nominal_pokok: nominal.nominal_pokok, nominal_wajib: nominal.nominal_wajib,
      minimal_sukarela: nominal.minimal_sukarela,
      tahun: Number(s.tahun),
      jasa_sukarela: Number(s.jasa_sukarela), jasa_pinjaman: Number(s.jasa_pinjaman),
      metode_jasa_sukarela: metodeJasaSukarela(s),
      shu_simpanan: Number(s.shu_simpanan), shu_jasa: Number(s.shu_jasa),
      dana_cadangan: Number(s.dana_cadangan), dana_pengurus: Number(s.dana_pengurus),
      dana_sosial: Number(s.dana_sosial),
      dana_anggota: Number(s.dana_anggota), dana_kesejahteraan: Number(s.dana_kesejahteraan),
      dana_pendidikan: Number(s.dana_pendidikan), dana_pembangunan: Number(s.dana_pembangunan),
      pajak_persen: Number(s.pajak_persen),
      aktif: s.status_aktif === true
    };
  }).sort(function(a, b) { return b.tahun - a.tahun; });
}

/** Tambah/ubah satu baris setting_rat. status_aktif TIDAK diubah di sini. */
function apiSettingSimpan(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;

  const tahun = Number(d.tahun);
  if (!tahun || tahun < 2020 || tahun > 2100) throw new Error('Tahun tidak valid.');

  pastikanKolomAlokasi();
  const calon = {
    jasa_sukarela: Number(d.jasa_sukarela), jasa_pinjaman: Number(d.jasa_pinjaman),
    shu_simpanan: Number(d.shu_simpanan), shu_jasa: Number(d.shu_jasa),
    pajak_persen: Number(d.pajak_persen || 0)
  };
  pastikanKolomJasaSukarela();
  if (d.metode_jasa_sukarela !== undefined) {
    calon.metode_jasa_sukarela = metodeJasaSukarela({ metode_jasa_sukarela: d.metode_jasa_sukarela });
  }
  POS_ALOKASI.forEach(function(p) { calon[p[0]] = Number(d[p[0]]); });
  validateAlokasiSHU(calon); // 7 pos = 100%, split simpanan/jasa = 100

  pastikanKolomNominalSimpanan();
  const lamaNominal = getNominalSimpanan(tahun);
  Object.keys(DEFAULT_NOMINAL_SIMPANAN).forEach(function(k) {
    // Form lama (tanpa isian nominal) → pertahankan nilai yang sudah ada
    const nilai = d[k] === undefined || d[k] === '' ? lamaNominal[k] : Number(d[k]);
    if (!(nilai > 0)) throw new Error('Nominal ' + k.replace(/_/g, ' ') + ' harus lebih dari 0.');
    calon[k] = nilai;
  });

  const ada = getSettingRAT(tahun);
  if (ada) {
    updateRowByField(SHEET.SETTING_RAT, 'tahun', tahun, calon);
    logAktivitas('EDIT', SHEET.SETTING_RAT, String(tahun), null, calon);
    return '✅ Setting tahun ' + tahun + ' diperbarui.';
  }

  calon.tahun = tahun;
  calon.status_aktif = false; // tahun baru selalu nonaktif dulu — diaktifkan via tutup buku
  appendRowFromObject(SHEET.SETTING_RAT, calon);
  logAktivitas('INSERT', SHEET.SETTING_RAT, String(tahun), null, calon);
  return '✅ Setting tahun ' + tahun + ' dibuat (status: belum aktif — akan aktif ' +
    'otomatis saat tutup buku tahun sebelumnya).';
}

function apiCekTutupBuku() {
  requireRole(['admin']);
  MODE_SENYAP = true;
  const tahun = getTahunAktif();
  if (!tahun) return { tahun: null, masalah: ['Tidak ada tahun aktif di setting_rat.'] };
  return { tahun: tahun, masalah: cekKesiapanTutupBuku(tahun) };
}

/** Eksekusi tutup buku. Wajib menyertakan tahun yang diketik pengguna. */
function apiTutupBukuApp(tahunKonfirmasi) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  const tahun = getTahunAktif();
  if (!tahun) throw new Error('Tidak ada tahun aktif.');
  if (Number(tahunKonfirmasi) !== Number(tahun)) {
    throw new Error('Konfirmasi tahun tidak cocok (Anda mengetik "' + tahunKonfirmasi +
      '", tahun aktif adalah ' + tahun + '). Tutup buku dibatalkan.');
  }
  const laporan = prosesTutupBuku(tahun);
  return '✅ TUTUP BUKU TAHUN ' + tahun + ' SELESAI\n\n• ' + laporan.join('\n• ') +
    '\n\nTahun ' + (tahun + 1) + ' sekarang aktif.';
}

// ============================================================
// 7. GRAFIK TREN BULANAN (untuk Dashboard)
// ============================================================

/**
 * Data per bulan tahun aktif:
 * - saldoKas[]  : saldo kas akhir tiap bulan (kumulatif)
 * - setoran[]   : total setoran simpanan per bulan
 */
function apiTrenBulanan() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;

  const tahun = getTahunAktif();
  const kasMasuk = new Array(13).fill(0), kasKeluar = new Array(13).fill(0);
  const setoran = new Array(13).fill(0);

  sheetToObjects(SHEET.KAS).forEach(function(k) {
    if (Number(k.tahun) !== tahun || k.status_lock === STATUS_LOCK.VOID) return;
    const b = Number(k.bulan);
    if (b >= 1 && b <= 12) {
      kasMasuk[b] += Number(k.masuk) || 0;
      kasKeluar[b] += Number(k.keluar) || 0;
    }
  });

  sheetToObjects(SHEET.SIMPANAN).forEach(function(s) {
    if (Number(s.tahun) !== tahun || s.status_lock === STATUS_LOCK.VOID) return;
    const b = Number(s.bulan);
    if (b >= 1 && b <= 12) setoran[b] += Number(s.jumlah_setoran) || 0;
  });

  const saldoKas = [];
  // Mulai dari saldo awal kas migrasi + akumulasi tahun-tahun sebelumnya
  let kumulatif = sumKolom(SHEET.SALDO_AWAL, 'nominal', function(r) {
    return r.jenis === 'kas';
  });
  sheetToObjects(SHEET.KAS).forEach(function(k) {
    if (Number(k.tahun) < tahun && k.status_lock !== STATUS_LOCK.VOID) {
      kumulatif += (Number(k.masuk) || 0) - (Number(k.keluar) || 0);
    }
  });
  for (let b = 1; b <= 12; b++) {
    kumulatif += kasMasuk[b] - kasKeluar[b];
    saldoKas.push(kumulatif);
  }

  return { tahun: tahun, saldoKas: saldoKas, setoran: setoran.slice(1) };
}

// ============================================================
// 8. IMPORT MASSAL SALDO AWAL (dari teks CSV)
// ============================================================

const JENIS_SALDO_VALID = {
  'simpanan_pokok': 'simpanan_pokok', 'pokok': 'simpanan_pokok',
  'simpanan_wajib': 'simpanan_wajib', 'wajib': 'simpanan_wajib',
  'simpanan_sukarela': 'simpanan_sukarela', 'sukarela': 'simpanan_sukarela',
  'piutang': 'piutang',
  // ---- milik KOPERASI (tanpa anggota) — posisi neraca saat mulai memakai aplikasi ----
  'kas': 'kas',
  'inventaris': 'inventaris', 'penyusutan': 'penyusutan',
  'cadangan': 'cadangan', 'modal': 'modal',
  'dana_anggota': 'dana_anggota', 'dana_pengurus': 'dana_pengurus',
  'dana_kesejahteraan': 'dana_kesejahteraan', 'dana_pendidikan': 'dana_pendidikan',
  'dana_sosial': 'dana_sosial', 'dana_pembangunan': 'dana_pembangunan',
  'shu': 'shu', 'ekuitas': 'shu'   // SHU belum dibagi / penyeimbang ekuitas lama
};
/**
 * Angka dari tempelan Excel/CSV: "4.370.495,82" (format Indonesia), "4370495.82",
 * "4,370,495.82", "Rp 1.000.000", "10000". Hasil dibulatkan ke sen.
 */
function angkaImpor(v) {
  if (typeof v === 'number') return Math.round(v * 100) / 100;
  let s = String(v || '').replace(/rp|\s/gi, '');
  if (!s) return 0;
  const koma = s.lastIndexOf(','), titik = s.lastIndexOf('.');
  if (koma > -1 && titik > -1) {
    // pemisah yang paling kanan = desimal
    s = koma > titik ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  } else if (koma > -1) {
    s = /^-?\d{1,3}(,\d{3})+$/.test(s) ? s.replace(/,/g, '') : s.replace(',', '.');
  } else if (titik > -1 && /^-?\d{1,3}(\.\d{3})+$/.test(s)) {
    s = s.replace(/\./g, '');            // 1.000.000 = ribuan
  }
  const n = Number(s);
  return isNaN(n) ? NaN : Math.round(n * 100) / 100;
}

const JENIS_SALDO_ANGGOTA = { simpanan_pokok: 1, simpanan_wajib: 1, simpanan_sukarela: 1, piutang: 1 };

/**
 * @param {Object} d { tahun, baris: [{id_anggota, jenis, nominal}, ...] }
 */
function apiImportSaldoAwal(d) {
  const profil = requireRole(['admin']);
  MODE_SENYAP = true;

  const tahun = Number(d.tahun) || getTahunAktif();
  const baris = d.baris || [];
  if (baris.length === 0) throw new Error('Tidak ada baris untuk diimpor.');
  if (baris.length > 500) throw new Error('Maksimal 500 baris per impor.');

  const hasil = { sukses: 0, gagal: [] };

  // Peta nama → id untuk pencocokan berbasis nama (dari Excel lama)
  const norm = function(t) {
    return String(t || '').toLowerCase().replace(/\s+/g, ' ').trim();
  };
  const petaNama = {};
  sheetToObjects(SHEET.ANGGOTA).forEach(function(a) {
    petaNama[norm(a.nama)] = a.id_anggota;
  });

  baris.forEach(function(r, i) {
    const nomorBaris = i + 1;
    try {
      let idAnggota = String(r.id_anggota || '').trim();
      const jenisRaw = String(r.jenis || '').trim().toLowerCase().replace(/\s+/g, '_');
      const nominal = angkaImpor(r.nominal);

      // Baris milik KOPERASI (kas / shu-ekuitas): tidak terikat anggota
      if (!JENIS_SALDO_VALID[jenisRaw]) throw new Error('jenis "' + r.jenis + '" tidak dikenal ' +
        '(pakai: pokok/wajib/sukarela/piutang, atau untuk koperasi: kas/inventaris/' +
        'penyusutan/cadangan/modal/dana_anggota/dana_pengurus/dana_kesejahteraan/' +
        'dana_pendidikan/dana_sosial/dana_pembangunan/shu)');
      const tanpaAnggota = !JENIS_SALDO_ANGGOTA[JENIS_SALDO_VALID[jenisRaw]];
      if (tanpaAnggota) {
        idAnggota = '-';
      } else if (!getAnggota(idAnggota)) {
        // Cocokkan: coba sebagai ID dulu, lalu sebagai NAMA
        const dariNama = petaNama[norm(idAnggota)];
        if (dariNama) {
          idAnggota = dariNama;
        } else {
          throw new Error('anggota "' + r.id_anggota +
            '" tidak ditemukan (baik sebagai ID maupun nama)');
        }
      }
      const jenis = JENIS_SALDO_VALID[jenisRaw];
      if (!jenis) throw new Error('jenis "' + r.jenis + '" tidak dikenal ' +
        '(pakai: pokok/wajib/sukarela/piutang, atau untuk koperasi: kas/inventaris/' +
        'penyusutan/cadangan/modal/dana_anggota/dana_pengurus/dana_kesejahteraan/' +
        'dana_pendidikan/dana_sosial/dana_pembangunan/shu)');
      if (!nominal || nominal <= 0) throw new Error('nominal tidak valid');

      const idSal = generateId(SHEET.SALDO_AWAL, ID_PREFIX.SALDO_AWAL, 1, tahun, null);
      appendRowFromObject(SHEET.SALDO_AWAL, {
        id: idSal, tahun: tahun, id_anggota: idAnggota,
        jenis: jenis, nominal: nominal,
        keterangan: 'Import CSV oleh ' + profil.email
      });
      hasil.sukses++;
    } catch (e) {
      hasil.gagal.push('Baris ' + nomorBaris + ': ' + e.message);
    }
  });

  logAktivitas('INSERT', SHEET.SALDO_AWAL, 'IMPORT-' + tahun, null,
    { sukses: hasil.sukses, gagal: hasil.gagal.length });
  const cek = cekSaldoAwal();

  return '✅ Import selesai.\nBerhasil: ' + hasil.sukses + ' baris' +
    (hasil.gagal.length > 0
      ? '\nGagal: ' + hasil.gagal.length + '\n' + hasil.gagal.join('\n') +
        '\n\n(Baris yang gagal tidak tersimpan — perbaiki lalu impor ulang ' +
        'HANYA baris tersebut agar tidak dobel)'
      : '') +
    '\n\nKeseimbangan saldo awal: aktiva ' + formatRupiah(cek.aktiva) + ' · pasiva ' +
    formatRupiah(cek.pasiva) + (cek.balance ? ' ✅ BALANCE'
      : ' ❌ selisih ' + formatRupiah(cek.selisih) + ' (tambahkan jenis "shu" sebagai penyeimbang ' +
        'bila ada ekuitas lama yang belum tercatat)') +
    '\n\n⚠️ Untuk jenis "piutang", jangan lupa buat juga pinjaman MIGRASI ' +
    'di form Transaksi → tab Migrasi, lalu cek dengan Validasi Migrasi.';
}

// ============================================================
// 8b. IMPORT MASSAL PINJAMAN MIGRASI
// ============================================================

/**
 * @param {Object} d { baris: [{nama, sisa, tenor, bulan, tahun}, ...] }
 */
function apiImportPinjamanMigrasi(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;

  const baris = d.baris || [];
  if (baris.length === 0) throw new Error('Tidak ada baris untuk diimpor.');
  if (baris.length > 200) throw new Error('Maksimal 200 baris per impor.');

  const norm = function(t) {
    return String(t || '').toLowerCase().replace(/\s+/g, ' ').trim();
  };
  const petaNama = {};
  sheetToObjects(SHEET.ANGGOTA).forEach(function(a) {
    petaNama[norm(a.nama)] = a.id_anggota;
  });

  // Pengaman anti-dobel: anggota yang SUDAH punya pinjaman MIGRASI aktif
  const sudahMigrasi = {};
  sheetToObjects(SHEET.PINJAMAN).forEach(function(p) {
    if (p.status_lock !== STATUS_LOCK.VOID &&
        String(p.keterangan || '').indexOf('MIGRASI') === 0) {
      sudahMigrasi[p.id_anggota] = true;
    }
  });

  const hasil = { sukses: 0, gagal: [] };

  baris.forEach(function(r, i) {
    const nomorBaris = i + 1;
    try {
      let id = String(r.nama || '').trim();
      if (!getAnggota(id)) {
        const dariNama = petaNama[norm(id)];
        if (!dariNama) throw new Error('anggota "' + r.nama + '" tidak ditemukan');
        id = dariNama;
      }
      if (sudahMigrasi[id]) {
        throw new Error('"' + r.nama + '" sudah punya pinjaman MIGRASI aktif — dilewati ' +
          'agar tidak dobel');
      }

      const sisa = angkaImpor(r.sisa);
      const tenor = Number(r.tenor);
      const bulan = Number(r.bulan) || 12;
      const tahun = Number(r.tahun) || (getTahunAktif() - 1);
      if (!sisa || sisa <= 0) throw new Error('sisa pinjaman tidak valid');
      if (!tenor || tenor <= 0) {
        throw new Error('tenor kosong/tidak valid — isi dulu kolom tenor_sisa');
      }

      const angsuran = Number(String(r.angsuran || '').replace(/[.\s]/g, '')) || 0;
      inputPinjamanMigrasi({
        id_anggota: id, sisa_pokok: sisa, tenor_sisa: tenor, angsuran_pokok: angsuran,
        tanggal_asli: new Date(tahun, bulan - 1, 1),
        keterangan_tambahan: 'Import massal'
      });
      sudahMigrasi[id] = true; // jaga-jaga kalau satu nama muncul dua kali di CSV
      hasil.sukses++;
    } catch (e) {
      hasil.gagal.push('Baris ' + nomorBaris + ': ' + e.message);
    }
  });

  logAktivitas('INSERT', SHEET.PINJAMAN, 'IMPORT-MIGRASI', null,
    { sukses: hasil.sukses, gagal: hasil.gagal.length });

  return '✅ Import pinjaman migrasi selesai.\nBerhasil: ' + hasil.sukses + ' pinjaman' +
    (hasil.gagal.length > 0
      ? '\nGagal/dilewati: ' + hasil.gagal.length + '\n' + hasil.gagal.join('\n')
      : '') +
    '\n\nLangkah berikutnya: jalankan Validasi Migrasi (menu 🔧 SETUP di spreadsheet) ' +
    'untuk memastikan total pinjaman MIGRASI = total saldo awal piutang.';
}

// ===== jasa.js =====
/**
 * ============================================================
 * JASA.GS — Hitung & Posting Jasa Sukarela Bulanan
 * Koperasi App
 * Sesuai dokumen: REKAP_Koperasi_v2.md (BAGIAN 2A, 3, 12, 15, 17, 18)
 * ============================================================
 *
 * DUA TAHAP TERPISAH (SENGAJA, bukan satu langkah):
 *
 * 1. hitungJasaSukarelaBulanan()
 *    Menghitung jasa untuk semua anggota, status = DRAFT.
 *    Masih BISA dikoreksi/dihapus manual di sheet jika ada yang salah.
 *    BELUM masuk jurnal, BELUM menambah saldo simpanan siapapun.
 *
 * 2. postingJasaSukarela()
 *    Mengunci semua baris DRAFT bulan ini jadi POSTED.
 *    SETELAH ini: jurnal dibuat, saldo sukarela anggota bertambah,
 *    TIDAK BISA diubah lagi tanpa proses VOID.
 *
 * Alur normal (checklist BAGIAN 15, akhir bulan tgl 25-31):
 * [ ] 1. Klik "Hitung Jasa Sukarela Bulan Ini" → cek angkanya di sheet
 * [ ] 2. Kalau sudah benar → Klik "Posting Jasa Sukarela (Kunci)"
 * ============================================================
 */

// ============================================================
// TAHAP 1: HITUNG (DRAFT)
// ============================================================

/**
 * Hitung jasa sukarela untuk SEMUA anggota aktif, bulan & tahun
 * tertentu. Hasilnya berstatus DRAFT — masih boleh dikoreksi.
 *
 * @param {number} [tahun]  default: tahun aktif
 * @param {number} [bulan]  default: bulan berjalan sekarang
 */
function hitungJasaSukarelaBulanan(tahun, bulan) {
  const sekarang = new Date();
  tahun = tahun || getTahunAktif();
  bulan = bulan || (sekarang.getMonth() + 1);

  const setting = validateTahunAktif(tahun);
  pastikanKolomJasaSukarela();
  const persenJasa = Number(setting.jasa_sukarela);
  const metode = metodeJasaSukarela(setting);

  const anggotaAktif = getRowsByFilter(SHEET.ANGGOTA, function(row) {
    return row.status === 'aktif';
  });

  const hasil = { dibuat: [], dilewatiSudahAda: [], dilewatiSaldoNol: [] };

  anggotaAktif.forEach(function(anggota) {
    // Lewati jika bulan ini SUDAH pernah dihitung untuk anggota ini
    const sudahAda = getRowsByFilter(SHEET.JASA_SUKARELA, function(row) {
      return row.id_anggota === anggota.id_anggota &&
             Number(row.tahun) === Number(tahun) &&
             Number(row.bulan) === Number(bulan);
    });
    if (sudahAda.length > 0) {
      hasil.dilewatiSudahAda.push(anggota.nama + ' (' + sudahAda[0].status_posting + ')');
      return;
    }

    const j = hitungJasaSukarelaAnggota(anggota.id_anggota, tahun, bulan, setting);
    const saldoAwalBulan = j.saldoAwal;

    // Aturan: tidak diberikan jika dasar hitung = 0 (anggota baru / saldo sempat habis)
    if (j.dasar <= 0) {
      hasil.dilewatiSaldoNol.push(anggota.nama);
      return;
    }

    const nominalJasa = j.nominal;

    const idJasa = generateId(SHEET.JASA_SUKARELA, ID_PREFIX.JASA_SUKARELA, 1, tahun, bulan);

    const row = {
      id: idJasa,
      tahun: tahun,
      bulan: bulan,
      id_anggota: anggota.id_anggota,
      saldo_awal_bulan: saldoAwalBulan,
      dasar_jasa: j.dasar,
      metode_jasa: metode,
      persen_jasa: persenJasa,
      nominal_jasa: nominalJasa,
      status_posting: 'DRAFT'
    };
    appendRowFromObject(SHEET.JASA_SUKARELA, row);
    logAktivitas('INSERT', SHEET.JASA_SUKARELA, idJasa, null, row);

    hasil.dibuat.push(anggota.nama + ': ' + formatRupiah(nominalJasa));
  });

  const pesan =
    '✅ HITUNG JASA SUKARELA — Tahun ' + tahun + ' Bulan ' + bulan + ' (status: DRAFT)\n' +
    'Rumus: ' + KET_METODE_JASA[metode].replace('{p}', persenJasa) + '\n\n' +
    'Dibuat: ' + hasil.dibuat.length + '\n' +
    (hasil.dibuat.length ? hasil.dibuat.join('\n') + '\n\n' : '\n') +
    'Dilewati (saldo 0 / anggota baru): ' + hasil.dilewatiSaldoNol.length + '\n' +
    'Dilewati (sudah pernah dihitung): ' + hasil.dilewatiSudahAda.length + '\n\n' +
    '➡️ Cek angkanya di sheet rekap_jasa_sukarela. Jika sudah benar, ' +
    'jalankan "Posting Jasa Sukarela (Kunci)" untuk mengunci & membuat jurnal.';

  tampilkanPesan(pesan);
  return hasil;
}

// ============================================================
// RUMUS JASA SUKARELA (bisa berubah tiap RAT → dipilih per tahun buku)
// ============================================================

/**
 * 'terendah' (bawaan, RAT 2026): jasa = saldo sukarela TERENDAH pada bulan itu
 *            × persen/100. Persen dibaca PER BULAN.
 * 'harian'  (aturan lama):      jasa = saldo akhir bulan lalu × persen/100 × 30/365.
 *            Persen dibaca PER TAHUN.
 */
const KET_METODE_JASA = {
  terendah: 'saldo sukarela terendah bulan ini × {p}% per bulan',
  harian: 'saldo akhir bulan lalu × {p}% per tahun × 30/365'
};
function metodeJasaSukarela(setting) {
  const m = String((setting && setting.metode_jasa_sukarela) || '').trim().toLowerCase();
  return KET_METODE_JASA[m] ? m : 'terendah';
}

/** Kolom baru di setting_rat & rekap_jasa_sukarela (sheet lama belum punya). */
function pastikanKolomJasaSukarela() {
  tambahKolomBila_(SHEET.SETTING_RAT, ['metode_jasa_sukarela']);
  tambahKolomBila_(SHEET.JASA_SUKARELA, ['dasar_jasa', 'metode_jasa']);
}
function tambahKolomBila_(namaSheet, kolom) {
  const header = getHeader(namaSheet).map(function(h) { return String(h).trim(); });
  const kurang = kolom.filter(function(k) { return header.indexOf(k) === -1; });
  if (!kurang.length) return;
  getSheet(namaSheet).getRange(1, header.length + 1, 1, kurang.length).setValues([kurang])
    .setFontWeight('bold');
  hapusCache(namaSheet);
}

/**
 * Saldo sukarela TERENDAH selama satu bulan: dimulai dari saldo akhir bulan
 * lalu, lalu saldo di akhir setiap hari yang ada setoran / pengambilan
 * sukarela (yang sudah disetujui). Jasa bulan itu sendiri tidak dihitung.
 */
function getSaldoSukarelaTerendahBulan(idAnggota, tahun, bulan) {
  const awal = getSaldoSukarelaAkhirBulan(idAnggota, tahun, bulan);
  const bulanIni = function(r) {
    return String(r.id_anggota) === String(idAnggota) && r.status_lock !== STATUS_LOCK.VOID &&
      Number(r.tahun) === Number(tahun) && Number(r.bulan) === Number(bulan);
  };
  const hari = function(tgl) {
    const d = new Date(tgl);
    return isNaN(d.getTime()) ? 1 : d.getDate();
  };
  const mutasi = {};   // tanggal → perubahan saldo hari itu
  getRowsByFilter(SHEET.SIMPANAN, function(r) {
    return bulanIni(r) && r.jenis_simpanan === 'sukarela';
  }).forEach(function(r) {
    const h = hari(r.tanggal); mutasi[h] = (mutasi[h] || 0) + (Number(r.jumlah_setoran) || 0);
  });
  getRowsByFilter(SHEET.PENGAMBILAN, function(r) {
    return bulanIni(r) && r.jenis_simpanan === 'sukarela' && r.status_approval === 'APPROVED';
  }).forEach(function(r) {
    const h = hari(r.tanggal); mutasi[h] = (mutasi[h] || 0) - (Number(r.jumlah) || 0);
  });
  let saldo = awal, terendah = awal;
  Object.keys(mutasi).map(Number).sort(function(a, b) { return a - b; }).forEach(function(h) {
    saldo += mutasi[h];
    if (saldo < terendah) terendah = saldo;
  });
  return { saldoAwal: awal, terendah: terendah };
}

/**
 * Jasa sukarela satu anggota untuk satu bulan, sesuai rumus tahun buku itu.
 * @return {{saldoAwal:number, dasar:number, nominal:number, metode:string}}
 */
function hitungJasaSukarelaAnggota(idAnggota, tahun, bulan, setting) {
  setting = setting || getSettingRAT(tahun) || {};
  const persen = Number(setting.jasa_sukarela) || 0;
  const metode = metodeJasaSukarela(setting);
  if (metode === 'harian') {
    const saldoAwal = getSaldoSukarelaAkhirBulan(idAnggota, tahun, bulan);
    return { saldoAwal: saldoAwal, dasar: saldoAwal, metode: metode,
      nominal: saldoAwal > 0 ? Math.round(saldoAwal * persen / 100 * 30 / 365) : 0 };
  }
  const s = getSaldoSukarelaTerendahBulan(idAnggota, tahun, bulan);
  return { saldoAwal: s.saldoAwal, dasar: s.terendah, metode: metode,
    nominal: s.terendah > 0 ? Math.round(s.terendah * persen / 100) : 0 };
}

// ============================================================
// TAHAP 2: POSTING (KUNCI — DRAFT → POSTED)
// ============================================================

/**
 * Kunci semua baris DRAFT bulan & tahun tertentu jadi POSTED.
 * Setelah ini: jurnal dibuat, saldo sukarela anggota bertambah.
 * TIDAK BISA diubah lagi tanpa proses VOID manual.
 *
 * @param {number} [tahun]  default: tahun aktif
 * @param {number} [bulan]  default: bulan berjalan sekarang
 */
function postingJasaSukarela(tahun, bulan) {
  const sekarang = new Date();
  tahun = tahun || getTahunAktif();
  bulan = bulan || (sekarang.getMonth() + 1);

  validateTahunAktif(tahun);

  const draftBulanIni = getRowsByFilter(SHEET.JASA_SUKARELA, function(row) {
    return Number(row.tahun) === Number(tahun) && Number(row.bulan) === Number(bulan) &&
           row.status_posting === 'DRAFT';
  });

  if (draftBulanIni.length === 0) {
    tampilkanPesan('Tidak ada data DRAFT untuk tahun ' + tahun + ' bulan ' + bulan +
      '.\n\nJalankan "Hitung Jasa Sukarela Bulan Ini" dulu sebelum posting, ' +
      'atau mungkin bulan ini sudah pernah di-posting sebelumnya.');
    return { diposting: 0 };
  }

  const tanggalPosting = new Date(tahun, bulan - 1, 28); // akhir bulan, sesuai checklist tgl 25-31

  draftBulanIni.forEach(function(row) {
    const anggota = getAnggota(row.id_anggota);
    const namaAnggota = anggota ? anggota.nama : row.id_anggota;

    // Update status ke POSTED
    updateRowByField(SHEET.JASA_SUKARELA, 'id', row.id, { status_posting: 'POSTED' });

    // Jurnal: Debit Beban Jasa Simpanan Sukarela (biaya umum, seperti Excel
    // "Biaya Operasional"), Kredit Simpanan Sukarela — jasa MENAMBAH saldo
    // sukarela anggota (non-tunai).
    catatJurnal(tanggalPosting, tahun, bulan, row.id,
      'Posting jasa sukarela — ' + namaAnggota,
      [
        { kode_akun: AKUN.BEBAN_JASA_SUKARELA, debit: row.nominal_jasa, kredit: 0 },
        { kode_akun: AKUN.SIMPANAN_SUKARELA, debit: 0, kredit: row.nominal_jasa }
      ]
    );

    logAktivitas('EDIT', SHEET.JASA_SUKARELA, row.id,
      { status_posting: 'DRAFT' }, { status_posting: 'POSTED' });
  });

  const totalJasa = draftBulanIni.reduce(function(s, r) { return s + Number(r.nominal_jasa); }, 0);

  tampilkanPesan(
    '✅ POSTING SELESAI — Tahun ' + tahun + ' Bulan ' + bulan + '\n\n' +
    'Jumlah anggota diposting: ' + draftBulanIni.length + '\n' +
    'Total jasa sukarela: ' + formatRupiah(totalJasa) + '\n\n' +
    '⚠️ Data ini sekarang TERKUNCI. Saldo simpanan sukarela anggota ' +
    'sudah bertambah sesuai jasa masing-masing.'
  );

  return { diposting: draftBulanIni.length, totalJasa: totalJasa };
}

// ============================================================
// PEMBUNGKUS UNTUK MENU (tanpa parameter)
// ============================================================

function menuHitungJasaSukarela() {
  hitungJasaSukarelaBulanan();
}

function menuPostingJasaSukarela() {
  const ui = SpreadsheetApp.getUi();
  const jawab = ui.alert(
    '⚠️ Konfirmasi Posting',
    'Posting akan MENGUNCI data jasa sukarela bulan ini dan membuat jurnal. ' +
    'Setelah ini tidak bisa diubah lagi tanpa proses VOID.\n\nLanjutkan?',
    ui.ButtonSet.YES_NO
  );
  if (jawab !== ui.Button.YES) return;
  postingJasaSukarela();
}

// ===== kas_beban.js =====
/**
 * ============================================================
 * KAS_BEBAN.GS — Mutasi Kas & Beban Operasional
 * Koperasi App
 * Sesuai dokumen: REKAP_Koperasi_v2.md (BAGIAN 3, 12 #8)
 * ============================================================
 *
 * 1. inputBebanOperasional() — mencatat pengeluaran kas untuk
 *    beban/kesejahteraan (transport jenguk, THR, uang duduk RAT,
 *    souvenir, insentif, operasional umum).
 *    Jurnal #8: Debit akun 50x, Kredit 101 Kas.
 *    Tidak ada sheet baru — tercatat sebagai baris transaksi_kas
 *    (kategori "operasional") + jurnal, sesuai desain.
 *
 * 2. apiMutasiKas() — daftar mutasi kas untuk halaman Kas di
 *    aplikasi, bisa difilter tahun & bulan.
 * ============================================================
 */

/**
 * Kategori pengeluaran kas (mengikuti "Rekapitulasi Pengeluaran" Excel).
 * Semua dicatat: Debit akun, Kredit Kas.
 * - Biaya (5xx)            → mengurangi SHU tahun berjalan
 * - Pengeluaran dana (21x) → memakai dana hasil alokasi SHU tahun lalu
 * - Inventaris (111)       → aset, bukan biaya
 */
const AKUN_BEBAN = [
  { kode: '501', label: 'Biaya Operasional', grup: 'Biaya-biaya Umum' },
  { kode: '511', label: 'Brosur RAT', grup: 'Biaya-biaya Umum' },
  { kode: '512', label: 'Administrasi/ATK', grup: 'Biaya-biaya Umum' },
  { kode: '513', label: 'Konsumsi RAT', grup: 'Biaya-biaya Umum' },
  { kode: '514', label: 'Pengolahan Neraca', grup: 'Biaya-biaya Umum' },
  { kode: '515', label: 'Rapat Pengurus', grup: 'Biaya-biaya Umum' },
  { kode: '507', label: 'Uang Duduk', grup: 'Biaya-biaya Umum' },
  { kode: '516', label: 'Door Prize', grup: 'Biaya-biaya Umum' },
  { kode: '508', label: 'Souvenir RAT', grup: 'Biaya-biaya Umum' },
  { kode: '517', label: 'Transport Belanja', grup: 'Biaya-biaya Umum' },
  { kode: '505', label: 'Transport Kesejahteraan (jenguk/melahirkan)', grup: 'Biaya-biaya Umum' },
  { kode: '518', label: 'Honor BP & Pembina', grup: 'Biaya-biaya Umum' },
  { kode: '520', label: 'Biaya Lain-lain', grup: 'Biaya-biaya Umum' },
  { kode: '509', label: 'Honor Pengurus', grup: 'Honor & THR' },
  { kode: '506', label: 'THR', grup: 'Honor & THR' },
  { kode: '211', label: 'Dana Anggota (pembagian SHU)', grup: 'Pengeluaran Dana-dana (dari SHU)' },
  { kode: '212', label: 'Dana Pengurus', grup: 'Pengeluaran Dana-dana (dari SHU)' },
  { kode: '213', label: 'Dana Kesejahteraan Pegawai', grup: 'Pengeluaran Dana-dana (dari SHU)' },
  { kode: '214', label: 'Dana Pendidikan', grup: 'Pengeluaran Dana-dana (dari SHU)' },
  { kode: '215', label: 'Dana Sosial', grup: 'Pengeluaran Dana-dana (dari SHU)' },
  { kode: '216', label: 'Dana Pembangunan Daerah Kerja', grup: 'Pengeluaran Dana-dana (dari SHU)' },
  { kode: '111', label: 'Pembelian Inventaris', grup: 'Aset' }
];

/**
 * Catat satu beban operasional: kas keluar + jurnal.
 * @param {Object} d { kode_akun, nominal, keterangan }
 */
function inputBebanOperasional(d) {
  const tanggal = d.tanggal ? new Date(d.tanggal) : new Date();   // tanggal: migrasi data lama
  const tahun = tanggal.getFullYear();
  const bulan = tanggal.getMonth() + 1;

  validateTahunAktif(tahun);

  const akun = AKUN_BEBAN.find(function(a) { return a.kode === String(d.kode_akun); });
  if (!akun) throw new Error('Kategori beban tidak dikenal: ' + d.kode_akun);

  const nominal = Number(d.nominal);
  if (!nominal || nominal <= 0) throw new Error('Nominal beban harus lebih dari 0.');
  if (!d.keterangan || String(d.keterangan).trim() === '') {
    throw new Error('Keterangan wajib diisi untuk pengeluaran ' +
      '(contoh: "Transport jenguk Bu Nina").');
  }

  const saldoKas = getSaldoKas();
  if (nominal > saldoKas) {
    throw new Error('Saldo kas (' + formatRupiah(saldoKas) +
      ') tidak cukup untuk pengeluaran ' + formatRupiah(nominal) + '.');
  }
  pastikanCOA();

  // Pengeluaran dana: beri peringatan bila melebihi saldo dana (tidak diblokir,
  // karena saldo dana lama dari Excel mungkin belum dimigrasi)
  let peringatan = '';
  if (/^21/.test(akun.kode)) {
    const n = hitungNeraca(tahun);
    const dana = n.danaDana.find(function(r) { return r.kode === akun.kode; });
    const saldoDana = dana ? dana.nilai : 0;
    if (nominal > saldoDana) {
      peringatan = '\n⚠️ Melebihi saldo ' + akun.label + ' (' + formatRupiah(saldoDana) +
        '). Periksa kembali atau catat saldo dana lama.';
    }
  }

  // Kas keluar — idKas sekaligus jadi ID pencatatan beban ini
  const idKas = catatKasKeluar({
    tanggal: tanggal, tahun: tahun, bulan: bulan,
    kategori: 'operasional',
    referensi: akun.kode,
    keterangan: '[' + akun.label + '] ' + String(d.keterangan).trim(),
    nominal: nominal
  });

  // Jurnal #8: Debit akun beban, Kredit Kas
  catatJurnal(tanggal, tahun, bulan, idKas,
    akun.label + ' — ' + String(d.keterangan).trim(),
    [
      { kode_akun: akun.kode, debit: nominal, kredit: 0 },
      { kode_akun: AKUN.KAS, debit: 0, kredit: nominal }
    ]
  );

  logAktivitas('INSERT', SHEET.KAS, idKas, null,
    { jenis: 'beban', akun: akun.kode, nominal: nominal, keterangan: d.keterangan });

  return {
    id_kas: idKas,
    pesan: '✅ ' + akun.label + ' sebesar ' + formatRupiah(nominal) +
      ' tercatat.\nSisa kas: ' + formatRupiah(getSaldoKas()) + peringatan
  };
}

/**
 * Pemasukan lain-lain (bukan simpanan/angsuran), mis. bunga bank,
 * sisa konsumsi RAT. Kas masuk + jurnal Debit Kas, Kredit 403.
 */
function inputPemasukanLain(d) {
  const tanggal = d.tanggal ? new Date(d.tanggal) : new Date();
  const tahun = tanggal.getFullYear(), bulan = tanggal.getMonth() + 1;
  validateTahunAktif(tahun);
  const nominal = Number(d.nominal);
  if (!nominal || nominal <= 0) throw new Error('Nominal pemasukan harus lebih dari 0.');
  const ket = String(d.keterangan || '').trim();
  if (!ket) throw new Error('Keterangan wajib diisi (contoh: "Bunga bank").');
  pastikanCOA();

  const idKas = catatKasMasuk({ tanggal: tanggal, tahun: tahun, bulan: bulan,
    kategori: 'pendapatan_lain', referensi: AKUN.PENDAPATAN_LAIN,
    keterangan: '[Penghasilan Lain-lain] ' + ket, nominal: nominal });
  catatJurnal(tanggal, tahun, bulan, idKas, 'Penghasilan lain-lain — ' + ket, [
    { kode_akun: AKUN.KAS, debit: nominal, kredit: 0 },
    { kode_akun: AKUN.PENDAPATAN_LAIN, debit: 0, kredit: nominal }
  ]);
  logAktivitas('INSERT', SHEET.KAS, idKas, null, { jenis: 'pemasukan_lain', nominal: nominal, keterangan: ket });
  return { id_kas: idKas, pesan: '✅ Penghasilan lain-lain ' + formatRupiah(nominal) +
    ' tercatat.\nSaldo kas: ' + formatRupiah(getSaldoKas()) };
}

/**
 * Penyusutan inventaris (tanpa kas): Debit Beban Penyusutan (519),
 * Kredit Akumulasi Penyusutan (112).
 */
function inputPenyusutan(d) {
  const tanggal = d.tanggal ? new Date(d.tanggal) : new Date();
  const tahun = tanggal.getFullYear(), bulan = tanggal.getMonth() + 1;
  validateTahunAktif(tahun);
  const nominal = Number(d.nominal);
  if (!nominal || nominal <= 0) throw new Error('Nominal penyusutan harus lebih dari 0.');
  pastikanCOA();
  const n = hitungNeraca(tahun);
  const nilaiBuku = n.aktivaTetap.reduce(function(s, r) { return s + r.nilai; }, 0);
  if (nominal > nilaiBuku) {
    throw new Error('Penyusutan melebihi nilai buku inventaris (' + formatRupiah(nilaiBuku) + ').');
  }
  const ref = 'SUSUT-' + tahun + '-' + Utilities.getUuid().slice(0, 8);
  const ket = String(d.keterangan || '').trim() || 'Penyusutan inventaris ' + tahun;
  catatJurnal(tanggal, tahun, bulan, ref, ket, [
    { kode_akun: AKUN.BEBAN_PENYUSUTAN, debit: nominal, kredit: 0 },
    { kode_akun: AKUN.AKUM_PENYUSUTAN, debit: 0, kredit: nominal }
  ]);
  logAktivitas('INSERT', SHEET.JURNAL, ref, null, { jenis: 'penyusutan', nominal: nominal, keterangan: ket });
  return { pesan: '✅ Penyusutan ' + formatRupiah(nominal) + ' dijurnal (tanpa kas).' };
}

// ============================================================
// API UNTUK APLIKASI WEB
// ============================================================

function apiSubmitBeban(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  return inputBebanOperasional(d).pesan;
}

function apiDaftarAkunBeban() {
  requireRole(['admin', 'operator']);
  return AKUN_BEBAN;
}

function apiSubmitPemasukanLain(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  return inputPemasukanLain(d).pesan;
}

function apiSubmitPenyusutan(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  return inputPenyusutan(d).pesan;
}

/**
 * Mutasi kas untuk halaman Kas.
 * @param {Object} f { tahun, bulan } — bulan 0 = semua bulan
 */
function apiMutasiKas(f) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;

  const tahun = Number(f.tahun) || getTahunAktif();
  const bulan = Number(f.bulan) || 0;

  const rows = getRowsByFilter(SHEET.KAS, function(k) {
    if (Number(k.tahun) !== tahun) return false;
    if (bulan > 0 && Number(k.bulan) !== bulan) return false;
    return true;
  });

  let totMasuk = 0, totKeluar = 0;
  const daftar = rows.map(function(k) {
    const masuk = Number(k.masuk) || 0;
    const keluar = Number(k.keluar) || 0;
    if (k.status_lock !== STATUS_LOCK.VOID) {
      totMasuk += masuk; totKeluar += keluar;
    }
    return {
      id: k.id_kas,
      tgl: Utilities.formatDate(new Date(k.tanggal),
        Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      kategori: k.kategori,
      keterangan: k.keterangan || '',
      referensi: k.referensi || '',
      masuk: masuk, keluar: keluar,
      lock: k.status_lock
    };
  }).reverse(); // terbaru di atas

  return {
    tahun: tahun, bulan: bulan,
    rows: daftar,
    totMasuk: totMasuk, totKeluar: totKeluar,
    saldoKasTotal: getSaldoKas()
  };
}

// ===== keluar_anggota.js =====
/**
 * ============================================================
 * KELUAR_ANGGOTA.GS — Proses Anggota Keluar
 * Koperasi App
 * ============================================================
 *
 * Dipakai saat anggota berhenti (pindah tugas, pensiun, lulus).
 *
 * ALUR (semua dalam satu proses, khusus admin):
 * 1. Rekap: total simpanan (pokok + wajib + sukarela) dan sisa
 *    semua pinjaman aktif → uang bersih = simpanan − sisa pinjaman.
 * 2. Pinjaman aktif DILUNASI dari simpanan: dicatat sebagai
 *    pembayaran angsuran (kas masuk + jurnal, porsi pokok/jasa
 *    sama seperti pembayaran manual).
 * 3. Seluruh simpanan DIKEMBALIKAN: satu baris transaksi_pengambilan
 *    per jenis (status APPROVED) + kas keluar + jurnal
 *    Debit 201/202/203, Kredit 101 Kas.
 * 4. Status anggota → 'keluar'.
 *
 * Kas bersih yang berubah = uang bersih (simpanan − utang).
 * Jika utang > simpanan, anggota harus menyetor kekurangannya
 * (tercatat sebagai bagian dari pelunasan).
 *
 * Salah proses? Setiap baris (angsuran pelunasan & pengambilan)
 * bisa di-VOID satu per satu seperti transaksi biasa, lalu admin
 * mengaktifkan kembali anggota dari daftar anggota.
 * ============================================================
 */

/** Kode akun simpanan per jenis. (Fungsi, bukan konstanta: urutan
 *  pemuatan file Apps Script tidak dijamin, AKUN bisa belum ada.) */
function akunSimpanan(jenis) {
  return jenis === 'pokok' ? AKUN.SIMPANAN_POKOK
       : jenis === 'wajib' ? AKUN.SIMPANAN_WAJIB : AKUN.SIMPANAN_SUKARELA;
}

/**
 * Rekap hak & kewajiban anggota yang akan keluar.
 * @return {Object} { anggota, saldo, pinjaman[], totalSimpanan, totalUtang, netto, masalah[] }
 */
function hitungRekapKeluar(idAnggota) {
  const anggota = getAnggota(idAnggota);
  if (!anggota) throw new Error('Anggota ' + idAnggota + ' tidak ditemukan.');

  const saldo = {
    pokok: getSaldoSimpanan(idAnggota, 'pokok'),
    wajib: getSaldoSimpanan(idAnggota, 'wajib'),
    sukarela: getSaldoSimpanan(idAnggota, 'sukarela')
  };
  const totalSimpanan = saldo.pokok + saldo.wajib + saldo.sukarela;

  const pinjaman = getPinjamanAktif(idAnggota).map(function(p) {
    return { id: p.id_pinjaman, keterangan: String(p.keterangan || ''),
      sisa: getSisaPinjaman(p.id_pinjaman) };
  }).filter(function(p) { return p.sisa > 0; });
  const totalUtang = pinjaman.reduce(function(s, p) { return s + p.sisa; }, 0);

  const masalah = [];
  if (anggota.status === 'keluar') masalah.push('Anggota ini sudah berstatus keluar.');

  const pending = getRowsByFilter(SHEET.PENGAMBILAN, function(r) {
    return String(r.id_anggota) === String(idAnggota) &&
           r.status_approval === 'PENDING' && r.status_lock !== STATUS_LOCK.VOID;
  });
  if (pending.length > 0) {
    masalah.push('Masih ada ' + pending.length + ' pengajuan pengambilan PENDING (' +
      pending.map(function(r) { return r.id_pengambilan; }).join(', ') +
      '). Approve atau tolak dulu di halaman Approval.');
  }

  const draftJasa = getRowsByFilter(SHEET.JASA_SUKARELA, function(r) {
    return String(r.id_anggota) === String(idAnggota) && r.status_posting === 'DRAFT';
  });
  if (draftJasa.length > 0) {
    masalah.push('Ada jasa sukarela berstatus DRAFT untuk anggota ini. ' +
      'Posting jasa bulan ini dulu supaya haknya ikut dihitung.');
  }

  ['pokok', 'wajib', 'sukarela'].forEach(function(jenis) {
    if (saldo[jenis] < 0) {
      masalah.push('Saldo ' + jenis + ' minus (' + formatRupiah(saldo[jenis]) +
        '). Periksa data transaksi anggota ini dulu.');
    }
  });

  return {
    anggota: anggota, saldo: saldo, pinjaman: pinjaman,
    totalSimpanan: totalSimpanan, totalUtang: totalUtang,
    netto: totalSimpanan - totalUtang, masalah: masalah
  };
}

/** Proses keluar lengkap. @return {Object} { pesan, netto } */
function prosesKeluarAnggota(idAnggota, alasan) {
  const tahun = getTahunAktif();
  validateTahunAktif(tahun);

  alasan = String(alasan || '').trim();
  if (!alasan) throw new Error('Alasan keluar wajib diisi (contoh: "Pindah tugas").');

  const rekap = hitungRekapKeluar(idAnggota);
  if (rekap.masalah.length > 0) {
    throw new Error('Proses keluar dibatalkan:\n• ' + rekap.masalah.join('\n• '));
  }

  // Uang yang benar-benar keluar dari kas = netto (pelunasan masuk dulu)
  const saldoKas = getSaldoKas();
  if (rekap.netto > saldoKas) {
    throw new Error('Saldo kas (' + formatRupiah(saldoKas) + ') tidak cukup untuk ' +
      'mengembalikan ' + formatRupiah(rekap.netto) + '.');
  }

  const nama = rekap.anggota.nama;

  // ---------- 1. LUNASI PINJAMAN DARI SIMPANAN ----------
  rekap.pinjaman.forEach(function(p) {
    inputAngsuranManual({
      id_pinjaman: p.id, nominal: p.sisa,
      keterangan: 'Pelunasan dari simpanan — anggota keluar'
    });
  });

  // ---------- 2. KEMBALIKAN SELURUH SIMPANAN ----------
  const tanggal = new Date();
  const bulan = tanggal.getMonth() + 1;
  const idPengembalian = [];

  ['pokok', 'wajib', 'sukarela'].forEach(function(jenis) {
    const jumlah = rekap.saldo[jenis];
    if (!(jumlah > 0)) return;

    const idAmbil = generateId(SHEET.PENGAMBILAN, ID_PREFIX.PENGAMBILAN, 1, tahun, null);
    const row = {
      id_pengambilan: idAmbil,
      tanggal: tanggal, tahun: tahun, bulan: bulan,
      id_anggota: idAnggota,
      jenis_simpanan: jenis,
      jumlah: jumlah,
      status_approval: 'APPROVED',
      keterangan: 'KELUAR ANGGOTA — ' + alasan,
      user_input: getUserEmail(),
      timestamp: new Date(),
      status_lock: STATUS_LOCK.OPEN
    };
    appendRowFromObject(SHEET.PENGAMBILAN, row);

    catatKasKeluar({
      tanggal: tanggal, tahun: tahun, bulan: bulan,
      kategori: 'pengambilan', referensi: idAmbil,
      keterangan: 'Pengembalian simpanan ' + jenis + ' (keluar) — ' + nama,
      nominal: jumlah
    });

    catatJurnal(tanggal, tahun, bulan, idAmbil,
      'Pengembalian simpanan ' + jenis + ' — ' + nama + ' keluar',
      [
        { kode_akun: akunSimpanan(jenis), debit: jumlah, kredit: 0 },
        { kode_akun: AKUN.KAS, debit: 0, kredit: jumlah }
      ]
    );

    logAktivitas('INSERT', SHEET.PENGAMBILAN, idAmbil, null, row);
    idPengembalian.push(idAmbil);
  });

  // ---------- 3. STATUS ANGGOTA ----------
  updateRowByField(SHEET.ANGGOTA, 'id_anggota', idAnggota, { status: 'keluar' });
  logAktivitas('EDIT', SHEET.ANGGOTA, idAnggota, { status: rekap.anggota.status }, {
    status: 'keluar', alasan: alasan,
    total_simpanan: rekap.totalSimpanan, total_utang: rekap.totalUtang,
    netto: rekap.netto, pengembalian: idPengembalian,
    pinjaman_dilunasi: rekap.pinjaman.map(function(p) { return p.id; })
  });

  const baris = [
    '✅ ' + nama + ' resmi keluar.',
    'Total simpanan dikembalikan: ' + formatRupiah(rekap.totalSimpanan),
    'Sisa pinjaman dilunasi: ' + formatRupiah(rekap.totalUtang)
  ];
  baris.push(rekap.netto >= 0
    ? '➡️ Serahkan uang bersih ' + formatRupiah(rekap.netto) + ' kepada anggota.'
    : '➡️ Anggota wajib menyetor kekurangan ' + formatRupiah(-rekap.netto) + '.');

  return { pesan: baris.join('\n'), netto: rekap.netto };
}

// ============================================================
// API UNTUK APLIKASI WEB
// ============================================================

function apiKeluarRekap(idAnggota) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  const r = hitungRekapKeluar(idAnggota);
  return {
    id: r.anggota.id_anggota, nama: r.anggota.nama, status: r.anggota.status,
    saldo: r.saldo, pinjaman: r.pinjaman,
    totalSimpanan: r.totalSimpanan, totalUtang: r.totalUtang,
    netto: r.netto, masalah: r.masalah, saldoKas: getSaldoKas()
  };
}

function apiProsesKeluar(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  return prosesKeluarAnggota(d.id, d.alasan).pesan;
}

// ===== koreksi_data.js =====
/**
 * ============================================================
 * KOREKSI_DATA.GS — Koreksi Jurnal Jasa Pinjaman (data lama)
 * Koperasi App
 * ============================================================
 *
 * MASALAH YANG DIKOREKSI:
 * Versi lama pinjaman.js mencatat pencairan sebagai
 *   Debit Piutang (total_tagihan) / Kredit Kas (nominal)
 *                                 / Kredit Pendapatan Jasa (total_jasa)
 * padahal jasa juga diakui lagi setiap angsuran dibayar.
 * Akibatnya pendapatan jasa pinjaman DOBEL dan piutang tersisa
 * sebesar total_jasa meskipun pinjaman sudah lunas.
 *
 * KOREKSI: untuk setiap pinjaman yang terkena, dibuat jurnal
 *   Debit 401 Pendapatan Jasa Pinjaman / Kredit 103 Piutang
 * sebesar total_jasa, dengan referensi = id_pinjaman (supaya kalau
 * pinjaman itu kelak di-VOID, jurnal koreksinya ikut dibalik).
 *
 * CARA PAKAI (dari editor Apps Script):
 * 1. Jalankan koreksiJasaPinjaman_cek      → lihat daftar di Logs,
 *    TIDAK menulis apa pun.
 * 2. Jika daftarnya benar, jalankan koreksiJasaPinjaman_jalankan.
 * Aman dijalankan berulang: pinjaman yang sudah dikoreksi dilewati.
 *
 * Pinjaman dari tahun yang SUDAH DITUTUP tidak dikoreksi otomatis
 * (SHU-nya sudah dibagi) — hanya dilaporkan, putuskan manual.
 * ============================================================
 */

const PENANDA_KOREKSI_JASA = 'KOREKSI jasa dobel';

function koreksiJasaPinjaman_cek() {
  return koreksiJasaPinjaman(true);
}

function koreksiJasaPinjaman_jalankan() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return koreksiJasaPinjaman(false);
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

/**
 * @param {boolean} simulasi  true = hanya laporan, tanpa menulis
 * @return {Object} { dikoreksi: [], dilewati: [], tahunTertutup: [] }
 */
function koreksiJasaPinjaman(simulasi) {
  MODE_SENYAP = true;
  const hasil = { dikoreksi: [], dilewati: [], tahunTertutup: [] };

  const pinjamanList = getRowsByFilter(SHEET.PINJAMAN, function(p) {
    return p.status_lock !== STATUS_LOCK.VOID && Number(p.total_jasa) > 0;
  });

  pinjamanList.forEach(function(p) {
    const id = String(p.id_pinjaman);
    const jurnalPinjaman = getRowsByFilter(SHEET.JURNAL, function(j) {
      return String(j.referensi) === id;
    });

    // Terkena bug hanya jika jurnal pencairannya punya Kredit 401
    const kreditJasa = jurnalPinjaman.reduce(function(s, j) {
      return String(j.kode_akun) === AKUN.PENDAPATAN_JASA_PINJAMAN
        ? s + (Number(j.kredit) || 0) : s;
    }, 0);
    const sudahDikoreksi = jurnalPinjaman.some(function(j) {
      return String(j.keterangan).indexOf(PENANDA_KOREKSI_JASA) === 0;
    });

    if (kreditJasa <= 0 || sudahDikoreksi) {
      hasil.dilewati.push(id + (sudahDikoreksi ? ' (sudah dikoreksi)' : ' (format baru)'));
      return;
    }

    const tahun = Number(p.tahun);
    const setting = getSettingRAT(tahun);
    if (!setting || setting.status_aktif !== true) {
      hasil.tahunTertutup.push(id + ' — tahun ' + tahun + ', jasa ' + formatRupiah(kreditJasa));
      return;
    }

    if (!simulasi) {
      const kini = new Date();
      catatJurnal(kini, tahun, kini.getMonth() + 1, id,
        PENANDA_KOREKSI_JASA + ' — pencairan ' + id,
        [
          { kode_akun: AKUN.PENDAPATAN_JASA_PINJAMAN, debit: kreditJasa, kredit: 0 },
          { kode_akun: AKUN.PIUTANG_PINJAMAN, debit: 0, kredit: kreditJasa }
        ]
      );
      logAktivitas('EDIT', SHEET.JURNAL, id, null,
        { koreksi: 'jasa pinjaman dobel', nominal: kreditJasa });
    }
    hasil.dikoreksi.push(id + ' — ' + formatRupiah(kreditJasa));
  });

  Logger.log((simulasi ? '🔍 SIMULASI (tidak ada yang ditulis)' : '✅ KOREKSI SELESAI') +
    '\n\nDikoreksi: ' + hasil.dikoreksi.length + '\n' + hasil.dikoreksi.join('\n') +
    '\n\nTahun sudah ditutup (tidak dikoreksi): ' + hasil.tahunTertutup.length + '\n' +
    hasil.tahunTertutup.join('\n') +
    '\n\nDilewati: ' + hasil.dilewati.length);
  return hasil;
}

// ============================================================
// KONVERSI PINJAMAN AKTIF: JASA FLAT → JASA MENURUN
// ============================================================
//
// Keputusan koperasi: jasa pinjaman = persen × SISA POKOK tiap bulan
// (seperti Excel pembukuan). Pinjaman aktif yang dibuat sebelum itu
// masih memakai hitungan flat. Konversi ini:
//   - sisa pokok  = nominal − SUM(angsuran_pokok yang sudah dibayar)
//   - angsuran pokok/bulan = porsi pokok dari angsuran lama
//     (untuk pinjaman MIGRASI: angsuran lama memang sudah pokok)
//   - jasa_persen: pinjaman biasa tetap persen lamanya;
//     pinjaman MIGRASI (dulu 0%) memakai jasa_pinjaman tahun aktif
//   - sisa jasa flat yang belum ditagih DIHAPUS; mulai bulan depan
//     jasa dihitung dari sisa pokok.
// Jurnal tidak berubah (piutang sudah berbasis pokok sejak koreksi).
//
// CARA PAKAI (dari editor Apps Script):
// 1. Pastikan jasa_pinjaman tahun aktif di Pengaturan sudah benar
//    (Excel: 1,5 %/bulan).
// 2. Jalankan koreksiJasaPinjaman_cek / _jalankan dulu (jika ada).
// 3. Jalankan konversiPinjamanMenurun_cek → baca Logs.
// 4. Jika benar, jalankan konversiPinjamanMenurun_jalankan.

function konversiPinjamanMenurun_cek() {
  return konversiPinjamanMenurun(true);
}

function konversiPinjamanMenurun_jalankan() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return konversiPinjamanMenurun(false);
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function konversiPinjamanMenurun(simulasi) {
  MODE_SENYAP = true;
  if (!simulasi) pastikanKolomPinjaman();
  const hasil = { dikonversi: [], ditahan: [] };

  const tahunAktif = getTahunAktif();
  const setting = tahunAktif ? getSettingRAT(tahunAktif) : null;
  const jasaSetting = setting ? Number(setting.jasa_pinjaman) : 0;

  const daftar = getRowsByFilter(SHEET.PINJAMAN, function(p) {
    return p.status === 'aktif' && p.status_lock !== STATUS_LOCK.VOID && !isJasaMenurun(p);
  });

  daftar.forEach(function(p) {
    const id = String(p.id_pinjaman);

    // Pinjaman yang terkena jasa dobel tapi belum dikoreksi → tahan dulu
    const jurnal = getRowsByFilter(SHEET.JURNAL, function(j) { return String(j.referensi) === id; });
    const kreditJasa = jurnal.some(function(j) {
      return String(j.kode_akun) === AKUN.PENDAPATAN_JASA_PINJAMAN && Number(j.kredit) > 0;
    });
    const sudahDikoreksi = jurnal.some(function(j) {
      return String(j.keterangan).indexOf(PENANDA_KOREKSI_JASA) === 0;
    });
    if (kreditJasa && !sudahDikoreksi) {
      hasil.ditahan.push(id + ' — jalankan koreksiJasaPinjaman dulu');
      return;
    }

    const angsuran = getRowsByFilter(SHEET.ANGSURAN, function(a) {
      return String(a.id_pinjaman) === id && a.status_lock !== STATUS_LOCK.VOID;
    });
    const pokokDibayar = angsuran.reduce(function(s, a) { return s + (Number(a.angsuran_pokok) || 0); }, 0);
    const jasaDibayar = angsuran.reduce(function(s, a) { return s + (Number(a.jasa) || 0); }, 0);
    const nominal = Number(p.nominal) || 0;
    const sisaPokok = nominal - pokokDibayar;
    const sisaLama = getSisaPinjaman(id);   // hitungan flat (pokok + jasa)

    const migrasi = String(p.keterangan || '').indexOf('MIGRASI') === 0;
    const rasio = Number(p.total_tagihan) > 0 ? nominal / Number(p.total_tagihan) : 1;
    const angsuranPokok = Math.max(1, Math.round(Number(p.angsuran_perbulan) * rasio));
    const jasaPersen = migrasi || !(Number(p.jasa_persen) > 0) ? jasaSetting : Number(p.jasa_persen);

    const perubahan = {
      metode_jasa: 'menurun',
      jasa_persen: jasaPersen,
      total_tagihan: nominal,
      total_jasa: jasaDibayar,
      angsuran_perbulan: angsuranPokok
    };
    if (sisaPokok <= 0) perubahan.status = 'lunas';

    if (!simulasi) {
      updateRowByRowNumber(SHEET.PINJAMAN, p.__row, perubahan);
      logAktivitas('EDIT', SHEET.PINJAMAN, id, {
        metode_jasa: 'flat', jasa_persen: p.jasa_persen, total_tagihan: p.total_tagihan,
        total_jasa: p.total_jasa, angsuran_perbulan: p.angsuran_perbulan
      }, Object.assign({ konversi: 'jasa menurun' }, perubahan));
    }
    hasil.dikonversi.push(id + (migrasi ? ' (migrasi)' : '') +
      ' — sisa lama ' + formatRupiah(sisaLama) + ' → sisa pokok ' + formatRupiah(sisaPokok) +
      ', pokok ' + formatRupiah(angsuranPokok) + '/bln, jasa ' + jasaPersen + '%/bln' +
      ' (±' + formatRupiah(Math.round(Math.max(0, sisaPokok) * jasaPersen / 100)) + ' bln depan)');
  });

  Logger.log((simulasi ? '🔍 SIMULASI konversi (tidak ada yang ditulis)' : '✅ KONVERSI SELESAI') +
    '\n\nJasa pinjaman tahun aktif di Pengaturan: ' + jasaSetting + '%/bulan' +
    '\n\nDikonversi: ' + hasil.dikonversi.length + '\n' + hasil.dikonversi.join('\n') +
    '\n\nDitahan: ' + hasil.ditahan.length + '\n' + hasil.ditahan.join('\n'));
  return hasil;
}

// ===== laporan.js =====
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

// ===== laporan_app.js =====
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

// ===== log_email_backup.js =====
/**
 * ============================================================
 * LOG_EMAIL_BACKUP.GS — Log Aktivitas, Pengingat Email, Backup
 * Koperasi App
 * ============================================================
 *
 * TIGA FITUR DALAM SATU FILE:
 *
 * 1. apiLogAktivitas — penampil log untuk halaman "Log" di
 *    aplikasi (khusus admin). Data dari sheet log_aktivitas.
 *
 * 2. PENGINGAT EMAIL OTOMATIS — satu trigger harian mengecek
 *    tanggal:
 *      • Tgl 10: jika angsuran bulan ini BELUM digenerate →
 *        email pengingat ke semua admin
 *      • Tgl 25: jika jasa sukarela bulan ini BELUM diposting →
 *        email pengingat ke semua admin
 *
 * 3. BACKUP OTOMATIS — setiap tanggal 1, salinan spreadsheet
 *    disimpan ke folder "Backup Koperasi" di Drive pemilik.
 *
 * ⚙️ AKTIVASI (SEKALI SAJA, dari editor Apps Script):
 *    Pilih fungsi `pasangTugasHarian` di dropdown → Run.
 *    Trigger harian jam 06.00–07.00 akan terpasang otomatis.
 *    Untuk mematikan: jalankan `hapusTugasHarian`.
 * ============================================================
 */

// ============================================================
// 1. PENAMPIL LOG AKTIVITAS (khusus admin)
// ============================================================

/**
 * @param {Object} f { aksi: 'SEMUA'|'INSERT'|..., limit: 50|100|200 }
 */
function apiLogAktivitas(f) {
  requireRole(['admin']);
  MODE_SENYAP = true;

  const aksi = (f && f.aksi) || 'SEMUA';
  const limit = Number((f && f.limit) || 100);

  let rows = sheetToObjects(SHEET.LOG);
  if (aksi !== 'SEMUA') {
    rows = rows.filter(function(r) { return r.aksi === aksi; });
  }

  // Terbaru di atas, potong sesuai limit
  rows = rows.slice(-limit).reverse();

  return rows.map(function(r) {
    let ringkas = '';
    try {
      const d = r.data_baru ? JSON.parse(r.data_baru) : null;
      if (d) {
        ringkas = Object.keys(d).slice(0, 3).map(function(k) {
          return k + ': ' + String(d[k]).substring(0, 30);
        }).join(' · ');
      }
    } catch (e) { ringkas = String(r.data_baru || '').substring(0, 60); }

    return {
      waktu: r.timestamp
        ? Utilities.formatDate(new Date(r.timestamp),
            Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
        : '',
      user: String(r.user || '').split('@')[0],
      aksi: r.aksi,
      sheet: r.sheet_target,
      ref: r.id_referensi,
      ringkas: ringkas
    };
  });
}

// ============================================================
// 2 & 3. TUGAS HARIAN: PENGINGAT + BACKUP
// ============================================================

/** Folder Drive tempat cadangan (bisa diatur di identitas_lokal.js → folder_backup). */
function namaFolderBackup_() {
  const lokal = typeof identitasLokal_ === 'function' ? identitasLokal_() : {};
  return lokal.folder_backup || 'Backup Koperasi ' + getProfilKoperasi().nama_koperasi.replace(/"/g, '');
}

/** Dipanggil trigger setiap pagi. Mengecek tanggal & bertindak. */
function tugasHarian(e) {
  // Trigger waktu mengirim event berisi triggerUid milik proyek ini.
  // Panggilan lain (mis. dari browser) wajib lolos penjaga akses.
  const dariTrigger = !!(e && e.triggerUid) && ScriptApp.getProjectTriggers()
    .some(function(t) { return String(t.getUniqueId()) === String(e.triggerUid); });
  if (dariTrigger) __AKSES_PUBLIK = true; else pastikanAksesData();

  const kini = new Date();
  const tanggal = kini.getDate();

  try {
    bersihkanSesiKedaluwarsa();   // buang token login yang sudah lewat masa
    if (tanggal === 1) backupSpreadsheet();
    if (tanggal === 10) pengingatAngsuran();
    if (tanggal === 25) pengingatJasa();
  } catch (e) {
    // Kirim error ke pemilik supaya tidak diam-diam gagal
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
      '⚠️ Tugas harian koperasi GAGAL',
      'Terjadi error pada tugas harian tanggal ' + kini + ':\n\n' + e.message);
  }
}

/** Email semua admin aktif. */
function kirimKeAdmin(judul, isi) {
  pastikanAksesData();
  pastikanSheetPengguna();
  const adminList = sheetToObjects(SHEET_PENGGUNA).filter(function(u) {
    return u.peran === 'admin' && u.status === 'aktif';
  });
  if (adminList.length === 0) return;

  adminList.forEach(function(u) {
    MailApp.sendEmail(String(u.email), judul,
      'Assalamu\u2019alaikum / Selamat pagi,\n\n' + isi +
      '\n\n—\nPesan otomatis dari aplikasi Koperasi ' + getProfilKoperasi().nama_koperasi.replace(/"/g, '') + '.');
  });
}

/** Tgl 10: ingatkan generate angsuran JIKA belum dilakukan bulan ini. */
function pengingatAngsuran() {
  const kini = new Date();
  const tahun = kini.getFullYear(), bulan = kini.getMonth() + 1;

  const sudah = getRowsByFilter(SHEET.ANGSURAN, function(r) {
    return Number(r.tahun) === tahun && Number(r.bulan) === bulan &&
           r.status_lock !== STATUS_LOCK.VOID;
  });
  if (sudah.length > 0) return; // sudah digenerate → tidak perlu mengingatkan

  const pinjamanAktif = getRowsByFilter(SHEET.PINJAMAN, function(p) {
    return p.status === 'aktif' && p.status_lock !== STATUS_LOCK.VOID;
  });
  if (pinjamanAktif.length === 0) return; // tidak ada yang perlu ditagih

  kirimKeAdmin('🔔 Pengingat: Generate Angsuran bulan ' + bulan + '/' + tahun,
    'Hari ini tanggal 10 — waktunya proses angsuran bulanan.\n\n' +
    'Angsuran bulan ' + bulan + '/' + tahun + ' BELUM digenerate, padahal ada ' +
    pinjamanAktif.length + ' pinjaman aktif.\n\n' +
    'Buka aplikasi → menu Proses Bulanan → "Generate Angsuran Bulan Ini", ' +
    'lalu cetak struk potongan untuk bendahara gaji.');
}

/** Tgl 25: ingatkan hitung+posting jasa JIKA belum diposting bulan ini. */
function pengingatJasa() {
  const kini = new Date();
  const tahun = kini.getFullYear(), bulan = kini.getMonth() + 1;

  const sudahPosted = getRowsByFilter(SHEET.JASA_SUKARELA, function(r) {
    return Number(r.tahun) === tahun && Number(r.bulan) === bulan &&
           r.status_posting === 'POSTED';
  });
  if (sudahPosted.length > 0) return;

  const masihDraft = getRowsByFilter(SHEET.JASA_SUKARELA, function(r) {
    return Number(r.tahun) === tahun && Number(r.bulan) === bulan &&
           r.status_posting === 'DRAFT';
  });

  kirimKeAdmin('🔔 Pengingat: Jasa Sukarela bulan ' + bulan + '/' + tahun,
    'Hari ini tanggal 25 — waktunya proses jasa sukarela akhir bulan.\n\n' +
    (masihDraft.length > 0
      ? 'Ada ' + masihDraft.length + ' baris jasa berstatus DRAFT yang belum ' +
        'diposting. Periksa angkanya lalu klik "Posting Jasa (Kunci)".'
      : 'Jasa sukarela bulan ini belum dihitung. Buka aplikasi → Proses Bulanan → ' +
        '"Hitung Jasa Sukarela", periksa hasilnya, lalu posting.'));
}

/** Tgl 1: salin spreadsheet ke folder backup di Drive. */
function backupSpreadsheet() {
  pastikanAksesData();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fileAsli = DriveApp.getFileById(ss.getId());

  // Cari / buat folder backup
  let folder;
  const cari = DriveApp.getFoldersByName(namaFolderBackup_());
  folder = cari.hasNext() ? cari.next() : DriveApp.createFolder(namaFolderBackup_());

  const stempel = Utilities.formatDate(new Date(),
    Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const namaBackup = 'BACKUP ' + stempel + ' — ' + ss.getName();

  fileAsli.makeCopy(namaBackup, folder);

  // Rapikan: simpan maksimal 12 backup terakhir, hapus yang lebih tua
  const semua = [];
  const iter = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (iter.hasNext()) semua.push(iter.next());
  semua.sort(function(a, b) { return b.getDateCreated() - a.getDateCreated(); });
  for (let i = 12; i < semua.length; i++) semua[i].setTrashed(true);

  kirimKeAdmin('💾 Backup bulanan berhasil — ' + stempel,
    'Salinan spreadsheet koperasi berhasil dibuat:\n"' + namaBackup + '"\n\n' +
    'Lokasi: folder Drive "' + namaFolderBackup_() + '" milik pemilik sistem.\n' +
    'Sistem menyimpan 12 backup terakhir; yang lebih lama dihapus otomatis.');
}

// ============================================================
// PASANG / HAPUS TRIGGER (jalankan dari editor, sekali saja)
// ============================================================

function pasangTugasHarian() {
  pastikanAksesData();
  // Hindari trigger dobel
  hapusTugasHarian();

  ScriptApp.newTrigger('tugasHarian')
    .timeBased()
    .everyDays(1)
    .atHour(6)     // antara jam 06.00–07.00
    .create();

  Logger.log('✅ Trigger harian terpasang (jam 06.00–07.00).');
  Logger.log('Tgl 1: backup | Tgl 10: pengingat angsuran | Tgl 25: pengingat jasa.');
}

function hapusTugasHarian() {
  pastikanAksesData();
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'tugasHarian') ScriptApp.deleteTrigger(t);
  });
  Logger.log('Trigger tugasHarian (jika ada) sudah dihapus.');
}

/** Tes manual dari editor: paksa kirim pengingat & backup sekarang. */
function test_tugasHarian_paksa() {
  Logger.log('Menguji backup…'); backupSpreadsheet();
  Logger.log('Menguji pengingat angsuran…'); pengingatAngsuran();
  Logger.log('Menguji pengingat jasa…'); pengingatJasa();
  Logger.log('✅ Selesai. Cek email admin & folder Drive "' + namaFolderBackup_() + '".');
}
