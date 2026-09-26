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