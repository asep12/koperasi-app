/**
 * ============================================================
 * PINJAMAN.GS — Input Pinjaman Baru & Pinjaman Migrasi
 * Koperasi App
 * Sesuai dokumen: REKAP_Koperasi_v2.md (BAGIAN 3, 4, 12, 17, 18)
 * ============================================================
 *
 * DUA JALUR INPUT — JANGAN TERTUKAR:
 *
 * 1. inputPinjamanBaru()
 *    Untuk pinjaman yang BENAR-BENAR dicairkan sekarang.
 *    → Kas KELUAR + jurnal dibuat (uang sungguhan keluar hari ini).
 *
 * 2. inputPinjamanMigrasi()
 *    Untuk pinjaman LAMA dari Excel yang masih berjalan saat
 *    sistem ini mulai dipakai. Uangnya sudah keluar di masa lalu,
 *    SEBELUM sistem ini ada.
 *    → TIDAK ADA kas keluar baru, TIDAK ADA jurnal baru — hanya
 *      catatan historis supaya "Generate Angsuran" bisa jalan.
 *    → keterangan WAJIB diawali "MIGRASI" (dicek otomatis).
 * ============================================================
 */

// ============================================================
// 1. PINJAMAN BARU (pencairan sungguhan)
// ============================================================

/**
 * Input pinjaman yang dicairkan hari ini (jasa MENURUN).
 *
 * - Anggota BELUM punya pinjaman aktif → pinjaman baru (induk).
 * - Anggota SUDAH punya pinjaman aktif → pinjaman TAMBAHAN yang
 *   digabung ke sisa lama (sesuai kebiasaan di Excel): baris baru
 *   dengan id_induk, angsuran pokok induk dihitung ulang dari
 *   (sisa lama + tambahan) / tenor.
 *
 * Jasa tidak dihitung di muka: tiap bulan jasa = jasa_persen ×
 * sisa pokok awal bulan (lihat generateSatuAngsuran).
 *
 * @param {Object} data
 * @param {string} data.id_anggota
 * @param {number} data.nominal          pokok yang dicairkan sekarang
 * @param {number} data.tenor            lama cicilan (bulan) untuk total sisa
 * @param {number} [data.jasa_persen]    default: setting tahun aktif
 * @param {number} [data.angsuran_pokok] default: total / tenor dibulatkan ke atas
 * @param {string} [data.keterangan]
 * @param {Date}   [data.tanggal]        default: hari ini
 * @return {Object} { id_pinjaman, id_induk, pesan }
 */
function inputPinjamanBaru(data) {
  const tanggal = data.tanggal ? new Date(data.tanggal) : new Date();
  const tahun = tanggal.getFullYear();
  const bulan = tanggal.getMonth() + 1;

  const setting = validateTahunAktif(tahun);
  const anggota = pastikanAnggotaAktif(data.id_anggota);
  validatePinjaman(data, tahun);
  pastikanKolomPinjaman();

  const nominal = Number(data.nominal);
  const tenor = Number(data.tenor);
  const induk = getPinjamanAktif(data.id_anggota)[0] || null;
  const jasaPersen = induk ? Number(induk.jasa_persen)
    : (data.jasa_persen === undefined || data.jasa_persen === '' || data.jasa_persen === null
        ? Number(setting.jasa_pinjaman) : Number(data.jasa_persen));
  const sisaLama = induk ? getSisaPinjaman(induk.id_pinjaman) : 0;
  const totalPokok = sisaLama + nominal;
  const angsuranPokok = Number(data.angsuran_pokok) > 0
    ? Math.round(Number(data.angsuran_pokok)) : hitungAngsuranPokok(totalPokok, tenor);
  if (angsuranPokok > totalPokok) {
    throw new Error('Angsuran pokok per bulan tidak boleh melebihi total pinjaman.');
  }

  const idPinjaman = generateId(SHEET.PINJAMAN, ID_PREFIX.PINJAMAN, 1, tahun, null);
  const rowPinjaman = {
    id_pinjaman: idPinjaman,
    tanggal: tanggal, tahun: tahun, bulan: bulan,
    id_anggota: data.id_anggota,
    nominal: nominal,
    tenor: tenor,
    jasa_persen: jasaPersen,
    total_jasa: 0,                 // jasa menurun: dihitung tiap bulan
    total_tagihan: nominal,        // pokok saja
    angsuran_perbulan: angsuranPokok,
    status: induk ? 'tambahan' : 'aktif',
    keterangan: data.keterangan || '',
    user_input: getUserEmail(),
    timestamp: new Date(),
    status_lock: STATUS_LOCK.OPEN,
    metode_jasa: 'menurun',
    id_induk: induk ? induk.id_pinjaman : '',
    angsuran_sebelumnya: induk ? Number(induk.angsuran_perbulan) : ''
  };
  appendRowFromObject(SHEET.PINJAMAN, rowPinjaman);

  if (induk) {
    updateRowByRowNumber(SHEET.PINJAMAN, induk.__row,
      { angsuran_perbulan: angsuranPokok, tenor: tenor });
    logAktivitas('EDIT', SHEET.PINJAMAN, induk.id_pinjaman,
      { angsuran_perbulan: induk.angsuran_perbulan, tenor: induk.tenor },
      { angsuran_perbulan: angsuranPokok, tenor: tenor, tambahan: idPinjaman });
  }

  // ---------- KAS KELUAR (pencairan) ----------
  catatKasKeluar({
    tanggal: tanggal, tahun: tahun, bulan: bulan,
    kategori: 'pinjaman',
    referensi: idPinjaman,
    keterangan: (induk ? 'Tambahan pinjaman — ' : 'Pencairan pinjaman — ') + anggota.nama,
    nominal: nominal
  });

  // ---------- JURNAL (BAGIAN 12 #4) ----------
  // Debit Piutang = Kredit Kas, sebesar POKOK. Jasa diakui saat angsuran dibayar.
  catatJurnal(tanggal, tahun, bulan, idPinjaman,
    (induk ? 'Tambahan pinjaman — ' : 'Pencairan pinjaman — ') + anggota.nama,
    [
      { kode_akun: AKUN.PIUTANG_PINJAMAN, debit: nominal, kredit: 0 },
      { kode_akun: AKUN.KAS, debit: 0, kredit: nominal }
    ]
  );

  logAktivitas('INSERT', SHEET.PINJAMAN, idPinjaman, null, rowPinjaman);

  const jasaPertama = Math.round(totalPokok * jasaPersen / 100);
  return {
    id_pinjaman: idPinjaman,
    id_induk: induk ? induk.id_pinjaman : idPinjaman,
    pesan: '✅ ' + (induk ? 'Tambahan pinjaman ' : 'Pinjaman ') + formatRupiah(nominal) +
      ' untuk ' + anggota.nama + ' berhasil dicairkan.\n' +
      (induk ? 'Digabung ke ' + induk.id_pinjaman + ' (sisa lama ' + formatRupiah(sisaLama) +
        ') → total sisa pokok ' + formatRupiah(totalPokok) + '\n' : '') +
      'Angsuran pokok: ' + formatRupiah(angsuranPokok) + '/bulan\n' +
      'Jasa ' + jasaPersen + '%/bulan dari sisa pokok (bulan depan ±' +
      formatRupiah(jasaPertama) + ', lalu menurun)'
  };
}

// ============================================================
// 2. PINJAMAN MIGRASI (data historis dari Excel lama)
// ============================================================

/**
 * Input pinjaman lama yang masih berjalan, hasil migrasi dari
 * Excel. TIDAK membuat kas keluar / jurnal baru — murni catatan
 * historis. Sisa piutangnya harus SAMA dengan baris "piutang"
 * yang sudah Anda input di sheet saldo_awal untuk anggota ini.
 *
 * @param {Object} data
 * @param {string} data.id_anggota
 * @param {number} data.sisa_pokok      SISA pokok yang belum lunas
 * @param {number} data.tenor_sisa      SISA bulan angsuran
 * @param {Date}   data.tanggal_asli    tanggal pencairan ASLI (dulu)
 * @param {string} [data.keterangan_tambahan]
 * @return {Object} { id_pinjaman, pesan }
 */
function inputPinjamanMigrasi(data) {
  const tanggalAsli = new Date(data.tanggal_asli);
  const tahunPencairanAsli = tanggalAsli.getFullYear();
  const bulanPencairanAsli = tanggalAsli.getMonth() + 1;

  const anggota = pastikanAnggotaAktif(data.id_anggota);
  pastikanKolomPinjaman();

  if (!data.sisa_pokok || Number(data.sisa_pokok) <= 0) {
    throw new Error('Sisa pokok harus lebih dari 0.');
  }
  if (!data.tenor_sisa || Number(data.tenor_sisa) <= 0) {
    throw new Error('Sisa tenor harus lebih dari 0 bulan.');
  }
  const aktif = getPinjamanAktif(data.id_anggota);
  if (aktif.length > 0) {
    throw new Error(anggota.nama + ' sudah punya pinjaman aktif (' + aktif[0].id_pinjaman +
      '). Satu anggota hanya punya satu pinjaman aktif — gunakan pinjaman tambahan.');
  }

  // Piutang di Excel = sisa POKOK; jasa ditagih tiap bulan dari sisa itu
  // (jasa menurun), sama seperti pinjaman baru.
  const sisaPokok = Math.round(Number(data.sisa_pokok));
  const tahunAktif = getTahunAktif();
  const setting = tahunAktif ? getSettingRAT(tahunAktif) : null;
  const jasaPersen = data.jasa_persen === undefined || data.jasa_persen === '' || data.jasa_persen === null
    ? (setting ? Number(setting.jasa_pinjaman) : 0) : Number(data.jasa_persen);
  const angsuranPokok = Number(data.angsuran_pokok) > 0
    ? Math.round(Number(data.angsuran_pokok)) : hitungAngsuranPokok(sisaPokok, data.tenor_sisa);

  const idPinjaman = generateId(SHEET.PINJAMAN, ID_PREFIX.PINJAMAN, 1, tahunPencairanAsli, null);

  const keteranganFinal = 'MIGRASI ' + tahunPencairanAsli +
    (data.keterangan_tambahan ? ' — ' + data.keterangan_tambahan : '');

  const rowPinjaman = {
    id_pinjaman: idPinjaman,
    tanggal: tanggalAsli,
    tahun: tahunPencairanAsli,
    bulan: bulanPencairanAsli,
    id_anggota: data.id_anggota,
    nominal: sisaPokok,            // untuk migrasi, nominal = sisa pokok per migrasi
    tenor: Number(data.tenor_sisa),
    jasa_persen: jasaPersen,
    total_jasa: 0,
    total_tagihan: sisaPokok,
    angsuran_perbulan: angsuranPokok,
    status: 'aktif',
    keterangan: keteranganFinal,
    user_input: getUserEmail(),
    timestamp: new Date(),
    status_lock: STATUS_LOCK.OPEN,
    metode_jasa: 'menurun',
    id_induk: '',
    angsuran_sebelumnya: ''
  };
  appendRowFromObject(SHEET.PINJAMAN, rowPinjaman);

  // TIDAK ADA kas keluar. TIDAK ADA jurnal baru.
  // (uangnya sudah keluar di masa lalu; piutangnya tercatat di saldo_awal)

  logAktivitas('INSERT', SHEET.PINJAMAN, idPinjaman, null, rowPinjaman);

  return {
    id_pinjaman: idPinjaman,
    pesan: '✅ Pinjaman migrasi untuk ' + anggota.nama + ' tercatat.\n' +
      'ID: ' + idPinjaman + '\n' +
      'Sisa pokok: ' + formatRupiah(sisaPokok) + '\n' +
      'Angsuran pokok: ' + formatRupiah(angsuranPokok) + '/bulan + jasa ' + jasaPersen +
      '% dari sisa\n\n' +
      '⚠️ PENTING: pastikan nominal ' + formatRupiah(sisaPokok) +
      ' ini SAMA dengan baris "piutang" untuk ' + anggota.nama + ' di sheet saldo_awal. ' +
      'Jalankan menu "🔍 Validasi Migrasi Pinjaman" untuk mengecek otomatis.'
  };
}

/**
 * Validasi: total sisa pinjaman migrasi (transaksi_pinjaman ber-
 * keterangan MIGRASI) harus sama dengan total piutang di saldo_awal,
 * per anggota. Dipanggil manual lewat menu untuk cek konsistensi.
 */
function validasiMigrasiPinjaman() {
  const piutangSaldoAwal = getRowsByFilter(SHEET.SALDO_AWAL, function(row) {
    return row.jenis === 'piutang';
  });

  const hasil = [];
  let semuaCocok = true;

  piutangSaldoAwal.forEach(function(row) {
    const anggota = getAnggota(row.id_anggota);
    const namaAnggota = anggota ? anggota.nama : row.id_anggota;

    const pinjamanMigrasi = getRowsByFilter(SHEET.PINJAMAN, function(p) {
      return p.id_anggota === row.id_anggota &&
             String(p.keterangan).indexOf('MIGRASI') === 0 &&
             p.status_lock !== STATUS_LOCK.VOID;
    });

    const totalSisaPinjaman = pinjamanMigrasi.reduce(function(sum, p) {
      return sum + getSisaPinjaman(p.id_pinjaman);
    }, 0);

    const cocok = Math.abs(totalSisaPinjaman - Number(row.nominal)) < 1;
    if (!cocok) semuaCocok = false;

    hasil.push(
      (cocok ? '✅' : '❌') + ' ' + namaAnggota + ' (' + row.id_anggota + '): ' +
      'saldo_awal piutang = ' + formatRupiah(row.nominal) +
      ', sisa transaksi_pinjaman migrasi = ' + formatRupiah(totalSisaPinjaman) +
      (pinjamanMigrasi.length === 0 ? ' ⚠️ BELUM ADA baris pinjaman migrasi!' : '')
    );
  });

  tampilkanPesan(
    (semuaCocok ? '✅ SEMUA COCOK\n\n' : '⚠️ ADA YANG TIDAK COCOK\n\n') +
    (hasil.length > 0 ? hasil.join('\n') : '(Tidak ada baris piutang di saldo_awal)')
  );
}

// ============================================================
// VALIDASI PINJAMAN BARU (BAGIAN 17)
// ============================================================

function validatePinjaman(data, tahun) {
  if (!data.nominal || Number(data.nominal) <= 0) {
    throw new Error('Nominal pinjaman harus lebih dari 0.');
  }
  if (!data.tenor || Number(data.tenor) <= 0) {
    throw new Error('Tenor harus lebih dari 0 bulan.');
  }

  // Tenor maksimal 30 bulan
  if (Number(data.tenor) > 30) {
    throw new Error('Tenor maksimal 30 bulan.');
  }

  if (data.jasa_persen !== undefined && data.jasa_persen !== '' && data.jasa_persen !== null) {
    const j = Number(data.jasa_persen);
    if (isNaN(j) || j < 0 || j > 10) throw new Error('Jasa pinjaman harus 0–10 %/bulan.');
  }

  // Pinjaman aktif yang sudah ada → pinjaman ini jadi TAMBAHAN (digabung),
  // asalkan pinjaman lama sudah memakai jasa menurun.
  const pinjamanAktif = getPinjamanAktif(data.id_anggota);
  if (pinjamanAktif.length > 1) {
    throw new Error('Anggota ini punya lebih dari satu pinjaman aktif (' +
      pinjamanAktif.map(function(p) { return p.id_pinjaman; }).join(', ') +
      '). Rapikan dulu (lunasi/void) sebelum menambah pinjaman.');
  }
  if (pinjamanAktif.length === 1 && !isJasaMenurun(pinjamanAktif[0])) {
    throw new Error('Pinjaman aktif ' + pinjamanAktif[0].id_pinjaman + ' masih memakai ' +
      'hitungan lama (jasa flat). Jalankan konversi ke jasa menurun dulu ' +
      '(koreksi_data: konversiPinjamanMenurun_cek / _jalankan).');
  }

  const kasTersedia = getSaldoKas();
  if (kasTersedia < Number(data.nominal)) {
    throw new Error('Saldo kas tidak mencukupi. Kas tersedia: ' +
      formatRupiah(kasTersedia) + ', pinjaman diminta: ' + formatRupiah(data.nominal));
  }
}

/** Ambil semua pinjaman berstatus aktif milik satu anggota. */
function getPinjamanAktif(idAnggota) {
  return getRowsByFilter(SHEET.PINJAMAN, function(row) {
    return row.id_anggota === idAnggota && row.status === 'aktif' &&
           row.status_lock !== STATUS_LOCK.VOID;
  });
}

/**
 * Cek & update status pinjaman jadi 'lunas' jika sisa sudah 0.
 * Dipanggil dari angsuran.gs setiap kali ada angsuran baru masuk.
 */
function perbaruiStatusPinjamanJikaLunas(idPinjaman) {
  const sisa = getSisaPinjaman(idPinjaman);
  if (sisa <= 0) {
    updateRowByField(SHEET.PINJAMAN, 'id_pinjaman', idPinjaman, { status: 'lunas' });
    logAktivitas('EDIT', SHEET.PINJAMAN, idPinjaman, { status: 'aktif' }, { status: 'lunas' });
    return true;
  }
  return false;
}