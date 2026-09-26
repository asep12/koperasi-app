/**
 * ============================================================
 * SIMPANAN.GS — Input Setoran Simpanan
 * Koperasi App
 * Sesuai dokumen: REKAP_Koperasi_v2.md (BAGIAN 3, 12, 17)
 * ============================================================
 *
 * ALUR SETIAP SETORAN:
 * 1. Validasi (tahun aktif, anggota aktif, aturan nominal)
 * 2. Simpan ke transaksi_simpanan
 * 3. Simpan ke transaksi_kas (kas masuk)
 * 4. Simpan jurnal seimbang ke jurnal_umum
 * 5. Catat ke log_aktivitas
 *
 * Dipanggil dari menu (ui.gs, dibuat belakangan) lewat dialog,
 * TAPI bisa juga dipanggil langsung dari editor untuk testing
 * — lihat contoh di test_simpanan.gs.
 * ============================================================
 */

/**
 * Fungsi utama: input satu setoran simpanan.
 *
 * @param {Object} data
 * @param {string} data.id_anggota
 * @param {string} data.jenis_simpanan  'pokok' / 'wajib' / 'sukarela'
 * @param {number} data.jumlah_setoran
 * @param {string} [data.keterangan]
 * @param {Date}   [data.tanggal]  default: hari ini
 * @return {Object} { id_transaksi, pesan }
 */
function inputSetoranSimpanan(data) {
  // ---------- 1. VALIDASI ----------
  const tanggal = data.tanggal ? new Date(data.tanggal) : new Date();
  const tahun = tanggal.getFullYear();
  const bulan = tanggal.getMonth() + 1;

  validateTahunAktif(tahun);
  const anggota = pastikanAnggotaAktif(data.id_anggota);
  validateSetoran(data, tahun);

  // ---------- 2. SIMPAN transaksi_simpanan ----------
  const idTransaksi = generateId(SHEET.SIMPANAN, ID_PREFIX.SIMPANAN, 1, tahun, null);

  const rowSimpanan = {
    id_transaksi: idTransaksi,
    tanggal: tanggal,
    tahun: tahun,
    bulan: bulan,
    id_anggota: data.id_anggota,
    jenis_simpanan: data.jenis_simpanan,
    jumlah_setoran: data.jumlah_setoran,
    keterangan: data.keterangan || '',
    user_input: getUserEmail(),
    timestamp: new Date(),
    status_lock: STATUS_LOCK.OPEN
  };
  appendRowFromObject(SHEET.SIMPANAN, rowSimpanan);

  // ---------- 3. SIMPAN transaksi_kas (kas masuk) ----------
  catatKasMasuk({
    tanggal: tanggal,
    tahun: tahun,
    bulan: bulan,
    kategori: 'setoran',
    referensi: idTransaksi,
    keterangan: 'Setoran ' + data.jenis_simpanan + ' — ' + anggota.nama,
    nominal: data.jumlah_setoran
  });

  // ---------- 4. JURNAL (Debit Kas, Kredit akun simpanan sesuai jenis) ----------
  const akunSimpanan = data.jenis_simpanan === 'pokok'    ? AKUN.SIMPANAN_POKOK
                      : data.jenis_simpanan === 'wajib'   ? AKUN.SIMPANAN_WAJIB
                      : AKUN.SIMPANAN_SUKARELA;

  catatJurnal(tanggal, tahun, bulan, idTransaksi,
    'Setoran simpanan ' + data.jenis_simpanan + ' — ' + anggota.nama,
    [
      { kode_akun: AKUN.KAS,    debit: data.jumlah_setoran, kredit: 0 },
      { kode_akun: akunSimpanan, debit: 0, kredit: data.jumlah_setoran }
    ]
  );

  // ---------- 5. LOG ----------
  logAktivitas('INSERT', SHEET.SIMPANAN, idTransaksi, null, rowSimpanan);

  return {
    id_transaksi: idTransaksi,
    pesan: '✅ Setoran ' + data.jenis_simpanan + ' sebesar ' +
      formatRupiah(data.jumlah_setoran) + ' untuk ' + anggota.nama + ' berhasil disimpan.'
  };
}

// ============================================================
// VALIDASI SETORAN (BAGIAN 17)
// ============================================================

function validateSetoran(data, tahun) {
  // Jenis simpanan harus valid
  const jenisValid = ['pokok', 'wajib', 'sukarela'];
  if (jenisValid.indexOf(data.jenis_simpanan) === -1) {
    throw new Error('Jenis simpanan harus salah satu dari: ' + jenisValid.join(', '));
  }

  // Jumlah harus angka positif
  if (!data.jumlah_setoran || Number(data.jumlah_setoran) <= 0) {
    throw new Error('Jumlah setoran harus lebih dari 0.');
  }

  // Nominal mengikuti Pengaturan tahun buku (halaman Tahun Buku & RAT)
  const nominal = getNominalSimpanan(tahun);
  const jumlah = Number(data.jumlah_setoran);
  if (data.jenis_simpanan === 'sukarela' && jumlah < nominal.minimal_sukarela) {
    throw new Error('Simpanan sukarela minimal ' + formatRupiah(nominal.minimal_sukarela) + '.');
  }
  if (data.jenis_simpanan === 'pokok' && jumlah !== nominal.nominal_pokok) {
    throw new Error('Simpanan pokok tahun ' + tahun + ' harus ' +
      formatRupiah(nominal.nominal_pokok) + ' (sesuai Pengaturan).');
  }
  if (data.jenis_simpanan === 'wajib' && jumlah !== nominal.nominal_wajib) {
    throw new Error('Simpanan wajib tahun ' + tahun + ' harus ' +
      formatRupiah(nominal.nominal_wajib) + ' per bulan (sesuai Pengaturan).');
  }

  // Simpanan pokok hanya boleh 1x seumur hidup anggota
  if (data.jenis_simpanan === 'pokok') {
    const sudahPernah = getRowsByFilter(SHEET.SIMPANAN, function(row) {
      return row.id_anggota === data.id_anggota &&
             row.jenis_simpanan === 'pokok' &&
             row.status_lock !== STATUS_LOCK.VOID;
    });
    const dariMigrasi = getRowsByFilter(SHEET.SALDO_AWAL, function(row) {
      return row.id_anggota === data.id_anggota && row.jenis === 'simpanan_pokok';
    });
    if (sudahPernah.length > 0 || dariMigrasi.length > 0) {
      throw new Error('Anggota ini sudah pernah setor simpanan pokok ' +
        '(baik transaksi baru maupun migrasi). Simpanan pokok hanya sekali.');
    }
  }

  // Simpanan wajib: cegah dobel input bulan yang sama (opsional tapi membantu)
  if (data.jenis_simpanan === 'wajib') {
    const bulanIni = new Date(data.tanggal || new Date()).getMonth() + 1;
    const sudahSetorBulanIni = getRowsByFilter(SHEET.SIMPANAN, function(row) {
      return row.id_anggota === data.id_anggota &&
             row.jenis_simpanan === 'wajib' &&
             Number(row.tahun) === Number(tahun) &&
             Number(row.bulan) === Number(bulanIni) &&
             row.status_lock !== STATUS_LOCK.VOID;
    });
    if (sudahSetorBulanIni.length > 0) {
      throw new Error('Simpanan wajib bulan ini untuk anggota ini sudah pernah diinput ' +
        '(cek transaksi_simpanan). Jika ini koreksi, VOID dulu yang lama.');
    }
  }
}

// ============================================================
// PENGAMBILAN SIMPANAN SUKARELA (input awal — status PENDING)
// Approval-nya di file terpisah: approval.gs (dibuat belakangan)
// ============================================================

/**
 * Input pengajuan pengambilan sukarela. Status awal SELALU PENDING —
 * baru diproses (kas keluar + jurnal) setelah di-APPROVE lewat menu.
 */
function inputPengajuanPengambilan(data) {
  const tanggal = data.tanggal ? new Date(data.tanggal) : new Date();
  const tahun = tanggal.getFullYear();
  const bulan = tanggal.getMonth() + 1;

  validateTahunAktif(tahun);
  const anggota = pastikanAnggotaAktif(data.id_anggota);

  if (!data.jumlah || Number(data.jumlah) <= 0) {
    throw new Error('Jumlah pengambilan harus lebih dari 0.');
  }

  const saldoSukarela = getSaldoSimpanan(data.id_anggota, 'sukarela');
  if (Number(data.jumlah) > saldoSukarela) {
    throw new Error('Jumlah pengambilan (' + formatRupiah(data.jumlah) +
      ') melebihi saldo sukarela (' + formatRupiah(saldoSukarela) + ').');
  }

  const idPengambilan = generateId(SHEET.PENGAMBILAN, ID_PREFIX.PENGAMBILAN, 1, tahun, null);

  const row = {
    id_pengambilan: idPengambilan,
    tanggal: tanggal,
    tahun: tahun,
    bulan: bulan,
    id_anggota: data.id_anggota,
    jenis_simpanan: 'sukarela',
    jumlah: data.jumlah,
    status_approval: 'PENDING',
    keterangan: data.keterangan || '',
    user_input: getUserEmail(),
    timestamp: new Date(),
    status_lock: STATUS_LOCK.OPEN
  };
  appendRowFromObject(SHEET.PENGAMBILAN, row);
  logAktivitas('INSERT', SHEET.PENGAMBILAN, idPengambilan, null, row);

  return {
    id_pengambilan: idPengambilan,
    pesan: '✅ Pengajuan pengambilan ' + formatRupiah(data.jumlah) + ' untuk ' +
      anggota.nama + ' tersimpan dengan status PENDING. ' +
      'Perlu di-APPROVE lewat menu sebelum kas keluar & jurnal dibuat.'
  };
}

// ============================================================
// HELPER: catat kas masuk/keluar + jurnal (dipakai modul lain juga)
// ============================================================

/** Catat satu baris kas MASUK. Dipakai simpanan.gs, angsuran.gs, dll. */
function catatKasMasuk(p) {
  const idKas = generateId(SHEET.KAS, ID_PREFIX.KAS, 1, p.tahun, null);
  appendRowFromObject(SHEET.KAS, {
    id_kas: idKas,
    tanggal: p.tanggal,
    tahun: p.tahun,
    bulan: p.bulan,
    kategori: p.kategori,
    referensi: p.referensi,
    keterangan: p.keterangan,
    masuk: p.nominal,
    keluar: 0,
    user_input: getUserEmail(),
    timestamp: new Date(),
    status_lock: STATUS_LOCK.OPEN
  });
  return idKas;
}

/** Catat satu baris kas KELUAR. */
function catatKasKeluar(p) {
  const idKas = generateId(SHEET.KAS, ID_PREFIX.KAS, 1, p.tahun, null);
  appendRowFromObject(SHEET.KAS, {
    id_kas: idKas,
    tanggal: p.tanggal,
    tahun: p.tahun,
    bulan: p.bulan,
    kategori: p.kategori,
    referensi: p.referensi,
    keterangan: p.keterangan,
    masuk: 0,
    keluar: p.nominal,
    user_input: getUserEmail(),
    timestamp: new Date(),
    status_lock: STATUS_LOCK.OPEN
  });
  return idKas;
}

/**
 * Catat satu set entri jurnal yang HARUS balance (debit = kredit).
 * @param {Date} tanggal
 * @param {number} tahun
 * @param {number} bulan
 * @param {string} referensi   ID transaksi asal
 * @param {string} keterangan
 * @param {Array<{kode_akun:string, debit:number, kredit:number}>} entries
 */
function catatJurnal(tanggal, tahun, bulan, referensi, keterangan, entries) {
  validateJurnalBalance(entries);

  const coaList = sheetToObjects(SHEET.COA_AKUN);

  // skip baris nol (mis. jasa 0 pada pinjaman migrasi)
  const isi = entries.filter(function(e) {
    return !(Number(e.debit) === 0 && Number(e.kredit) === 0);
  });
  if (isi.length === 0) return;

  const idList = generateIdBeruntun(SHEET.JURNAL, ID_PREFIX.JURNAL, 1, tahun, null, isi.length);

  // Semua baris jurnal satu transaksi ditulis SEKALIGUS
  appendRowsFromObjects(SHEET.JURNAL, isi.map(function(entry, i) {
    const akun = coaList.find(function(a) { return String(a.kode) === String(entry.kode_akun); });
    const namaAkun = akun ? akun.nama_akun : '(akun tidak dikenal: ' + entry.kode_akun + ')';

    return {
      id_jurnal: idList[i],
      tanggal: tanggal,
      tahun: tahun,
      bulan: bulan,
      kode_akun: entry.kode_akun,
      nama_akun: namaAkun,
      debit: entry.debit,
      kredit: entry.kredit,
      referensi: referensi,
      keterangan: keterangan
    };
  }));
}

/** Validasi wajib: total debit harus sama dengan total kredit (BAGIAN 12). */
function validateJurnalBalance(entries) {
  const totalDebit = entries.reduce(function(s, e) { return s + Number(e.debit); }, 0);
  const totalKredit = entries.reduce(function(s, e) { return s + Number(e.kredit); }, 0);
  if (Math.abs(totalDebit - totalKredit) > 0.01) {
    throw new Error('Jurnal tidak balance! Debit=' + totalDebit + ', Kredit=' + totalKredit);
  }
}