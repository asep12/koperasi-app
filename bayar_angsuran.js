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