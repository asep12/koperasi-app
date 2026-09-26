/**
 * ============================================================
 * WEBAPP.GS — Server Aplikasi Web Penuh (v2)
 * Koperasi App
 * ============================================================
 *
 * MENGGANTIKAN: app.gs, ui.gs, form.html — HAPUS ketiganya dari
 * project Apps Script (supaya tidak ada dua doGet).
 * index.html juga DIGANTI TOTAL dengan versi baru.
 *
 * FITUR:
 * - Login berbasis akun Google (email dicek ke sheet "pengguna")
 * - Peran: admin (semua) / operator (harian, tanpa void, posting,
 *   approve, & kelola pengguna)
 * - CRUD Anggota, riwayat transaksi per anggota, VOID transaksi
 * - Semua endpoint input/approval/bulanan/laporan
 *
 * DEPLOY: Deploy → New deployment → Web app
 *   Execute as: Me | Who has access: Anyone with Google account
 *   (WAJIB "Anyone with Google account" agar email pengunjung
 *    terbaca untuk login)
 *
 * PENTING SAAT PERTAMA KALI:
 * Buka URL aplikasi dengan akun Google Anda sendiri (pemilik
 * script) — sistem otomatis mendaftarkan Anda sebagai ADMIN
 * pertama. Setelah itu tambahkan pengguna lain lewat halaman
 * Pengguna di dalam aplikasi.
 * ============================================================
 */

// ============================================================
// SHEET PENGGUNA (dibuat otomatis saat pertama dibutuhkan)
// ============================================================

const SHEET_PENGGUNA = 'pengguna';

function pastikanSheetPengguna() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_PENGGUNA);
  if (sheet) return sheet;

  sheet = ss.insertSheet(SHEET_PENGGUNA);
  const header = ['email', 'nama', 'peran', 'status', 'ditambahkan_oleh', 'timestamp'];
  sheet.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#0E4D3C');
  sheet.setFrozenRows(1);
  sheet.setTabColor('#0E4D3C');
  return sheet;
}

/**
 * Profil pengguna saat ini. null = belum masuk / nonaktif.
 *
 * Sumber utama: sesi login mandiri (__SESI_AKTIF, disetel oleh apiJalan
 * di sesi.gs). Session.getActiveUser() hanya dipakai sebagai cadangan
 * untuk pemanggilan dari editor Apps Script, karena Google TIDAK
 * memberikan email pengunjung web app di luar domain Workspace.
 */
function getProfilAktif() {
  pastikanSheetPengguna();

  let email = String(
    (typeof __SESI_AKTIF !== 'undefined' && __SESI_AKTIF) || ''
  ).trim().toLowerCase();

  if (!email) {
    email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  }
  if (!email) return null;

  const pemilik = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const daftar = sheetToObjects(SHEET_PENGGUNA);

  // Bootstrap: daftar kosong + yang membuka adalah pemilik script → jadi admin
  if (daftar.length === 0 && email === pemilik) {
    appendRowFromObject(SHEET_PENGGUNA, {
      email: email, nama: 'Pemilik Sistem', peran: 'admin', status: 'aktif',
      ditambahkan_oleh: 'sistem', timestamp: new Date()
    });
    return { email: email, nama: 'Pemilik Sistem', peran: 'admin' };
  }

  const p = daftar.find(function(u) {
    // trim() penting: email yang diketik/ditempel manual sering
    // membawa spasi tak terlihat di depan/belakang
    return String(u.email).trim().toLowerCase() === email;
  });
  if (!p) return null;

  // Status toleran: 'Aktif', 'AKTIF', ada spasi, atau kolom dikosongkan
  // saat baris ditambahkan manual di sheet → tetap dianggap aktif.
  const status = String(p.status || '').trim().toLowerCase();
  if (status !== '' && status !== 'aktif') return null;

  const peran = String(p.peran || '').trim().toLowerCase() === 'admin'
    ? 'admin' : 'operator';
  return { email: email, nama: p.nama || email.split('@')[0], peran: peran };
}

/**
 * DIAGNOSA AKSES — jalankan dari editor Apps Script (menu Run),
 * lalu buka View > Logs untuk melihat isi sheet `pengguna` apa adanya,
 * termasuk spasi tersembunyi dan huruf besar/kecil.
 */
function diagnosaAkses() {
  const daftar = sheetToObjects(SHEET_PENGGUNA);
  Logger.log('Jumlah baris terdaftar: ' + daftar.length);
  daftar.forEach(function(u, i) {
    Logger.log('Baris ' + (i + 2) +
      ' | email=[' + u.email + ']' +
      ' | panjang=' + String(u.email).length +
      ' (setelah trim=' + String(u.email).trim().length + ')' +
      ' | peran=[' + u.peran + ']' +
      ' | status=[' + u.status + ']');
  });
  Logger.log('Pemilik script: ' + Session.getEffectiveUser().getEmail());
  Logger.log('Email pembuka saat ini: [' +
    (Session.getActiveUser().getEmail() || '(KOSONG)') + ']');
  return daftar.length;
}

/** Penjaga peran di sisi server. */
function requireRole(daftarPeran) {
  const profil = getProfilAktif();
  if (!profil) throw new Error('Akses ditolak: akun Anda tidak terdaftar.');
  if (daftarPeran.indexOf(profil.peran) === -1) {
    throw new Error('Akses ditolak: fitur ini khusus ' + daftarPeran.join('/') + '.');
  }
  return profil;
}

// ============================================================
// PINTU MASUK WEB
// ============================================================

function doGet() {
  const judul = jalankanPublik_(function() {
    pastikanKolomSandi();
    pastikanKolomNominalSimpanan();
    pastikanKolomPinjaman();
    pastikanKolomAlokasi();
    pastikanCOA();
    pastikanKolomAnggota();
    return getProfilKoperasi();
  });
  let nama = String(judul.nama_koperasi || 'Koperasi').replace(/"/g, '').trim();
  if (nama === nama.toUpperCase()) nama = nama.toLowerCase().replace(/(^|\s)\S/g, function(h) { return h.toUpperCase(); });
  const kata = nama.split(/\s+/).filter(String);

  // Tidak ada lagi penolakan di sini: halaman selalu disajikan,
  // lalu klien menampilkan layar LOGIN. Identitas ditentukan oleh
  // sesi login mandiri (sesi.gs), bukan oleh akun Google pengunjung.
  const t = HtmlService.createTemplateFromFile('index');
  t.loginJudul = /^koperasi\b/i.test(nama) ? nama : 'Koperasi ' + nama;
  t.loginSub = judul.instansi || judul.jenis_koperasi || 'Aplikasi Pembukuan Koperasi';
  t.loginInisial = (kata.length > 1 ? kata[0][0] + kata[1][0] : nama.slice(0, 2)).toUpperCase();
  return t.evaluate()
    .setTitle(t.loginJudul)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// API: DATA UMUM
// ============================================================

function apiDataAwal() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const tahun = getTahunAktif();
  const setting = tahun ? getSettingRAT(tahun) : null;
  return {
    tahunAktif: tahun,
    jasaPinjaman: setting ? Number(setting.jasa_pinjaman) : 0,
    nominal: getNominalSimpanan(tahun),
    profil: getProfilKoperasi(),
    saldoKas: getSaldoKas()
  };
}

function apiDashboard() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;

  const aktif = getRowsByFilter(SHEET.ANGGOTA, function(a) { return a.status === 'aktif'; });
  const pinjaman = getRowsByFilter(SHEET.PINJAMAN, function(p) {
    return p.status === 'aktif' && p.status_lock !== STATUS_LOCK.VOID;
  });
  let piutang = 0;
  pinjaman.forEach(function(p) { piutang += getSisaPinjaman(p.id_pinjaman); });

  let simpanan = 0;
  aktif.forEach(function(a) {
    simpanan += getSaldoSimpanan(a.id_anggota, 'pokok') +
                getSaldoSimpanan(a.id_anggota, 'wajib') +
                getSaldoSimpanan(a.id_anggota, 'sukarela');
  });

  return {
    saldoKas: getSaldoKas(),
    jumlahAnggota: aktif.length,
    totalSimpanan: simpanan,
    jumlahPinjaman: pinjaman.length,
    totalPiutang: piutang,
    jumlahPending: getDaftarPending().length,
    tahunAktif: getTahunAktif()
  };
}

// ============================================================
// API: ANGGOTA (CRUD + DETAIL + RIWAYAT)
// ============================================================

function apiAnggotaList() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  return sheetToObjects(SHEET.ANGGOTA).map(function(a) {
    return {
      id: a.id_anggota, nip: a.nip_nis || '', nama: a.nama,
      jabatan: a.jabatan || '', telepon: a.telepon || '',
      status: a.status, sukarela_rutin: Number(a.sukarela_rutin) || 0
    };
  });
}

/** Setoran sukarela rutin (potong gaji) per bulan: kosong/0 atau ≥ minimal sukarela. */
function validasiSukarelaRutin(nilai) {
  const n = Math.round(Number(nilai) || 0);
  if (n < 0) throw new Error('Sukarela rutin tidak boleh minus.');
  const minimal = getNominalSimpanan(getTahunAktif()).minimal_sukarela;
  if (n > 0 && n < minimal) {
    throw new Error('Sukarela rutin minimal ' + formatRupiah(minimal) + ' (atau 0 bila tidak ada).');
  }
  return n;
}

function apiAnggotaTambah(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  if (!d.nama || String(d.nama).trim() === '') throw new Error('Nama wajib diisi.');

  pastikanKolomAnggota();
  const id = generateIdAnggota();
  const row = {
    id_anggota: id, nip_nis: d.nip || '', nama: String(d.nama).trim(),
    jabatan: d.jabatan || 'PNS', alamat: d.alamat || '', telepon: d.telepon || '',
    tanggal_masuk: new Date(), status: 'aktif',
    sukarela_rutin: validasiSukarelaRutin(d.sukarela_rutin)
  };
  appendRowFromObject(SHEET.ANGGOTA, row);
  logAktivitas('INSERT', SHEET.ANGGOTA, id, null, row);
  const pesan = '✅ Anggota ' + row.nama + ' terdaftar dengan ID ' + id + '.';

  // Opsional: langsung setor simpanan pokok (nominal dari Pengaturan)
  if (d.setorPokok) {
    try {
      const nominal = getNominalSimpanan(getTahunAktif());
      inputSetoranSimpanan({ id_anggota: id, jenis_simpanan: 'pokok',
        jumlah_setoran: nominal.nominal_pokok, keterangan: 'Setoran pokok saat mendaftar' });
      return pesan + '\nSimpanan pokok ' + formatRupiah(nominal.nominal_pokok) + ' tercatat.';
    } catch (e) {
      return pesan + '\n⚠️ Simpanan pokok BELUM tercatat: ' + e.message +
        ' — input manual di menu Transaksi.';
    }
  }
  return pesan;
}

function apiAnggotaEdit(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const lama = getAnggota(d.id);
  if (!lama) throw new Error('Anggota tidak ditemukan.');

  pastikanKolomAnggota();
  const baru = {
    nama: String(d.nama).trim(), nip_nis: d.nip || '',
    jabatan: d.jabatan || lama.jabatan, telepon: d.telepon || '',
    alamat: d.alamat || '',
    sukarela_rutin: validasiSukarelaRutin(d.sukarela_rutin)
  };
  updateRowByField(SHEET.ANGGOTA, 'id_anggota', d.id, baru);
  logAktivitas('EDIT', SHEET.ANGGOTA, d.id, lama, baru);
  return '✅ Data ' + baru.nama + ' diperbarui.';
}

function apiAnggotaStatus(id) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const lama = getAnggota(id);
  if (!lama) throw new Error('Anggota tidak ditemukan.');
  if (lama.status === 'keluar') {
    // Mengaktifkan kembali anggota yang sudah keluar hanya boleh admin
    // (biasanya setelah proses keluarnya di-VOID karena salah input).
    requireRole(['admin']);
  }
  const statusBaru = lama.status === 'aktif' ? 'nonaktif' : 'aktif';
  updateRowByField(SHEET.ANGGOTA, 'id_anggota', id, { status: statusBaru });
  logAktivitas('EDIT', SHEET.ANGGOTA, id, { status: lama.status }, { status: statusBaru });
  return '✅ ' + lama.nama + ' sekarang ' + statusBaru + '.';
}

/** Detail satu anggota: profil + saldo + seluruh riwayat. */
function apiAnggotaDetail(id) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const a = getAnggota(id);
  if (!a) throw new Error('Anggota tidak ditemukan.');

  const fmtTgl = function(t) {
    try { return Utilities.formatDate(new Date(t), Session.getScriptTimeZone(), 'dd/MM/yyyy'); }
    catch (e) { return String(t || ''); }
  };

  const setoran = getRowsByFilter(SHEET.SIMPANAN, function(r) { return r.id_anggota === id; })
    .map(function(r) { return { id: r.id_transaksi, tgl: fmtTgl(r.tanggal),
      jenis: r.jenis_simpanan, jumlah: Number(r.jumlah_setoran), lock: r.status_lock }; })
    .reverse();

  const pengambilan = getRowsByFilter(SHEET.PENGAMBILAN, function(r) { return r.id_anggota === id; })
    .map(function(r) { return { id: r.id_pengambilan, tgl: fmtTgl(r.tanggal),
      jumlah: Number(r.jumlah), approval: r.status_approval, lock: r.status_lock }; })
    .reverse();

  const pinjaman = getRowsByFilter(SHEET.PINJAMAN, function(r) { return r.id_anggota === id; })
    .map(function(r) { return { id: r.id_pinjaman, tgl: fmtTgl(r.tanggal),
      nominal: Number(r.nominal), tenor: r.tenor, tagihan: Number(r.total_tagihan),
      sisa: getSisaPinjaman(r.id_pinjaman), status: r.status, lock: r.status_lock,
      ket: r.keterangan || '' }; })
    .reverse();

  const angsuran = getRowsByFilter(SHEET.ANGSURAN, function(r) { return r.id_anggota === id; })
    .map(function(r) { return { id: r.id_angsuran, tgl: fmtTgl(r.tanggal),
      bulan: r.bulan + '/' + r.tahun, pokok: Number(r.angsuran_pokok),
      jasa: Number(r.jasa), total: Number(r.total_bayar), lock: r.status_lock }; })
    .reverse();

  return {
    profil: { id: a.id_anggota, nama: a.nama, nip: a.nip_nis || '',
      jabatan: a.jabatan || '', telepon: a.telepon || '', alamat: a.alamat || '',
      masuk: fmtTgl(a.tanggal_masuk), status: a.status,
      sukarela_rutin: Number(a.sukarela_rutin) || 0 },
    saldo: {
      pokok: getSaldoSimpanan(id, 'pokok'),
      wajib: getSaldoSimpanan(id, 'wajib'),
      sukarela: getSaldoSimpanan(id, 'sukarela')
    },
    setoran: setoran, pengambilan: pengambilan,
    pinjaman: pinjaman, angsuran: angsuran
  };
}

// ============================================================
// API: VOID TRANSAKSI (khusus admin — jejak audit tetap ada)
// ============================================================

/** Buat jurnal pembalik untuk semua entri jurnal ber-referensi tertentu. */
function buatJurnalBalik(referensi, alasan) {
  const entri = getRowsByFilter(SHEET.JURNAL, function(j) { return j.referensi === referensi; });
  const kini = new Date();
  entri.forEach(function(j) {
    const idJurnal = generateId(SHEET.JURNAL, ID_PREFIX.JURNAL, 1, Number(j.tahun), null);
    appendRowFromObject(SHEET.JURNAL, {
      id_jurnal: idJurnal, tanggal: kini, tahun: Number(j.tahun), bulan: Number(j.bulan),
      kode_akun: j.kode_akun, nama_akun: j.nama_akun,
      debit: Number(j.kredit), kredit: Number(j.debit),   // DIBALIK
      referensi: 'VOID-' + referensi,
      keterangan: 'Pembalikan (void): ' + (alasan || referensi)
    });
  });
  return entri.length;
}

/** Void baris kas yang menunjuk referensi tertentu. */
function voidKasByReferensi(referensi) {
  const rows = getRowsByFilter(SHEET.KAS, function(k) {
    return k.referensi === referensi && k.status_lock === STATUS_LOCK.OPEN;
  });
  rows.forEach(function(k) {
    updateRowByRowNumber(SHEET.KAS, k.__row, { status_lock: STATUS_LOCK.VOID });
  });
  return rows.length;
}

function apiVoidTransaksi(d) {
  const profil = requireRole(['admin']);
  MODE_SENYAP = true;
  const jenis = d.jenis, id = d.id, alasan = d.alasan || '';

  if (jenis === 'setoran') {
    const row = getRowByField(SHEET.SIMPANAN, 'id_transaksi', id);
    if (!row) throw new Error('Setoran tidak ditemukan.');
    if (row.status_lock !== STATUS_LOCK.OPEN)
      throw new Error('Hanya transaksi OPEN yang bisa di-void (status: ' + row.status_lock + ').');
    updateRowByRowNumber(SHEET.SIMPANAN, row.__row, { status_lock: STATUS_LOCK.VOID,
      keterangan: (row.keterangan ? row.keterangan + ' | ' : '') + 'VOID: ' + alasan });
    voidKasByReferensi(id);
    buatJurnalBalik(id, alasan);
    logAktivitas('VOID', SHEET.SIMPANAN, id, row, { alasan: alasan, oleh: profil.email });
    return '✅ Setoran ' + id + ' di-void. Kas & jurnal sudah dibalik.';
  }

  if (jenis === 'angsuran') {
    const row = getRowByField(SHEET.ANGSURAN, 'id_angsuran', id);
    if (!row) throw new Error('Angsuran tidak ditemukan.');
    if (row.status_lock !== STATUS_LOCK.OPEN)
      throw new Error('Hanya transaksi OPEN yang bisa di-void.');
    updateRowByRowNumber(SHEET.ANGSURAN, row.__row, { status_lock: STATUS_LOCK.VOID,
      keterangan: (row.keterangan ? row.keterangan + ' | ' : '') + 'VOID: ' + alasan });
    voidKasByReferensi(id);
    buatJurnalBalik(id, alasan);
    // Pinjaman yang tadinya lunas bisa kembali aktif
    const sisa = getSisaPinjaman(row.id_pinjaman);
    if (sisa > 0) updateRowByField(SHEET.PINJAMAN, 'id_pinjaman', row.id_pinjaman,
      { status: 'aktif' });
    logAktivitas('VOID', SHEET.ANGSURAN, id, row, { alasan: alasan, oleh: profil.email });
    return '✅ Angsuran ' + id + ' di-void. Sisa pinjaman kembali ' + formatRupiah(sisa) + '.';
  }

  if (jenis === 'pengambilan') {
    const row = getRowByField(SHEET.PENGAMBILAN, 'id_pengambilan', id);
    if (!row) throw new Error('Pengambilan tidak ditemukan.');
    if (row.status_lock !== STATUS_LOCK.OPEN)
      throw new Error('Hanya transaksi OPEN yang bisa di-void.');
    updateRowByRowNumber(SHEET.PENGAMBILAN, row.__row, { status_lock: STATUS_LOCK.VOID,
      keterangan: (row.keterangan ? row.keterangan + ' | ' : '') + 'VOID: ' + alasan });
    if (row.status_approval === 'APPROVED') {
      voidKasByReferensi(id);
      buatJurnalBalik(id, alasan);
    }
    logAktivitas('VOID', SHEET.PENGAMBILAN, id, row, { alasan: alasan, oleh: profil.email });
    return '✅ Pengambilan ' + id + ' di-void.';
  }

  if (jenis === 'pinjaman') {
    const row = getRowByField(SHEET.PINJAMAN, 'id_pinjaman', id);
    if (!row) throw new Error('Pinjaman tidak ditemukan.');
    if (row.status_lock !== STATUS_LOCK.OPEN)
      throw new Error('Hanya transaksi OPEN yang bisa di-void.');
    const adaAngsuran = getRowsByFilter(SHEET.ANGSURAN, function(r) {
      return r.id_pinjaman === id && r.status_lock !== STATUS_LOCK.VOID;
    });
    if (adaAngsuran.length > 0)
      throw new Error('Pinjaman ini sudah punya ' + adaAngsuran.length +
        ' angsuran. Void semua angsurannya dulu.');
    const tambahan = getRowsByFilter(SHEET.PINJAMAN, function(r) {
      return String(r.id_induk || '') === String(id) && r.status_lock !== STATUS_LOCK.VOID;
    });
    if (tambahan.length > 0)
      throw new Error('Pinjaman ini punya pinjaman tambahan (' +
        tambahan.map(function(r) { return r.id_pinjaman; }).join(', ') +
        '). Void pinjaman tambahannya dulu.');

    // Pinjaman TAMBAHAN: sisa induk tidak boleh jadi minus setelah dibatalkan
    const idInduk = String(row.id_induk || '');
    const induk = idInduk ? getRowByField(SHEET.PINJAMAN, 'id_pinjaman', idInduk) : null;
    if (induk && getSisaPinjaman(idInduk) - Number(row.nominal) < 0)
      throw new Error('Angsuran ' + idInduk + ' yang sudah dibayar melebihi pinjaman ' +
        'tanpa tambahan ini. Void angsuran terakhirnya dulu.');

    updateRowByRowNumber(SHEET.PINJAMAN, row.__row, { status_lock: STATUS_LOCK.VOID,
      status: induk ? 'tambahan' : 'lunas',
      keterangan: (row.keterangan ? row.keterangan + ' | ' : '') + 'VOID: ' + alasan });
    voidKasByReferensi(id);
    buatJurnalBalik(id, alasan);

    if (induk) {
      // Kembalikan angsuran pokok induk seperti sebelum tambahan
      const pulih = {};
      if (row.angsuran_sebelumnya !== '' && row.angsuran_sebelumnya !== undefined)
        pulih.angsuran_perbulan = Number(row.angsuran_sebelumnya);
      pulih.status = getSisaPinjaman(idInduk) > 0 ? 'aktif' : 'lunas';
      updateRowByRowNumber(SHEET.PINJAMAN, induk.__row, pulih);
      logAktivitas('EDIT', SHEET.PINJAMAN, idInduk,
        { angsuran_perbulan: induk.angsuran_perbulan, status: induk.status },
        Object.assign({ void_tambahan: id }, pulih));
    }
    logAktivitas('VOID', SHEET.PINJAMAN, id, row, { alasan: alasan, oleh: profil.email });
    return '✅ Pinjaman ' + id + ' di-void.' +
      (induk ? ' Sisa ' + idInduk + ' kembali ' + formatRupiah(getSisaPinjaman(idInduk)) + '.' : '');
  }

  throw new Error('Jenis transaksi tidak dikenal: ' + jenis);
}

// ============================================================
// API: INPUT TRANSAKSI (memakai mesin yang sudah ada)
// ============================================================

function apiSubmitSetoran(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  return inputSetoranSimpanan({
    id_anggota: d.id_anggota, jenis_simpanan: d.jenis_simpanan,
    jumlah_setoran: Number(d.jumlah), keterangan: d.keterangan || ''
  }).pesan;
}

function apiSubmitPengambilan(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  return inputPengajuanPengambilan({
    id_anggota: d.id_anggota, jumlah: Number(d.jumlah), keterangan: d.keterangan || ''
  }).pesan;
}

function apiSubmitPinjaman(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const hasil = inputPinjamanBaru({
    id_anggota: d.id_anggota, nominal: Number(d.nominal),
    tenor: Number(d.tenor), keterangan: d.keterangan || '',
    jasa_persen: d.jasa_persen, angsuran_pokok: d.angsuran_pokok
  });
  return { pesan: hasil.pesan, struk: dataStruk('pinjaman', hasil.id_pinjaman) };
}

function apiSubmitMigrasi(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  return inputPinjamanMigrasi({
    id_anggota: d.id_anggota, sisa_pokok: Number(d.sisa_pokok),
    tenor_sisa: Number(d.tenor_sisa),
    angsuran_pokok: d.angsuran_pokok, jasa_persen: d.jasa_persen,
    tanggal_asli: new Date(Number(d.tahun_asli), Number(d.bulan_asli) - 1, 1),
    keterangan_tambahan: d.keterangan || ''
  }).pesan;
}

function apiSaldoAnggota(id) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  return {
    pokok: getSaldoSimpanan(id, 'pokok'),
    wajib: getSaldoSimpanan(id, 'wajib'),
    sukarela: getSaldoSimpanan(id, 'sukarela')
  };
}

// ============================================================
// API: APPROVAL
// ============================================================

function apiPendingList() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  return getDaftarPending().map(function(r) {
    const a = getAnggota(r.id_anggota);
    return {
      id: r.id_pengambilan, nama: a ? a.nama : r.id_anggota,
      jumlah: Number(r.jumlah),
      tanggal: Utilities.formatDate(new Date(r.tanggal),
        Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      keterangan: r.keterangan || '',
      saldo: getSaldoSimpanan(r.id_anggota, 'sukarela')
    };
  });
}

function apiApprove(id) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  return approvePengambilan(id).pesan;
}

function apiReject(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  return rejectPengambilan(d.id, d.alasan).pesan;
}

// ============================================================
// API: PROSES BULANAN
// ============================================================

function apiGenerateAngsuran() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const h = generateAngsuranBulanan();
  return '✅ Generate angsuran selesai.\nBerhasil: ' + h.berhasil.length +
    '\nDilewati (sudah ada): ' + h.dilewati.length +
    (h.gagal.length ? '\nGagal: ' + h.gagal.length + '\n' + h.gagal.join('\n') : '');
}

function apiCetakStruk() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const h = cetakStrukPotongan();
  return '✅ Struk siap di sheet struk_potongan.\nAnggota: ' + h.jumlahAnggota +
    '\nTotal potongan: ' + formatRupiah(h.totalPotongan);
}

function apiHitungJasa() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const h = hitungJasaSukarelaBulanan();
  return '✅ Hitung jasa selesai (DRAFT).\nDibuat: ' + h.dibuat.length +
    '\nDilewati (saldo 0): ' + h.dilewatiSaldoNol.length +
    '\nDilewati (sudah ada): ' + h.dilewatiSudahAda.length;
}

function apiPostingJasa() {
  requireRole(['admin']);
  MODE_SENYAP = true;
  const h = postingJasaSukarela();
  return h.diposting > 0
    ? '✅ Posting selesai. ' + h.diposting + ' anggota, total ' +
      formatRupiah(h.totalJasa) + '. Data terkunci.'
    : 'Tidak ada DRAFT untuk diposting.';
}

// ============================================================
// API: LAPORAN
// ============================================================

function apiLaporan(nama) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  if (nama === 'labarugi') {
    return '✅ Laba Rugi selesai. SHU Kotor: ' + formatRupiah(generateLaporanLabaRugi());
  }
  if (nama === 'neraca') {
    const h = generateLaporanNeraca();
    return '✅ Neraca selesai. Aktiva ' + formatRupiah(h.totalAktiva) +
      (h.balance ? ' — BALANCE' : ' — ❌ TIDAK BALANCE');
  }
  if (nama === 'shu') {
    const h = generateLaporanSHU();
    return '✅ Laporan SHU selesai. Netto anggota: ' + formatRupiah(h.alokasi.shuNetto);
  }
  if (nama === 'rekap_simpanan') { generateRekapSimpanan(); return '✅ Rekap simpanan selesai.'; }
  if (nama === 'rekap_pinjaman') { generateRekapPinjaman(); return '✅ Rekap pinjaman selesai.'; }
  if (nama === 'semua') {
    generateLaporanLabaRugi(); generateLaporanNeraca(); generateLaporanSHU();
    generateRekapSimpanan(); generateRekapPinjaman();
    return '✅ Semua laporan selesai digenerate.';
  }
  throw new Error('Jenis laporan tidak dikenal.');
}

function apiDataSHU() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const tahun = getTahunAktif();
  const alokasi = hitungAlokasiDana(tahun);
  const setting = getSettingRAT(tahun);
  const baris = daftarSHUAnggota(tahun).map(function(r) {
    return { nama: r.nama, status: r.status, simpanan: r.simpanan, jasa: r.jasa,
      shuSimpanan: r.shuSimpanan, shuJasa: r.shuJasa, total: r.total };
  });
  return { tahun: tahun, alokasi: alokasi, baris: baris,
    shuSimpananPersen: Number(setting.shu_simpanan), shuJasaPersen: Number(setting.shu_jasa) };
}

// ============================================================
// API: KELOLA PENGGUNA (khusus admin)
// ============================================================

function apiPenggunaList() {
  requireRole(['admin']);
  MODE_SENYAP = true;
  return sheetToObjects(SHEET_PENGGUNA).map(function(u) {
    return { email: u.email, nama: u.nama, peran: u.peran, status: u.status };
  });
}

function apiPenggunaTambah(d) {
  const profil = requireRole(['admin']);
  MODE_SENYAP = true;
  const email = String(d.email || '').trim().toLowerCase();
  if (!email || email.indexOf('@') === -1) throw new Error('Email tidak valid.');

  const sudahAda = getRowsByFilter(SHEET_PENGGUNA, function(u) {
    return String(u.email).toLowerCase() === email;
  });
  if (sudahAda.length > 0) throw new Error('Email ini sudah terdaftar.');

  validasiKekuatanSandi(d.sandi);
  const garam = buatGaram();

  const peran = d.peran === 'admin' ? 'admin' : 'operator';
  appendRowFromObject(SHEET_PENGGUNA, {
    email: email, nama: d.nama || email.split('@')[0], peran: peran,
    status: 'aktif', ditambahkan_oleh: profil.email, timestamp: new Date(),
    sandi_hash: hashSandi(String(d.sandi), garam), garam: garam,
    wajib_ganti: 'ya', terakhir_login: ''
  });
  logAktivitas('INSERT', SHEET_PENGGUNA, email, null, { peran: peran });
  return '✅ ' + email + ' terdaftar sebagai ' + peran + '.';
}

function apiPenggunaStatus(email) {
  const profil = requireRole(['admin']);
  MODE_SENYAP = true;
  const target = String(email).toLowerCase();
  if (target === profil.email) throw new Error('Tidak bisa menonaktifkan akun sendiri.');

  const row = getRowsByFilter(SHEET_PENGGUNA, function(u) {
    return String(u.email).toLowerCase() === target;
  })[0];
  if (!row) throw new Error('Pengguna tidak ditemukan.');

  const baru = row.status === 'aktif' ? 'nonaktif' : 'aktif';
  updateRowByRowNumber(SHEET_PENGGUNA, row.__row, { status: baru });
  logAktivitas('EDIT', SHEET_PENGGUNA, target, { status: row.status }, { status: baru });
  return '✅ ' + target + ' sekarang ' + baru + '.';
}