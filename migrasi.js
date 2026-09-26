/**
 * ============================================================
 * MIGRASI.GS — Pindahkan pembukuan Excel 2026 ke aplikasi
 * Koperasi App
 * ============================================================
 *
 * Sumber: Excel pembukuan (diekstrak ke migrasi_data.js oleh tools/ekstrak_excel.py).
 * Hasil:  anggota, saldo per 31 Des 2025, lalu SEMUA transaksi
 *         Januari s/d bulan terakhir yang terisi di Excel.
 *
 * Langkah (dijalankan berurutan dari halaman Pengaturan → Import Data Awal):
 *   cadangan  → salinan spreadsheet ke Drive
 *   bersihkan → kosongkan data transaksi & anggota (pengguna, pengaturan,
 *               identitas, bagan akun TIDAK disentuh)
 *   awal      → anggota + saldo 31 Des 2025 + pinjaman berjalan
 *   bulan1..N → setoran, angsuran, pinjaman baru, pengambilan, kas lain
 *   selesai   → status anggota, angsuran rutin, rekonsiliasi akhir
 *
 * Tiap langkah ATOMIK: semua baris ditampung di memori (mode tulis massal)
 * dan dicocokkan dengan angka Excel (saldo tiap anggota, piutang, kas)
 * SEBELUM ditulis. Ada selisih → langkah dibatalkan, tidak ada yang tertulis.
 * ============================================================
 */

const KUNCI_MIGRASI = 'MIGRASI_EXCEL_2026';

function daftarLangkahMigrasi_() {
  const d = dataMigrasi_();
  const l = ['cadangan', 'bersihkan', 'awal'];
  for (let b = 1; b <= d.bulan_terisi; b++) l.push('bulan' + b);
  l.push('selesai');
  return l;
}

function statusMigrasi_() {
  const s = PropertiesService.getScriptProperties().getProperty(KUNCI_MIGRASI);
  return s ? JSON.parse(s) : { selesai: [] };
}
function simpanStatusMigrasi_(st) {
  PropertiesService.getScriptProperties().setProperty(KUNCI_MIGRASI, JSON.stringify(st));
}

/** Ringkasan untuk layar: isi Excel, kondisi data sekarang, langkah yang sudah jalan. */
/** migrasi_data.js dihapus dari Apps Script setelah migrasi (berisi data pribadi anggota). */
function adaDataMigrasi_() {
  return typeof DATA_MIGRASI !== 'undefined' || typeof DATA_MIGRASI_2026 !== 'undefined';
}
/** Aturan tahun buku untuk migrasi (bisa diatur di migrasi_data.js → "aturan"). */
function aturanMigrasi_() {
  return Object.assign({
    jasa_pinjaman: 1.5, jasa_sukarela: 1, metode_jasa_tahun_ini: 'terendah', metode_jasa_tahun_lalu: 'harian',
    nominal_pokok: 10000, nominal_wajib: 50000, minimal_sukarela: 10000
  }, dataMigrasi_().aturan || {});
}

/** Data hasil tools/ekstrak_excel.py (file migrasi_data.js). */
function dataMigrasi_() {
  if (typeof DATA_MIGRASI !== 'undefined') return DATA_MIGRASI;
  if (typeof DATA_MIGRASI_2026 !== 'undefined') return DATA_MIGRASI_2026;
  throw new Error('File migrasi_data.js belum ada. Buat dengan tools/ekstrak_excel.py lalu clasp push.');
}

function apiMigrasiStatus() {
  requireRole(['admin']);
  MODE_SENYAP = true;
  if (!adaDataMigrasi_()) return { tanpaData: true, selesai: statusMigrasi_().selesai };
  const d = dataMigrasi_();
  const jumlah = {};
  [SHEET.ANGGOTA, SHEET.SIMPANAN, SHEET.PINJAMAN, SHEET.ANGSURAN, SHEET.KAS].forEach(function(n) {
    jumlah[n] = sheetToObjects(n).length;
  });
  const akhir = d.bulan[d.bulan.length - 1];
  return {
    langkah: daftarLangkahMigrasi_(),
    selesai: statusMigrasi_().selesai,
    excel: {
      sumber: d.sumber, tahun: d.tahun, bulanTerisi: d.bulan_terisi,
      namaBulanAkhir: NAMA_BULAN_PANJANG[d.bulan_terisi - 1],
      anggota: d.anggota.length, kasAkhir: akhir.cek_kas.saldo_akhir,
      transaksi: d.bulan.reduce(function(s, b) {
        return s + Object.keys(b.setoran).length + Object.keys(b.ambil).length +
          Object.keys(b.angsuran).length + Object.keys(b.pinjaman).length + b.kas_lain.length;
      }, 0)
    },
    sekarang: {
      anggota: jumlah[SHEET.ANGGOTA], simpanan: jumlah[SHEET.SIMPANAN], pinjaman: jumlah[SHEET.PINJAMAN],
      angsuran: jumlah[SHEET.ANGSURAN], kas: jumlah[SHEET.KAS], saldoKas: getSaldoKas(), tahunAktif: getTahunAktif()
    }
  };
}

/** Jalankan satu langkah. d: { langkah } */
function apiMigrasiLangkah(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  if (!adaDataMigrasi_()) throw new Error('Data migrasi sudah dihapus dari aplikasi (migrasi telah selesai).');
  const langkah = String(d && d.langkah || '');
  const semua = daftarLangkahMigrasi_();
  const i = semua.indexOf(langkah);
  if (i === -1) throw new Error('Langkah migrasi tidak dikenal: ' + langkah);

  const st = statusMigrasi_();
  if (langkah === 'cadangan' || langkah === 'bersihkan') {
    // boleh diulang kapan saja (mulai dari awal)
  } else {
    const perlu = semua[i - 1];
    if (st.selesai.indexOf(perlu) === -1) throw new Error('Jalankan langkah "' + perlu + '" dulu.');
    if (st.selesai.indexOf(langkah) !== -1) throw new Error('Langkah "' + langkah + '" sudah pernah dijalankan.');
    if (i > semua.indexOf('awal')) pastikanHasilAwalLengkap_(st);
    if (/^bulan[0-9]+$/.test(langkah)) pastikanBulanBelumAda_(st, Number(langkah.replace('bulan', '')));
  }

  let pesan;
  if (langkah === 'cadangan') {
    pesan = migrasiCadangan_();
    st.selesai = ['cadangan'];
  } else if (langkah === 'bersihkan') {
    if (st.selesai.indexOf('cadangan') === -1) throw new Error('Buat cadangan dulu.');
    pesan = migrasiBersihkan_();
    st.selesai = ['cadangan', 'bersihkan'];
  } else {
    if (langkah === 'awal') {
      // mulai dari sheet bersih: percobaan sebelumnya bisa tertulis sebagian sebelum ditolak Sheets
      kosongkanSheetMigrasi_();
      perbaruiValidasiSheet();   // daftar pilihan lama menolak nilai baru
    }
    mulaiTulisMassal();
    try {
      if (langkah === 'awal') pesan = migrasiAwal_();
      else if (langkah === 'selesai') pesan = migrasiSelesai_();
      else pesan = migrasiBulan_(Number(langkah.replace('bulan', '')));
      const ditulis = selesaiTulisMassal();
      const r = Object.keys(ditulis).map(function(k) { return ditulis[k] + ' baris ' + k; });
      if (r.length) pesan += ' (' + r.join(', ') + ')';
    } catch (e) {
      batalTulisMassal();
      throw e;
    }
    st.selesai.push(langkah);
    if (langkah === 'selesai') {
      // log "status → keluar": dipakai laporan RAT untuk menghitung anggota keluar tahun ini
      __MIGRASI_KELUAR.forEach(function(id) {
        logAktivitas('EDIT', SHEET.ANGGOTA, id, { status: 'aktif' },
          { status: 'keluar', alasan: 'Migrasi Excel: seluruh simpanan sudah diambil' });
      });
    }
  }
  simpanStatusMigrasi_(st);
  logAktivitas('INSERT', 'migrasi', langkah, null, { pesan: pesan });
  return { langkah: langkah, pesan: pesan, selesai: st.selesai };
}

/**
 * Pastikan anggota & pinjaman dari langkah "awal" benar-benar ada di sheet.
 * Kalau tidak (mis. penulisan ditolak Google Sheets setelah langkah dicatat),
 * status dimundurkan ke "cadangan" supaya diulang dari "bersihkan" — tanpa data ganda.
 */
function pastikanHasilAwalLengkap_(st) {
  const d = dataMigrasi_();
  const anggota = sheetToObjects(SHEET.ANGGOTA).length;
  const pinjaman = sheetToObjects(SHEET.PINJAMAN).length;
  const perluPinjaman = Object.keys(d.awal.piutang).filter(function(k) { return d.awal.piutang[k] > 0; }).length;
  if (anggota !== d.anggota.length || pinjaman < perluPinjaman) {
    st.selesai = st.selesai.indexOf('cadangan') !== -1 ? ['cadangan'] : [];
    simpanStatusMigrasi_(st);
    throw new Error('Data langkah "Anggota & saldo" belum tersimpan lengkap di spreadsheet (anggota ' + anggota + ' dari ' +
      d.anggota.length + ', pinjaman ' + pinjaman + ' dari ' + perluPinjaman + '). Migrasi dimundurkan: klik ' +
      '"Lanjutkan migrasi" untuk mengulang mulai dari "Kosongkan data lama".');
  }
}

/** Bulan yang gagal di tengah penulisan bisa meninggalkan sebagian baris → ulang dari awal, bukan dobel. */
function pastikanBulanBelumAda_(st, b) {
  const T = dataMigrasi_().tahun;
  const ada = sheetToObjects(SHEET.KAS).some(function(r) { return Number(r.tahun) === T && Number(r.bulan) === b; });
  if (ada) {
    st.selesai = st.selesai.indexOf('cadangan') !== -1 ? ['cadangan'] : [];
    simpanStatusMigrasi_(st);
    throw new Error('Sebagian data ' + NAMA_BULAN_PANJANG[b - 1] + ' sudah tertulis dari percobaan sebelumnya. ' +
      'Migrasi dimundurkan: klik tombol migrasi lagi untuk mengulang dari "Kosongkan data lama" (tanpa data ganda).');
  }
}

// ============================================================
// LANGKAH-LANGKAH
// ============================================================

function migrasiCadangan_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cari = DriveApp.getFoldersByName(namaFolderBackup_());
  const folder = cari.hasNext() ? cari.next() : DriveApp.createFolder(namaFolderBackup_());
  const nama = 'SEBELUM MIGRASI ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH.mm') +
    ' — ' + ss.getName();
  const salinan = DriveApp.getFileById(ss.getId()).makeCopy(nama, folder);
  return 'Cadangan dibuat: "' + nama + '" di folder Drive "' + namaFolderBackup_() + '" (' + salinan.getUrl() + ').';
}

const SHEET_MIGRASI_DIKOSONGKAN = ['ANGGOTA', 'SALDO_AWAL', 'SIMPANAN', 'PENGAMBILAN', 'PINJAMAN', 'ANGSURAN',
  'KAS', 'JURNAL', 'JASA_SUKARELA'];

function kosongkanSheetMigrasi_() {
  const hasil = [];
  SHEET_MIGRASI_DIKOSONGKAN.forEach(function(k) {
    const sheet = getSheet(SHEET[k]);
    const n = sheet.getLastRow() - 1;
    if (n > 0) sheet.getRange(2, 1, n, Math.max(sheet.getLastColumn(), 1)).clearContent();
    hasil.push(SHEET[k] + ': ' + Math.max(n, 0));
    hapusCache(SHEET[k]);
  });
  SpreadsheetApp.flush();
  return hasil;
}

function migrasiBersihkan_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hasil = kosongkanSheetMigrasi_();
  ['DASHBOARD', 'lap_shu', 'lap_neraca', 'lap_labarugi', 'rekap_pinjaman', 'rekap_angsuran', 'rekap_simpanan',
    'struk_potongan'].forEach(function(n) { const s = ss.getSheetByName(n); if (s) s.clear(); });

  pastikanCOA();
  pastikanKolomPinjaman();
  pastikanKolomJasaSukarela();
  pastikanKolomNominalSimpanan();
  pastikanKolomAlokasi();
  perbaruiValidasiSheet();
  siapkanTahunBukuMigrasi_();
  return 'Data lama dikosongkan (' + hasil.join(', ') + '). Tahun buku ' + dataMigrasi_().tahun +
    ' aktif; jasa pinjaman ' + String(aturanMigrasi_().jasa_pinjaman).replace('.', ',') +
    '%/bulan menurun, jasa sukarela ' + String(aturanMigrasi_().jasa_sukarela).replace('.', ',') + '%.';
}

/** Setting 2025 (ditutup, aturan jasa lama) & 2026 (aktif, aturan RAT 2026). */
function siapkanTahunBukuMigrasi_() {
  const A = aturanMigrasi_();
  const T = dataMigrasi_().tahun;
  const semua = sheetToObjects(SHEET.SETTING_RAT);
  const contoh = semua.find(function(r) { return Number(r.tahun) === T; }) || semua[0] || {};
  [T - 1, T].forEach(function(th) {
    const nilai = {
      jasa_pinjaman: A.jasa_pinjaman, jasa_sukarela: A.jasa_sukarela,
      metode_jasa_sukarela: th === T ? A.metode_jasa_tahun_ini : A.metode_jasa_tahun_lalu,
      shu_simpanan: Number(contoh.shu_simpanan) || 40, shu_jasa: Number(contoh.shu_jasa) || 60,
      nominal_pokok: A.nominal_pokok, nominal_wajib: A.nominal_wajib, minimal_sukarela: A.minimal_sukarela,
      status_aktif: th === T
    };
    const ada = getSettingRAT(th);
    if (ada) updateRowByField(SHEET.SETTING_RAT, 'tahun', th, nilai);
    else {
      const baru = Object.assign({}, contoh, nilai, { tahun: th, tanggal_rat: '' });
      delete baru.__row;
      appendRowFromObject(SHEET.SETTING_RAT, baru);
    }
  });
  sheetToObjects(SHEET.SETTING_RAT).forEach(function(r) {
    if (Number(r.tahun) !== T && r.status_aktif === true) updateRowByRowNumber(SHEET.SETTING_RAT, r.__row, { status_aktif: false });
  });
  hapusCache(SHEET.SETTING_RAT);
}

/** Peta kode Excel (P1, N3, …) → ID anggota di aplikasi. PNS dulu, lalu Non PNS. */
function petaIdMigrasi_() {
  const urut = dataMigrasi_().anggota.slice().sort(function(a, b) {
    return (a.grup === b.grup ? 0 : a.grup === 'PNS' ? -1 : 1) || a.no - b.no;
  });
  const peta = {};
  urut.forEach(function(a, i) { peta[a.kode] = 'AGT-' + String(i + 1).padStart(4, '0'); });
  return { peta: peta, urut: urut };
}

function tgl_(b, hari) { return new Date(dataMigrasi_().tahun, b - 1, hari); }
const r2_ = function(x) { return Math.round((Number(x) || 0) * 100) / 100; };

/** Angsuran pokok berikutnya (bulan > b) yang tercatat di Excel untuk anggota ini. */
function angsuranBerikut_(kode, sesudahBulan) {
  const d = dataMigrasi_();
  for (let i = 0; i < d.bulan.length; i++) {
    if (d.bulan[i].bulan <= sesudahBulan) continue;
    const a = d.bulan[i].angsuran[kode];
    if (a && a[0] > 0) return a[0];
  }
  return 0;
}

function migrasiAwal_() {
  const d = dataMigrasi_(), T = d.tahun;
  const P = petaIdMigrasi_();
  const ket = 'Migrasi Excel ' + T + ' — posisi 31 Desember ' + (T - 1);

  // ---- anggota ----
  P.urut.forEach(function(a) {
    const s = d.awal.simpanan[a.kode] || [0, 0, 0];
    let masuk = '';
    d.bulan.forEach(function(b) {
      const st = b.setoran[a.kode];
      if (!masuk && st && st[0] > 0 && s[0] < 1) masuk = tgl_(b.bulan, 10);
    });
    appendRowFromObject(SHEET.ANGGOTA, {
      id_anggota: P.peta[a.kode], nip_nis: '', nama: a.nama, jabatan: a.grup, alamat: '', telepon: '',
      tanggal_masuk: masuk, status: 'aktif', sukarela_rutin: ''
    });
  });

  // ---- saldo awal per anggota ----
  const tahunSaldo = T - 1;
  const tambahSaldo = function(idAnggota, jenis, nominal) {
    if (Math.abs(nominal) < 0.005) return;
    appendRowFromObject(SHEET.SALDO_AWAL, {
      id: generateId(SHEET.SALDO_AWAL, ID_PREFIX.SALDO_AWAL, 1, tahunSaldo, null),
      tahun: tahunSaldo, id_anggota: idAnggota, jenis: jenis, nominal: r2_(nominal), keterangan: ket
    });
  };
  P.urut.forEach(function(a) {
    const s = d.awal.simpanan[a.kode] || [0, 0, 0];
    tambahSaldo(P.peta[a.kode], 'simpanan_pokok', s[0]);
    tambahSaldo(P.peta[a.kode], 'simpanan_wajib', s[1]);
    tambahSaldo(P.peta[a.kode], 'simpanan_sukarela', s[2]);
    tambahSaldo(P.peta[a.kode], 'piutang', d.awal.piutang[a.kode] || 0);
  });
  const k = d.awal.koperasi;
  ['kas', 'inventaris', 'penyusutan', 'cadangan', 'dana_pendidikan', 'dana_sosial', 'dana_pembangunan', 'shu']
    .forEach(function(j) { tambahSaldo('-', j, k[j] || 0); });
  const cek = cekSaldoAwal();
  if (Math.abs(cek.selisih) >= 5) {
    throw new Error('Saldo awal tidak seimbang (selisih Rp ' + cek.selisih.toFixed(2) + ').');
  }

  // ---- pinjaman berjalan per 31 Des ----
  let nPinjaman = 0;
  Object.keys(d.awal.piutang).forEach(function(kode) {
    const sisa = d.awal.piutang[kode];
    if (!(sisa > 0)) return;
    const angs = Math.min(angsuranBerikut_(kode, 0) || sisa, sisa);
    inputPinjamanMigrasi({
      id_anggota: P.peta[kode], sisa_pokok: sisa, tenor_sisa: Math.ceil(sisa / angs),
      angsuran_pokok: angs, jasa_persen: aturanMigrasi_().jasa_pinjaman, tanggal_asli: new Date(tahunSaldo, 11, 31),
      keterangan_tambahan: 'sisa per 31 Des ' + tahunSaldo + ' (Excel)'
    });
    nPinjaman++;
  });
  return P.urut.length + ' anggota, saldo 31 Des ' + tahunSaldo + ' seimbang (' + formatRupiah(cek.aktiva) +
    '), ' + nPinjaman + ' pinjaman berjalan';
}

function migrasiBulan_(b) {
  const d = dataMigrasi_(), T = d.tahun;
  const data = d.bulan[b - 1];
  if (!data || data.bulan !== b) throw new Error('Data bulan ' + b + ' tidak ada.');
  const P = petaIdMigrasi_();
  const nama = function(kode) { return (d.anggota.find(function(a) { return a.kode === kode; }) || {}).nama || kode; };
  const tanda = tandaPotongGaji(T, b) + ' (migrasi Excel)';
  PERINGATAN_MIGRASI = [];
  const n = { angsuran: 0, setoran: 0, pinjaman: 0, ambil: 0, lain: 0 };

  // 1. ANGSURAN (tgl 10) — pokok & jasa persis seperti Excel
  Object.keys(data.angsuran).forEach(function(kode) {
    const pokok = r2_(data.angsuran[kode][0]), jasa = r2_(data.angsuran[kode][1]);
    let induk = getPinjamanAktif(P.peta[kode])[0];
    let ketAngs = 'Generate otomatis — potong gaji (migrasi Excel)';
    if (!induk && pokok === 0 && jasa > 0) {
      // Excel mencatat jasa tanpa sisa pinjaman → dicatat pada pinjaman terakhir anggota (agar ikut SHU jasa)
      induk = getRowsByFilter(SHEET.PINJAMAN, function(p) {
        return p.id_anggota === P.peta[kode] && !String(p.id_induk || '') && p.status_lock !== STATUS_LOCK.VOID;
      }).pop();
      ketAngs = 'Jasa tanpa sisa pinjaman (sesuai Excel, perlu dicek) — migrasi Excel';
      if (induk) PERINGATAN_MIGRASI.push(NAMA_BULAN_PANJANG[b - 1] + ': ' + nama(kode) + ' membayar jasa ' +
        formatRupiah(jasa) + ' tanpa sisa pinjaman di Excel.');
    }
    if (!induk) throw new Error('Bulan ' + b + ': angsuran ' + nama(kode) + ' tetapi tidak ada pinjaman.\n' +
      diagnosaPinjamanMigrasi_(P.peta[kode]));
    const sisa = getSisaPinjaman(induk.id_pinjaman);
    if (pokok > sisa + 0.5) throw new Error('Bulan ' + b + ': angsuran ' + nama(kode) + ' melebihi sisa pinjaman.');
    const tanggal = tgl_(b, 10), total = r2_(pokok + jasa);
    const id = generateId(SHEET.ANGSURAN, ID_PREFIX.ANGSURAN, 1, T, null);
    appendRowFromObject(SHEET.ANGSURAN, {
      id_angsuran: id, tanggal: tanggal, tahun: T, bulan: b, id_pinjaman: induk.id_pinjaman,
      id_anggota: P.peta[kode], angsuran_pokok: pokok, jasa: jasa, denda: 0, total_bayar: total,
      keterangan: ketAngs, user_input: getUserEmail(),
      timestamp: new Date(), status_lock: STATUS_LOCK.OPEN
    });
    catatKasMasuk({ tanggal: tanggal, tahun: T, bulan: b, kategori: 'angsuran', referensi: id,
      keterangan: 'Angsuran pinjaman — ' + nama(kode), nominal: total });
    catatJurnal(tanggal, T, b, id, 'Angsuran pinjaman — ' + nama(kode), [
      { kode_akun: AKUN.KAS, debit: total, kredit: 0 },
      { kode_akun: AKUN.PIUTANG_PINJAMAN, debit: 0, kredit: pokok },
      { kode_akun: AKUN.PENDAPATAN_JASA_PINJAMAN, debit: 0, kredit: jasa }
    ]);
    if (induk.status === 'aktif') perbaruiStatusPinjamanJikaLunas(induk.id_pinjaman);
    n.angsuran++;
  });

  // 2. SETORAN SIMPANAN (tgl 10)
  Object.keys(data.setoran).forEach(function(kode) {
    ['pokok', 'wajib', 'sukarela'].forEach(function(jenis, j) {
      const jumlah = r2_(data.setoran[kode][j]);
      if (!jumlah) return;
      const tanggal = tgl_(b, 10);
      const id = generateId(SHEET.SIMPANAN, ID_PREFIX.SIMPANAN, 1, T, null);
      appendRowFromObject(SHEET.SIMPANAN, {
        id_transaksi: id, tanggal: tanggal, tahun: T, bulan: b, id_anggota: P.peta[kode],
        jenis_simpanan: jenis, jumlah_setoran: jumlah, keterangan: tanda, user_input: getUserEmail(),
        timestamp: new Date(), status_lock: STATUS_LOCK.OPEN
      });
      catatKasMasuk({ tanggal: tanggal, tahun: T, bulan: b, kategori: 'setoran', referensi: id,
        keterangan: 'Setoran ' + jenis + ' — ' + nama(kode), nominal: jumlah });
      catatJurnal(tanggal, T, b, id, 'Setoran simpanan ' + jenis + ' — ' + nama(kode), [
        { kode_akun: AKUN.KAS, debit: jumlah, kredit: 0 },
        { kode_akun: akunSimpananMigrasi_(jenis), debit: 0, kredit: jumlah }
      ]);
      n.setoran++;
    });
  });

  // 3. PINJAMAN BARU / TAMBAHAN (tgl 15)
  Object.keys(data.pinjaman).forEach(function(kode) {
    const nominal = r2_(data.pinjaman[kode]);
    const idAgt = P.peta[kode], tanggal = tgl_(b, 15);
    const induk = getPinjamanAktif(idAgt)[0] || null;
    const sisaLama = induk ? getSisaPinjaman(induk.id_pinjaman) : 0;
    const total = sisaLama + nominal;
    const angs = Math.min(angsuranBerikut_(kode, b) || Math.ceil(total / 10), total);
    const tenor = Math.max(1, Math.ceil(total / angs));
    const id = generateId(SHEET.PINJAMAN, ID_PREFIX.PINJAMAN, 1, T, null);
    appendRowFromObject(SHEET.PINJAMAN, {
      id_pinjaman: id, tanggal: tanggal, tahun: T, bulan: b, id_anggota: idAgt, nominal: nominal, tenor: tenor,
      jasa_persen: induk ? Number(induk.jasa_persen) : aturanMigrasi_().jasa_pinjaman, total_jasa: 0, total_tagihan: nominal,
      angsuran_perbulan: angs, status: induk ? 'tambahan' : 'aktif', keterangan: 'Migrasi Excel ' + T,
      user_input: getUserEmail(), timestamp: new Date(), status_lock: STATUS_LOCK.OPEN, metode_jasa: 'menurun',
      id_induk: induk ? induk.id_pinjaman : '', angsuran_sebelumnya: induk ? Number(induk.angsuran_perbulan) : ''
    });
    if (induk) updateRowByRowNumber(SHEET.PINJAMAN, induk.__row, { angsuran_perbulan: angs, tenor: tenor });
    const ketP = (induk ? 'Tambahan pinjaman — ' : 'Pencairan pinjaman — ') + nama(kode);
    catatKasKeluar({ tanggal: tanggal, tahun: T, bulan: b, kategori: 'pinjaman', referensi: id, keterangan: ketP, nominal: nominal });
    catatJurnal(tanggal, T, b, id, ketP, [
      { kode_akun: AKUN.PIUTANG_PINJAMAN, debit: nominal, kredit: 0 },
      { kode_akun: AKUN.KAS, debit: 0, kredit: nominal }
    ]);
    n.pinjaman++;
  });

  // 4. PENGAMBILAN SIMPANAN (tgl 20) — sudah disetujui
  Object.keys(data.ambil).forEach(function(kode) {
    ['pokok', 'wajib', 'sukarela'].forEach(function(jenis, j) {
      const jumlah = r2_(data.ambil[kode][j]);
      if (!jumlah) return;
      const tanggal = tgl_(b, 20);
      const id = generateId(SHEET.PENGAMBILAN, ID_PREFIX.PENGAMBILAN, 1, T, null);
      appendRowFromObject(SHEET.PENGAMBILAN, {
        id_pengambilan: id, tanggal: tanggal, tahun: T, bulan: b, id_anggota: P.peta[kode],
        jenis_simpanan: jenis, jumlah: jumlah, status_approval: 'APPROVED', keterangan: 'Migrasi Excel ' + T,
        user_input: getUserEmail(), timestamp: new Date(), status_lock: STATUS_LOCK.OPEN
      });
      catatKasKeluar({ tanggal: tanggal, tahun: T, bulan: b, kategori: 'pengambilan', referensi: id,
        keterangan: 'Pengambilan ' + jenis + ' — ' + nama(kode), nominal: jumlah });
      catatJurnal(tanggal, T, b, id, 'Pengambilan simpanan ' + jenis + ' — ' + nama(kode), [
        { kode_akun: akunSimpananMigrasi_(jenis), debit: jumlah, kredit: 0 },
        { kode_akun: AKUN.KAS, debit: 0, kredit: jumlah }
      ]);
      n.ambil++;
    });
  });

  // 5. KAS LAIN: pembagian SHU tahun lalu (tanpa kas) & pengeluaran/biaya (tgl 25)
  data.kas_lain.forEach(function(x) {
    if (x.jenis === 'alokasi_shu') { catatAlokasiSHUMigrasi_(x, b); n.lain++; return; }
    if (x.jenis === 'beban') {
      inputBebanOperasional({ kode_akun: x.akun, nominal: r2_(x.nominal), keterangan: x.ket + ' (migrasi Excel)',
        tanggal: tgl_(b, 25) });
      n.lain++;
      return;
    }
    if (x.jenis === 'masuk') {
      inputPemasukanLain({ nominal: r2_(x.nominal), keterangan: x.ket + ' (migrasi Excel)', tanggal: tgl_(b, 25) });
      n.lain++;
      return;
    }
    if (x.jenis === 'penyusutan') {
      inputPenyusutan({ nominal: r2_(x.nominal), keterangan: x.ket + ' (migrasi Excel)', tanggal: tgl_(b, 28) });
      n.lain++;
      return;
    }
    throw new Error('Jenis kas lain tidak dikenal: ' + x.jenis);
  });

  cocokkanDenganExcel_(data, P);
  return (PERINGATAN_MIGRASI.length ? '⚠️ ' + PERINGATAN_MIGRASI.join(' ') + ' ' : '') +
    NAMA_BULAN_PANJANG[b - 1] + ': ' + n.setoran + ' setoran, ' + n.angsuran + ' angsuran, ' + n.pinjaman +
    ' pinjaman, ' + n.ambil + ' pengambilan, ' + n.lain + ' kas lain — cocok dengan Excel (kas ' +
    formatRupiah(data.cek_kas.saldo_akhir) + ')';
}

function akunSimpananMigrasi_(jenis) {
  return jenis === 'pokok' ? AKUN.SIMPANAN_POKOK : jenis === 'wajib' ? AKUN.SIMPANAN_WAJIB : AKUN.SIMPANAN_SUKARELA;
}

/** Jurnal pembagian SHU tahun lalu sesuai keputusan RAT (tanpa kas). */
function catatAlokasiSHUMigrasi_(x, b) {
  const d = dataMigrasi_(), T = d.tahun;
  const persen = d.alokasi_shu_persen;
  const akun = { cadangan: '304', dana_anggota: '211', dana_pengurus: '212', dana_kesejahteraan: '213',
    dana_pendidikan: '214', dana_sosial: '215', dana_pembangunan: '216' };
  const total = r2_(x.nominal);
  const baris = [{ kode_akun: AKUN.SHU_DITAHAN, debit: total, kredit: 0 }];
  let terpakai = 0;
  const kunci = Object.keys(akun);
  kunci.forEach(function(k, i) {
    const nilai = i === kunci.length - 1 ? r2_(total - terpakai) : r2_(total * persen[k] / 100);
    terpakai = r2_(terpakai + nilai);
    baris.push({ kode_akun: akun[k], debit: 0, kredit: nilai });
  });
  catatJurnal(tgl_(b, 25), T, b, 'SHU-' + x.tahun_shu, 'Pembagian SHU ' + x.tahun_shu + ' sesuai keputusan RAT', baris);
}

/** Bandingkan hasil aplikasi dengan angka Excel akhir bulan; beda > Rp 1 → batalkan langkah. */
function cocokkanDenganExcel_(data, P) {
  const beda = [];
  const nama = function(kode) { return (dataMigrasi_().anggota.find(function(a) { return a.kode === kode; }) || {}).nama; };
  Object.keys(P.peta).forEach(function(kode) {
    const id = P.peta[kode];
    const exp = data.cek_simpanan[kode] || [0, 0, 0];
    ['pokok', 'wajib', 'sukarela'].forEach(function(j, i) {
      const s = getSaldoSimpanan(id, j);
      if (Math.abs(s - exp[i]) > 1) beda.push(nama(kode) + ' ' + j + ': aplikasi ' + formatRupiah(s) + ', Excel ' + formatRupiah(exp[i]));
    });
    const piutang = getPinjamanAktif(id).reduce(function(t, p) { return t + getSisaPinjaman(p.id_pinjaman); }, 0);
    const expP = (data.cek_piutang[kode] || [0, 0, 0, 0])[3];
    if (Math.abs(piutang - expP) > 1) beda.push(nama(kode) + ' piutang: aplikasi ' + formatRupiah(piutang) + ', Excel ' + formatRupiah(expP));
  });
  const kas = getSaldoKas();
  if (Math.abs(kas - data.cek_kas.saldo_akhir) > 1) {
    beda.push('Saldo kas: aplikasi ' + formatRupiah(kas) + ', Excel ' + formatRupiah(data.cek_kas.saldo_akhir));
  }
  if (beda.length) {
    throw new Error('Bulan ' + data.bulan + ' tidak cocok dengan Excel, tidak ada yang disimpan:\n• ' + beda.slice(0, 15).join('\n• '));
  }
}

function migrasiSelesai_() {
  const d = dataMigrasi_();
  const P = petaIdMigrasi_();
  const akhir = d.bulan[d.bulan.length - 1];
  const jumlah = { aktif: 0, nonaktif: 0, keluar: 0 };
  const keluarTahunIni = [];
  P.urut.forEach(function(a) {
    const id = P.peta[a.kode];
    const pw = getSaldoSimpanan(id, 'pokok') + getSaldoSimpanan(id, 'wajib');
    const suk = getSaldoSimpanan(id, 'sukarela');
    const sisa = getPinjamanAktif(id).reduce(function(t, p) { return t + getSisaPinjaman(p.id_pinjaman); }, 0);
    const status = pw >= 1 ? 'aktif' : (sisa >= 1 || suk >= 1 ? 'nonaktif' : 'keluar');
    const rutin = status === 'aktif' && akhir.setoran[a.kode] ? Math.round(akhir.setoran[a.kode][2]) : '';
    const baris = getRowByField(SHEET.ANGGOTA, 'id_anggota', id);
    updateRowByRowNumber(SHEET.ANGGOTA, baris.__row, { status: status, sukarela_rutin: rutin || '' });
    jumlah[status]++;
    const awal = d.awal.simpanan[a.kode] || [0, 0, 0];
    if (status === 'keluar' && awal[0] + awal[1] >= 1) keluarTahunIni.push(id);
  });

  // angsuran rutin = potongan terakhir di Excel; tenor menyesuaikan sisa
  let nPinjaman = 0;
  sheetToObjects(SHEET.PINJAMAN).filter(function(p) { return p.status === 'aktif' && !String(p.id_induk || ''); })
    .forEach(function(p) {
      const kode = Object.keys(P.peta).find(function(k) { return P.peta[k] === p.id_anggota; });
      let angs = 0;
      d.bulan.forEach(function(b) { const a = b.angsuran[kode]; if (a && a[0] > 0) angs = a[0]; });
      const sisa = getSisaPinjaman(p.id_pinjaman);
      if (angs > 0 && sisa > 0) {
        updateRowByRowNumber(SHEET.PINJAMAN, p.__row, { angsuran_perbulan: Math.min(angs, sisa), tenor: Math.ceil(sisa / Math.min(angs, sisa)) });
      }
      nPinjaman++;
    });
  // anggota yang keluar tahun ini dicatat di log SETELAH data tersimpan (lihat apiMigrasiLangkah)
  __MIGRASI_KELUAR = keluarTahunIni;
  const kas = getSaldoKas();
  return 'Anggota: ' + jumlah.aktif + ' aktif, ' + jumlah.nonaktif + ' nonaktif (masih ada pinjaman/sukarela), ' +
    jumlah.keluar + ' keluar. ' + nPinjaman + ' pinjaman aktif. Saldo kas ' + formatRupiah(kas) +
    ' = Excel ' + formatRupiah(akhir.cek_kas.saldo_akhir);
}
let __MIGRASI_KELUAR = [];
let PERINGATAN_MIGRASI = [];

/** Keterangan untuk pesan error: isi sheet pinjaman & anggota apa adanya. */
function diagnosaPinjamanMigrasi_(idAnggota) {
  const pj = sheetToObjects(SHEET.PINJAMAN);
  const agt = sheetToObjects(SHEET.ANGGOTA);
  const hitung = {};
  pj.forEach(function(p) { const k = JSON.stringify(p.status) + '/' + JSON.stringify(p.status_lock); hitung[k] = (hitung[k] || 0) + 1; });
  const idGanda = agt.length - Object.keys(agt.reduce(function(m, a) { m[a.id_anggota] = 1; return m; }, {})).length;
  const milik = pj.filter(function(p) { return String(p.id_anggota).trim() === idAnggota; });
  return 'DIAGNOSA — kolom transaksi_pinjaman: ' + getHeader(SHEET.PINJAMAN).join(', ') +
    ' | baris pinjaman: ' + pj.length + ' (status/lock: ' + JSON.stringify(hitung) + ')' +
    ' | milik ' + idAnggota + ': ' + JSON.stringify(milik.map(function(p) {
      return [p.id_pinjaman, p.id_anggota, p.status, p.status_lock, p.id_induk, p.__row]; })) +
    ' | contoh: ' + JSON.stringify(pj.slice(0, 2).map(function(p) { return [p.id_pinjaman, p.id_anggota, p.status, p.status_lock]; })) +
    ' | anggota: ' + agt.length + ' baris, ID ganda ' + idGanda +
    ' | saldo_awal: ' + sheetToObjects(SHEET.SALDO_AWAL).length + ' baris | langkah selesai: ' +
    JSON.stringify(statusMigrasi_().selesai);
}
