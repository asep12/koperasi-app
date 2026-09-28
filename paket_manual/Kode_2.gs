/**
 * Kode_2.gs — bagian 2 dari 3 kode Koperasi App.
 * Hak Cipta (c) Asep Hanuryana (Threads: @asep94). Seluruh hak dilindungi; lihat LICENSE.
 *
 * FILE INI DIBUAT OTOMATIS oleh tools/buat_paket_manual.py. Jangan disunting di sini:
 * ubah file aslinya di repositori, lalu jalankan ulang alat tersebut.
 *
 * Berisi: migrasi.js, pinjaman.js, profil_koperasi.js, rat.js, reset_data.js, ringkasan_struk.js, sesi.js, setup.js, simpanan.js, tagihan.js
 */

// ===== migrasi.js =====
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

// ===== pinjaman.js =====
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

// ===== profil_koperasi.js =====
/**
 * ============================================================
 * PROFIL_KOPERASI.GS — Identitas Koperasi (kop struk & laporan)
 * Koperasi App
 * ============================================================
 *
 * Disimpan di sheet `profil_koperasi` (kolom: kunci | nilai),
 * dibuat otomatis saat pertama kali disimpan dari aplikasi.
 * Isian mengikuti sheet "Ident_Kop" di Excel pembukuan lama:
 * jenis, nama, badan hukum, alamat, pengurus & pengawas.
 *
 * Logo disimpan sebagai data URL (gambar kecil, sudah diperkecil
 * di browser) — batas satu sel Google Sheets 50.000 karakter.
 * ============================================================
 */

const SHEET_PROFIL = 'profil_koperasi';

/** Lambang Koperasi Indonesia (dari kop surat laporan RAT), PNG 260×265. */
const LOGO_KOPERASI_INDONESIA = 'data:image/png;base64,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';
/** Isi sel logo bila admin sengaja menghapus logo (supaya tidak kembali ke bawaan). */
const TANPA_LOGO = 'tanpa';

const PROFIL_BAWAAN = {
  // Identitas asli koperasi diisi dari identitas_lokal.js (tidak masuk Git) atau
  // dari Pengaturan → Identitas Koperasi di aplikasi.
  jenis_koperasi: 'KOPERASI SIMPAN PINJAM',
  nama_koperasi: 'KOPERASI',
  no_badan_hukum: '',
  alamat: '',
  kota: '',                     // dipakai di "<kota>, 31 Desember <tahun>"
  instansi: '',                 // sekolah/kantor tempat koperasi (untuk laporan RAT)
  telepon: '',
  email: '',
  // Nama pengurus & pengawas: diisi dari identitas_lokal.js (tidak masuk Git)
  ketua: '', sekretaris: '', bendahara: '',
  pengurus_anggota: '',        // satu nama per baris (untuk laporan RAT)
  pengawas_ketua: '', pengawas_anggota: '',
  logo: LOGO_KOPERASI_INDONESIA
};

const BATAS_LOGO = 45000; // karakter data URL

/** Profil lengkap (nilai kosong di sheet → pakai bawaan). */
function getProfilKoperasi() {
  return cacheTurunan('profilKoperasi', function() {
    const lokal = typeof identitasLokal_ === 'function' ? identitasLokal_() : {};
    const hasil = Object.assign({}, PROFIL_BAWAAN, lokal);
    const ada = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PROFIL);
    if (!ada) return hasil;
    sheetToObjects(SHEET_PROFIL).forEach(function(r) {
      const k = String(r.kunci || '').trim();
      if (PROFIL_BAWAAN.hasOwnProperty(k) && String(r.nilai || '').trim() !== '') {
        hasil[k] = String(r.nilai);
      }
    });
    if (hasil.logo === TANPA_LOGO) hasil.logo = '';   // logo sengaja dihapus admin
    return hasil;
  });
}

function pastikanSheetProfil_() {
  pastikanAksesData();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_PROFIL);
  if (sheet) return sheet;
  sheet = ss.insertSheet(SHEET_PROFIL);
  sheet.getRange(1, 1, 1, 2).setValues([['kunci', 'nilai']])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#0E4D3C');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 420);
  return sheet;
}

// ============================================================
// API
// ============================================================

function apiProfilKoperasi() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  return getProfilKoperasi();
}

function apiProfilSimpan(d) {
  const profil = requireRole(['admin']);
  MODE_SENYAP = true;
  d = d || {};

  const baru = {};
  Object.keys(PROFIL_BAWAAN).forEach(function(k) {
    baru[k] = String(d[k] === undefined || d[k] === null ? '' : d[k]).trim();
  });
  if (!baru.nama_koperasi) throw new Error('Nama koperasi wajib diisi.');
  if (baru.logo && baru.logo !== TANPA_LOGO) {
    if (!/^data:image\/(png|jpeg|webp);base64,[A-Za-z0-9+/=]+$/.test(baru.logo)) {
      throw new Error('Format logo tidak dikenali. Gunakan gambar PNG atau JPG.');
    }
    if (baru.logo.length > BATAS_LOGO) {
      throw new Error('Logo terlalu besar. Pilih gambar yang lebih sederhana/kecil.');
    }
  }

  const lama = getProfilKoperasi();
  const sheet = pastikanSheetProfil_();
  const baris = Object.keys(PROFIL_BAWAAN).map(function(k) { return [k, baru[k]]; });
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  sheet.getRange(2, 1, baris.length, 2).setNumberFormat('@').setValues(baris);
  hapusCache(SHEET_PROFIL);

  const catatLama = Object.assign({}, lama), catatBaru = Object.assign({}, baru);
  catatLama.logo = lama.logo ? '(ada logo)' : '';
  catatBaru.logo = baru.logo ? '(ada logo)' : '';
  logAktivitas('EDIT', SHEET_PROFIL, 'identitas', catatLama,
    Object.assign(catatBaru, { oleh: profil.email }));
  return '✅ Identitas koperasi disimpan.';
}

// ===== rat.js =====
/**
 * ============================================================
 * RAT.GS — Laporan Rapat Anggota Tahunan (file Word .docx)
 * Koperasi App
 * ============================================================
 *
 * Menyusun paket laporan RAT (mengikuti "Pembukuan Koperasi 2025.docx"):
 *   1. Undangan RAT            4. Laporan pertanggungjawaban pengurus
 *   2. Susunan acara           5. Laporan badan pengawas + analisis keuangan
 *   3. Tata tertib             6. Rencana kerja tahun berikutnya
 *
 * Angka (SHU + terbilang, keanggotaan, aktiva/pasiva, likuiditas,
 * solvabilitas, rentabilitas, nominal simpanan, jasa, pembagian SHU)
 * diambil dari data aplikasi. Kalimat-kalimat tetap bisa disunting di Word.
 *
 * File .docx disusun langsung sebagai XML (WordprocessingML) lalu di-zip
 * dengan Utilities.zip — tanpa pustaka luar, tanpa izin Google Docs.
 * ============================================================
 */

const FOLDER_RAT = 'Laporan RAT Koperasi';
const KOLOM_RAT = ['nomor_undangan_rat', 'tanggal_undangan_rat', 'waktu_rat', 'tempat_rat',
  'acara_tambahan_rat', 'predikat_kesehatan', 'pinjaman_maksimal', 'tenor_maksimal'];
const BAWAAN_RAT = {
  waktu_rat: '13.00 WIB s.d. selesai', tempat_rat: '',
  acara_tambahan_rat: 'Door Prize', predikat_kesehatan: 'SEHAT',
  pinjaman_maksimal: 60000000, tenor_maksimal: 30
};

// ============================================================
// DATA
// ============================================================

/** Data untuk form & laporan RAT satu tahun buku (tanpa menulis apa pun). */
function apiDataRAT(tahun) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  return kumpulkanDataRAT_(Number(tahun) || tahunRATBawaan_());
}

/** RAT biasanya digelar awal tahun untuk tahun buku sebelumnya. */
function tahunRATBawaan_() {
  const aktif = getTahunAktif() || new Date().getFullYear();
  if (new Date().getMonth() < 6 && getSettingRAT(aktif - 1)) return aktif - 1;
  return aktif;
}

function kumpulkanDataRAT_(tahun) {
  const setting = getSettingRAT(tahun);
  if (!setting) throw new Error('Tahun buku ' + tahun + ' belum ada di Pengaturan.');
  const rencana = getSettingRAT(tahun + 1) || setting;
  const nominalRencana = getNominalSimpanan(getSettingRAT(tahun + 1) ? tahun + 1 : tahun);
  const profil = getProfilKoperasi();
  const nr = hitungNeraca(tahun);
  const isian = {};
  KOLOM_RAT.forEach(function(k) {
    const v = setting[k];
    isian[k] = v === undefined || v === null || v === '' ? (BAWAAN_RAT[k] !== undefined ? BAWAAN_RAT[k] : '') : v;
  });
  if (!isian.tempat_rat) isian.tempat_rat = profil.instansi || '';
  isian.tanggal_rat = tglISO_(setting.tanggal_rat);
  isian.tanggal_undangan_rat = tglISO_(isian.tanggal_undangan_rat);
  isian.pinjaman_maksimal = Number(isian.pinjaman_maksimal) || BAWAAN_RAT.pinjaman_maksimal;
  isian.tenor_maksimal = Number(isian.tenor_maksimal) || BAWAAN_RAT.tenor_maksimal;

  return {
    tahun: tahun,
    profil: profil,
    isian: isian,
    anggota: mutasiAnggota_(tahun),
    neraca: nr,
    analisis: analisisKeuangan_(nr),
    nominal: getNominalSimpanan(tahun),
    rencana: {
      tahun: tahun + 1,
      nominal: nominalRencana,
      jasaSukarela: Number(rencana.jasa_sukarela) || 0,
      metodeJasa: metodeJasaSukarela(rencana),
      jasaPinjaman: Number(rencana.jasa_pinjaman) || 0,
      shuSimpanan: Number(rencana.shu_simpanan) || 0,
      shuJasa: Number(rencana.shu_jasa) || 0
    }
  };
}

function tglISO_(v) {
  if (!v) return '';
  const d = v instanceof Date ? v : new Date(v);
  if (isNaN(d.getTime())) return '';
  const p = function(n) { return (n < 10 ? '0' : '') + n; };
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
}

/**
 * Keanggotaan satu tahun buku: masuk dari tanggal_masuk, keluar dari
 * catatan log "status → keluar". Anggota berstatus keluar tanpa catatan
 * tanggal dianggap keluar sebelum tahun tersebut.
 */
function mutasiAnggota_(tahun) {
  const akhirTahun = new Date(tahun, 11, 31, 23, 59, 59);
  const awalTahun = new Date(tahun, 0, 1);
  const tglKeluar = {};
  sheetToObjects(SHEET.LOG).forEach(function(l) {
    if (String(l.sheet_target) !== SHEET.ANGGOTA || String(l.aksi) !== 'EDIT') return;
    let baru = {};
    try { baru = JSON.parse(l.data_baru || '{}'); } catch (e) { return; }
    if (baru.status === 'keluar') tglKeluar[String(l.id_referensi)] = new Date(l.timestamp);
  });
  let akhir = 0, masuk = 0, keluar = 0;
  sheetToObjects(SHEET.ANGGOTA).forEach(function(a) {
    const id = String(a.id_anggota);
    const tMasuk = a.tanggal_masuk ? new Date(a.tanggal_masuk) : null;
    if (tMasuk && !isNaN(tMasuk.getTime()) && tMasuk > akhirTahun) return;   // baru masuk setelahnya
    let tKeluar = null;
    if (a.status === 'keluar') tKeluar = tglKeluar[id] || new Date(0);
    if (tMasuk && tMasuk >= awalTahun && tMasuk <= akhirTahun) masuk++;
    if (tKeluar && tKeluar >= awalTahun && tKeluar <= akhirTahun) keluar++;
    if (!tKeluar || tKeluar > akhirTahun) akhir++;
  });
  return { awal: akhir - masuk + keluar, masuk: masuk, keluar: keluar, akhir: akhir };
}

/** Likuiditas, solvabilitas, rentabilitas — seperti laporan badan pengawas. */
function analisisKeuangan_(nr) {
  const jml = function(a) { return a.reduce(function(s, r) { return s + (Number(r.nilai) || 0); }, 0); };
  const aktivaLancar = jml(nr.aktivaLancar);
  const aktivaTetap = jml(nr.aktivaTetap);
  const utangPendek = jml(nr.hutangLancar) + jml(nr.danaDana);
  const utangPanjang = 0;
  const modalSendiri = jml(nr.modalSendiri);
  const persen = function(a, b) { return b ? Math.round(a / b * 10000) / 100 : 0; };
  return {
    aktivaLancar: aktivaLancar, aktivaTetap: aktivaTetap, totalAktiva: nr.totalAktiva,
    utangPendek: utangPendek, utangPanjang: utangPanjang, modalSendiri: modalSendiri,
    shu: nr.shuBerjalan,
    likuiditas: persen(aktivaLancar, utangPendek),
    solvabilitas: persen(nr.totalAktiva, utangPendek + utangPanjang),
    rentabilitas: persen(nr.shuBerjalan, modalSendiri)
  };
}

// ============================================================
// BUAT FILE
// ============================================================

/**
 * Simpan isian RAT ke setting tahun buku, susun .docx, simpan ke Google Drive
 * (folder "Laporan RAT Koperasi"), kembalikan juga isi file untuk diunduh.
 */
function apiBuatLaporanRAT(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  d = d || {};
  const tahun = Number(d.tahun);
  if (!getSettingRAT(tahun)) throw new Error('Tahun buku ' + tahun + ' belum ada di Pengaturan.');
  if (!d.tanggal_rat) throw new Error('Tanggal RAT wajib diisi.');

  // simpan isian supaya tahun depan tinggal menyesuaikan
  tambahKolomBila_(SHEET.SETTING_RAT, KOLOM_RAT);
  const simpan = { tanggal_rat: new Date(d.tanggal_rat + 'T00:00:00') };
  KOLOM_RAT.forEach(function(k) { if (d[k] !== undefined) simpan[k] = String(d[k]).trim(); });
  if (d.tanggal_undangan_rat) simpan.tanggal_undangan_rat = new Date(d.tanggal_undangan_rat + 'T00:00:00');
  simpan.pinjaman_maksimal = Number(d.pinjaman_maksimal) || BAWAAN_RAT.pinjaman_maksimal;
  simpan.tenor_maksimal = Number(d.tenor_maksimal) || BAWAAN_RAT.tenor_maksimal;
  updateRowByField(SHEET.SETTING_RAT, 'tahun', tahun, simpan);
  hapusCache(SHEET.SETTING_RAT);

  const data = kumpulkanDataRAT_(tahun);
  // angka keanggotaan boleh dikoreksi manual di form
  ['awal', 'masuk', 'keluar', 'akhir'].forEach(function(k) {
    const v = d['anggota_' + k];
    if (v !== undefined && v !== '' && !isNaN(Number(v))) data.anggota[k] = Number(v);
  });

  const blob = susunDocxRAT_(data);
  const namaFile = 'Laporan RAT Tahun Buku ' + tahun + '.docx';
  blob.setName(namaFile);

  let url = '';
  try {
    const cari = DriveApp.getFoldersByName(FOLDER_RAT);
    const folder = cari.hasNext() ? cari.next() : DriveApp.createFolder(FOLDER_RAT);
    const lama = folder.getFilesByName(namaFile);
    while (lama.hasNext()) lama.next().setTrashed(true);   // ganti versi sebelumnya
    url = folder.createFile(blob).getUrl();
  } catch (e) {
    url = '';   // Drive gagal → file tetap bisa diunduh langsung
  }
  logAktivitas('INSERT', SHEET.SETTING_RAT, String(tahun), null, { laporan_rat: namaFile });
  return {
    nama: namaFile, url: url,
    base64: Utilities.base64Encode(blob.getBytes()),
    ringkas: { shu: data.analisis.shu, anggota: data.anggota.akhir,
      likuiditas: data.analisis.likuiditas, solvabilitas: data.analisis.solvabilitas,
      rentabilitas: data.analisis.rentabilitas, balance: data.neraca.balance }
  };
}

// ============================================================
// FORMAT ANGKA & TANGGAL
// ============================================================

const BULAN_RAT = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus',
  'September', 'Oktober', 'November', 'Desember'];
const HARI_RAT = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', "Jum'at", 'Sabtu'];

function tglPanjang_(iso, pakaiHari) {
  if (!iso) return '';
  const d = new Date(iso + 'T00:00:00');
  if (isNaN(d.getTime())) return '';
  return (pakaiHari ? HARI_RAT[d.getDay()] + ', ' : '') + d.getDate() + ' ' + BULAN_RAT[d.getMonth()] + ' ' + d.getFullYear();
}

/** 1234567.5 → "1.234.567,50"; bilangan bulat tanpa sen. */
function angkaRAT_(n) {
  n = Number(n) || 0;
  const neg = n < 0; n = Math.abs(n);
  const bulat = Math.floor(n + 1e-9);
  const sen = Math.round((n - bulat) * 100);
  const s = String(bulat).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return (neg ? '-' : '') + s + (sen ? ',' + (sen < 10 ? '0' : '') + sen : '');
}
function rpRAT_(n) { return 'Rp ' + angkaRAT_(n); }
function persenRAT_(n) { return String(Math.round(n * 100) / 100).replace('.', ',') + ' %'; }

/** Terbilang bahasa Indonesia, mis. 1250000 → "satu juta dua ratus lima puluh ribu rupiah". */
function terbilang(n) {
  const satuan = ['', 'satu', 'dua', 'tiga', 'empat', 'lima', 'enam', 'tujuh', 'delapan', 'sembilan',
    'sepuluh', 'sebelas'];
  const t = function(x) {
    if (x < 12) return satuan[x];
    if (x < 20) return satuan[x - 10] + ' belas';
    if (x < 100) return satuan[Math.floor(x / 10)] + ' puluh' + (x % 10 ? ' ' + satuan[x % 10] : '');
    if (x < 200) return 'seratus' + (x - 100 ? ' ' + t(x - 100) : '');
    if (x < 1000) return satuan[Math.floor(x / 100)] + ' ratus' + (x % 100 ? ' ' + t(x % 100) : '');
    if (x < 2000) return 'seribu' + (x - 1000 ? ' ' + t(x - 1000) : '');
    if (x < 1e6) return t(Math.floor(x / 1000)) + ' ribu' + (x % 1000 ? ' ' + t(x % 1000) : '');
    if (x < 1e9) return t(Math.floor(x / 1e6)) + ' juta' + (x % 1e6 ? ' ' + t(x % 1e6) : '');
    if (x < 1e12) return t(Math.floor(x / 1e9)) + ' miliar' + (x % 1e9 ? ' ' + t(x % 1e9) : '');
    return t(Math.floor(x / 1e12)) + ' triliun' + (x % 1e12 ? ' ' + t(x % 1e12) : '');
  };
  n = Math.abs(Number(n) || 0);
  const bulat = Math.floor(n + 1e-9);
  const sen = Math.round((n - bulat) * 100);
  let s = bulat === 0 ? 'nol' : t(bulat);
  s += ' rupiah';
  if (sen) s += ' ' + t(sen) + ' sen';
  return s.replace(/\s+/g, ' ').trim();
}

// ============================================================
// PENYUSUN WORDPROCESSINGML
// ============================================================

const F4_LEBAR = 12240, F4_TINGGI = 18720;          // 8,5 × 13 inci (F4), satuan twip
const TEPI = { atas: 1134, bawah: 1134, kiri: 1701, kanan: 1134 };   // 2 cm; kiri 3 cm
const LEBAR_ISI = F4_LEBAR - TEPI.kiri - TEPI.kanan; // 9405 twip

function xmlEsc_(s) {
  return String(s === undefined || s === null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Potongan XML Word (hasil R_/P_) vs teks biasa dari pengguna. */
function apakahXml_(s) { return typeof s === 'string' && s.indexOf('<w:') === 0; }

/** Satu run teks. o: {b, i, u, sz (pt), caps} */
function R_(teks, o) {
  o = o || {};
  let pr = '';
  if (o.b) pr += '<w:b/>';
  if (o.i) pr += '<w:i/>';
  if (o.caps) pr += '<w:caps/>';
  if (o.u) pr += '<w:u w:val="single"/>';
  if (o.sz) pr += '<w:sz w:val="' + Math.round(o.sz * 2) + '"/><w:szCs w:val="' + Math.round(o.sz * 2) + '"/>';
  const bagian = String(teks === undefined || teks === null ? '' : teks).split('\t');
  return bagian.map(function(t, i) {
    return (i ? '<w:r>' + (pr ? '<w:rPr>' + pr + '</w:rPr>' : '') + '<w:tab/></w:r>' : '') +
      (t ? '<w:r>' + (pr ? '<w:rPr>' + pr + '</w:rPr>' : '') + '<w:t xml:space="preserve">' + xmlEsc_(t) + '</w:t></w:r>' : '');
  }).join('');
}

/**
 * Paragraf. isi: string (teks biasa) atau hasil R_() digabung.
 * o: {al: 'center'|'right'|'both', sb, sa (pt spasi sebelum/sesudah), ind (twip kiri), hang (twip gantung),
 *     tabs: [{pos, al, leader}], b, sz, pb (halaman baru), kn (tetap bersama berikutnya), garis (garis bawah)}
 */
function P_(isi, o) {
  o = o || {};
  if (!apakahXml_(isi)) isi = R_(isi, { b: o.b, i: o.i, sz: o.sz, u: o.u });
  let pr = '';
  if (o.kn) pr += '<w:keepNext/>';
  if (o.pb) pr += '<w:pageBreakBefore/>';
  if (o.garis) pr += '<w:pBdr><w:bottom w:val="' + (o.garis === 'ganda' ? 'double' : 'single') + '" w:sz="' +
    (o.garis === 'ganda' ? 6 : 8) + '" w:space="4" w:color="000000"/></w:pBdr>';
  if (o.tabs) pr += '<w:tabs>' + o.tabs.map(function(t) {
    return '<w:tab w:val="' + (t.al || 'left') + '"' + (t.leader ? ' w:leader="' + t.leader + '"' : '') + ' w:pos="' + t.pos + '"/>';
  }).join('') + '</w:tabs>';
  pr += '<w:spacing w:before="' + Math.round((o.sb || 0) * 20) + '" w:after="' + Math.round((o.sa === undefined ? 6 : o.sa) * 20) + '"' +
    (o.rapat ? ' w:line="240" w:lineRule="auto"' : '') + '/>';
  if (o.ind || o.hang) pr += '<w:ind w:left="' + ((o.ind || 0) + (o.hang || 0)) + '"' + (o.hang ? ' w:hanging="' + o.hang + '"' : '') + '/>';
  if (o.al) pr += '<w:jc w:val="' + o.al + '"/>';
  return '<w:p><w:pPr>' + pr + '</w:pPr>' + isi + '</w:p>';
}

/** Item bernomor dengan indentasi gantung: "1.  teks yang panjang ...". */
function NO_(nomor, teks, o) {
  o = o || {};
  const ind = o.ind || 0, lebarNo = o.lebarNo || 425;
  return P_(R_(nomor + '\t', { b: o.b }) + (apakahXml_(teks) ? teks : R_(teks, { b: o.b })),
    { ind: ind, hang: lebarNo, tabs: [{ pos: ind + lebarNo }], sa: o.sa === undefined ? 3 : o.sa, al: o.al || 'both' });
}

/** Tabel. baris: array baris, tiap sel {isi: xml paragraf, lebar, al, b, span, atas (garis atas), bawah}. */
function T_(lebar, baris, o) {
  o = o || {};
  const total = lebar.reduce(function(a, b) { return a + b; }, 0);
  const bingkai = o.garis
    ? '<w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/><w:left w:val="single" w:sz="4" w:color="000000"/>' +
      '<w:bottom w:val="single" w:sz="4" w:color="000000"/><w:right w:val="single" w:sz="4" w:color="000000"/>' +
      '<w:insideH w:val="single" w:sz="4" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:color="000000"/></w:tblBorders>'
    : '<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>' +
      '<w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>';
  let x = '<w:tbl><w:tblPr><w:tblW w:w="' + total + '" w:type="dxa"/>' +
    (o.ind ? '<w:tblInd w:w="' + o.ind + '" w:type="dxa"/>' : '') + bingkai +
    '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="' + (o.pad === undefined ? 60 : o.pad) + '" w:type="dxa"/>' +
    '<w:right w:w="' + (o.pad === undefined ? 60 : o.pad) + '" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>' +
    lebar.map(function(w) { return '<w:gridCol w:w="' + w + '"/>'; }).join('') + '</w:tblGrid>';
  baris.forEach(function(row) {
    x += '<w:tr><w:trPr><w:cantSplit/></w:trPr>';
    let kol = 0;
    row.forEach(function(sel) {
      if (typeof sel === 'string' || typeof sel === 'number') sel = { t: String(sel) };
      const span = sel.span || 1;
      let w = 0; for (let i = 0; i < span; i++) w += lebar[kol + i];
      kol += span;
      let tcPr = '<w:tcW w:w="' + w + '" w:type="dxa"/>' + (span > 1 ? '<w:gridSpan w:val="' + span + '"/>' : '');
      if (sel.atas || sel.bawah) tcPr += '<w:tcBorders>' + (sel.atas ? '<w:top w:val="single" w:sz="8" w:color="000000"/>' : '') +
        (sel.bawah ? '<w:bottom w:val="single" w:sz="8" w:color="000000"/>' : '') + '</w:tcBorders>';
      if (sel.vmid) tcPr += '<w:vAlign w:val="center"/>';
      if (sel.arsir) tcPr += '<w:shd w:val="clear" w:color="auto" w:fill="E7ECE9"/>';
      let isi = sel.isi || P_(R_(sel.t || '', { b: sel.b, i: sel.i }), { al: sel.al, sa: sel.sa === undefined ? 1 : sel.sa, sb: sel.sb || 1 });
      if (o.kn) isi = isi.replace(/<w:pPr>/g, '<w:pPr><w:keepNext/>');
      x += '<w:tc><w:tcPr>' + tcPr + '</w:tcPr>' + isi + '</w:tc>';
    });
    x += '</w:tr>';
  });
  return x + '</w:tbl>' + (o.tanpaSpasi ? '' : P_('', { sa: 2, kn: o.kn }));
}

/** Blok tanda tangan: satu atau dua kolom. kolom: [{atas, jabatan, nama}] */
function TTD_(kolom, o) {
  o = o || {};
  const lebar = kolom.length === 1 ? [LEBAR_ISI - 4200, 4200] : [LEBAR_ISI / 2, LEBAR_ISI / 2];
  const sel = function(k) {
    return { isi: P_(R_(k.atas || ''), { al: 'center', sa: 0 }) + P_(R_(k.jabatan || ''), { al: 'center', sa: 0 }) +
      P_('', { sa: 0 }) + P_('', { sa: 0 }) + P_('', { sa: 0 }) +
      P_(R_(k.nama || '(................................................)', { b: true, u: !!k.nama }), { al: 'center', sa: 0 }) };
  };
  const baris = kolom.length === 1 ? [[{ t: '' }, sel(kolom[0])]] : [kolom.map(sel)];
  // blok tanda tangan tidak boleh terpisah dari baris di atasnya
  return P_('', { sa: o.sb === undefined ? 6 : o.sb, kn: true }) + T_(lebar, baris, { kn: true });
}

/** Judul bagian (halaman baru). sub: string atau array baris. */
function JUDUL_(teks, sub, o) {
  o = o || {};
  const baris = [].concat(sub || []);
  return P_(R_(teks, { b: true, sz: 13 }), { al: 'center', sa: 0, pb: o.pb, kn: true }) +
    baris.map(function(t, i) {
      return P_(R_(t, { b: true, sz: 12 }), { al: 'center', sa: i === baris.length - 1 ? 12 : 0, kn: true });
    }).join('') + (baris.length ? '' : P_('', { sa: 6 }));
}
function SUBJ_(teks, o) {
  o = o || {};
  return P_(R_(teks, { b: true }), { sb: o.sb === undefined ? 8 : o.sb, sa: 4, kn: true });
}

/** Pecahan untuk analisis rasio: [label pembilang / penyebut] × 100% = hasil */
function PECAHAN_(atas, bawah, hasil) {
  const lb = [3700, 1000, 3000];
  return T_(lb, [
    [{ t: atas, al: 'center', sa: 0 }, { t: '× 100 %', al: 'center', sa: 0 }, { t: '= ' + hasil, b: true, sa: 0 }],
    [{ t: bawah, al: 'center', atas: true, sa: 0 }, { t: '' }, { t: '' }]
  ], { ind: 567 });
}

// ============================================================
// ISI LAPORAN
// ============================================================

function susunDocxRAT_(D) {
  const T = D.tahun, P = D.profil, I = D.isian, A = D.anggota, N = D.neraca, K = D.analisis;
  const kota = P.kota || '..........';
  const instansi = P.instansi || 'sekolah/instansi';
  const nama = 'Koperasi ' + P.nama_koperasi;
  const tglRAT = tglPanjang_(I.tanggal_rat, true);
  const bulanRAT = (function() {
    const d = new Date(I.tanggal_rat + 'T00:00:00');
    return isNaN(d.getTime()) ? '' : BULAN_RAT[d.getMonth()] + ' ' + d.getFullYear();
  })();
  const tglTtd = kota + ', ' + (tglPanjang_(I.tanggal_rat) || bulanRAT);
  const pengurusTTD = [
    { jabatan: 'Ketua', nama: P.ketua }, { jabatan: 'Sekretaris', nama: P.sekretaris }];
  let x = '';

  // ---------- 1. UNDANGAN ----------
  x += kopSurat_(P);
  x += T_([1300, 300, 3900, 3900], [
    ['Nomor', ':', I.nomor_undangan_rat || '......./KOP/...../' + (T + 1), { t: kota + ', ' + (tglPanjang_(I.tanggal_undangan_rat) || '.................'), al: 'right' }],
    ['Lampiran', ':', '-', ''],
    ['Hal', ':', { t: 'Undangan Rapat Anggota Tahunan', b: true }, '']
  ]);
  x += P_('Kepada Yth.', { ind: 5100, sa: 0, sb: 6 });
  x += P_('Bapak/Ibu Pembina dan Anggota', { ind: 5100, sa: 0 });
  x += P_(nama, { ind: 5100, sa: 0 });
  x += P_('di Tempat', { ind: 5100, sa: 14 });
  x += P_(R_('Assalamu’alaikum Wr. Wb.', { i: true }), { sa: 6 });
  x += P_('Sehubungan dengan akan dilaksanakannya Rapat Anggota Tahunan (RAT) ' + nama + ' Tahun Buku ' + T +
    ', kami pengurus koperasi mengundang Bapak/Ibu untuk hadir pada:', { al: 'both' });
  x += T_([2300, 300, 6200], [
    ['Hari / Tanggal', ':', tglRAT || '.......................'],
    ['Waktu', ':', I.waktu_rat],
    ['Tempat', ':', I.tempat_rat],
    ['Acara', ':', { isi: P_('1.  Laporan Pertanggungjawaban Pengurus Tahun Buku ' + T, { sa: 0, sb: 1 }) +
      (I.acara_tambahan_rat ? P_('2.  ' + I.acara_tambahan_rat, { sa: 1 }) : '') }]
  ], { ind: 567 });
  x += P_('Mengingat pentingnya acara tersebut, kami mengharapkan kehadiran Bapak/Ibu tepat pada waktunya. ' +
    'Atas perhatian dan kehadirannya kami ucapkan terima kasih.', { al: 'both' });
  x += P_(R_('Wassalamu’alaikum Wr. Wb.', { i: true }), { sa: 4 });
  x += P_('Pengurus ' + nama, { al: 'center', sb: 8, sa: 0 });
  x += TTD_(pengurusTTD, { sb: 0 });

  // ---------- 2. SUSUNAN ACARA ----------
  x += JUDUL_('SUSUNAN ACARA RAPAT ANGGOTA TAHUNAN (RAT)', 'TAHUN BUKU ' + T, { pb: true });
  [
    'Pembukaan', 'Pengesahan Acara', 'Sambutan Ketua ' + nama,
    'Sambutan Kepala ' + instansi, 'Pembacaan Tata Tertib RAT',
    'Laporan Pertanggungjawaban Pengurus ' + nama, 'Laporan Badan Pengawas ' + nama,
    'Pandangan Umum', 'Pengesahan Pertanggungjawaban Pengurus Tahun Buku ' + T,
    'Program Kerja dan RAB Tahun ' + (T + 1), 'Penutup'
  ].forEach(function(a, i) { x += NO_((i + 1) + '.', a, { ind: 567, sa: 6 }); });
  x += P_(tglTtd, { al: 'right', sb: 12, sa: 0 });
  x += TTD_(pengurusTTD, { sb: 0 });

  // ---------- 3. TATA TERTIB ----------
  const kuorum = Math.floor(A.akhir * 3 / 4) + 1;
  x += JUDUL_('TATA TERTIB RAPAT ANGGOTA TAHUNAN (RAT)', 'TAHUN BUKU ' + T, { pb: true });
  const tt = [
    'Rapat ini adalah Rapat Anggota Tahunan ' + nama + (P.instansi ? ' ' + P.instansi : '') + ' Tahun Buku ' + T + '.',
    'Pimpinan Rapat adalah Pengurus ' + nama + '.',
    ['Peserta Rapat terdiri dari:', ['Semua Pengurus;', 'Ketua dan Anggota Badan Pengawas;',
      'Semua anggota yang pada tanggal 31 Desember ' + T + ' masih aktif menjadi anggota ' + nama + '.']],
    'Hak bicara bagi anggota diatur berdasarkan kebijaksanaan Pimpinan Rapat.',
    'Rapat Anggota Tahunan ini sah apabila dihadiri oleh lebih dari ¾ jumlah anggota, yaitu sekurang-kurangnya ' +
      kuorum + ' orang dari ' + A.akhir + ' anggota.',
    'Keputusan dianggap sah apabila disetujui oleh sekurang-kurangnya ¾ dari anggota yang hadir.',
    'Setiap peserta wajib menjaga ketertiban selama RAT berlangsung.',
    'Dalam menggunakan hak suara, anggota harus obyektif, ringkas, dan jelas.',
    ['Pimpinan Rapat berhak:', ['Mengambil kebijaksanaan untuk mengatasi segala sesuatu yang dapat mengganggu kelancaran jalannya rapat;',
      'Mengambil tindakan seperlunya terhadap peserta yang tidak mengindahkan tata tertib, berupa teguran dan peringatan.']],
    'Hal-hal yang dianggap perlu namun belum tercantum dalam tata tertib ini akan diatur dan ditentukan kemudian berdasarkan kesepakatan peserta rapat.'
  ];
  tt.forEach(function(item, i) {
    if (typeof item === 'string') { x += NO_((i + 1) + '.', item, { sa: 5 }); return; }
    x += NO_((i + 1) + '.', item[0], { sa: 3 });
    item[1].forEach(function(s, j) { x += NO_(String.fromCharCode(97 + j) + '.', s, { ind: 425, sa: 3 }); });
  });
  x += P_(tglTtd, { al: 'right', sb: 12, sa: 0 });
  x += TTD_(pengurusTTD, { sb: 0 });

  // ---------- 4. LAPORAN PENGURUS ----------
  x += JUDUL_('LAPORAN PERTANGGUNGJAWABAN PENGURUS', [nama.toUpperCase(), 'TAHUN BUKU ' + T], { pb: true });
  x += P_(R_('Assalamu’alaikum Wr. Wb.', { i: true }), { sa: 4 });
  x += P_('Yang terhormat:', { sa: 0 });
  x += NO_('•', 'Kepala ' + instansi + ';', { sa: 0 });
  x += NO_('•', 'Saudara-saudara Pengurus, Badan Pengawas, para Anggota, dan Undangan.', { sa: 6 });
  x += P_('Puji syukur kita panjatkan kepada Allah SWT, karena berkat inayah-Nya pada hari ini kita dapat hadir bersama dalam ' +
    'Rapat Anggota Tahunan ' + nama + '. Shalawat beserta salam semoga selalu tercurah kepada Nabi Muhammad SAW, keluarga, ' +
    'sahabat, serta para pengikutnya hingga akhir zaman.', { al: 'both' });
  x += P_('Kami mengucapkan terima kasih kepada Bapak, Ibu, dan Saudara atas kehadirannya memenuhi undangan kami. ' +
    'Laporan ini kami sajikan dengan sistematika sebagai berikut.', { al: 'both' });

  x += SUBJ_('I.  PENDAHULUAN');
  x += P_('Rapat Anggota Tahunan (RAT) adalah forum tertinggi pengambilan keputusan dalam koperasi, sehingga wajib dilaksanakan ' +
    'setiap tahun. Laporan pertanggungjawaban ini disampaikan kepada anggota melalui forum RAT, terutama dititikberatkan pada ' +
    'kegiatan yang telah diprogramkan pada RAT tahun lalu.', { al: 'both', ind: 425 });
  x += P_('Susunan pengurus dan badan pengawas:', { ind: 425, sa: 2 });
  // satu nama per baris (koma tidak dipakai sebagai pemisah karena ada di gelar, mis. "S.Ag, M.Pd")
  const perBaris = function(v) { return String(v || '').split(/\r?\n|;/).map(function(s) { return s.trim(); }).filter(String); };
  const anggotaPengurus = perBaris(P.pengurus_anggota);
  const pengawasAnggota = perBaris(P.pengawas_anggota);
  const susunan = [['Ketua', P.ketua], ['Sekretaris', P.sekretaris], ['Bendahara', P.bendahara]];
  anggotaPengurus.forEach(function(n, i) { susunan.push([i ? '' : 'Anggota', n]); });
  susunan.push(['Badan Pengawas', (P.pengawas_ketua ? P.pengawas_ketua + ' (Ketua)' : '')]);
  pengawasAnggota.forEach(function(n) { susunan.push(['', n]); });
  x += T_([2600, 300, 5900], susunan.filter(function(s) { return s[1]; }).map(function(s) {
    return [s[0], s[0] ? ':' : '', s[1]];
  }), { ind: 425 });
  x += P_('Selama satu tahun kami berusaha melaksanakan program yang telah ditetapkan, namun masih ada sebagian program yang belum ' +
    'dapat dilaksanakan mengingat berbagai keterbatasan.', { al: 'both', ind: 425 });

  x += SUBJ_('II.  PELAKSANAAN DAN EVALUASI PROGRAM');
  x += P_('Kegiatan yang telah kami lakukan adalah sebagai berikut.', { ind: 425 });
  const bidang = [
    ['Bidang Administrasi dan Organisasi', 'Menyempurnakan kelengkapan administrasi dan melaksanakan konsolidasi organisasi dalam bentuk mekanisme kerja. Pencatatan transaksi dilakukan melalui aplikasi koperasi.'],
    ['Bidang Usaha', 'Melanjutkan usaha simpan pinjam bagi anggota.'],
    ['Bidang Kesejahteraan', 'Berjalan sesuai dengan harapan.'],
    ['Bidang Permodalan', 'Simpanan pokok dan wajib berjalan sesuai harapan. Simpanan sukarela terus digalakkan. Sisa Hasil Usaha (SHU) tahun ' + T +
      ' sebesar ' + rpRAT_(K.shu) + ' (' + kapital_(terbilang(K.shu)) + '). Neraca terlampir untuk dapat ditelaah para anggota.'],
    ['Bidang Sarana', 'Tempat dan sarana kegiatan koperasi memanfaatkan fasilitas sekolah.'],
    ['Keanggotaan', 'Jumlah anggota per 31 Desember ' + T + ' sebanyak ' + A.akhir + ' orang.']
  ];
  bidang.forEach(function(b, i) {
    x += NO_(String.fromCharCode(97 + i) + '.', R_(b[0], { b: true }), { ind: 425, sa: 1 });
    x += P_(b[1], { ind: 850, al: 'both', sa: 5 });
  });

  x += SUBJ_('III.  PENUTUP');
  x += P_((I.predikat_kesehatan ? 'Berdasarkan penilaian kesehatan koperasi oleh Dinas Koperasi, UKM, Perindustrian dan Perdagangan ' +
    (P.kota ? 'Kota ' + P.kota : 'setempat') + ', Alhamdulillah ' + nama + ' mendapat predikat “' + I.predikat_kesehatan + '”. Namun demikian, ' : '') +
    'mengingat keterbatasan kemampuan dan waktu, tentu masih terdapat kekurangan dalam kegiatan koperasi selama kepengurusan kami. ' +
    'Dengan segala kerendahan hati kami mohon maaf, dan semoga laporan pertanggungjawaban ini dapat diterima dan disahkan.',
    { al: 'both', ind: 425 });
  x += P_('Akhirnya kami menghaturkan terima kasih kepada semua pihak yang telah membantu jalannya koperasi kita.', { al: 'both', ind: 425 });
  x += P_(R_('Wabillahi taufik wal hidayah. Wassalamu’alaikum Wr. Wb.', { i: true }), { ind: 425, kn: true });
  x += TTD_([{ atas: tglTtd, jabatan: 'Ketua', nama: P.ketua }]);

  // ---------- 5. LAPORAN BADAN PENGAWAS ----------
  x += JUDUL_('LAPORAN BADAN PENGAWAS', 'PADA RAT TAHUN BUKU ' + T, { pb: true });
  x += P_(R_('Assalamu’alaikum Wr. Wb.', { i: true }), { sa: 4 });
  x += P_('Puji syukur kehadirat Allah SWT yang telah memberikan rahmat dan hidayah-Nya sehingga kita dapat bersama-sama menghadiri ' +
    'Rapat Anggota Tahunan Tahun Buku ' + T + '. Berdasarkan keputusan RAT tahun lalu, Badan Pengawas mendapat mandat dari anggota ' +
    'untuk melaksanakan pengawasan terhadap jalannya koperasi. Hasil pemeriksaan tahun buku ' + T + ' adalah sebagai berikut.', { al: 'both' });

  x += SUBJ_('I.  BIDANG ORGANISASI');
  x += P_('Tegak dan majunya koperasi ditopang oleh organisasi/kepengurusan yang sehat, administrasi yang sehat, dan usaha yang sehat. ' +
    'Keadaan keanggotaan:', { al: 'both', ind: 425 });
  x += T_([5200, 300, 1500], [
    ['Jumlah anggota awal tahun ' + T, ':', { t: A.awal + ' orang', al: 'right' }],
    ['Anggota masuk tahun ' + T, ':', { t: A.masuk + ' orang', al: 'right' }],
    ['Anggota keluar tahun ' + T, ':', { t: A.keluar + ' orang', al: 'right' }],
    [{ t: 'Jumlah anggota akhir tahun ' + T, b: true }, ':', { t: A.akhir + ' orang', b: true, al: 'right', atas: true }]
  ], { ind: 425 });
  x += P_('Pada dasarnya kepengurusan berjalan dengan baik.', { ind: 425 });

  x += SUBJ_('II.  BIDANG ADMINISTRASI');
  x += NO_('a.', R_('Administrasi Organisasi', { b: true }), { ind: 425, sa: 1 });
  x += P_('Buku daftar anggota, daftar pengurus, daftar badan pengawas, notulen rapat, inventaris, agenda, catatan kejadian penting, ' +
    'dan catatan lainnya. Pada dasarnya buku-buku tersebut telah sesuai dengan kenyataan dan beberapa perlu dilengkapi.', { ind: 850, al: 'both' });
  x += NO_('b.', R_('Administrasi Keuangan', { b: true }), { ind: 425, sa: 1 });
  x += P_('Buku kas umum, daftar simpanan kolektif dan perorangan, daftar pinjaman, daftar potongan bulanan, catatan pendukung, ' +
    'dan bundel bukti keuangan. Penyelenggaraan administrasi keuangan berdasarkan bukti-bukti yang sah dan dibukukan dengan tertib.', { ind: 850, al: 'both' });

  x += SUBJ_('III.  BIDANG KEUANGAN');
  x += NO_('a.', R_('Permodalan', { b: true }), { ind: 425, sa: 1 });
  x += T_([4500, 300, 2200], [
    ['Simpanan Pokok', ':', { t: rpRAT_(D.nominal.nominal_pokok) + ',-', al: 'right' }],
    ['Simpanan Wajib (per bulan)', ':', { t: rpRAT_(D.nominal.nominal_wajib) + ',-', al: 'right' }],
    ['Simpanan Sukarela terendah (per bulan)', ':', { t: rpRAT_(D.nominal.minimal_sukarela) + ',-', al: 'right' }]
  ], { ind: 850 });
  x += NO_('b.', R_('Kondisi aktiva dan pasiva per 31 Desember ' + T, { b: true }), { ind: 425, sa: 3 });
  x += tabelNeraca_(N);

  x += SUBJ_('IV.  ANALISIS KEUANGAN');
  x += NO_('1.', R_('Likuiditas', { b: true }) + R_(' (kemampuan melunasi utang jangka pendek)'), { ind: 425, sa: 2 });
  x += T_([5600, 300, 2200], [
    ['Aktiva lancar (kas & piutang pinjaman)', ':', { t: rpRAT_(K.aktivaLancar), al: 'right' }],
    ['Utang jangka pendek (simpanan sukarela & dana-dana)', ':', { t: rpRAT_(K.utangPendek), al: 'right' }]
  ], { ind: 850 });
  x += PECAHAN_('Aktiva Lancar', 'Utang Jangka Pendek', persenRAT_(K.likuiditas));
  x += P_(R_('Artinya: setiap utang yang harus segera dibayar sebesar Rp 1,- dijamin dengan harta lancar sebesar Rp ' +
    kaliRp_(K.likuiditas) + ' (' + (K.likuiditas >= 100 ? 'likuid' : 'kurang likuid') + ').', { i: true }), { ind: 850, al: 'both' });

  x += NO_('2.', R_('Solvabilitas', { b: true }) + R_(' (kemampuan membayar seluruh utang)'), { ind: 425, sa: 2 });
  x += T_([5600, 300, 2200], [
    ['Jumlah aktiva', ':', { t: rpRAT_(K.totalAktiva), al: 'right' }],
    ['Utang jangka pendek + jangka panjang', ':', { t: rpRAT_(K.utangPendek + K.utangPanjang), al: 'right' }]
  ], { ind: 850 });
  x += PECAHAN_('Jumlah Aktiva', 'Jumlah Utang', persenRAT_(K.solvabilitas));
  x += P_(R_('Artinya: setiap kewajiban koperasi sebesar Rp 1,- dijamin oleh harta sebesar Rp ' +
    kaliRp_(K.solvabilitas) + ' (' + (K.solvabilitas >= 100 ? 'solvabel' : 'kurang solvabel') + ').', { i: true }),
    { ind: 850, al: 'both' });

  x += NO_('3.', R_('Rentabilitas', { b: true }) + R_(' (kemampuan modal sendiri menghasilkan SHU)'), { ind: 425, sa: 2 });
  x += T_([5600, 300, 2200], [
    ['Sisa Hasil Usaha tahun ' + T, ':', { t: rpRAT_(K.shu), al: 'right' }],
    ['Modal sendiri (simpanan pokok, wajib, cadangan)', ':', { t: rpRAT_(K.modalSendiri), al: 'right' }]
  ], { ind: 850 });
  x += PECAHAN_('Sisa Hasil Usaha', 'Modal Sendiri', persenRAT_(K.rentabilitas));
  x += P_(R_('Artinya: setiap modal sendiri sebesar Rp 1,- menghasilkan keuntungan sebesar Rp ' +
    kaliRp_(K.rentabilitas) + '.', { i: true }), { ind: 850, al: 'both' });

  x += SUBJ_('V.  KESIMPULAN DAN SARAN');
  x += P_('Jalannya kegiatan koperasi dapat menyejahterakan anggota. Saran-saran:', { ind: 425 });
  x += NO_('a.', 'Agar simpan pinjam dan simpanan sukarela ditingkatkan;', { ind: 425 });
  x += NO_('b.', 'Agar pembagian tugas dalam kepengurusan ditingkatkan.', { ind: 425 });
  x += SUBJ_('VI.  PENUTUP');
  x += P_('Akhirnya kami mohon maaf apabila terdapat kekurangan dan kekhilafan, semoga laporan ini bermanfaat bagi semua pihak.', { ind: 425, al: 'both', kn: true });
  x += P_(R_('Wassalamu’alaikum Wr. Wb.', { i: true }), { ind: 425, kn: true });
  x += TTD_([{ atas: tglTtd, jabatan: 'Badan Pengawas', nama: P.pengawas_ketua }]);

  // ---------- 6. RENCANA KERJA ----------
  const RK = D.rencana;
  x += JUDUL_('RENCANA KERJA', 'TAHUN ' + RK.tahun, { pb: true });
  x += SUBJ_('A.  BIDANG ORGANISASI DAN ADMINISTRASI', { sb: 2 });
  x += NO_('1.', 'Meningkatkan pengelolaan koperasi dalam bidang administrasi, keanggotaan, serta tata niaga secara keseluruhan.', { ind: 425 });
  x += SUBJ_('B.  BIDANG USAHA');
  x += NO_('1.', 'Meningkatkan pelayanan terhadap anggota sesuai ketentuan koperasi.', { ind: 425 });
  x += NO_('2.', 'Meningkatkan pelayanan jumlah pinjaman bagi anggota sesuai kemampuan, dengan kriteria:', { ind: 425 });
  [
    'Permohonan pinjaman diajukan minimal satu bulan sebelumnya;',
    'Besarnya pinjaman maksimal ' + rpRAT_(I.pinjaman_maksimal) + ',-, dapat melebihi apabila keuangan koperasi tersedia;',
    'Pemenuhan permohonan pinjaman disesuaikan dengan keadaan keuangan koperasi;',
    'Pinjaman diangsur tiap bulan dengan cara pemotongan gaji oleh bendahara gaji ' + instansi + ';',
    'Pinjaman diangsur maksimal ' + I.tenor_maksimal + ' bulan dengan jasa ' + String(RK.jasaPinjaman).replace('.', ',') +
      '% per bulan dari sisa pinjaman (menurun);',
    'Batas pembayaran setiap angsuran tanggal 10 tiap bulan;',
    'Apabila peminjam 3 kali berturut-turut tidak dapat membayar sesuai waktu yang ditentukan, setelah dibina pada bulan ke-1 dan ke-2, ' +
      'yang bersangkutan dinyatakan keluar sebagai anggota ' + nama + ';',
    'Peminjam yang masih memiliki tunggakan tidak diberi pinjaman baru kecuali dalam keadaan darurat;',
    'Pinjaman diprioritaskan bagi anggota yang belum pernah meminjam.'
  ].forEach(function(s, i) { x += NO_(String.fromCharCode(97 + i) + '.', s, { ind: 850 }); });
  x += NO_('3.', 'Melayani kebutuhan barang yang bersifat insidental.', { ind: 425 });
  x += NO_('4.', 'Melakukan promosi kepada anggota apabila dana koperasi belum terserap.', { ind: 425 });

  x += SUBJ_('C.  BIDANG KESEJAHTERAAN');
  [
    'Memberikan bantuan transport sebesar Rp 100.000,- untuk menjenguk anggota yang melahirkan atau dirawat di rumah sakit, serta menghadiri undangan anggota;',
    'Memberikan hadiah Idul Fitri sebesar Rp 350.000,- per anggota;',
    'Memberikan insentif bulanan kepada pengurus sesuai keuangan koperasi;',
    'Memberikan uang duduk bagi anggota yang mengikuti RAT sebesar Rp 100.000,- atau menurut kondisi keuangan koperasi;',
    'Memberikan cendera mata kepada semua anggota pada saat RAT;',
    'Pengambilan simpanan sukarela dapat dilakukan 2 kali dalam satu tahun, dengan mendaftar sebulan sebelumnya;',
    RK.metodeJasa === 'harian'
      ? 'Simpanan sukarela mendapat jasa ' + String(RK.jasaSukarela).replace('.', ',') + '% per tahun, dihitung setiap bulan dari saldo akhir bulan sebelumnya;'
      : 'Setiap bulan simpanan sukarela mendapat jasa ' + String(RK.jasaSukarela).replace('.', ',') +
        '% dari jumlah simpanan sukarela terendah pada bulan tersebut (disesuaikan dengan jasa bank);',
    'Simpanan sukarela tidak diperhitungkan dalam pembagian SHU;',
    'Pembagian SHU: SHU simpanan ' + RK.shuSimpanan + '% (simpanan pokok dan wajib) dan SHU pinjaman ' + RK.shuJasa + '%.'
  ].forEach(function(s, i) { x += NO_((i + 1) + '.', s, { ind: 425 }); });

  x += SUBJ_('D.  BIDANG PERMODALAN');
  x += T_([4500, 300, 2200], [
    ['Simpanan Pokok bagi anggota baru', ':', { t: rpRAT_(RK.nominal.nominal_pokok) + ',-', al: 'right' }],
    ['Simpanan Wajib per bulan', ':', { t: rpRAT_(RK.nominal.nominal_wajib) + ',-', al: 'right' }],
    ['Simpanan Sukarela minimal', ':', { t: rpRAT_(RK.nominal.minimal_sukarela) + ',-', al: 'right' }]
  ], { ind: 425, kn: true });
  x += T_([LEBAR_ISI - 4700, 1800, 300, 2600], [
    ['', 'Ditetapkan di', ':', kota], ['', 'Tanggal', ':', tglPanjang_(I.tanggal_rat) || bulanRAT]
  ], { kn: true });
  x += TTD_([{ jabatan: 'Pimpinan Rapat', nama: '' }], { sb: 0 });

  return bungkusDocx_(x, P);
}

/** 200,22 % → "2,00" (rupiah per Rp 1) */
function kaliRp_(persen) { return String((persen / 100).toFixed(2)).replace('.', ','); }
function kapital_(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }

/** Tabel aktiva & pasiva (neraca ringkas) untuk laporan pengawas. */
function tabelNeraca_(N) {
  const lb = [500, 5300, 2800];
  const baris = [[{ t: 'AKTIVA', b: true, span: 3, arsir: true }]];
  let no = 0;
  N.aktivaLancar.concat(N.aktivaTetap).forEach(function(r) {
    baris.push([{ t: (++no) + '.', al: 'right' }, r.kode === '103' ? 'Usaha Simpan Pinjam (piutang pinjaman)' : r.label,
      { t: rpRAT_(r.nilai), al: 'right' }]);
  });
  baris.push(['', { t: 'JUMLAH AKTIVA', b: true }, { t: rpRAT_(N.totalAktiva), b: true, al: 'right', atas: true }]);
  baris.push([{ t: 'PASIVA', b: true, span: 3, arsir: true }]);
  no = 0;
  const pasiva = N.modalSendiri.concat(N.hutangLancar, N.danaDana);
  pasiva.forEach(function(r) {
    baris.push([{ t: (++no) + '.', al: 'right' }, r.label, { t: rpRAT_(r.nilai), al: 'right' }]);
  });
  baris.push([{ t: (++no) + '.', al: 'right' }, 'SHU tahun ' + N.tahun, { t: rpRAT_(N.shuBerjalan), al: 'right' }]);
  baris.push(['', { t: 'JUMLAH PASIVA', b: true }, { t: rpRAT_(N.totalPasiva), b: true, al: 'right', atas: true }]);
  return T_(lb, baris, { ind: 850 });
}

/** Kop surat dengan logo (bila ada) + garis ganda. */
function kopSurat_(P) {
  const teks = P_(R_(P.jenis_koperasi || '', { b: true, sz: 11 }), { al: 'center', sa: 0 }) +
    P_(R_('“' + (P.nama_koperasi || '') + '”', { b: true, sz: 16 }), { al: 'center', sa: 0 }) +
    (P.no_badan_hukum ? P_(R_('Badan Hukum Nomor ' + P.no_badan_hukum, { sz: 10 }), { al: 'center', sa: 0 }) : '') +
    P_(R_((P.alamat || '') + (P.telepon ? ' · Telp. ' + P.telepon : '') + (P.email ? ' · ' + P.email : ''), { sz: 10 }), { al: 'center', sa: 0 });
  const logo = logoDocx_(P.logo);
  const x = logo
    ? T_([1400, LEBAR_ISI - 2800, 1400], [[{ isi: P_(logo.xml, { al: 'center', sa: 0 }), vmid: true }, { isi: teks, vmid: true }, { t: '' }]],
      { tanpaSpasi: true })
    : teks;
  return x + P_('', { garis: 'ganda', sa: 10 });
}

/** Gambar logo (data URL PNG/JPEG) → drawing inline 2,3 cm. */
function logoDocx_(dataUrl) {
  const m = /^data:image\/(png|jpeg);base64,(.+)$/.exec(String(dataUrl || ''));
  if (!m) return null;
  // tinggi 2,3 cm, lebar mengikuti perbandingan gambar (PNG: ukuran ada di header IHDR)
  const ukuran = ukuranGambar_(m[1], m[2]);
  const cy = Math.round(2.3 * 360000);
  const cx = Math.round(cy * ukuran.lebar / ukuran.tinggi);
  return {
    jenis: m[1], base64: m[2],
    xml: '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + cx + '" cy="' + cy + '"/>' +
      '<wp:docPr id="1" name="Logo"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:nvPicPr><pic:cNvPr id="1" name="logo"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdLogo"/>' +
      '<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
  };
}

/** Lebar × tinggi gambar dari isi base64 (PNG & JPEG); gagal → persegi. */
function ukuranGambar_(jenis, b64) {
  try {
    const b = Utilities.base64Decode(b64).map(function(x) { return x & 255; });
    if (jenis === 'png') {
      return { lebar: (b[16] << 24 | b[17] << 16 | b[18] << 8 | b[19]) >>> 0, tinggi: (b[20] << 24 | b[21] << 16 | b[22] << 8 | b[23]) >>> 0 };
    }
    for (let i = 2; i < b.length - 9;) {            // JPEG: cari penanda SOFn
      if (b[i] !== 0xFF) { i++; continue; }
      const m = b[i + 1], len = b[i + 2] << 8 | b[i + 3];
      if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) {
        return { tinggi: b[i + 5] << 8 | b[i + 6], lebar: b[i + 7] << 8 | b[i + 8] };
      }
      i += 2 + len;
    }
  } catch (e) { /* abaikan */ }
  return { lebar: 1, tinggi: 1 };
}

/** Rangkai bagian-bagian .docx lalu zip. */
function bungkusDocx_(isiBody, P) {
  const logo = logoDocx_(P.logo);
  const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
  const sect = '<w:sectPr><w:footerReference w:type="default" r:id="rIdKaki"/>' +
    '<w:pgSz w:w="' + F4_LEBAR + '" w:h="' + F4_TINGGI + '"/>' +
    '<w:pgMar w:top="' + TEPI.atas + '" w:right="' + TEPI.kanan + '" w:bottom="' + TEPI.bawah + '" w:left="' + TEPI.kiri +
    '" w:header="567" w:footer="567" w:gutter="0"/><w:titlePg/></w:sectPr>';
  const dokumen = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + ns + '><w:body>' +
    isiBody + sect + '</w:body></w:document>';
  const gaya = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr>' +
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/>' +
    '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="id-ID"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>' +
    '<w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>' +
    '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>' +
    '<w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>';
  const kaki = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ' + ns + '>' +
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">Laporan RAT ' +
    xmlEsc_('Koperasi ' + (P.nama_koperasi || '')) + ' — halaman </w:t></w:r><w:r><w:rPr><w:sz w:val="18"/></w:rPr>' +
    '<w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
    '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>2</w:t></w:r>' +
    '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>';
  const jenisTipe = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '</Types>';
  const relAkar = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    '</Relationships>';
  const relDok = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rIdGaya" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '<Relationship Id="rIdKaki" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' +
    (logo ? '<Relationship Id="rIdLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.' +
      (logo.jenis === 'png' ? 'png' : 'jpeg') + '"/>' : '') + '</Relationships>';
  const inti = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Laporan RAT</dc:title><dc:creator>' +
    xmlEsc_('Koperasi ' + (P.nama_koperasi || '')) + '</dc:creator></cp:coreProperties>';

  const b = function(isi, nama) { return Utilities.newBlob(isi, 'application/xml', nama); };
  const berkas = [
    b(jenisTipe, '[Content_Types].xml'), b(relAkar, '_rels/.rels'),
    b(dokumen, 'word/document.xml'), b(gaya, 'word/styles.xml'), b(kaki, 'word/footer1.xml'),
    b(relDok, 'word/_rels/document.xml.rels'), b(inti, 'docProps/core.xml')
  ];
  if (logo) {
    berkas.push(Utilities.newBlob(Utilities.base64Decode(logo.base64), 'image/' + logo.jenis,
      'word/media/logo.' + (logo.jenis === 'png' ? 'png' : 'jpeg')));
  }
  const zip = Utilities.zip(berkas, 'laporan-rat.docx');
  return Utilities.newBlob(zip.getBytes(), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'laporan-rat.docx');
}

// ===== reset_data.js =====
/**
 * ============================================================
 * RESET_DATA.GS — Bersihkan Data Uji Coba, Pertahankan Master
 * ============================================================
 *
 * BEDANYA DENGAN "Hapus Semua & Setup Ulang":
 * - Setup Ulang  = menghapus SEMUA sheet termasuk anggota (total)
 * - Reset Data   = hanya MENGOSONGKAN ISI sheet transaksi & laporan;
 *                  sheet-nya tetap ada (header, dropdown, warna utuh)
 *
 * YANG DIPERTAHANKAN (tidak disentuh sama sekali):
 * - anggota      (semua data anggota aman)
 * - coa_akun     (daftar akun)
 * - setting_rat  (baris tahun dipertahankan — lihat catatan bawah)
 *
 * YANG DIKOSONGKAN (isi dihapus, header tetap):
 * - saldo_awal, transaksi_simpanan, transaksi_pengambilan,
 *   transaksi_pinjaman, transaksi_angsuran, transaksi_kas,
 *   jurnal_umum, rekap_jasa_sukarela, realisasi_shu, log_aktivitas
 * - semua sheet laporan (lap_*, rekap_*, struk_potongan) + DASHBOARD
 *
 * CATATAN setting_rat: baris tahun TIDAK dihapus otomatis (keputusan
 * RAT itu data penting). Tapi setelah tes tutup buku, status_aktif
 * Anda mungkin sudah pindah ke tahun berikutnya — fungsi ini akan
 * MENAWARKAN mengembalikan status_aktif ke tahun yang Anda pilih.
 *
 * CARA PAKAI: Menu 🔧 SETUP → 🧹 Reset Data (Anggota Aman)
 * ============================================================
 */

// Sheet yang isinya dikosongkan (baris 2 ke bawah), header dipertahankan
const SHEET_DIKOSONGKAN = [
  'saldo_awal',
  'transaksi_simpanan',
  'transaksi_pengambilan',
  'transaksi_pinjaman',
  'transaksi_angsuran',
  'transaksi_kas',
  'jurnal_umum',
  'rekap_jasa_sukarela',
  'realisasi_shu',
  'log_aktivitas'
];

// Sheet laporan yang dibersihkan total (isinya memang generate-ulang)
const SHEET_LAPORAN_DIBERSIHKAN = [
  'DASHBOARD', 'lap_shu', 'lap_neraca', 'lap_labarugi',
  'rekap_pinjaman', 'rekap_angsuran', 'rekap_simpanan', 'struk_potongan'
];

/**
 * Fungsi inti reset. Bisa dipanggil dari menu atau test.
 * @param {number|null} tahunAktifKembali  jika diisi, status_aktif di
 *        setting_rat dikembalikan ke tahun ini (yang lain jadi FALSE)
 */
function resetDataTransaksi(tahunAktifKembali) {
  pastikanAksesData();
  const hasil = [];

  // 1. Kosongkan sheet transaksi (baris 2 ke bawah, header aman)
  SHEET_DIKOSONGKAN.forEach(function(nama) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nama);
    if (!sheet) { hasil.push('⏭️ ' + nama + ' tidak ditemukan, dilewati'); return; }

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      hasil.push('🧹 ' + nama + ': ' + (lastRow - 1) + ' baris dihapus');
    } else {
      hasil.push('✔️ ' + nama + ': sudah kosong');
    }
    hapusCache(nama);
  });

  // 2. Bersihkan sheet laporan (generate-ulang kapan saja)
  SHEET_LAPORAN_DIBERSIHKAN.forEach(function(nama) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nama);
    if (sheet) { sheet.clear(); hapusCache(nama); }
  });
  hasil.push('🧹 Semua sheet laporan & DASHBOARD dibersihkan');

  // 3. (Opsional) Kembalikan status_aktif setting_rat
  if (tahunAktifKembali) {
    const semuaSetting = sheetToObjects(SHEET.SETTING_RAT);
    semuaSetting.forEach(function(row) {
      const harusAktif = Number(row.tahun) === Number(tahunAktifKembali);
      if (row.status_aktif !== harusAktif) {
        updateRowByRowNumber(SHEET.SETTING_RAT, row.__row, { status_aktif: harusAktif });
      }
    });
    hasil.push('⚙️ status_aktif dikembalikan ke tahun ' + tahunAktifKembali);
  }

  hasil.push('');
  hasil.push('✅ DIPERTAHANKAN: anggota (' +
    sheetToObjects(SHEET.ANGGOTA).length + ' baris), coa_akun, setting_rat');

  return hasil;
}

// ============================================================
// MENU (dengan konfirmasi)
// ============================================================

function menuResetDataTransaksi() {
  const ui = SpreadsheetApp.getUi();

  const konfirmasi = ui.alert(
    '🧹 RESET DATA (Anggota Aman)',
    'Ini akan MENGHAPUS ISI semua sheet transaksi, jurnal, log, dan laporan — ' +
    'tapi data ANGGOTA, coa_akun, dan setting_rat TETAP AMAN.\n\n' +
    'Cocok untuk membersihkan data uji coba sebelum mulai input data asli.\n\n' +
    'Data transaksi yang terhapus TIDAK BISA dikembalikan. Lanjutkan?',
    ui.ButtonSet.YES_NO
  );
  if (konfirmasi !== ui.Button.YES) return;

  // Tawarkan kembalikan tahun aktif (berguna setelah tes tutup buku)
  const respTahun = ui.prompt(
    'Tahun aktif',
    'Setelah reset, tahun berapa yang jadi tahun aktif di setting_rat?\n' +
    '(Ketik angka tahun, contoh: 2026 — atau kosongkan untuk tidak mengubah)',
    ui.ButtonSet.OK_CANCEL
  );
  if (respTahun.getSelectedButton() !== ui.Button.OK) return;

  const teksTahun = respTahun.getResponseText().trim();
  let tahunAktif = null;
  if (teksTahun !== '') {
    tahunAktif = parseInt(teksTahun, 10);
    if (isNaN(tahunAktif) || !getSettingRAT(tahunAktif)) {
      ui.alert('Tahun "' + teksTahun + '" tidak ada di setting_rat. Reset dibatalkan — ' +
        'tidak ada data yang diubah.');
      return;
    }
  }

  const hasil = resetDataTransaksi(tahunAktif);
  ui.alert('✅ RESET SELESAI\n\n' + hasil.join('\n'));
}

// ============================================================
// VERSI EDITOR (untuk testing, tanpa dialog)
// ============================================================

/** Reset dari editor — tahun aktif dikembalikan ke 2026. Ubah jika perlu. */
function test_resetDataTransaksi() {
  MODE_SENYAP = true;
  const hasil = resetDataTransaksi(2026);  // ⚠️ ganti tahun jika perlu
  Logger.log(hasil.join('\n'));
}

// ===== ringkasan_struk.js =====
/**
 * ============================================================
 * RINGKASAN_STRUK.GS — Dashboard per Periode, Riwayat Transaksi,
 *                      dan Data Struk (printer thermal / kuitansi)
 * Koperasi App
 * ============================================================
 *
 * 1. apiDashboardPeriode  — ringkasan kas masuk/keluar, grafik
 *    bulanan, dan komposisi simpanan untuk rentang tanggal.
 * 2. apiRiwayatTransaksi  — gabungan setoran, pengambilan,
 *    angsuran, dan pinjaman (terbaru di atas), bisa difilter.
 * 3. apiStruk / dataStruk — isi bukti transaksi untuk dicetak.
 *
 * Semua fungsi di sini hanya MEMBACA data (tidak dikunci).
 * Tanggal dari klien berformat 'yyyy-MM-dd'.
 * ============================================================
 */

const NAMA_BULAN_PENDEK = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun',
  'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

/** 'yyyy-MM-dd' → Date (awal hari). akhirHari=true → 23:59:59.999 */
function parseTanggalKlien(teks, akhirHari) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(teks || ''));
  if (!m) return null;
  return akhirHari
    ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 23, 59, 59, 999)
    : new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function formatTgl(t, pola) {
  try {
    return Utilities.formatDate(new Date(t), Session.getScriptTimeZone(), pola || 'dd/MM/yyyy');
  } catch (e) {
    return String(t || '');
  }
}

/** Rentang dari filter klien; default 6 bulan terakhir s/d hari ini. */
function rentangDariFilter(f) {
  f = f || {};
  const kini = new Date();
  let dari = parseTanggalKlien(f.dari, false);
  let sampai = parseTanggalKlien(f.sampai, true);
  if (!sampai) sampai = new Date(kini.getFullYear(), kini.getMonth(), kini.getDate(), 23, 59, 59, 999);
  if (!dari) dari = new Date(sampai.getFullYear(), sampai.getMonth() - 5, 1);
  if (dari > sampai) throw new Error('Tanggal "dari" tidak boleh setelah tanggal "sampai".');
  return { dari: dari, sampai: sampai };
}

function dalamRentang(tanggal, r) {
  const t = new Date(tanggal);
  return !isNaN(t.getTime()) && t >= r.dari && t <= r.sampai;
}

// ============================================================
// 1. DASHBOARD PER PERIODE
// ============================================================

function apiDashboardPeriode(f) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const r = rentangDariFilter(f);

  // Ember per bulan dalam rentang (maksimal 24 bulan terakhir rentang)
  const ember = [];
  const indeks = {};
  let kursor = new Date(r.dari.getFullYear(), r.dari.getMonth(), 1);
  while (kursor <= r.sampai) {
    const kunci = kursor.getFullYear() + '-' + kursor.getMonth();
    indeks[kunci] = ember.length;
    ember.push({ label: NAMA_BULAN_PENDEK[kursor.getMonth()] + ' ' +
      String(kursor.getFullYear()).slice(2), masuk: 0, keluar: 0 });
    kursor = new Date(kursor.getFullYear(), kursor.getMonth() + 1, 1);
  }

  let masuk = 0, keluar = 0;
  const perKategori = {};
  sheetToObjects(SHEET.KAS).forEach(function(k) {
    if (k.status_lock === STATUS_LOCK.VOID || !dalamRentang(k.tanggal, r)) return;
    const m = Number(k.masuk) || 0, kl = Number(k.keluar) || 0;
    masuk += m; keluar += kl;
    const kat = String(k.kategori || 'lainnya');
    if (!perKategori[kat]) perKategori[kat] = { masuk: 0, keluar: 0 };
    perKategori[kat].masuk += m; perKategori[kat].keluar += kl;

    const t = new Date(k.tanggal);
    const i = indeks[t.getFullYear() + '-' + t.getMonth()];
    if (i !== undefined) { ember[i].masuk += m; ember[i].keluar += kl; }
  });

  let jasaPinjaman = 0;
  sheetToObjects(SHEET.ANGSURAN).forEach(function(a) {
    if (a.status_lock === STATUS_LOCK.VOID || !dalamRentang(a.tanggal, r)) return;
    jasaPinjaman += Number(a.jasa) || 0;
  });

  // Komposisi simpanan saat ini (semua anggota)
  const komposisi = { pokok: 0, wajib: 0, sukarela: 0 };
  const peta = getPetaSaldoSimpanan();
  Object.keys(peta).forEach(function(id) {
    komposisi.pokok += peta[id].pokok;
    komposisi.wajib += peta[id].wajib;
    komposisi.sukarela += peta[id].sukarela;
  });

  return {
    dari: formatTgl(r.dari, 'yyyy-MM-dd'), sampai: formatTgl(r.sampai, 'yyyy-MM-dd'),
    masuk: masuk, keluar: keluar, bersih: masuk - keluar,
    jasaPinjaman: jasaPinjaman,
    saldoKasAkhir: getSaldoKas(r.sampai),
    perKategori: perKategori,
    bulanan: ember.slice(-24),
    komposisi: komposisi
  };
}

// ============================================================
// 2. RIWAYAT TRANSAKSI (gabungan)
// ============================================================

/**
 * @param {Object} f { jenis: 'semua'|'setoran'|'pengambilan'|'angsuran'|'pinjaman',
 *                     dari, sampai }
 */
function apiRiwayatTransaksi(f) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  f = f || {};
  const r = rentangDariFilter(f);
  const jenis = f.jenis || 'semua';
  const ikut = function(j) { return jenis === 'semua' || jenis === j; };

  const namaAnggota = {};
  sheetToObjects(SHEET.ANGGOTA).forEach(function(a) {
    namaAnggota[String(a.id_anggota)] = a.nama;
  });

  const hasil = [];
  const tambah = function(row, isi) {
    if (!dalamRentang(row.tanggal, r)) return;
    const idAgt = String(row.id_anggota || '');
    isi.id_anggota = idAgt;
    isi.nama = namaAnggota[idAgt] || idAgt;
    isi.tgl = formatTgl(row.tanggal);
    isi.urut = new Date(row.tanggal).getTime() || 0;
    isi.urutWaktu = new Date(row.timestamp || row.tanggal).getTime() || 0;
    isi.keterangan = String(row.keterangan || '');
    isi.lock = row.status_lock || '';
    hasil.push(isi);
  };

  if (ikut('setoran')) sheetToObjects(SHEET.SIMPANAN).forEach(function(s) {
    tambah(s, { jenis: 'setoran', id: s.id_transaksi,
      uraian: 'Setoran ' + s.jenis_simpanan, masuk: Number(s.jumlah_setoran) || 0, keluar: 0,
      status: s.status_lock });
  });
  if (ikut('pengambilan')) sheetToObjects(SHEET.PENGAMBILAN).forEach(function(p) {
    tambah(p, { jenis: 'pengambilan', id: p.id_pengambilan,
      uraian: 'Tarik ' + (p.jenis_simpanan || 'sukarela'), masuk: 0, keluar: Number(p.jumlah) || 0,
      status: p.status_lock === STATUS_LOCK.VOID ? 'VOID' : p.status_approval });
  });
  if (ikut('angsuran')) sheetToObjects(SHEET.ANGSURAN).forEach(function(a) {
    tambah(a, { jenis: 'angsuran', id: a.id_angsuran,
      uraian: 'Angsuran ' + a.id_pinjaman, masuk: Number(a.total_bayar) || 0, keluar: 0,
      status: a.status_lock });
  });
  if (ikut('pinjaman')) sheetToObjects(SHEET.PINJAMAN).forEach(function(p) {
    const migrasi = String(p.keterangan || '').indexOf('MIGRASI') === 0;
    tambah(p, { jenis: 'pinjaman', id: p.id_pinjaman,
      uraian: migrasi ? 'Pinjaman migrasi'
        : String(p.id_induk || '') ? 'Tambahan pinjaman (' + p.id_induk + ')' : 'Pencairan pinjaman',
      masuk: 0, keluar: migrasi ? 0 : Number(p.nominal) || 0,
      status: p.status_lock === STATUS_LOCK.VOID ? 'VOID' : p.status });
  });

  hasil.sort(function(a, b) { return (b.urut - a.urut) || (b.urutWaktu - a.urutWaktu); });
  hasil.forEach(function(h) { delete h.urut; delete h.urutWaktu; });
  return {
    dari: formatTgl(r.dari, 'yyyy-MM-dd'), sampai: formatTgl(r.sampai, 'yyyy-MM-dd'),
    rows: hasil
  };
}

// ============================================================
// 3. STRUK
// ============================================================

function namaPetugas(email) {
  const u = String(email || '').trim().toLowerCase();
  if (!u || u === 'unknown') return '-';
  let daftar = [];
  try { daftar = sheetToObjects(SHEET_PENGGUNA); } catch (e) { /* sheet belum ada */ }
  const p = daftar.find(function(x) { return String(x.email).trim().toLowerCase() === u; });
  return p && p.nama ? String(p.nama) : u.split('@')[0];
}

/**
 * Isi struk satu transaksi.
 * @return {Object} { judul, no, tanggal, nama, id_anggota, rincian:[{l,v,rp}],
 *                    utama:{l,v}, catatan, petugas, dicetak, batal }
 */
function dataStruk(jenis, id) {
  const cari = {
    setoran: [SHEET.SIMPANAN, 'id_transaksi'],
    pengambilan: [SHEET.PENGAMBILAN, 'id_pengambilan'],
    angsuran: [SHEET.ANGSURAN, 'id_angsuran'],
    pinjaman: [SHEET.PINJAMAN, 'id_pinjaman']
  }[jenis];
  if (!cari) throw new Error('Jenis struk tidak dikenal: ' + jenis);
  const row = getRowByField(cari[0], cari[1], id);
  if (!row) throw new Error('Transaksi ' + id + ' tidak ditemukan.');

  const anggota = getAnggota(row.id_anggota);
  const s = {
    no: String(id),
    tanggal: formatTgl(row.tanggal) + (row.timestamp ? ' ' + formatTgl(row.timestamp, 'HH:mm') : ''),
    nama: anggota ? String(anggota.nama) : String(row.id_anggota),
    id_anggota: String(row.id_anggota),
    rincian: [], catatan: '',
    petugas: namaPetugas(row.user_input),
    dicetak: formatTgl(new Date(), 'dd/MM/yyyy HH:mm'),
    batal: row.status_lock === STATUS_LOCK.VOID
  };

  if (jenis === 'setoran') {
    s.judul = 'BUKTI SETORAN SIMPANAN';
    s.rincian.push({ l: 'Jenis', v: 'Simpanan ' + row.jenis_simpanan });
    s.utama = { l: 'JUMLAH', v: Number(row.jumlah_setoran) };
    s.rincian2 = [{ l: 'Saldo ' + row.jenis_simpanan + ' kini',
      v: getSaldoSimpanan(row.id_anggota, row.jenis_simpanan), rp: true }];
  } else if (jenis === 'pengambilan') {
    s.judul = 'BUKTI PENGAMBILAN SIMPANAN';
    s.rincian.push({ l: 'Jenis', v: 'Simpanan ' + (row.jenis_simpanan || 'sukarela') });
    s.rincian.push({ l: 'Status', v: String(row.status_approval) });
    s.utama = { l: 'JUMLAH', v: Number(row.jumlah) };
    s.rincian2 = [{ l: 'Saldo ' + (row.jenis_simpanan || 'sukarela') + ' kini',
      v: getSaldoSimpanan(row.id_anggota, row.jenis_simpanan || 'sukarela'), rp: true }];
  } else if (jenis === 'angsuran') {
    s.judul = 'BUKTI PEMBAYARAN ANGSURAN';
    const pinjaman = getRowByField(SHEET.PINJAMAN, 'id_pinjaman', row.id_pinjaman);
    // Urutan angsuran & sisa SETELAH pembayaran ini (hanya yang tidak VOID)
    const semua = getRowsByFilter(SHEET.ANGSURAN, function(a) {
      return String(a.id_pinjaman) === String(row.id_pinjaman) &&
             a.status_lock !== STATUS_LOCK.VOID;
    }).sort(function(a, b) { return a.__row - b.__row; });
    const posisi = semua.findIndex(function(a) { return a.__row === row.__row; });
    const menurun = isJasaMenurun(pinjaman);
    let terbayar = 0;
    for (let i = 0; i <= posisi; i++) {
      terbayar += Number(menurun ? semua[i].angsuran_pokok : semua[i].total_bayar) || 0;
    }

    s.rincian.push({ l: 'No. pinjaman', v: String(row.id_pinjaman) });
    if (posisi >= 0) {
      s.rincian.push({ l: 'Angsuran ke', v: String(posisi + 1) +
        (pinjaman && !menurun ? ' dari ' + pinjaman.tenor : '') });
    }
    s.rincian.push({ l: 'Pokok', v: Number(row.angsuran_pokok), rp: true });
    s.rincian.push({ l: 'Jasa', v: Number(row.jasa), rp: true });
    if (Number(row.denda) > 0) s.rincian.push({ l: 'Denda', v: Number(row.denda), rp: true });
    s.utama = { l: 'TOTAL BAYAR', v: Number(row.total_bayar) };
    if (pinjaman && posisi >= 0) {
      // Jasa menurun: sisa POKOK = pokok dicairkan s/d saat bayar − pokok terbayar
      const batas = new Date(row.timestamp || row.tanggal);
      const pokokCair = menurun
        ? getBarisPokokPinjaman(pinjaman.id_pinjaman).reduce(function(t, r) {
            return new Date(r.timestamp || r.tanggal) <= batas || !String(r.id_induk || '')
              ? t + (Number(r.nominal) || 0) : t;
          }, 0)
        : Number(pinjaman.total_tagihan);
      const sisa = pokokCair - terbayar;
      s.rincian2 = [{ l: menurun ? 'Sisa pokok' : 'Sisa pinjaman', v: sisa, rp: true }];
      if (sisa <= 0) s.catatan = 'PINJAMAN LUNAS';
    }
  } else {
    const idInduk = String(row.id_induk || '');
    s.judul = idInduk ? 'BUKTI TAMBAHAN PINJAMAN' : 'BUKTI PENCAIRAN PINJAMAN';
    if (idInduk) s.rincian.push({ l: 'Digabung ke', v: idInduk });
    s.rincian.push({ l: 'Tenor', v: row.tenor + ' bulan' });
    s.utama = { l: 'POKOK DICAIRKAN', v: Number(row.nominal) };
    if (isJasaMenurun(row)) {
      s.rincian.push({ l: 'Jasa', v: row.jasa_persen + '%/bln dr sisa' });
      s.rincian2 = [{ l: 'Angsuran pokok/bln', v: Number(row.angsuran_perbulan), rp: true }];
      if (idInduk) {
        s.rincian2.unshift({ l: 'Total sisa pokok', v: getSisaPinjaman(idInduk), rp: true });
      }
    } else {
      s.rincian.push({ l: 'Jasa', v: row.jasa_persen + '% / bulan' });
      s.rincian2 = [
        { l: 'Total tagihan', v: Number(row.total_tagihan), rp: true },
        { l: 'Angsuran/bulan', v: Number(row.angsuran_perbulan), rp: true }
      ];
    }
  }
  if (!s.rincian2) s.rincian2 = [];
  if (!s.catatan && row.keterangan) s.catatan = String(row.keterangan);
  return s;
}

function apiStruk(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  return dataStruk(d.jenis, d.id);
}

// ===== sesi.js =====
/**
 * ============================================================
 * SESI.GS — LOGIN MANDIRI (email + kata sandi)
 * Koperasi App
 * ============================================================
 *
 * LATAR BELAKANG:
 * Google Apps Script TIDAK memberitahu identitas pengunjung
 * (Session.getActiveUser() kosong) bila web app dideploy dengan
 * "Execute as: Me" untuk pengguna di luar domain Workspace.
 * Karena itu login berbasis akun Google tidak bisa diandalkan.
 * File ini menggantinya dengan login sendiri: email + sandi,
 * disimpan ter-hash di sheet `pengguna`.
 *
 * KEAMANAN:
 * - Sandi TIDAK PERNAH disimpan apa adanya. Yang disimpan adalah
 *   hash SHA-256 dengan garam (salt) acak per pengguna,
 *   diulang 600 kali agar sulit ditebak paksa.
 * - Token sesi acak, berlaku 8 jam, disimpan di CacheService.
 * - Semua panggilan API wajib menyertakan token (lihat apiJalan).
 *
 * LANGKAH PERTAMA (WAJIB, sekali saja):
 * Buka editor Apps Script → pilih fungsi `buatAdminPertama` →
 * ubah dulu email/nama/sandi di dalamnya → klik Run.
 * ============================================================
 */

const KOLOM_SANDI = ['sandi_hash', 'garam', 'wajib_ganti', 'terakhir_login'];
const MASA_SESI_DETIK = 8 * 60 * 60;   // 8 jam
const PUTARAN_HASH = 600;

// Diisi oleh apiJalan sebelum fungsi API dijalankan.
let __SESI_AKTIF = null;

// ============================================================
// PENYIAPAN SHEET
// ============================================================

/** Tambahkan kolom sandi ke sheet `pengguna` bila belum ada. */
function pastikanKolomSandi() {
  const sheet = pastikanSheetPengguna();
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h).trim(); });

  const kurang = KOLOM_SANDI.filter(function(k) { return header.indexOf(k) === -1; });
  if (kurang.length === 0) return sheet;

  sheet.getRange(1, lastCol + 1, 1, kurang.length).setValues([kurang])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#0E4D3C');
  return sheet;
}

// ============================================================
// SANDI
// ============================================================

function buatGaram() {
  return Utilities.getUuid().replace(/-/g, '') +
         Utilities.getUuid().replace(/-/g, '').slice(0, 8);
}

/** Hash sandi berulang agar tahan tebak-paksa. */
function hashSandi(sandi, garam) {
  let nilai = garam + '::' + sandi;
  for (let i = 0; i < PUTARAN_HASH; i++) {
    const byte = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256, nilai, Utilities.Charset.UTF_8);
    nilai = byte.map(function(b) {
      return ((b < 0 ? b + 256 : b) + 0x100).toString(16).slice(1);
    }).join('');
  }
  return nilai;
}

function validasiKekuatanSandi(sandi) {
  if (String(sandi || '').length < 6) {
    throw new Error('Kata sandi minimal 6 karakter.');
  }
}

// ============================================================
// TOKEN SESI
// ============================================================

function simpanSesi_(email) {
  const token = Utilities.getUuid() + '-' + Utilities.getUuid();
  const isi = JSON.stringify({
    email: email,
    kedaluwarsa: Date.now() + MASA_SESI_DETIK * 1000
  });
  CacheService.getScriptCache().put('sesi_' + token, isi, MASA_SESI_DETIK);
  // Cadangan bila cache dibersihkan Google sebelum waktunya
  PropertiesService.getScriptProperties().setProperty('sesi_' + token, isi);
  return token;
}

function bacaSesi_(token) {
  if (!token) return null;
  const kunci = 'sesi_' + token;
  let isi = CacheService.getScriptCache().get(kunci);
  if (!isi) isi = PropertiesService.getScriptProperties().getProperty(kunci);
  if (!isi) return null;

  let data;
  try { data = JSON.parse(isi); } catch (e) { return null; }
  if (!data.kedaluwarsa || Date.now() > data.kedaluwarsa) {
    hapusSesi_(token);
    return null;
  }
  return data;
}

function hapusSesi_(token) {
  if (!token) return;
  CacheService.getScriptCache().remove('sesi_' + token);
  PropertiesService.getScriptProperties().deleteProperty('sesi_' + token);
}

/** Bersihkan sesi kedaluwarsa (dipanggil tugas harian). */
function bersihkanSesiKedaluwarsa() {
  const props = PropertiesService.getScriptProperties();
  const semua = props.getProperties();
  let hapus = 0;
  Object.keys(semua).forEach(function(k) {
    if (k.indexOf('sesi_') !== 0) return;
    try {
      const d = JSON.parse(semua[k]);
      if (!d.kedaluwarsa || Date.now() > d.kedaluwarsa) {
        props.deleteProperty(k); hapus++;
      }
    } catch (e) { props.deleteProperty(k); hapus++; }
  });
  return hapus;
}

// ============================================================
// PENGGUNA
// ============================================================

function cariBarisPengguna(email) {
  const target = String(email || '').trim().toLowerCase();
  if (!target) return null;
  const daftar = sheetToObjects(SHEET_PENGGUNA);
  return daftar.find(function(u) {
    return String(u.email).trim().toLowerCase() === target;
  }) || null;
}

/**
 * JALANKAN SEKALI DARI EDITOR untuk membuat admin pertama.
 *
 * Kata sandi TIDAK ditulis di kode (kode ini ikut tersalin ke mana-mana
 * lewat clasp/berbagi proyek). Isi dulu Script Properties:
 *   Project Settings (ikon ⚙️) → Script Properties → Add script property
 *     ADMIN_EMAIL = email admin
 *     ADMIN_NAMA  = nama admin
 *     ADMIN_SANDI = kata sandi (minimal 6 karakter)
 * lalu Run fungsi ini. ADMIN_SANDI otomatis dihapus setelah dipakai.
 */
function buatAdminPertama() {
  const props = PropertiesService.getScriptProperties();
  const EMAIL = props.getProperty('ADMIN_EMAIL') || '';
  const NAMA  = props.getProperty('ADMIN_NAMA') || '';
  const SANDI = props.getProperty('ADMIN_SANDI') || '';
  if (!EMAIL || !SANDI) {
    throw new Error('Isi Script Properties ADMIN_EMAIL dan ADMIN_SANDI dulu ' +
      '(Project Settings → Script Properties), lalu jalankan lagi.');
  }

  pastikanKolomSandi();
  validasiKekuatanSandi(SANDI);

  const email = EMAIL.trim().toLowerCase();
  const garam = buatGaram();
  const hash = hashSandi(SANDI, garam);
  const ada = cariBarisPengguna(email);

  if (ada) {
    const sheet = pastikanKolomSandi();
    updateRowByField(SHEET_PENGGUNA, 'email', ada.email, {
      sandi_hash: hash, garam: garam, peran: 'admin',
      status: 'aktif', wajib_ganti: 'tidak'
    });
    Logger.log('Sandi admin diperbarui untuk ' + email);
  } else {
    appendRowFromObject(SHEET_PENGGUNA, {
      email: email, nama: NAMA || email.split('@')[0], peran: 'admin', status: 'aktif',
      ditambahkan_oleh: 'sistem', timestamp: new Date(),
      sandi_hash: hash, garam: garam, wajib_ganti: 'tidak',
      terakhir_login: ''
    });
    Logger.log('Admin pertama dibuat: ' + email);
  }
  props.deleteProperty('ADMIN_SANDI');
  Logger.log('Silakan masuk ke web app dengan email dan sandi tersebut. ' +
    'ADMIN_SANDI sudah dihapus dari Script Properties.');
  return 'OK';
}

// ============================================================
// API LOGIN
// ============================================================

/** Masuk. @return {Object} {token, nama, peran, email, wajibGanti} */
function apiMasuk(d) {
  return jalankanPublik_(function() { return masuk_(d); });
}

function masuk_(d) {
  pastikanKolomSandi();
  const email = String(d.email || '').trim().toLowerCase();
  const sandi = String(d.sandi || '');
  const pesanUmum = 'Email atau kata sandi salah.';

  if (!email || !sandi) throw new Error('Email dan kata sandi wajib diisi.');
  cekBatasLogin_(email);

  const u = cariBarisPengguna(email);
  if (!u) { catatGagalLogin_(email); throw new Error(pesanUmum); }

  const status = String(u.status || '').trim().toLowerCase();
  if (status && status !== 'aktif') {
    throw new Error('Akun Anda dinonaktifkan. Hubungi admin koperasi.');
  }
  if (!u.sandi_hash || !u.garam) {
    throw new Error('Akun ini belum punya kata sandi. Minta admin ' +
      'mengatur ulang kata sandi Anda.');
  }
  if (hashSandi(sandi, String(u.garam)) !== String(u.sandi_hash)) {
    catatGagalLogin_(email);
    throw new Error(pesanUmum);
  }
  CacheService.getScriptCache().remove('gagal_' + email);

  updateRowByField(SHEET_PENGGUNA, 'email', u.email,
    { terakhir_login: new Date() });

  const peran = String(u.peran || '').trim().toLowerCase() === 'admin'
    ? 'admin' : 'operator';

  return {
    token: simpanSesi_(email),
    email: email,
    nama: u.nama || email.split('@')[0],
    peran: peran,
    wajibGanti: String(u.wajib_ganti || '').trim().toLowerCase() === 'ya'
  };
}

/** Maksimal 5 kali salah sandi per email; setelah itu dikunci 15 menit. */
const BATAS_GAGAL_LOGIN = 5;
const KUNCI_LOGIN_DETIK = 15 * 60;
function cekBatasLogin_(email) {
  const n = Number(CacheService.getScriptCache().get('gagal_' + email) || 0);
  if (n >= BATAS_GAGAL_LOGIN) {
    throw new Error('Terlalu banyak percobaan masuk yang salah. Coba lagi 15 menit lagi, ' +
      'atau minta admin mengatur ulang kata sandi.');
  }
}
function catatGagalLogin_(email) {
  const c = CacheService.getScriptCache();
  const n = Number(c.get('gagal_' + email) || 0) + 1;
  c.put('gagal_' + email, String(n), KUNCI_LOGIN_DETIK);
}

function apiKeluar(token) {
  hapusSesi_(token);
  return 'Anda telah keluar.';
}

/** Cek token masih berlaku (dipakai saat halaman dimuat ulang). */
function apiCekSesi(token) {
  return jalankanPublik_(function() { return cekSesi_(token); });
}

function cekSesi_(token) {
  const sesi = bacaSesi_(token);
  if (!sesi) return null;
  const u = cariBarisPengguna(sesi.email);
  if (!u) return null;
  const status = String(u.status || '').trim().toLowerCase();
  if (status && status !== 'aktif') return null;

  return {
    email: sesi.email,
    nama: u.nama || sesi.email.split('@')[0],
    peran: String(u.peran || '').trim().toLowerCase() === 'admin'
      ? 'admin' : 'operator',
    wajibGanti: String(u.wajib_ganti || '').trim().toLowerCase() === 'ya'
  };
}

/** Ganti sandi sendiri. */
function apiGantiSandi(d) {
  const profil = requireRole(['admin', 'operator']);
  const u = cariBarisPengguna(profil.email);
  if (!u) throw new Error('Pengguna tidak ditemukan.');

  if (hashSandi(String(d.lama || ''), String(u.garam)) !== String(u.sandi_hash)) {
    throw new Error('Kata sandi lama salah.');
  }
  validasiKekuatanSandi(d.baru);
  if (String(d.baru) === String(d.lama)) {
    throw new Error('Kata sandi baru harus berbeda dari yang lama.');
  }

  const garam = buatGaram();
  updateRowByField(SHEET_PENGGUNA, 'email', u.email, {
    sandi_hash: hashSandi(String(d.baru), garam),
    garam: garam, wajib_ganti: 'tidak'
  });
  return '✅ Kata sandi berhasil diganti.';
}

/** Admin mengatur ulang sandi pengguna lain. */
function apiResetSandi(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  const u = cariBarisPengguna(d.email);
  if (!u) throw new Error('Pengguna tidak ditemukan.');
  validasiKekuatanSandi(d.sandi);

  const garam = buatGaram();
  updateRowByField(SHEET_PENGGUNA, 'email', u.email, {
    sandi_hash: hashSandi(String(d.sandi), garam),
    garam: garam, wajib_ganti: 'ya'
  });
  logAktivitas('UPDATE', SHEET_PENGGUNA, u.email, null, { aksi: 'reset sandi' });
  return '✅ Kata sandi ' + u.email + ' diatur ulang. ' +
    'Pengguna akan diminta menggantinya saat masuk.';
}

// ============================================================
// PENJAGA AKSES DATA
// ============================================================
//
// PENTING: di Apps Script, SEMUA fungsi server yang namanya tidak
// diakhiri "_" bisa dipanggil langsung dari browser lewat
// google.script.run — tanpa melewati apiJalan & pengecekan token.
// Karena itu setiap akses sheet (getSheet) & aksi sensitif lain
// (email, Drive, trigger, setup/reset) wajib lolos penjaga ini.
//
// Diizinkan jika salah satu:
// 1. Dipanggil lewat apiJalan dengan token sesi yang sah
// 2. Jalur publik yang memang sengaja (doGet, apiMasuk, apiCekSesi)
// 3. Pengguna Google yang sedang menjalankan (editor Apps Script /
//    menu spreadsheet) adalah pemilik script atau terdaftar AKTIF
//    di sheet `pengguna`.

let __AKSES_PUBLIK = false;   // disetel sementara oleh jalur publik
let __AKSES_OK = false;       // hasil cek pengguna Google (per eksekusi)

function pastikanAksesData() {
  if (__AKSES_PUBLIK || __AKSES_OK) return;
  if (typeof __SESI_AKTIF !== 'undefined' && __SESI_AKTIF) return;

  // Tanpa sesi login, hanya boleh bila kode berjalan SEBAGAI orang yang
  // menjalankannya: editor/menu spreadsheet/pemicu (aktif === efektif).
  // Di web app (dijalankan sebagai pemilik) pengunjung selalu aktif ≠ efektif,
  // sehingga panggilan google.script.run langsung ke fungsi mana pun —
  // termasuk oleh pengguna terdaftar yang memakai akun Google satu domain —
  // ditolak dan wajib lewat apiJalan (token + cek peran).
  const aktif = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (aktif) {
    const efektif = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
    if (aktif === efektif) {
      __AKSES_OK = true;
      return;
    }
  }
  throw new Error('Akses ditolak. Silakan masuk lewat aplikasi koperasi.');
}

/** Jalankan fn sebagai jalur publik yang disengaja (login, halaman awal). */
function jalankanPublik_(fn) {
  const sebelum = __AKSES_PUBLIK;
  __AKSES_PUBLIK = true;
  try { return fn(); } finally { __AKSES_PUBLIK = sebelum; }
}

// ============================================================
// GERBANG SEMUA API
// ============================================================

/**
 * Satu-satunya pintu panggilan dari klien.
 * Menyetel sesi aktif, lalu menjalankan fungsi API yang diminta.
 */
function apiJalan(nama, arg, token) {
  if (!/^api[A-Za-z0-9]+$/.test(String(nama))) {
    throw new Error('Fungsi tidak dikenal.');
  }
  if (nama === 'apiJalan' || nama === 'apiMasuk' || nama === 'apiCekSesi') {
    throw new Error('Fungsi tidak boleh dipanggil dari sini.');
  }

  const sesi = bacaSesi_(token);
  if (!sesi) throw new Error('SESI_HABIS');
  __SESI_AKTIF = sesi.email;

  const fn = globalThis[nama];
  if (typeof fn !== 'function') throw new Error('Fungsi ' + nama + ' tidak ada.');

  // Semua API yang MENULIS dijalankan bergiliran. Tanpa kunci, dua
  // operator yang menyimpan bersamaan bisa mendapat ID kembar
  // (generateId = nomor terbesar + 1) atau Generate Angsuran dobel.
  const lock = API_BACA_SAJA[nama] ? null : LockService.getScriptLock();
  if (lock && !lock.tryLock(30000)) {
    throw new Error('Sistem sedang memproses transaksi lain. Coba lagi sebentar.');
  }

  try {
    return arg === undefined || arg === null ? fn() : fn(arg);
  } finally {
    __SESI_AKTIF = null;
    if (lock) {
      SpreadsheetApp.flush(); // pastikan tulisan tersimpan sebelum giliran berikutnya
      lock.releaseLock();
    }
  }
}

/**
 * API yang hanya MEMBACA data — boleh jalan bersamaan tanpa kunci.
 * API baru yang tidak terdaftar di sini otomatis dikunci (aman).
 */
const API_BACA_SAJA = {
  apiDataAwal: true, apiDashboard: true, apiAnggotaList: true,
  apiAnggotaDetail: true, apiSaldoAnggota: true, apiPendingList: true,
  apiDataSHU: true, apiPenggunaList: true, apiPinjamanAnggota: true,
  apiDaftarAkunBeban: true, apiMutasiKas: true, apiSettingList: true,
  apiCekTutupBuku: true, apiTrenBulanan: true, apiLogAktivitas: true,
  apiKeluar: true, apiKeluarRekap: true,
  apiDashboardPeriode: true, apiRiwayatTransaksi: true, apiStruk: true,
  apiProfilKoperasi: true, apiTagihanBulanan: true,
  apiLapBukuKas: true, apiLapLabaRugi: true, apiLapNeraca: true, apiLapPenjelasan: true,
  apiLapSHU: true, apiStatusBulanan: true, apiDataRAT: true, apiMigrasiStatus: true
};

// ===== setup.js =====
/**
 * ============================================================
 * SETUP.GS v2 — Setup Otomatis Struktur Sheet Koperasi
 * Koperasi App
 * Sesuai dokumen: REKAP_Koperasi_v2.md (9 Juli 2026)
 * ============================================================
 *
 * PERUBAHAN v2:
 * - Catatan penjelas (note) otomatis terpasang di header
 *   setting_rat, saldo_awal, dan transaksi_pinjaman
 * - Perbaikan urutkanTab (bisa dijalankan langsung dari editor)
 * - Menu tambahan: Pasang Catatan Header
 *
 * CARA PAKAI (sekali saja):
 * 1. Buka Google Spreadsheet kosong yang baru
 * 2. Menu: Extensions → Apps Script
 * 3. Hapus isi Code.gs, paste seluruh file ini
 * 4. Simpan (Ctrl+S), lalu refresh spreadsheet-nya
 * 5. Akan muncul menu "🔧 SETUP" di atas → klik "Jalankan Setup Lengkap"
 * 6. Izinkan permission saat diminta (klik Advanced → Go to project)
 *
 * Script ini AMAN dijalankan berulang:
 * - Sheet yang sudah ada TIDAK ditimpa, hanya dilewati
 * - Untuk mengulang dari nol, gunakan menu "⚠️ Hapus Semua & Setup Ulang"
 * ============================================================
 */

// ============================================================
// KONFIGURASI STRUKTUR SEMUA SHEET
// Sumber kebenaran tunggal — sesuai BAGIAN 4 dokumen rekap v2
// ============================================================

const SHEET_CONFIG = [

  // ---------- DASHBOARD ----------
  {
    nama: 'DASHBOARD',
    warna: '#4a86e8', // biru
    header: [],  // dashboard diisi manual/script lain nanti
    keterangan: 'Ringkasan — diisi belakangan'
  },

  // ---------- KELOMPOK A: MASTER DATA (BIRU) ----------
  {
    nama: 'anggota',
    warna: '#4a86e8',
    header: ['id_anggota', 'nip_nis', 'nama', 'jabatan', 'alamat',
             'telepon', 'tanggal_masuk', 'status', 'sukarela_rutin']
  },
  {
    nama: 'setting_rat',
    warna: '#4a86e8',
    header: ['tahun', 'jasa_sukarela', 'jasa_pinjaman', 'shu_simpanan',
             'shu_jasa', 'dana_cadangan', 'dana_pengurus', 'dana_sosial',
             'tanggal_rat', 'status_aktif',
             'nominal_pokok', 'nominal_wajib', 'minimal_sukarela',
             'dana_anggota', 'dana_kesejahteraan', 'dana_pendidikan',
             'dana_pembangunan', 'pajak_persen'],
    // Nilai awal mengikuti Excel pembukuan: jasa sukarela 1%/tahun,
    // jasa pinjaman 1,5%/bulan (menurun), alokasi 25/50/10/5/5/2,5/2,5, pajak 0,5%
    dataAwal: [
      [2026, 1, 1.5, 40, 60, 25, 10, 2.5, '', true, 10000, 50000, 10000,
       50, 5, 5, 2.5, 0.5]
    ]
  },
  {
    nama: 'coa_akun',
    warna: '#4a86e8',
    header: ['kode', 'nama_akun', 'jenis'],
    dataAwal: [
      ['101', 'Kas', 'Aktiva'],
      ['102', 'Bank', 'Aktiva'],
      ['103', 'Piutang Pinjaman', 'Aktiva'],
      ['201', 'Simpanan Pokok', 'Kewajiban'],
      ['202', 'Simpanan Wajib', 'Kewajiban'],
      ['203', 'Simpanan Sukarela', 'Kewajiban'],
      ['301', 'Modal Koperasi', 'Ekuitas'],
      ['302', 'SHU Tahun Berjalan', 'Ekuitas'],
      ['303', 'SHU Ditahan', 'Ekuitas'],
      ['401', 'Pendapatan Jasa Pinjaman', 'Pendapatan'],
      ['402', 'Pendapatan Jasa Sukarela', 'Pendapatan'],
      ['501', 'Beban Operasional', 'Beban'],
      ['502', 'Dana Cadangan', 'Beban/Alokasi'],
      ['503', 'Dana Sosial', 'Beban/Alokasi'],
      ['504', 'Dana Pengurus', 'Beban/Alokasi'],
      ['505', 'Transport Kesejahteraan', 'Beban'],
      ['506', 'THR Lebaran', 'Beban'],
      ['507', 'Uang Duduk RAT', 'Beban'],
      ['508', 'Souvenir RAT', 'Beban'],
      ['509', 'Insentif Pengurus', 'Beban']
    ]
  },
  {
    nama: 'saldo_awal',
    warna: '#4a86e8',
    header: ['id', 'tahun', 'id_anggota', 'jenis', 'nominal', 'keterangan']
  },

  // ---------- KELOMPOK B: TRANSAKSI (HIJAU) ----------
  {
    nama: 'transaksi_simpanan',
    warna: '#6aa84f',
    header: ['id_transaksi', 'tanggal', 'tahun', 'bulan', 'id_anggota',
             'jenis_simpanan', 'jumlah_setoran', 'keterangan',
             'user_input', 'timestamp', 'status_lock']
  },
  {
    nama: 'transaksi_pengambilan',
    warna: '#6aa84f',
    header: ['id_pengambilan', 'tanggal', 'tahun', 'bulan', 'id_anggota',
             'jenis_simpanan', 'jumlah', 'status_approval', 'keterangan',
             'user_input', 'timestamp', 'status_lock']
  },
  {
    nama: 'transaksi_pinjaman',
    warna: '#6aa84f',
    header: ['id_pinjaman', 'tanggal', 'tahun', 'bulan', 'id_anggota',
             'nominal', 'tenor', 'jasa_persen', 'total_jasa',
             'total_tagihan', 'angsuran_perbulan', 'status', 'keterangan',
             'user_input', 'timestamp', 'status_lock',
             'metode_jasa', 'id_induk', 'angsuran_sebelumnya']
  },
  {
    nama: 'transaksi_angsuran',
    warna: '#6aa84f',
    header: ['id_angsuran', 'tanggal', 'tahun', 'bulan', 'id_pinjaman',
             'id_anggota', 'angsuran_pokok', 'jasa', 'denda', 'total_bayar',
             'keterangan', 'user_input', 'timestamp', 'status_lock']
  },
  {
    nama: 'transaksi_kas',
    warna: '#6aa84f',
    header: ['id_kas', 'tanggal', 'tahun', 'bulan', 'kategori', 'referensi',
             'keterangan', 'masuk', 'keluar',
             'user_input', 'timestamp', 'status_lock']
  },
  {
    nama: 'jurnal_umum',
    warna: '#f1c232', // kuning
    header: ['id_jurnal', 'tanggal', 'tahun', 'bulan', 'kode_akun',
             'nama_akun', 'debit', 'kredit', 'referensi', 'keterangan']
  },

  // ---------- KELOMPOK C: PERIODIK / DERIVED (KUNING) ----------
  {
    nama: 'rekap_jasa_sukarela',
    warna: '#f1c232',
    header: ['id', 'tahun', 'bulan', 'id_anggota', 'saldo_awal_bulan',
             'persen_jasa', 'nominal_jasa', 'status_posting']
  },
  {
    nama: 'realisasi_shu',
    warna: '#f1c232',
    header: ['id', 'tahun', 'id_anggota', 'shu_hak', 'shu_dibayar',
             'tanggal_bayar', 'metode', 'status', 'keterangan']
  },

  // ---------- KELOMPOK D: LOG (MERAH) ----------
  {
    nama: 'log_aktivitas',
    warna: '#cc0000',
    header: ['id', 'timestamp', 'user', 'aksi', 'sheet_target',
             'id_referensi', 'data_lama', 'data_baru']
  },

  // ---------- LAPORAN (ABU-ABU, read-only, generate otomatis) ----------
  { nama: 'lap_shu',         warna: '#999999', header: [], keterangan: 'Generate otomatis' },
  { nama: 'lap_neraca',      warna: '#999999', header: [], keterangan: 'Generate otomatis' },
  { nama: 'lap_labarugi',    warna: '#999999', header: [], keterangan: 'Generate otomatis' },
  { nama: 'rekap_pinjaman',  warna: '#999999', header: [], keterangan: 'Generate otomatis' },
  { nama: 'rekap_angsuran',  warna: '#999999', header: [], keterangan: 'Generate otomatis' },
  { nama: 'rekap_simpanan',  warna: '#999999', header: [], keterangan: 'Generate otomatis' },
  { nama: 'struk_potongan',  warna: '#999999', header: [], keterangan: 'Generate otomatis' }
];

// ============================================================
// VALIDASI DROPDOWN PER SHEET
// ============================================================

// Daftar pilihan (dropdown) di sheet. Sifatnya MEMBANTU saat sheet diedit
// manual: nilai di luar daftar hanya diberi tanda peringatan, tidak ditolak —
// supaya fitur baru tidak macet karena daftar ini lupa diperbarui.
const JABATAN_ANGGOTA = ['PNS', 'P3K Penuh Waktu', 'P3K Paruh Waktu', 'Honorer', 'Non PNS',
  'guru', 'karyawan', 'siswa'];
const VALIDASI_CONFIG = {
  'anggota': {
    'jabatan': JABATAN_ANGGOTA,
    'status':  ['aktif', 'nonaktif', 'keluar']
  },
  'saldo_awal': {
    'jenis': ['simpanan_pokok', 'simpanan_wajib', 'simpanan_sukarela', 'piutang', 'kas', 'inventaris',
      'penyusutan', 'cadangan', 'modal', 'dana_anggota', 'dana_pengurus', 'dana_kesejahteraan',
      'dana_pendidikan', 'dana_sosial', 'dana_pembangunan', 'shu']
  },
  'transaksi_simpanan': {
    'jenis_simpanan': ['pokok', 'wajib', 'sukarela'],
    'status_lock': ['OPEN', 'LOCKED', 'VOID']
  },
  'transaksi_pengambilan': {
    'jenis_simpanan': ['sukarela', 'wajib', 'pokok'],
    'status_approval': ['PENDING', 'APPROVED', 'REJECTED'],
    'status_lock': ['OPEN', 'LOCKED', 'VOID']
  },
  'transaksi_pinjaman': {
    'status': ['aktif', 'lunas', 'tambahan'],
    'status_lock': ['OPEN', 'LOCKED', 'VOID']
  },
  'transaksi_angsuran': {
    'status_lock': ['OPEN', 'LOCKED', 'VOID']
  },
  'transaksi_kas': {
    'kategori': ['setoran', 'angsuran', 'pengambilan', 'operasional', 'pinjaman', 'shu', 'jasa',
      'pendapatan_lain'],
    'status_lock': ['OPEN', 'LOCKED', 'VOID']
  },
  'rekap_jasa_sukarela': {
    'status_posting': ['DRAFT', 'POSTED']
  },
  'realisasi_shu': {
    'metode': ['tunai', 'transfer', 'potong_simpanan', 'ditahan'],
    'status': ['PENDING', 'PAID', 'PARTIAL']
  },
  'log_aktivitas': {
    'aksi': ['INSERT', 'EDIT', 'DELETE', 'VOID', 'APPROVE', 'REJECT']
  }
};

// ============================================================
// CATATAN PENJELAS HEADER (muncul saat kursor diarahkan ke sel)
// Format: NOTES_CONFIG[namaSheet][namaKolom] = teks catatan
// ============================================================

const NOTES_CONFIG = {

  // ---------- setting_rat ----------
  'setting_rat': {
    'tahun':
      'Tahun buku koperasi.\n' +
      'Contoh: 2026\n' +
      'Satu baris = satu tahun buku. Jangan ada tahun kembar.',

    'jasa_sukarela':
      'Persen jasa simpanan sukarela PER TAHUN (seperti Excel).\n' +
      'Isi ANGKA PERSEN saja, TANPA tanda %.\n' +
      'Contoh: 1 artinya 1% per tahun.\n' +
      'Jasa bulanan = saldo sukarela akhir bulan lalu × persen × 30/365.',

    'jasa_pinjaman':
      'Persen jasa (bunga) pinjaman PER BULAN.\n' +
      'Isi ANGKA PERSEN saja, TANPA tanda %.\n' +
      'Contoh: 1 artinya 1% per bulan dari pokok pinjaman.',

    'shu_simpanan':
      'Porsi SHU anggota yang dibagi berdasarkan SIMPANAN.\n' +
      'Isi ANGKA PERSEN. Contoh: 40 artinya 40%.\n' +
      'PENTING: shu_simpanan + shu_jasa harus = 100.\n' +
      'Basis: simpanan pokok + wajib saja (sukarela TIDAK ikut).',

    'shu_jasa':
      'Porsi SHU anggota yang dibagi berdasarkan JASA PINJAMAN\n' +
      '(bunga yang dibayar anggota lewat angsuran).\n' +
      'Isi ANGKA PERSEN. Contoh: 60 artinya 60%.\n' +
      'PENTING: shu_simpanan + shu_jasa harus = 100.',

    'dana_cadangan':
      'Persen dari SHU KOTOR yang disisihkan jadi dana cadangan.\n' +
      'Contoh: 20 artinya 20%.\n' +
      'Dipotong DULU sebelum SHU dibagi ke anggota\n' +
      '(Pasal 45 UU 25/1992).\n' +
      'Dana ini TIDAK PERNAH dibagikan ke anggota.',

    'dana_pengurus':
      'Persen dari SHU KOTOR untuk dana/insentif pengurus.\n' +
      'Contoh: 10 artinya 10%.\n' +
      'Dipotong dulu sebelum SHU dibagi ke anggota.',

    'dana_sosial':
      'Persen dari SHU KOTOR untuk dana sosial/kesejahteraan\n' +
      '(transport jenguk, THR, uang duduk RAT, dll).\n' +
      'Contoh: 10 artinya 10%.\n' +
      'Dipotong dulu sebelum SHU dibagi ke anggota.\n\n' +
      'CATATAN TOTAL: cadangan + pengurus + sosial = 40,\n' +
      'sisanya 60% jadi SHU anggota (lalu di-split\n' +
      'sesuai shu_simpanan & shu_jasa).',

    'tanggal_rat':
      'Tanggal pelaksanaan RAT yang memutuskan angka-angka\n' +
      'di baris ini. Format tanggal biasa, contoh: 15/01/2026.',

    'status_aktif':
      'TRUE  = tahun buku yang SEDANG BERJALAN.\n' +
      'FALSE = tahun sudah ditutup (tutup buku).\n' +
      'Hanya boleh ada SATU baris TRUE.\n' +
      'Input transaksi hanya diterima untuk tahun TRUE.'
  },

  // ---------- saldo_awal ----------
  'saldo_awal': {
    'id':
      'ID otomatis, format: SAL-TAHUN-NOMOR.\n' +
      'Contoh: SAL-2026-0001.\n' +
      'Nanti diisi otomatis oleh script. Jika input manual,\n' +
      'ikuti format dan jangan ada nomor kembar.',

    'tahun':
      'Tahun buku saat migrasi dimulai.\n' +
      'Contoh: 2026 (saldo per 1 Januari 2026).',

    'id_anggota':
      'ID anggota pemilik saldo. Harus sudah ada di sheet anggota.\n' +
      'Format: AGT-0001.',

    'jenis':
      'Jenis saldo yang dimigrasi dari Excel lama:\n' +
      '• simpanan_pokok    = saldo simpanan pokok\n' +
      '• simpanan_wajib    = akumulasi simpanan wajib\n' +
      '• simpanan_sukarela = saldo sukarela terakhir\n' +
      '• piutang = SISA pinjaman yang belum lunas\n' +
      '  (WAJIB dibuatkan juga baris pinjaman MIGRASI\n' +
      '   di transaksi_pinjaman — lihat catatan di sana)\n' +
      '• kas / shu = saldo kas koperasi / SHU ditahan.',

    'nominal':
      'Nilai saldo dalam RUPIAH penuh, tanpa titik/koma.\n' +
      'Contoh: 2000000 artinya Rp 2.000.000.',

    'keterangan':
      'Catatan bebas. Disarankan tulis sumber datanya,\n' +
      'contoh: "Migrasi dari Excel 2025, sheet Desember".'
  },

  // ---------- transaksi_pinjaman ----------
  'transaksi_pinjaman': {
    'id_pinjaman':
      'ID otomatis dari script, format: PJM-TAHUN-NOMOR.\n' +
      'Contoh: PJM-2026-0001. Jangan isi manual.',

    'tanggal':
      'Tanggal pencairan pinjaman.\n' +
      'Untuk pinjaman MIGRASI: isi tanggal pencairan ASLI\n' +
      'dari catatan lama.',

    'nominal':
      'Jumlah pokok pinjaman yang dicairkan (Rupiah penuh).\n' +
      'Untuk pinjaman MIGRASI: isi SISA POKOK yang belum\n' +
      'terbayar, bukan nominal awal.',

    'tenor':
      'Lama cicilan dalam BULAN. Maksimal 30.\n' +
      'Untuk pinjaman MIGRASI: isi SISA bulan angsuran.',

    'jasa_persen':
      'Persen jasa per bulan — otomatis diambil dari\n' +
      'setting_rat.jasa_pinjaman tahun berjalan.\n' +
      'Contoh: 1 artinya 1% per bulan.',

    'total_jasa':
      'DIHITUNG OTOMATIS oleh script:\n' +
      'nominal × jasa_persen% × tenor. Jangan isi manual.',

    'total_tagihan':
      'DIHITUNG OTOMATIS: nominal + total_jasa.\n' +
      'Sisa pinjaman = total_tagihan − jumlah angsuran masuk\n' +
      '(tidak disimpan, dihitung real-time).',

    'angsuran_perbulan':
      'DIHITUNG OTOMATIS: total_tagihan ÷ tenor.\n' +
      'Ini yang dipotong dari gaji tiap bulan.',

    'status':
      'aktif = masih ada sisa angsuran.\n' +
      'lunas = sudah terbayar penuh (diubah otomatis\n' +
      'oleh script saat sisa = 0).',

    'keterangan':
      'Catatan bebas.\n' +
      'KHUSUS pinjaman hasil migrasi dari Excel lama:\n' +
      'WAJIB tulis "MIGRASI 2025" (atau tahun sumbernya)\n' +
      'agar script tahu ini pinjaman lama.',

    'status_lock':
      'OPEN = normal, bisa diedit.\n' +
      'LOCKED = terkunci tutup buku.\n' +
      'VOID = dibatalkan (jangan hapus baris — void saja\n' +
      'agar jejak audit tetap ada).'
  }
};

// ============================================================
// MENU
// ============================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // ----- MENU HARIAN untuk admin -----
  ui.createMenu('📋 KOPERASI')
    .addItem('✅ Approve/Reject Pengambilan Sukarela', 'menuApprovalPengambilan')
    .addSeparator()
    .addItem('📅 Generate Angsuran Bulan Ini', 'menuGenerateAngsuranBulanIni')
    .addItem('🧾 Cetak Struk Potongan Gaji', 'menuCetakStrukPotongan')
    .addItem('💹 Hitung Jasa Sukarela Bulan Ini', 'menuHitungJasaSukarela')
    .addItem('🔒 Posting Jasa Sukarela (Kunci)', 'menuPostingJasaSukarela')
    .addSeparator()
    .addItem('📊 Refresh Dashboard', 'refreshDashboard')
    .addItem('📈 Laporan Laba Rugi', 'menuLaporanLabaRugi')
    .addItem('⚖️ Laporan Neraca', 'menuLaporanNeraca')
    .addItem('💰 Laporan SHU per Anggota', 'menuLaporanSHU')
    .addItem('📑 Rekap Simpanan', 'menuRekapSimpanan')
    .addItem('📑 Rekap Pinjaman', 'menuRekapPinjaman')
    .addItem('🗂️ Generate SEMUA Laporan', 'menuGenerateSemuaLaporan')
    .addToUi();

  // ----- MENU TEKNIS / TAHUNAN -----
  ui.createMenu('🔧 SETUP')
    .addItem('▶️ Jalankan Setup Lengkap', 'setupLengkap')
    .addItem('🔍 Cek Struktur (audit sheet & header)', 'cekStruktur')
    .addItem('📝 Pasang Catatan Header', 'pasangSemuaNotes')
    .addItem('🆔 Generate ID Anggota Kosong', 'generateIdAnggotaMassal')
    .addItem('🔍 Validasi Migrasi Pinjaman', 'validasiMigrasiPinjaman')
    .addSeparator()
    .addItem('🔎 Cek Kesiapan Tutup Buku', 'menuCekKesiapanTutupBuku')
    .addItem('📕 TUTUP BUKU TAHUN', 'menuTutupBukuTahun')
    .addSeparator()
    .addItem('🧹 Reset Data (Anggota Aman)', 'menuResetDataTransaksi')
    .addItem('⚠️ Hapus Semua & Setup Ulang', 'setupUlangDariNol')
    .addToUi();
}

// ============================================================
// FUNGSI UTAMA: SETUP LENGKAP
// ============================================================

function setupLengkap() {
  pastikanAksesData();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hasil = { dibuat: [], dilewati: [] };

  SHEET_CONFIG.forEach(function(cfg) {
    let sheet = ss.getSheetByName(cfg.nama);

    if (sheet) {
      hasil.dilewati.push(cfg.nama);
      return; // sudah ada → jangan sentuh, aman untuk data
    }

    sheet = ss.insertSheet(cfg.nama);
    hasil.dibuat.push(cfg.nama);

    // Warna tab
    sheet.setTabColor(cfg.warna);

    // Header
    if (cfg.header.length > 0) {
      const headerRange = sheet.getRange(1, 1, 1, cfg.header.length);
      headerRange.setValues([cfg.header]);
      headerRange
        .setFontWeight('bold')
        .setFontColor('#ffffff')
        .setBackground(cfg.warna)
        .setHorizontalAlignment('center');
      sheet.setFrozenRows(1);

      // Lebar kolom otomatis menyesuaikan header
      for (let c = 1; c <= cfg.header.length; c++) {
        sheet.setColumnWidth(c, Math.max(110, cfg.header[c - 1].length * 9 + 30));
      }

      // Rapikan: hapus kolom kosong berlebih di kanan
      const maxCols = sheet.getMaxColumns();
      if (maxCols > cfg.header.length) {
        sheet.deleteColumns(cfg.header.length + 1, maxCols - cfg.header.length);
      }

      // Dropdown validasi
      pasangValidasi(sheet, cfg.nama, cfg.header);
    } else if (cfg.keterangan) {
      sheet.getRange('A1').setValue('[' + cfg.nama + '] ' + cfg.keterangan)
        .setFontStyle('italic').setFontColor('#999999');
    }

    // Data awal (setting_rat & coa_akun)
    if (cfg.dataAwal && cfg.dataAwal.length > 0) {
      sheet.getRange(2, 1, cfg.dataAwal.length, cfg.dataAwal[0].length)
        .setValues(cfg.dataAwal);
    }
  });

  // Pasang catatan penjelas di header (setting_rat, saldo_awal, dll)
  pasangSemuaNotes(true); // true = mode diam, tanpa popup sendiri

  // Urutkan tab sesuai urutan SHEET_CONFIG
  urutkanTab(ss);

  // Hapus "Sheet1" bawaan jika masih ada dan kosong
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && sheet1.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }

  // Laporan hasil
  const pesan =
    '✅ SETUP SELESAI\n\n' +
    'Sheet dibuat baru (' + hasil.dibuat.length + '):\n' +
    (hasil.dibuat.length ? '• ' + hasil.dibuat.join('\n• ') : '(tidak ada)') +
    '\n\nSheet sudah ada, dilewati (' + hasil.dilewati.length + '):\n' +
    (hasil.dilewati.length ? '• ' + hasil.dilewati.join('\n• ') : '(tidak ada)') +
    '\n\nCatatan penjelas terpasang di header setting_rat,\n' +
    'saldo_awal, dan transaksi_pinjaman — arahkan kursor\n' +
    'ke sel header untuk membacanya.\n\n' +
    'Langkah berikutnya:\n' +
    '1. Cek tanggal_rat di setting_rat\n' +
    '2. Input data anggota\n' +
    '3. Input saldo_awal (migrasi Excel lama)\n' +
    '4. Buat pinjaman sintetis MIGRASI di transaksi_pinjaman';
  tampilkanPesan(pesan);
}

// ============================================================
// PASANG CATATAN HEADER
// Bisa dijalankan sendiri (menu / editor) atau dari setupLengkap
// ============================================================

function pasangSemuaNotes(diam) {
  pastikanAksesData();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalTerpasang = 0;
  const detail = [];

  Object.keys(NOTES_CONFIG).forEach(function(namaSheet) {
    const sheet = ss.getSheetByName(namaSheet);
    if (!sheet || sheet.getLastColumn() === 0) return;

    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0].map(String);
    let terpasang = 0;

    header.forEach(function(nama, i) {
      const catatan = NOTES_CONFIG[namaSheet][nama];
      if (catatan) {
        sheet.getRange(1, i + 1).setNote(catatan);
        terpasang++;
      }
    });

    totalTerpasang += terpasang;
    detail.push('• ' + namaSheet + ': ' + terpasang + ' catatan');
  });

  if (diam !== true) {
    tampilkanPesan(
      '✅ ' + totalTerpasang + ' catatan terpasang:\n\n' + detail.join('\n') +
      '\n\nArahkan kursor ke sel header untuk melihat penjelasannya.'
    );
  }
}

// ============================================================
// AUDIT STRUKTUR — cek sheet & header yang hilang/beda
// ============================================================

function cekStruktur() {
  pastikanAksesData();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masalah = [];

  SHEET_CONFIG.forEach(function(cfg) {
    const sheet = ss.getSheetByName(cfg.nama);
    if (!sheet) {
      masalah.push('❌ Sheet "' + cfg.nama + '" TIDAK ADA');
      return;
    }
    if (cfg.header.length === 0) return;

    const headerAda = sheet.getRange(1, 1, 1, cfg.header.length)
      .getValues()[0].map(String);
    cfg.header.forEach(function(h, i) {
      if (headerAda[i] !== h) {
        masalah.push('⚠️ ' + cfg.nama + ' kolom ' + (i + 1) +
          ': seharusnya "' + h + '", terisi "' + headerAda[i] + '"');
      }
    });
  });

  tampilkanPesan(
    masalah.length === 0
      ? '✅ Struktur LENGKAP & SESUAI dokumen rekap v2.\nSemua ' +
        SHEET_CONFIG.length + ' sheet ada dengan header benar.'
      : '🔍 HASIL AUDIT — ditemukan ' + masalah.length + ' masalah:\n\n' +
        masalah.join('\n')
  );
}

// ============================================================
// RESET TOTAL — hati-hati, menghapus semua data!
// ============================================================

function setupUlangDariNol() {
  pastikanAksesData();
  const ui = SpreadsheetApp.getUi();
  const jawab = ui.alert(
    '⚠️ PERINGATAN',
    'Ini akan MENGHAPUS SEMUA sheet sistem beserta ISINYA, lalu membuat ulang ' +
    'dari nol.\n\nSemua data transaksi akan HILANG dan tidak bisa dikembalikan.\n\n' +
    'Yakin lanjut?',
    ui.ButtonSet.YES_NO
  );
  if (jawab !== ui.Button.YES) return;

  const konfirmasi2 = ui.alert(
    '⚠️ KONFIRMASI TERAKHIR',
    'Sekali lagi: SEMUA DATA akan terhapus permanen. Lanjutkan?',
    ui.ButtonSet.YES_NO
  );
  if (konfirmasi2 !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Sheet sementara agar spreadsheet tidak pernah kosong total
  const temp = ss.insertSheet('_temp_setup_');

  SHEET_CONFIG.forEach(function(cfg) {
    const sheet = ss.getSheetByName(cfg.nama);
    if (sheet) ss.deleteSheet(sheet);
  });

  setupLengkap();

  const tempAkhir = ss.getSheetByName('_temp_setup_');
  if (tempAkhir) ss.deleteSheet(tempAkhir);
}

// ============================================================
// HELPER
// ============================================================

/**
 * Pasang ulang daftar pilihan di semua sheet sesuai VALIDASI_CONFIG terbaru
 * (sheet lama dibuat dengan daftar yang lebih sempit & menolak nilai baru).
 */
function perbaruiValidasiSheet() {
  pastikanAksesData();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(VALIDASI_CONFIG).forEach(function(nama) {
    const sheet = ss.getSheetByName(nama);
    if (!sheet || sheet.getLastColumn() === 0) return;
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    // Hapus dulu SEMUA aturan lama di baris data: kolom yang pernah bergeser bisa masih
    // membawa aturan kolom lain (mis. "OPEN/LOCKED/VOID" menempel di kolom lain).
    const maksBaris = Number(sheet.getMaxRows && sheet.getMaxRows()) || 1001;
    const maksKolom = Number(sheet.getMaxColumns && sheet.getMaxColumns()) || header.length;
    if (maksBaris > 1) sheet.getRange(2, 1, maksBaris - 1, maksKolom).clearDataValidations();
    pasangValidasi(sheet, nama, header);
  });
}

function pasangValidasi(sheet, namaSheet, header) {
  const aturan = VALIDASI_CONFIG[namaSheet];
  if (!aturan) return;

  Object.keys(aturan).forEach(function(namaKolom) {
    const idx = header.indexOf(namaKolom);
    if (idx === -1) return;

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(aturan[namaKolom], true)
      .setAllowInvalid(true)
      .setHelpText('Pilih: ' + aturan[namaKolom].join(' / '))
      .build();

    // Terapkan ke baris data (maks. 1000 baris pertama, tidak melewati ukuran sheet)
    const maks = typeof sheet.getMaxRows === 'function' ? Number(sheet.getMaxRows()) || 1001 : 1001;
    const jumlah = Math.min(1000, maks - 1);
    if (jumlah > 0) sheet.getRange(2, idx + 1, jumlah, 1).setDataValidation(rule);
  });
}

function urutkanTab(ss) {
  // Jika dipanggil langsung dari editor (tanpa parameter),
  // ambil spreadsheet aktif — perbaikan error "Cannot read
  // properties of undefined"
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  // Susun tab sesuai urutan SHEET_CONFIG (kiri ke kanan)
  for (let i = SHEET_CONFIG.length - 1; i >= 0; i--) {
    const sheet = ss.getSheetByName(SHEET_CONFIG[i].nama);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(1);
    }
  }
  // Kembali ke DASHBOARD
  const dash = ss.getSheetByName('DASHBOARD');
  if (dash) ss.setActiveSheet(dash);
}

// ===== simpanan.js =====
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

// ===== tagihan.js =====
/**
 * ============================================================
 * TAGIHAN.GS — Tagihan Bulanan (Potong Gaji) & Slip Tagihan
 * Koperasi App
 * ============================================================
 *
 * Pengganti sheet "Tag (1..12)" + "Str (1..12)" di Excel:
 * satu layar untuk melihat tagihan semua anggota, mencetak
 * Daftar Tagihan & Slip Tagihan, lalu membukukan potong gaji.
 *
 * Isi tagihan per anggota (bulan, tahun):
 *   1. Simpanan Pokok   — hanya jika belum pernah setor pokok
 *   2. Simpanan Wajib   — nominal dari Pengaturan tahun itu
 *   3. Simpanan Sukarela— setoran rutin anggota (kolom anggota.sukarela_rutin)
 *   4. Angsuran Pinjaman (pokok)
 *   5. Jasa Pinjaman    — jasa menurun: persen × sisa pokok awal bulan
 * + informasi: sisa piutang setelah bayar, saldo sukarela, dan jasa
 *   sukarela bulan ini (menambah saldo, TIDAK dipotong dari gaji —
 *   di Excel lama jasa ini tercampur ke baris sukarela).
 *
 * PROSES POTONG GAJI membukukan item yang belum tercatat untuk anggota
 * yang dipilih. Aman diulang: yang sudah tercatat dilewati.
 * ============================================================
 */

const NAMA_BULAN_PANJANG = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

function pastikanKolomAnggota() {
  return pastikanKolomTambahan(SHEET.ANGGOTA, ['sukarela_rutin']);
}

function tandaPotongGaji(tahun, bulan) {
  return 'Potong gaji ' + NAMA_BULAN_PANJANG[bulan - 1] + ' ' + tahun;
}

/** Rincian angsuran bulan (tahun, bulan) untuk satu pinjaman induk aktif. */
function rincianAngsuranBulan(p, tahun, bulan) {
  const idP = String(p.id_pinjaman);
  const sisa = getSisaPinjaman(idP);
  const sudah = getRowsByFilter(SHEET.ANGSURAN, function(a) {
    return String(a.id_pinjaman) === idP && a.status_lock !== STATUS_LOCK.VOID &&
           Number(a.tahun) === Number(tahun) && Number(a.bulan) === Number(bulan) &&
           String(a.keterangan || '').indexOf('Generate otomatis') === 0;
  })[0];
  if (sudah) {
    return { id: idP, pokok: Number(sudah.angsuran_pokok) || 0, jasa: Number(sudah.jasa) || 0,
      sisaSetelah: sisa, sudah: true };
  }
  if (sisa <= 0) return null;
  if (isJasaMenurun(p)) {
    if (new Date(p.tanggal) >= new Date(tahun, bulan - 1, 1)) {
      return { id: idP, pokok: 0, jasa: 0, sisaSetelah: sisa, sudah: false, baru: true };
    }
    const pokok = Math.min(Number(p.angsuran_perbulan) || 0, sisa);
    return { id: idP, pokok: pokok, jasa: hitungJasaMenurun(p, tahun, bulan),
      sisaSetelah: sisa - pokok, sudah: false };
  }
  // data lama (jasa flat)
  const total = Math.min(Number(p.angsuran_perbulan) || 0, sisa);
  const rasio = Number(p.total_tagihan) > 0 ? Number(p.nominal) / Number(p.total_tagihan) : 1;
  const pokok = Math.round(total * rasio);
  return { id: idP, pokok: pokok, jasa: total - pokok, sisaSetelah: sisa - total, sudah: false };
}

/**
 * Tagihan semua anggota aktif (+ anggota nonaktif yang masih punya pinjaman).
 * @return {Object} { tahun, bulan, namaBulan, baris: [...], total: {...} }
 */
function hitungTagihanBulanan(tahun, bulan) {
  tahun = Number(tahun); bulan = Number(bulan);
  if (!(bulan >= 1 && bulan <= 12)) throw new Error('Bulan tidak valid.');
  const setting = getSettingRAT(tahun);
  if (!setting) throw new Error('Setting tahun ' + tahun + ' belum ada.');
  const nominal = getNominalSimpanan(tahun);
  const tanda = tandaPotongGaji(tahun, bulan);

  const simpananBulanIni = getRowsByFilter(SHEET.SIMPANAN, function(r) {
    return r.status_lock !== STATUS_LOCK.VOID && Number(r.tahun) === tahun && Number(r.bulan) === bulan;
  });
  const sudahPokok = {};
  getRowsByFilter(SHEET.SIMPANAN, function(r) {
    return r.status_lock !== STATUS_LOCK.VOID && r.jenis_simpanan === 'pokok';
  }).forEach(function(r) { sudahPokok[String(r.id_anggota)] = true; });
  sheetToObjects(SHEET.SALDO_AWAL).forEach(function(r) {
    if (r.jenis === 'simpanan_pokok' && !isSaldoAwalSalinan(r)) sudahPokok[String(r.id_anggota)] = true;
  });

  // Pinjaman aktif + pinjaman yang LUNAS oleh angsuran bulan ini (angsuran
  // terakhirnya tetap harus tercetak di tagihan bulan itu)
  const lunasBulanIni = {};
  getRowsByFilter(SHEET.ANGSURAN, function(a) {
    return a.status_lock !== STATUS_LOCK.VOID && Number(a.tahun) === tahun &&
           Number(a.bulan) === bulan && String(a.keterangan || '').indexOf('Generate otomatis') === 0;
  }).forEach(function(a) { lunasBulanIni[String(a.id_pinjaman)] = true; });
  const aktifPinjaman = {};
  getRowsByFilter(SHEET.PINJAMAN, function(p) {
    return p.status_lock !== STATUS_LOCK.VOID && !String(p.id_induk || '') &&
      (p.status === 'aktif' || (p.status === 'lunas' && lunasBulanIni[String(p.id_pinjaman)]));
  }).forEach(function(p) {
    const id = String(p.id_anggota);
    (aktifPinjaman[id] = aktifPinjaman[id] || []).push(p);
  });

  const baris = [];
  sheetToObjects(SHEET.ANGGOTA).forEach(function(a) {
    const id = String(a.id_anggota);
    const aktif = a.status === 'aktif';
    const pinjaman = aktifPinjaman[id] || [];
    if (!aktif && pinjaman.length === 0) return;

    const setoranIni = simpananBulanIni.filter(function(r) { return String(r.id_anggota) === id; });
    const wajibSudah = setoranIni.some(function(r) { return r.jenis_simpanan === 'wajib'; });
    const sukarelaSudah = setoranIni.some(function(r) {
      return r.jenis_simpanan === 'sukarela' && String(r.keterangan || '').indexOf(tanda) === 0;
    });
    const pokokSudahIni = setoranIni.some(function(r) { return r.jenis_simpanan === 'pokok'; });

    const angs = pinjaman.map(function(p) { return rincianAngsuranBulan(p, tahun, bulan); })
      .filter(function(x) { return x; });
    const angsPokok = angs.reduce(function(s, x) { return s + x.pokok; }, 0);
    const jasa = angs.reduce(function(s, x) { return s + x.jasa; }, 0);
    const sisa = angs.length ? angs.reduce(function(s, x) { return s + x.sisaSetelah; }, 0)
      : pinjaman.reduce(function(s, p) { return s + getSisaPinjaman(p.id_pinjaman); }, 0);

    const pokok = aktif && (!sudahPokok[id] || pokokSudahIni) ? nominal.nominal_pokok : 0;
    const wajib = aktif ? nominal.nominal_wajib : 0;
    const sukarela = aktif ? Math.round(Number(a.sukarela_rutin) || 0) : 0;

    baris.push({
      id: id, nama: String(a.nama), nip: String(a.nip_nis || ''), status: a.status,
      pokok: pokok, wajib: wajib, sukarela: sukarela,
      angsuranPokok: angsPokok, jasa: jasa,
      jumlah: pokok + wajib + sukarela + angsPokok + jasa,
      sisaPiutang: sisa,
      saldoSukarela: getSaldoSimpanan(id, 'sukarela'),
      jasaSukarela: aktif ? hitungJasaSukarelaAnggota(id, tahun, bulan, setting).nominal : 0,
      pinjamanBaru: angs.some(function(x) { return x.baru; }),
      status_proses: {
        pokok: pokok === 0 || pokokSudahIni,
        wajib: wajib === 0 || wajibSudah,
        sukarela: sukarela === 0 || sukarelaSudah,
        angsuran: angs.every(function(x) { return x.sudah || x.baru || x.pokok + x.jasa === 0; })
      }
    });
  });

  const total = { pokok: 0, wajib: 0, sukarela: 0, angsuranPokok: 0, jasa: 0, jumlah: 0, sisaPiutang: 0 };
  baris.forEach(function(b) { Object.keys(total).forEach(function(k) { total[k] += b[k]; }); });
  baris.forEach(function(b) {
    const s = b.status_proses;
    b.selesai = s.pokok && s.wajib && s.sukarela && s.angsuran;
  });
  return { tahun: tahun, bulan: bulan, namaBulan: NAMA_BULAN_PANJANG[bulan - 1],
    baris: baris, total: total };
}

/**
 * Bukukan potong gaji untuk anggota terpilih. Item yang sudah tercatat dilewati.
 * @return {Object} { diproses: [..], dilewati: n, gagal: [..] }
 */
/**
 * @param {string[]} daftarId       anggota yang dibukukan
 * @param {string[]} [tanpaAngsuran] anggota yang angsurannya TIDAK dipotong bulan ini
 *                                   (hanya simpanan yang dibukukan; angsuran tetap tertagih)
 */
function prosesPotongGaji(tahun, bulan, daftarId, tanpaAngsuran) {
  tahun = Number(tahun); bulan = Number(bulan);
  validateTahunAktif(tahun);
  const tanggal = new Date(tahun, bulan - 1, 10);   // tanggal potong gaji (seperti angsuran)
  const tanda = tandaPotongGaji(tahun, bulan);
  const pilih = {};
  (daftarId || []).forEach(function(id) { pilih[String(id)] = true; });
  const lewatiAngsuran = {};
  (tanpaAngsuran || []).forEach(function(id) { lewatiAngsuran[String(id)] = true; });

  const tagihan = hitungTagihanBulanan(tahun, bulan);
  const hasil = { diproses: [], dilewati: 0, gagal: [] };

  tagihan.baris.forEach(function(b) {
    if (!pilih[b.id]) return;
    if (b.selesai) { hasil.dilewati++; return; }
    const langkah = [];
    const coba = function(label, fn) {
      try { fn(); langkah.push(label); }
      catch (e) { hasil.gagal.push(b.nama + ' — ' + label + ': ' + e.message); }
    };
    if (!b.status_proses.pokok) coba('pokok', function() {
      inputSetoranSimpanan({ id_anggota: b.id, jenis_simpanan: 'pokok', jumlah_setoran: b.pokok,
        tanggal: tanggal, keterangan: tanda });
    });
    if (!b.status_proses.wajib) coba('wajib', function() {
      inputSetoranSimpanan({ id_anggota: b.id, jenis_simpanan: 'wajib', jumlah_setoran: b.wajib,
        tanggal: tanggal, keterangan: tanda });
    });
    if (!b.status_proses.sukarela) coba('sukarela', function() {
      inputSetoranSimpanan({ id_anggota: b.id, jenis_simpanan: 'sukarela', jumlah_setoran: b.sukarela,
        tanggal: tanggal, keterangan: tanda });
    });
    if (!b.status_proses.angsuran && !lewatiAngsuran[b.id]) {
      getPinjamanAktif(b.id).forEach(function(p) {
        const r = rincianAngsuranBulan(p, tahun, bulan);
        if (!r || r.sudah || r.baru || r.pokok + r.jasa === 0) return;
        coba('angsuran ' + p.id_pinjaman, function() { generateSatuAngsuran(p, tahun, bulan); });
      });
    }
    if (lewatiAngsuran[b.id] && !b.status_proses.angsuran) hasil.tanpaAngsuran = (hasil.tanpaAngsuran || 0) + 1;
    if (langkah.length) hasil.diproses.push(b.nama + ': ' + langkah.join(', '));
  });

  logAktivitas('INSERT', SHEET.SIMPANAN, 'POTONG-GAJI-' + tahun + '-' + bulan, null,
    { anggota: Object.keys(pilih).length, diproses: hasil.diproses.length, gagal: hasil.gagal.length });
  return hasil;
}

// ============================================================
// API
// ============================================================

function apiTagihanBulanan(f) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const kini = new Date();
  const t = hitungTagihanBulanan(Number(f && f.tahun) || getTahunAktif() || kini.getFullYear(),
    Number(f && f.bulan) || kini.getMonth() + 1);
  t.profil = getProfilKoperasi();
  return t;
}

function apiProsesPotongGaji(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  const ids = d.ids || [];
  if (ids.length === 0) throw new Error('Pilih minimal satu anggota.');
  if (ids.length > 40) throw new Error('Maksimal 40 anggota per proses (aplikasi memprosesnya bertahap).');
  return prosesPotongGaji(d.tahun, d.bulan, ids, d.tanpaAngsuran || []);
}

/**
 * Batalkan SELURUH pembukuan potong gaji satu bulan: setoran bertanda
 * "Potong gaji <bulan> <tahun>" dan angsuran hasil generate/potong gaji bulan itu.
 * Semua dibatalkan dengan jejak (VOID + jurnal balik), sekaligus atau tidak sama sekali.
 */
function apiBatalkanPotongGaji(d) {
  const profil = requireRole(['admin']);
  MODE_SENYAP = true;
  const tahun = Number(d && d.tahun), bulan = Number(d && d.bulan);
  const alasan = String(d && d.alasan || '').trim();
  if (!tahun || !(bulan >= 1 && bulan <= 12)) throw new Error('Bulan tidak valid.');
  if (!alasan) throw new Error('Alasan pembatalan wajib diisi.');
  validateTahunAktif(tahun);
  const tanda = tandaPotongGaji(tahun, bulan);
  const ket = 'Batal potong gaji ' + NAMA_BULAN_PANJANG[bulan - 1] + ' ' + tahun + ': ' + alasan;

  const setoran = getRowsByFilter(SHEET.SIMPANAN, function(r) {
    return r.status_lock === STATUS_LOCK.OPEN && Number(r.tahun) === tahun && Number(r.bulan) === bulan &&
      String(r.keterangan || '').indexOf(tanda) === 0;
  });
  const angsuran = getRowsByFilter(SHEET.ANGSURAN, function(r) {
    return r.status_lock === STATUS_LOCK.OPEN && Number(r.tahun) === tahun && Number(r.bulan) === bulan &&
      String(r.keterangan || '').indexOf('Generate otomatis') === 0;
  });
  if (!setoran.length && !angsuran.length) {
    throw new Error('Tidak ada pembukuan potong gaji ' + NAMA_BULAN_PANJANG[bulan - 1] + ' ' + tahun + ' yang bisa dibatalkan.');
  }
  const total = setoran.reduce(function(s, r) { return s + (Number(r.jumlah_setoran) || 0); }, 0) +
    angsuran.reduce(function(s, r) { return s + (Number(r.total_bayar) || 0); }, 0);
  const anggota = {};
  setoran.concat(angsuran).forEach(function(r) { anggota[r.id_anggota] = true; });

  mulaiTulisMassal();
  try {
    angsuran.forEach(function(r) { voidAngsuranInti_(r, ket); });
    setoran.forEach(function(r) { voidSetoranInti_(r, ket); });
    selesaiTulisMassal();
  } catch (e) {
    batalTulisMassal();
    throw e;
  }
  logAktivitas('VOID', SHEET.SIMPANAN, 'POTONG-GAJI-' + tahun + '-' + bulan, null,
    { setoran: setoran.length, angsuran: angsuran.length, total: total, alasan: alasan, oleh: profil.email });
  return {
    anggota: Object.keys(anggota).length, setoran: setoran.length, angsuran: angsuran.length, total: total,
    pesan: 'Pembukuan potong gaji ' + NAMA_BULAN_PANJANG[bulan - 1] + ' ' + tahun + ' dibatalkan: ' +
      setoran.length + ' setoran & ' + angsuran.length + ' angsuran dari ' + Object.keys(anggota).length +
      ' anggota (' + formatRupiah(total) + '). Kas & jurnal sudah dibalik; data tetap tercatat sebagai "Dibatalkan".'
  };
}

/**
 * Status proses satu bulan (untuk Beranda "Perlu dikerjakan" & langkah Bulanan).
 * Hanya membaca data.
 */
function apiStatusBulanan(f) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const kini = new Date();
  const tahun = Number(f && f.tahun) || getTahunAktif() || kini.getFullYear();
  const bulan = Number(f && f.bulan) || kini.getMonth() + 1;
  let tagihan = { baris: [] };
  try { tagihan = hitungTagihanBulanan(tahun, bulan); } catch (e) { /* setting tahun belum ada */ }
  const jasa = getRowsByFilter(SHEET.JASA_SUKARELA, function(r) {
    return Number(r.tahun) === tahun && Number(r.bulan) === bulan;
  });
  return {
    tahun: tahun, bulan: bulan, namaBulan: NAMA_BULAN_PANJANG[bulan - 1],
    anggotaDitagih: tagihan.baris.length,
    belumDibukukan: tagihan.baris.filter(function(b) { return !b.selesai; }).length,
    jasaDraft: jasa.filter(function(r) { return r.status_posting === 'DRAFT'; }).length,
    jasaPosted: jasa.filter(function(r) { return r.status_posting === 'POSTED'; }).length,
    pengajuanMenunggu: getDaftarPending().length
  };
}
