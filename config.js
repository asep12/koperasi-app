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