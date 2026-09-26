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