/**
 * ============================================================
 * TEST.GS — Fungsi Pembungkus untuk Testing dari Editor
 * ============================================================
 *
 * KENAPA FILE INI ADA:
 * Banyak fungsi di config.gs & database.gs butuh PARAMETER
 * (contoh: getSheet(namaSheet), getSaldoSimpanan(id, jenis)).
 * Tapi dropdown "Run" di editor Apps Script TIDAK BISA
 * mengisi parameter — kalau langsung dijalankan, parameternya
 * jadi "undefined" dan muncul error.
 *
 * Fungsi-fungsi di bawah ini TIDAK punya parameter (aman
 * dijalankan langsung dari dropdown Run), dan di dalamnya
 * baru memanggil fungsi asli dengan argumen yang benar.
 *
 * CARA PAKAI:
 * 1. Pilih salah satu fungsi test_xxx di dropdown atas editor
 * 2. Klik Run (▶️)
 * 3. Lihat hasilnya: menu View → Logs (atau Ctrl+Enter)
 * ============================================================
 */

/** Tes: baca sheet anggota, harus muncul tanpa error */
function test_getSheet() {
  const sheet = getSheet(SHEET.ANGGOTA);
  Logger.log('✅ Berhasil baca sheet: ' + sheet.getName());
}

/** Tes: tahun mana yang sedang aktif */
function test_getTahunAktif() {
  const tahun = getTahunAktif();
  Logger.log('Tahun aktif saat ini: ' + tahun);
}

/** Tes: baca setting_rat tahun 2026 */
function test_getSettingRAT() {
  const setting = getSettingRAT(2026);
  Logger.log(JSON.stringify(setting, null, 2));
}

/** Tes: validasi tahun aktif — coba tahun yang benar */
function test_validateTahunAktif() {
  try {
    const setting = validateTahunAktif(2026);
    Logger.log('✅ Tahun 2026 valid & aktif: ' + JSON.stringify(setting));
  } catch (e) {
    Logger.log('❌ ' + e.message);
  }
}

/** Tes: validasi tahun yang TIDAK ada — harus muncul error yang jelas */
function test_validateTahunAktif_gagal() {
  try {
    validateTahunAktif(1999);
    Logger.log('⚠️ Seharusnya error, tapi tidak error!');
  } catch (e) {
    Logger.log('✅ Berhasil ditolak sesuai harapan: ' + e.message);
  }
}

/** Tes: baca semua data anggota jadi array of objects */
function test_sheetToObjects() {
  const data = sheetToObjects(SHEET.ANGGOTA);
  Logger.log('Jumlah baris anggota: ' + data.length);
  Logger.log(JSON.stringify(data, null, 2));
}

/** Tes: generate ID baru untuk transaksi_simpanan tahun 2026 */
function test_generateId() {
  const id = generateId(SHEET.SIMPANAN, ID_PREFIX.SIMPANAN, 1, 2026, null);
  Logger.log('ID baru yang akan dipakai: ' + id);
}

/** Tes: generate ID anggota baru */
function test_generateIdAnggota() {
  const id = generateIdAnggota();
  Logger.log('ID anggota baru: ' + id);
}

/** Tes: saldo kas real-time */
function test_getSaldoKas() {
  const saldo = getSaldoKas();
  Logger.log('Saldo kas saat ini: ' + formatRupiah(saldo));
}

/**
 * Tes: catat satu baris dummy ke log_aktivitas.
 * Jalankan ini untuk memastikan logAktivitas() berfungsi.
 * Setelah dijalankan, cek sheet log_aktivitas — harus ada baris baru.
 */
function test_logAktivitas() {
  logAktivitas('INSERT', SHEET.ANGGOTA, 'AGT-TEST', null, { nama: 'Test Dummy' });
  Logger.log('✅ Log tercatat. Cek sheet log_aktivitas untuk verifikasi.');
}

/**
 * Tes GABUNGAN: jalankan semua tes sekaligus.
 * Paling praktis — cukup jalankan fungsi ini satu kali untuk
 * memastikan config.gs & database.gs berfungsi semua.
 */
function test_semuaFungsi() {
  Logger.log('===== MULAI TES SEMUA FUNGSI =====');

  Logger.log('\n--- 1. getSheet ---');
  test_getSheet();

  Logger.log('\n--- 2. getTahunAktif ---');
  test_getTahunAktif();

  Logger.log('\n--- 3. getSettingRAT ---');
  test_getSettingRAT();

  Logger.log('\n--- 4. validateTahunAktif (benar) ---');
  test_validateTahunAktif();

  Logger.log('\n--- 5. validateTahunAktif (harus gagal) ---');
  test_validateTahunAktif_gagal();

  Logger.log('\n--- 6. sheetToObjects ---');
  test_sheetToObjects();

  Logger.log('\n--- 7. generateId ---');
  test_generateId();

  Logger.log('\n--- 8. generateIdAnggota ---');
  test_generateIdAnggota();

  Logger.log('\n--- 9. getSaldoKas ---');
  test_getSaldoKas();

  Logger.log('\n===== SELESAI. Jika tidak ada ❌ di atas, semua aman. =====');
}