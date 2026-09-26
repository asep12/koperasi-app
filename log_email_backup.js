/**
 * ============================================================
 * LOG_EMAIL_BACKUP.GS — Log Aktivitas, Pengingat Email, Backup
 * Koperasi App
 * ============================================================
 *
 * TIGA FITUR DALAM SATU FILE:
 *
 * 1. apiLogAktivitas — penampil log untuk halaman "Log" di
 *    aplikasi (khusus admin). Data dari sheet log_aktivitas.
 *
 * 2. PENGINGAT EMAIL OTOMATIS — satu trigger harian mengecek
 *    tanggal:
 *      • Tgl 10: jika angsuran bulan ini BELUM digenerate →
 *        email pengingat ke semua admin
 *      • Tgl 25: jika jasa sukarela bulan ini BELUM diposting →
 *        email pengingat ke semua admin
 *
 * 3. BACKUP OTOMATIS — setiap tanggal 1, salinan spreadsheet
 *    disimpan ke folder "Backup Koperasi" di Drive pemilik.
 *
 * ⚙️ AKTIVASI (SEKALI SAJA, dari editor Apps Script):
 *    Pilih fungsi `pasangTugasHarian` di dropdown → Run.
 *    Trigger harian jam 06.00–07.00 akan terpasang otomatis.
 *    Untuk mematikan: jalankan `hapusTugasHarian`.
 * ============================================================
 */

// ============================================================
// 1. PENAMPIL LOG AKTIVITAS (khusus admin)
// ============================================================

/**
 * @param {Object} f { aksi: 'SEMUA'|'INSERT'|..., limit: 50|100|200 }
 */
function apiLogAktivitas(f) {
  requireRole(['admin']);
  MODE_SENYAP = true;

  const aksi = (f && f.aksi) || 'SEMUA';
  const limit = Number((f && f.limit) || 100);

  let rows = sheetToObjects(SHEET.LOG);
  if (aksi !== 'SEMUA') {
    rows = rows.filter(function(r) { return r.aksi === aksi; });
  }

  // Terbaru di atas, potong sesuai limit
  rows = rows.slice(-limit).reverse();

  return rows.map(function(r) {
    let ringkas = '';
    try {
      const d = r.data_baru ? JSON.parse(r.data_baru) : null;
      if (d) {
        ringkas = Object.keys(d).slice(0, 3).map(function(k) {
          return k + ': ' + String(d[k]).substring(0, 30);
        }).join(' · ');
      }
    } catch (e) { ringkas = String(r.data_baru || '').substring(0, 60); }

    return {
      waktu: r.timestamp
        ? Utilities.formatDate(new Date(r.timestamp),
            Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
        : '',
      user: String(r.user || '').split('@')[0],
      aksi: r.aksi,
      sheet: r.sheet_target,
      ref: r.id_referensi,
      ringkas: ringkas
    };
  });
}

// ============================================================
// 2 & 3. TUGAS HARIAN: PENGINGAT + BACKUP
// ============================================================

/** Folder Drive tempat cadangan (bisa diatur di identitas_lokal.js → folder_backup). */
function namaFolderBackup_() {
  const lokal = typeof identitasLokal_ === 'function' ? identitasLokal_() : {};
  return lokal.folder_backup || 'Backup Koperasi ' + getProfilKoperasi().nama_koperasi.replace(/"/g, '');
}

/** Dipanggil trigger setiap pagi. Mengecek tanggal & bertindak. */
function tugasHarian(e) {
  // Trigger waktu mengirim event berisi triggerUid milik proyek ini.
  // Panggilan lain (mis. dari browser) wajib lolos penjaga akses.
  const dariTrigger = !!(e && e.triggerUid) && ScriptApp.getProjectTriggers()
    .some(function(t) { return String(t.getUniqueId()) === String(e.triggerUid); });
  if (dariTrigger) __AKSES_PUBLIK = true; else pastikanAksesData();

  const kini = new Date();
  const tanggal = kini.getDate();

  try {
    bersihkanSesiKedaluwarsa();   // buang token login yang sudah lewat masa
    if (tanggal === 1) backupSpreadsheet();
    if (tanggal === 10) pengingatAngsuran();
    if (tanggal === 25) pengingatJasa();
  } catch (e) {
    // Kirim error ke pemilik supaya tidak diam-diam gagal
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
      '⚠️ Tugas harian koperasi GAGAL',
      'Terjadi error pada tugas harian tanggal ' + kini + ':\n\n' + e.message);
  }
}

/** Email semua admin aktif. */
function kirimKeAdmin(judul, isi) {
  pastikanAksesData();
  pastikanSheetPengguna();
  const adminList = sheetToObjects(SHEET_PENGGUNA).filter(function(u) {
    return u.peran === 'admin' && u.status === 'aktif';
  });
  if (adminList.length === 0) return;

  adminList.forEach(function(u) {
    MailApp.sendEmail(String(u.email), judul,
      'Assalamu\u2019alaikum / Selamat pagi,\n\n' + isi +
      '\n\n—\nPesan otomatis dari aplikasi Koperasi ' + getProfilKoperasi().nama_koperasi.replace(/"/g, '') + '.');
  });
}

/** Tgl 10: ingatkan generate angsuran JIKA belum dilakukan bulan ini. */
function pengingatAngsuran() {
  const kini = new Date();
  const tahun = kini.getFullYear(), bulan = kini.getMonth() + 1;

  const sudah = getRowsByFilter(SHEET.ANGSURAN, function(r) {
    return Number(r.tahun) === tahun && Number(r.bulan) === bulan &&
           r.status_lock !== STATUS_LOCK.VOID;
  });
  if (sudah.length > 0) return; // sudah digenerate → tidak perlu mengingatkan

  const pinjamanAktif = getRowsByFilter(SHEET.PINJAMAN, function(p) {
    return p.status === 'aktif' && p.status_lock !== STATUS_LOCK.VOID;
  });
  if (pinjamanAktif.length === 0) return; // tidak ada yang perlu ditagih

  kirimKeAdmin('🔔 Pengingat: Generate Angsuran bulan ' + bulan + '/' + tahun,
    'Hari ini tanggal 10 — waktunya proses angsuran bulanan.\n\n' +
    'Angsuran bulan ' + bulan + '/' + tahun + ' BELUM digenerate, padahal ada ' +
    pinjamanAktif.length + ' pinjaman aktif.\n\n' +
    'Buka aplikasi → menu Proses Bulanan → "Generate Angsuran Bulan Ini", ' +
    'lalu cetak struk potongan untuk bendahara gaji.');
}

/** Tgl 25: ingatkan hitung+posting jasa JIKA belum diposting bulan ini. */
function pengingatJasa() {
  const kini = new Date();
  const tahun = kini.getFullYear(), bulan = kini.getMonth() + 1;

  const sudahPosted = getRowsByFilter(SHEET.JASA_SUKARELA, function(r) {
    return Number(r.tahun) === tahun && Number(r.bulan) === bulan &&
           r.status_posting === 'POSTED';
  });
  if (sudahPosted.length > 0) return;

  const masihDraft = getRowsByFilter(SHEET.JASA_SUKARELA, function(r) {
    return Number(r.tahun) === tahun && Number(r.bulan) === bulan &&
           r.status_posting === 'DRAFT';
  });

  kirimKeAdmin('🔔 Pengingat: Jasa Sukarela bulan ' + bulan + '/' + tahun,
    'Hari ini tanggal 25 — waktunya proses jasa sukarela akhir bulan.\n\n' +
    (masihDraft.length > 0
      ? 'Ada ' + masihDraft.length + ' baris jasa berstatus DRAFT yang belum ' +
        'diposting. Periksa angkanya lalu klik "Posting Jasa (Kunci)".'
      : 'Jasa sukarela bulan ini belum dihitung. Buka aplikasi → Proses Bulanan → ' +
        '"Hitung Jasa Sukarela", periksa hasilnya, lalu posting.'));
}

/** Tgl 1: salin spreadsheet ke folder backup di Drive. */
function backupSpreadsheet() {
  pastikanAksesData();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const fileAsli = DriveApp.getFileById(ss.getId());

  // Cari / buat folder backup
  let folder;
  const cari = DriveApp.getFoldersByName(namaFolderBackup_());
  folder = cari.hasNext() ? cari.next() : DriveApp.createFolder(namaFolderBackup_());

  const stempel = Utilities.formatDate(new Date(),
    Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const namaBackup = 'BACKUP ' + stempel + ' — ' + ss.getName();

  fileAsli.makeCopy(namaBackup, folder);

  // Rapikan: simpan maksimal 12 backup terakhir, hapus yang lebih tua
  const semua = [];
  const iter = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (iter.hasNext()) semua.push(iter.next());
  semua.sort(function(a, b) { return b.getDateCreated() - a.getDateCreated(); });
  for (let i = 12; i < semua.length; i++) semua[i].setTrashed(true);

  kirimKeAdmin('💾 Backup bulanan berhasil — ' + stempel,
    'Salinan spreadsheet koperasi berhasil dibuat:\n"' + namaBackup + '"\n\n' +
    'Lokasi: folder Drive "' + namaFolderBackup_() + '" milik pemilik sistem.\n' +
    'Sistem menyimpan 12 backup terakhir; yang lebih lama dihapus otomatis.');
}

// ============================================================
// PASANG / HAPUS TRIGGER (jalankan dari editor, sekali saja)
// ============================================================

function pasangTugasHarian() {
  pastikanAksesData();
  // Hindari trigger dobel
  hapusTugasHarian();

  ScriptApp.newTrigger('tugasHarian')
    .timeBased()
    .everyDays(1)
    .atHour(6)     // antara jam 06.00–07.00
    .create();

  Logger.log('✅ Trigger harian terpasang (jam 06.00–07.00).');
  Logger.log('Tgl 1: backup | Tgl 10: pengingat angsuran | Tgl 25: pengingat jasa.');
}

function hapusTugasHarian() {
  pastikanAksesData();
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'tugasHarian') ScriptApp.deleteTrigger(t);
  });
  Logger.log('Trigger tugasHarian (jika ada) sudah dihapus.');
}

/** Tes manual dari editor: paksa kirim pengingat & backup sekarang. */
function test_tugasHarian_paksa() {
  Logger.log('Menguji backup…'); backupSpreadsheet();
  Logger.log('Menguji pengingat angsuran…'); pengingatAngsuran();
  Logger.log('Menguji pengingat jasa…'); pengingatJasa();
  Logger.log('✅ Selesai. Cek email admin & folder Drive "' + namaFolderBackup_() + '".');
}