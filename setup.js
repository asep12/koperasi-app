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