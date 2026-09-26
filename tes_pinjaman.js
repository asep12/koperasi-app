/**
 * ============================================================
 * TEST_PINJAMAN.GS — Uji Coba Modul Pinjaman
 * ============================================================
 * GANTI 'AGT-0001' / 'AGT-0002' dengan ID anggota yang benar-
 * benar ada di sheet anggota Anda sebelum menjalankan.
 *
 * CATATAN: pinjaman baru butuh saldo kas mencukupi. Jalankan
 * test_semuaFungsiSimpanan (dari test_simpanan.gs) dulu supaya
 * ada kas masuk, baru coba pinjaman.
 * ============================================================
 */

/** Tes: input pinjaman baru normal */
function test_inputPinjamanBaru() {
  try {
    const hasil = inputPinjamanBaru({
      id_anggota: 'AGT-0002',   // ⚠️ ganti — pastikan anggota ini BELUM punya pinjaman aktif
      nominal: 3000000,
      tenor: 10,
      keterangan: 'Tes pinjaman baru'
    });
    Logger.log('✅ ' + hasil.pesan);
    Logger.log('ID pinjaman: ' + hasil.id_pinjaman);
  } catch (e) {
    Logger.log('❌ ' + e.message);
  }
}

/** Tes: tenor melebihi 30 bulan — HARUS gagal */
function test_pinjamanTenorTerlaluLama_gagal() {
  try {
    inputPinjamanBaru({
      id_anggota: 'AGT-0002',
      nominal: 1000000,
      tenor: 36
    });
    Logger.log('⚠️ Seharusnya gagal, tapi malah berhasil!');
  } catch (e) {
    Logger.log('✅ Berhasil ditolak sesuai harapan: ' + e.message);
  }
}

/** Tes: anggota yang masih punya pinjaman aktif ajukan pinjaman lagi — HARUS gagal */
function test_pinjamanDobel_gagal() {
  try {
    // Jalankan dua kali berturut-turut untuk anggota yang sama
    inputPinjamanBaru({ id_anggota: 'AGT-0002', nominal: 500000, tenor: 5 });
    Logger.log('⚠️ Seharusnya gagal (masih ada pinjaman aktif), tapi malah berhasil!');
  } catch (e) {
    Logger.log('✅ Berhasil ditolak sesuai harapan: ' + e.message);
  }
}

/** Tes: cek jurnal pinjaman balance */
function test_cekJurnalPinjaman() {
  const hasil = inputPinjamanBaru({
    id_anggota: 'AGT-0003',   // ⚠️ ganti — anggota lain yang belum punya pinjaman aktif
    nominal: 2000000,
    tenor: 8,
    keterangan: 'Tes cek jurnal pinjaman'
  });

  const jurnal = getRowsByFilter(SHEET.JURNAL, function(row) {
    return row.referensi === hasil.id_pinjaman;
  });
  Logger.log('Baris jurnal: ' + JSON.stringify(jurnal));

  const totalDebit = jurnal.reduce(function(s, j) { return s + Number(j.debit); }, 0);
  const totalKredit = jurnal.reduce(function(s, j) { return s + Number(j.kredit); }, 0);
  Logger.log('Total Debit: ' + totalDebit + ' | Total Kredit: ' + totalKredit +
    (totalDebit === totalKredit ? ' ✅ BALANCE' : ' ❌ TIDAK BALANCE'));

  const sisa = getSisaPinjaman(hasil.id_pinjaman);
  Logger.log('Sisa pinjaman (harus = total_tagihan, karena belum ada angsuran): ' +
    formatRupiah(sisa));
}

/**
 * Tes: input pinjaman MIGRASI (contoh dari dokumen rekap:
 * AGT-0001, sisa Rp 5.000.000, sisa tenor 10 bulan)
 */
function test_inputPinjamanMigrasi() {
  try {
    const hasil = inputPinjamanMigrasi({
      id_anggota: 'AGT-0003',        // ⚠️ ganti sesuai anggota Anda
      sisa_pokok: 5000000,
      tenor_sisa: 10,
      tanggal_asli: new Date(2025, 5, 1), // Juni 2025, contoh
      keterangan_tambahan: 'Data dari Excel Desember 2025'
    });
    Logger.log('✅ ' + hasil.pesan);

    // Pastikan TIDAK ada kas keluar / jurnal baru untuk pinjaman migrasi
    const kas = getRowsByFilter(SHEET.KAS, function(row) {
      return row.referensi === hasil.id_pinjaman;
    });
    const jurnal = getRowsByFilter(SHEET.JURNAL, function(row) {
      return row.referensi === hasil.id_pinjaman;
    });
    Logger.log('Kas terkait (harus KOSONG): ' + kas.length + ' baris');
    Logger.log('Jurnal terkait (harus KOSONG): ' + jurnal.length + ' baris');
  } catch (e) {
    Logger.log('❌ ' + e.message);
  }
}

/** Jalankan semua tes pinjaman sekaligus */
function test_semuaFungsiPinjaman() {
  Logger.log('===== MULAI TES MODUL PINJAMAN =====');
  Logger.log('\n--- 1. Pinjaman baru normal ---');
  test_inputPinjamanBaru();
  Logger.log('\n--- 2. Tenor terlalu lama (harus gagal) ---');
  test_pinjamanTenorTerlaluLama_gagal();
  Logger.log('\n--- 3. Pinjaman dobel untuk anggota yg sama (harus gagal) ---');
  test_pinjamanDobel_gagal();
  Logger.log('\n--- 4. Cek jurnal pinjaman balance ---');
  test_cekJurnalPinjaman();
  Logger.log('\n--- 5. Pinjaman migrasi (tanpa kas/jurnal baru) ---');
  test_inputPinjamanMigrasi();
  Logger.log('\n===== SELESAI. Cek juga langsung ke sheet transaksi_pinjaman, ' +
    'transaksi_kas, dan jurnal_umum untuk verifikasi visual. =====');
}