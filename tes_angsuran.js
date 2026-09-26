/**
 * ============================================================
 * TEST_ANGSURAN.GS — Uji Coba Modul Angsuran
 * ============================================================
 * SYARAT: harus sudah ada minimal 1 pinjaman aktif.
 * Jalankan test_semuaFungsiPinjaman (dari test_pinjaman.gs)
 * dulu jika belum ada pinjaman sama sekali.
 * ============================================================
 */

/** Tes: generate angsuran untuk bulan & tahun berjalan */
function test_generateAngsuranBulanan() {
  const hasil = generateAngsuranBulanan(); // default: tahun aktif, bulan sekarang
  Logger.log('Berhasil: ' + hasil.berhasil.length);
  Logger.log('Dilewati: ' + hasil.dilewati.length);
  Logger.log('Gagal: ' + hasil.gagal.length);
  Logger.log(JSON.stringify(hasil, null, 2));
}

/** Tes: generate ulang di bulan yang SAMA — semua harus DILEWATI (bukan dobel) */
function test_generateAngsuranDobel_harusDilewati() {
  const hasilPertama = generateAngsuranBulanan();
  const hasilKedua = generateAngsuranBulanan();

  Logger.log('Generate pertama — berhasil: ' + hasilPertama.berhasil.length);
  Logger.log('Generate kedua (harus semua dilewati) — dilewati: ' +
    hasilKedua.dilewati.length + ', berhasil: ' + hasilKedua.berhasil.length);

  if (hasilKedua.berhasil.length === 0) {
    Logger.log('✅ Berhasil: tidak ada dobel generate.');
  } else {
    Logger.log('❌ MASALAH: ada baris yang dobel ter-generate!');
  }
}

/** Tes: cek jurnal angsuran balance untuk satu ID angsuran */
function test_cekJurnalAngsuran() {
  const hasil = generateAngsuranBulanan();
  if (hasil.berhasil.length === 0) {
    Logger.log('Tidak ada angsuran baru untuk dicek (mungkin sudah pernah digenerate).');
    return;
  }

  // Ambil ID angsuran dari hasil pertama, format "PJM-xxx → ANG-xxx"
  const idAngsuran = hasil.berhasil[0].split(' → ')[1];
  const jurnal = getRowsByFilter(SHEET.JURNAL, function(row) {
    return row.referensi === idAngsuran;
  });
  Logger.log('Jurnal untuk ' + idAngsuran + ': ' + JSON.stringify(jurnal));

  const totalDebit = jurnal.reduce(function(s, j) { return s + Number(j.debit); }, 0);
  const totalKredit = jurnal.reduce(function(s, j) { return s + Number(j.kredit); }, 0);
  Logger.log('Total Debit: ' + totalDebit + ' | Total Kredit: ' + totalKredit +
    (totalDebit === totalKredit ? ' ✅ BALANCE' : ' ❌ TIDAK BALANCE'));
}

/**
 * Tes SIMULASI PELUNASAN: generate angsuran berulang kali dengan
 * bulan berbeda-beda sampai satu pinjaman lunas, lalu cek statusnya
 * otomatis berubah jadi 'lunas' dan sisa jadi PERSIS 0.
 *
 * ⚠️ Ini akan membuat banyak baris angsuran dummy (sejumlah tenor).
 * Gunakan hanya di spreadsheet TESTING, bukan data asli.
 */
function test_simulasiPelunasanPenuh() {
  // Cari satu pinjaman aktif dengan tenor pendek untuk disimulasikan
  const pinjamanAktif = getRowsByFilter(SHEET.PINJAMAN, function(row) {
    return row.status === 'aktif';
  });

  if (pinjamanAktif.length === 0) {
    Logger.log('Tidak ada pinjaman aktif untuk disimulasikan.');
    return;
  }

  const target = pinjamanAktif[0];
  Logger.log('Simulasi pelunasan untuk: ' + target.id_pinjaman +
    ' (tenor ' + target.tenor + ' bulan, total tagihan ' +
    formatRupiah(target.total_tagihan) + ')');

  const tahunMulai = Number(target.tahun);
  let bulanJalan = Number(target.bulan) + 1;
  let tahunJalan = tahunMulai;

  for (let i = 0; i < Number(target.tenor) + 1; i++) {
    if (bulanJalan > 12) { bulanJalan = 1; tahunJalan++; }

    const sisaSaatIni = getSisaPinjaman(target.id_pinjaman);
    if (sisaSaatIni <= 0) {
      Logger.log('✅ Lunas setelah ' + i + ' kali angsuran. Sisa akhir: ' +
        formatRupiah(sisaSaatIni) + ' (harus 0)');
      break;
    }

    try {
      const idAngsuran = generateSatuAngsuran(target, tahunJalan, bulanJalan);
      Logger.log('  Bulan ' + bulanJalan + '/' + tahunJalan + ' → ' + idAngsuran +
        ' | sisa sesudah: ' + formatRupiah(getSisaPinjaman(target.id_pinjaman)));
    } catch (e) {
      Logger.log('  Bulan ' + bulanJalan + '/' + tahunJalan + ' → ❌ ' + e.message);
      break;
    }

    bulanJalan++;
  }

  const pinjamanAkhir = getRowByField(SHEET.PINJAMAN, 'id_pinjaman', target.id_pinjaman);
  Logger.log('Status akhir pinjaman: ' + pinjamanAkhir.status +
    (pinjamanAkhir.status === 'lunas' ? ' ✅' : ' ❌ (seharusnya lunas)'));
}

/** Tes: cetak struk potongan bulan berjalan */
function test_cetakStrukPotongan() {
  const hasil = cetakStrukPotongan();
  Logger.log('Jumlah anggota di struk: ' + hasil.jumlahAnggota);
  Logger.log('Total potongan: ' + formatRupiah(hasil.totalPotongan));
}

/** Jalankan tes utama (bukan simulasi pelunasan, karena itu bikin banyak data) */
function test_semuaFungsiAngsuran() {
  Logger.log('===== MULAI TES MODUL ANGSURAN =====');
  Logger.log('\n--- 1. Generate angsuran bulan ini ---');
  test_generateAngsuranBulanan();
  Logger.log('\n--- 2. Generate dobel (harus dilewati) ---');
  test_generateAngsuranDobel_harusDilewati();
  Logger.log('\n--- 3. Cetak struk potongan ---');
  test_cetakStrukPotongan();
  Logger.log('\n===== SELESAI. Cek sheet transaksi_angsuran, transaksi_kas, ' +
    'jurnal_umum, dan struk_potongan untuk verifikasi visual. =====\n' +
    'Untuk uji pelunasan penuh sampai status "lunas", jalankan ' +
    'test_simulasiPelunasanPenuh() secara terpisah.');
}