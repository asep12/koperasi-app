/**
 * ============================================================
 * TEST_APPROVAL.GS — Uji Coba Modul Approval Pengambilan
 * ============================================================
 * GANTI 'AGT-0001' dengan ID anggota yang punya saldo sukarela.
 * ============================================================
 */

/** Tes lengkap: ajukan → approve → cek kas & jurnal & saldo */
function test_alurApprovePengambilan() {
  MODE_SENYAP = true;

  const idAnggota = 'AGT-0003';  // ⚠️ ganti — anggota dengan saldo sukarela cukup
  const jumlah = 50000;

  const saldoSebelum = getSaldoSimpanan(idAnggota, 'sukarela');
  const kasSebelum = getSaldoKas();
  Logger.log('Saldo sukarela sebelum: ' + formatRupiah(saldoSebelum));
  Logger.log('Kas sebelum: ' + formatRupiah(kasSebelum));

  // 1. Ajukan
  const pengajuan = inputPengajuanPengambilan({
    id_anggota: idAnggota,
    jumlah: jumlah,
    keterangan: 'Tes alur approval'
  });
  Logger.log('1️⃣ ' + pengajuan.pesan);

  // 2. Approve
  const hasil = approvePengambilan(pengajuan.id_pengambilan);
  Logger.log('2️⃣ ' + hasil.pesan);

  // 3. Verifikasi
  const saldoSesudah = getSaldoSimpanan(idAnggota, 'sukarela');
  const kasSesudah = getSaldoKas();
  Logger.log('Saldo sukarela sesudah: ' + formatRupiah(saldoSesudah) +
    (saldoSesudah === saldoSebelum - jumlah ? ' ✅' : ' ❌ tidak sesuai'));
  Logger.log('Kas sesudah: ' + formatRupiah(kasSesudah) +
    (kasSesudah === kasSebelum - jumlah ? ' ✅' : ' ❌ tidak sesuai'));

  // 4. Cek jurnal balance
  const jurnal = getRowsByFilter(SHEET.JURNAL, function(j) {
    return j.referensi === pengajuan.id_pengambilan;
  });
  const totalDebit = jurnal.reduce(function(s, j) { return s + Number(j.debit); }, 0);
  const totalKredit = jurnal.reduce(function(s, j) { return s + Number(j.kredit); }, 0);
  Logger.log('Jurnal — Debit: ' + totalDebit + ', Kredit: ' + totalKredit +
    (totalDebit === totalKredit ? ' ✅ BALANCE' : ' ❌ TIDAK BALANCE'));
}

/** Tes: reject pengajuan — tidak boleh ada kas/jurnal terbentuk */
function test_alurRejectPengambilan() {
  MODE_SENYAP = true;

  const idAnggota = 'AGT-0003';  // ⚠️ ganti
  const kasSebelum = getSaldoKas();

  const pengajuan = inputPengajuanPengambilan({
    id_anggota: idAnggota,
    jumlah: 30000,
    keterangan: 'Tes alur reject'
  });
  Logger.log('1️⃣ ' + pengajuan.pesan);

  const hasil = rejectPengambilan(pengajuan.id_pengambilan, 'Tes penolakan');
  Logger.log('2️⃣ ' + hasil.pesan);

  const kasSesudah = getSaldoKas();
  const kas = getRowsByFilter(SHEET.KAS, function(row) {
    return row.referensi === pengajuan.id_pengambilan;
  });
  Logger.log('Kas tidak berubah: ' + (kasSesudah === kasSebelum ? '✅' : '❌'));
  Logger.log('Baris kas terkait (harus 0): ' + kas.length +
    (kas.length === 0 ? ' ✅' : ' ❌'));
}

/** Tes: approve dobel — kedua kali HARUS gagal */
function test_approveDobel_gagal() {
  MODE_SENYAP = true;

  const pengajuan = inputPengajuanPengambilan({
    id_anggota: 'AGT-0003',  // ⚠️ ganti
    jumlah: 20000,
    keterangan: 'Tes approve dobel'
  });
  approvePengambilan(pengajuan.id_pengambilan);
  Logger.log('Approve pertama: ✅');

  try {
    approvePengambilan(pengajuan.id_pengambilan);
    Logger.log('⚠️ Approve kedua seharusnya gagal, tapi malah berhasil!');
  } catch (e) {
    Logger.log('✅ Approve kedua berhasil ditolak: ' + e.message);
  }
}

/** Tes: pengambilan melebihi saldo — approve HARUS gagal */
function test_approveMelebihiSaldo_gagal() {
  MODE_SENYAP = true;

  const idAnggota = 'AGT-0003';  // ⚠️ ganti
  const saldo = getSaldoSimpanan(idAnggota, 'sukarela');

  try {
    // Pengajuan yang melebihi saldo seharusnya sudah ditolak di input,
    // jadi tes ini memastikan validasi lapis pertama juga bekerja
    inputPengajuanPengambilan({
      id_anggota: idAnggota,
      jumlah: saldo + 1000000,
      keterangan: 'Tes melebihi saldo'
    });
    Logger.log('⚠️ Pengajuan melebihi saldo seharusnya ditolak, tapi lolos!');
  } catch (e) {
    Logger.log('✅ Pengajuan melebihi saldo berhasil ditolak: ' + e.message);
  }
}

/** Jalankan semua tes approval */
function test_semuaFungsiApproval() {
  MODE_SENYAP = true;
  Logger.log('===== MULAI TES MODUL APPROVAL =====');
  Logger.log('\n--- 1. Alur approve lengkap ---');
  test_alurApprovePengambilan();
  Logger.log('\n--- 2. Alur reject ---');
  test_alurRejectPengambilan();
  Logger.log('\n--- 3. Approve dobel (harus gagal) ---');
  test_approveDobel_gagal();
  Logger.log('\n--- 4. Pengajuan melebihi saldo (harus gagal) ---');
  test_approveMelebihiSaldo_gagal();
  Logger.log('\n===== SELESAI. Cek log_aktivitas — harus ada baris APPROVE ' +
    'dan REJECT dengan email approver. =====');
}