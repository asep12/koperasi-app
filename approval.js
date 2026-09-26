/**
 * ============================================================
 * APPROVAL.GS — Approve/Reject Pengambilan Simpanan Sukarela
 * Koperasi App
 * Sesuai dokumen: REKAP_Koperasi_v2.md (BAGIAN 3, 12, 17 — keputusan v2 #7)
 * ============================================================
 *
 * KEPUTUSAN v2: status_approval TIDAK BOLEH diubah manual di sel.
 * Harus lewat menu ini, supaya:
 * 1. Tercatat di log_aktivitas SIAPA yang approve/reject + kapan
 * 2. Saldo divalidasi ULANG saat approve (bukan cuma saat pengajuan
 *    — saldo bisa berubah di antara pengajuan dan approval)
 * 3. Kas keluar + jurnal hanya dibuat saat benar-benar APPROVED
 *
 * ALUR:
 * PENDING (dari inputPengajuanPengambilan di simpanan.gs)
 *    ├── APPROVED → kas keluar + jurnal + log APPROVE
 *    └── REJECTED → hanya ubah status + log REJECT (tanpa kas/jurnal)
 * ============================================================
 */

// ============================================================
// FUNGSI INTI (bisa dipanggil dari test/editor dengan parameter)
// ============================================================

/**
 * APPROVE satu pengajuan pengambilan.
 * @param {string} idPengambilan  contoh: 'AMB-2026-0001'
 * @return {Object} { pesan }
 */
function approvePengambilan(idPengambilan) {
  const row = getRowByField(SHEET.PENGAMBILAN, 'id_pengambilan', idPengambilan);
  if (!row) throw new Error('Pengajuan ' + idPengambilan + ' tidak ditemukan.');

  if (row.status_approval !== 'PENDING') {
    throw new Error('Pengajuan ' + idPengambilan + ' berstatus ' + row.status_approval +
      ' — hanya yang PENDING yang bisa di-approve.');
  }
  if (row.status_lock === STATUS_LOCK.VOID) {
    throw new Error('Pengajuan ini sudah di-VOID.');
  }

  validateTahunAktif(Number(row.tahun));
  const anggota = pastikanAnggotaAktif(row.id_anggota);

  // VALIDASI ULANG SALDO — saldo bisa berubah sejak pengajuan dibuat
  const saldoSukarela = getSaldoSimpanan(row.id_anggota, 'sukarela');
  if (Number(row.jumlah) > saldoSukarela) {
    throw new Error('Saldo sukarela ' + anggota.nama + ' saat ini (' +
      formatRupiah(saldoSukarela) + ') tidak cukup untuk pengambilan ' +
      formatRupiah(row.jumlah) + '. Pengajuan TIDAK di-approve.');
  }

  // VALIDASI KAS — kas koperasi harus cukup untuk bayar tunai
  const saldoKas = getSaldoKas();
  if (Number(row.jumlah) > saldoKas) {
    throw new Error('Saldo kas koperasi (' + formatRupiah(saldoKas) +
      ') tidak cukup untuk pengambilan ' + formatRupiah(row.jumlah) + '.');
  }

  // ---------- UBAH STATUS ----------
  updateRowByField(SHEET.PENGAMBILAN, 'id_pengambilan', idPengambilan,
    { status_approval: 'APPROVED' });

  // ---------- KAS KELUAR ----------
  const tanggal = new Date();
  catatKasKeluar({
    tanggal: tanggal,
    tahun: Number(row.tahun),
    bulan: Number(row.bulan),
    kategori: 'pengambilan',
    referensi: idPengambilan,
    keterangan: 'Pengambilan sukarela — ' + anggota.nama,
    nominal: Number(row.jumlah)
  });

  // ---------- JURNAL (BAGIAN 12 #6) ----------
  // Debit Simpanan Sukarela, Kredit Kas
  catatJurnal(tanggal, Number(row.tahun), Number(row.bulan), idPengambilan,
    'Pengambilan simpanan sukarela — ' + anggota.nama,
    [
      { kode_akun: AKUN.SIMPANAN_SUKARELA, debit: Number(row.jumlah), kredit: 0 },
      { kode_akun: AKUN.KAS, debit: 0, kredit: Number(row.jumlah) }
    ]
  );

  // ---------- LOG APPROVE ----------
  logAktivitas('APPROVE', SHEET.PENGAMBILAN, idPengambilan,
    { status_approval: 'PENDING' },
    { status_approval: 'APPROVED', approver: getUserEmail() });

  return {
    pesan: '✅ APPROVED — Pengambilan ' + formatRupiah(row.jumlah) + ' untuk ' +
      anggota.nama + ' disetujui.\nKas keluar & jurnal sudah dibuat otomatis.\n' +
      'Sisa saldo sukarela anggota: ' +
      formatRupiah(getSaldoSimpanan(row.id_anggota, 'sukarela'))
  };
}

/**
 * REJECT satu pengajuan pengambilan.
 * Tidak ada kas/jurnal — hanya ubah status + catat log.
 * @param {string} idPengambilan
 * @param {string} [alasan]
 */
function rejectPengambilan(idPengambilan, alasan) {
  const row = getRowByField(SHEET.PENGAMBILAN, 'id_pengambilan', idPengambilan);
  if (!row) throw new Error('Pengajuan ' + idPengambilan + ' tidak ditemukan.');

  if (row.status_approval !== 'PENDING') {
    throw new Error('Pengajuan ' + idPengambilan + ' berstatus ' + row.status_approval +
      ' — hanya yang PENDING yang bisa di-reject.');
  }

  updateRowByField(SHEET.PENGAMBILAN, 'id_pengambilan', idPengambilan, {
    status_approval: 'REJECTED',
    keterangan: (row.keterangan ? row.keterangan + ' | ' : '') +
      'DITOLAK: ' + (alasan || 'tanpa alasan tertulis')
  });

  logAktivitas('REJECT', SHEET.PENGAMBILAN, idPengambilan,
    { status_approval: 'PENDING' },
    { status_approval: 'REJECTED', approver: getUserEmail(), alasan: alasan || '' });

  const anggota = getAnggota(row.id_anggota);
  return {
    pesan: '🚫 REJECTED — Pengajuan ' + formatRupiah(row.jumlah) + ' untuk ' +
      (anggota ? anggota.nama : row.id_anggota) + ' ditolak.'
  };
}

/** Ambil daftar semua pengajuan yang masih PENDING. */
function getDaftarPending() {
  return getRowsByFilter(SHEET.PENGAMBILAN, function(row) {
    return row.status_approval === 'PENDING' && row.status_lock !== STATUS_LOCK.VOID;
  });
}

// ============================================================
// MENU INTERAKTIF (dipanggil dari menu spreadsheet)
// ============================================================

/**
 * Menu: tampilkan daftar PENDING, minta admin pilih ID,
 * lalu pilih Approve / Reject.
 */
function menuApprovalPengambilan() {
  const ui = SpreadsheetApp.getUi();
  const pendingList = getDaftarPending();

  if (pendingList.length === 0) {
    ui.alert('Tidak ada pengajuan pengambilan yang berstatus PENDING. ✅');
    return;
  }

  // Susun daftar untuk ditampilkan
  const daftar = pendingList.map(function(row, i) {
    const anggota = getAnggota(row.id_anggota);
    return (i + 1) + '. ' + row.id_pengambilan + ' — ' +
      (anggota ? anggota.nama : row.id_anggota) + ' — ' +
      formatRupiah(row.jumlah) +
      (row.keterangan ? ' (' + row.keterangan + ')' : '');
  }).join('\n');

  const respId = ui.prompt(
    '📋 Pengajuan PENDING (' + pendingList.length + ')',
    daftar + '\n\nKetik ID pengajuan yang mau diproses (contoh: ' +
    pendingList[0].id_pengambilan + '):',
    ui.ButtonSet.OK_CANCEL
  );
  if (respId.getSelectedButton() !== ui.Button.OK) return;

  const idPilihan = respId.getResponseText().trim();
  const target = pendingList.find(function(r) { return r.id_pengambilan === idPilihan; });
  if (!target) {
    ui.alert('ID "' + idPilihan + '" tidak ada di daftar PENDING. Proses dibatalkan.');
    return;
  }

  const anggota = getAnggota(target.id_anggota);
  const respAksi = ui.alert(
    'Proses ' + idPilihan,
    'Pengambilan ' + formatRupiah(target.jumlah) + ' oleh ' +
    (anggota ? anggota.nama : target.id_anggota) + '\n\n' +
    'YES = APPROVE (setujui, kas keluar)\n' +
    'NO  = REJECT (tolak)\n' +
    'CANCEL = batal, tidak melakukan apa-apa',
    ui.ButtonSet.YES_NO_CANCEL
  );

  try {
    if (respAksi === ui.Button.YES) {
      const hasil = approvePengambilan(idPilihan);
      ui.alert(hasil.pesan);
    } else if (respAksi === ui.Button.NO) {
      const respAlasan = ui.prompt('Alasan penolakan (opsional):', ui.ButtonSet.OK);
      const hasil = rejectPengambilan(idPilihan, respAlasan.getResponseText());
      ui.alert(hasil.pesan);
    }
    // CANCEL: tidak melakukan apa-apa
  } catch (e) {
    ui.alert('❌ GAGAL: ' + e.message);
  }
}