/**
 * ============================================================
 * KELUAR_ANGGOTA.GS — Proses Anggota Keluar
 * Koperasi App
 * ============================================================
 *
 * Dipakai saat anggota berhenti (pindah tugas, pensiun, lulus).
 *
 * ALUR (semua dalam satu proses, khusus admin):
 * 1. Rekap: total simpanan (pokok + wajib + sukarela) dan sisa
 *    semua pinjaman aktif → uang bersih = simpanan − sisa pinjaman.
 * 2. Pinjaman aktif DILUNASI dari simpanan: dicatat sebagai
 *    pembayaran angsuran (kas masuk + jurnal, porsi pokok/jasa
 *    sama seperti pembayaran manual).
 * 3. Seluruh simpanan DIKEMBALIKAN: satu baris transaksi_pengambilan
 *    per jenis (status APPROVED) + kas keluar + jurnal
 *    Debit 201/202/203, Kredit 101 Kas.
 * 4. Status anggota → 'keluar'.
 *
 * Kas bersih yang berubah = uang bersih (simpanan − utang).
 * Jika utang > simpanan, anggota harus menyetor kekurangannya
 * (tercatat sebagai bagian dari pelunasan).
 *
 * Salah proses? Setiap baris (angsuran pelunasan & pengambilan)
 * bisa di-VOID satu per satu seperti transaksi biasa, lalu admin
 * mengaktifkan kembali anggota dari daftar anggota.
 * ============================================================
 */

/** Kode akun simpanan per jenis. (Fungsi, bukan konstanta: urutan
 *  pemuatan file Apps Script tidak dijamin, AKUN bisa belum ada.) */
function akunSimpanan(jenis) {
  return jenis === 'pokok' ? AKUN.SIMPANAN_POKOK
       : jenis === 'wajib' ? AKUN.SIMPANAN_WAJIB : AKUN.SIMPANAN_SUKARELA;
}

/**
 * Rekap hak & kewajiban anggota yang akan keluar.
 * @return {Object} { anggota, saldo, pinjaman[], totalSimpanan, totalUtang, netto, masalah[] }
 */
function hitungRekapKeluar(idAnggota) {
  const anggota = getAnggota(idAnggota);
  if (!anggota) throw new Error('Anggota ' + idAnggota + ' tidak ditemukan.');

  const saldo = {
    pokok: getSaldoSimpanan(idAnggota, 'pokok'),
    wajib: getSaldoSimpanan(idAnggota, 'wajib'),
    sukarela: getSaldoSimpanan(idAnggota, 'sukarela')
  };
  const totalSimpanan = saldo.pokok + saldo.wajib + saldo.sukarela;

  const pinjaman = getPinjamanAktif(idAnggota).map(function(p) {
    return { id: p.id_pinjaman, keterangan: String(p.keterangan || ''),
      sisa: getSisaPinjaman(p.id_pinjaman) };
  }).filter(function(p) { return p.sisa > 0; });
  const totalUtang = pinjaman.reduce(function(s, p) { return s + p.sisa; }, 0);

  const masalah = [];
  if (anggota.status === 'keluar') masalah.push('Anggota ini sudah berstatus keluar.');

  const pending = getRowsByFilter(SHEET.PENGAMBILAN, function(r) {
    return String(r.id_anggota) === String(idAnggota) &&
           r.status_approval === 'PENDING' && r.status_lock !== STATUS_LOCK.VOID;
  });
  if (pending.length > 0) {
    masalah.push('Masih ada ' + pending.length + ' pengajuan pengambilan PENDING (' +
      pending.map(function(r) { return r.id_pengambilan; }).join(', ') +
      '). Approve atau tolak dulu di halaman Approval.');
  }

  const draftJasa = getRowsByFilter(SHEET.JASA_SUKARELA, function(r) {
    return String(r.id_anggota) === String(idAnggota) && r.status_posting === 'DRAFT';
  });
  if (draftJasa.length > 0) {
    masalah.push('Ada jasa sukarela berstatus DRAFT untuk anggota ini. ' +
      'Posting jasa bulan ini dulu supaya haknya ikut dihitung.');
  }

  ['pokok', 'wajib', 'sukarela'].forEach(function(jenis) {
    if (saldo[jenis] < 0) {
      masalah.push('Saldo ' + jenis + ' minus (' + formatRupiah(saldo[jenis]) +
        '). Periksa data transaksi anggota ini dulu.');
    }
  });

  return {
    anggota: anggota, saldo: saldo, pinjaman: pinjaman,
    totalSimpanan: totalSimpanan, totalUtang: totalUtang,
    netto: totalSimpanan - totalUtang, masalah: masalah
  };
}

/** Proses keluar lengkap. @return {Object} { pesan, netto } */
function prosesKeluarAnggota(idAnggota, alasan) {
  const tahun = getTahunAktif();
  validateTahunAktif(tahun);

  alasan = String(alasan || '').trim();
  if (!alasan) throw new Error('Alasan keluar wajib diisi (contoh: "Pindah tugas").');

  const rekap = hitungRekapKeluar(idAnggota);
  if (rekap.masalah.length > 0) {
    throw new Error('Proses keluar dibatalkan:\n• ' + rekap.masalah.join('\n• '));
  }

  // Uang yang benar-benar keluar dari kas = netto (pelunasan masuk dulu)
  const saldoKas = getSaldoKas();
  if (rekap.netto > saldoKas) {
    throw new Error('Saldo kas (' + formatRupiah(saldoKas) + ') tidak cukup untuk ' +
      'mengembalikan ' + formatRupiah(rekap.netto) + '.');
  }

  const nama = rekap.anggota.nama;

  // ---------- 1. LUNASI PINJAMAN DARI SIMPANAN ----------
  rekap.pinjaman.forEach(function(p) {
    inputAngsuranManual({
      id_pinjaman: p.id, nominal: p.sisa,
      keterangan: 'Pelunasan dari simpanan — anggota keluar'
    });
  });

  // ---------- 2. KEMBALIKAN SELURUH SIMPANAN ----------
  const tanggal = new Date();
  const bulan = tanggal.getMonth() + 1;
  const idPengembalian = [];

  ['pokok', 'wajib', 'sukarela'].forEach(function(jenis) {
    const jumlah = rekap.saldo[jenis];
    if (!(jumlah > 0)) return;

    const idAmbil = generateId(SHEET.PENGAMBILAN, ID_PREFIX.PENGAMBILAN, 1, tahun, null);
    const row = {
      id_pengambilan: idAmbil,
      tanggal: tanggal, tahun: tahun, bulan: bulan,
      id_anggota: idAnggota,
      jenis_simpanan: jenis,
      jumlah: jumlah,
      status_approval: 'APPROVED',
      keterangan: 'KELUAR ANGGOTA — ' + alasan,
      user_input: getUserEmail(),
      timestamp: new Date(),
      status_lock: STATUS_LOCK.OPEN
    };
    appendRowFromObject(SHEET.PENGAMBILAN, row);

    catatKasKeluar({
      tanggal: tanggal, tahun: tahun, bulan: bulan,
      kategori: 'pengambilan', referensi: idAmbil,
      keterangan: 'Pengembalian simpanan ' + jenis + ' (keluar) — ' + nama,
      nominal: jumlah
    });

    catatJurnal(tanggal, tahun, bulan, idAmbil,
      'Pengembalian simpanan ' + jenis + ' — ' + nama + ' keluar',
      [
        { kode_akun: akunSimpanan(jenis), debit: jumlah, kredit: 0 },
        { kode_akun: AKUN.KAS, debit: 0, kredit: jumlah }
      ]
    );

    logAktivitas('INSERT', SHEET.PENGAMBILAN, idAmbil, null, row);
    idPengembalian.push(idAmbil);
  });

  // ---------- 3. STATUS ANGGOTA ----------
  updateRowByField(SHEET.ANGGOTA, 'id_anggota', idAnggota, { status: 'keluar' });
  logAktivitas('EDIT', SHEET.ANGGOTA, idAnggota, { status: rekap.anggota.status }, {
    status: 'keluar', alasan: alasan,
    total_simpanan: rekap.totalSimpanan, total_utang: rekap.totalUtang,
    netto: rekap.netto, pengembalian: idPengembalian,
    pinjaman_dilunasi: rekap.pinjaman.map(function(p) { return p.id; })
  });

  const baris = [
    '✅ ' + nama + ' resmi keluar.',
    'Total simpanan dikembalikan: ' + formatRupiah(rekap.totalSimpanan),
    'Sisa pinjaman dilunasi: ' + formatRupiah(rekap.totalUtang)
  ];
  baris.push(rekap.netto >= 0
    ? '➡️ Serahkan uang bersih ' + formatRupiah(rekap.netto) + ' kepada anggota.'
    : '➡️ Anggota wajib menyetor kekurangan ' + formatRupiah(-rekap.netto) + '.');

  return { pesan: baris.join('\n'), netto: rekap.netto };
}

// ============================================================
// API UNTUK APLIKASI WEB
// ============================================================

function apiKeluarRekap(idAnggota) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  const r = hitungRekapKeluar(idAnggota);
  return {
    id: r.anggota.id_anggota, nama: r.anggota.nama, status: r.anggota.status,
    saldo: r.saldo, pinjaman: r.pinjaman,
    totalSimpanan: r.totalSimpanan, totalUtang: r.totalUtang,
    netto: r.netto, masalah: r.masalah, saldoKas: getSaldoKas()
  };
}

function apiProsesKeluar(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  return prosesKeluarAnggota(d.id, d.alasan).pesan;
}
