/**
 * ============================================================
 * JASA.GS — Hitung & Posting Jasa Sukarela Bulanan
 * Koperasi App
 * Sesuai dokumen: REKAP_Koperasi_v2.md (BAGIAN 2A, 3, 12, 15, 17, 18)
 * ============================================================
 *
 * DUA TAHAP TERPISAH (SENGAJA, bukan satu langkah):
 *
 * 1. hitungJasaSukarelaBulanan()
 *    Menghitung jasa untuk semua anggota, status = DRAFT.
 *    Masih BISA dikoreksi/dihapus manual di sheet jika ada yang salah.
 *    BELUM masuk jurnal, BELUM menambah saldo simpanan siapapun.
 *
 * 2. postingJasaSukarela()
 *    Mengunci semua baris DRAFT bulan ini jadi POSTED.
 *    SETELAH ini: jurnal dibuat, saldo sukarela anggota bertambah,
 *    TIDAK BISA diubah lagi tanpa proses VOID.
 *
 * Alur normal (checklist BAGIAN 15, akhir bulan tgl 25-31):
 * [ ] 1. Klik "Hitung Jasa Sukarela Bulan Ini" → cek angkanya di sheet
 * [ ] 2. Kalau sudah benar → Klik "Posting Jasa Sukarela (Kunci)"
 * ============================================================
 */

// ============================================================
// TAHAP 1: HITUNG (DRAFT)
// ============================================================

/**
 * Hitung jasa sukarela untuk SEMUA anggota aktif, bulan & tahun
 * tertentu. Hasilnya berstatus DRAFT — masih boleh dikoreksi.
 *
 * @param {number} [tahun]  default: tahun aktif
 * @param {number} [bulan]  default: bulan berjalan sekarang
 */
function hitungJasaSukarelaBulanan(tahun, bulan) {
  const sekarang = new Date();
  tahun = tahun || getTahunAktif();
  bulan = bulan || (sekarang.getMonth() + 1);

  const setting = validateTahunAktif(tahun);
  pastikanKolomJasaSukarela();
  const persenJasa = Number(setting.jasa_sukarela);
  const metode = metodeJasaSukarela(setting);

  const anggotaAktif = getRowsByFilter(SHEET.ANGGOTA, function(row) {
    return row.status === 'aktif';
  });

  const hasil = { dibuat: [], dilewatiSudahAda: [], dilewatiSaldoNol: [] };

  anggotaAktif.forEach(function(anggota) {
    // Lewati jika bulan ini SUDAH pernah dihitung untuk anggota ini
    const sudahAda = getRowsByFilter(SHEET.JASA_SUKARELA, function(row) {
      return row.id_anggota === anggota.id_anggota &&
             Number(row.tahun) === Number(tahun) &&
             Number(row.bulan) === Number(bulan);
    });
    if (sudahAda.length > 0) {
      hasil.dilewatiSudahAda.push(anggota.nama + ' (' + sudahAda[0].status_posting + ')');
      return;
    }

    const j = hitungJasaSukarelaAnggota(anggota.id_anggota, tahun, bulan, setting);
    const saldoAwalBulan = j.saldoAwal;

    // Aturan: tidak diberikan jika dasar hitung = 0 (anggota baru / saldo sempat habis)
    if (j.dasar <= 0) {
      hasil.dilewatiSaldoNol.push(anggota.nama);
      return;
    }

    const nominalJasa = j.nominal;

    const idJasa = generateId(SHEET.JASA_SUKARELA, ID_PREFIX.JASA_SUKARELA, 1, tahun, bulan);

    const row = {
      id: idJasa,
      tahun: tahun,
      bulan: bulan,
      id_anggota: anggota.id_anggota,
      saldo_awal_bulan: saldoAwalBulan,
      dasar_jasa: j.dasar,
      metode_jasa: metode,
      persen_jasa: persenJasa,
      nominal_jasa: nominalJasa,
      status_posting: 'DRAFT'
    };
    appendRowFromObject(SHEET.JASA_SUKARELA, row);
    logAktivitas('INSERT', SHEET.JASA_SUKARELA, idJasa, null, row);

    hasil.dibuat.push(anggota.nama + ': ' + formatRupiah(nominalJasa));
  });

  const pesan =
    '✅ HITUNG JASA SUKARELA — Tahun ' + tahun + ' Bulan ' + bulan + ' (status: DRAFT)\n' +
    'Rumus: ' + KET_METODE_JASA[metode].replace('{p}', persenJasa) + '\n\n' +
    'Dibuat: ' + hasil.dibuat.length + '\n' +
    (hasil.dibuat.length ? hasil.dibuat.join('\n') + '\n\n' : '\n') +
    'Dilewati (saldo 0 / anggota baru): ' + hasil.dilewatiSaldoNol.length + '\n' +
    'Dilewati (sudah pernah dihitung): ' + hasil.dilewatiSudahAda.length + '\n\n' +
    '➡️ Cek angkanya di sheet rekap_jasa_sukarela. Jika sudah benar, ' +
    'jalankan "Posting Jasa Sukarela (Kunci)" untuk mengunci & membuat jurnal.';

  tampilkanPesan(pesan);
  return hasil;
}

// ============================================================
// RUMUS JASA SUKARELA (bisa berubah tiap RAT → dipilih per tahun buku)
// ============================================================

/**
 * 'terendah' (bawaan, RAT 2026): jasa = saldo sukarela TERENDAH pada bulan itu
 *            × persen/100. Persen dibaca PER BULAN.
 * 'harian'  (aturan lama):      jasa = saldo akhir bulan lalu × persen/100 × 30/365.
 *            Persen dibaca PER TAHUN.
 */
const KET_METODE_JASA = {
  terendah: 'saldo sukarela terendah bulan ini × {p}% per bulan',
  harian: 'saldo akhir bulan lalu × {p}% per tahun × 30/365'
};
function metodeJasaSukarela(setting) {
  const m = String((setting && setting.metode_jasa_sukarela) || '').trim().toLowerCase();
  return KET_METODE_JASA[m] ? m : 'terendah';
}

/** Kolom baru di setting_rat & rekap_jasa_sukarela (sheet lama belum punya). */
function pastikanKolomJasaSukarela() {
  tambahKolomBila_(SHEET.SETTING_RAT, ['metode_jasa_sukarela']);
  tambahKolomBila_(SHEET.JASA_SUKARELA, ['dasar_jasa', 'metode_jasa']);
}
function tambahKolomBila_(namaSheet, kolom) {
  const header = getHeader(namaSheet).map(function(h) { return String(h).trim(); });
  const kurang = kolom.filter(function(k) { return header.indexOf(k) === -1; });
  if (!kurang.length) return;
  getSheet(namaSheet).getRange(1, header.length + 1, 1, kurang.length).setValues([kurang])
    .setFontWeight('bold');
  hapusCache(namaSheet);
}

/**
 * Saldo sukarela TERENDAH selama satu bulan: dimulai dari saldo akhir bulan
 * lalu, lalu saldo di akhir setiap hari yang ada setoran / pengambilan
 * sukarela (yang sudah disetujui). Jasa bulan itu sendiri tidak dihitung.
 */
function getSaldoSukarelaTerendahBulan(idAnggota, tahun, bulan) {
  const awal = getSaldoSukarelaAkhirBulan(idAnggota, tahun, bulan);
  const bulanIni = function(r) {
    return String(r.id_anggota) === String(idAnggota) && r.status_lock !== STATUS_LOCK.VOID &&
      Number(r.tahun) === Number(tahun) && Number(r.bulan) === Number(bulan);
  };
  const hari = function(tgl) {
    const d = new Date(tgl);
    return isNaN(d.getTime()) ? 1 : d.getDate();
  };
  const mutasi = {};   // tanggal → perubahan saldo hari itu
  getRowsByFilter(SHEET.SIMPANAN, function(r) {
    return bulanIni(r) && r.jenis_simpanan === 'sukarela';
  }).forEach(function(r) {
    const h = hari(r.tanggal); mutasi[h] = (mutasi[h] || 0) + (Number(r.jumlah_setoran) || 0);
  });
  getRowsByFilter(SHEET.PENGAMBILAN, function(r) {
    return bulanIni(r) && r.jenis_simpanan === 'sukarela' && r.status_approval === 'APPROVED';
  }).forEach(function(r) {
    const h = hari(r.tanggal); mutasi[h] = (mutasi[h] || 0) - (Number(r.jumlah) || 0);
  });
  let saldo = awal, terendah = awal;
  Object.keys(mutasi).map(Number).sort(function(a, b) { return a - b; }).forEach(function(h) {
    saldo += mutasi[h];
    if (saldo < terendah) terendah = saldo;
  });
  return { saldoAwal: awal, terendah: terendah };
}

/**
 * Jasa sukarela satu anggota untuk satu bulan, sesuai rumus tahun buku itu.
 * @return {{saldoAwal:number, dasar:number, nominal:number, metode:string}}
 */
function hitungJasaSukarelaAnggota(idAnggota, tahun, bulan, setting) {
  setting = setting || getSettingRAT(tahun) || {};
  const persen = Number(setting.jasa_sukarela) || 0;
  const metode = metodeJasaSukarela(setting);
  if (metode === 'harian') {
    const saldoAwal = getSaldoSukarelaAkhirBulan(idAnggota, tahun, bulan);
    return { saldoAwal: saldoAwal, dasar: saldoAwal, metode: metode,
      nominal: saldoAwal > 0 ? Math.round(saldoAwal * persen / 100 * 30 / 365) : 0 };
  }
  const s = getSaldoSukarelaTerendahBulan(idAnggota, tahun, bulan);
  return { saldoAwal: s.saldoAwal, dasar: s.terendah, metode: metode,
    nominal: s.terendah > 0 ? Math.round(s.terendah * persen / 100) : 0 };
}

// ============================================================
// TAHAP 2: POSTING (KUNCI — DRAFT → POSTED)
// ============================================================

/**
 * Kunci semua baris DRAFT bulan & tahun tertentu jadi POSTED.
 * Setelah ini: jurnal dibuat, saldo sukarela anggota bertambah.
 * TIDAK BISA diubah lagi tanpa proses VOID manual.
 *
 * @param {number} [tahun]  default: tahun aktif
 * @param {number} [bulan]  default: bulan berjalan sekarang
 */
function postingJasaSukarela(tahun, bulan) {
  const sekarang = new Date();
  tahun = tahun || getTahunAktif();
  bulan = bulan || (sekarang.getMonth() + 1);

  validateTahunAktif(tahun);

  const draftBulanIni = getRowsByFilter(SHEET.JASA_SUKARELA, function(row) {
    return Number(row.tahun) === Number(tahun) && Number(row.bulan) === Number(bulan) &&
           row.status_posting === 'DRAFT';
  });

  if (draftBulanIni.length === 0) {
    tampilkanPesan('Tidak ada data DRAFT untuk tahun ' + tahun + ' bulan ' + bulan +
      '.\n\nJalankan "Hitung Jasa Sukarela Bulan Ini" dulu sebelum posting, ' +
      'atau mungkin bulan ini sudah pernah di-posting sebelumnya.');
    return { diposting: 0 };
  }

  const tanggalPosting = new Date(tahun, bulan - 1, 28); // akhir bulan, sesuai checklist tgl 25-31

  draftBulanIni.forEach(function(row) {
    const anggota = getAnggota(row.id_anggota);
    const namaAnggota = anggota ? anggota.nama : row.id_anggota;

    // Update status ke POSTED
    updateRowByField(SHEET.JASA_SUKARELA, 'id', row.id, { status_posting: 'POSTED' });

    // Jurnal: Debit Beban Jasa Simpanan Sukarela (biaya umum, seperti Excel
    // "Biaya Operasional"), Kredit Simpanan Sukarela — jasa MENAMBAH saldo
    // sukarela anggota (non-tunai).
    catatJurnal(tanggalPosting, tahun, bulan, row.id,
      'Posting jasa sukarela — ' + namaAnggota,
      [
        { kode_akun: AKUN.BEBAN_JASA_SUKARELA, debit: row.nominal_jasa, kredit: 0 },
        { kode_akun: AKUN.SIMPANAN_SUKARELA, debit: 0, kredit: row.nominal_jasa }
      ]
    );

    logAktivitas('EDIT', SHEET.JASA_SUKARELA, row.id,
      { status_posting: 'DRAFT' }, { status_posting: 'POSTED' });
  });

  const totalJasa = draftBulanIni.reduce(function(s, r) { return s + Number(r.nominal_jasa); }, 0);

  tampilkanPesan(
    '✅ POSTING SELESAI — Tahun ' + tahun + ' Bulan ' + bulan + '\n\n' +
    'Jumlah anggota diposting: ' + draftBulanIni.length + '\n' +
    'Total jasa sukarela: ' + formatRupiah(totalJasa) + '\n\n' +
    '⚠️ Data ini sekarang TERKUNCI. Saldo simpanan sukarela anggota ' +
    'sudah bertambah sesuai jasa masing-masing.'
  );

  return { diposting: draftBulanIni.length, totalJasa: totalJasa };
}

// ============================================================
// PEMBUNGKUS UNTUK MENU (tanpa parameter)
// ============================================================

function menuHitungJasaSukarela() {
  hitungJasaSukarelaBulanan();
}

function menuPostingJasaSukarela() {
  const ui = SpreadsheetApp.getUi();
  const jawab = ui.alert(
    '⚠️ Konfirmasi Posting',
    'Posting akan MENGUNCI data jasa sukarela bulan ini dan membuat jurnal. ' +
    'Setelah ini tidak bisa diubah lagi tanpa proses VOID.\n\nLanjutkan?',
    ui.ButtonSet.YES_NO
  );
  if (jawab !== ui.Button.YES) return;
  postingJasaSukarela();
}