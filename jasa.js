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
 * @param {Object} [opsi]   { semuaAnggota: true } → termasuk anggota yang KINI keluar/nonaktif
 *                          (dipakai saat menyusulkan bulan lalu: status sekarang belum tentu
 *                          status bulan itu; yang saldonya 0 tetap dilewati)
 */
function hitungJasaSukarelaBulanan(tahun, bulan, opsi) {
  const sekarang = new Date();
  tahun = tahun || getTahunAktif();
  bulan = bulan || (sekarang.getMonth() + 1);

  const setting = validateTahunAktif(tahun);
  pastikanKolomJasaSukarela();
  const persenJasa = Number(setting.jasa_sukarela);
  const metode = metodeJasaSukarela(setting);

  const semua = !!(opsi && opsi.semuaAnggota);
  const anggotaAktif = getRowsByFilter(SHEET.ANGGOTA, function(row) {
    return semua ? !!String(row.id_anggota || '') : row.status === 'aktif';
  });

  const hasil = { dibuat: [], dilewatiSudahAda: [], dilewatiSaldoNol: [], rincian: [] };

  denganTulisMassal(function() {   // semua baris DRAFT ditulis sekaligus di akhir
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
    // atau jasanya dibulatkan jadi Rp 0 (saldo sangat kecil)
    if (j.dasar <= 0 || j.nominal <= 0) {
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
    hasil.rincian.push({ id: String(anggota.id_anggota), nama: String(anggota.nama),
      status: String(anggota.status || ''), nominal: nominalJasa });
  });
  });
  logAktivitas('INSERT', SHEET.JASA_SUKARELA, 'HITUNG-JASA-' + tahun + '-' + bulan, null,
    { dibuat: hasil.dibuat.length });

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

  denganTulisMassal(function() {   // status POSTED & jurnal ditulis sekaligus di akhir
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
  });

  const totalJasa = draftBulanIni.reduce(function(s, r) { return s + Number(r.nominal_jasa); }, 0);
  logAktivitas('EDIT', SHEET.JASA_SUKARELA, 'POSTING-JASA-' + tahun + '-' + bulan,
    { status_posting: 'DRAFT' }, { status_posting: 'POSTED', anggota: draftBulanIni.length, total: totalJasa });

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
// SUSULAN — bulan-bulan yang belum pernah dihitung
// ============================================================

/**
 * Hitung & posting jasa sukarela bulan 1..sampaiBulan yang BELUM punya jasa POSTED,
 * berurutan: jasa Januari diposting dulu sehingga ikut menjadi saldo dasar Februari,
 * dst. (hasilnya sama seperti bila setiap bulan dikerjakan tepat waktu).
 *
 * simpan=false → PRATINJAU: dihitung penuh di memori lalu dibuang, tidak ada yang tertulis.
 * simpan=true  → ditulis sekaligus (atau tidak sama sekali bila ada yang gagal).
 */
function jasaSukarelaSusulan(tahun, sampaiBulan, simpan) {
  tahun = Number(tahun); sampaiBulan = Number(sampaiBulan);
  if (!(sampaiBulan >= 1 && sampaiBulan <= 12)) throw new Error('Bulan tidak valid.');
  validateTahunAktif(tahun);
  pastikanKolomJasaSukarela();

  const hasil = { tahun: tahun, sampaiBulan: sampaiBulan, bulan: [], anggota: [], total: 0, disimpan: !!simpan };
  const perAnggota = {};
  const barisBulan = function(b) {
    return getRowsByFilter(SHEET.JASA_SUKARELA, function(r) {
      return r.status_lock !== STATUS_LOCK.VOID && Number(r.tahun) === tahun && Number(r.bulan) === b;
    });
  };
  mulaiTulisMassal();
  try {
    for (let b = 1; b <= sampaiBulan; b++) {
      const info = { bulan: b, namaBulan: NAMA_BULAN_PANJANG[b - 1], anggota: 0, total: 0 };
      const lama = barisBulan(b);
      if (lama.some(function(r) { return r.status_posting === 'POSTED'; })) {
        info.status = 'sudah';
        info.anggota = lama.length;
        info.total = lama.reduce(function(s, r) { return s + (Number(r.nominal_jasa) || 0); }, 0);
        hasil.bulan.push(info);
        continue;
      }
      hitungJasaSukarelaBulanan(tahun, b, { semuaAnggota: true });
      // semua DRAFT bulan itu (baru dihitung + DRAFT lama bila ada) → rincian, lalu posting
      barisBulan(b).forEach(function(r) {
        const id = String(r.id_anggota), a = getAnggota(id), nominal = Number(r.nominal_jasa) || 0;
        const p = perAnggota[id] = perAnggota[id] ||
          { id: id, nama: a ? String(a.nama) : id, status: a ? String(a.status || '') : '', jasa: {}, total: 0 };
        p.jasa[b] = nominal; p.total += nominal;
        info.anggota++; info.total += nominal;
      });
      if (info.anggota) postingJasaSukarela(tahun, b);
      info.status = info.anggota ? 'baru' : 'kosong';
      hasil.total += info.total;
      hasil.bulan.push(info);
    }
    if (simpan) selesaiTulisMassal(); else batalTulisMassal();
  } catch (e) {
    batalTulisMassal();
    throw e;
  }
  hasil.anggota = Object.keys(perAnggota).map(function(k) { return perAnggota[k]; })
    .sort(function(a, b) { return a.id < b.id ? -1 : 1; });
  if (simpan) {
    logAktivitas('INSERT', SHEET.JASA_SUKARELA, 'SUSULAN-JASA-' + tahun + '-1-' + sampaiBulan, null,
      { bulan: hasil.bulan.filter(function(x) { return x.status === 'baru'; }).length,
        anggota: hasil.anggota.length, total: hasil.total });
  }
  return hasil;
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