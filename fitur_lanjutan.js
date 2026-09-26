/**
 * ============================================================
 * FITUR_LANJUTAN.GS — Kuitansi, Tahun Buku, Grafik, Import
 * Koperasi App
 * ============================================================
 * Endpoint untuk 4 fitur:
 * 5. Kuitansi/bukti setoran (data untuk dicetak dari aplikasi)
 * 6. Pengaturan RAT + Cek Kesiapan + Tutup Buku dari aplikasi
 * 7. Grafik tren bulanan untuk Dashboard
 * 8. Import massal saldo_awal dari CSV
 * ============================================================
 */

// ============================================================
// 5. SETORAN + DATA KUITANSI
// ============================================================

/** Seperti apiSubmitSetoran, tapi mengembalikan data bukti untuk dicetak. */
function apiSetoranDenganBukti(d) {
  const profil = requireRole(['admin', 'operator']);
  MODE_SENYAP = true;

  const hasil = inputSetoranSimpanan({
    id_anggota: d.id_anggota, jenis_simpanan: d.jenis_simpanan,
    jumlah_setoran: Number(d.jumlah), keterangan: d.keterangan || ''
  });

  return { pesan: hasil.pesan, struk: dataStruk('setoran', hasil.id_transaksi) };
}

// ============================================================
// 6. PENGATURAN TAHUN BUKU (setting_rat) + TUTUP BUKU
// ============================================================

function apiSettingList() {
  requireRole(['admin']);
  MODE_SENYAP = true;
  return sheetToObjects(SHEET.SETTING_RAT).map(function(s) {
    const nominal = getNominalSimpanan(Number(s.tahun));
    return {
      nominal_pokok: nominal.nominal_pokok, nominal_wajib: nominal.nominal_wajib,
      minimal_sukarela: nominal.minimal_sukarela,
      tahun: Number(s.tahun),
      jasa_sukarela: Number(s.jasa_sukarela), jasa_pinjaman: Number(s.jasa_pinjaman),
      metode_jasa_sukarela: metodeJasaSukarela(s),
      shu_simpanan: Number(s.shu_simpanan), shu_jasa: Number(s.shu_jasa),
      dana_cadangan: Number(s.dana_cadangan), dana_pengurus: Number(s.dana_pengurus),
      dana_sosial: Number(s.dana_sosial),
      dana_anggota: Number(s.dana_anggota), dana_kesejahteraan: Number(s.dana_kesejahteraan),
      dana_pendidikan: Number(s.dana_pendidikan), dana_pembangunan: Number(s.dana_pembangunan),
      pajak_persen: Number(s.pajak_persen),
      aktif: s.status_aktif === true
    };
  }).sort(function(a, b) { return b.tahun - a.tahun; });
}

/** Tambah/ubah satu baris setting_rat. status_aktif TIDAK diubah di sini. */
function apiSettingSimpan(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;

  const tahun = Number(d.tahun);
  if (!tahun || tahun < 2020 || tahun > 2100) throw new Error('Tahun tidak valid.');

  pastikanKolomAlokasi();
  const calon = {
    jasa_sukarela: Number(d.jasa_sukarela), jasa_pinjaman: Number(d.jasa_pinjaman),
    shu_simpanan: Number(d.shu_simpanan), shu_jasa: Number(d.shu_jasa),
    pajak_persen: Number(d.pajak_persen || 0)
  };
  pastikanKolomJasaSukarela();
  if (d.metode_jasa_sukarela !== undefined) {
    calon.metode_jasa_sukarela = metodeJasaSukarela({ metode_jasa_sukarela: d.metode_jasa_sukarela });
  }
  POS_ALOKASI.forEach(function(p) { calon[p[0]] = Number(d[p[0]]); });
  validateAlokasiSHU(calon); // 7 pos = 100%, split simpanan/jasa = 100

  pastikanKolomNominalSimpanan();
  const lamaNominal = getNominalSimpanan(tahun);
  Object.keys(DEFAULT_NOMINAL_SIMPANAN).forEach(function(k) {
    // Form lama (tanpa isian nominal) → pertahankan nilai yang sudah ada
    const nilai = d[k] === undefined || d[k] === '' ? lamaNominal[k] : Number(d[k]);
    if (!(nilai > 0)) throw new Error('Nominal ' + k.replace(/_/g, ' ') + ' harus lebih dari 0.');
    calon[k] = nilai;
  });

  const ada = getSettingRAT(tahun);
  if (ada) {
    updateRowByField(SHEET.SETTING_RAT, 'tahun', tahun, calon);
    logAktivitas('EDIT', SHEET.SETTING_RAT, String(tahun), null, calon);
    return '✅ Setting tahun ' + tahun + ' diperbarui.';
  }

  calon.tahun = tahun;
  calon.status_aktif = false; // tahun baru selalu nonaktif dulu — diaktifkan via tutup buku
  appendRowFromObject(SHEET.SETTING_RAT, calon);
  logAktivitas('INSERT', SHEET.SETTING_RAT, String(tahun), null, calon);
  return '✅ Setting tahun ' + tahun + ' dibuat (status: belum aktif — akan aktif ' +
    'otomatis saat tutup buku tahun sebelumnya).';
}

function apiCekTutupBuku() {
  requireRole(['admin']);
  MODE_SENYAP = true;
  const tahun = getTahunAktif();
  if (!tahun) return { tahun: null, masalah: ['Tidak ada tahun aktif di setting_rat.'] };
  return { tahun: tahun, masalah: cekKesiapanTutupBuku(tahun) };
}

/** Eksekusi tutup buku. Wajib menyertakan tahun yang diketik pengguna. */
function apiTutupBukuApp(tahunKonfirmasi) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  const tahun = getTahunAktif();
  if (!tahun) throw new Error('Tidak ada tahun aktif.');
  if (Number(tahunKonfirmasi) !== Number(tahun)) {
    throw new Error('Konfirmasi tahun tidak cocok (Anda mengetik "' + tahunKonfirmasi +
      '", tahun aktif adalah ' + tahun + '). Tutup buku dibatalkan.');
  }
  const laporan = prosesTutupBuku(tahun);
  return '✅ TUTUP BUKU TAHUN ' + tahun + ' SELESAI\n\n• ' + laporan.join('\n• ') +
    '\n\nTahun ' + (tahun + 1) + ' sekarang aktif.';
}

// ============================================================
// 7. GRAFIK TREN BULANAN (untuk Dashboard)
// ============================================================

/**
 * Data per bulan tahun aktif:
 * - saldoKas[]  : saldo kas akhir tiap bulan (kumulatif)
 * - setoran[]   : total setoran simpanan per bulan
 */
function apiTrenBulanan() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;

  const tahun = getTahunAktif();
  const kasMasuk = new Array(13).fill(0), kasKeluar = new Array(13).fill(0);
  const setoran = new Array(13).fill(0);

  sheetToObjects(SHEET.KAS).forEach(function(k) {
    if (Number(k.tahun) !== tahun || k.status_lock === STATUS_LOCK.VOID) return;
    const b = Number(k.bulan);
    if (b >= 1 && b <= 12) {
      kasMasuk[b] += Number(k.masuk) || 0;
      kasKeluar[b] += Number(k.keluar) || 0;
    }
  });

  sheetToObjects(SHEET.SIMPANAN).forEach(function(s) {
    if (Number(s.tahun) !== tahun || s.status_lock === STATUS_LOCK.VOID) return;
    const b = Number(s.bulan);
    if (b >= 1 && b <= 12) setoran[b] += Number(s.jumlah_setoran) || 0;
  });

  const saldoKas = [];
  // Mulai dari saldo awal kas migrasi + akumulasi tahun-tahun sebelumnya
  let kumulatif = sumKolom(SHEET.SALDO_AWAL, 'nominal', function(r) {
    return r.jenis === 'kas';
  });
  sheetToObjects(SHEET.KAS).forEach(function(k) {
    if (Number(k.tahun) < tahun && k.status_lock !== STATUS_LOCK.VOID) {
      kumulatif += (Number(k.masuk) || 0) - (Number(k.keluar) || 0);
    }
  });
  for (let b = 1; b <= 12; b++) {
    kumulatif += kasMasuk[b] - kasKeluar[b];
    saldoKas.push(kumulatif);
  }

  return { tahun: tahun, saldoKas: saldoKas, setoran: setoran.slice(1) };
}

// ============================================================
// 8. IMPORT MASSAL SALDO AWAL (dari teks CSV)
// ============================================================

const JENIS_SALDO_VALID = {
  'simpanan_pokok': 'simpanan_pokok', 'pokok': 'simpanan_pokok',
  'simpanan_wajib': 'simpanan_wajib', 'wajib': 'simpanan_wajib',
  'simpanan_sukarela': 'simpanan_sukarela', 'sukarela': 'simpanan_sukarela',
  'piutang': 'piutang',
  // ---- milik KOPERASI (tanpa anggota) — posisi neraca saat mulai memakai aplikasi ----
  'kas': 'kas',
  'inventaris': 'inventaris', 'penyusutan': 'penyusutan',
  'cadangan': 'cadangan', 'modal': 'modal',
  'dana_anggota': 'dana_anggota', 'dana_pengurus': 'dana_pengurus',
  'dana_kesejahteraan': 'dana_kesejahteraan', 'dana_pendidikan': 'dana_pendidikan',
  'dana_sosial': 'dana_sosial', 'dana_pembangunan': 'dana_pembangunan',
  'shu': 'shu', 'ekuitas': 'shu'   // SHU belum dibagi / penyeimbang ekuitas lama
};
/**
 * Angka dari tempelan Excel/CSV: "4.370.495,82" (format Indonesia), "4370495.82",
 * "4,370,495.82", "Rp 1.000.000", "10000". Hasil dibulatkan ke sen.
 */
function angkaImpor(v) {
  if (typeof v === 'number') return Math.round(v * 100) / 100;
  let s = String(v || '').replace(/rp|\s/gi, '');
  if (!s) return 0;
  const koma = s.lastIndexOf(','), titik = s.lastIndexOf('.');
  if (koma > -1 && titik > -1) {
    // pemisah yang paling kanan = desimal
    s = koma > titik ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  } else if (koma > -1) {
    s = /^-?\d{1,3}(,\d{3})+$/.test(s) ? s.replace(/,/g, '') : s.replace(',', '.');
  } else if (titik > -1 && /^-?\d{1,3}(\.\d{3})+$/.test(s)) {
    s = s.replace(/\./g, '');            // 1.000.000 = ribuan
  }
  const n = Number(s);
  return isNaN(n) ? NaN : Math.round(n * 100) / 100;
}

const JENIS_SALDO_ANGGOTA = { simpanan_pokok: 1, simpanan_wajib: 1, simpanan_sukarela: 1, piutang: 1 };

/**
 * @param {Object} d { tahun, baris: [{id_anggota, jenis, nominal}, ...] }
 */
function apiImportSaldoAwal(d) {
  const profil = requireRole(['admin']);
  MODE_SENYAP = true;

  const tahun = Number(d.tahun) || getTahunAktif();
  const baris = d.baris || [];
  if (baris.length === 0) throw new Error('Tidak ada baris untuk diimpor.');
  if (baris.length > 500) throw new Error('Maksimal 500 baris per impor.');

  const hasil = { sukses: 0, gagal: [] };

  // Peta nama → id untuk pencocokan berbasis nama (dari Excel lama)
  const norm = function(t) {
    return String(t || '').toLowerCase().replace(/\s+/g, ' ').trim();
  };
  const petaNama = {};
  sheetToObjects(SHEET.ANGGOTA).forEach(function(a) {
    petaNama[norm(a.nama)] = a.id_anggota;
  });

  baris.forEach(function(r, i) {
    const nomorBaris = i + 1;
    try {
      let idAnggota = String(r.id_anggota || '').trim();
      const jenisRaw = String(r.jenis || '').trim().toLowerCase().replace(/\s+/g, '_');
      const nominal = angkaImpor(r.nominal);

      // Baris milik KOPERASI (kas / shu-ekuitas): tidak terikat anggota
      if (!JENIS_SALDO_VALID[jenisRaw]) throw new Error('jenis "' + r.jenis + '" tidak dikenal ' +
        '(pakai: pokok/wajib/sukarela/piutang, atau untuk koperasi: kas/inventaris/' +
        'penyusutan/cadangan/modal/dana_anggota/dana_pengurus/dana_kesejahteraan/' +
        'dana_pendidikan/dana_sosial/dana_pembangunan/shu)');
      const tanpaAnggota = !JENIS_SALDO_ANGGOTA[JENIS_SALDO_VALID[jenisRaw]];
      if (tanpaAnggota) {
        idAnggota = '-';
      } else if (!getAnggota(idAnggota)) {
        // Cocokkan: coba sebagai ID dulu, lalu sebagai NAMA
        const dariNama = petaNama[norm(idAnggota)];
        if (dariNama) {
          idAnggota = dariNama;
        } else {
          throw new Error('anggota "' + r.id_anggota +
            '" tidak ditemukan (baik sebagai ID maupun nama)');
        }
      }
      const jenis = JENIS_SALDO_VALID[jenisRaw];
      if (!jenis) throw new Error('jenis "' + r.jenis + '" tidak dikenal ' +
        '(pakai: pokok/wajib/sukarela/piutang, atau untuk koperasi: kas/inventaris/' +
        'penyusutan/cadangan/modal/dana_anggota/dana_pengurus/dana_kesejahteraan/' +
        'dana_pendidikan/dana_sosial/dana_pembangunan/shu)');
      if (!nominal || nominal <= 0) throw new Error('nominal tidak valid');

      const idSal = generateId(SHEET.SALDO_AWAL, ID_PREFIX.SALDO_AWAL, 1, tahun, null);
      appendRowFromObject(SHEET.SALDO_AWAL, {
        id: idSal, tahun: tahun, id_anggota: idAnggota,
        jenis: jenis, nominal: nominal,
        keterangan: 'Import CSV oleh ' + profil.email
      });
      hasil.sukses++;
    } catch (e) {
      hasil.gagal.push('Baris ' + nomorBaris + ': ' + e.message);
    }
  });

  logAktivitas('INSERT', SHEET.SALDO_AWAL, 'IMPORT-' + tahun, null,
    { sukses: hasil.sukses, gagal: hasil.gagal.length });
  const cek = cekSaldoAwal();

  return '✅ Import selesai.\nBerhasil: ' + hasil.sukses + ' baris' +
    (hasil.gagal.length > 0
      ? '\nGagal: ' + hasil.gagal.length + '\n' + hasil.gagal.join('\n') +
        '\n\n(Baris yang gagal tidak tersimpan — perbaiki lalu impor ulang ' +
        'HANYA baris tersebut agar tidak dobel)'
      : '') +
    '\n\nKeseimbangan saldo awal: aktiva ' + formatRupiah(cek.aktiva) + ' · pasiva ' +
    formatRupiah(cek.pasiva) + (cek.balance ? ' ✅ BALANCE'
      : ' ❌ selisih ' + formatRupiah(cek.selisih) + ' (tambahkan jenis "shu" sebagai penyeimbang ' +
        'bila ada ekuitas lama yang belum tercatat)') +
    '\n\n⚠️ Untuk jenis "piutang", jangan lupa buat juga pinjaman MIGRASI ' +
    'di form Transaksi → tab Migrasi, lalu cek dengan Validasi Migrasi.';
}

// ============================================================
// 8b. IMPORT MASSAL PINJAMAN MIGRASI
// ============================================================

/**
 * @param {Object} d { baris: [{nama, sisa, tenor, bulan, tahun}, ...] }
 */
function apiImportPinjamanMigrasi(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;

  const baris = d.baris || [];
  if (baris.length === 0) throw new Error('Tidak ada baris untuk diimpor.');
  if (baris.length > 200) throw new Error('Maksimal 200 baris per impor.');

  const norm = function(t) {
    return String(t || '').toLowerCase().replace(/\s+/g, ' ').trim();
  };
  const petaNama = {};
  sheetToObjects(SHEET.ANGGOTA).forEach(function(a) {
    petaNama[norm(a.nama)] = a.id_anggota;
  });

  // Pengaman anti-dobel: anggota yang SUDAH punya pinjaman MIGRASI aktif
  const sudahMigrasi = {};
  sheetToObjects(SHEET.PINJAMAN).forEach(function(p) {
    if (p.status_lock !== STATUS_LOCK.VOID &&
        String(p.keterangan || '').indexOf('MIGRASI') === 0) {
      sudahMigrasi[p.id_anggota] = true;
    }
  });

  const hasil = { sukses: 0, gagal: [] };

  baris.forEach(function(r, i) {
    const nomorBaris = i + 1;
    try {
      let id = String(r.nama || '').trim();
      if (!getAnggota(id)) {
        const dariNama = petaNama[norm(id)];
        if (!dariNama) throw new Error('anggota "' + r.nama + '" tidak ditemukan');
        id = dariNama;
      }
      if (sudahMigrasi[id]) {
        throw new Error('"' + r.nama + '" sudah punya pinjaman MIGRASI aktif — dilewati ' +
          'agar tidak dobel');
      }

      const sisa = angkaImpor(r.sisa);
      const tenor = Number(r.tenor);
      const bulan = Number(r.bulan) || 12;
      const tahun = Number(r.tahun) || (getTahunAktif() - 1);
      if (!sisa || sisa <= 0) throw new Error('sisa pinjaman tidak valid');
      if (!tenor || tenor <= 0) {
        throw new Error('tenor kosong/tidak valid — isi dulu kolom tenor_sisa');
      }

      const angsuran = Number(String(r.angsuran || '').replace(/[.\s]/g, '')) || 0;
      inputPinjamanMigrasi({
        id_anggota: id, sisa_pokok: sisa, tenor_sisa: tenor, angsuran_pokok: angsuran,
        tanggal_asli: new Date(tahun, bulan - 1, 1),
        keterangan_tambahan: 'Import massal'
      });
      sudahMigrasi[id] = true; // jaga-jaga kalau satu nama muncul dua kali di CSV
      hasil.sukses++;
    } catch (e) {
      hasil.gagal.push('Baris ' + nomorBaris + ': ' + e.message);
    }
  });

  logAktivitas('INSERT', SHEET.PINJAMAN, 'IMPORT-MIGRASI', null,
    { sukses: hasil.sukses, gagal: hasil.gagal.length });

  return '✅ Import pinjaman migrasi selesai.\nBerhasil: ' + hasil.sukses + ' pinjaman' +
    (hasil.gagal.length > 0
      ? '\nGagal/dilewati: ' + hasil.gagal.length + '\n' + hasil.gagal.join('\n')
      : '') +
    '\n\nLangkah berikutnya: jalankan Validasi Migrasi (menu 🔧 SETUP di spreadsheet) ' +
    'untuk memastikan total pinjaman MIGRASI = total saldo awal piutang.';
}