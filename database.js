/**
 * ============================================================
 * DATABASE.GS — Fungsi Baca/Tulis Generik ke Sheet
 * Koperasi App
 * Sesuai dokumen: REKAP_Koperasi_v2.md (BAGIAN 6)
 * ============================================================
 *
 * File ini adalah lapisan "akses data" — semua modul lain
 * (simpanan.gs, pinjaman.gs, dst) memakai fungsi di sini
 * untuk baca/tulis sheet, TIDAK langsung memanggil
 * SpreadsheetApp / getRange di file masing-masing.
 *
 * Keuntungan: kalau nanti ada perubahan struktur sheet,
 * cukup diperbaiki di satu tempat ini.
 * ============================================================
 */

// ============================================================
// AKSES SHEET DASAR
// ============================================================

/** Ambil objek Sheet dari nama. Error jelas jika sheet tidak ada. */
function getSheet(namaSheet) {
  pastikanAksesData(); // lihat sesi.js — cegah panggilan langsung dari browser
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet);
  if (!sheet) {
    throw new Error('Sheet "' + namaSheet + '" tidak ditemukan. ' +
      'Jalankan menu 🔧 SETUP → Jalankan Setup Lengkap dulu.');
  }
  return sheet;
}

/** Ambil array header (baris 1) dari sebuah sheet. */
function getHeader(namaSheet) {
  if (__TULIS_MASSAL && __TULIS_MASSAL.sheet[namaSheet]) return __TULIS_MASSAL.sheet[namaSheet].header.slice();
  const sheet = getSheet(namaSheet);
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

// ============================================================
// BACA DATA — sheet → array of objects
// ============================================================

// ============================================================
// CACHE — sheet hanya dibaca SEKALI per eksekusi
// Tanpa ini, fungsi seperti hitungSHU bisa membaca sheet yang
// sama ratusan kali dan berjalan bermenit-menit.
// Cache otomatis dihapus saat ada penulisan ke sheet tersebut.
// ============================================================

let __CACHE_SHEET = {};

// Cache turunan (peta saldo, sisa pinjaman, data SHU) — dihitung dari
// beberapa sheet sekaligus, jadi ikut dibuang setiap ada penulisan.
let __CACHE_TURUNAN = {};

/** Hapus cache satu sheet (dipanggil otomatis setiap ada penulisan). */
function hapusCache(namaSheet) {
  if (!(__TULIS_MASSAL && __TULIS_MASSAL.sheet[namaSheet])) delete __CACHE_SHEET[namaSheet];
  __CACHE_TURUNAN = {};
}

/** Hapus semua cache (jarang perlu — hanya jika edit sheet manual di tengah eksekusi). */
function hapusSemuaCache() {
  __CACHE_SHEET = {};
  __CACHE_TURUNAN = {};
}

/** Ambil nilai cache turunan; kalau belum ada, hitung sekali lalu simpan. */
function cacheTurunan(kunci, hitungFn) {
  if (!__CACHE_TURUNAN.hasOwnProperty(kunci)) __CACHE_TURUNAN[kunci] = hitungFn();
  return __CACHE_TURUNAN[kunci];
}

/**
 * Baca seluruh isi sheet (kecuali header) jadi array of objects.
 * Setiap objek juga punya properti tersembunyi __row (nomor baris asli
 * di spreadsheet, 1-based) supaya bisa dipakai untuk update/hapus.
 *
 * HASIL DI-CACHE per eksekusi — pembacaan kedua dst dari sheet yang
 * sama langsung dari memori (jauh lebih cepat).
 *
 * Contoh hasil: [{ id_anggota: 'AGT-0001', nama: 'Budi', ..., __row: 2 }, ...]
 */
function sheetToObjects(namaSheet) {
  if (__CACHE_SHEET[namaSheet]) return __CACHE_SHEET[namaSheet];

  const sheet = getSheet(namaSheet);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol === 0) {
    __CACHE_SHEET[namaSheet] = [];
    return __CACHE_SHEET[namaSheet];
  }

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const hasil = data.map(function(row, i) {
    const obj = {};
    header.forEach(function(kolom, j) { obj[kolom] = row[j]; });
    obj.__row = i + 2; // baris asli di sheet (1-based, header = baris 1)
    return obj;
  });

  __CACHE_SHEET[namaSheet] = hasil;
  return hasil;
}

/**
 * Cari SATU baris berdasarkan nilai kolom tertentu.
 * @return {Object|null}
 */
function getRowByField(namaSheet, namaKolom, nilai) {
  const data = sheetToObjects(namaSheet);
  const found = data.find(function(row) {
    return String(row[namaKolom]) === String(nilai);
  });
  return found || null;
}

/**
 * Cari BANYAK baris yang cocok dengan fungsi filter.
 * @param {function(Object): boolean} filterFn
 * @return {Object[]}
 */
function getRowsByFilter(namaSheet, filterFn) {
  return sheetToObjects(namaSheet).filter(filterFn);
}

// ============================================================
// TULIS DATA — objects → sheet
// ============================================================

/**
 * Teks yang diawali = + - @ akan dibaca Google Sheets sebagai RUMUS
 * (mis. keterangan "=IMPORTXML(...)" bisa mengirim data ke luar).
 * Beri tanda petik di depan agar disimpan sebagai teks biasa;
 * saat dibaca kembali, Sheets mengembalikan teks aslinya tanpa petik.
 */
function amanSel_(v) {
  return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v;
}
function amanBaris_(baris) {
  return baris.map(function(r) { return Array.isArray(r) ? r.map(amanSel_) : amanSel_(r); });
}

// ============================================================
// MODE TULIS MASSAL (impor/migrasi ribuan baris dalam satu eksekusi)
// ------------------------------------------------------------
// Selama aktif, baris baru & perubahan ditampung di memori (dan langsung
// terlihat oleh sheetToObjects/generateId), lalu ditulis sekaligus oleh
// selesaiTulisMassal(). Gagal di tengah → batalTulisMassal(): tidak ada
// yang tertulis sama sekali. Fungsi bisnis (setoran, angsuran, dst.)
// tidak perlu diubah.
// ============================================================

let __TULIS_MASSAL = null;

/** Selama mode ini, logAktivitas per baris dilewati (tulis satu log ringkasan sesudahnya). */
function mulaiTulisMassal() {
  if (__TULIS_MASSAL) throw new Error('Mode tulis massal sudah aktif.');
  __TULIS_MASSAL = { sheet: {} };
}

function bufferSheet_(namaSheet) {
  let b = __TULIS_MASSAL.sheet[namaSheet];
  if (b) return b;
  const sheet = getSheet(namaSheet);
  const lastCol = sheet.getLastColumn();
  const header = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  sheetToObjects(namaSheet);                         // muat isi lama ke cache
  b = __TULIS_MASSAL.sheet[namaSheet] = {
    header: header, baris: [], ubah: [], barisAwal: Math.max(sheet.getLastRow(), 1) + 1, nomor: {}
  };
  return b;
}

function tampungBaris_(namaSheet, dataObjek) {
  const b = bufferSheet_(namaSheet);
  const obj = {};
  b.baris.push(b.header.map(function(kolom) {
    const ada = dataObjek.hasOwnProperty(kolom);
    obj[kolom] = ada ? dataObjek[kolom] : '';
    return ada ? amanSel_(dataObjek[kolom]) : '';
  }));
  obj.__row = b.barisAwal + b.baris.length - 1;
  __CACHE_SHEET[namaSheet].push(obj);
  __CACHE_TURUNAN = {};
  return obj.__row;
}

/** Tulis semua tampungan ke sheet. @return {Object} jumlah baris per sheet */
function selesaiTulisMassal() {
  const m = __TULIS_MASSAL;
  if (!m) return {};
  const ringkas = {};
  Object.keys(m.sheet).forEach(function(nama) {
    const b = m.sheet[nama];
    const sheet = getSheet(nama);
    if (b.baris.length) {
      pastikanKapasitasBaris_(sheet, b.barisAwal + b.baris.length - 1);
      sheet.getRange(b.barisAwal, 1, b.baris.length, b.header.length).setValues(b.baris);
      ringkas[nama] = b.baris.length;
    }
    tulisPerubahanMassal_(sheet, nama, b);
  });
  __TULIS_MASSAL = null;
  // Apps Script menunda penulisan ke sheet; paksa sekarang supaya penolakan
  // (mis. aturan validasi sel) muncul DI SINI, bukan setelah langkah dianggap berhasil.
  SpreadsheetApp.flush();
  hapusSemuaCache();
  return ringkas;
}

/**
 * Tulis perubahan sel pada baris LAMA per blok baris berdekatan (satu setValues per blok),
 * bukan satu setValue per sel — mis. membatalkan 150 transaksi: ±10 panggilan, bukan ±500.
 * Sel di dalam blok yang tidak berubah ditulis ulang dengan nilai yang sama dari cache.
 */
function tulisPerubahanMassal_(sheet, namaSheet, b) {
  if (!b.ubah.length) return;
  const data = __CACHE_SHEET[namaSheet] || [];
  const nilai = {};                                   // baris → { kolom: nilai } (perubahan terakhir menang)
  b.ubah.forEach(function(u) { (nilai[u[0]] = nilai[u[0]] || {})[u[1]] = u[2]; });
  const baris = Object.keys(nilai).map(Number).sort(function(x, y) { return x - y; });
  const objBaris = function(r) { const o = data[r - 2]; return o && o.__row === r ? o : null; };

  let i = 0;
  while (i < baris.length) {
    let j = i;
    while (j + 1 < baris.length && baris[j + 1] - baris[j] <= 3) j++;
    const r0 = baris[i], r1 = baris[j];
    let c0 = Infinity, c1 = 0, lengkap = true;
    for (let r = r0; r <= r1; r++) {
      if (!objBaris(r)) lengkap = false;
      Object.keys(nilai[r] || {}).forEach(function(c) { c = Number(c); c0 = Math.min(c0, c); c1 = Math.max(c1, c); });
    }
    if (!lengkap) {
      // cadangan: baris tak ada di cache → tulis sel yang berubah saja
      for (let k = i; k <= j; k++) {
        Object.keys(nilai[baris[k]]).forEach(function(c) {
          sheet.getRange(baris[k], Number(c)).setValue(nilai[baris[k]][c]);
        });
      }
    } else {
      const blok = [];
      for (let r = r0; r <= r1; r++) {
        const o = objBaris(r), ubah = nilai[r] || {}, isi = [];
        for (let c = c0; c <= c1; c++) {
          isi.push(ubah.hasOwnProperty(c) ? ubah[c] : amanSel_(o[b.header[c - 1]]));
        }
        blok.push(isi);
      }
      sheet.getRange(r0, c0, blok.length, c1 - c0 + 1).setValues(blok);
    }
    i = j + 1;
  }
}

/**
 * getRange().setValues() TIDAK menambah baris sendiri (beda dengan appendRow);
 * sheet baru hanya punya 1.000 baris. Tambah baris kosong bila perlu.
 */
function pastikanKapasitasBaris_(sheet, barisTerakhir) {
  const maks = sheet.getMaxRows();
  if (typeof maks === 'number' && barisTerakhir > maks) sheet.insertRowsAfter(maks, barisTerakhir - maks + 100);
}

/**
 * Jalankan fn() dalam mode tulis massal: semua tulisan dikirim sekaligus di akhir,
 * atau tidak sama sekali bila fn() gagal. Bila mode sudah aktif, fn() langsung dijalankan.
 */
function denganTulisMassal(fn) {
  if (__TULIS_MASSAL) return fn();
  mulaiTulisMassal();
  try {
    const hasil = fn();
    selesaiTulisMassal();
    return hasil;
  } catch (e) {
    batalTulisMassal();
    throw e;
  }
}

/** Buang semua tampungan (dipakai saat terjadi error). */
function batalTulisMassal() {
  __TULIS_MASSAL = null;
  hapusSemuaCache();
}

/**
 * Tambah satu baris baru ke sheet berdasarkan objek.
 * Kolom yang tidak ada di objek akan dikosongkan.
 * Urutan diambil otomatis dari header sheet (bukan urutan objek),
 * jadi aman meskipun properti objek ditulis acak.
 *
 * @param {string} namaSheet
 * @param {Object} dataObjek  contoh: { id_anggota: 'AGT-0001', nama: 'Budi' }
 * @return {number} nomor baris tempat data baru disimpan
 */
function appendRowFromObject(namaSheet, dataObjek) {
  if (__TULIS_MASSAL) return tampungBaris_(namaSheet, dataObjek);
  const sheet = getSheet(namaSheet);
  const header = getHeader(namaSheet);

  const baris = header.map(function(kolom) {
    return dataObjek.hasOwnProperty(kolom) ? amanSel_(dataObjek[kolom]) : '';
  });

  sheet.appendRow(baris);
  hapusCache(namaSheet); // data berubah → cache tidak berlaku lagi
  return sheet.getLastRow();
}

/**
 * Tambah BANYAK baris sekaligus (satu kali tulis, jauh lebih cepat
 * daripada appendRow berulang). Aturan kolom sama dengan appendRowFromObject.
 * @param {string} namaSheet
 * @param {Object[]} daftarObjek
 */
function appendRowsFromObjects(namaSheet, daftarObjek) {
  if (!daftarObjek || daftarObjek.length === 0) return;
  if (__TULIS_MASSAL) { daftarObjek.forEach(function(o) { tampungBaris_(namaSheet, o); }); return; }
  const sheet = getSheet(namaSheet);
  const header = getHeader(namaSheet);

  const baris = daftarObjek.map(function(obj) {
    return header.map(function(kolom) {
      return obj.hasOwnProperty(kolom) ? amanSel_(obj[kolom]) : '';
    });
  });

  pastikanKapasitasBaris_(sheet, sheet.getLastRow() + baris.length);
  sheet.getRange(sheet.getLastRow() + 1, 1, baris.length, header.length).setValues(baris);
  hapusCache(namaSheet);
}

/**
 * Update satu baris yang sudah ada, berdasarkan nomor baris (__row).
 * Hanya kolom yang ada di dataBaru yang diubah — kolom lain tidak disentuh.
 *
 * @param {string} namaSheet
 * @param {number} nomorBaris  ambil dari properti __row hasil sheetToObjects
 * @param {Object} dataBaru    contoh: { status: 'nonaktif' }
 */
function updateRowByRowNumber(namaSheet, nomorBaris, dataBaru) {
  if (__TULIS_MASSAL) {
    const b = bufferSheet_(namaSheet);
    const obj = __CACHE_SHEET[namaSheet].find(function(r) { return r.__row === nomorBaris; });
    Object.keys(dataBaru).forEach(function(kolom) {
      const idx = b.header.indexOf(kolom);
      if (idx === -1) throw new Error('Kolom "' + kolom + '" tidak ada di sheet "' + namaSheet + '"');
      if (obj) obj[kolom] = dataBaru[kolom];
      if (nomorBaris >= b.barisAwal) b.baris[nomorBaris - b.barisAwal][idx] = amanSel_(dataBaru[kolom]);
      else b.ubah.push([nomorBaris, idx + 1, amanSel_(dataBaru[kolom])]);
    });
    __CACHE_TURUNAN = {};
    return;
  }
  const sheet = getSheet(namaSheet);
  const header = getHeader(namaSheet);

  Object.keys(dataBaru).forEach(function(kolom) {
    const idx = header.indexOf(kolom);
    if (idx === -1) {
      throw new Error('Kolom "' + kolom + '" tidak ada di sheet "' + namaSheet + '"');
    }
    sheet.getRange(nomorBaris, idx + 1).setValue(amanSel_(dataBaru[kolom]));
  });
  hapusCache(namaSheet); // data berubah → cache tidak berlaku lagi
}

/**
 * Cari satu baris berdasarkan kolom tertentu, lalu update.
 * Gabungan getRowByField + updateRowByRowNumber, lebih ringkas dipakai.
 *
 * @return {boolean} true jika baris ditemukan & diupdate, false jika tidak ketemu
 */
function updateRowByField(namaSheet, namaKolomCari, nilaiCari, dataBaru) {
  const row = getRowByField(namaSheet, namaKolomCari, nilaiCari);
  if (!row) return false;
  updateRowByRowNumber(namaSheet, row.__row, dataBaru);
  return true;
}

// ============================================================
// FUNGSI SUM / AGREGAT — dipakai untuk hitung saldo, dsb
// ============================================================

/**
 * Jumlahkan nilai satu kolom numerik dari baris yang lolos filter.
 * Baris berstatus VOID otomatis diabaikan (jika sheet punya status_lock).
 *
 * @param {string} namaSheet
 * @param {string} kolomAngka   nama kolom yang mau dijumlah
 * @param {function(Object): boolean} filterFn  filter tambahan (opsional)
 * @return {number}
 */
function sumKolom(namaSheet, kolomAngka, filterFn) {
  const data = sheetToObjects(namaSheet);
  return data.reduce(function(total, row) {
    if (row.status_lock === STATUS_LOCK.VOID) return total; // skip VOID
    if (filterFn && !filterFn(row)) return total;
    const nilai = Number(row[kolomAngka]) || 0;
    return total + nilai;
  }, 0);
}

// ============================================================
// LOOKUP UMUM (dipakai lintas modul)
// ============================================================

/** Ambil data satu anggota berdasarkan id_anggota. @return {Object|null} */
function getAnggota(idAnggota) {
  return getRowByField(SHEET.ANGGOTA, 'id_anggota', idAnggota);
}

/** Cek apakah anggota berstatus aktif. Error jika anggota tidak ditemukan. */
function pastikanAnggotaAktif(idAnggota) {
  const anggota = getAnggota(idAnggota);
  if (!anggota) throw new Error('Anggota ' + idAnggota + ' tidak ditemukan.');
  if (anggota.status !== 'aktif') {
    throw new Error('Anggota ' + idAnggota + ' (' + anggota.nama + ') berstatus ' +
      (anggota.status || 'nonaktif') + '.');
  }
  return anggota;
}

/**
 * Saldo kas real-time (BAGIAN 18).
 * @param {Date|null} sampaiTanggal  jika diisi, hanya hitung s/d tanggal ini
 */
function getSaldoKas(sampaiTanggal) {
  const filterFn = sampaiTanggal
    ? function(row) { return new Date(row.tanggal) <= sampaiTanggal; }
    : null;
  const masuk = sumKolom(SHEET.KAS, 'masuk', filterFn);
  const keluar = sumKolom(SHEET.KAS, 'keluar', filterFn);

  // Saldo awal kas hasil migrasi (saldo_awal jenis 'kas').
  // Selalu ikut dihitung karena mendahului semua transaksi.
  // Aman dari dobel: tutup buku TIDAK menulis baris kas baru
  // (transaksi_kas bersifat kumulatif lintas tahun).
  const saldoAwalKas = sumKolom(SHEET.SALDO_AWAL, 'nominal', function(row) {
    return row.jenis === 'kas' && !isSaldoAwalSalinan(row);
  });

  return saldoAwalKas + masuk - keluar;
}

/**
 * Peta total angsuran terbayar per pinjaman: id_pinjaman → rupiah.
 * Dihitung SEKALI per eksekusi (bukan sekali per pinjaman).
 */
function getPetaTotalAngsuran() {
  return cacheTurunan('totalAngsuran', function() {
    const peta = {};
    sheetToObjects(SHEET.ANGSURAN).forEach(function(row) {
      if (row.status_lock === STATUS_LOCK.VOID) return;
      const id = String(row.id_pinjaman);
      peta[id] = (peta[id] || 0) + (Number(row.total_bayar) || 0);
    });
    return peta;
  });
}

/** Peta total POKOK angsuran terbayar per pinjaman (tidak VOID). */
function getPetaAngsuranPokok() {
  return cacheTurunan('angsuranPokok', function() {
    const peta = {};
    sheetToObjects(SHEET.ANGSURAN).forEach(function(row) {
      if (row.status_lock === STATUS_LOCK.VOID) return;
      const id = String(row.id_pinjaman);
      peta[id] = (peta[id] || 0) + (Number(row.angsuran_pokok) || 0);
    });
    return peta;
  });
}

/** Pinjaman memakai jasa menurun (persen × sisa pokok tiap bulan)? */
function isJasaMenurun(pinjaman) {
  return String(pinjaman && pinjaman.metode_jasa || '').trim().toLowerCase() === 'menurun';
}

/**
 * Baris pokok milik satu pinjaman induk: induk + semua pinjaman
 * tambahan (id_induk), yang tidak VOID.
 */
function getBarisPokokPinjaman(idInduk) {
  const id = String(idInduk);
  return getRowsByFilter(SHEET.PINJAMAN, function(r) {
    return r.status_lock !== STATUS_LOCK.VOID &&
           (String(r.id_pinjaman) === id || String(r.id_induk || '') === id);
  });
}

/**
 * Sisa pinjaman real-time (BAGIAN 18).
 * - Jasa MENURUN: sisa POKOK = pokok dicairkan (induk + tambahan)
 *                 − SUM(angsuran_pokok). Jasa ditagih terpisah tiap bulan.
 * - Jasa FLAT (data lama): sisa = total_tagihan − SUM(total_bayar).
 * - Baris pinjaman TAMBAHAN selalu 0 (sisanya ikut induk).
 */
function getSisaPinjaman(idPinjaman) {
  const pinjaman = getRowByField(SHEET.PINJAMAN, 'id_pinjaman', idPinjaman);
  if (!pinjaman) throw new Error('Pinjaman ' + idPinjaman + ' tidak ditemukan.');

  if (isJasaMenurun(pinjaman)) {
    if (String(pinjaman.id_induk || '')) return 0;
    const pokok = getBarisPokokPinjaman(idPinjaman).reduce(function(s, r) {
      return s + (Number(r.nominal) || 0);
    }, 0);
    return pokok - (getPetaAngsuranPokok()[String(idPinjaman)] || 0);
  }

  const totalAngsuran = getPetaTotalAngsuran()[String(idPinjaman)] || 0;
  return Number(pinjaman.total_tagihan) - totalAngsuran;
}

/**
 * Sisa pokok pada AWAL bulan (tahun, bulan) — dasar hitung jasa menurun.
 * Pokok yang dicairkan dalam bulan itu & angsuran bulan itu belum dihitung
 * (sesuai Excel: jasa Januari = 1,5% × sisa per 1 Januari).
 */
function getSisaPokokAwalBulan(idInduk, tahun, bulan) {
  const awal = new Date(tahun, bulan - 1, 1);
  const pokok = getBarisPokokPinjaman(idInduk).reduce(function(s, r) {
    return new Date(r.tanggal) < awal ? s + (Number(r.nominal) || 0) : s;
  }, 0);
  const periode = tahun * 12 + bulan;
  const dibayar = getRowsByFilter(SHEET.ANGSURAN, function(a) {
    return String(a.id_pinjaman) === String(idInduk) && a.status_lock !== STATUS_LOCK.VOID &&
           Number(a.tahun) * 12 + Number(a.bulan) < periode;
  }).reduce(function(s, a) { return s + (Number(a.angsuran_pokok) || 0); }, 0);
  return pokok - dibayar;
}

/** Jasa menurun bulan (tahun, bulan) untuk satu pinjaman induk. */
function hitungJasaMenurun(pinjaman, tahun, bulan) {
  const dasar = Math.max(0, getSisaPokokAwalBulan(pinjaman.id_pinjaman, tahun, bulan));
  return Math.round(dasar * Number(pinjaman.jasa_persen || 0) / 100);
}

/** Angsuran pokok per bulan: dibulatkan KE ATAS ke ribuan (angsuran terakhir lebih kecil). */
function hitungAngsuranPokok(total, tenor) {
  return Math.ceil(Number(total) / Number(tenor) / 1000) * 1000;
}

/**
 * Tambahkan kolom baru ke sheet bila belum ada (sheet lama dibuat
 * sebelum fitur tertentu). Aman dipanggil berulang.
 */
function pastikanKolomTambahan(namaSheet, daftarKolom) {
  const header = getHeader(namaSheet).map(function(h) { return String(h).trim(); });
  const kurang = daftarKolom.filter(function(k) { return header.indexOf(k) === -1; });
  if (kurang.length === 0) return false;
  getSheet(namaSheet).getRange(1, header.length + 1, 1, kurang.length)
    .setValues([kurang]).setFontWeight('bold');
  hapusCache(namaSheet);
  return true;
}

const KOLOM_PINJAMAN_MENURUN = ['metode_jasa', 'id_induk', 'angsuran_sebelumnya'];
function pastikanKolomPinjaman() {
  return pastikanKolomTambahan(SHEET.PINJAMAN, KOLOM_PINJAMAN_MENURUN);
}

const JENIS_SALDO_AWAL_SIMPANAN = {
  simpanan_pokok: 'pokok', simpanan_wajib: 'wajib', simpanan_sukarela: 'sukarela'
};

/**
 * Peta saldo simpanan SEMUA anggota: id_anggota → {pokok, wajib, sukarela}.
 * Setiap sheet dibaca satu kali saja, lalu dipakai berulang oleh
 * dashboard, SHU, rekap, dll. (Dulu setiap panggilan getSaldoSimpanan
 * menyisir ulang 4 sheet → dashboard & SHU sangat lambat.)
 */
function getPetaSaldoSimpanan(sampaiTahun) {
  // sampaiTahun (opsional): saldo per 31 Desember tahun itu (untuk SHU tahun lalu)
  const batas = Number(sampaiTahun) || 0;
  const ikut = function(row) { return !batas || Number(row.tahun) <= batas; };
  return cacheTurunan('saldoSimpanan_' + batas, function() {
    const peta = {};
    const tambah = function(id, jenis, nilai) {
      id = String(id);
      if (!peta[id]) peta[id] = { pokok: 0, wajib: 0, sukarela: 0 };
      if (peta[id].hasOwnProperty(jenis)) peta[id][jenis] += Number(nilai) || 0;
    };

    sheetToObjects(SHEET.SIMPANAN).forEach(function(row) {
      if (row.status_lock === STATUS_LOCK.VOID || !ikut(row)) return;
      tambah(row.id_anggota, row.jenis_simpanan, row.jumlah_setoran);
    });
    sheetToObjects(SHEET.PENGAMBILAN).forEach(function(row) {
      if (row.status_lock === STATUS_LOCK.VOID || !ikut(row)) return;
      if (row.status_approval !== 'APPROVED') return;
      tambah(row.id_anggota, row.jenis_simpanan, -(Number(row.jumlah) || 0));
    });
    sheetToObjects(SHEET.JASA_SUKARELA).forEach(function(row) {
      if (row.status_lock === STATUS_LOCK.VOID || !ikut(row)) return;
      if (row.status_posting !== 'POSTED') return;
      tambah(row.id_anggota, 'sukarela', row.nominal_jasa);
    });
    sheetToObjects(SHEET.SALDO_AWAL).forEach(function(row) {
      if (row.status_lock === STATUS_LOCK.VOID || isSaldoAwalSalinan(row) || !ikut(row)) return;
      const jenis = JENIS_SALDO_AWAL_SIMPANAN[row.jenis];
      if (jenis) tambah(row.id_anggota, jenis, row.nominal);
    });
    return peta;
  });
}

/**
 * Saldo simpanan per anggota per jenis (BAGIAN 18).
 * @param {string} idAnggota
 * @param {string} jenisSimpanan  'pokok' / 'wajib' / 'sukarela'
 */
function getSaldoSimpanan(idAnggota, jenisSimpanan) {
  const saldo = getPetaSaldoSimpanan()[String(idAnggota)];
  return saldo ? (saldo[jenisSimpanan] || 0) : 0;
}

/**
 * Saldo simpanan SUKARELA per anggota, dihitung "sampai akhir bulan
 * sebelum" (tahun, bulan) tertentu — dipakai sebagai basis hitung
 * jasa sukarela (BAGIAN 2A: "basis = saldo akhir bulan sebelumnya").
 *
 * Contoh: getSaldoSukarelaAkhirBulan('AGT-0001', 2026, 3) menghitung
 * saldo sukarela per akhir Februari 2026 (SEBELUM bulan Maret),
 * dipakai untuk hitung jasa sukarela bulan Maret.
 *
 * Saldo migrasi (saldo_awal) selalu dihitung ikut (dianggap sebelum
 * semua periode transaksi sistem berjalan).
 */
function getSaldoSukarelaAkhirBulan(idAnggota, tahun, bulan) {
  const sebelumPeriode = function(row) {
    const t = Number(row.tahun), b = Number(row.bulan);
    return (t < tahun) || (t === tahun && b < bulan);
  };

  const totalSetor = sumKolom(SHEET.SIMPANAN, 'jumlah_setoran', function(row) {
    return row.id_anggota === idAnggota && row.jenis_simpanan === 'sukarela' &&
           sebelumPeriode(row);
  });

  const totalAmbil = sumKolom(SHEET.PENGAMBILAN, 'jumlah', function(row) {
    return row.id_anggota === idAnggota && row.jenis_simpanan === 'sukarela' &&
           row.status_approval === 'APPROVED' && sebelumPeriode(row);
  });

  const totalJasa = sumKolom(SHEET.JASA_SUKARELA, 'nominal_jasa', function(row) {
    return row.id_anggota === idAnggota && row.status_posting === 'POSTED' &&
           sebelumPeriode(row);
  });

  const saldoAwal = sumKolom(SHEET.SALDO_AWAL, 'nominal', function(row) {
    return row.id_anggota === idAnggota && row.jenis === 'simpanan_sukarela' &&
           !isSaldoAwalSalinan(row);
  });

  return saldoAwal + totalSetor + totalJasa - totalAmbil;
}