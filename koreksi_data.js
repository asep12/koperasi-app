/**
 * ============================================================
 * KOREKSI_DATA.GS — Koreksi Jurnal Jasa Pinjaman (data lama)
 * Koperasi App
 * ============================================================
 *
 * MASALAH YANG DIKOREKSI:
 * Versi lama pinjaman.js mencatat pencairan sebagai
 *   Debit Piutang (total_tagihan) / Kredit Kas (nominal)
 *                                 / Kredit Pendapatan Jasa (total_jasa)
 * padahal jasa juga diakui lagi setiap angsuran dibayar.
 * Akibatnya pendapatan jasa pinjaman DOBEL dan piutang tersisa
 * sebesar total_jasa meskipun pinjaman sudah lunas.
 *
 * KOREKSI: untuk setiap pinjaman yang terkena, dibuat jurnal
 *   Debit 401 Pendapatan Jasa Pinjaman / Kredit 103 Piutang
 * sebesar total_jasa, dengan referensi = id_pinjaman (supaya kalau
 * pinjaman itu kelak di-VOID, jurnal koreksinya ikut dibalik).
 *
 * CARA PAKAI (dari editor Apps Script):
 * 1. Jalankan koreksiJasaPinjaman_cek      → lihat daftar di Logs,
 *    TIDAK menulis apa pun.
 * 2. Jika daftarnya benar, jalankan koreksiJasaPinjaman_jalankan.
 * Aman dijalankan berulang: pinjaman yang sudah dikoreksi dilewati.
 *
 * Pinjaman dari tahun yang SUDAH DITUTUP tidak dikoreksi otomatis
 * (SHU-nya sudah dibagi) — hanya dilaporkan, putuskan manual.
 * ============================================================
 */

const PENANDA_KOREKSI_JASA = 'KOREKSI jasa dobel';

function koreksiJasaPinjaman_cek() {
  return koreksiJasaPinjaman(true);
}

function koreksiJasaPinjaman_jalankan() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return koreksiJasaPinjaman(false);
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

/**
 * @param {boolean} simulasi  true = hanya laporan, tanpa menulis
 * @return {Object} { dikoreksi: [], dilewati: [], tahunTertutup: [] }
 */
function koreksiJasaPinjaman(simulasi) {
  MODE_SENYAP = true;
  const hasil = { dikoreksi: [], dilewati: [], tahunTertutup: [] };

  const pinjamanList = getRowsByFilter(SHEET.PINJAMAN, function(p) {
    return p.status_lock !== STATUS_LOCK.VOID && Number(p.total_jasa) > 0;
  });

  pinjamanList.forEach(function(p) {
    const id = String(p.id_pinjaman);
    const jurnalPinjaman = getRowsByFilter(SHEET.JURNAL, function(j) {
      return String(j.referensi) === id;
    });

    // Terkena bug hanya jika jurnal pencairannya punya Kredit 401
    const kreditJasa = jurnalPinjaman.reduce(function(s, j) {
      return String(j.kode_akun) === AKUN.PENDAPATAN_JASA_PINJAMAN
        ? s + (Number(j.kredit) || 0) : s;
    }, 0);
    const sudahDikoreksi = jurnalPinjaman.some(function(j) {
      return String(j.keterangan).indexOf(PENANDA_KOREKSI_JASA) === 0;
    });

    if (kreditJasa <= 0 || sudahDikoreksi) {
      hasil.dilewati.push(id + (sudahDikoreksi ? ' (sudah dikoreksi)' : ' (format baru)'));
      return;
    }

    const tahun = Number(p.tahun);
    const setting = getSettingRAT(tahun);
    if (!setting || setting.status_aktif !== true) {
      hasil.tahunTertutup.push(id + ' — tahun ' + tahun + ', jasa ' + formatRupiah(kreditJasa));
      return;
    }

    if (!simulasi) {
      const kini = new Date();
      catatJurnal(kini, tahun, kini.getMonth() + 1, id,
        PENANDA_KOREKSI_JASA + ' — pencairan ' + id,
        [
          { kode_akun: AKUN.PENDAPATAN_JASA_PINJAMAN, debit: kreditJasa, kredit: 0 },
          { kode_akun: AKUN.PIUTANG_PINJAMAN, debit: 0, kredit: kreditJasa }
        ]
      );
      logAktivitas('EDIT', SHEET.JURNAL, id, null,
        { koreksi: 'jasa pinjaman dobel', nominal: kreditJasa });
    }
    hasil.dikoreksi.push(id + ' — ' + formatRupiah(kreditJasa));
  });

  Logger.log((simulasi ? '🔍 SIMULASI (tidak ada yang ditulis)' : '✅ KOREKSI SELESAI') +
    '\n\nDikoreksi: ' + hasil.dikoreksi.length + '\n' + hasil.dikoreksi.join('\n') +
    '\n\nTahun sudah ditutup (tidak dikoreksi): ' + hasil.tahunTertutup.length + '\n' +
    hasil.tahunTertutup.join('\n') +
    '\n\nDilewati: ' + hasil.dilewati.length);
  return hasil;
}

// ============================================================
// KONVERSI PINJAMAN AKTIF: JASA FLAT → JASA MENURUN
// ============================================================
//
// Keputusan koperasi: jasa pinjaman = persen × SISA POKOK tiap bulan
// (seperti Excel pembukuan). Pinjaman aktif yang dibuat sebelum itu
// masih memakai hitungan flat. Konversi ini:
//   - sisa pokok  = nominal − SUM(angsuran_pokok yang sudah dibayar)
//   - angsuran pokok/bulan = porsi pokok dari angsuran lama
//     (untuk pinjaman MIGRASI: angsuran lama memang sudah pokok)
//   - jasa_persen: pinjaman biasa tetap persen lamanya;
//     pinjaman MIGRASI (dulu 0%) memakai jasa_pinjaman tahun aktif
//   - sisa jasa flat yang belum ditagih DIHAPUS; mulai bulan depan
//     jasa dihitung dari sisa pokok.
// Jurnal tidak berubah (piutang sudah berbasis pokok sejak koreksi).
//
// CARA PAKAI (dari editor Apps Script):
// 1. Pastikan jasa_pinjaman tahun aktif di Pengaturan sudah benar
//    (Excel: 1,5 %/bulan).
// 2. Jalankan koreksiJasaPinjaman_cek / _jalankan dulu (jika ada).
// 3. Jalankan konversiPinjamanMenurun_cek → baca Logs.
// 4. Jika benar, jalankan konversiPinjamanMenurun_jalankan.

function konversiPinjamanMenurun_cek() {
  return konversiPinjamanMenurun(true);
}

function konversiPinjamanMenurun_jalankan() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return konversiPinjamanMenurun(false);
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function konversiPinjamanMenurun(simulasi) {
  MODE_SENYAP = true;
  if (!simulasi) pastikanKolomPinjaman();
  const hasil = { dikonversi: [], ditahan: [] };

  const tahunAktif = getTahunAktif();
  const setting = tahunAktif ? getSettingRAT(tahunAktif) : null;
  const jasaSetting = setting ? Number(setting.jasa_pinjaman) : 0;

  const daftar = getRowsByFilter(SHEET.PINJAMAN, function(p) {
    return p.status === 'aktif' && p.status_lock !== STATUS_LOCK.VOID && !isJasaMenurun(p);
  });

  daftar.forEach(function(p) {
    const id = String(p.id_pinjaman);

    // Pinjaman yang terkena jasa dobel tapi belum dikoreksi → tahan dulu
    const jurnal = getRowsByFilter(SHEET.JURNAL, function(j) { return String(j.referensi) === id; });
    const kreditJasa = jurnal.some(function(j) {
      return String(j.kode_akun) === AKUN.PENDAPATAN_JASA_PINJAMAN && Number(j.kredit) > 0;
    });
    const sudahDikoreksi = jurnal.some(function(j) {
      return String(j.keterangan).indexOf(PENANDA_KOREKSI_JASA) === 0;
    });
    if (kreditJasa && !sudahDikoreksi) {
      hasil.ditahan.push(id + ' — jalankan koreksiJasaPinjaman dulu');
      return;
    }

    const angsuran = getRowsByFilter(SHEET.ANGSURAN, function(a) {
      return String(a.id_pinjaman) === id && a.status_lock !== STATUS_LOCK.VOID;
    });
    const pokokDibayar = angsuran.reduce(function(s, a) { return s + (Number(a.angsuran_pokok) || 0); }, 0);
    const jasaDibayar = angsuran.reduce(function(s, a) { return s + (Number(a.jasa) || 0); }, 0);
    const nominal = Number(p.nominal) || 0;
    const sisaPokok = nominal - pokokDibayar;
    const sisaLama = getSisaPinjaman(id);   // hitungan flat (pokok + jasa)

    const migrasi = String(p.keterangan || '').indexOf('MIGRASI') === 0;
    const rasio = Number(p.total_tagihan) > 0 ? nominal / Number(p.total_tagihan) : 1;
    const angsuranPokok = Math.max(1, Math.round(Number(p.angsuran_perbulan) * rasio));
    const jasaPersen = migrasi || !(Number(p.jasa_persen) > 0) ? jasaSetting : Number(p.jasa_persen);

    const perubahan = {
      metode_jasa: 'menurun',
      jasa_persen: jasaPersen,
      total_tagihan: nominal,
      total_jasa: jasaDibayar,
      angsuran_perbulan: angsuranPokok
    };
    if (sisaPokok <= 0) perubahan.status = 'lunas';

    if (!simulasi) {
      updateRowByRowNumber(SHEET.PINJAMAN, p.__row, perubahan);
      logAktivitas('EDIT', SHEET.PINJAMAN, id, {
        metode_jasa: 'flat', jasa_persen: p.jasa_persen, total_tagihan: p.total_tagihan,
        total_jasa: p.total_jasa, angsuran_perbulan: p.angsuran_perbulan
      }, Object.assign({ konversi: 'jasa menurun' }, perubahan));
    }
    hasil.dikonversi.push(id + (migrasi ? ' (migrasi)' : '') +
      ' — sisa lama ' + formatRupiah(sisaLama) + ' → sisa pokok ' + formatRupiah(sisaPokok) +
      ', pokok ' + formatRupiah(angsuranPokok) + '/bln, jasa ' + jasaPersen + '%/bln' +
      ' (±' + formatRupiah(Math.round(Math.max(0, sisaPokok) * jasaPersen / 100)) + ' bln depan)');
  });

  Logger.log((simulasi ? '🔍 SIMULASI konversi (tidak ada yang ditulis)' : '✅ KONVERSI SELESAI') +
    '\n\nJasa pinjaman tahun aktif di Pengaturan: ' + jasaSetting + '%/bulan' +
    '\n\nDikonversi: ' + hasil.dikonversi.length + '\n' + hasil.dikonversi.join('\n') +
    '\n\nDitahan: ' + hasil.ditahan.length + '\n' + hasil.ditahan.join('\n'));
  return hasil;
}
