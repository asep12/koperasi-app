/**
 * ============================================================
 * TAGIHAN.GS — Tagihan Bulanan (Potong Gaji) & Slip Tagihan
 * Koperasi App
 * ============================================================
 *
 * Pengganti sheet "Tag (1..12)" + "Str (1..12)" di Excel:
 * satu layar untuk melihat tagihan semua anggota, mencetak
 * Daftar Tagihan & Slip Tagihan, lalu membukukan potong gaji.
 *
 * Isi tagihan per anggota (bulan, tahun):
 *   1. Simpanan Pokok   — hanya jika belum pernah setor pokok
 *   2. Simpanan Wajib   — nominal dari Pengaturan tahun itu
 *   3. Simpanan Sukarela— setoran rutin anggota (kolom anggota.sukarela_rutin)
 *   4. Angsuran Pinjaman (pokok)
 *   5. Jasa Pinjaman    — jasa menurun: persen × sisa pokok awal bulan
 * + informasi: sisa piutang setelah bayar, saldo sukarela, dan jasa
 *   sukarela bulan ini (menambah saldo, TIDAK dipotong dari gaji —
 *   di Excel lama jasa ini tercampur ke baris sukarela).
 *
 * PROSES POTONG GAJI membukukan item yang belum tercatat untuk anggota
 * yang dipilih. Aman diulang: yang sudah tercatat dilewati.
 * ============================================================
 */

const NAMA_BULAN_PANJANG = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

function pastikanKolomAnggota() {
  return pastikanKolomTambahan(SHEET.ANGGOTA, ['sukarela_rutin']);
}

function tandaPotongGaji(tahun, bulan) {
  return 'Potong gaji ' + NAMA_BULAN_PANJANG[bulan - 1] + ' ' + tahun;
}

/** Rincian angsuran bulan (tahun, bulan) untuk satu pinjaman induk aktif. */
function rincianAngsuranBulan(p, tahun, bulan) {
  const idP = String(p.id_pinjaman);
  const sisa = getSisaPinjaman(idP);
  const sudah = getRowsByFilter(SHEET.ANGSURAN, function(a) {
    return String(a.id_pinjaman) === idP && a.status_lock !== STATUS_LOCK.VOID &&
           Number(a.tahun) === Number(tahun) && Number(a.bulan) === Number(bulan) &&
           String(a.keterangan || '').indexOf('Generate otomatis') === 0;
  })[0];
  if (sudah) {
    return { id: idP, pokok: Number(sudah.angsuran_pokok) || 0, jasa: Number(sudah.jasa) || 0,
      sisaSetelah: sisa, sudah: true };
  }
  if (sisa <= 0) return null;
  if (isJasaMenurun(p)) {
    if (new Date(p.tanggal) >= new Date(tahun, bulan - 1, 1)) {
      return { id: idP, pokok: 0, jasa: 0, sisaSetelah: sisa, sudah: false, baru: true };
    }
    const pokok = Math.min(Number(p.angsuran_perbulan) || 0, sisa);
    return { id: idP, pokok: pokok, jasa: hitungJasaMenurun(p, tahun, bulan),
      sisaSetelah: sisa - pokok, sudah: false };
  }
  // data lama (jasa flat)
  const total = Math.min(Number(p.angsuran_perbulan) || 0, sisa);
  const rasio = Number(p.total_tagihan) > 0 ? Number(p.nominal) / Number(p.total_tagihan) : 1;
  const pokok = Math.round(total * rasio);
  return { id: idP, pokok: pokok, jasa: total - pokok, sisaSetelah: sisa - total, sudah: false };
}

/**
 * Tagihan semua anggota aktif (+ anggota nonaktif yang masih punya pinjaman).
 * @return {Object} { tahun, bulan, namaBulan, baris: [...], total: {...} }
 */
function hitungTagihanBulanan(tahun, bulan) {
  tahun = Number(tahun); bulan = Number(bulan);
  if (!(bulan >= 1 && bulan <= 12)) throw new Error('Bulan tidak valid.');
  const setting = getSettingRAT(tahun);
  if (!setting) throw new Error('Setting tahun ' + tahun + ' belum ada.');
  const nominal = getNominalSimpanan(tahun);
  const tanda = tandaPotongGaji(tahun, bulan);

  const simpananBulanIni = getRowsByFilter(SHEET.SIMPANAN, function(r) {
    return r.status_lock !== STATUS_LOCK.VOID && Number(r.tahun) === tahun && Number(r.bulan) === bulan;
  });
  const sudahPokok = {};
  getRowsByFilter(SHEET.SIMPANAN, function(r) {
    return r.status_lock !== STATUS_LOCK.VOID && r.jenis_simpanan === 'pokok';
  }).forEach(function(r) { sudahPokok[String(r.id_anggota)] = true; });
  sheetToObjects(SHEET.SALDO_AWAL).forEach(function(r) {
    if (r.jenis === 'simpanan_pokok' && !isSaldoAwalSalinan(r)) sudahPokok[String(r.id_anggota)] = true;
  });

  // Pinjaman aktif + pinjaman yang LUNAS oleh angsuran bulan ini (angsuran
  // terakhirnya tetap harus tercetak di tagihan bulan itu)
  const lunasBulanIni = {};
  getRowsByFilter(SHEET.ANGSURAN, function(a) {
    return a.status_lock !== STATUS_LOCK.VOID && Number(a.tahun) === tahun &&
           Number(a.bulan) === bulan && String(a.keterangan || '').indexOf('Generate otomatis') === 0;
  }).forEach(function(a) { lunasBulanIni[String(a.id_pinjaman)] = true; });
  const aktifPinjaman = {};
  getRowsByFilter(SHEET.PINJAMAN, function(p) {
    return p.status_lock !== STATUS_LOCK.VOID && !String(p.id_induk || '') &&
      (p.status === 'aktif' || (p.status === 'lunas' && lunasBulanIni[String(p.id_pinjaman)]));
  }).forEach(function(p) {
    const id = String(p.id_anggota);
    (aktifPinjaman[id] = aktifPinjaman[id] || []).push(p);
  });

  const baris = [];
  sheetToObjects(SHEET.ANGGOTA).forEach(function(a) {
    const id = String(a.id_anggota);
    const aktif = a.status === 'aktif';
    const pinjaman = aktifPinjaman[id] || [];
    if (!aktif && pinjaman.length === 0) return;

    const setoranIni = simpananBulanIni.filter(function(r) { return String(r.id_anggota) === id; });
    const wajibSudah = setoranIni.some(function(r) { return r.jenis_simpanan === 'wajib'; });
    const sukarelaSudah = setoranIni.some(function(r) {
      return r.jenis_simpanan === 'sukarela' && String(r.keterangan || '').indexOf(tanda) === 0;
    });
    const pokokSudahIni = setoranIni.some(function(r) { return r.jenis_simpanan === 'pokok'; });

    const angs = pinjaman.map(function(p) { return rincianAngsuranBulan(p, tahun, bulan); })
      .filter(function(x) { return x; });
    const angsPokok = angs.reduce(function(s, x) { return s + x.pokok; }, 0);
    const jasa = angs.reduce(function(s, x) { return s + x.jasa; }, 0);
    const sisa = angs.length ? angs.reduce(function(s, x) { return s + x.sisaSetelah; }, 0)
      : pinjaman.reduce(function(s, p) { return s + getSisaPinjaman(p.id_pinjaman); }, 0);

    const pokok = aktif && (!sudahPokok[id] || pokokSudahIni) ? nominal.nominal_pokok : 0;
    const wajib = aktif ? nominal.nominal_wajib : 0;
    const sukarela = aktif ? Math.round(Number(a.sukarela_rutin) || 0) : 0;

    baris.push({
      id: id, nama: String(a.nama), nip: String(a.nip_nis || ''), status: a.status,
      pokok: pokok, wajib: wajib, sukarela: sukarela,
      angsuranPokok: angsPokok, jasa: jasa,
      jumlah: pokok + wajib + sukarela + angsPokok + jasa,
      sisaPiutang: sisa,
      // saldo SETELAH potongan bulan ini (seperti slip Excel): s.d. akhir bulan tagihan,
      // ditambah sukarela bulan ini bila belum dibukukan — sama sebelum & sesudah diproses
      saldoSukarela: getSaldoSukarelaAkhirBulan(id, tahun, bulan + 1) + (sukarelaSudah ? 0 : sukarela),
      jasaSukarela: aktif ? hitungJasaSukarelaAnggota(id, tahun, bulan, setting).nominal : 0,
      pinjamanBaru: angs.some(function(x) { return x.baru; }),
      status_proses: {
        pokok: pokok === 0 || pokokSudahIni,
        wajib: wajib === 0 || wajibSudah,
        sukarela: sukarela === 0 || sukarelaSudah,
        angsuran: angs.every(function(x) { return x.sudah || x.baru || x.pokok + x.jasa === 0; })
      }
    });
  });

  const total = { pokok: 0, wajib: 0, sukarela: 0, angsuranPokok: 0, jasa: 0, jumlah: 0, sisaPiutang: 0 };
  baris.forEach(function(b) { Object.keys(total).forEach(function(k) { total[k] += b[k]; }); });
  baris.forEach(function(b) {
    const s = b.status_proses;
    b.selesai = s.pokok && s.wajib && s.sukarela && s.angsuran;
  });
  return { tahun: tahun, bulan: bulan, namaBulan: NAMA_BULAN_PANJANG[bulan - 1],
    baris: baris, total: total };
}

/**
 * Bukukan potong gaji untuk anggota terpilih. Item yang sudah tercatat dilewati.
 * @return {Object} { diproses: [..], dilewati: n, gagal: [..] }
 */
/**
 * @param {string[]} daftarId       anggota yang dibukukan
 * @param {string[]} [tanpaAngsuran] anggota yang angsurannya TIDAK dipotong bulan ini
 *                                   (hanya simpanan yang dibukukan; angsuran tetap tertagih)
 */
function prosesPotongGaji(tahun, bulan, daftarId, tanpaAngsuran) {
  tahun = Number(tahun); bulan = Number(bulan);
  validateTahunAktif(tahun);
  const tanggal = new Date(tahun, bulan - 1, 10);   // tanggal potong gaji (seperti angsuran)
  const tanda = tandaPotongGaji(tahun, bulan);
  const pilih = {};
  (daftarId || []).forEach(function(id) { pilih[String(id)] = true; });
  const lewatiAngsuran = {};
  (tanpaAngsuran || []).forEach(function(id) { lewatiAngsuran[String(id)] = true; });

  const tagihan = hitungTagihanBulanan(tahun, bulan);
  const hasil = { diproses: [], dilewati: 0, gagal: [] };

  // Semua baris (setoran, angsuran, kas, jurnal) ditampung di memori lalu ditulis
  // sekaligus: ±50 panggilan Sheets, bukan ±50 per anggota.
  denganTulisMassal(function() {
  tagihan.baris.forEach(function(b) {
    if (!pilih[b.id]) return;
    if (b.selesai) { hasil.dilewati++; return; }
    const langkah = [];
    const coba = function(label, fn) {
      try { fn(); langkah.push(label); }
      catch (e) { hasil.gagal.push(b.nama + ' — ' + label + ': ' + e.message); }
    };
    if (!b.status_proses.pokok) coba('pokok', function() {
      inputSetoranSimpanan({ id_anggota: b.id, jenis_simpanan: 'pokok', jumlah_setoran: b.pokok,
        tanggal: tanggal, keterangan: tanda });
    });
    if (!b.status_proses.wajib) coba('wajib', function() {
      inputSetoranSimpanan({ id_anggota: b.id, jenis_simpanan: 'wajib', jumlah_setoran: b.wajib,
        tanggal: tanggal, keterangan: tanda });
    });
    if (!b.status_proses.sukarela) coba('sukarela', function() {
      inputSetoranSimpanan({ id_anggota: b.id, jenis_simpanan: 'sukarela', jumlah_setoran: b.sukarela,
        tanggal: tanggal, keterangan: tanda });
    });
    if (!b.status_proses.angsuran && !lewatiAngsuran[b.id]) {
      getPinjamanAktif(b.id).forEach(function(p) {
        const r = rincianAngsuranBulan(p, tahun, bulan);
        if (!r || r.sudah || r.baru || r.pokok + r.jasa === 0) return;
        coba('angsuran ' + p.id_pinjaman, function() { generateSatuAngsuran(p, tahun, bulan); });
      });
    }
    if (lewatiAngsuran[b.id] && !b.status_proses.angsuran) hasil.tanpaAngsuran = (hasil.tanpaAngsuran || 0) + 1;
    if (langkah.length) hasil.diproses.push(b.nama + ': ' + langkah.join(', '));
  });
  });

  logAktivitas('INSERT', SHEET.SIMPANAN, 'POTONG-GAJI-' + tahun + '-' + bulan, null,
    { anggota: Object.keys(pilih).length, diproses: hasil.diproses.length, gagal: hasil.gagal.length });
  return hasil;
}

// ============================================================
// API
// ============================================================

function apiTagihanBulanan(f) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const kini = new Date();
  const t = hitungTagihanBulanan(Number(f && f.tahun) || getTahunAktif() || kini.getFullYear(),
    Number(f && f.bulan) || kini.getMonth() + 1);
  t.profil = getProfilKoperasi();
  const pulih = cariPulihPotongGaji_(t.tahun, t.bulan);
  t.bisaDipulihkan = pulih.setoran.length + pulih.angsuran.length;
  return t;
}

function apiProsesPotongGaji(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  const ids = d.ids || [];
  if (ids.length === 0) throw new Error('Pilih minimal satu anggota.');
  if (ids.length > 100) throw new Error('Maksimal 100 anggota per proses (aplikasi memprosesnya bertahap).');
  return prosesPotongGaji(d.tahun, d.bulan, ids, d.tanpaAngsuran || []);
}

/**
 * Batalkan SELURUH pembukuan potong gaji satu bulan: setoran bertanda
 * "Potong gaji <bulan> <tahun>" dan angsuran hasil generate/potong gaji bulan itu.
 * Semua dibatalkan dengan jejak (VOID + jurnal balik), sekaligus atau tidak sama sekali.
 */
/**
 * Baris potong gaji bulan itu yang DIBATALKAN lewat "Batalkan pembukuan bulan ini" dan belum
 * dibukukan ulang (tidak ada baris OPEN pengganti untuk anggota/jenis atau pinjaman yang sama).
 */
function cariPulihPotongGaji_(tahun, bulan) {
  const penanda = ' | VOID: Batal potong gaji ' + NAMA_BULAN_PANJANG[bulan - 1] + ' ' + tahun;
  const tanda = tandaPotongGaji(tahun, bulan);
  const bulanIni = function(r) { return Number(r.tahun) === tahun && Number(r.bulan) === bulan; };
  const simpananAktif = {}, angsuranAktif = {};
  getRowsByFilter(SHEET.SIMPANAN, function(r) {
    return r.status_lock === STATUS_LOCK.OPEN && bulanIni(r) && String(r.keterangan || '').indexOf(tanda) === 0;
  }).forEach(function(r) { simpananAktif[r.id_anggota + '|' + r.jenis_simpanan] = true; });
  getRowsByFilter(SHEET.ANGSURAN, function(r) {
    return r.status_lock === STATUS_LOCK.OPEN && bulanIni(r) && String(r.keterangan || '').indexOf('Generate otomatis') === 0;
  }).forEach(function(r) { angsuranAktif[String(r.id_pinjaman)] = true; });
  const setoran = [], angsuran = [];
  getRowsByFilter(SHEET.SIMPANAN, function(r) {
    return r.status_lock === STATUS_LOCK.VOID && bulanIni(r) && String(r.keterangan || '').indexOf(penanda) !== -1;
  }).forEach(function(r) {
    const k = r.id_anggota + '|' + r.jenis_simpanan;
    if (!simpananAktif[k]) { simpananAktif[k] = true; setoran.push(r); }
  });
  getRowsByFilter(SHEET.ANGSURAN, function(r) {
    return r.status_lock === STATUS_LOCK.VOID && bulanIni(r) && String(r.keterangan || '').indexOf(penanda) !== -1;
  }).forEach(function(r) {
    const k = String(r.id_pinjaman);
    // pinjaman yang sudah lunas/kurang dari pokok angsuran ini (mis. dilunasi manual) → dipulihkan = kelebihan bayar
    if (!angsuranAktif[k] && (Number(r.angsuran_pokok) || 0) <= getSisaPinjaman(k) + 0.5) {
      angsuranAktif[k] = true; angsuran.push(r);
    }
  });
  return { penanda: penanda, setoran: setoran, angsuran: angsuran,
    total: setoran.reduce(function(s, r) { return s + (Number(r.jumlah_setoran) || 0); }, 0) +
      angsuran.reduce(function(s, r) { return s + (Number(r.total_bayar) || 0); }, 0) };
}

/** Kembalikan satu baris yang dibatalkan: OPEN lagi, kas aktif lagi, jurnal pembalikannya dibalik. */
function pulihkanBaris_(namaSheet, row, id, penanda) {
  const ket = String(row.keterangan || '');
  updateRowByRowNumber(namaSheet, row.__row, { status_lock: STATUS_LOCK.OPEN,
    keterangan: ket.slice(0, ket.indexOf(penanda)) + ' | dipulihkan' });
  getRowsByFilter(SHEET.KAS, function(k) {
    return k.referensi === id && k.status_lock === STATUS_LOCK.VOID;
  }).forEach(function(k) { updateRowByRowNumber(SHEET.KAS, k.__row, { status_lock: STATUS_LOCK.OPEN }); });
  buatJurnalBalik('VOID-' + id, 'Pemulihan pembukuan yang dibatalkan');
}

/**
 * Pulihkan pembukuan potong gaji satu bulan yang dibatalkan (mis. tidak sengaja).
 * d.simpan=false → hanya ringkasan.
 */
function apiPulihkanPotongGaji(d) {
  const profil = requireRole(['admin']);
  MODE_SENYAP = true;
  const tahun = Number(d && d.tahun), bulan = Number(d && d.bulan);
  if (!tahun || !(bulan >= 1 && bulan <= 12)) throw new Error('Bulan tidak valid.');
  validateTahunAktif(tahun);
  const c = cariPulihPotongGaji_(tahun, bulan);
  const anggota = {};
  c.setoran.concat(c.angsuran).forEach(function(r) { anggota[r.id_anggota] = true; });
  const hasil = { setoran: c.setoran.length, angsuran: c.angsuran.length, total: c.total,
    anggota: Object.keys(anggota).length, disimpan: !!(d && d.simpan) };
  if (!hasil.setoran && !hasil.angsuran) {
    throw new Error('Tidak ada pembukuan potong gaji ' + NAMA_BULAN_PANJANG[bulan - 1] + ' ' + tahun + ' yang bisa dipulihkan.');
  }
  if (!hasil.disimpan) return hasil;
  const pinjaman = {};
  denganTulisMassal(function() {
    c.setoran.forEach(function(r) { pulihkanBaris_(SHEET.SIMPANAN, r, r.id_transaksi, c.penanda); });
    c.angsuran.forEach(function(r) {
      pulihkanBaris_(SHEET.ANGSURAN, r, r.id_angsuran, c.penanda);
      pinjaman[String(r.id_pinjaman)] = true;
    });
    Object.keys(pinjaman).forEach(function(id) {
      if (getSisaPinjaman(id) <= 0) updateRowByField(SHEET.PINJAMAN, 'id_pinjaman', id, { status: 'lunas' });
    });
  });
  logAktivitas('EDIT', SHEET.SIMPANAN, 'PULIH-POTONG-GAJI-' + tahun + '-' + bulan, null,
    { setoran: hasil.setoran, angsuran: hasil.angsuran, total: hasil.total, oleh: profil.email });
  hasil.pesan = 'Pembukuan potong gaji ' + NAMA_BULAN_PANJANG[bulan - 1] + ' ' + tahun + ' dipulihkan: ' +
    hasil.setoran + ' setoran & ' + hasil.angsuran + ' angsuran dari ' + hasil.anggota + ' anggota (' +
    formatRupiah(hasil.total) + '). Kas & jurnal kembali seperti sebelum dibatalkan.';
  return hasil;
}

function apiBatalkanPotongGaji(d) {
  const profil = requireRole(['admin']);
  MODE_SENYAP = true;
  const tahun = Number(d && d.tahun), bulan = Number(d && d.bulan);
  const alasan = String(d && d.alasan || '').trim();
  if (!tahun || !(bulan >= 1 && bulan <= 12)) throw new Error('Bulan tidak valid.');
  if (!alasan) throw new Error('Alasan pembatalan wajib diisi.');
  validateTahunAktif(tahun);
  const tanda = tandaPotongGaji(tahun, bulan);
  const ket = 'Batal potong gaji ' + NAMA_BULAN_PANJANG[bulan - 1] + ' ' + tahun + ': ' + alasan;

  // Data hasil migrasi Excel adalah arsip pembukuan lama — tidak ikut dibatalkan dari sini
  const bukanMigrasi = function(r) { return String(r.keterangan || '').indexOf('migrasi Excel') === -1; };
  const setoran = getRowsByFilter(SHEET.SIMPANAN, function(r) {
    return r.status_lock === STATUS_LOCK.OPEN && Number(r.tahun) === tahun && Number(r.bulan) === bulan &&
      String(r.keterangan || '').indexOf(tanda) === 0 && bukanMigrasi(r);
  });
  const angsuran = getRowsByFilter(SHEET.ANGSURAN, function(r) {
    return r.status_lock === STATUS_LOCK.OPEN && Number(r.tahun) === tahun && Number(r.bulan) === bulan &&
      String(r.keterangan || '').indexOf('Generate otomatis') === 0 && bukanMigrasi(r);
  });
  if (!setoran.length && !angsuran.length) {
    const adaMigrasi = getRowsByFilter(SHEET.SIMPANAN, function(r) {
      return r.status_lock === STATUS_LOCK.OPEN && Number(r.tahun) === tahun && Number(r.bulan) === bulan && !bukanMigrasi(r);
    }).length;
    if (adaMigrasi) {
      throw new Error('Pembukuan ' + NAMA_BULAN_PANJANG[bulan - 1] + ' ' + tahun + ' berasal dari migrasi Excel dan tidak ' +
        'bisa dibatalkan dari sini. Koreksi transaksi tertentu lewat Riwayat → Void.');
    }
    throw new Error('Tidak ada pembukuan potong gaji ' + NAMA_BULAN_PANJANG[bulan - 1] + ' ' + tahun + ' yang bisa dibatalkan.');
  }
  const total = setoran.reduce(function(s, r) { return s + (Number(r.jumlah_setoran) || 0); }, 0) +
    angsuran.reduce(function(s, r) { return s + (Number(r.total_bayar) || 0); }, 0);
  const anggota = {};
  setoran.concat(angsuran).forEach(function(r) { anggota[r.id_anggota] = true; });

  mulaiTulisMassal();
  try {
    angsuran.forEach(function(r) { voidAngsuranInti_(r, ket); });
    setoran.forEach(function(r) { voidSetoranInti_(r, ket); });
    selesaiTulisMassal();
  } catch (e) {
    batalTulisMassal();
    throw e;
  }
  logAktivitas('VOID', SHEET.SIMPANAN, 'POTONG-GAJI-' + tahun + '-' + bulan, null,
    { setoran: setoran.length, angsuran: angsuran.length, total: total, alasan: alasan, oleh: profil.email });
  return {
    anggota: Object.keys(anggota).length, setoran: setoran.length, angsuran: angsuran.length, total: total,
    pesan: 'Pembukuan potong gaji ' + NAMA_BULAN_PANJANG[bulan - 1] + ' ' + tahun + ' dibatalkan: ' +
      setoran.length + ' setoran & ' + angsuran.length + ' angsuran dari ' + Object.keys(anggota).length +
      ' anggota (' + formatRupiah(total) + '). Kas & jurnal sudah dibalik; data tetap tercatat sebagai "Dibatalkan".'
  };
}

/**
 * Status proses satu bulan (untuk Beranda "Perlu dikerjakan" & langkah Bulanan).
 * Hanya membaca data.
 */
function apiStatusBulanan(f) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const kini = new Date();
  const tahun = Number(f && f.tahun) || getTahunAktif() || kini.getFullYear();
  const bulan = Number(f && f.bulan) || kini.getMonth() + 1;
  let tagihan = { baris: [] };
  try { tagihan = hitungTagihanBulanan(tahun, bulan); } catch (e) { /* setting tahun belum ada */ }
  const jasa = getRowsByFilter(SHEET.JASA_SUKARELA, function(r) {
    return Number(r.tahun) === tahun && Number(r.bulan) === bulan;
  });
  return {
    tahun: tahun, bulan: bulan, namaBulan: NAMA_BULAN_PANJANG[bulan - 1],
    anggotaDitagih: tagihan.baris.length,
    belumDibukukan: tagihan.baris.filter(function(b) { return !b.selesai; }).length,
    jasaDraft: jasa.filter(function(r) { return r.status_posting === 'DRAFT'; }).length,
    jasaPosted: jasa.filter(function(r) { return r.status_posting === 'POSTED'; }).length,
    pengajuanMenunggu: getDaftarPending().length
  };
}
