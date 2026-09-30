/**
 * Kode_3.gs — bagian 3 dari 3 kode Koperasi App.
 * Hak Cipta (c) Asep Hanuryana (Threads: @asep94). Seluruh hak dilindungi; lihat LICENSE.
 *
 * FILE INI DIBUAT OTOMATIS oleh tools/buat_paket_manual.py. Jangan disunting di sini:
 * ubah file aslinya di repositori, lalu jalankan ulang alat tersebut.
 *
 * Berisi: tagihan.js, tutup_buku.js, webapp.js
 */

// ===== tagihan.js =====
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

  // Anggota yang mendaftar SETELAH tanggal potong gaji (tgl 10) baru ditagih bulan berikutnya
  const batasMasuk = new Date(tahun, bulan - 1, 11);
  const baris = [];
  sheetToObjects(SHEET.ANGGOTA).forEach(function(a) {
    const id = String(a.id_anggota);
    const aktif = a.status === 'aktif';
    const pinjaman = aktifPinjaman[id] || [];
    if (!aktif && pinjaman.length === 0) return;

    const setoranIni = simpananBulanIni.filter(function(r) { return String(r.id_anggota) === id; });
    const masuk = a.tanggal_masuk ? new Date(a.tanggal_masuk) : null;
    if (masuk && !isNaN(masuk.getTime()) && masuk >= batasMasuk && pinjaman.length === 0 &&
        !setoranIni.some(function(r) { return String(r.keterangan || '').indexOf(tanda) === 0; })) return;
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

// ===== tutup_buku.js =====
/**
 * ============================================================
 * TUTUP_BUKU.GS — Proses Tutup Buku Tahunan
 * Koperasi App
 * Sesuai dokumen: REKAP_Koperasi_v2.md (BAGIAN 8, 12 #9, 16, 17)
 * ============================================================
 *
 * INI PROSES PALING "BERAT" DI SISTEM — dijalankan SEKALI SETAHUN
 * setelah RAT. Urutannya (sesuai BAGIAN 8 + 16):
 *
 * 1. VALIDASI PRA-TUTUP (bisa dicek dulu terpisah lewat menu):
 *    - Semua jasa sukarela tahun ini sudah POSTED (tidak ada DRAFT)
 *    - Jurnal balance (total debit = total kredit)
 *    - Setting RAT tahun BARU sudah dibuat di setting_rat
 * 2. Buat jurnal alokasi dana dari SHU (BAGIAN 12 #9):
 *    Debit 302 = Kredit 502 + 503 + 504
 * 3. LOCK semua transaksi tahun berjalan (status_lock → LOCKED)
 * 4. Hitung saldo akhir → tulis ke saldo_awal tahun baru
 * 5. Pindahkan status_aktif: tahun lama FALSE, tahun baru TRUE
 * 6. Generate semua laporan final (arsip)
 *
 * SETELAH TUTUP BUKU: input transaksi ke tahun lama otomatis
 * DITOLAK oleh validateTahunAktif (sudah terpasang di semua modul).
 * ============================================================
 */

// ============================================================
// VALIDASI PRA-TUTUP BUKU (BAGIAN 17)
// Bisa dijalankan terpisah untuk cek kesiapan tanpa eksekusi
// ============================================================

/**
 * Cek semua syarat tutup buku. Mengembalikan daftar masalah
 * (kosong = siap tutup buku).
 */
function cekKesiapanTutupBuku(tahunLama) {
  const masalah = [];
  const tahunBaru = Number(tahunLama) + 1;

  // 1. Setting tahun lama harus ada & aktif
  const settingLama = getSettingRAT(tahunLama);
  if (!settingLama) {
    masalah.push('Setting RAT tahun ' + tahunLama + ' tidak ditemukan.');
  } else if (settingLama.status_aktif !== true) {
    masalah.push('Tahun ' + tahunLama + ' tidak berstatus aktif — mungkin sudah ditutup?');
  }

  // 2. Setting tahun BARU harus sudah dibuat (hasil keputusan RAT)
  const settingBaru = getSettingRAT(tahunBaru);
  if (!settingBaru) {
    masalah.push('Setting RAT tahun ' + tahunBaru + ' BELUM dibuat. ' +
      'Tambahkan baris tahun ' + tahunBaru + ' di sheet setting_rat ' +
      '(isi sesuai keputusan RAT, status_aktif = FALSE dulu — ' +
      'nanti diaktifkan otomatis oleh proses tutup buku).');
  } else {
    try { validateAlokasiSHU(settingBaru); }
    catch (e) { masalah.push('Setting tahun ' + tahunBaru + ': ' + e.message); }
  }
  if (settingLama) {
    try { validateAlokasiSHU(settingLama); }
    catch (e) { masalah.push('Setting tahun ' + tahunLama + ': ' + e.message); }
  }

  // 3. Tidak boleh ada jasa sukarela DRAFT
  const draftJasa = getRowsByFilter(SHEET.JASA_SUKARELA, function(row) {
    return Number(row.tahun) === Number(tahunLama) && row.status_posting === 'DRAFT';
  });
  if (draftJasa.length > 0) {
    masalah.push('Masih ada ' + draftJasa.length + ' baris jasa sukarela berstatus DRAFT. ' +
      'Posting dulu (atau hapus jika salah) sebelum tutup buku.');
  }

  // 4. Jurnal harus balance
  const saldo = getSaldoPerAkun(tahunLama);
  let totalDebit = 0, totalKredit = 0;
  Object.keys(saldo).forEach(function(kode) {
    totalDebit += saldo[kode].debit;
    totalKredit += saldo[kode].kredit;
  });
  if (Math.abs(totalDebit - totalKredit) > 1) {
    masalah.push('Jurnal tahun ' + tahunLama + ' TIDAK BALANCE. Debit=' +
      formatRupiah(totalDebit) + ', Kredit=' + formatRupiah(totalKredit) +
      '. Telusuri jurnal_umum sebelum tutup buku.');
  }

  return masalah;
}

/** Menu: cek kesiapan tanpa mengeksekusi apapun. */
function menuCekKesiapanTutupBuku() {
  const tahun = getTahunAktif();
  if (!tahun) { tampilkanPesan('Tidak ada tahun aktif di setting_rat.'); return; }

  const masalah = cekKesiapanTutupBuku(tahun);
  tampilkanPesan(masalah.length === 0
    ? '✅ SIAP TUTUP BUKU tahun ' + tahun + '.\nSemua syarat terpenuhi.'
    : '⚠️ BELUM SIAP tutup buku tahun ' + tahun + ':\n\n• ' + masalah.join('\n• '));
}

// ============================================================
// PROSES TUTUP BUKU
// ============================================================

/**
 * Eksekusi tutup buku tahun aktif.
 * JANGAN panggil langsung — gunakan menuTutupBukuTahun yang
 * ada konfirmasi bertingkat.
 */
function prosesTutupBuku(tahunLama) {
  const tahunBaru = Number(tahunLama) + 1;

  // ---------- 0. VALIDASI FINAL ----------
  const masalah = cekKesiapanTutupBuku(tahunLama);
  if (masalah.length > 0) {
    throw new Error('Tutup buku dibatalkan:\n• ' + masalah.join('\n• '));
  }

  const laporan = [];

  // ---------- 1. JURNAL PENUTUP + ALOKASI SHU (7 pos) ----------
  // Pendapatan & biaya tahun lama dinolkan; SHU dipindah ke Cadangan
  // (modal sendiri) & dana-dana (hutang lancar) sesuai persen alokasi.
  const alokasi = buatJurnalPenutup(tahunLama);
  if (alokasi.shuKotor > 0) {
    laporan.push('Jurnal penutup dibuat. SHU ' + formatRupiah(alokasi.shuKotor) + ' dialokasikan: ' +
      alokasi.pos.map(function(p) { return p.label + ' ' + formatRupiah(p.nominal); }).join(', ') + '.');
  } else if (alokasi.shuKotor < 0) {
    laporan.push('Jurnal penutup dibuat. Rugi ' + formatRupiah(-alokasi.shuKotor) +
      ' mengurangi Cadangan.');
  } else {
    laporan.push('SHU = 0 → jurnal penutup tanpa alokasi.');
  }

  // ---------- 2. SALDO ANGGOTA ----------
  // Tidak disalin ke saldo_awal: saldo simpanan & piutang dihitung
  // kumulatif dari seluruh transaksi lintas tahun. (Versi lama menyalin
  // saldo sehingga simpanan terhitung dua kali setelah tutup buku.)

  // ---------- 3. LOCK SEMUA TRANSAKSI TAHUN LAMA ----------
  const sheetTransaksi = [SHEET.SIMPANAN, SHEET.PENGAMBILAN, SHEET.PINJAMAN,
                          SHEET.ANGSURAN, SHEET.KAS];
  let totalDikunci = 0;

  sheetTransaksi.forEach(function(namaSheet) {
    const sheet = getSheet(namaSheet);
    const header = getHeader(namaSheet);
    const idxTahun = header.indexOf('tahun');
    const idxLock = header.indexOf('status_lock');
    if (idxTahun === -1 || idxLock === -1) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    // Baca & tulis kolom lock sekaligus (cepat, bukan per sel)
    const dataTahun = sheet.getRange(2, idxTahun + 1, lastRow - 1, 1).getValues();
    const dataLock = sheet.getRange(2, idxLock + 1, lastRow - 1, 1).getValues();

    let berubah = 0;
    for (let i = 0; i < dataTahun.length; i++) {
      if (Number(dataTahun[i][0]) === Number(tahunLama) &&
          dataLock[i][0] === STATUS_LOCK.OPEN) {
        dataLock[i][0] = STATUS_LOCK.LOCKED;
        berubah++;
      }
    }
    if (berubah > 0) {
      sheet.getRange(2, idxLock + 1, lastRow - 1, 1).setValues(dataLock);
      hapusCache(namaSheet);
      totalDikunci += berubah;
    }
  });
  laporan.push(totalDikunci + ' transaksi tahun ' + tahunLama + ' dikunci (LOCKED).');

  // ---------- 4. PINDAHKAN STATUS AKTIF ----------
  updateRowByField(SHEET.SETTING_RAT, 'tahun', tahunLama, { status_aktif: false });
  updateRowByField(SHEET.SETTING_RAT, 'tahun', tahunBaru, { status_aktif: true });
  laporan.push('Tahun aktif dipindah: ' + tahunLama + ' → ' + tahunBaru + '.');

  // ---------- 5. GENERATE LAPORAN FINAL (ARSIP) ----------
  const modeSebelumnya = MODE_SENYAP;
  MODE_SENYAP = true;
  generateLaporanLabaRugi(tahunLama);
  generateLaporanNeraca(tahunLama);
  generateLaporanSHU(tahunLama);
  MODE_SENYAP = modeSebelumnya;
  laporan.push('Laporan final tahun ' + tahunLama + ' digenerate (lap_labarugi, lap_neraca, lap_shu).');

  // ---------- 6. LOG ----------
  logAktivitas('EDIT', SHEET.SETTING_RAT, 'TUTUP-BUKU-' + tahunLama,
    { tahun_aktif: tahunLama }, { tahun_aktif: tahunBaru, detail: laporan.join(' | ') });

  return laporan;
}

// ============================================================
// MENU DENGAN KONFIRMASI BERTINGKAT
// ============================================================

function menuTutupBukuTahun() {
  const ui = SpreadsheetApp.getUi();
  const tahun = getTahunAktif();
  if (!tahun) { ui.alert('Tidak ada tahun aktif di setting_rat.'); return; }

  // Cek kesiapan dulu, tampilkan hasilnya
  const masalah = cekKesiapanTutupBuku(tahun);
  if (masalah.length > 0) {
    ui.alert('⚠️ BELUM SIAP tutup buku tahun ' + tahun + ':\n\n• ' +
      masalah.join('\n• ') + '\n\nSelesaikan dulu, lalu coba lagi.');
    return;
  }

  const konfirmasi1 = ui.alert(
    '⚠️ TUTUP BUKU TAHUN ' + tahun,
    'Proses ini akan:\n' +
    '1. Membuat jurnal alokasi dana dari SHU\n' +
    '2. MENGUNCI semua transaksi tahun ' + tahun + ' (tidak bisa diedit lagi)\n' +
    '3. Membuat saldo_awal tahun ' + (tahun + 1) + '\n' +
    '4. Memindahkan tahun aktif ke ' + (tahun + 1) + '\n' +
    '5. Meng-generate laporan final\n\n' +
    'Proses TIDAK BISA di-undo otomatis. Lanjutkan?',
    ui.ButtonSet.YES_NO
  );
  if (konfirmasi1 !== ui.Button.YES) return;

  const konfirmasi2 = ui.prompt(
    '⚠️ KONFIRMASI TERAKHIR',
    'Ketik angka tahun yang mau ditutup (' + tahun + ') untuk melanjutkan:',
    ui.ButtonSet.OK_CANCEL
  );
  if (konfirmasi2.getSelectedButton() !== ui.Button.OK) return;
  if (konfirmasi2.getResponseText().trim() !== String(tahun)) {
    ui.alert('Angka tahun tidak cocok. Tutup buku dibatalkan.');
    return;
  }

  try {
    const laporan = prosesTutupBuku(tahun);
    ui.alert('✅ TUTUP BUKU TAHUN ' + tahun + ' SELESAI\n\n• ' + laporan.join('\n• ') +
      '\n\nTahun ' + (tahun + 1) + ' sekarang aktif. Selamat tahun buku baru! 🎉');
  } catch (e) {
    ui.alert('❌ GAGAL: ' + e.message);
  }
}
// ============================================================
// TEST DARI EDITOR (tanpa parameter, tanpa dialog konfirmasi)
// ============================================================

/**
 * Cek kesiapan tutup buku dari editor — TIDAK mengeksekusi apapun,
 * hanya menampilkan daftar syarat yang belum terpenuhi di log.
 */
function test_cekKesiapanTutupBuku() {
  MODE_SENYAP = true;
  const tahun = getTahunAktif();
  if (!tahun) { Logger.log('❌ Tidak ada tahun aktif di setting_rat.'); return; }

  const masalah = cekKesiapanTutupBuku(tahun);
  Logger.log(masalah.length === 0
    ? '✅ SIAP tutup buku tahun ' + tahun
    : '⚠️ BELUM SIAP tahun ' + tahun + ':\n• ' + masalah.join('\n• '));
}

/**
 * Eksekusi tutup buku dari editor untuk TAHUN AKTIF saat ini.
 * ⚠️ TANPA konfirmasi — hanya untuk testing di spreadsheet dummy!
 * Di penggunaan asli, SELALU pakai menu 📕 TUTUP BUKU TAHUN.
 */
function test_prosesTutupBukuTahunAktif() {
  MODE_SENYAP = true;
  const tahun = getTahunAktif();
  if (!tahun) { Logger.log('❌ Tidak ada tahun aktif.'); return; }

  Logger.log('Menutup buku tahun ' + tahun + '...');
  try {
    const laporan = prosesTutupBuku(tahun);
    Logger.log('✅ SELESAI:\n• ' + laporan.join('\n• '));
    Logger.log('Tahun aktif sekarang: ' + getTahunAktif());
  } catch (e) {
    Logger.log('❌ ' + e.message);
  }
}

// ===== webapp.js =====
/**
 * ============================================================
 * WEBAPP.GS — Server Aplikasi Web Penuh (v2)
 * Koperasi App
 * ============================================================
 *
 * MENGGANTIKAN: app.gs, ui.gs, form.html — HAPUS ketiganya dari
 * project Apps Script (supaya tidak ada dua doGet).
 * index.html juga DIGANTI TOTAL dengan versi baru.
 *
 * FITUR:
 * - Login berbasis akun Google (email dicek ke sheet "pengguna")
 * - Peran: admin (semua) / operator (harian, tanpa void, posting,
 *   approve, & kelola pengguna)
 * - CRUD Anggota, riwayat transaksi per anggota, VOID transaksi
 * - Semua endpoint input/approval/bulanan/laporan
 *
 * DEPLOY: Deploy → New deployment → Web app
 *   Execute as: Me | Who has access: Anyone with Google account
 *   (WAJIB "Anyone with Google account" agar email pengunjung
 *    terbaca untuk login)
 *
 * PENTING SAAT PERTAMA KALI:
 * Buka URL aplikasi dengan akun Google Anda sendiri (pemilik
 * script) — sistem otomatis mendaftarkan Anda sebagai ADMIN
 * pertama. Setelah itu tambahkan pengguna lain lewat halaman
 * Pengguna di dalam aplikasi.
 * ============================================================
 */

// ============================================================
// SHEET PENGGUNA (dibuat otomatis saat pertama dibutuhkan)
// ============================================================

const SHEET_PENGGUNA = 'pengguna';

function pastikanSheetPengguna() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_PENGGUNA);
  if (sheet) return sheet;

  sheet = ss.insertSheet(SHEET_PENGGUNA);
  const header = ['email', 'nama', 'peran', 'status', 'ditambahkan_oleh', 'timestamp'];
  sheet.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#0E4D3C');
  sheet.setFrozenRows(1);
  sheet.setTabColor('#0E4D3C');
  return sheet;
}

/**
 * Profil pengguna saat ini. null = belum masuk / nonaktif.
 *
 * Sumber utama: sesi login mandiri (__SESI_AKTIF, disetel oleh apiJalan
 * di sesi.gs). Session.getActiveUser() hanya dipakai sebagai cadangan
 * untuk pemanggilan dari editor Apps Script, karena Google TIDAK
 * memberikan email pengunjung web app di luar domain Workspace.
 */
function getProfilAktif() {
  pastikanSheetPengguna();

  let email = String(
    (typeof __SESI_AKTIF !== 'undefined' && __SESI_AKTIF) || ''
  ).trim().toLowerCase();

  if (!email) {
    email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  }
  if (!email) return null;

  const pemilik = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const daftar = sheetToObjects(SHEET_PENGGUNA);

  // Bootstrap: daftar kosong + yang membuka adalah pemilik script → jadi admin
  if (daftar.length === 0 && email === pemilik) {
    appendRowFromObject(SHEET_PENGGUNA, {
      email: email, nama: 'Pemilik Sistem', peran: 'admin', status: 'aktif',
      ditambahkan_oleh: 'sistem', timestamp: new Date()
    });
    return { email: email, nama: 'Pemilik Sistem', peran: 'admin' };
  }

  const p = daftar.find(function(u) {
    // trim() penting: email yang diketik/ditempel manual sering
    // membawa spasi tak terlihat di depan/belakang
    return String(u.email).trim().toLowerCase() === email;
  });
  if (!p) return null;

  // Status toleran: 'Aktif', 'AKTIF', ada spasi, atau kolom dikosongkan
  // saat baris ditambahkan manual di sheet → tetap dianggap aktif.
  const status = String(p.status || '').trim().toLowerCase();
  if (status !== '' && status !== 'aktif') return null;

  const peran = String(p.peran || '').trim().toLowerCase() === 'admin'
    ? 'admin' : 'operator';
  return { email: email, nama: p.nama || email.split('@')[0], peran: peran };
}

/**
 * DIAGNOSA AKSES — jalankan dari editor Apps Script (menu Run),
 * lalu buka View > Logs untuk melihat isi sheet `pengguna` apa adanya,
 * termasuk spasi tersembunyi dan huruf besar/kecil.
 */
function diagnosaAkses() {
  const daftar = sheetToObjects(SHEET_PENGGUNA);
  Logger.log('Jumlah baris terdaftar: ' + daftar.length);
  daftar.forEach(function(u, i) {
    Logger.log('Baris ' + (i + 2) +
      ' | email=[' + u.email + ']' +
      ' | panjang=' + String(u.email).length +
      ' (setelah trim=' + String(u.email).trim().length + ')' +
      ' | peran=[' + u.peran + ']' +
      ' | status=[' + u.status + ']');
  });
  Logger.log('Pemilik script: ' + Session.getEffectiveUser().getEmail());
  Logger.log('Email pembuka saat ini: [' +
    (Session.getActiveUser().getEmail() || '(KOSONG)') + ']');
  return daftar.length;
}

/** Penjaga peran di sisi server. */
function requireRole(daftarPeran) {
  const profil = getProfilAktif();
  if (!profil) throw new Error('Akses ditolak: akun Anda tidak terdaftar.');
  if (daftarPeran.indexOf(profil.peran) === -1) {
    throw new Error('Akses ditolak: fitur ini khusus ' + daftarPeran.join('/') + '.');
  }
  return profil;
}

// ============================================================
// PINTU MASUK WEB
// ============================================================

function doGet() {
  const judul = jalankanPublik_(function() {
    pastikanKolomSandi();
    pastikanKolomNominalSimpanan();
    pastikanKolomPinjaman();
    pastikanKolomAlokasi();
    pastikanCOA();
    pastikanKolomAnggota();
    return getProfilKoperasi();
  });
  let nama = String(judul.nama_koperasi || 'Koperasi').replace(/"/g, '').trim();
  if (nama === nama.toUpperCase()) nama = nama.toLowerCase().replace(/(^|\s)\S/g, function(h) { return h.toUpperCase(); });
  const kata = nama.split(/\s+/).filter(String);

  // Tidak ada lagi penolakan di sini: halaman selalu disajikan,
  // lalu klien menampilkan layar LOGIN. Identitas ditentukan oleh
  // sesi login mandiri (sesi.gs), bukan oleh akun Google pengunjung.
  const t = HtmlService.createTemplateFromFile('index');
  t.loginJudul = /^koperasi\b/i.test(nama) ? nama : 'Koperasi ' + nama;
  t.loginSub = judul.instansi || judul.jenis_koperasi || 'Aplikasi Pembukuan Koperasi';
  t.loginInisial = (kata.length > 1 ? kata[0][0] + kata[1][0] : nama.slice(0, 2)).toUpperCase();
  return t.evaluate()
    .setTitle(t.loginJudul)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// API: DATA UMUM
// ============================================================

function apiDataAwal() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const tahun = getTahunAktif();
  const setting = tahun ? getSettingRAT(tahun) : null;
  return {
    tahunAktif: tahun,
    jasaPinjaman: setting ? Number(setting.jasa_pinjaman) : 0,
    nominal: getNominalSimpanan(tahun),
    profil: getProfilKoperasi(),
    saldoKas: getSaldoKas()
  };
}

function apiDashboard() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;

  const aktif = getRowsByFilter(SHEET.ANGGOTA, function(a) { return a.status === 'aktif'; });
  const pinjaman = getRowsByFilter(SHEET.PINJAMAN, function(p) {
    return p.status === 'aktif' && p.status_lock !== STATUS_LOCK.VOID;
  });
  let piutang = 0;
  pinjaman.forEach(function(p) { piutang += getSisaPinjaman(p.id_pinjaman); });

  let simpanan = 0;
  aktif.forEach(function(a) {
    simpanan += getSaldoSimpanan(a.id_anggota, 'pokok') +
                getSaldoSimpanan(a.id_anggota, 'wajib') +
                getSaldoSimpanan(a.id_anggota, 'sukarela');
  });

  return {
    saldoKas: getSaldoKas(),
    jumlahAnggota: aktif.length,
    totalSimpanan: simpanan,
    jumlahPinjaman: pinjaman.length,
    totalPiutang: piutang,
    jumlahPending: getDaftarPending().length,
    tahunAktif: getTahunAktif()
  };
}

// ============================================================
// API: ANGGOTA (CRUD + DETAIL + RIWAYAT)
// ============================================================

function apiAnggotaList() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  return sheetToObjects(SHEET.ANGGOTA).map(function(a) {
    return {
      id: a.id_anggota, nip: a.nip_nis || '', nama: a.nama,
      jabatan: a.jabatan || '', telepon: a.telepon || '',
      status: a.status, sukarela_rutin: Number(a.sukarela_rutin) || 0
    };
  });
}

/** Setoran sukarela rutin (potong gaji) per bulan: kosong/0 atau ≥ minimal sukarela. */
function validasiSukarelaRutin(nilai) {
  const n = Math.round(Number(nilai) || 0);
  if (n < 0) throw new Error('Sukarela rutin tidak boleh minus.');
  const minimal = getNominalSimpanan(getTahunAktif()).minimal_sukarela;
  if (n > 0 && n < minimal) {
    throw new Error('Sukarela rutin minimal ' + formatRupiah(minimal) + ' (atau 0 bila tidak ada).');
  }
  return n;
}

function apiAnggotaTambah(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  if (!d.nama || String(d.nama).trim() === '') throw new Error('Nama wajib diisi.');

  pastikanKolomAnggota();
  const id = generateIdAnggota();
  const row = {
    id_anggota: id, nip_nis: d.nip || '', nama: String(d.nama).trim(),
    jabatan: d.jabatan || 'PNS', alamat: d.alamat || '', telepon: d.telepon || '',
    tanggal_masuk: new Date(), status: 'aktif',
    sukarela_rutin: validasiSukarelaRutin(d.sukarela_rutin)
  };
  appendRowFromObject(SHEET.ANGGOTA, row);
  logAktivitas('INSERT', SHEET.ANGGOTA, id, null, row);
  const pesan = '✅ Anggota ' + row.nama + ' terdaftar dengan ID ' + id + '.';

  // Opsional: langsung setor simpanan pokok (nominal dari Pengaturan)
  if (d.setorPokok) {
    try {
      const nominal = getNominalSimpanan(getTahunAktif());
      inputSetoranSimpanan({ id_anggota: id, jenis_simpanan: 'pokok',
        jumlah_setoran: nominal.nominal_pokok, keterangan: 'Setoran pokok saat mendaftar' });
      return pesan + '\nSimpanan pokok ' + formatRupiah(nominal.nominal_pokok) + ' tercatat.';
    } catch (e) {
      return pesan + '\n⚠️ Simpanan pokok BELUM tercatat: ' + e.message +
        ' — input manual di menu Transaksi.';
    }
  }
  return pesan;
}

function apiAnggotaEdit(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const lama = getAnggota(d.id);
  if (!lama) throw new Error('Anggota tidak ditemukan.');

  pastikanKolomAnggota();
  const baru = {
    nama: String(d.nama).trim(), nip_nis: d.nip || '',
    jabatan: d.jabatan || lama.jabatan, telepon: d.telepon || '',
    alamat: d.alamat || '',
    sukarela_rutin: validasiSukarelaRutin(d.sukarela_rutin)
  };
  updateRowByField(SHEET.ANGGOTA, 'id_anggota', d.id, baru);
  logAktivitas('EDIT', SHEET.ANGGOTA, d.id, lama, baru);
  return '✅ Data ' + baru.nama + ' diperbarui.';
}

function apiAnggotaStatus(id) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const lama = getAnggota(id);
  if (!lama) throw new Error('Anggota tidak ditemukan.');
  if (lama.status === 'keluar') {
    // Mengaktifkan kembali anggota yang sudah keluar hanya boleh admin
    // (biasanya setelah proses keluarnya di-VOID karena salah input).
    requireRole(['admin']);
  }
  const statusBaru = lama.status === 'aktif' ? 'nonaktif' : 'aktif';
  updateRowByField(SHEET.ANGGOTA, 'id_anggota', id, { status: statusBaru });
  logAktivitas('EDIT', SHEET.ANGGOTA, id, { status: lama.status }, { status: statusBaru });
  return '✅ ' + lama.nama + ' sekarang ' + statusBaru + '.';
}

/** Detail satu anggota: profil + saldo + seluruh riwayat. */
function apiAnggotaDetail(id) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const a = getAnggota(id);
  if (!a) throw new Error('Anggota tidak ditemukan.');

  const fmtTgl = function(t) {
    try { return Utilities.formatDate(new Date(t), Session.getScriptTimeZone(), 'dd/MM/yyyy'); }
    catch (e) { return String(t || ''); }
  };

  const setoran = getRowsByFilter(SHEET.SIMPANAN, function(r) { return r.id_anggota === id; })
    .map(function(r) { return { id: r.id_transaksi, tgl: fmtTgl(r.tanggal),
      jenis: r.jenis_simpanan, jumlah: Number(r.jumlah_setoran), lock: r.status_lock }; })
    .reverse();

  const pengambilan = getRowsByFilter(SHEET.PENGAMBILAN, function(r) { return r.id_anggota === id; })
    .map(function(r) { return { id: r.id_pengambilan, tgl: fmtTgl(r.tanggal),
      jumlah: Number(r.jumlah), approval: r.status_approval, lock: r.status_lock }; })
    .reverse();

  const pinjaman = getRowsByFilter(SHEET.PINJAMAN, function(r) { return r.id_anggota === id; })
    .map(function(r) { return { id: r.id_pinjaman, tgl: fmtTgl(r.tanggal),
      nominal: Number(r.nominal), tenor: r.tenor, tagihan: Number(r.total_tagihan),
      sisa: getSisaPinjaman(r.id_pinjaman), status: r.status, lock: r.status_lock,
      ket: r.keterangan || '' }; })
    .reverse();

  const angsuran = getRowsByFilter(SHEET.ANGSURAN, function(r) { return r.id_anggota === id; })
    .map(function(r) { return { id: r.id_angsuran, tgl: fmtTgl(r.tanggal),
      bulan: r.bulan + '/' + r.tahun, pokok: Number(r.angsuran_pokok),
      jasa: Number(r.jasa), total: Number(r.total_bayar), lock: r.status_lock }; })
    .reverse();

  return {
    profil: { id: a.id_anggota, nama: a.nama, nip: a.nip_nis || '',
      jabatan: a.jabatan || '', telepon: a.telepon || '', alamat: a.alamat || '',
      masuk: fmtTgl(a.tanggal_masuk), status: a.status,
      sukarela_rutin: Number(a.sukarela_rutin) || 0 },
    saldo: {
      pokok: getSaldoSimpanan(id, 'pokok'),
      wajib: getSaldoSimpanan(id, 'wajib'),
      sukarela: getSaldoSimpanan(id, 'sukarela')
    },
    setoran: setoran, pengambilan: pengambilan,
    pinjaman: pinjaman, angsuran: angsuran
  };
}

// ============================================================
// API: VOID TRANSAKSI (khusus admin — jejak audit tetap ada)
// ============================================================

/** Buat jurnal pembalik untuk semua entri jurnal ber-referensi tertentu. */
function buatJurnalBalik(referensi, alasan) {
  const entri = getRowsByFilter(SHEET.JURNAL, function(j) { return j.referensi === referensi; });
  const kini = new Date();
  entri.forEach(function(j) {
    const idJurnal = generateId(SHEET.JURNAL, ID_PREFIX.JURNAL, 1, Number(j.tahun), null);
    appendRowFromObject(SHEET.JURNAL, {
      id_jurnal: idJurnal, tanggal: kini, tahun: Number(j.tahun), bulan: Number(j.bulan),
      kode_akun: j.kode_akun, nama_akun: j.nama_akun,
      debit: Number(j.kredit), kredit: Number(j.debit),   // DIBALIK
      referensi: 'VOID-' + referensi,
      keterangan: 'Pembalikan (void): ' + (alasan || referensi)
    });
  });
  return entri.length;
}

/** Void baris kas yang menunjuk referensi tertentu. */
function voidKasByReferensi(referensi) {
  const rows = getRowsByFilter(SHEET.KAS, function(k) {
    return k.referensi === referensi && k.status_lock === STATUS_LOCK.OPEN;
  });
  rows.forEach(function(k) {
    updateRowByRowNumber(SHEET.KAS, k.__row, { status_lock: STATUS_LOCK.VOID });
  });
  return rows.length;
}

/** Batalkan satu setoran (baris OPEN): tandai VOID, kas & jurnal dibalik. */
function voidSetoranInti_(row, alasan) {
  updateRowByRowNumber(SHEET.SIMPANAN, row.__row, { status_lock: STATUS_LOCK.VOID,
    keterangan: (row.keterangan ? row.keterangan + ' | ' : '') + 'VOID: ' + alasan });
  voidKasByReferensi(row.id_transaksi);
  buatJurnalBalik(row.id_transaksi, alasan);
}

/** Batalkan satu angsuran (baris OPEN); pinjaman yang sempat lunas kembali aktif. @return sisa pinjaman */
function voidAngsuranInti_(row, alasan) {
  updateRowByRowNumber(SHEET.ANGSURAN, row.__row, { status_lock: STATUS_LOCK.VOID,
    keterangan: (row.keterangan ? row.keterangan + ' | ' : '') + 'VOID: ' + alasan });
  voidKasByReferensi(row.id_angsuran);
  buatJurnalBalik(row.id_angsuran, alasan);
  const sisa = getSisaPinjaman(row.id_pinjaman);
  if (sisa > 0) updateRowByField(SHEET.PINJAMAN, 'id_pinjaman', row.id_pinjaman, { status: 'aktif' });
  return sisa;
}

function apiVoidTransaksi(d) {
  const profil = requireRole(['admin']);
  MODE_SENYAP = true;
  const jenis = d.jenis, id = d.id, alasan = d.alasan || '';

  if (jenis === 'setoran') {
    const row = getRowByField(SHEET.SIMPANAN, 'id_transaksi', id);
    if (!row) throw new Error('Setoran tidak ditemukan.');
    if (row.status_lock !== STATUS_LOCK.OPEN)
      throw new Error('Hanya transaksi OPEN yang bisa di-void (status: ' + row.status_lock + ').');
    voidSetoranInti_(row, alasan);
    logAktivitas('VOID', SHEET.SIMPANAN, id, row, { alasan: alasan, oleh: profil.email });
    return '✅ Setoran ' + id + ' di-void. Kas & jurnal sudah dibalik.';
  }

  if (jenis === 'angsuran') {
    const row = getRowByField(SHEET.ANGSURAN, 'id_angsuran', id);
    if (!row) throw new Error('Angsuran tidak ditemukan.');
    if (row.status_lock !== STATUS_LOCK.OPEN)
      throw new Error('Hanya transaksi OPEN yang bisa di-void.');
    const sisa = voidAngsuranInti_(row, alasan);
    logAktivitas('VOID', SHEET.ANGSURAN, id, row, { alasan: alasan, oleh: profil.email });
    return '✅ Angsuran ' + id + ' di-void. Sisa pinjaman kembali ' + formatRupiah(sisa) + '.';
  }

  if (jenis === 'pengambilan') {
    const row = getRowByField(SHEET.PENGAMBILAN, 'id_pengambilan', id);
    if (!row) throw new Error('Pengambilan tidak ditemukan.');
    if (row.status_lock !== STATUS_LOCK.OPEN)
      throw new Error('Hanya transaksi OPEN yang bisa di-void.');
    updateRowByRowNumber(SHEET.PENGAMBILAN, row.__row, { status_lock: STATUS_LOCK.VOID,
      keterangan: (row.keterangan ? row.keterangan + ' | ' : '') + 'VOID: ' + alasan });
    if (row.status_approval === 'APPROVED') {
      voidKasByReferensi(id);
      buatJurnalBalik(id, alasan);
    }
    logAktivitas('VOID', SHEET.PENGAMBILAN, id, row, { alasan: alasan, oleh: profil.email });
    return '✅ Pengambilan ' + id + ' di-void.';
  }

  if (jenis === 'pinjaman') {
    const row = getRowByField(SHEET.PINJAMAN, 'id_pinjaman', id);
    if (!row) throw new Error('Pinjaman tidak ditemukan.');
    if (row.status_lock !== STATUS_LOCK.OPEN)
      throw new Error('Hanya transaksi OPEN yang bisa di-void.');
    const adaAngsuran = getRowsByFilter(SHEET.ANGSURAN, function(r) {
      return r.id_pinjaman === id && r.status_lock !== STATUS_LOCK.VOID;
    });
    if (adaAngsuran.length > 0)
      throw new Error('Pinjaman ini sudah punya ' + adaAngsuran.length +
        ' angsuran. Void semua angsurannya dulu.');
    const tambahan = getRowsByFilter(SHEET.PINJAMAN, function(r) {
      return String(r.id_induk || '') === String(id) && r.status_lock !== STATUS_LOCK.VOID;
    });
    if (tambahan.length > 0)
      throw new Error('Pinjaman ini punya pinjaman tambahan (' +
        tambahan.map(function(r) { return r.id_pinjaman; }).join(', ') +
        '). Void pinjaman tambahannya dulu.');

    // Pinjaman TAMBAHAN: sisa induk tidak boleh jadi minus setelah dibatalkan
    const idInduk = String(row.id_induk || '');
    const induk = idInduk ? getRowByField(SHEET.PINJAMAN, 'id_pinjaman', idInduk) : null;
    if (induk && getSisaPinjaman(idInduk) - Number(row.nominal) < 0)
      throw new Error('Angsuran ' + idInduk + ' yang sudah dibayar melebihi pinjaman ' +
        'tanpa tambahan ini. Void angsuran terakhirnya dulu.');

    updateRowByRowNumber(SHEET.PINJAMAN, row.__row, { status_lock: STATUS_LOCK.VOID,
      status: induk ? 'tambahan' : 'lunas',
      keterangan: (row.keterangan ? row.keterangan + ' | ' : '') + 'VOID: ' + alasan });
    voidKasByReferensi(id);
    buatJurnalBalik(id, alasan);

    if (induk) {
      // Kembalikan angsuran pokok induk seperti sebelum tambahan
      const pulih = {};
      if (row.angsuran_sebelumnya !== '' && row.angsuran_sebelumnya !== undefined)
        pulih.angsuran_perbulan = Number(row.angsuran_sebelumnya);
      pulih.status = getSisaPinjaman(idInduk) > 0 ? 'aktif' : 'lunas';
      updateRowByRowNumber(SHEET.PINJAMAN, induk.__row, pulih);
      logAktivitas('EDIT', SHEET.PINJAMAN, idInduk,
        { angsuran_perbulan: induk.angsuran_perbulan, status: induk.status },
        Object.assign({ void_tambahan: id }, pulih));
    }
    logAktivitas('VOID', SHEET.PINJAMAN, id, row, { alasan: alasan, oleh: profil.email });
    return '✅ Pinjaman ' + id + ' di-void.' +
      (induk ? ' Sisa ' + idInduk + ' kembali ' + formatRupiah(getSisaPinjaman(idInduk)) + '.' : '');
  }

  throw new Error('Jenis transaksi tidak dikenal: ' + jenis);
}

// ============================================================
// API: INPUT TRANSAKSI (memakai mesin yang sudah ada)
// ============================================================

function apiSubmitSetoran(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  return inputSetoranSimpanan({
    id_anggota: d.id_anggota, jenis_simpanan: d.jenis_simpanan,
    jumlah_setoran: Number(d.jumlah), keterangan: d.keterangan || ''
  }).pesan;
}

function apiSubmitPengambilan(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  return inputPengajuanPengambilan({
    id_anggota: d.id_anggota, jumlah: Number(d.jumlah), keterangan: d.keterangan || ''
  }).pesan;
}

function apiSubmitPinjaman(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const hasil = inputPinjamanBaru({
    id_anggota: d.id_anggota, nominal: Number(d.nominal),
    tenor: Number(d.tenor), keterangan: d.keterangan || '',
    jasa_persen: d.jasa_persen, angsuran_pokok: d.angsuran_pokok
  });
  return { pesan: hasil.pesan, struk: dataStruk('pinjaman', hasil.id_pinjaman) };
}

function apiSubmitMigrasi(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  return inputPinjamanMigrasi({
    id_anggota: d.id_anggota, sisa_pokok: Number(d.sisa_pokok),
    tenor_sisa: Number(d.tenor_sisa),
    angsuran_pokok: d.angsuran_pokok, jasa_persen: d.jasa_persen,
    tanggal_asli: new Date(Number(d.tahun_asli), Number(d.bulan_asli) - 1, 1),
    keterangan_tambahan: d.keterangan || ''
  }).pesan;
}

function apiSaldoAnggota(id) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  return {
    pokok: getSaldoSimpanan(id, 'pokok'),
    wajib: getSaldoSimpanan(id, 'wajib'),
    sukarela: getSaldoSimpanan(id, 'sukarela')
  };
}

// ============================================================
// API: APPROVAL
// ============================================================

function apiPendingList() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  return getDaftarPending().map(function(r) {
    const a = getAnggota(r.id_anggota);
    return {
      id: r.id_pengambilan, nama: a ? a.nama : r.id_anggota,
      jumlah: Number(r.jumlah),
      tanggal: Utilities.formatDate(new Date(r.tanggal),
        Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      keterangan: r.keterangan || '',
      saldo: getSaldoSimpanan(r.id_anggota, 'sukarela')
    };
  });
}

function apiApprove(id) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  return approvePengambilan(id).pesan;
}

function apiReject(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  return rejectPengambilan(d.id, d.alasan).pesan;
}

// ============================================================
// API: PROSES BULANAN
// ============================================================

function apiGenerateAngsuran() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const h = generateAngsuranBulanan();
  return '✅ Generate angsuran selesai.\nBerhasil: ' + h.berhasil.length +
    '\nDilewati (sudah ada): ' + h.dilewati.length +
    (h.gagal.length ? '\nGagal: ' + h.gagal.length + '\n' + h.gagal.join('\n') : '');
}

function apiCetakStruk() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const h = cetakStrukPotongan();
  return '✅ Struk siap di sheet struk_potongan.\nAnggota: ' + h.jumlahAnggota +
    '\nTotal potongan: ' + formatRupiah(h.totalPotongan);
}

/** Bulan jasa dari UI (halaman Bulanan); tanpa isian → bulan berjalan. */
function periodeJasa_(d) {
  const sekarang = new Date();
  const tahun = Number(d && d.tahun) || getTahunAktif();
  const bulan = Number(d && d.bulan) || (sekarang.getMonth() + 1);
  if (!(bulan >= 1 && bulan <= 12)) throw new Error('Bulan tidak valid.');
  if (tahun * 12 + bulan > sekarang.getFullYear() * 12 + sekarang.getMonth() + 1) {
    throw new Error('Jasa ' + NAMA_BULAN_PANJANG[bulan - 1] + ' ' + tahun +
      ' belum bisa dihitung — bulannya belum berjalan.');
  }
  return { tahun: tahun, bulan: bulan, nama: NAMA_BULAN_PANJANG[bulan - 1] + ' ' + tahun };
}

function apiHitungJasa(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const p = periodeJasa_(d);
  const h = hitungJasaSukarelaBulanan(p.tahun, p.bulan);
  return ringkasanJasaBulan_(p, '✅ Hitung jasa ' + p.nama + ' selesai (DRAFT).\nDibuat: ' + h.dibuat.length +
    '\nDilewati (saldo 0): ' + h.dilewatiSaldoNol.length +
    '\nDilewati (sudah ada): ' + h.dilewatiSudahAda.length);
}

/**
 * Rincian jasa satu bulan (semua baris DRAFT/POSTED, bukan hanya yang baru dibuat) untuk
 * pratinjau di layar, plus bulan-bulan sebelumnya di tahun itu yang belum punya jasa POSTED
 * (saldo dasar bulan ini belum memuat jasa bulan-bulan itu).
 */
function ringkasanJasaBulan_(p, pesan) {
  const setting = getSettingRAT(p.tahun) || {};
  const semua = sheetToObjects(SHEET.JASA_SUKARELA).filter(function(r) {
    return Number(r.tahun) === p.tahun && r.status_posting !== 'VOID';
  });
  const rincian = semua.filter(function(r) { return Number(r.bulan) === p.bulan; }).map(function(r) {
    const a = getAnggota(String(r.id_anggota));
    return { id: String(r.id_anggota), nama: a ? String(a.nama) : String(r.id_anggota),
      status: a ? String(a.status || '') : '', dasar: Number(r.dasar_jasa || r.saldo_awal_bulan) || 0,
      nominal: Number(r.nominal_jasa) || 0, posting: r.status_posting };
  }).sort(function(x, y) { return x.id < y.id ? -1 : 1; });
  const bulanKosong = [];
  for (let b = 1; b < p.bulan; b++) {
    if (!semua.some(function(r) { return Number(r.bulan) === b && r.status_posting === 'POSTED'; })) {
      bulanKosong.push(NAMA_BULAN_PANJANG[b - 1]);
    }
  }
  return {
    pesan: pesan, tahun: p.tahun, bulan: p.bulan, nama: p.nama,
    rumus: KET_METODE_JASA[metodeJasaSukarela(setting)].replace('{p}', Number(setting.jasa_sukarela) || 0),
    rincian: rincian, total: rincian.reduce(function(s, r) { return s + r.nominal; }, 0),
    bulanKosong: bulanKosong
  };
}

function apiPostingJasa(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  const p = periodeJasa_(d);
  const h = postingJasaSukarela(p.tahun, p.bulan);
  return h.diposting > 0
    ? '✅ Posting jasa ' + p.nama + ' selesai. ' + h.diposting + ' anggota, total ' +
      formatRupiah(h.totalJasa) + '. Data terkunci.'
    : 'Tidak ada DRAFT jasa ' + p.nama + ' untuk diposting.';
}

/** Batalkan jasa sukarela bulan terpilih & sesudahnya. d.simpan=false → ringkasan saja. */
function apiBatalkanJasa(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  const p = periodeJasa_(d);
  return batalkanJasaSukarela(p.tahun, p.bulan, d && d.alasan, !!(d && d.simpan));
}

/** Jasa sukarela bulan-bulan yang terlewat (Januari s.d. bulan terpilih). d.simpan=false → pratinjau. */
function apiJasaSusulan(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  const p = periodeJasa_(d);
  return jasaSukarelaSusulan(p.tahun, p.bulan, !!(d && d.simpan));
}

// ============================================================
// API: LAPORAN
// ============================================================

function apiLaporan(nama) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  if (nama === 'labarugi') {
    return '✅ Laba Rugi selesai. SHU Kotor: ' + formatRupiah(generateLaporanLabaRugi());
  }
  if (nama === 'neraca') {
    const h = generateLaporanNeraca();
    return '✅ Neraca selesai. Aktiva ' + formatRupiah(h.totalAktiva) +
      (h.balance ? ' — BALANCE' : ' — ❌ TIDAK BALANCE');
  }
  if (nama === 'shu') {
    const h = generateLaporanSHU();
    return '✅ Laporan SHU selesai. Netto anggota: ' + formatRupiah(h.alokasi.shuNetto);
  }
  if (nama === 'rekap_simpanan') { generateRekapSimpanan(); return '✅ Rekap simpanan selesai.'; }
  if (nama === 'rekap_pinjaman') { generateRekapPinjaman(); return '✅ Rekap pinjaman selesai.'; }
  if (nama === 'semua') {
    generateLaporanLabaRugi(); generateLaporanNeraca(); generateLaporanSHU();
    generateRekapSimpanan(); generateRekapPinjaman();
    return '✅ Semua laporan selesai digenerate.';
  }
  throw new Error('Jenis laporan tidak dikenal.');
}

function apiDataSHU() {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const tahun = getTahunAktif();
  const alokasi = hitungAlokasiDana(tahun);
  const setting = getSettingRAT(tahun);
  const baris = daftarSHUAnggota(tahun).map(function(r) {
    return { nama: r.nama, status: r.status, simpanan: r.simpanan, jasa: r.jasa,
      shuSimpanan: r.shuSimpanan, shuJasa: r.shuJasa, total: r.total };
  });
  return { tahun: tahun, alokasi: alokasi, baris: baris,
    shuSimpananPersen: Number(setting.shu_simpanan), shuJasaPersen: Number(setting.shu_jasa) };
}

// ============================================================
// API: KELOLA PENGGUNA (khusus admin)
// ============================================================

function apiPenggunaList() {
  requireRole(['admin']);
  MODE_SENYAP = true;
  return sheetToObjects(SHEET_PENGGUNA).map(function(u) {
    return { email: u.email, nama: u.nama, peran: u.peran, status: u.status };
  });
}

function apiPenggunaTambah(d) {
  const profil = requireRole(['admin']);
  MODE_SENYAP = true;
  const email = String(d.email || '').trim().toLowerCase();
  if (!email || email.indexOf('@') === -1) throw new Error('Email tidak valid.');

  const sudahAda = getRowsByFilter(SHEET_PENGGUNA, function(u) {
    return String(u.email).toLowerCase() === email;
  });
  if (sudahAda.length > 0) throw new Error('Email ini sudah terdaftar.');

  validasiKekuatanSandi(d.sandi);
  const garam = buatGaram();

  const peran = d.peran === 'admin' ? 'admin' : 'operator';
  appendRowFromObject(SHEET_PENGGUNA, {
    email: email, nama: d.nama || email.split('@')[0], peran: peran,
    status: 'aktif', ditambahkan_oleh: profil.email, timestamp: new Date(),
    sandi_hash: hashSandi(String(d.sandi), garam), garam: garam,
    wajib_ganti: 'ya', terakhir_login: ''
  });
  logAktivitas('INSERT', SHEET_PENGGUNA, email, null, { peran: peran });
  return '✅ ' + email + ' terdaftar sebagai ' + peran + '.';
}

function apiPenggunaStatus(email) {
  const profil = requireRole(['admin']);
  MODE_SENYAP = true;
  const target = String(email).toLowerCase();
  if (target === profil.email) throw new Error('Tidak bisa menonaktifkan akun sendiri.');

  const row = getRowsByFilter(SHEET_PENGGUNA, function(u) {
    return String(u.email).toLowerCase() === target;
  })[0];
  if (!row) throw new Error('Pengguna tidak ditemukan.');

  const baru = row.status === 'aktif' ? 'nonaktif' : 'aktif';
  updateRowByRowNumber(SHEET_PENGGUNA, row.__row, { status: baru });
  logAktivitas('EDIT', SHEET_PENGGUNA, target, { status: row.status }, { status: baru });
  return '✅ ' + target + ' sekarang ' + baru + '.';
}
