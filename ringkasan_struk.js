/**
 * ============================================================
 * RINGKASAN_STRUK.GS — Dashboard per Periode, Riwayat Transaksi,
 *                      dan Data Struk (printer thermal / kuitansi)
 * Koperasi App
 * ============================================================
 *
 * 1. apiDashboardPeriode  — ringkasan kas masuk/keluar, grafik
 *    bulanan, dan komposisi simpanan untuk rentang tanggal.
 * 2. apiRiwayatTransaksi  — gabungan setoran, pengambilan,
 *    angsuran, dan pinjaman (terbaru di atas), bisa difilter.
 * 3. apiStruk / dataStruk — isi bukti transaksi untuk dicetak.
 *
 * Semua fungsi di sini hanya MEMBACA data (tidak dikunci).
 * Tanggal dari klien berformat 'yyyy-MM-dd'.
 * ============================================================
 */

const NAMA_BULAN_PENDEK = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun',
  'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

/** 'yyyy-MM-dd' → Date (awal hari). akhirHari=true → 23:59:59.999 */
function parseTanggalKlien(teks, akhirHari) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(teks || ''));
  if (!m) return null;
  return akhirHari
    ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 23, 59, 59, 999)
    : new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function formatTgl(t, pola) {
  try {
    return Utilities.formatDate(new Date(t), Session.getScriptTimeZone(), pola || 'dd/MM/yyyy');
  } catch (e) {
    return String(t || '');
  }
}

/** Rentang dari filter klien; default 6 bulan terakhir s/d hari ini. */
function rentangDariFilter(f) {
  f = f || {};
  const kini = new Date();
  let dari = parseTanggalKlien(f.dari, false);
  let sampai = parseTanggalKlien(f.sampai, true);
  if (!sampai) sampai = new Date(kini.getFullYear(), kini.getMonth(), kini.getDate(), 23, 59, 59, 999);
  if (!dari) dari = new Date(sampai.getFullYear(), sampai.getMonth() - 5, 1);
  if (dari > sampai) throw new Error('Tanggal "dari" tidak boleh setelah tanggal "sampai".');
  return { dari: dari, sampai: sampai };
}

function dalamRentang(tanggal, r) {
  const t = new Date(tanggal);
  return !isNaN(t.getTime()) && t >= r.dari && t <= r.sampai;
}

// ============================================================
// 1. DASHBOARD PER PERIODE
// ============================================================

function apiDashboardPeriode(f) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  const r = rentangDariFilter(f);

  // Ember per bulan dalam rentang (maksimal 24 bulan terakhir rentang)
  const ember = [];
  const indeks = {};
  let kursor = new Date(r.dari.getFullYear(), r.dari.getMonth(), 1);
  while (kursor <= r.sampai) {
    const kunci = kursor.getFullYear() + '-' + kursor.getMonth();
    indeks[kunci] = ember.length;
    ember.push({ label: NAMA_BULAN_PENDEK[kursor.getMonth()] + ' ' +
      String(kursor.getFullYear()).slice(2), masuk: 0, keluar: 0 });
    kursor = new Date(kursor.getFullYear(), kursor.getMonth() + 1, 1);
  }

  let masuk = 0, keluar = 0;
  const perKategori = {};
  sheetToObjects(SHEET.KAS).forEach(function(k) {
    if (k.status_lock === STATUS_LOCK.VOID || !dalamRentang(k.tanggal, r)) return;
    const m = Number(k.masuk) || 0, kl = Number(k.keluar) || 0;
    masuk += m; keluar += kl;
    const kat = String(k.kategori || 'lainnya');
    if (!perKategori[kat]) perKategori[kat] = { masuk: 0, keluar: 0 };
    perKategori[kat].masuk += m; perKategori[kat].keluar += kl;

    const t = new Date(k.tanggal);
    const i = indeks[t.getFullYear() + '-' + t.getMonth()];
    if (i !== undefined) { ember[i].masuk += m; ember[i].keluar += kl; }
  });

  let jasaPinjaman = 0;
  sheetToObjects(SHEET.ANGSURAN).forEach(function(a) {
    if (a.status_lock === STATUS_LOCK.VOID || !dalamRentang(a.tanggal, r)) return;
    jasaPinjaman += Number(a.jasa) || 0;
  });

  // Komposisi simpanan saat ini (semua anggota)
  const komposisi = { pokok: 0, wajib: 0, sukarela: 0 };
  const peta = getPetaSaldoSimpanan();
  Object.keys(peta).forEach(function(id) {
    komposisi.pokok += peta[id].pokok;
    komposisi.wajib += peta[id].wajib;
    komposisi.sukarela += peta[id].sukarela;
  });

  return {
    dari: formatTgl(r.dari, 'yyyy-MM-dd'), sampai: formatTgl(r.sampai, 'yyyy-MM-dd'),
    masuk: masuk, keluar: keluar, bersih: masuk - keluar,
    jasaPinjaman: jasaPinjaman,
    saldoKasAkhir: getSaldoKas(r.sampai),
    perKategori: perKategori,
    bulanan: ember.slice(-24),
    komposisi: komposisi
  };
}

// ============================================================
// 2. RIWAYAT TRANSAKSI (gabungan)
// ============================================================

/**
 * @param {Object} f { jenis: 'semua'|'setoran'|'pengambilan'|'angsuran'|'pinjaman',
 *                     dari, sampai }
 */
function apiRiwayatTransaksi(f) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  f = f || {};
  const r = rentangDariFilter(f);
  const jenis = f.jenis || 'semua';
  const ikut = function(j) { return jenis === 'semua' || jenis === j; };

  const namaAnggota = {};
  sheetToObjects(SHEET.ANGGOTA).forEach(function(a) {
    namaAnggota[String(a.id_anggota)] = a.nama;
  });

  const hasil = [];
  const tambah = function(row, isi) {
    if (!dalamRentang(row.tanggal, r)) return;
    const idAgt = String(row.id_anggota || '');
    isi.id_anggota = idAgt;
    isi.nama = namaAnggota[idAgt] || idAgt;
    isi.tgl = formatTgl(row.tanggal);
    isi.urut = new Date(row.tanggal).getTime() || 0;
    isi.urutWaktu = new Date(row.timestamp || row.tanggal).getTime() || 0;
    isi.keterangan = String(row.keterangan || '');
    isi.lock = row.status_lock || '';
    hasil.push(isi);
  };

  if (ikut('setoran')) sheetToObjects(SHEET.SIMPANAN).forEach(function(s) {
    tambah(s, { jenis: 'setoran', id: s.id_transaksi,
      uraian: 'Setoran ' + s.jenis_simpanan, masuk: Number(s.jumlah_setoran) || 0, keluar: 0,
      status: s.status_lock });
  });
  if (ikut('pengambilan')) sheetToObjects(SHEET.PENGAMBILAN).forEach(function(p) {
    tambah(p, { jenis: 'pengambilan', id: p.id_pengambilan,
      uraian: 'Tarik ' + (p.jenis_simpanan || 'sukarela'), masuk: 0, keluar: Number(p.jumlah) || 0,
      status: p.status_lock === STATUS_LOCK.VOID ? 'VOID' : p.status_approval });
  });
  if (ikut('angsuran')) sheetToObjects(SHEET.ANGSURAN).forEach(function(a) {
    tambah(a, { jenis: 'angsuran', id: a.id_angsuran,
      uraian: 'Angsuran ' + a.id_pinjaman, masuk: Number(a.total_bayar) || 0, keluar: 0,
      status: a.status_lock });
  });
  if (ikut('pinjaman')) sheetToObjects(SHEET.PINJAMAN).forEach(function(p) {
    const migrasi = String(p.keterangan || '').indexOf('MIGRASI') === 0;
    tambah(p, { jenis: 'pinjaman', id: p.id_pinjaman,
      uraian: migrasi ? 'Pinjaman migrasi'
        : String(p.id_induk || '') ? 'Tambahan pinjaman (' + p.id_induk + ')' : 'Pencairan pinjaman',
      masuk: 0, keluar: migrasi ? 0 : Number(p.nominal) || 0,
      status: p.status_lock === STATUS_LOCK.VOID ? 'VOID' : p.status });
  });

  hasil.sort(function(a, b) { return (b.urut - a.urut) || (b.urutWaktu - a.urutWaktu); });
  hasil.forEach(function(h) { delete h.urut; delete h.urutWaktu; });
  return {
    dari: formatTgl(r.dari, 'yyyy-MM-dd'), sampai: formatTgl(r.sampai, 'yyyy-MM-dd'),
    rows: hasil
  };
}

// ============================================================
// 3. STRUK
// ============================================================

function namaPetugas(email) {
  const u = String(email || '').trim().toLowerCase();
  if (!u || u === 'unknown') return '-';
  let daftar = [];
  try { daftar = sheetToObjects(SHEET_PENGGUNA); } catch (e) { /* sheet belum ada */ }
  const p = daftar.find(function(x) { return String(x.email).trim().toLowerCase() === u; });
  return p && p.nama ? String(p.nama) : u.split('@')[0];
}

/**
 * Isi struk satu transaksi.
 * @return {Object} { judul, no, tanggal, nama, id_anggota, rincian:[{l,v,rp}],
 *                    utama:{l,v}, catatan, petugas, dicetak, batal }
 */
function dataStruk(jenis, id) {
  const cari = {
    setoran: [SHEET.SIMPANAN, 'id_transaksi'],
    pengambilan: [SHEET.PENGAMBILAN, 'id_pengambilan'],
    angsuran: [SHEET.ANGSURAN, 'id_angsuran'],
    pinjaman: [SHEET.PINJAMAN, 'id_pinjaman']
  }[jenis];
  if (!cari) throw new Error('Jenis struk tidak dikenal: ' + jenis);
  const row = getRowByField(cari[0], cari[1], id);
  if (!row) throw new Error('Transaksi ' + id + ' tidak ditemukan.');

  const anggota = getAnggota(row.id_anggota);
  const s = {
    no: String(id),
    tanggal: formatTgl(row.tanggal) + (row.timestamp ? ' ' + formatTgl(row.timestamp, 'HH:mm') : ''),
    nama: anggota ? String(anggota.nama) : String(row.id_anggota),
    id_anggota: String(row.id_anggota),
    rincian: [], catatan: '',
    petugas: namaPetugas(row.user_input),
    dicetak: formatTgl(new Date(), 'dd/MM/yyyy HH:mm'),
    batal: row.status_lock === STATUS_LOCK.VOID
  };

  if (jenis === 'setoran') {
    s.judul = 'BUKTI SETORAN SIMPANAN';
    s.rincian.push({ l: 'Jenis', v: 'Simpanan ' + row.jenis_simpanan });
    s.utama = { l: 'JUMLAH', v: Number(row.jumlah_setoran) };
    s.rincian2 = [{ l: 'Saldo ' + row.jenis_simpanan + ' kini',
      v: getSaldoSimpanan(row.id_anggota, row.jenis_simpanan), rp: true }];
  } else if (jenis === 'pengambilan') {
    s.judul = 'BUKTI PENGAMBILAN SIMPANAN';
    s.rincian.push({ l: 'Jenis', v: 'Simpanan ' + (row.jenis_simpanan || 'sukarela') });
    s.rincian.push({ l: 'Status', v: String(row.status_approval) });
    s.utama = { l: 'JUMLAH', v: Number(row.jumlah) };
    s.rincian2 = [{ l: 'Saldo ' + (row.jenis_simpanan || 'sukarela') + ' kini',
      v: getSaldoSimpanan(row.id_anggota, row.jenis_simpanan || 'sukarela'), rp: true }];
  } else if (jenis === 'angsuran') {
    s.judul = 'BUKTI PEMBAYARAN ANGSURAN';
    const pinjaman = getRowByField(SHEET.PINJAMAN, 'id_pinjaman', row.id_pinjaman);
    // Urutan angsuran & sisa SETELAH pembayaran ini (hanya yang tidak VOID)
    const semua = getRowsByFilter(SHEET.ANGSURAN, function(a) {
      return String(a.id_pinjaman) === String(row.id_pinjaman) &&
             a.status_lock !== STATUS_LOCK.VOID;
    }).sort(function(a, b) { return a.__row - b.__row; });
    const posisi = semua.findIndex(function(a) { return a.__row === row.__row; });
    const menurun = isJasaMenurun(pinjaman);
    let terbayar = 0;
    for (let i = 0; i <= posisi; i++) {
      terbayar += Number(menurun ? semua[i].angsuran_pokok : semua[i].total_bayar) || 0;
    }

    s.rincian.push({ l: 'No. pinjaman', v: String(row.id_pinjaman) });
    if (posisi >= 0) {
      s.rincian.push({ l: 'Angsuran ke', v: String(posisi + 1) +
        (pinjaman && !menurun ? ' dari ' + pinjaman.tenor : '') });
    }
    s.rincian.push({ l: 'Pokok', v: Number(row.angsuran_pokok), rp: true });
    s.rincian.push({ l: 'Jasa', v: Number(row.jasa), rp: true });
    if (Number(row.denda) > 0) s.rincian.push({ l: 'Denda', v: Number(row.denda), rp: true });
    s.utama = { l: 'TOTAL BAYAR', v: Number(row.total_bayar) };
    if (pinjaman && posisi >= 0) {
      // Jasa menurun: sisa POKOK = pokok dicairkan s/d saat bayar − pokok terbayar
      const batas = new Date(row.timestamp || row.tanggal);
      const pokokCair = menurun
        ? getBarisPokokPinjaman(pinjaman.id_pinjaman).reduce(function(t, r) {
            return new Date(r.timestamp || r.tanggal) <= batas || !String(r.id_induk || '')
              ? t + (Number(r.nominal) || 0) : t;
          }, 0)
        : Number(pinjaman.total_tagihan);
      const sisa = pokokCair - terbayar;
      s.rincian2 = [{ l: menurun ? 'Sisa pokok' : 'Sisa pinjaman', v: sisa, rp: true }];
      if (sisa <= 0) s.catatan = 'PINJAMAN LUNAS';
    }
  } else {
    const idInduk = String(row.id_induk || '');
    s.judul = idInduk ? 'BUKTI TAMBAHAN PINJAMAN' : 'BUKTI PENCAIRAN PINJAMAN';
    if (idInduk) s.rincian.push({ l: 'Digabung ke', v: idInduk });
    s.rincian.push({ l: 'Tenor', v: row.tenor + ' bulan' });
    s.utama = { l: 'POKOK DICAIRKAN', v: Number(row.nominal) };
    if (isJasaMenurun(row)) {
      s.rincian.push({ l: 'Jasa', v: row.jasa_persen + '%/bln dr sisa' });
      s.rincian2 = [{ l: 'Angsuran pokok/bln', v: Number(row.angsuran_perbulan), rp: true }];
      if (idInduk) {
        s.rincian2.unshift({ l: 'Total sisa pokok', v: getSisaPinjaman(idInduk), rp: true });
      }
    } else {
      s.rincian.push({ l: 'Jasa', v: row.jasa_persen + '% / bulan' });
      s.rincian2 = [
        { l: 'Total tagihan', v: Number(row.total_tagihan), rp: true },
        { l: 'Angsuran/bulan', v: Number(row.angsuran_perbulan), rp: true }
      ];
    }
  }
  if (!s.rincian2) s.rincian2 = [];
  if (!s.catatan && row.keterangan) s.catatan = String(row.keterangan);
  return s;
}

function apiStruk(d) {
  requireRole(['admin', 'operator']);
  MODE_SENYAP = true;
  return dataStruk(d.jenis, d.id);
}
