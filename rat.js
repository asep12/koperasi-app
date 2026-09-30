/**
 * ============================================================
 * RAT.GS — Laporan Rapat Anggota Tahunan (file Word .docx)
 * Koperasi App
 * ============================================================
 *
 * Menyusun paket laporan RAT (mengikuti "Pembukuan Koperasi 2025.docx"):
 *   1. Undangan RAT            4. Laporan pertanggungjawaban pengurus
 *   2. Susunan acara           5. Laporan badan pengawas + analisis keuangan
 *   3. Tata tertib             6. Rencana kerja tahun berikutnya
 *
 * Angka (SHU + terbilang, keanggotaan, aktiva/pasiva, likuiditas,
 * solvabilitas, rentabilitas, nominal simpanan, jasa, pembagian SHU)
 * diambil dari data aplikasi. Kalimat-kalimat tetap bisa disunting di Word.
 *
 * File .docx disusun langsung sebagai XML (WordprocessingML) lalu di-zip
 * dengan Utilities.zip — tanpa pustaka luar, tanpa izin Google Docs.
 * ============================================================
 */

const FOLDER_RAT = 'Laporan RAT Koperasi';
const KOLOM_RAT = ['nomor_undangan_rat', 'tanggal_undangan_rat', 'waktu_rat', 'tempat_rat',
  'acara_tambahan_rat', 'predikat_kesehatan', 'pinjaman_maksimal', 'tenor_maksimal'];
const BAWAAN_RAT = {
  waktu_rat: '13.00 WIB s.d. selesai', tempat_rat: '',
  acara_tambahan_rat: 'Door Prize', predikat_kesehatan: 'SEHAT',
  pinjaman_maksimal: 60000000, tenor_maksimal: 30
};

// ============================================================
// DATA
// ============================================================

/** Data untuk form & laporan RAT satu tahun buku (tanpa menulis apa pun). */
function apiDataRAT(tahun) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  return kumpulkanDataRAT_(Number(tahun) || tahunRATBawaan_());
}

/** RAT biasanya digelar awal tahun untuk tahun buku sebelumnya. */
function tahunRATBawaan_() {
  const aktif = getTahunAktif() || new Date().getFullYear();
  if (new Date().getMonth() < 6 && getSettingRAT(aktif - 1)) return aktif - 1;
  return aktif;
}

function kumpulkanDataRAT_(tahun) {
  const setting = getSettingRAT(tahun);
  if (!setting) throw new Error('Tahun buku ' + tahun + ' belum ada di Pengaturan.');
  const rencana = getSettingRAT(tahun + 1) || setting;
  const nominalRencana = getNominalSimpanan(getSettingRAT(tahun + 1) ? tahun + 1 : tahun);
  const profil = getProfilKoperasi();
  const nr = hitungNeraca(tahun);
  const isian = {};
  KOLOM_RAT.forEach(function(k) {
    const v = setting[k];
    isian[k] = v === undefined || v === null || v === '' ? (BAWAAN_RAT[k] !== undefined ? BAWAAN_RAT[k] : '') : v;
  });
  if (!isian.tempat_rat) isian.tempat_rat = profil.instansi || '';
  isian.tanggal_rat = tglISO_(setting.tanggal_rat);
  isian.tanggal_undangan_rat = tglISO_(isian.tanggal_undangan_rat);
  isian.pinjaman_maksimal = Number(isian.pinjaman_maksimal) || BAWAAN_RAT.pinjaman_maksimal;
  isian.tenor_maksimal = Number(isian.tenor_maksimal) || BAWAAN_RAT.tenor_maksimal;

  return {
    tahun: tahun,
    profil: profil,
    isian: isian,
    anggota: mutasiAnggota_(tahun),
    neraca: nr,
    analisis: analisisKeuangan_(nr),
    nominal: getNominalSimpanan(tahun),
    rencana: {
      tahun: tahun + 1,
      nominal: nominalRencana,
      jasaSukarela: Number(rencana.jasa_sukarela) || 0,
      metodeJasa: metodeJasaSukarela(rencana),
      jasaPinjaman: Number(rencana.jasa_pinjaman) || 0,
      shuSimpanan: Number(rencana.shu_simpanan) || 0,
      shuJasa: Number(rencana.shu_jasa) || 0
    }
  };
}

function tglISO_(v) {
  if (!v) return '';
  const d = v instanceof Date ? v : new Date(v);
  if (isNaN(d.getTime())) return '';
  const p = function(n) { return (n < 10 ? '0' : '') + n; };
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
}

/**
 * Keanggotaan satu tahun buku: masuk dari tanggal_masuk, keluar dari
 * catatan log "status → keluar". Anggota berstatus keluar tanpa catatan
 * tanggal dianggap keluar sebelum tahun tersebut.
 */
function mutasiAnggota_(tahun) {
  const akhirTahun = new Date(tahun, 11, 31, 23, 59, 59);
  const awalTahun = new Date(tahun, 0, 1);
  const tglKeluar = {};
  sheetToObjects(SHEET.LOG).forEach(function(l) {
    if (String(l.sheet_target) !== SHEET.ANGGOTA || String(l.aksi) !== 'EDIT') return;
    let baru = {};
    try { baru = JSON.parse(l.data_baru || '{}'); } catch (e) { return; }
    if (baru.status === 'keluar') tglKeluar[String(l.id_referensi)] = new Date(l.timestamp);
  });
  let akhir = 0, masuk = 0, keluar = 0;
  sheetToObjects(SHEET.ANGGOTA).forEach(function(a) {
    const id = String(a.id_anggota);
    const tMasuk = a.tanggal_masuk ? new Date(a.tanggal_masuk) : null;
    if (tMasuk && !isNaN(tMasuk.getTime()) && tMasuk > akhirTahun) return;   // baru masuk setelahnya
    let tKeluar = null;
    if (a.status === 'keluar') tKeluar = tglKeluar[id] || new Date(0);
    if (tMasuk && tMasuk >= awalTahun && tMasuk <= akhirTahun) masuk++;
    if (tKeluar && tKeluar >= awalTahun && tKeluar <= akhirTahun) keluar++;
    if (!tKeluar || tKeluar > akhirTahun) akhir++;
  });
  return { awal: akhir - masuk + keluar, masuk: masuk, keluar: keluar, akhir: akhir };
}

/** Likuiditas, solvabilitas, rentabilitas — seperti laporan badan pengawas. */
function analisisKeuangan_(nr) {
  const jml = function(a) { return a.reduce(function(s, r) { return s + (Number(r.nilai) || 0); }, 0); };
  const aktivaLancar = jml(nr.aktivaLancar);
  const aktivaTetap = jml(nr.aktivaTetap);
  const utangPendek = jml(nr.hutangLancar) + jml(nr.danaDana);
  const utangPanjang = 0;
  const modalSendiri = jml(nr.modalSendiri);
  const persen = function(a, b) { return b ? Math.round(a / b * 10000) / 100 : 0; };
  return {
    aktivaLancar: aktivaLancar, aktivaTetap: aktivaTetap, totalAktiva: nr.totalAktiva,
    utangPendek: utangPendek, utangPanjang: utangPanjang, modalSendiri: modalSendiri,
    shu: nr.shuBerjalan,
    likuiditas: persen(aktivaLancar, utangPendek),
    solvabilitas: persen(nr.totalAktiva, utangPendek + utangPanjang),
    rentabilitas: persen(nr.shuBerjalan, modalSendiri)
  };
}

// ============================================================
// BUAT FILE
// ============================================================

/**
 * Simpan isian RAT ke setting tahun buku, susun .docx, simpan ke Google Drive
 * (folder "Laporan RAT Koperasi"), kembalikan juga isi file untuk diunduh.
 */
function apiBuatLaporanRAT(d) {
  requireRole(['admin']);
  MODE_SENYAP = true;
  d = d || {};
  const tahun = Number(d.tahun);
  if (!getSettingRAT(tahun)) throw new Error('Tahun buku ' + tahun + ' belum ada di Pengaturan.');
  if (!d.tanggal_rat) throw new Error('Tanggal RAT wajib diisi.');

  // simpan isian supaya tahun depan tinggal menyesuaikan
  tambahKolomBila_(SHEET.SETTING_RAT, KOLOM_RAT);
  const simpan = { tanggal_rat: new Date(d.tanggal_rat + 'T00:00:00') };
  KOLOM_RAT.forEach(function(k) { if (d[k] !== undefined) simpan[k] = String(d[k]).trim(); });
  if (d.tanggal_undangan_rat) simpan.tanggal_undangan_rat = new Date(d.tanggal_undangan_rat + 'T00:00:00');
  simpan.pinjaman_maksimal = Number(d.pinjaman_maksimal) || BAWAAN_RAT.pinjaman_maksimal;
  simpan.tenor_maksimal = Number(d.tenor_maksimal) || BAWAAN_RAT.tenor_maksimal;
  updateRowByField(SHEET.SETTING_RAT, 'tahun', tahun, simpan);
  hapusCache(SHEET.SETTING_RAT);

  const data = kumpulkanDataRAT_(tahun);
  // angka keanggotaan boleh dikoreksi manual di form
  ['awal', 'masuk', 'keluar', 'akhir'].forEach(function(k) {
    const v = d['anggota_' + k];
    if (v !== undefined && v !== '' && !isNaN(Number(v))) data.anggota[k] = Number(v);
  });

  const blob = susunDocxRAT_(data);
  const namaFile = 'Laporan RAT Tahun Buku ' + tahun + '.docx';
  blob.setName(namaFile);

  let url = '';
  try {
    const cari = DriveApp.getFoldersByName(FOLDER_RAT);
    const folder = cari.hasNext() ? cari.next() : DriveApp.createFolder(FOLDER_RAT);
    const lama = folder.getFilesByName(namaFile);
    while (lama.hasNext()) lama.next().setTrashed(true);   // ganti versi sebelumnya
    url = folder.createFile(blob).getUrl();
  } catch (e) {
    url = '';   // Drive gagal → file tetap bisa diunduh langsung
  }
  logAktivitas('INSERT', SHEET.SETTING_RAT, String(tahun), null, { laporan_rat: namaFile });
  return {
    nama: namaFile, url: url,
    base64: Utilities.base64Encode(blob.getBytes()),
    ringkas: { shu: data.analisis.shu, anggota: data.anggota.akhir,
      likuiditas: data.analisis.likuiditas, solvabilitas: data.analisis.solvabilitas,
      rentabilitas: data.analisis.rentabilitas, balance: data.neraca.balance }
  };
}

// ============================================================
// FORMAT ANGKA & TANGGAL
// ============================================================

const BULAN_RAT = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus',
  'September', 'Oktober', 'November', 'Desember'];
const HARI_RAT = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', "Jum'at", 'Sabtu'];

function tglPanjang_(iso, pakaiHari) {
  if (!iso) return '';
  const d = new Date(iso + 'T00:00:00');
  if (isNaN(d.getTime())) return '';
  return (pakaiHari ? HARI_RAT[d.getDay()] + ', ' : '') + d.getDate() + ' ' + BULAN_RAT[d.getMonth()] + ' ' + d.getFullYear();
}

/** 1234567.5 → "1.234.567,50"; bilangan bulat tanpa sen. */
function angkaRAT_(n) {
  n = Number(n) || 0;
  const neg = n < 0; n = Math.abs(n);
  const bulat = Math.floor(n + 1e-9);
  const sen = Math.round((n - bulat) * 100);
  const s = String(bulat).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return (neg ? '-' : '') + s + (sen ? ',' + (sen < 10 ? '0' : '') + sen : '');
}
function rpRAT_(n) { return 'Rp ' + angkaRAT_(n); }
function persenRAT_(n) { return String(Math.round(n * 100) / 100).replace('.', ',') + ' %'; }

/** Terbilang bahasa Indonesia, mis. 1250000 → "satu juta dua ratus lima puluh ribu rupiah". */
function terbilang(n) {
  const satuan = ['', 'satu', 'dua', 'tiga', 'empat', 'lima', 'enam', 'tujuh', 'delapan', 'sembilan',
    'sepuluh', 'sebelas'];
  const t = function(x) {
    if (x < 12) return satuan[x];
    if (x < 20) return satuan[x - 10] + ' belas';
    if (x < 100) return satuan[Math.floor(x / 10)] + ' puluh' + (x % 10 ? ' ' + satuan[x % 10] : '');
    if (x < 200) return 'seratus' + (x - 100 ? ' ' + t(x - 100) : '');
    if (x < 1000) return satuan[Math.floor(x / 100)] + ' ratus' + (x % 100 ? ' ' + t(x % 100) : '');
    if (x < 2000) return 'seribu' + (x - 1000 ? ' ' + t(x - 1000) : '');
    if (x < 1e6) return t(Math.floor(x / 1000)) + ' ribu' + (x % 1000 ? ' ' + t(x % 1000) : '');
    if (x < 1e9) return t(Math.floor(x / 1e6)) + ' juta' + (x % 1e6 ? ' ' + t(x % 1e6) : '');
    if (x < 1e12) return t(Math.floor(x / 1e9)) + ' miliar' + (x % 1e9 ? ' ' + t(x % 1e9) : '');
    return t(Math.floor(x / 1e12)) + ' triliun' + (x % 1e12 ? ' ' + t(x % 1e12) : '');
  };
  n = Math.abs(Number(n) || 0);
  const bulat = Math.floor(n + 1e-9);
  const sen = Math.round((n - bulat) * 100);
  let s = bulat === 0 ? 'nol' : t(bulat);
  s += ' rupiah';
  if (sen) s += ' ' + t(sen) + ' sen';
  return s.replace(/\s+/g, ' ').trim();
}

// ============================================================
// PENYUSUN WORDPROCESSINGML
// ============================================================

const F4_LEBAR = 12240, F4_TINGGI = 18720;          // 8,5 × 13 inci (F4), satuan twip
const TEPI = { atas: 1134, bawah: 1134, kiri: 1701, kanan: 1134 };   // 2 cm; kiri 3 cm
const LEBAR_ISI = F4_LEBAR - TEPI.kiri - TEPI.kanan; // 9405 twip

function xmlEsc_(s) {
  return String(s === undefined || s === null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** Potongan XML Word (hasil R_/P_) vs teks biasa dari pengguna. */
function apakahXml_(s) { return typeof s === 'string' && s.indexOf('<w:') === 0; }

/** Satu run teks. o: {b, i, u, sz (pt), caps} */
function R_(teks, o) {
  o = o || {};
  let pr = '';
  if (o.b) pr += '<w:b/>';
  if (o.i) pr += '<w:i/>';
  if (o.caps) pr += '<w:caps/>';
  if (o.u) pr += '<w:u w:val="single"/>';
  if (o.sz) pr += '<w:sz w:val="' + Math.round(o.sz * 2) + '"/><w:szCs w:val="' + Math.round(o.sz * 2) + '"/>';
  const bagian = String(teks === undefined || teks === null ? '' : teks).split('\t');
  return bagian.map(function(t, i) {
    return (i ? '<w:r>' + (pr ? '<w:rPr>' + pr + '</w:rPr>' : '') + '<w:tab/></w:r>' : '') +
      (t ? '<w:r>' + (pr ? '<w:rPr>' + pr + '</w:rPr>' : '') + '<w:t xml:space="preserve">' + xmlEsc_(t) + '</w:t></w:r>' : '');
  }).join('');
}

/**
 * Paragraf. isi: string (teks biasa) atau hasil R_() digabung.
 * o: {al: 'center'|'right'|'both', sb, sa (pt spasi sebelum/sesudah), ind (twip kiri), hang (twip gantung),
 *     tabs: [{pos, al, leader}], b, sz, pb (halaman baru), kn (tetap bersama berikutnya), garis (garis bawah)}
 */
function P_(isi, o) {
  o = o || {};
  if (!apakahXml_(isi)) isi = R_(isi, { b: o.b, i: o.i, sz: o.sz, u: o.u });
  let pr = '';
  if (o.kn) pr += '<w:keepNext/>';
  if (o.pb) pr += '<w:pageBreakBefore/>';
  if (o.garis) pr += '<w:pBdr><w:bottom w:val="' + (o.garis === 'ganda' ? 'double' : 'single') + '" w:sz="' +
    (o.garis === 'ganda' ? 6 : 8) + '" w:space="4" w:color="000000"/></w:pBdr>';
  if (o.tabs) pr += '<w:tabs>' + o.tabs.map(function(t) {
    return '<w:tab w:val="' + (t.al || 'left') + '"' + (t.leader ? ' w:leader="' + t.leader + '"' : '') + ' w:pos="' + t.pos + '"/>';
  }).join('') + '</w:tabs>';
  pr += '<w:spacing w:before="' + Math.round((o.sb || 0) * 20) + '" w:after="' + Math.round((o.sa === undefined ? 6 : o.sa) * 20) + '"' +
    (o.rapat ? ' w:line="240" w:lineRule="auto"' : '') + '/>';
  if (o.ind || o.hang) pr += '<w:ind w:left="' + ((o.ind || 0) + (o.hang || 0)) + '"' + (o.hang ? ' w:hanging="' + o.hang + '"' : '') + '/>';
  if (o.al) pr += '<w:jc w:val="' + o.al + '"/>';
  return '<w:p><w:pPr>' + pr + '</w:pPr>' + isi + '</w:p>';
}

/** Item bernomor dengan indentasi gantung: "1.  teks yang panjang ...". */
function NO_(nomor, teks, o) {
  o = o || {};
  const ind = o.ind || 0, lebarNo = o.lebarNo || 425;
  return P_(R_(nomor + '\t', { b: o.b }) + (apakahXml_(teks) ? teks : R_(teks, { b: o.b })),
    { ind: ind, hang: lebarNo, tabs: [{ pos: ind + lebarNo }], sa: o.sa === undefined ? 3 : o.sa, al: o.al || 'both' });
}

/** Tabel. baris: array baris, tiap sel {isi: xml paragraf, lebar, al, b, span, atas (garis atas), bawah}. */
function T_(lebar, baris, o) {
  o = o || {};
  const total = lebar.reduce(function(a, b) { return a + b; }, 0);
  const bingkai = o.garis
    ? '<w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/><w:left w:val="single" w:sz="4" w:color="000000"/>' +
      '<w:bottom w:val="single" w:sz="4" w:color="000000"/><w:right w:val="single" w:sz="4" w:color="000000"/>' +
      '<w:insideH w:val="single" w:sz="4" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:color="000000"/></w:tblBorders>'
    : '<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>' +
      '<w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>';
  let x = '<w:tbl><w:tblPr><w:tblW w:w="' + total + '" w:type="dxa"/>' +
    (o.ind ? '<w:tblInd w:w="' + o.ind + '" w:type="dxa"/>' : '') + bingkai +
    '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="' + (o.pad === undefined ? 60 : o.pad) + '" w:type="dxa"/>' +
    '<w:right w:w="' + (o.pad === undefined ? 60 : o.pad) + '" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>' +
    lebar.map(function(w) { return '<w:gridCol w:w="' + w + '"/>'; }).join('') + '</w:tblGrid>';
  baris.forEach(function(row) {
    x += '<w:tr><w:trPr><w:cantSplit/></w:trPr>';
    let kol = 0;
    row.forEach(function(sel) {
      if (typeof sel === 'string' || typeof sel === 'number') sel = { t: String(sel) };
      const span = sel.span || 1;
      let w = 0; for (let i = 0; i < span; i++) w += lebar[kol + i];
      kol += span;
      let tcPr = '<w:tcW w:w="' + w + '" w:type="dxa"/>' + (span > 1 ? '<w:gridSpan w:val="' + span + '"/>' : '');
      if (sel.atas || sel.bawah) tcPr += '<w:tcBorders>' + (sel.atas ? '<w:top w:val="single" w:sz="8" w:color="000000"/>' : '') +
        (sel.bawah ? '<w:bottom w:val="single" w:sz="8" w:color="000000"/>' : '') + '</w:tcBorders>';
      if (sel.vmid) tcPr += '<w:vAlign w:val="center"/>';
      if (sel.arsir) tcPr += '<w:shd w:val="clear" w:color="auto" w:fill="E7ECE9"/>';
      let isi = sel.isi || P_(R_(sel.t || '', { b: sel.b, i: sel.i }), { al: sel.al, sa: sel.sa === undefined ? 1 : sel.sa, sb: sel.sb || 1 });
      if (o.kn) isi = isi.replace(/<w:pPr>/g, '<w:pPr><w:keepNext/>');
      x += '<w:tc><w:tcPr>' + tcPr + '</w:tcPr>' + isi + '</w:tc>';
    });
    x += '</w:tr>';
  });
  return x + '</w:tbl>' + (o.tanpaSpasi ? '' : P_('', { sa: 2, kn: o.kn }));
}

/** Blok tanda tangan: satu atau dua kolom. kolom: [{atas, jabatan, nama}] */
function TTD_(kolom, o) {
  o = o || {};
  const lebar = kolom.length === 1 ? [LEBAR_ISI - 4200, 4200] : [LEBAR_ISI / 2, LEBAR_ISI / 2];
  const sel = function(k) {
    return { isi: P_(R_(k.atas || ''), { al: 'center', sa: 0 }) + P_(R_(k.jabatan || ''), { al: 'center', sa: 0 }) +
      P_('', { sa: 0 }) + P_('', { sa: 0 }) + P_('', { sa: 0 }) +
      P_(R_(k.nama || '(................................................)', { b: true, u: !!k.nama }), { al: 'center', sa: 0 }) };
  };
  const baris = kolom.length === 1 ? [[{ t: '' }, sel(kolom[0])]] : [kolom.map(sel)];
  // blok tanda tangan tidak boleh terpisah dari baris di atasnya
  return P_('', { sa: o.sb === undefined ? 6 : o.sb, kn: true }) + T_(lebar, baris, { kn: true });
}

/** Judul bagian (halaman baru). sub: string atau array baris. */
function JUDUL_(teks, sub, o) {
  o = o || {};
  const baris = [].concat(sub || []);
  return P_(R_(teks, { b: true, sz: 13 }), { al: 'center', sa: 0, pb: o.pb, kn: true }) +
    baris.map(function(t, i) {
      return P_(R_(t, { b: true, sz: 12 }), { al: 'center', sa: i === baris.length - 1 ? 12 : 0, kn: true });
    }).join('') + (baris.length ? '' : P_('', { sa: 6 }));
}
function SUBJ_(teks, o) {
  o = o || {};
  return P_(R_(teks, { b: true }), { sb: o.sb === undefined ? 8 : o.sb, sa: 4, kn: true });
}

/** Pecahan untuk analisis rasio: [label pembilang / penyebut] × 100% = hasil */
function PECAHAN_(atas, bawah, hasil) {
  const lb = [3700, 1000, 3000];
  return T_(lb, [
    [{ t: atas, al: 'center', sa: 0 }, { t: '× 100 %', al: 'center', sa: 0 }, { t: '= ' + hasil, b: true, sa: 0 }],
    [{ t: bawah, al: 'center', atas: true, sa: 0 }, { t: '' }, { t: '' }]
  ], { ind: 567 });
}

// ============================================================
// ISI LAPORAN
// ============================================================

function susunDocxRAT_(D) {
  const T = D.tahun, P = D.profil, I = D.isian, A = D.anggota, N = D.neraca, K = D.analisis;
  const kota = P.kota || '..........';
  const instansi = P.instansi || 'sekolah/instansi';
  const nama = 'Koperasi ' + P.nama_koperasi;
  const tglRAT = tglPanjang_(I.tanggal_rat, true);
  const bulanRAT = (function() {
    const d = new Date(I.tanggal_rat + 'T00:00:00');
    return isNaN(d.getTime()) ? '' : BULAN_RAT[d.getMonth()] + ' ' + d.getFullYear();
  })();
  const tglTtd = kota + ', ' + (tglPanjang_(I.tanggal_rat) || bulanRAT);
  const pengurusTTD = [
    { jabatan: 'Ketua', nama: P.ketua }, { jabatan: 'Sekretaris', nama: P.sekretaris }];
  let x = '';

  // ---------- 1. UNDANGAN ----------
  x += kopSurat_(P);
  x += T_([1300, 300, 3900, 3900], [
    ['Nomor', ':', I.nomor_undangan_rat || '......./KOP/...../' + (T + 1), { t: kota + ', ' + (tglPanjang_(I.tanggal_undangan_rat) || '.................'), al: 'right' }],
    ['Lampiran', ':', '-', ''],
    ['Hal', ':', { t: 'Undangan Rapat Anggota Tahunan', b: true }, '']
  ]);
  x += P_('Kepada Yth.', { ind: 5100, sa: 0, sb: 6 });
  x += P_('Bapak/Ibu Pembina dan Anggota', { ind: 5100, sa: 0 });
  x += P_(nama, { ind: 5100, sa: 0 });
  x += P_('di Tempat', { ind: 5100, sa: 14 });
  x += P_(R_('Assalamu’alaikum Wr. Wb.', { i: true }), { sa: 6 });
  x += P_('Sehubungan dengan akan dilaksanakannya Rapat Anggota Tahunan (RAT) ' + nama + ' Tahun Buku ' + T +
    ', kami pengurus koperasi mengundang Bapak/Ibu untuk hadir pada:', { al: 'both' });
  x += T_([2300, 300, 6200], [
    ['Hari / Tanggal', ':', tglRAT || '.......................'],
    ['Waktu', ':', I.waktu_rat],
    ['Tempat', ':', I.tempat_rat],
    ['Acara', ':', { isi: P_('1.  Laporan Pertanggungjawaban Pengurus Tahun Buku ' + T, { sa: 0, sb: 1 }) +
      (I.acara_tambahan_rat ? P_('2.  ' + I.acara_tambahan_rat, { sa: 1 }) : '') }]
  ], { ind: 567 });
  x += P_('Mengingat pentingnya acara tersebut, kami mengharapkan kehadiran Bapak/Ibu tepat pada waktunya. ' +
    'Atas perhatian dan kehadirannya kami ucapkan terima kasih.', { al: 'both' });
  x += P_(R_('Wassalamu’alaikum Wr. Wb.', { i: true }), { sa: 4 });
  x += P_('Pengurus ' + nama, { al: 'center', sb: 8, sa: 0 });
  x += TTD_(pengurusTTD, { sb: 0 });

  // ---------- 2. SUSUNAN ACARA ----------
  x += JUDUL_('SUSUNAN ACARA RAPAT ANGGOTA TAHUNAN (RAT)', 'TAHUN BUKU ' + T, { pb: true });
  [
    'Pembukaan', 'Pengesahan Acara', 'Sambutan Ketua ' + nama,
    'Sambutan Kepala ' + instansi, 'Pembacaan Tata Tertib RAT',
    'Laporan Pertanggungjawaban Pengurus ' + nama, 'Laporan Badan Pengawas ' + nama,
    'Pandangan Umum', 'Pengesahan Pertanggungjawaban Pengurus Tahun Buku ' + T,
    'Program Kerja dan RAB Tahun ' + (T + 1), 'Penutup'
  ].forEach(function(a, i) { x += NO_((i + 1) + '.', a, { ind: 567, sa: 6 }); });
  x += P_(tglTtd, { al: 'right', sb: 12, sa: 0 });
  x += TTD_(pengurusTTD, { sb: 0 });

  // ---------- 3. TATA TERTIB ----------
  const kuorum = Math.floor(A.akhir * 3 / 4) + 1;
  x += JUDUL_('TATA TERTIB RAPAT ANGGOTA TAHUNAN (RAT)', 'TAHUN BUKU ' + T, { pb: true });
  const tt = [
    'Rapat ini adalah Rapat Anggota Tahunan ' + nama + (P.instansi ? ' ' + P.instansi : '') + ' Tahun Buku ' + T + '.',
    'Pimpinan Rapat adalah Pengurus ' + nama + '.',
    ['Peserta Rapat terdiri dari:', ['Semua Pengurus;', 'Ketua dan Anggota Badan Pengawas;',
      'Semua anggota yang pada tanggal 31 Desember ' + T + ' masih aktif menjadi anggota ' + nama + '.']],
    'Hak bicara bagi anggota diatur berdasarkan kebijaksanaan Pimpinan Rapat.',
    'Rapat Anggota Tahunan ini sah apabila dihadiri oleh lebih dari ¾ jumlah anggota, yaitu sekurang-kurangnya ' +
      kuorum + ' orang dari ' + A.akhir + ' anggota.',
    'Keputusan dianggap sah apabila disetujui oleh sekurang-kurangnya ¾ dari anggota yang hadir.',
    'Setiap peserta wajib menjaga ketertiban selama RAT berlangsung.',
    'Dalam menggunakan hak suara, anggota harus obyektif, ringkas, dan jelas.',
    ['Pimpinan Rapat berhak:', ['Mengambil kebijaksanaan untuk mengatasi segala sesuatu yang dapat mengganggu kelancaran jalannya rapat;',
      'Mengambil tindakan seperlunya terhadap peserta yang tidak mengindahkan tata tertib, berupa teguran dan peringatan.']],
    'Hal-hal yang dianggap perlu namun belum tercantum dalam tata tertib ini akan diatur dan ditentukan kemudian berdasarkan kesepakatan peserta rapat.'
  ];
  tt.forEach(function(item, i) {
    if (typeof item === 'string') { x += NO_((i + 1) + '.', item, { sa: 5 }); return; }
    x += NO_((i + 1) + '.', item[0], { sa: 3 });
    item[1].forEach(function(s, j) { x += NO_(String.fromCharCode(97 + j) + '.', s, { ind: 425, sa: 3 }); });
  });
  x += P_(tglTtd, { al: 'right', sb: 12, sa: 0 });
  x += TTD_(pengurusTTD, { sb: 0 });

  // ---------- 4. LAPORAN PENGURUS ----------
  x += JUDUL_('LAPORAN PERTANGGUNGJAWABAN PENGURUS', [nama.toUpperCase(), 'TAHUN BUKU ' + T], { pb: true });
  x += P_(R_('Assalamu’alaikum Wr. Wb.', { i: true }), { sa: 4 });
  x += P_('Yang terhormat:', { sa: 0 });
  x += NO_('•', 'Kepala ' + instansi + ';', { sa: 0 });
  x += NO_('•', 'Saudara-saudara Pengurus, Badan Pengawas, para Anggota, dan Undangan.', { sa: 6 });
  x += P_('Puji syukur kita panjatkan kepada Allah SWT, karena berkat inayah-Nya pada hari ini kita dapat hadir bersama dalam ' +
    'Rapat Anggota Tahunan ' + nama + '. Shalawat beserta salam semoga selalu tercurah kepada Nabi Muhammad SAW, keluarga, ' +
    'sahabat, serta para pengikutnya hingga akhir zaman.', { al: 'both' });
  x += P_('Kami mengucapkan terima kasih kepada Bapak, Ibu, dan Saudara atas kehadirannya memenuhi undangan kami. ' +
    'Laporan ini kami sajikan dengan sistematika sebagai berikut.', { al: 'both' });

  x += SUBJ_('I.  PENDAHULUAN');
  x += P_('Rapat Anggota Tahunan (RAT) adalah forum tertinggi pengambilan keputusan dalam koperasi, sehingga wajib dilaksanakan ' +
    'setiap tahun. Laporan pertanggungjawaban ini disampaikan kepada anggota melalui forum RAT, terutama dititikberatkan pada ' +
    'kegiatan yang telah diprogramkan pada RAT tahun lalu.', { al: 'both', ind: 425 });
  x += P_('Susunan pengurus dan badan pengawas:', { ind: 425, sa: 2 });
  // satu nama per baris (koma tidak dipakai sebagai pemisah karena ada di gelar, mis. "S.Ag, M.Pd")
  const perBaris = function(v) { return String(v || '').split(/\r?\n|;/).map(function(s) { return s.trim(); }).filter(String); };
  const anggotaPengurus = perBaris(P.pengurus_anggota);
  const pengawasAnggota = perBaris(P.pengawas_anggota);
  const susunan = [['Ketua', P.ketua], ['Sekretaris', P.sekretaris], ['Bendahara', P.bendahara]];
  anggotaPengurus.forEach(function(n, i) { susunan.push([i ? '' : 'Anggota', n]); });
  susunan.push(['Badan Pengawas', (P.pengawas_ketua ? P.pengawas_ketua + ' (Ketua)' : '')]);
  pengawasAnggota.forEach(function(n) { susunan.push(['', n]); });
  x += T_([2600, 300, 5900], susunan.filter(function(s) { return s[1]; }).map(function(s) {
    return [s[0], s[0] ? ':' : '', s[1]];
  }), { ind: 425 });
  x += P_('Selama satu tahun kami berusaha melaksanakan program yang telah ditetapkan, namun masih ada sebagian program yang belum ' +
    'dapat dilaksanakan mengingat berbagai keterbatasan.', { al: 'both', ind: 425 });

  x += SUBJ_('II.  PELAKSANAAN DAN EVALUASI PROGRAM');
  x += P_('Kegiatan yang telah kami lakukan adalah sebagai berikut.', { ind: 425 });
  const bidang = [
    ['Bidang Administrasi dan Organisasi', 'Menyempurnakan kelengkapan administrasi dan melaksanakan konsolidasi organisasi dalam bentuk mekanisme kerja. Pencatatan transaksi dilakukan melalui aplikasi koperasi.'],
    ['Bidang Usaha', 'Melanjutkan usaha simpan pinjam bagi anggota.'],
    ['Bidang Kesejahteraan', 'Berjalan sesuai dengan harapan.'],
    ['Bidang Permodalan', 'Simpanan pokok dan wajib berjalan sesuai harapan. Simpanan sukarela terus digalakkan. Sisa Hasil Usaha (SHU) tahun ' + T +
      ' sebesar ' + rpRAT_(K.shu) + ' (' + kapital_(terbilang(K.shu)) + '). Neraca terlampir untuk dapat ditelaah para anggota.'],
    ['Bidang Sarana', 'Tempat dan sarana kegiatan koperasi memanfaatkan fasilitas sekolah.'],
    ['Keanggotaan', 'Jumlah anggota per 31 Desember ' + T + ' sebanyak ' + A.akhir + ' orang.']
  ];
  bidang.forEach(function(b, i) {
    x += NO_(String.fromCharCode(97 + i) + '.', R_(b[0], { b: true }), { ind: 425, sa: 1 });
    x += P_(b[1], { ind: 850, al: 'both', sa: 5 });
  });

  x += SUBJ_('III.  PENUTUP');
  x += P_((I.predikat_kesehatan ? 'Berdasarkan penilaian kesehatan koperasi oleh Dinas Koperasi, UKM, Perindustrian dan Perdagangan ' +
    (P.kota ? 'Kota ' + P.kota : 'setempat') + ', Alhamdulillah ' + nama + ' mendapat predikat “' + I.predikat_kesehatan + '”. Namun demikian, ' : '') +
    'mengingat keterbatasan kemampuan dan waktu, tentu masih terdapat kekurangan dalam kegiatan koperasi selama kepengurusan kami. ' +
    'Dengan segala kerendahan hati kami mohon maaf, dan semoga laporan pertanggungjawaban ini dapat diterima dan disahkan.',
    { al: 'both', ind: 425 });
  x += P_('Akhirnya kami menghaturkan terima kasih kepada semua pihak yang telah membantu jalannya koperasi kita.', { al: 'both', ind: 425 });
  x += P_(R_('Wabillahi taufik wal hidayah. Wassalamu’alaikum Wr. Wb.', { i: true }), { ind: 425, kn: true });
  x += TTD_([{ atas: tglTtd, jabatan: 'Ketua', nama: P.ketua }]);

  // ---------- 5. LAPORAN BADAN PENGAWAS ----------
  x += JUDUL_('LAPORAN BADAN PENGAWAS', 'PADA RAT TAHUN BUKU ' + T, { pb: true });
  x += P_(R_('Assalamu’alaikum Wr. Wb.', { i: true }), { sa: 4 });
  x += P_('Puji syukur kehadirat Allah SWT yang telah memberikan rahmat dan hidayah-Nya sehingga kita dapat bersama-sama menghadiri ' +
    'Rapat Anggota Tahunan Tahun Buku ' + T + '. Berdasarkan keputusan RAT tahun lalu, Badan Pengawas mendapat mandat dari anggota ' +
    'untuk melaksanakan pengawasan terhadap jalannya koperasi. Hasil pemeriksaan tahun buku ' + T + ' adalah sebagai berikut.', { al: 'both' });

  x += SUBJ_('I.  BIDANG ORGANISASI');
  x += P_('Tegak dan majunya koperasi ditopang oleh organisasi/kepengurusan yang sehat, administrasi yang sehat, dan usaha yang sehat. ' +
    'Keadaan keanggotaan:', { al: 'both', ind: 425 });
  x += T_([5200, 300, 1500], [
    ['Jumlah anggota awal tahun ' + T, ':', { t: A.awal + ' orang', al: 'right' }],
    ['Anggota masuk tahun ' + T, ':', { t: A.masuk + ' orang', al: 'right' }],
    ['Anggota keluar tahun ' + T, ':', { t: A.keluar + ' orang', al: 'right' }],
    [{ t: 'Jumlah anggota akhir tahun ' + T, b: true }, ':', { t: A.akhir + ' orang', b: true, al: 'right', atas: true }]
  ], { ind: 425 });
  x += P_('Pada dasarnya kepengurusan berjalan dengan baik.', { ind: 425 });

  x += SUBJ_('II.  BIDANG ADMINISTRASI');
  x += NO_('a.', R_('Administrasi Organisasi', { b: true }), { ind: 425, sa: 1 });
  x += P_('Buku daftar anggota, daftar pengurus, daftar badan pengawas, notulen rapat, inventaris, agenda, catatan kejadian penting, ' +
    'dan catatan lainnya. Pada dasarnya buku-buku tersebut telah sesuai dengan kenyataan dan beberapa perlu dilengkapi.', { ind: 850, al: 'both' });
  x += NO_('b.', R_('Administrasi Keuangan', { b: true }), { ind: 425, sa: 1 });
  x += P_('Buku kas umum, daftar simpanan kolektif dan perorangan, daftar pinjaman, daftar potongan bulanan, catatan pendukung, ' +
    'dan bundel bukti keuangan. Penyelenggaraan administrasi keuangan berdasarkan bukti-bukti yang sah dan dibukukan dengan tertib.', { ind: 850, al: 'both' });

  x += SUBJ_('III.  BIDANG KEUANGAN');
  x += NO_('a.', R_('Permodalan', { b: true }), { ind: 425, sa: 1 });
  x += T_([4500, 300, 2200], [
    ['Simpanan Pokok', ':', { t: rpRAT_(D.nominal.nominal_pokok) + ',-', al: 'right' }],
    ['Simpanan Wajib (per bulan)', ':', { t: rpRAT_(D.nominal.nominal_wajib) + ',-', al: 'right' }],
    ['Simpanan Sukarela terendah (per bulan)', ':', { t: rpRAT_(D.nominal.minimal_sukarela) + ',-', al: 'right' }]
  ], { ind: 850 });
  x += NO_('b.', R_('Kondisi aktiva dan pasiva per 31 Desember ' + T, { b: true }), { ind: 425, sa: 3 });
  x += tabelNeraca_(N);

  x += SUBJ_('IV.  ANALISIS KEUANGAN');
  x += NO_('1.', R_('Likuiditas', { b: true }) + R_(' (kemampuan melunasi utang jangka pendek)'), { ind: 425, sa: 2 });
  x += T_([5600, 300, 2200], [
    ['Aktiva lancar (kas & piutang pinjaman)', ':', { t: rpRAT_(K.aktivaLancar), al: 'right' }],
    ['Utang jangka pendek (simpanan sukarela & dana-dana)', ':', { t: rpRAT_(K.utangPendek), al: 'right' }]
  ], { ind: 850 });
  x += PECAHAN_('Aktiva Lancar', 'Utang Jangka Pendek', persenRAT_(K.likuiditas));
  x += P_(R_('Artinya: setiap utang yang harus segera dibayar sebesar Rp 1,- dijamin dengan harta lancar sebesar Rp ' +
    kaliRp_(K.likuiditas) + ' (' + (K.likuiditas >= 100 ? 'likuid' : 'kurang likuid') + ').', { i: true }), { ind: 850, al: 'both' });

  x += NO_('2.', R_('Solvabilitas', { b: true }) + R_(' (kemampuan membayar seluruh utang)'), { ind: 425, sa: 2 });
  x += T_([5600, 300, 2200], [
    ['Jumlah aktiva', ':', { t: rpRAT_(K.totalAktiva), al: 'right' }],
    ['Utang jangka pendek + jangka panjang', ':', { t: rpRAT_(K.utangPendek + K.utangPanjang), al: 'right' }]
  ], { ind: 850 });
  x += PECAHAN_('Jumlah Aktiva', 'Jumlah Utang', persenRAT_(K.solvabilitas));
  x += P_(R_('Artinya: setiap kewajiban koperasi sebesar Rp 1,- dijamin oleh harta sebesar Rp ' +
    kaliRp_(K.solvabilitas) + ' (' + (K.solvabilitas >= 100 ? 'solvabel' : 'kurang solvabel') + ').', { i: true }),
    { ind: 850, al: 'both' });

  x += NO_('3.', R_('Rentabilitas', { b: true }) + R_(' (kemampuan modal sendiri menghasilkan SHU)'), { ind: 425, sa: 2 });
  x += T_([5600, 300, 2200], [
    ['Sisa Hasil Usaha tahun ' + T, ':', { t: rpRAT_(K.shu), al: 'right' }],
    ['Modal sendiri (simpanan pokok, wajib, cadangan)', ':', { t: rpRAT_(K.modalSendiri), al: 'right' }]
  ], { ind: 850 });
  x += PECAHAN_('Sisa Hasil Usaha', 'Modal Sendiri', persenRAT_(K.rentabilitas));
  x += P_(R_('Artinya: setiap modal sendiri sebesar Rp 1,- menghasilkan keuntungan sebesar Rp ' +
    kaliRp_(K.rentabilitas) + '.', { i: true }), { ind: 850, al: 'both' });

  x += SUBJ_('V.  KESIMPULAN DAN SARAN');
  x += P_('Jalannya kegiatan koperasi dapat menyejahterakan anggota. Saran-saran:', { ind: 425 });
  x += NO_('a.', 'Agar simpan pinjam dan simpanan sukarela ditingkatkan;', { ind: 425 });
  x += NO_('b.', 'Agar pembagian tugas dalam kepengurusan ditingkatkan.', { ind: 425 });
  x += SUBJ_('VI.  PENUTUP');
  x += P_('Akhirnya kami mohon maaf apabila terdapat kekurangan dan kekhilafan, semoga laporan ini bermanfaat bagi semua pihak.', { ind: 425, al: 'both', kn: true });
  x += P_(R_('Wassalamu’alaikum Wr. Wb.', { i: true }), { ind: 425, kn: true });
  x += TTD_([{ atas: tglTtd, jabatan: 'Badan Pengawas', nama: P.pengawas_ketua }]);

  // ---------- 6. RENCANA KERJA ----------
  const RK = D.rencana;
  x += JUDUL_('RENCANA KERJA', 'TAHUN ' + RK.tahun, { pb: true });
  x += SUBJ_('A.  BIDANG ORGANISASI DAN ADMINISTRASI', { sb: 2 });
  x += NO_('1.', 'Meningkatkan pengelolaan koperasi dalam bidang administrasi, keanggotaan, serta tata niaga secara keseluruhan.', { ind: 425 });
  x += SUBJ_('B.  BIDANG USAHA');
  x += NO_('1.', 'Meningkatkan pelayanan terhadap anggota sesuai ketentuan koperasi.', { ind: 425 });
  x += NO_('2.', 'Meningkatkan pelayanan jumlah pinjaman bagi anggota sesuai kemampuan, dengan kriteria:', { ind: 425 });
  [
    'Permohonan pinjaman diajukan minimal satu bulan sebelumnya;',
    'Besarnya pinjaman maksimal ' + rpRAT_(I.pinjaman_maksimal) + ',-, dapat melebihi apabila keuangan koperasi tersedia;',
    'Pemenuhan permohonan pinjaman disesuaikan dengan keadaan keuangan koperasi;',
    'Pinjaman diangsur tiap bulan dengan cara pemotongan gaji oleh bendahara gaji ' + instansi + ';',
    'Pinjaman diangsur maksimal ' + I.tenor_maksimal + ' bulan dengan jasa ' + String(RK.jasaPinjaman).replace('.', ',') +
      '% per bulan dari sisa pinjaman (menurun);',
    'Batas pembayaran setiap angsuran tanggal 10 tiap bulan;',
    'Apabila peminjam 3 kali berturut-turut tidak dapat membayar sesuai waktu yang ditentukan, setelah dibina pada bulan ke-1 dan ke-2, ' +
      'yang bersangkutan dinyatakan keluar sebagai anggota ' + nama + ';',
    'Peminjam yang masih memiliki tunggakan tidak diberi pinjaman baru kecuali dalam keadaan darurat;',
    'Pinjaman diprioritaskan bagi anggota yang belum pernah meminjam.'
  ].forEach(function(s, i) { x += NO_(String.fromCharCode(97 + i) + '.', s, { ind: 850 }); });
  x += NO_('3.', 'Melayani kebutuhan barang yang bersifat insidental.', { ind: 425 });
  x += NO_('4.', 'Melakukan promosi kepada anggota apabila dana koperasi belum terserap.', { ind: 425 });

  x += SUBJ_('C.  BIDANG KESEJAHTERAAN');
  [
    'Memberikan bantuan transport sebesar Rp 100.000,- untuk menjenguk anggota yang melahirkan atau dirawat di rumah sakit, serta menghadiri undangan anggota;',
    'Memberikan hadiah Idul Fitri sebesar Rp 350.000,- per anggota;',
    'Memberikan insentif bulanan kepada pengurus sesuai keuangan koperasi;',
    'Memberikan uang duduk bagi anggota yang mengikuti RAT sebesar Rp 100.000,- atau menurut kondisi keuangan koperasi;',
    'Memberikan cendera mata kepada semua anggota pada saat RAT;',
    'Pengambilan simpanan sukarela dapat dilakukan 2 kali dalam satu tahun, dengan mendaftar sebulan sebelumnya;',
    RK.metodeJasa === 'harian'
      ? 'Simpanan sukarela mendapat jasa ' + String(RK.jasaSukarela).replace('.', ',') + '% per tahun, dihitung setiap bulan dari saldo akhir bulan sebelumnya;'
      : RK.metodeJasa === 'terendah_tahunan'
      ? 'Setiap bulan simpanan sukarela mendapat jasa ' + String(RK.jasaSukarela).replace('.', ',') +
        '% per tahun dari jumlah simpanan sukarela terendah pada bulan tersebut (disesuaikan dengan jasa bank);'
      : 'Setiap bulan simpanan sukarela mendapat jasa ' + String(RK.jasaSukarela).replace('.', ',') +
        '% dari jumlah simpanan sukarela terendah pada bulan tersebut (disesuaikan dengan jasa bank);',
    'Simpanan sukarela tidak diperhitungkan dalam pembagian SHU;',
    'Pembagian SHU: SHU simpanan ' + RK.shuSimpanan + '% (simpanan pokok dan wajib) dan SHU pinjaman ' + RK.shuJasa + '%.'
  ].forEach(function(s, i) { x += NO_((i + 1) + '.', s, { ind: 425 }); });

  x += SUBJ_('D.  BIDANG PERMODALAN');
  x += T_([4500, 300, 2200], [
    ['Simpanan Pokok bagi anggota baru', ':', { t: rpRAT_(RK.nominal.nominal_pokok) + ',-', al: 'right' }],
    ['Simpanan Wajib per bulan', ':', { t: rpRAT_(RK.nominal.nominal_wajib) + ',-', al: 'right' }],
    ['Simpanan Sukarela minimal', ':', { t: rpRAT_(RK.nominal.minimal_sukarela) + ',-', al: 'right' }]
  ], { ind: 425, kn: true });
  x += T_([LEBAR_ISI - 4700, 1800, 300, 2600], [
    ['', 'Ditetapkan di', ':', kota], ['', 'Tanggal', ':', tglPanjang_(I.tanggal_rat) || bulanRAT]
  ], { kn: true });
  x += TTD_([{ jabatan: 'Pimpinan Rapat', nama: '' }], { sb: 0 });

  return bungkusDocx_(x, P);
}

/** 200,22 % → "2,00" (rupiah per Rp 1) */
function kaliRp_(persen) { return String((persen / 100).toFixed(2)).replace('.', ','); }
function kapital_(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }

/** Tabel aktiva & pasiva (neraca ringkas) untuk laporan pengawas. */
function tabelNeraca_(N) {
  const lb = [500, 5300, 2800];
  const baris = [[{ t: 'AKTIVA', b: true, span: 3, arsir: true }]];
  let no = 0;
  N.aktivaLancar.concat(N.aktivaTetap).forEach(function(r) {
    baris.push([{ t: (++no) + '.', al: 'right' }, r.kode === '103' ? 'Usaha Simpan Pinjam (piutang pinjaman)' : r.label,
      { t: rpRAT_(r.nilai), al: 'right' }]);
  });
  baris.push(['', { t: 'JUMLAH AKTIVA', b: true }, { t: rpRAT_(N.totalAktiva), b: true, al: 'right', atas: true }]);
  baris.push([{ t: 'PASIVA', b: true, span: 3, arsir: true }]);
  no = 0;
  const pasiva = N.modalSendiri.concat(N.hutangLancar, N.danaDana);
  pasiva.forEach(function(r) {
    baris.push([{ t: (++no) + '.', al: 'right' }, r.label, { t: rpRAT_(r.nilai), al: 'right' }]);
  });
  baris.push([{ t: (++no) + '.', al: 'right' }, 'SHU tahun ' + N.tahun, { t: rpRAT_(N.shuBerjalan), al: 'right' }]);
  baris.push(['', { t: 'JUMLAH PASIVA', b: true }, { t: rpRAT_(N.totalPasiva), b: true, al: 'right', atas: true }]);
  return T_(lb, baris, { ind: 850 });
}

/** Kop surat dengan logo (bila ada) + garis ganda. */
function kopSurat_(P) {
  const teks = P_(R_(P.jenis_koperasi || '', { b: true, sz: 11 }), { al: 'center', sa: 0 }) +
    P_(R_('“' + (P.nama_koperasi || '') + '”', { b: true, sz: 16 }), { al: 'center', sa: 0 }) +
    (P.no_badan_hukum ? P_(R_('Badan Hukum Nomor ' + P.no_badan_hukum, { sz: 10 }), { al: 'center', sa: 0 }) : '') +
    P_(R_((P.alamat || '') + (P.telepon ? ' · Telp. ' + P.telepon : '') + (P.email ? ' · ' + P.email : ''), { sz: 10 }), { al: 'center', sa: 0 });
  const logo = logoDocx_(P.logo);
  const x = logo
    ? T_([1400, LEBAR_ISI - 2800, 1400], [[{ isi: P_(logo.xml, { al: 'center', sa: 0 }), vmid: true }, { isi: teks, vmid: true }, { t: '' }]],
      { tanpaSpasi: true })
    : teks;
  return x + P_('', { garis: 'ganda', sa: 10 });
}

/** Gambar logo (data URL PNG/JPEG) → drawing inline 2,3 cm. */
function logoDocx_(dataUrl) {
  const m = /^data:image\/(png|jpeg);base64,(.+)$/.exec(String(dataUrl || ''));
  if (!m) return null;
  // tinggi 2,3 cm, lebar mengikuti perbandingan gambar (PNG: ukuran ada di header IHDR)
  const ukuran = ukuranGambar_(m[1], m[2]);
  const cy = Math.round(2.3 * 360000);
  const cx = Math.round(cy * ukuran.lebar / ukuran.tinggi);
  return {
    jenis: m[1], base64: m[2],
    xml: '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + cx + '" cy="' + cy + '"/>' +
      '<wp:docPr id="1" name="Logo"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:nvPicPr><pic:cNvPr id="1" name="logo"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdLogo"/>' +
      '<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
  };
}

/** Lebar × tinggi gambar dari isi base64 (PNG & JPEG); gagal → persegi. */
function ukuranGambar_(jenis, b64) {
  try {
    const b = Utilities.base64Decode(b64).map(function(x) { return x & 255; });
    if (jenis === 'png') {
      return { lebar: (b[16] << 24 | b[17] << 16 | b[18] << 8 | b[19]) >>> 0, tinggi: (b[20] << 24 | b[21] << 16 | b[22] << 8 | b[23]) >>> 0 };
    }
    for (let i = 2; i < b.length - 9;) {            // JPEG: cari penanda SOFn
      if (b[i] !== 0xFF) { i++; continue; }
      const m = b[i + 1], len = b[i + 2] << 8 | b[i + 3];
      if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) {
        return { tinggi: b[i + 5] << 8 | b[i + 6], lebar: b[i + 7] << 8 | b[i + 8] };
      }
      i += 2 + len;
    }
  } catch (e) { /* abaikan */ }
  return { lebar: 1, tinggi: 1 };
}

/** Rangkai bagian-bagian .docx lalu zip. */
function bungkusDocx_(isiBody, P) {
  const logo = logoDocx_(P.logo);
  const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
  const sect = '<w:sectPr><w:footerReference w:type="default" r:id="rIdKaki"/>' +
    '<w:pgSz w:w="' + F4_LEBAR + '" w:h="' + F4_TINGGI + '"/>' +
    '<w:pgMar w:top="' + TEPI.atas + '" w:right="' + TEPI.kanan + '" w:bottom="' + TEPI.bawah + '" w:left="' + TEPI.kiri +
    '" w:header="567" w:footer="567" w:gutter="0"/><w:titlePg/></w:sectPr>';
  const dokumen = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + ns + '><w:body>' +
    isiBody + sect + '</w:body></w:document>';
  const gaya = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr>' +
    '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/>' +
    '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="id-ID"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>' +
    '<w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>' +
    '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>' +
    '<w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>';
  const kaki = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ' + ns + '>' +
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve">Laporan RAT ' +
    xmlEsc_('Koperasi ' + (P.nama_koperasi || '')) + ' — halaman </w:t></w:r><w:r><w:rPr><w:sz w:val="18"/></w:rPr>' +
    '<w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
    '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>2</w:t></w:r>' +
    '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>';
  const jenisTipe = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '</Types>';
  const relAkar = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    '</Relationships>';
  const relDok = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rIdGaya" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '<Relationship Id="rIdKaki" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' +
    (logo ? '<Relationship Id="rIdLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.' +
      (logo.jenis === 'png' ? 'png' : 'jpeg') + '"/>' : '') + '</Relationships>';
  const inti = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Laporan RAT</dc:title><dc:creator>' +
    xmlEsc_('Koperasi ' + (P.nama_koperasi || '')) + '</dc:creator></cp:coreProperties>';

  const b = function(isi, nama) { return Utilities.newBlob(isi, 'application/xml', nama); };
  const berkas = [
    b(jenisTipe, '[Content_Types].xml'), b(relAkar, '_rels/.rels'),
    b(dokumen, 'word/document.xml'), b(gaya, 'word/styles.xml'), b(kaki, 'word/footer1.xml'),
    b(relDok, 'word/_rels/document.xml.rels'), b(inti, 'docProps/core.xml')
  ];
  if (logo) {
    berkas.push(Utilities.newBlob(Utilities.base64Decode(logo.base64), 'image/' + logo.jenis,
      'word/media/logo.' + (logo.jenis === 'png' ? 'png' : 'jpeg')));
  }
  const zip = Utilities.zip(berkas, 'laporan-rat.docx');
  return Utilities.newBlob(zip.getBytes(), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'laporan-rat.docx');
}
