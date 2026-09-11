/**
 * PengajuanService.gs
 * ------------------------------------------------------------
 * Logic server untuk halaman Form Pengajuan.
 * ------------------------------------------------------------
 */

/**
 * Daftar provinsi unik untuk dropdown Provinsi (section SPTJB & Trip),
 * diambil dari kolom "provinsi" di blok REF_TARIF (UH_PEGAWAI, fallback
 * ke HOTEL_PEJABAT kalau blok itu tidak ketemu).
 */
function getProvinsiList() {
  var blok = getBlokRefTarif_();
  var b = blok.UH_PEGAWAI || blok.HOTEL_PEJABAT;
  if (!b) return [];
  var colProv = findColIndex_(b.headers, ['provinsi']);
  if (colProv === -1) return [];
  var set = {};
  b.data.forEach(function (row) {
    var v = String(row[colProv] || '').trim();
    if (v) set[v] = true;
  });
  return Object.keys(set).sort();
}

/**
 * Daftar kab/kota per provinsi -- lihat getDaerahByProvinsi() di
 * WilayahService.gs (sumbernya API resmi wilayah Indonesia, bukan
 * REF_TARIF lagi, supaya cakupannya lengkap & bisa difilter per
 * provinsi yang dipilih).
 */


function getPegawaiList_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REF_PEGAWAI');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let r = 1; r < data.length; r++) {
    if (!data[r][0]) continue;
    out.push({ nama: data[r][0], nip: data[r][1], jabatan: data[r][2] });
  }
  return out;
}

// Hierarki Kegiatan > Output > SubOutput > Akun untuk dropdown bertingkat
// "Klasifikasi Anggaran" di Form Pengajuan.
function getRefAkunTree_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REF_AKUN');
  if (!sh) return { kegiatan: [], output: [], suboutput: [], akun: [] };
  const data = sh.getDataRange().getValues();
  const tree = { kegiatan: [], output: [], suboutput: [], akun: [] };
  const seen = {};

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const level = row[0], kode = row[1], nama = row[2];
    const kegiatanKode = row[3], outputKode = row[4], suboutputKode = row[5];
    if (!kode) continue;
    const dedupKey = level + '|' + kode + '|' + kegiatanKode + '|' + outputKode + '|' + suboutputKode;
    if (seen[dedupKey]) continue;
    seen[dedupKey] = true;

    if (level === 'Kegiatan') {
      tree.kegiatan.push({ kode: kode, nama: nama });
    } else if (level === 'Output') {
      tree.output.push({ kode: kode, nama: nama, kegiatanKode: kegiatanKode });
    } else if (level === 'SubOutput') {
      tree.suboutput.push({ kode: kode, nama: nama, kegiatanKode: kegiatanKode, outputKode: outputKode });
    } else if (level === 'Akun') {
      tree.akun.push({ kode: kode, nama: nama, kegiatanKode: kegiatanKode, outputKode: outputKode, suboutputKode: suboutputKode });
    }
  }
  return tree;
}


// Daftar SPTJB ringkas, untuk dropdown "pilih kegiatan yang sudah ada / buat baru"
function getSptjbListRingkas_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_SPTJB');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let r = data.length - 1; r >= 1; r--) {
    if (!data[r][0]) continue;
    out.push({ idSptjb: data[r][0], noSptjb: data[r][1], kegiatan: data[r][3], noAkun: data[r][9] });
  }
  return out;
}

/**
 * Submit form Pengajuan -- dipakai untuk BUAT BARU maupun EDIT data
 * yang sudah ada (kalau formData.idPengajuanEdit diisi). Satu kali
 * submit BISA berisi BEBERAPA Surat Tugas sekaligus (formData.tripList),
 * masing-masing dengan daftar orang & nominatif sendiri -- tapi semua
 * ikut 1 SPTJB & 1 No Akun yang sama.
 *
 * Status per Surat Tugas otomatis: 'Draft' kalau No SPM masih kosong,
 * 'Final' kalau No SPM sudah diisi.
 *
 * formData = {
 *   sptjbMode: 'baru' | 'existing',
 *   idSptjbExisting: '...',
 *   idPengajuanEdit: '...' // isi kalau EDIT (cuma 1 Surat Tugas saat edit), kosong kalau buat baru
 *   sptjbBaru: { notaDinas, kegiatan, jenis, provinsi, tglMulai, tglSelesai, noAkun, namaAkun, jumlahPengajuan, klasifikasiAnggaran },
 *   tripList: [
 *     { trip: { tujuan, provinsi, daerah, tglBerangkat, tglKembali, hotelVendor, noSpm, tahap, kegiatanTrip, noSt },
 *       orangList: [ { nama, nip, tipeId, noSkPeserta, namaKelompok,
 *                      transportPP, transportKedudukan, transportTujuan, transportLainnya,
 *                      uangHarian, penginapan, uangRepresentatif, totalNominatif }, ... ] },
 *     ... // bisa lebih dari 1 Surat Tugas
 *   ],
 *   token: '...'
 * }
 */
function submitPengajuan(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  const pengirim = getUserByToken_(formData.token);
  const emailPengirim = pengirim ? pengirim.email : '(tidak diketahui)';

  try {
    const shSptjb = ss.getSheetByName('DB_SPTJB');
    const shPengajuan = ss.getSheetByName('DB_PENGAJUAN');
    const isEdit = !!formData.idPengajuanEdit;

    let idSptjb;
    let noSptjbAuto = '';

    if (isEdit) {
      // Mode EDIT: cuma 1 Surat Tugas, pakai idTrip/idPengajuan/idSptjb
      // yang SUDAH ADA (baris lama dihapus, ditulis ulang).
      const existing = getPengajuanUntukEdit(formData.idPengajuanEdit);
      if (!existing) return { ok: false, error: 'Data pengajuan yang mau diedit tidak ditemukan.' };
      idSptjb = existing.idSptjb;
      hapusBarisPengajuan_(existing.idPengajuan);

      const noAkunTrip = lookupNoAkunSptjb_(idSptjb);
      const hasil = tulisSatuTrip_(shPengajuan, formData.tripList[0], idSptjb, noAkunTrip, emailPengirim,
        existing.idTrip, existing.idPengajuan);

      return {
        ok: true, isEdit: true, noSptjb: '', status: hasil.status,
        jumlahTrip: 1, jumlahOrang: hasil.jumlahOrang
      };
    }

    // ---- Mode BUAT BARU (bisa banyak Surat Tugas sekaligus) ----
    if (formData.sptjbMode === 'existing') {
      idSptjb = formData.idSptjbExisting;
    } else {
      const s = formData.sptjbBaru;
      idSptjb = 'SPTJB-2026-' + nextSequence_('SPTJB');
      noSptjbAuto = generateNoSptjb_();
      const rowIdx = shSptjb.getLastRow() + 1;
      shSptjb.getRange(rowIdx, 1, 1, 19).setValues([[
        idSptjb, noSptjbAuto, '', s.kegiatan || '', '',
        '', '', s.tglMulai || '', s.tglSelesai || '',
        s.noAkun || '', '', s.namaAkun || '', '', Number(s.jumlahPengajuan) || 0,
        'Aktif', emailPengirim, new Date(), '',
        s.klasifikasiAnggaran || ''
      ]]);
    }

    // No Akun trip SELALU ikut No Akun SPTJB-nya -- 1 SPTJB = 1 akun,
    // tidak bisa beda per Surat Tugas, walau Surat Tugasnya lebih dari 1.
    const noAkunTrip = (formData.sptjbMode === 'existing' ? lookupNoAkunSptjb_(idSptjb) : formData.sptjbBaru.noAkun) || '';

    let jumlahOrangTotal = 0;
    let statusTerakhir = '';
    formData.tripList.forEach(function (item) {
      const idTrip = 'TRIP-' + nextSequence_('TRIP');
      const idPengajuan = 'PENG-2026-' + nextSequence_('PENG');
      const hasil = tulisSatuTrip_(shPengajuan, item, idSptjb, noAkunTrip, emailPengirim, idTrip, idPengajuan);
      jumlahOrangTotal += hasil.jumlahOrang;
      statusTerakhir = hasil.status;
    });

    return {
      ok: true, isEdit: false, noSptjb: noSptjbAuto, status: statusTerakhir,
      jumlahTrip: formData.tripList.length, jumlahOrang: jumlahOrangTotal
    };
  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Tulis baris DB_PENGAJUAN untuk SATU Surat Tugas (1 idTrip/idPengajuan)
 * beserta semua orangnya. Dipakai berulang oleh submitPengajuan untuk
 * kasus banyak Surat Tugas sekaligus, dan sekali saja untuk kasus edit.
 */
function tulisSatuTrip_(shPengajuan, item, idSptjb, noAkunTrip, emailPengirim, idTrip, idPengajuan) {
  const t = item.trip;
  const status = t.noSpm ? 'Final' : 'Draft';

  const rows = item.orangList.map(function (o) {
    const idBaris = 'BR-' + nextSequence_('BARIS');
    return [
      idBaris, idTrip, idPengajuan, idSptjb,
      noAkunTrip,
      t.noSpm || '', t.tahap || '', t.kegiatanTrip || '',
      o.nama || '', o.nip || '', o.tipeId || '', t.noSk || '', o.namaKelompok || '',
      '', t.tujuan || '', t.provinsi || '', t.daerah || '',
      t.tglBerangkat || '', t.tglKembali || '', hitungLamaHari_(t.tglBerangkat, t.tglKembali),
      t.hotelVendor || '', item.orangList.length,
      status, '', '', '',
      emailPengirim, new Date(), '',
      t.noSt || '',
      Number(o.transportPP) || 0,
      Number(o.transportKedudukan) || 0,
      Number(o.transportTujuan) || 0,
      Number(o.uangHarian) || 0,
      Number(o.penginapan) || 0,
      Number(o.uangRepresentatif) || 0,
      Number(o.totalNominatif) || 0,
      Number(o.transportLainnya) || 0 // kolom baru, ditambah di UJUNG (lihat AddFieldsV4.gs)
    ];
  });

  const startRow = shPengajuan.getLastRow() + 1;
  shPengajuan.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

  return { status: status, jumlahOrang: rows.length };
}

/**
 * Hapus semua baris DB_PENGAJUAN milik satu ID_Pengajuan (dipakai
 * saat mode EDIT, sebelum menulis ulang baris terbaru).
 *
 * CATATAN KETERBATASAN: cara ini hapus+tulis ulang ID_Baris yang BARU
 * untuk semua orang (bukan update di tempat) -- jadi kalau pengajuan
 * yang diedit SUDAH punya data di DB_REALISASI yang tertaut ke
 * ID_Baris lama, tautannya akan putus (Realisasi jadi "nyantol" ke
 * ID_Baris yang sudah tidak ada). Aman untuk edit pengajuan yang
 * belum direalisasikan; untuk yang sudah, sebaiknya cek dulu manual
 * atau kabari saya kalau perlu dibuatkan proteksinya.
 */
function hapusBarisPengajuan_(idPengajuan) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_PENGAJUAN');
  const data = sh.getDataRange().getValues();
  for (let r = data.length - 1; r >= 1; r--) {
    if (data[r][2] === idPengajuan) sh.deleteRow(r + 1);
  }
}

/**
 * Dipanggil dari client untuk memuat ulang satu Pengajuan (semua
 * orang dalam 1 trip) ke form, supaya bisa diedit. TIDAK pakai
 * akhiran "_" karena dipanggil lewat google.script.run.
 */
function getPengajuanUntukEdit(idPengajuan) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_PENGAJUAN');
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  const rows = [];
  for (let r = 1; r < data.length; r++) {
    if (data[r][2] === idPengajuan) rows.push(data[r]);
  }
  if (rows.length === 0) return null;

  const first = rows[0];
  return {
    idPengajuan: idPengajuan,
    idTrip: first[1],
    idSptjb: first[3],
    trip: {
      tujuan: first[14], provinsi: first[15], daerah: first[16],
      tglBerangkat: formatTanggalInput_(first[17]), tglKembali: formatTanggalInput_(first[18]),
      hotelVendor: first[20], noSpm: first[5], noSt: first[29], noSk: first[11],
      tahap: first[6], kegiatanTrip: first[7]
    },
    orangList: rows.map(function (row) {
      return {
        nama: row[8], nip: row[9], tipeId: row[10], namaKelompok: row[12],
        transportPP: row[30], transportKedudukan: row[31], transportTujuan: row[32],
        uangHarian: row[33], penginapan: row[34], uangRepresentatif: row[35], totalNominatif: row[36],
        transportLainnya: row[37] || 0
      };
    })
  };
}

/**
 * Daftar ringkas semua Pengajuan (dikelompokkan per ID_Pengajuan)
 * untuk ditampilkan di form supaya bisa dipilih untuk di-Edit.
 * Dipanggil dari server (Code.gs doGet), bukan dari client.
 */
function getDaftarPengajuanUntukEdit_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_PENGAJUAN');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const map = {};
  const order = [];
  for (let r = 1; r < data.length; r++) {
    const idPengajuan = data[r][2];
    if (!idPengajuan) continue;
    if (!map[idPengajuan]) {
      map[idPengajuan] = {
        idPengajuan: idPengajuan,
        tujuan: data[r][14] || '-',
        tglBerangkat: formatTanggal_(data[r][17]),
        tglKembali: formatTanggal_(data[r][18]),
        jumlahOrang: 0,
        status: data[r][22] || '-',
        noSpm: data[r][5] || ''
      };
      order.push(idPengajuan);
    }
    map[idPengajuan].jumlahOrang++;
  }
  return order.map(function (id) { return map[id]; }).reverse();
}

function formatTanggalInput_(v) {
  if (!v) return '';
  try {
    if (Object.prototype.toString.call(v) === '[object Date]') {
      return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    return String(v);
  } catch (e) {
    return '';
  }
}

function lookupNoAkunSptjb_(idSptjb) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_SPTJB');
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === idSptjb) return data[r][9];
  }
  return '';
}

function hitungLamaHari_(tglMulai, tglSelesai) {
  try {
    const d1 = new Date(tglMulai);
    const d2 = new Date(tglSelesai);
    const diff = Math.round((d2 - d1) / (1000 * 60 * 60 * 24)) + 1;
    return diff > 0 ? diff : '';
  } catch (err) {
    return '';
  }
}

// Penomoran berurutan sederhana, disimpan di Properties supaya tidak
// perlu scan seluruh sheet tiap kali submit.
function toRoman_(num) {
  const map = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'],
    [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'],
    [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']
  ];
  let result = '';
  map.forEach(function (pair) {
    while (num >= pair[0]) {
      result += pair[1];
      num -= pair[0];
    }
  });
  return result;
}

/**
 * Generate No SPTJB otomatis dengan format resmi:
 *   3/{nomor urut per tahun}/PK.03.03/{bulan romawi}/{tahun}
 * Nomor urut reset tiap tahun (pakai tanggal SAAT INI, bukan
 * tanggal mulai kegiatan -- sesuai kebiasaan penomoran resmi).
 */
function generateNoSptjb_() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const props = PropertiesService.getScriptProperties();
  const propKey = 'SPTJB_COUNTER_' + year;
  const next = Number(props.getProperty(propKey) || '0') + 1;
  props.setProperty(propKey, String(next));
  return '3/' + next + '/PK.03.03/' + toRoman_(month) + '/' + year;
}

/**
 * Dipanggil dari client (google.script.run) untuk MENAMPILKAN di form
 * nomor SPTJB yang AKAN terbentuk kalau pengajuan disimpan sekarang --
 * TANPA memakai/mengunci nomornya (beda dari generateNoSptjb_ yang
 * menambah counter permanen). Jadi aman dipanggil berkali-kali
 * (refresh halaman, ganti tab, dst) tanpa bikin nomor "loncat" gara-gara
 * dibuka tapi tidak pernah disimpan.
 *
 * CATATAN: kalau form dibuka menjelang pergantian bulan/tahun, nomor
 * preview ini bisa beda tipis dari nomor final (yang dihitung ulang
 * saat submit) -- itu wajar, bukan bug.
 */
function previewNoSptjb() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const props = PropertiesService.getScriptProperties();
  const propKey = 'SPTJB_COUNTER_' + year;
  const next = Number(props.getProperty(propKey) || '0') + 1;
  return '3/' + next + '/PK.03.03/' + toRoman_(month) + '/' + year;
}

function nextSequence_(key) {
  const props = PropertiesService.getScriptProperties();
  const propKey = 'SEQ_' + key;
  let n = Number(props.getProperty(propKey) || '0') + 1;
  props.setProperty(propKey, String(n));
  return String(n).padStart(5, '0');
}

/**
 * PENTING: jalankan fungsi ini SEKALI SAJA sebelum pertama kali pakai
 * form Pengajuan/Realisasi -- supaya nomor urut baru (SPTJB-2026-xxxx,
 * PENG-2026-xxxx, dst) tidak bentrok dengan ID hasil migrasi historis.
 * Fungsi ini scan ID tertinggi yang sudah ada, lalu set Properties
 * mulai dari situ.
 */
function initSequencesV1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  function maxSuffix(sheetName, colIndex, prefix) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return 0;
    const data = sh.getDataRange().getValues();
    let max = 0;
    for (let r = 1; r < data.length; r++) {
      const v = String(data[r][colIndex] || '');
      if (v.indexOf(prefix) === 0) {
        const n = parseInt(v.slice(prefix.length), 10);
        if (!isNaN(n) && n > max) max = n;
      }
    }
    return max;
  }

  props.setProperty('SEQ_SPTJB', String(maxSuffix('DB_SPTJB', 0, 'SPTJB-2026-')));
  props.setProperty('SEQ_PENG', String(maxSuffix('DB_PENGAJUAN', 2, 'PENG-2026-')));
  // TRIP dan BARIS format lama beda (P/xxx, P/xxx/n) -- mulai dari 0 aman
  // karena prefix baru (TRIP-, BR-) tidak akan pernah bentrok dgn format lama.
  props.setProperty('SEQ_TRIP', '0');
  props.setProperty('SEQ_BARIS', '0');
  props.setProperty('SEQ_REAL', String(maxSuffix('DB_REALISASI', 0, 'REAL-')));

  // Seed nomor urut SPTJB per tahun dari data lama, format "3/{n}/PK.03.03/{romawi}/{tahun}"
  seedNoSptjbCounters_(ss, props);

  Logger.log('Sequence diinisialisasi: ' + JSON.stringify(props.getProperties()));
}

function seedNoSptjbCounters_(ss, props) {
  const sh = ss.getSheetByName('DB_SPTJB');
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  const maxByYear = {};
  const pattern = /^3\/(\d+)\/PK\.03\.03\/[IVXLCDM]+\/(\d{4})$/;
  for (let r = 1; r < data.length; r++) {
    const v = String(data[r][1] || '');
    const m = v.match(pattern);
    if (m) {
      const n = parseInt(m[1], 10);
      const year = m[2];
      if (!maxByYear[year] || n > maxByYear[year]) maxByYear[year] = n;
    }
  }
  Object.keys(maxByYear).forEach(function (year) {
    props.setProperty('SPTJB_COUNTER_' + year, String(maxByYear[year]));
  });
}

/**
 * ============================================================================
 * MODUL SARAN NOMINATIF OTOMATIS — form Pengajuan SIPERDIN
 * ----------------------------------------------------------------------------
 * Tujuan: begitu user isi nama pegawai + tujuan + tanggal di form Pengajuan,
 * sistem menarik batas MAKSIMAL tarif dari sheet REF_TARIF (sesuai PMK 32/2025)
 * untuk tiap komponen biaya, lalu mem-prefill field angka di form.
 * Field tsb tetap <input type="number"> biasa (BUKAN readonly) — user bisa
 * turunkan manual karena SBM adalah batas atas, bukan angka pasti.
 *
 * CATATAN PENTING:
 * - Nama judul blok & nama kolom di bawah ini adalah ASUMSI berdasarkan hasil
 *   kerja sebelumnya (REF_TARIF berisi 7 blok berjejer samping: REF_UH_PEGAWAI,
 *   REF_HOTEL_PEJABAT, REF_FULLBOARD_PESERTA, REF_SEWA_MOBIL,
 *   REF_TRANSPORT_PESERTA, REF_TAKSI, REF_TIKET_PESAWAT).
 *   Cek TITLE_KEYWORDS & kandidat nama kolom di findColIndex_ terhadap sheet
 *   REF_TARIF asli — sesuaikan string-nya kalau beda.
 * - Uang Representatif BELUM ada blok referensinya → sengaja dikembalikan 0
 *   dengan TODO, supaya user isi manual sampai blok itu ditambahkan.
 * ============================================================================
 */

var REF_TARIF_SHEET_NAME = 'REF_TARIF';

var TITLE_KEYWORDS = {
  UH_PEGAWAI: 'uh pegawai',
  HOTEL_PEJABAT: 'hotel pejabat',
  FULLBOARD_PESERTA: 'fullboard peserta',
  SEWA_MOBIL: 'sewa mobil',
  TRANSPORT_PESERTA: 'transport peserta',
  TAKSI: 'taksi',
  TIKET_PESAWAT: 'tiket pesawat'
};

/**
 * Fungsi utama — dipanggil dari Form.html lewat google.script.run.
 * @param {Object} p {namaPegawai, provinsiTujuan, kabKotaTujuan, lamaHari}
 * @return {Object} breakdown nominatif per komponen + total
 */
function hitungNominatifSaran(p) {
  var pegawai = getPegawaiByNama_(p.namaPegawai);
  var blok = getBlokRefTarif_();

  var transportPP = cariTiketPesawat_(blok.TIKET_PESAWAT, 'Jakarta', p.kabKotaTujuan, pegawai.golongan);
  // TAKSI di REF_TARIF diindeks per PROVINSI (bukan per kota) -- jadi
  // kedudukan (asal) pakai provinsi DKI Jakarta, tujuan pakai provinsi
  // tujuan trip, BUKAN nama kab/kota-nya.
  var transportKedudukan = cariTaksi_(blok.TAKSI, 'D.K.I. Jakarta');
  var transportTujuan = cariTaksi_(blok.TAKSI, p.provinsiTujuan);
  if (!transportTujuan) {
    transportTujuan = cariTransportPeserta_(blok.TRANSPORT_PESERTA, p.kabKotaTujuan);
  }
  var uhHarian = cariUhPegawai_(blok.UH_PEGAWAI, p.provinsiTujuan, 'luar_kota');
  var uangHarian = uhHarian * p.lamaHari;
  var tarifHotel = cariHotelPejabat_(blok.HOTEL_PEJABAT, p.provinsiTujuan, pegawai.golonganHotel);
  var penginapan = tarifHotel * Math.max(p.lamaHari - 1, 0);
  var uangRepresentatif = 0; // TODO: belum ada blok referensi, isi manual dulu

  var total = transportPP + transportKedudukan + transportTujuan
            + uangHarian + penginapan + uangRepresentatif;

  return {
    transportPP: transportPP,
    transportKedudukan: transportKedudukan,
    transportTujuan: transportTujuan,
    uangHarian: uangHarian,
    penginapan: penginapan,
    uangRepresentatif: uangRepresentatif,
    total: total,
    catatan: uangRepresentatif === 0
      ? 'Uang Representatif belum ada referensi tarif — isi manual jika berlaku.'
      : ''
  };
}

function getBlokRefTarif_(skipCache) {
  var cache = CacheService.getScriptCache();
  if (!skipCache) {
    var cached = cache.get('BLOK_REF_TARIF');
    if (cached) return JSON.parse(cached);
  }

  var sheet = SpreadsheetApp.getActive().getSheetByName(REF_TARIF_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + REF_TARIF_SHEET_NAME + '" tidak ditemukan di spreadsheet ini.');
  var lastCol = sheet.getLastColumn();
  var titleRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

  var blocks = {};
  var currentKey = null;
  var currentStart = null;

  for (var c = 0; c < lastCol; c++) {
    // pakai normalisasi_ (bukan toLowerCase/trim biasa) supaya spasi ganda
    // atau spasi tak biasa di judul sheet tidak bikin deteksi blok gagal
    var titleCell = normalisasi_(titleRow[c]);
    if (titleCell) {
      if (currentKey) {
        blocks[currentKey] = closeBlock_(sheet, currentStart, c, headerRow);
      }
      currentKey = matchTitleKeyword_(titleCell);
      currentStart = c;
    }
  }
  if (currentKey) {
    blocks[currentKey] = closeBlock_(sheet, currentStart, lastCol, headerRow);
  }

  cache.put('BLOK_REF_TARIF', JSON.stringify(blocks), 300); // cache 5 menit
  return blocks;
}

function closeBlock_(sheet, startCol, endCol, headerRow) {
  var width = endCol - startCol;
  var headers = headerRow.slice(startCol, endCol).map(function (h) {
    return String(h || '').toLowerCase().trim();
  });
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(3, startCol + 1, lastRow - 2, width).getValues();
  return { startCol: startCol, headers: headers, data: data };
}

function matchTitleKeyword_(titleLower) {
  for (var key in TITLE_KEYWORDS) {
    if (titleLower.indexOf(TITLE_KEYWORDS[key]) !== -1) return key;
  }
  return null;
}

function findColIndex_(headers, candidates) {
  for (var i = 0; i < headers.length; i++) {
    for (var j = 0; j < candidates.length; j++) {
      if (headers[i].indexOf(candidates[j]) !== -1) return i;
    }
  }
  return -1;
}

function normalisasi_(s) {
  return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

// --- Pencari per blok ---------------------------------------------------

function cariUhPegawai_(blokData, provinsi, jenis) {
  if (!blokData) return 0;
  var colProv = findColIndex_(blokData.headers, ['provinsi']);
  var colLuar = findColIndex_(blokData.headers, ['luar kota', 'luar_kota']);
  var colDalam = findColIndex_(blokData.headers, ['dalam kota', 'dalam_kota']);
  var colTarget = jenis === 'dalam_kota' ? colDalam : colLuar;
  if (colProv === -1 || colTarget === -1) return 0;

  var target = normalisasi_(provinsi);
  for (var i = 0; i < blokData.data.length; i++) {
    if (normalisasi_(blokData.data[i][colProv]) === target) {
      return Number(blokData.data[i][colTarget]) || 0;
    }
  }
  return 0;
}

function cariHotelPejabat_(blokData, provinsi, golonganHotel) {
  if (!blokData) return 0;
  var colProv = findColIndex_(blokData.headers, ['provinsi']);
  var colGol = findColIndex_(blokData.headers, [normalisasi_(golonganHotel)]);
  if (colProv === -1 || colGol === -1) return 0;

  var target = normalisasi_(provinsi);
  for (var i = 0; i < blokData.data.length; i++) {
    if (normalisasi_(blokData.data[i][colProv]) === target) {
      return Number(blokData.data[i][colGol]) || 0;
    }
  }
  return 0;
}

function cariTaksi_(blokData, provinsi) {
  if (!blokData) return 0;
  var colProv = findColIndex_(blokData.headers, ['provinsi']);
  var colTarif = findColIndex_(blokData.headers, ['tarif', 'taksi', 'nilai']);
  if (colProv === -1 || colTarif === -1) return 0;

  var target = normalisasi_(provinsi);
  for (var i = 0; i < blokData.data.length; i++) {
    if (normalisasi_(blokData.data[i][colProv]) === target) {
      return Number(blokData.data[i][colTarif]) || 0;
    }
  }
  return 0;
}

function cariTransportPeserta_(blokData, kabKota) {
  if (!blokData) return 0;
  var colTujuan = findColIndex_(blokData.headers, ['kab', 'kota', 'tujuan']);
  var colTarif = findColIndex_(blokData.headers, ['tarif', 'nilai']);
  if (colTujuan === -1 || colTarif === -1) return 0;

  var target = normalisasi_(kabKota);
  for (var i = 0; i < blokData.data.length; i++) {
    var cell = normalisasi_(blokData.data[i][colTujuan]);
    if (cell.indexOf(target) !== -1 || target.indexOf(cell) !== -1) {
      return Number(blokData.data[i][colTarif]) || 0;
    }
  }
  return 0;
}

function cariTiketPesawat_(blokData, kotaAsal, kotaTujuan, golongan) {
  if (!blokData) return 0;
  var colAsal = findColIndex_(blokData.headers, ['asal']);
  var colTujuan = findColIndex_(blokData.headers, ['tujuan']);
  var kelas = golonganTermasukBisnis_(golongan) ? 'bisnis' : 'ekonomi';
  var colTarif = findColIndex_(blokData.headers, [kelas]);
  if (colAsal === -1 || colTujuan === -1 || colTarif === -1) return 0;

  var tAsal = normalisasi_(kotaAsal);
  var tTujuan = normalisasi_(kotaTujuan);
  for (var i = 0; i < blokData.data.length; i++) {
    var asal = normalisasi_(blokData.data[i][colAsal]);
    var tujuan = normalisasi_(blokData.data[i][colTujuan]);
    if (asal.indexOf(tAsal) !== -1 &&
        (tujuan.indexOf(tTujuan) !== -1 || tTujuan.indexOf(tujuan) !== -1)) {
      return Number(blokData.data[i][colTarif]) || 0;
    }
  }
  return 0;
}

/**
 * Aturan kelas tiket bisnis biasanya untuk pejabat eselon I/II (golongan IV/d ke atas).
 * TODO: sesuaikan ambang golongan ini dengan aturan resmi di PMK 32/2025 jika beda.
 */
function golonganTermasukBisnis_(golongan) {
  if (!golongan) return false;
  var g = String(golongan).toUpperCase().replace(/\s+/g, '');
  return g.indexOf('IV/D') !== -1 || g.indexOf('IV/E') !== -1 || g.indexOf('IVD') !== -1 || g.indexOf('IVE') !== -1;
}

/**
 * Ambil data pegawai (golongan, golongan_hotel) dari REF_PEGAWAI berdasar nama.
 */
function getPegawaiByNama_(nama) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('REF_PEGAWAI');
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function (h) { return normalisasi_(h); });
  var colNama = findColIndex_(headers, ['nama']);
  var colGol = findColIndex_(headers, ['golongan']);
  var colGolHotel = findColIndex_(headers, ['golongan_hotel', 'golongan hotel']);

  var target = normalisasi_(nama);
  for (var i = 1; i < data.length; i++) {
    if (normalisasi_(data[i][colNama]) === target) {
      return {
        nama: data[i][colNama],
        golongan: colGol !== -1 ? data[i][colGol] : '',
        golonganHotel: colGolHotel !== -1 ? data[i][colGolHotel] : ''
      };
    }
  }
  return { nama: nama, golongan: '', golonganHotel: '' };
}

/**
 * DIAGNOSTIK — jalankan fungsi ini manual dari editor Apps Script
 * (pilih debugRefTarif di dropdown atas, klik Run), lalu buka
 * Executions / View > Logs untuk lihat hasilnya.
 * Ini akan menunjukkan blok apa saja yang berhasil terdeteksi di
 * REF_TARIF, dan nama kolom (header) yang terbaca di tiap blok —
 * supaya kita bisa cocokkan dengan TITLE_KEYWORDS & findColIndex_
 * kalau ada yang tidak match.
 */
function debugRefTarif() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(REF_TARIF_SHEET_NAME);
  if (!sheet) {
    Logger.log('❌ Sheet "' + REF_TARIF_SHEET_NAME + '" TIDAK DITEMUKAN. Cek nama sheet persis.');
    return;
  }

  var lastCol = sheet.getLastColumn();
  var titleRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  Logger.log('=== Baris judul (row 1), mentah ===');
  Logger.log(JSON.stringify(titleRow));

  var blok = getBlokRefTarif_(true); // skip cache, baca ulang dari sheet
  var keys = Object.keys(blok);
  Logger.log('=== Blok yang TERDETEKSI: ' + (keys.length ? keys.join(', ') : '(tidak ada satupun)') + ' ===');

  Object.keys(TITLE_KEYWORDS).forEach(function (key) {
    if (blok[key]) {
      Logger.log('✅ ' + key + ' → headers: ' + JSON.stringify(blok[key].headers) + ' | jumlah baris data: ' + blok[key].data.length);
    } else {
      Logger.log('❌ ' + key + ' TIDAK terdeteksi (cari kata kunci "' + TITLE_KEYWORDS[key] + '" di baris judul)');
    }
  });

  Logger.log('=== Tes getProvinsiList() ===');
  Logger.log(JSON.stringify(getProvinsiList()));
  Logger.log('=== Tes getDaerahByProvinsi("Jawa Barat") ===');
  Logger.log(JSON.stringify(getDaerahByProvinsi('Jawa Barat')));
}

/**
 * DIAGNOSTIK Transport Tiket -- jalankan manual dari editor, isi
 * parameter di baris pemanggilan paling bawah sesuai kasus yang
 * bermasalah, lalu cek Logs.
 */
function debugTiketPesawat() {
  var kotaTujuan = 'Kota Medan'; // GANTI sesuai kab/kota yang dicoba di form
  var golongan = 'III/a'; // GANTI sesuai golongan pegawai yang dicoba

  var blok = getBlokRefTarif_(true);
  var b = blok.TIKET_PESAWAT;
  if (!b) { Logger.log('❌ Blok TIKET_PESAWAT tidak terdeteksi sama sekali.'); return; }

  Logger.log('=== Header blok TIKET_PESAWAT ===');
  Logger.log(JSON.stringify(b.headers));
  Logger.log('Jumlah baris data: ' + b.data.length);
  Logger.log('=== 10 baris pertama (mentah) ===');
  Logger.log(JSON.stringify(b.data.slice(0, 10)));

  var kelas = golonganTermasukBisnis_(golongan) ? 'bisnis' : 'ekonomi';
  Logger.log('Golongan "' + golongan + '" -> kelas dipakai: ' + kelas);

  var hasil = cariTiketPesawat_(b, 'Jakarta', kotaTujuan, golongan);
  Logger.log('=== Hasil cariTiketPesawat_("Jakarta", "' + kotaTujuan + '", "' + golongan + '") ===');
  Logger.log('Tarif ditemukan: ' + hasil);

  // Tampilkan semua baris yang "tujuan"-nya mengandung sebagian kata dari kotaTujuan,
  // untuk lihat format asli penulisan kota di REF_TARIF
  var kw = normalisasi_(kotaTujuan).split(' ').pop(); // kata terakhir, mis. "medan" dari "kota medan"
  Logger.log('=== Baris yang mengandung kata "' + kw + '" di kolom manapun ===');
  b.data.forEach(function (row) {
    var gabung = row.join(' | ').toLowerCase();
    if (gabung.indexOf(kw) !== -1) Logger.log(JSON.stringify(row));
  });
}