/**
 * Code.gs
 * ------------------------------------------------------------
 * Kerangka utama aplikasi web SIPERDIN.
 *
 * ALUR (LOGIN MANUAL, bukan akun Google):
 * 1. Orang buka URL web app -> doGet(e) jalan.
 * 2. Cek parameter ?token=... di URL. Token dicocokkan ke sesi
 *    yang tersimpan di CacheService (lihat SessionService.gs).
 *    - Tidak ada token / token tidak valid/kadaluarsa -> tampilkan
 *      halaman Login.html.
 *    - Valid -> ambil data user (email/nama/role), lanjut routing.
 * 3. Routing berdasarkan ?page=... di URL:
 *      (kosong)      -> Index.html   (dashboard ringkas)
 *      ?page=pengajuan  -> halaman form pengajuan
 *      ?page=realisasi  -> halaman form realisasi
 *      ?page=users      -> kelola user (Admin only)
 *      ?page=logout     -> hapus sesi, kembali ke Login
 * 4. Menu.html di-include di setiap halaman, isinya menyesuaikan
 *    Role, dan setiap link nav ikut menyelipkan &token=... supaya
 *    sesi tetap "nempel" saat pindah halaman.
 *
 * PENTING soal token di halaman lain: kalau ada halaman yang bikin
 * redirect sendiri lewat JavaScript (mis. setelah submit form sukses
 * lalu window.top.location = appUrl + '?page=...'), pastikan token
 * ikut disertakan di URL itu -- kalau tidak, sesi akan "hilang" dan
 * user dilempar balik ke Login. Contoh yang sudah diperbaiki:
 * FormPengajuan.html.
 *
 * CARA DEPLOY:
 * Deploy > Manage deployments > edit > Who has access: "Anyone"
 * (BUKAN "Anyone with Google account") -- karena login sekarang
 * ditangani sendiri (manual), bukan lewat akun Google lagi.
 * Setiap kali source code diubah, perlu "New version" lagi di
 * deployment yang sama supaya perubahan kelihatan di URL yang sama.
 * ------------------------------------------------------------
 */

const PAGE_TITLE = 'SIPERDIN 2026';

// Token sesi request SAAT INI -- di-set di awal doGet(), dipakai oleh
// Menu.html (lewat getCurrentToken_()) supaya semua link nav ikut
// menyelipkan token tanpa perlu diteruskan manual ke tiap template.
var CURRENT_TOKEN = '';

function doGet(e) {
  const token = (e && e.parameter && e.parameter.token) || '';
  const page = (e && e.parameter && e.parameter.page) || 'index';

  if (page === 'logout') {
    logout(token);
    return renderLoginPage_();
  }

  CURRENT_TOKEN = token;
  const user = getUserByToken_(token);

  if (!user) {
    return renderLoginPage_();
  }

  // Halaman yang butuh role Admin -- tambah nama page di sini kalau perlu
  const adminOnlyPages = ['users', 'referensi', 'config'];
  if (adminOnlyPages.indexOf(page) !== -1 && user.role !== 'Admin') {
    return renderPage_('AccessDenied', { email: user.email, reason: 'Halaman ini khusus Admin.' });
  }

  switch (page) {
    case 'pengajuan':
      return renderPage_('FormPengajuan', {
        user: user, activePage: 'pengajuan',
        pegawaiList: getPegawaiList_(), sptjbList: getSptjbListRingkas_(),
        refAkun: getRefAkunTree_(), kodeProgram: getKodeProgram_(),
        daftarPengajuan: getDaftarPengajuanUntukEdit_()
      });
    case 'realisasi':
      return renderPage_('FormRealisasi', { user: user, activePage: 'realisasi', pengajuanList: getPengajuanListUntukRealisasi_() });
    case 'users':
      return renderPage_('KelolaUser', { user: user, activePage: 'users', userList: getUserList_() });
    case 'referensi':
      return renderPage_('Referensi', { user: user, activePage: 'referensi', refCounts: getRefCounts_() });
    case 'config':
      return renderPage_('Konfigurasi', { user: user, activePage: 'config', config: getConfig_() });
    case 'index':
    default:
      return renderPage_('Dashboard', { user: user, activePage: 'index', data: getDashboardData_() });
  }
}

// Dipakai oleh Menu.html & halaman lain lewat scriptlet <?= getCurrentToken_() ?>
// supaya link/redirect ikut membawa token sesi.
function getCurrentToken_() {
  return CURRENT_TOKEN;
}

function renderLoginPage_() {
  const tmpl = HtmlService.createTemplateFromFile('Login');
  tmpl.appUrl = ScriptApp.getService().getUrl();
  return tmpl.evaluate()
    .setTitle(PAGE_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ------------------------------------------------------------
// Render halaman: template HTML + inject data (user, activePage, dll)
// ------------------------------------------------------------
function renderPage_(templateName, data) {
  const tmpl = HtmlService.createTemplateFromFile(templateName);
  const merged = Object.assign({ appUrl: ScriptApp.getService().getUrl() }, data || {});
  Object.keys(merged).forEach(function (key) {
    tmpl[key] = merged[key];
  });
  return tmpl.evaluate()
    .setTitle(PAGE_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Dipakai oleh template HTML untuk include file lain dengan variabel:
// <?!= include('Menu', {user: user, activePage: activePage}); ?>
function include(filename, data) {
  const tmpl = HtmlService.createTemplateFromFile(filename);
  Object.keys(data || {}).forEach(function (key) {
    tmpl[key] = data[key];
  });
  return tmpl.evaluate().getContent();
}

// ------------------------------------------------------------
// Kumpulkan data ringkasan untuk Dashboard
// ------------------------------------------------------------
function getDashboardData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sptjbSheet = ss.getSheetByName('DB_SPTJB');
  const pengajuanSheet = ss.getSheetByName('DB_PENGAJUAN');
  const realisasiSheet = ss.getSheetByName('DB_REALISASI');

  const sptjbData = sptjbSheet ? sptjbSheet.getDataRange().getValues() : [[]];
  const pengajuanData = pengajuanSheet ? pengajuanSheet.getDataRange().getValues() : [[]];
  const realisasiData = realisasiSheet ? realisasiSheet.getDataRange().getValues() : [[]];

  // Kolom DB_PENGAJUAN: A=ID_Baris, B=ID_Trip, C=ID_Pengajuan, D=ID_SPTJB ...
  // Kolom DB_SPTJB: A=ID_SPTJB, B=No_SPTJB, C=Nota_Dinas, D=Kegiatan, E=Jenis,
  //   F=Provinsi, G=Daerah, H=Tanggal_Mulai, I=Tanggal_Selesai, N=Jumlah_Pengajuan_Rp
  // Kolom DB_REALISASI: B=ID_Baris (link), AF=Total_Realisasi (indeks 31, 0-based)

  // Hitung jumlah orang per ID_SPTJB dari DB_PENGAJUAN
  const orangPerSptjb = {};
  for (let r = 1; r < pengajuanData.length; r++) {
    const idSptjb = pengajuanData[r][3];
    if (!idSptjb) continue;
    orangPerSptjb[idSptjb] = (orangPerSptjb[idSptjb] || 0) + 1;
  }

  // Total realisasi (Rp) per ID_Baris -> perlu di-map ke ID_SPTJB lewat Pengajuan
  const idBarisToSptjb = {};
  for (let r = 1; r < pengajuanData.length; r++) {
    const idBaris = pengajuanData[r][0];
    const idSptjb = pengajuanData[r][3];
    if (idBaris) idBarisToSptjb[idBaris] = idSptjb;
  }
  const realisasiPerSptjb = {};
  let totalRealisasiKeseluruhan = 0;
  for (let r = 1; r < realisasiData.length; r++) {
    const idBaris = realisasiData[r][1];
    const totalRealisasi = Number(realisasiData[r][31]) || 0; // kolom AF
    totalRealisasiKeseluruhan += totalRealisasi;
    const idSptjb = idBarisToSptjb[idBaris];
    if (idSptjb) {
      realisasiPerSptjb[idSptjb] = (realisasiPerSptjb[idSptjb] || 0) + totalRealisasi;
    }
  }

  // Daftar SPTJB terbaru (maksimal 25 baris, urut dari bawah/terbaru)
  const rows = [];
  for (let r = sptjbData.length - 1; r >= 1; r--) {
    const row = sptjbData[r];
    const idSptjb = row[0];
    if (!idSptjb) continue;
    const jumlahPengajuan = Number(row[13]) || 0; // N: Jumlah_Pengajuan_Rp
    const totalRealisasi = realisasiPerSptjb[idSptjb] || 0;
    rows.push({
      idSptjb: idSptjb,
      noSptjb: row[1],
      kegiatan: row[3],
      provinsi: row[5],
      tanggalMulai: formatTanggal_(row[7]),
      jumlahOrang: orangPerSptjb[idSptjb] || 0,
      jumlahPengajuan: jumlahPengajuan,
      totalRealisasi: totalRealisasi,
      status: row[14] || '-'
    });
    if (rows.length >= 25) break;
  }

  return {
    totalSptjb: Math.max(sptjbData.length - 1, 0),
    totalPengajuan: Math.max(pengajuanData.length - 1, 0),
    totalRealisasi: Math.max(realisasiData.length - 1, 0),
    totalRupiahRealisasi: totalRealisasiKeseluruhan,
    rows: rows
  };
}

function formatTanggal_(v) {
  if (!v) return '-';
  try {
    if (Object.prototype.toString.call(v) === '[object Date]') {
      return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd MMM yyyy');
    }
    return String(v);
  } catch (err) {
    return String(v);
  }
}

function formatRupiah_(n) {
  n = Number(n) || 0;
  return 'Rp' + n.toLocaleString('id-ID');
}