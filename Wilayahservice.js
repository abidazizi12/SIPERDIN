/**
 * WilayahService.gs
 * ------------------------------------------------------------
 * Sumber data kab/kota Indonesia yang LENGKAP (514 kab/kota,
 * 38 provinsi -- data resmi Kemendagri/BPS), diambil dari API statis
 * publik https://emsifa.github.io/api-wilayah-indonesia/.
 *
 * KENAPA TIDAK PAKAI REF_TARIF SAJA: tabel SBM di REF_TARIF (blok
 * TRANSPORT_PESERTA/TAKSI) hanya berisi ~26 kota (kemungkinan cuma
 * ibukota provinsi, bukan daftar administratif lengkap) -- jadi tidak
 * cukup untuk dropdown Daerah yang perlu mencakup semua kabupaten
 * (mis. "Kab. Bandung Barat", "Kab. Purwakarta").
 *
 * Dropdown PROVINSI tetap dari REF_TARIF (lihat getProvinsiList di
 * PengajuanService.gs) karena namanya harus PERSIS cocok dengan yang
 * dipakai untuk mencari tarif SBM. Tapi ejaan provinsi di REF_TARIF
 * agak tidak baku (mis. "Sumatra Utara" bukan "Sumatera Utara", ada
 * juga yang hurufnya terpisah spasi seperti "B A L I") -- jadi perlu
 * PROVINSI_ALIAS di bawah untuk mencocokkan ke nama resmi API.
 *
 * Kalau ada provinsi baru/berubah di REF_TARIF dan tidak ada di
 * alias ini, getDaerahByProvinsi akan mengembalikan array kosong --
 * tambahkan pemetaannya di PROVINSI_ALIAS.
 * ------------------------------------------------------------
 */

var WILAYAH_API_BASE = 'https://emsifa.github.io/api-wilayah-indonesia/api';

// REF_TARIF (kiri) -> nama resmi di API wilayah (kanan, huruf besar)
var PROVINSI_ALIAS = {
  'aceh': 'ACEH',
  'bali': 'BALI',
  'banten': 'BANTEN',
  'bangkabelitung': 'KEPULAUAN BANGKA BELITUNG',
  'bengkulu': 'BENGKULU',
  'diyogyakarta': 'DI YOGYAKARTA',
  'dkijakarta': 'DKI JAKARTA',
  'gorontalo': 'GORONTALO',
  'jambi': 'JAMBI',
  'jawabarat': 'JAWA BARAT',
  'jawatengah': 'JAWA TENGAH',
  'jawatimur': 'JAWA TIMUR',
  'kalimantanbarat': 'KALIMANTAN BARAT',
  'kalimantanselatan': 'KALIMANTAN SELATAN',
  'kalimantantengah': 'KALIMANTAN TENGAH',
  'kalimantantimur': 'KALIMANTAN TIMUR',
  'kalimantanutara': 'KALIMANTAN UTARA',
  'kepulauanriau': 'KEPULAUAN RIAU',
  'lampung': 'LAMPUNG',
  'maluku': 'MALUKU',
  'malukuutara': 'MALUKU UTARA',
  'nusatenggarabarat': 'NUSA TENGGARA BARAT',
  'nusatenggaratimur': 'NUSA TENGGARA TIMUR',
  'papua': 'PAPUA',
  'papuabarat': 'PAPUA BARAT',
  'papuabaratdaya': 'PAPUA BARAT DAYA',
  'papuapegunungan': 'PAPUA PEGUNUNGAN',
  'papuaselatan': 'PAPUA SELATAN',
  'papuatengah': 'PAPUA TENGAH',
  'riau': 'RIAU',
  'sulawesibarat': 'SULAWESI BARAT',
  'sulawesiselatan': 'SULAWESI SELATAN',
  'sulawesitengah': 'SULAWESI TENGAH',
  'sulawesitenggara': 'SULAWESI TENGGARA',
  'sulawesiutara': 'SULAWESI UTARA',
  'sumatrabarat': 'SUMATERA BARAT',
  'sumatraselatan': 'SUMATERA SELATAN',
  'sumatrautara': 'SUMATERA UTARA'
};

/**
 * Dipanggil dari client (google.script.run) setiap kali dropdown
 * Provinsi berubah, untuk mengisi dropdown Daerah supaya HANYA
 * menampilkan kab/kota di provinsi itu.
 * @param {string} provinsiRefTarif nama provinsi persis seperti di REF_TARIF
 * @return {string[]} daftar kab/kota, mis. ["Kab. Bandung Barat", "Kota Bekasi", ...]
 */
function getDaerahByProvinsi(provinsiRefTarif) {
  try {
    const namaResmi = cariNamaProvinsiResmi_(provinsiRefTarif);
    if (!namaResmi) return [];

    const provinces = ambilProvincesWilayah_();
    const prov = provinces.find(function (p) { return p.name === namaResmi; });
    if (!prov) return [];

    const regencies = ambilRegenciesWilayah_(prov.id);
    return regencies.map(formatNamaKabKota_).sort();
  } catch (e) {
    console.error('Error getDaerahByProvinsi: ' + e.message);
    return [];
  }
}

function cariNamaProvinsiResmi_(provinsiRefTarif) {
  const key = String(provinsiRefTarif || '').toLowerCase().replace(/[^a-z]/g, '');
  return PROVINSI_ALIAS[key] || null;
}

/**
 * "KABUPATEN BANDUNG BARAT" -> "Kab. Bandung Barat"
 * "KOTA BEKASI" -> "Kota Bekasi"
 */
function formatNamaKabKota_(regency) {
  let nama = regency.name;
  let prefix = '';
  if (/^KABUPATEN\s/i.test(nama)) {
    prefix = 'Kab. ';
    nama = nama.replace(/^KABUPATEN\s/i, '');
  } else if (/^KOTA\s/i.test(nama)) {
    prefix = 'Kota ';
    nama = nama.replace(/^KOTA\s/i, '');
  }
  const titleCase = nama.toLowerCase().replace(/\b\w/g, function (c) { return c.toUpperCase(); });
  return prefix + titleCase;
}

function ambilProvincesWilayah_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('WILAYAH_PROVINCES');
  if (cached) return JSON.parse(cached);

  const resp = UrlFetchApp.fetch(WILAYAH_API_BASE + '/provinces.json', { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('Gagal ambil daftar provinsi dari API wilayah (HTTP ' + resp.getResponseCode() + ')');
  const data = JSON.parse(resp.getContentText());
  cache.put('WILAYAH_PROVINCES', JSON.stringify(data), 21600); // cache 6 jam
  return data;
}

function ambilRegenciesWilayah_(provinceId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'WILAYAH_REGENCIES_' + provinceId;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const resp = UrlFetchApp.fetch(WILAYAH_API_BASE + '/regencies/' + provinceId + '.json', { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('Gagal ambil daftar kab/kota dari API wilayah (HTTP ' + resp.getResponseCode() + ')');
  const data = JSON.parse(resp.getContentText());
  cache.put(cacheKey, JSON.stringify(data), 21600); // cache 6 jam
  return data;
}

/**
 * DIAGNOSTIK -- jalankan manual dari editor Apps Script (pilih
 * debugWilayah di dropdown fungsi, klik Run), lalu cek Executions/Logs.
 * Ini SENGAJA tidak pakai try/catch supaya kalau ada error (mis. izin
 * akses eksternal belum di-approve), errornya kelihatan jelas -- beda
 * dari getDaerahByProvinsi yang membungkam error demi keamanan client.
 *
 * PENTING: kalau ini baru pertama kali dijalankan, kemungkinan besar
 * akan muncul popup "Otorisasi diperlukan" di editor -- klik Lanjutkan
 * / Allow. Ini WAJIB dilakukan minimal SEKALI oleh kamu (pemilik
 * script) supaya web app (yang jalan sebagai "Me") boleh mengakses
 * API eksternal untuk SEMUA orang yang buka aplikasinya nanti.
 */
function debugWilayah() {
  Logger.log('=== Tes cariNamaProvinsiResmi_("Jawa Barat") ===');
  Logger.log(cariNamaProvinsiResmi_('Jawa Barat'));

  Logger.log('=== Tes ambilProvincesWilayah_() (fetch ke API eksternal) ===');
  const provinces = ambilProvincesWilayah_();
  Logger.log('Jumlah provinsi dari API: ' + provinces.length);
  Logger.log(JSON.stringify(provinces.slice(0, 5)) + ' ...');

  const prov = provinces.find(function (p) { return p.name === 'JAWA BARAT'; });
  Logger.log('=== Cocokkan "JAWA BARAT" di data API ===');
  Logger.log(JSON.stringify(prov));

  if (prov) {
    Logger.log('=== Tes ambilRegenciesWilayah_(' + prov.id + ') ===');
    const regencies = ambilRegenciesWilayah_(prov.id);
    Logger.log('Jumlah kab/kota: ' + regencies.length);
    Logger.log(JSON.stringify(regencies.slice(0, 10)) + ' ...');
  }

  Logger.log('=== Tes getDaerahByProvinsi("Jawa Barat") end-to-end ===');
  Logger.log(JSON.stringify(getDaerahByProvinsi('Jawa Barat')));
}