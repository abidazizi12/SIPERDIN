/**
 * ConfigService.gs
 * ------------------------------------------------------------
 * Menyimpan data institusional tetap (Satker, DIPA, NPWP, pejabat
 * penandatangan) yang dipakai berulang di semua dokumen cetak
 * (SPTJB, SPBY, Kwitansi, Daftar Nominatif) -- diisi 1x lewat
 * halaman Admin, bukan diketik ulang tiap submit Pengajuan.
 * ------------------------------------------------------------
 */

// Daftar field CONFIG beserta label yang ditampilkan di form Admin
const CONFIG_FIELDS = [
  { key: 'nama_satker', label: 'Nama Satuan Kerja', contoh: 'DIT. BINA PKK' },
  { key: 'kode_satker', label: 'Kode Satuan Kerja', contoh: '026.04.626041' },
  { key: 'no_dipa', label: 'Nomor DIPA', contoh: '026.04.1.626041/2026' },
  { key: 'tanggal_dipa', label: 'Tanggal DIPA', contoh: '1 Desember 2025' },
  { key: 'npwp', label: 'NPWP', contoh: '63.595.125.4-063.000' },
  { key: 'alamat_satker', label: 'Alamat Satker', contoh: 'Jl. Gatot Soebroto Kav. 51 Jakarta Selatan' },
  { key: 'kota_penandatanganan', label: 'Kota Penandatanganan', contoh: 'Jakarta' },
  { key: 'nama_ketua_tim_kerja', label: 'Nama Ketua Tim Kerja', contoh: 'Jaelani Efendi S.P.' },
  { key: 'nip_ketua_tim_kerja', label: 'NIP Ketua Tim Kerja', contoh: '19800211 200604 1 001' },
  { key: 'nama_ppk', label: 'Nama Pejabat Pembuat Komitmen (PPK)', contoh: 'Zulfikar Khomeini S.E.' },
  { key: 'nip_ppk', label: 'NIP PPK', contoh: '19800331 201101 1 007' },
  { key: 'nama_bendahara', label: 'Nama Bendahara Pengeluaran', contoh: 'Maylina Rahmad Eka Syahputri' },
  { key: 'nip_bendahara', label: 'NIP Bendahara', contoh: '19980512 201812 2 001' },
  { key: 'kode_program', label: 'Kode Program/Klasifikasi Rincian Output', contoh: '2175' }
];

function setupConfigV1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('CONFIG');
  if (!sh) {
    sh = ss.insertSheet('CONFIG');
  } else {
    sh.clear();
  }
  sh.getRange(1, 1, 1, 3).setValues([['Key', 'Label', 'Value']]);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1E40AF');
  sh.setFrozenRows(1);

  const rows = CONFIG_FIELDS.map(function (f) { return [f.key, f.label, '']; });
  sh.getRange(2, 1, rows.length, 3).setValues(rows);
  sh.autoResizeColumns(1, 3);

  Logger.log('Sheet CONFIG dibuat. Isi kolom Value lewat halaman Admin > Konfigurasi.');
}

function getConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('CONFIG');
  const values = {};
  if (sh) {
    const data = sh.getDataRange().getValues();
    for (let r = 1; r < data.length; r++) {
      if (data[r][0]) values[data[r][0]] = data[r][2] || '';
    }
  }
  // gabungkan dengan definisi field (label + contoh) supaya form tahu apa yang ditampilkan
  return CONFIG_FIELDS.map(function (f) {
    return { key: f.key, label: f.label, contoh: f.contoh, value: values[f.key] || '' };
  });
}

// Dipakai form Pengajuan untuk auto-generate Klasifikasi Anggaran lengkap
function getKodeProgram_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('CONFIG');
  if (!sh) return '';
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === 'kode_program') return data[r][2] || '';
  }
  return '';
}

function simpanKonfigurasi(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('CONFIG');
  if (!sh) return { ok: false, error: 'Sheet CONFIG belum dibuat. Jalankan setupConfigV1() dulu.' };

  try {
    const data = sh.getDataRange().getValues();
    for (let r = 1; r < data.length; r++) {
      const key = data[r][0];
      if (key && formData.hasOwnProperty(key)) {
        sh.getRange(r + 1, 3).setValue(formData[key]);
      }
    }
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}