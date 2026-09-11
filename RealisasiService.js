/**
 * RealisasiService.gs
 * ------------------------------------------------------------
 * Logic server untuk halaman Form Realisasi.
 * ------------------------------------------------------------
 */

// Daftar baris Pengajuan yang bisa dipilih untuk direalisasikan
// (label gabungan: Nama - Tujuan - No SPTJB - Tanggal)
function getPengajuanListUntukRealisasi_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shPengajuan = ss.getSheetByName('DB_PENGAJUAN');
  const shSptjb = ss.getSheetByName('DB_SPTJB');
  if (!shPengajuan) return [];

  const noSptjbMap = {};
  if (shSptjb) {
    const sptjbData = shSptjb.getDataRange().getValues();
    for (let r = 1; r < sptjbData.length; r++) {
      if (sptjbData[r][0]) noSptjbMap[sptjbData[r][0]] = sptjbData[r][1];
    }
  }

  const data = shPengajuan.getDataRange().getValues();
  const out = [];
  for (let r = data.length - 1; r >= 1; r--) {
    const idBaris = data[r][0];
    if (!idBaris) continue;
    out.push({
      idBaris: idBaris,
      nama: data[r][8],
      tujuan: data[r][14],
      provinsi: data[r][15],
      noSptjb: noSptjbMap[data[r][3]] || '',
      tglBerangkat: formatTanggal_(data[r][17])
    });
    if (out.length >= 300) break; // batasi biar dropdown tidak terlalu berat
  }
  return out;
}

/**
 * Ambil saran tarif MAKSIMAL (batas SBM) untuk satu provinsi, dari sheet
 * REF_TARIF (blok-blok berjejer samping). Dipakai untuk PRE-FILL form
 * Realisasi -- tetap bisa diubah manual karena tarif SBM adalah batas
 * atas, bukan angka pasti yang harus dipakai penuh.
 */
function getTarifSaran(provinsi) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REF_TARIF');
  if (!sh || !provinsi) return { uangHarian: 0, tarifHotel: 0 };

  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  const titles = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const headers = sh.getRange(2, 1, 1, lastCol).getValues()[0];
  const data = sh.getRange(1, 1, lastRow, lastCol).getValues();

  // deteksi batas tiap blok dari baris judul (merge cell -> hanya kolom
  // pertama blok yang terisi, kolom kosong tanpa header = pemisah blok)
  const blocks = {};
  let curTitle = null, curStart = null;
  for (let c = 0; c < lastCol; c++) {
    if (titles[c]) {
      if (curTitle) blocks[curTitle] = { start: curStart, end: c - 1 };
      curTitle = titles[c];
      curStart = c;
    } else if (!headers[c] && curTitle) {
      blocks[curTitle] = { start: curStart, end: c - 1 };
      curTitle = null;
    }
  }
  if (curTitle) blocks[curTitle] = { start: curStart, end: lastCol - 1 };

  function cariDiBlok(blockTitle, fieldName) {
    const b = blocks[blockTitle];
    if (!b) return null;
    const hdrs = headers.slice(b.start, b.end + 1);
    const provIdx = hdrs.indexOf('Provinsi');
    const fieldIdx = hdrs.indexOf(fieldName);
    if (provIdx === -1 || fieldIdx === -1) return null;
    const target = String(provinsi).trim().toLowerCase();
    for (let r = 2; r < data.length; r++) {
      const rowProv = String(data[r][b.start + provIdx] || '').trim().toLowerCase();
      if (rowProv === target) return data[r][b.start + fieldIdx];
    }
    return null;
  }

  return {
    uangHarian: Number(cariDiBlok('UH PEGAWAI', 'UH_Luar_Kota')) || 0,
    tarifHotel: Number(cariDiBlok('HOTEL PEJABAT', 'Tarif_Eselon_IV_Gol_III_II_I')) || 0
  };
}

/**
 * Submit form Realisasi. formData = { idBaris, ...field DB_REALISASI }
 */
function submitRealisasi(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sh = ss.getSheetByName('DB_REALISASI');
    const idRealisasi = 'REAL-' + nextSequence_('REAL');

    const totalTransport = (Number(formData.transportKedudukan) || 0) +
      (Number(formData.transportTujuan) || 0) + (Number(formData.transportLainnya) || 0);
    const totalHotel = (Number(formData.tarifHotel) || 0) * (Number(formData.malam) || 0);
    const totalUangHarian = (Number(formData.uangHarianSatuan) || 0) * (Number(formData.malam || formData.hari) || 0);
    const totalRealisasi = totalTransport + totalHotel + totalUangHarian +
      (Number(formData.hargaTiketBerangkat) || 0) + (Number(formData.hargaTiketPulang) || 0) +
      (Number(formData.uangRepresentative) || 0) + (Number(formData.totalHargaKontrak) || 0);

    const row = [
      idRealisasi, formData.idBaris,
      // kolom C (No_Akun) SENGAJA DILEWATI -- formula otomatis
      formData.noSk || '', formData.noSt || '', formData.tanggalSk || '', formData.tanggalSt || '',
      formData.saranaTransportasi || '',
      Number(formData.transportKedudukan) || 0, Number(formData.transportTujuan) || 0, Number(formData.transportLainnya) || 0,
      totalTransport,
      formData.tanggalBerangkat || '', Number(formData.hargaTiketBerangkat) || 0, formData.noTiketBerangkat || '', formData.seatBerangkat || '',
      formData.tanggalPulang || '', Number(formData.hargaTiketPulang) || 0, formData.noTiketPulang || '', formData.seatPulang || '',
      Number(formData.tarifHotel) || 0, Number(formData.malam) || 0, formData.namaHotel || '', totalHotel,
      Number(formData.uangHarianSatuan) || 0, totalUangHarian,
      Number(formData.uangRepresentative) || 0,
      formData.tanggalKontrak || '', Number(formData.hargaKontrak) || 0, formData.paxUnit || '', Number(formData.totalHargaKontrak) || 0,
      totalRealisasi,
      formData.kelengkapanSt ? 'Ya' : 'Belum',
      formData.kelengkapanSpj ? 'Ya' : 'Belum',
      formData.kelengkapanBast ? 'Ya' : 'Belum',
      formData.kelengkapanSk ? 'Ya' : 'Belum',
      formData.kelengkapanLaporan ? 'Ya' : 'Belum',
      'Selesai',
      Session.getActiveUser().getEmail(), new Date(), ''
    ];

    const startRow = sh.getLastRow() + 1;
    // tulis kolom A-B dulu, lalu D-AO (lewati kolom C/No_Akun yang formula)
    sh.getRange(startRow, 1, 1, 2).setValues([[row[0], row[1]]]);
    sh.getRange(startRow, 4, 1, row.length - 2).setValues([row.slice(2)]);

    return { ok: true, idRealisasi: idRealisasi, totalRealisasi: totalRealisasi };
  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}