/**
 * DummyData.gs
 * ------------------------------------------------------------
 * Jalankan fungsi seedDummyData() SEKALI dari editor Apps Script
 * untuk isi DB_SPTJB, DB_PENGAJUAN, & DB_REALISASI dengan data
 * contoh -- supaya ada isi untuk menggambarkan tampilan aplikasi
 * (Dashboard, dll) selama masa pengembangan.
 *
 * AMAN dijalankan di sheet yang sudah kosong (cuma header) -- fungsi
 * ini akan MENGHAPUS dulu semua baris data yang ada di ketiga sheet
 * itu (baris 2 ke bawah, header di baris 1 tetap aman), baru menulis
 * ulang data dummy dari nol. Kalau kamu masih punya data ASLI di
 * situ, JANGAN jalankan fungsi ini.
 *
 * SEMUA ANGKA NOMINATIF DI BAWAH DIHITUNG MANUAL DARI REF_TARIF ASLI
 * (bukan angka bulat karangan), supaya representatif:
 *
 * Trip 1 -- Jawa Barat, Kab. Bandung Barat, 2 hari (1 malam):
 *   UH Luar Kota Jawa Barat = 430.000 x 2 hari = 860.000
 *   Taksi kedudukan (DKI Jakarta) = 250.000, taksi tujuan (Jawa Barat) = 180.000
 *   Hotel: Jaelani (gol.hotel 3, Eselon III/Gol IV) = 1.366.000 x 1 malam
 *          Juwariyah (gol.hotel 4, Eselon IV/Gol III-II-I) = 735.000 x 1 malam
 *   Transport PP = 0 (Jakarta-Bandung tidak ada rute pesawat di REF_TIKET_PESAWAT,
 *                  jarak dekat -- umum ditempuh darat)
 *
 * Trip 2 -- Sumatra Utara, Kota Medan, 3 hari (2 malam), 3 orang gol.hotel 4:
 *   UH Luar Kota Sumatra Utara = 370.000 x 3 hari = 1.110.000
 *   Taksi kedudukan = 250.000, taksi tujuan (Sumatra Utara) = 278.000
 *   Hotel (Eselon IV/Gol III-II-I Sumut) = 699.000 x 2 malam = 1.398.000
 *   Tiket pesawat Jakarta-Medan kelas EKONOMI (semua golongan < III/d) = 4.054.000
 *
 * Trip 3 -- D.I. Yogyakarta, Kota Yogyakarta, 2 hari (1 malam), 3 orang gol.hotel 4:
 *   UH Luar Kota DIY = 420.000 x 2 hari = 840.000
 *   Taksi kedudukan = 250.000, taksi tujuan (DIY) = 258.000
 *   Hotel (Eselon IV/Gol III-II-I DIY) = 845.000 x 1 malam
 *   Tiket pesawat Jakarta-Yogyakarta kelas EKONOMI = 2.268.000
 *
 * Nama & NIP pegawai diambil dari REF_PEGAWAI asli (Jaelani Efendi,
 * Juwariyah, Nidhomul Haq, dkk) supaya matching NIP/golongan nyambung
 * benar kalau nanti dites di form Realisasi.
 * ------------------------------------------------------------
 */
function seedDummyData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shSptjb = ss.getSheetByName('DB_SPTJB');
  const shPengajuan = ss.getSheetByName('DB_PENGAJUAN');
  const shRealisasi = ss.getSheetByName('DB_REALISASI');
  if (!shSptjb || !shPengajuan) throw new Error('Sheet DB_SPTJB atau DB_PENGAJUAN tidak ditemukan.');

  clearDataRows_(shSptjb);
  clearDataRows_(shPengajuan);
  if (shRealisasi) clearDataRows_(shRealisasi);

  // ---- 3 SPTJB / kegiatan dummy (jumlahPengajuan = total nominatif trip) ----
  const sptjbRows = [
    ['SPTJB-2026-00001', '3/1/PK.03.03/VIII/2026', 'ND-001/2026', 'Koordinasi Perluasan Kesempatan Kerja TKM', 'Dalam Negeri',
      'Jawa Barat', 'Kab. Bandung Barat', '2026-08-05', '2026-08-06', '524111', '', 'Belanja Perjalanan Dinas Biasa', '',
      4681000, 'Aktif', 'demo@siperdin.local', new Date(), '', '2175.BDC.003.0B.524111'],
    ['SPTJB-2026-00002', '3/2/PK.03.03/VIII/2026', 'ND-002/2026', 'Pembinaan Tenaga Kerja Mandiri Pemula', 'Dalam Negeri',
      'Sumatra Utara', 'Kota Medan', '2026-08-12', '2026-08-14', '524111', '', 'Belanja Perjalanan Dinas Biasa', '',
      21270000, 'Aktif', 'demo@siperdin.local', new Date(), '', '2175.BDC.003.0A.524111'],
    ['SPTJB-2026-00003', '3/3/PK.03.03/IX/2026', 'ND-003/2026', 'Monitoring & Evaluasi Program TKM', 'Dalam Negeri',
      'D.I. Yogyakarta', 'Kota Yogyakarta', '2026-09-02', '2026-09-03', '521811', '', 'Belanja Bahan', '',
      13383000, 'Aktif', 'demo@siperdin.local', new Date(), '', '2175.BDC.003.0A.521811']
  ];
  shSptjb.getRange(2, 1, sptjbRows.length, sptjbRows[0].length).setValues(sptjbRows);

  // ---- Orang per trip -- nominatif = [transportPP, transportKedudukan, transportTujuan, uangHarian, penginapan, uangRepresentatif, total] ----
  const orangPerTrip = [
    { idSptjb: 'SPTJB-2026-00001', tujuan: 'Kab. Bandung Barat', provinsi: 'Jawa Barat', daerah: 'Kab. Bandung Barat',
      tglBerangkat: '2026-08-05', tglKembali: '2026-08-06', noAkun: '524111', hotelVendor: 'Hotel Grand Sunshine',
      orang: [
        { nama: 'Jaelani Efendi S.P.', nip: '198002112006041001', tipe: 'Penyelenggara', nominatif: [0, 250000, 180000, 860000, 1366000, 0, 2656000] },
        { nama: 'Juwariyah S.H.I., L.L.M.', nip: '198506152011012014', tipe: 'Penyelenggara', nominatif: [0, 250000, 180000, 860000, 735000, 0, 2025000] }
      ] },
    { idSptjb: 'SPTJB-2026-00002', tujuan: 'Kota Medan', provinsi: 'Sumatra Utara', daerah: 'Kota Medan',
      tglBerangkat: '2026-08-12', tglKembali: '2026-08-14', noAkun: '524111', hotelVendor: 'Hotel Antares Medan',
      orang: [
        { nama: 'Nidhomul Haq S.Psi.', nip: '198808312018011001', tipe: 'Penyelenggara', nominatif: [4054000, 250000, 278000, 1110000, 1398000, 0, 7090000] },
        { nama: 'Syahranitazli S.Kom.', nip: '20001023 2025052005', tipe: 'Peserta', nominatif: [4054000, 250000, 278000, 1110000, 1398000, 0, 7090000] },
        { nama: 'Dian Kusumawati S.Sos.', nip: '198601202009122004', tipe: 'Peserta', nominatif: [4054000, 250000, 278000, 1110000, 1398000, 0, 7090000] }
      ] },
    { idSptjb: 'SPTJB-2026-00003', tujuan: 'Kota Yogyakarta', provinsi: 'D.I. Yogyakarta', daerah: 'Kota Yogyakarta',
      tglBerangkat: '2026-09-02', tglKembali: '2026-09-03', noAkun: '521811', hotelVendor: '',
      orang: [
        { nama: 'Bagas Suryo Pambuka S.E.', nip: '198205222009011006', tipe: 'Narasumber', nominatif: [2268000, 250000, 258000, 840000, 845000, 0, 4461000] },
        { nama: 'Muhammad Sadad Mahmud', nip: '-', tipe: 'Penyelenggara', nominatif: [2268000, 250000, 258000, 840000, 845000, 0, 4461000] },
        { nama: 'Muhamad Irfan Ardiansyah', nip: '-', tipe: 'Peserta', nominatif: [2268000, 250000, 258000, 840000, 845000, 0, 4461000] }
      ] }
  ];

  const rows = [];
  const idBarisList = [];
  let barisCounter = 1;
  let tripCounter = 1;
  let pengajuanCounter = 1;

  orangPerTrip.forEach(function (trip) {
    const idTrip = 'TRIP-' + String(tripCounter++).padStart(5, '0');
    const idPengajuan = 'PENG-2026-' + String(pengajuanCounter++).padStart(5, '0');
    const lamaHari = hitungLamaHari_(trip.tglBerangkat, trip.tglKembali);

    trip.orang.forEach(function (o) {
      const idBaris = 'BR-' + String(barisCounter++).padStart(5, '0');
      const n = o.nominatif;
      rows.push([
        idBaris, idTrip, idPengajuan, trip.idSptjb,
        trip.noAkun, '', 'Tahap 1', 'Kegiatan Koordinasi TKM',
        o.nama, o.nip, o.tipe, '', '',
        '', trip.tujuan, trip.provinsi, trip.daerah,
        trip.tglBerangkat, trip.tglKembali, lamaHari,
        trip.hotelVendor, trip.orang.length,
        'Menunggu', '', '', '',
        'demo@siperdin.local', new Date(), '',
        '3.4/457/PK.03/VIII/2026',
        n[0], n[1], n[2], n[3], n[4], n[5], n[6]
      ]);
      idBarisList.push({ idBaris: idBaris, noAkun: trip.noAkun, tglBerangkat: trip.tglBerangkat, tglKembali: trip.tglKembali, lamaHari: lamaHari, hotelVendor: trip.hotelVendor, n: n });
    });
  });

  shPengajuan.getRange(2, 1, rows.length, rows[0].length).setValues(rows);

  // ---- DB_REALISASI: isi untuk 5 dari 8 baris (sisanya sengaja
  // dibiarkan "belum direalisasikan" supaya Dashboard menggambarkan
  // variasi status) -- angka disamakan dengan nominatif pengajuan ----
  if (shRealisasi) {
    const realisasiRows = [];
    idBarisList.slice(0, 5).forEach(function (item, idx) {
      const n = item.n;
      const malam = Math.max(item.lamaHari - 1, 0);
      const tarifHotel = malam > 0 ? Math.round(n[4] / malam) : 0;
      const totalTransport = n[1] + n[2];
      const totalHotel = n[4];
      const uangHarianSatuan = item.lamaHari > 0 ? Math.round(n[3] / item.lamaHari) : 0;
      const totalUangHarian = n[3];
      const totalRealisasi = n[0] + totalTransport + totalHotel + totalUangHarian + n[5];

      realisasiRows.push([
        'REAL-' + String(idx + 1).padStart(5, '0'), item.idBaris, item.noAkun, '', '3.4/457/PK.03/VIII/2026',
        '', '', n[0] > 0 ? 'Pesawat' : 'Darat', n[1], n[2], 0, totalTransport,
        item.tglBerangkat, n[0], n[0] > 0 ? ('JT-' + (1000 + idx)) : '', n[0] > 0 ? '12A' : '',
        item.tglKembali, 0, '', '',
        tarifHotel, malam, item.hotelVendor || '-', totalHotel,
        uangHarianSatuan, totalUangHarian, n[5],
        '', 0, 1, 0,
        totalRealisasi,
        'Ya', 'Ya', 'Ya', 'Ya', 'Ya',
        'Selesai', 'demo@siperdin.local', new Date(), ''
      ]);
    });
    shRealisasi.getRange(2, 1, realisasiRows.length, realisasiRows[0].length).setValues(realisasiRows);
    Logger.log('DB_REALISASI diisi: ' + realisasiRows.length + ' baris.');
  }

  Logger.log('Dummy data selesai: ' + sptjbRows.length + ' SPTJB, ' + rows.length + ' baris orang di DB_PENGAJUAN.');
}

function clearDataRows_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
}