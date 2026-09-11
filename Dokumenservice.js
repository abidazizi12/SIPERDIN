/**
 * DokumenService.gs
 * ------------------------------------------------------------
 * Logic untuk tab "Print" di halaman Pengajuan -- daftar SPTJB yang
 * siap dicetak, dan generator dokumennya.
 *
 * STATUS PER DOKUMEN:
 * - Daftar Nominatif: SUDAH JALAN (cetakDaftarNominatif). Formatnya
 *   diambil persis dari contoh PDF "Nominatif Perjadin" yang pernah
 *   dibagikan.
 * - SPTJB, SPBY, Kwitansi: BELUM DIBUAT. Saya belum pernah lihat
 *   format resmi dokumen-dokumen ini (cuma tahu namanya dari obrolan
 *   awal) -- daripada mengarang format dokumen resmi yang salah,
 *   fungsinya sengaja dibuat menolak dulu (lihat cetakSptjb dkk di
 *   bawah). Share contoh PDF-nya kalau mau saya bangunkan.
 * ------------------------------------------------------------
 */

/**
 * Daftar SPTJB untuk ditampilkan di tab Print (dipanggil server-side
 * saat render halaman, lihat Code.gs).
 */
function getSptjbListUntukPrint_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shSptjb = ss.getSheetByName('DB_SPTJB');
  const shPengajuan = ss.getSheetByName('DB_PENGAJUAN');
  if (!shSptjb) return [];

  const dataS = shSptjb.getDataRange().getValues();
  const dataP = shPengajuan ? shPengajuan.getDataRange().getValues() : [];

  const jumlahOrangPerSptjb = {};
  const jumlahTripPerSptjb = {};
  const tripSetPerSptjb = {};
  for (let r = 1; r < dataP.length; r++) {
    const idSptjb = dataP[r][3];
    const idTrip = dataP[r][1];
    if (!idSptjb) continue;
    jumlahOrangPerSptjb[idSptjb] = (jumlahOrangPerSptjb[idSptjb] || 0) + 1;
    if (!tripSetPerSptjb[idSptjb]) tripSetPerSptjb[idSptjb] = {};
    if (idTrip && !tripSetPerSptjb[idSptjb][idTrip]) {
      tripSetPerSptjb[idSptjb][idTrip] = true;
      jumlahTripPerSptjb[idSptjb] = (jumlahTripPerSptjb[idSptjb] || 0) + 1;
    }
  }

  const out = [];
  for (let r = dataS.length - 1; r >= 1; r--) {
    const idSptjb = dataS[r][0];
    if (!idSptjb) continue;
    out.push({
      idSptjb: idSptjb,
      noSptjb: dataS[r][1],
      kegiatan: dataS[r][3],
      noAkun: dataS[r][9],
      jumlahTrip: jumlahTripPerSptjb[idSptjb] || 0,
      jumlahOrang: jumlahOrangPerSptjb[idSptjb] || 0
    });
  }
  return out;
}

/**
 * Generate dokumen "Daftar Nominatif" (Google Docs) untuk 1 SPTJB --
 * menggabungkan SEMUA Surat Tugas & orang di bawah SPTJB itu jadi 1
 * tabel, persis format contoh PDF yang dibagikan.
 * @return {Object} {ok:true, url:'...'} atau {ok:false, error:'...'}
 */
function cetakDaftarNominatif(idSptjb) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shSptjb = ss.getSheetByName('DB_SPTJB');
    const shPengajuan = ss.getSheetByName('DB_PENGAJUAN');
    if (!shSptjb || !shPengajuan) return { ok: false, error: 'Sheet DB_SPTJB/DB_PENGAJUAN tidak ditemukan.' };

    const dataS = shSptjb.getDataRange().getValues();
    let sptjb = null;
    for (let r = 1; r < dataS.length; r++) {
      if (dataS[r][0] === idSptjb) {
        sptjb = {
          noSptjb: dataS[r][1], kegiatan: dataS[r][3], noAkun: dataS[r][9], namaAkun: dataS[r][11],
          klasifikasiAnggaran: dataS[r][18]
        };
        break;
      }
    }
    if (!sptjb) return { ok: false, error: 'SPTJB tidak ditemukan.' };

    const dataP = shPengajuan.getDataRange().getValues();
    const baris = [];
    const tujuanSet = {};
    let provinsi = '';
    for (let r = 1; r < dataP.length; r++) {
      if (dataP[r][3] !== idSptjb) continue;
      baris.push({
        nama: dataP[r][8], tujuan: dataP[r][14], provinsi: dataP[r][15],
        tglBerangkat: dataP[r][17], tglKembali: dataP[r][18], lamaHari: dataP[r][19],
        transportPP: dataP[r][30], transportKedudukan: dataP[r][31], transportTujuan: dataP[r][32],
        uangHarian: dataP[r][33], penginapan: dataP[r][34], uangRepresentatif: dataP[r][35],
        total: dataP[r][36]
      });
      if (dataP[r][14]) tujuanSet[dataP[r][14]] = true;
      provinsi = dataP[r][15] || provinsi;
    }
    if (baris.length === 0) return { ok: false, error: 'Belum ada data orang untuk SPTJB ini.' };

    const config = getConfig_();
    const daftarTujuan = Object.keys(tujuanSet).join(', ');

    const doc = DocumentApp.create('Daftar Nominatif - ' + sptjb.noSptjb + ' - ' + new Date().getTime());
    const body = doc.getBody();
    body.setPageWidth(841.89).setPageHeight(595.28); // A4 landscape (points)

    body.appendParagraph('Daftar Nominatif : Perjalanan Dinas dalam rangka ' + sptjb.kegiatan +
      ' di ' + daftarTujuan + (provinsi ? ', Provinsi ' + provinsi : ''))
      .setHeading(DocumentApp.ParagraphHeading.NORMAL).editAsText().setBold(true);
    body.appendParagraph('Dengan nomor : ' + sptjb.noSptjb);
    body.appendParagraph('1  Tanggal/No.DIPA : ' + (config.no_dipa || '-') + (config.tanggal_dipa ? ' Tanggal ' + config.tanggal_dipa : ''));
    body.appendParagraph('2  Klasifikasi Anggaran : ' + (sptjb.klasifikasiAnggaran || '-'));
    body.appendParagraph('3  Akun : ' + (sptjb.namaAkun || '-') + ' ' + (sptjb.noAkun || ''));
    body.appendParagraph('');

    const headers = ['No', 'Nama', 'Keberangkatan', 'Tgl Berangkat', 'Tgl Kembali', 'Lama',
      'Transport PP', 'Transport Kedudukan', 'Transport Tujuan', 'Uang Harian', 'Penginapan', 'Uang Representatif', 'Total'];
    const tableData = [headers];

    let totalKolom = [0, 0, 0, 0, 0, 0, 0];
    baris.forEach(function (b, idx) {
      const row = [
        String(idx + 1), b.nama, 'Jakarta - ' + b.tujuan,
        formatTanggalInput_(b.tglBerangkat), formatTanggalInput_(b.tglKembali), (b.lamaHari || '') + ' Hari',
        formatRupiah_(b.transportPP), formatRupiah_(b.transportKedudukan), formatRupiah_(b.transportTujuan),
        formatRupiah_(b.uangHarian), formatRupiah_(b.penginapan), formatRupiah_(b.uangRepresentatif),
        formatRupiah_(b.total)
      ];
      tableData.push(row);
      totalKolom[0] += Number(b.transportPP) || 0;
      totalKolom[1] += Number(b.transportKedudukan) || 0;
      totalKolom[2] += Number(b.transportTujuan) || 0;
      totalKolom[3] += Number(b.uangHarian) || 0;
      totalKolom[4] += Number(b.penginapan) || 0;
      totalKolom[5] += Number(b.uangRepresentatif) || 0;
      totalKolom[6] += Number(b.total) || 0;
    });
    tableData.push([
      '', '', '', '', '', 'JUMLAH',
      formatRupiah_(totalKolom[0]), formatRupiah_(totalKolom[1]), formatRupiah_(totalKolom[2]),
      formatRupiah_(totalKolom[3]), formatRupiah_(totalKolom[4]), formatRupiah_(totalKolom[5]), formatRupiah_(totalKolom[6])
    ]);

    const table = body.appendTable(tableData);
    table.getRow(0).editAsText().setBold(true);
    const lastRow = table.getNumRows() - 1;
    table.getRow(lastRow).editAsText().setBold(true);

    body.appendParagraph('');
    body.appendParagraph((config.kota_penandatanganan || 'Jakarta') + ', ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM yyyy'));
    body.appendParagraph('');
    const tandaTangan = body.appendTable([
      ['Setuju Dibayarkan,', 'Yang Bertanggungjawab', ''],
      ['Pejabat Pembuat Komitmen', 'Ketua Tim Kerja', 'Bendahara Pengeluaran'],
      ['', '', ''],
      ['', '', ''],
      [config.nama_ppk || '(Nama PPK -- isi di Konfigurasi)', config.nama_ketua_tim_kerja || '(Nama Ketua Tim Kerja -- isi di Konfigurasi)', config.nama_bendahara || '(Nama Bendahara -- isi di Konfigurasi)'],
      ['NIP. ' + (config.nip_ppk || '-'), 'NIP. ' + (config.nip_ketua_tim_kerja || '-'), 'NIP. ' + (config.nip_bendahara || '-')]
    ]);
    tandaTangan.setBorderWidth(0);

    doc.saveAndClose();
    return { ok: true, url: doc.getUrl() };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

/**
 * Placeholder -- belum dibuat, lihat catatan di atas file ini.
 */
function cetakSptjb(idSptjb) {
  return { ok: false, error: 'Format dokumen SPTJB resmi belum dikonfirmasi. Share contoh PDF-nya dulu ke Claude.' };
}
function cetakSpby(idSptjb) {
  return { ok: false, error: 'Format dokumen SPBY resmi belum dikonfirmasi. Share contoh PDF-nya dulu ke Claude.' };
}
function cetakKwitansi(idSptjb) {
  return { ok: false, error: 'Format dokumen Kwitansi resmi belum dikonfirmasi. Share contoh PDF-nya dulu ke Claude.' };
}
