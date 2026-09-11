/**
 * UserService.gs
 * ------------------------------------------------------------
 * Logic server untuk halaman Kelola User (Admin) & Referensi (Admin).
 * ------------------------------------------------------------
 */

function getUserList_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USERS');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let r = 1; r < data.length; r++) {
    if (!data[r][0]) continue;
    out.push({ email: data[r][0], nama: data[r][1], role: data[r][2], status: data[r][3] });
  }
  return out; // password (kolom F) sengaja TIDAK diikutkan ke client
}

function tambahUser(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('USERS');
  try {
    if (!formData.password) return { ok: false, error: 'Password awal wajib diisi.' };

    const data = sh.getDataRange().getValues();
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][0]).trim().toLowerCase() === String(formData.email).trim().toLowerCase()) {
        return { ok: false, error: 'Email ini sudah terdaftar.' };
      }
    }
    sh.getRange(sh.getLastRow() + 1, 1, 1, 6).setValues([[
      formData.email, formData.nama, formData.role, 'Aktif', new Date(), formData.password
    ]]);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function ubahStatusUser(email, statusBaru) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USERS');
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][0]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      sh.getRange(r + 1, 4).setValue(statusBaru);
      return { ok: true };
    }
  }
  return { ok: false, error: 'User tidak ditemukan.' };
}

/**
 * Dipanggil Admin dari halaman Kelola User untuk set/reset password
 * user tertentu (kolom F: Password).
 */
function setPasswordAdmin_(email, passwordBaru) {
  if (!passwordBaru) return { ok: false, error: 'Password baru tidak boleh kosong.' };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USERS');
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][0]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      sh.getRange(r + 1, 6).setValue(passwordBaru); // kolom F
      return { ok: true };
    }
  }
  return { ok: false, error: 'User tidak ditemukan.' };
}

function getRefCounts_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  function count(name) {
    const sh = ss.getSheetByName(name);
    return sh ? Math.max(sh.getLastRow() - 1, 0) : 0;
  }
  return {
    pegawai: count('REF_PEGAWAI'),
    tarifSheetAda: !!ss.getSheetByName('REF_TARIF'),
    spreadsheetUrl: ss.getUrl()
  };
}