/**
 * SessionService.gs
 * ------------------------------------------------------------
 * Login manual (email + password) menggunakan token sesi, supaya
 * web app bisa dibuka di Incognito / browser mana saja tanpa perlu
 * akun Google yang terdaftar di organisasi.
 *
 * CATATAN KEAMANAN: password disimpan APA ADANYA (plain text) di
 * kolom Password sheet USERS -- cukup untuk tool internal kecil,
 * tapi kalau datanya makin sensitif sebaiknya di-hash. Kabari kalau
 * mau saya tambahkan hashing (Utilities.computeDigest).
 *
 * Sesi disimpan di CacheService (bukan cookie), otomatis kadaluarsa
 * 6 jam (batas maksimal Apps Script) sejak AKTIVITAS TERAKHIR --
 * tiap kali token dipakai (buka halaman), masa berlakunya
 * diperpanjang lagi (sliding expiration).
 * ------------------------------------------------------------
 */

var SESSION_DURATION_SECONDS = 21600; // 6 jam, batas maksimal CacheService

/**
 * Dipanggil dari Login.html lewat google.script.run.
 * @return {Object} {ok:true, token, user} atau {ok:false, error}
 */
function login(email, password) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('USERS');
  if (!sh) return { ok: false, error: 'Sheet USERS tidak ditemukan.' };

  const data = sh.getDataRange().getValues();
  const targetEmail = String(email || '').trim().toLowerCase();

  for (let r = 1; r < data.length; r++) {
    const rowEmail = String(data[r][0] || '').trim().toLowerCase();
    if (rowEmail !== targetEmail) continue;

    const status = data[r][3];
    const passwordTersimpan = String(data[r][5] || ''); // kolom F

    if (status !== 'Aktif') return { ok: false, error: 'Akun ini nonaktif. Hubungi Admin.' };
    if (!passwordTersimpan) return { ok: false, error: 'Akun ini belum punya password. Hubungi Admin untuk di-set-kan.' };
    if (passwordTersimpan !== password) return { ok: false, error: 'Password salah.' };

    const user = { email: data[r][0], nama: data[r][1], role: data[r][2] };
    const token = Utilities.getUuid();
    CacheService.getScriptCache().put('SESSION_' + token, JSON.stringify(user), SESSION_DURATION_SECONDS);
    return { ok: true, token: token, user: user };
  }

  return { ok: false, error: 'Email tidak terdaftar.' };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('SESSION_' + token);
  return { ok: true };
}

/**
 * Ambil user dari token sesi. Kalau valid, masa berlaku sesi
 * diperpanjang lagi 6 jam (sliding expiration).
 */
function getUserByToken_(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const raw = cache.get('SESSION_' + token);
  if (!raw) return null;

  cache.put('SESSION_' + token, raw, SESSION_DURATION_SECONDS); // perpanjang sesi
  return JSON.parse(raw);
}