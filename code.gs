/**
 * =====================================================================
 *  PEMILIHAN OSIS — Backend (code.gs)
 *  Dijalankan di Google Apps Script (Google Sheet container).
 *
 *  CARA DEPLOY (FRONTEND IKUT DI-HOST DI APPS SCRIPT):
 *  1. Buat project Apps Script baru, tempel file ini (code.gs), lalu
 *     buat file HTML baru bernama `index` dan tempel seluruh isi
 *     index.html ke dalamnya.
 *  2. Buka Google Sheet terkait (atau buat via Extensions > Apps Script,
 *     file ini otomatis terhubung ke spreadsheet aktif).
 *  3. Jalankan fungsi `CreateDatabase` sekali (via editor atau menu
 *     "Pemilihan OSIS > Buat Database").
 *  4. Deploy > New deployment > Web app:
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     Buka URL Web App tersebut langsung di browser — halaman aplikasi
 *     (index.html) dan API-nya dilayani dari URL yang sama.
 *
 *  Catatan:
 *  - Password admin disimpan polos sesuai spesifikasi. Untuk produksi
 *    nyata, pertimbangkan hashing.
 *  - Login default awal: username `admin`, password `admin123`.
 *    Segera ubah setelah deploy pertama.
 *  - Foto paslon TIDAK diunggah lewat aplikasi. Admin mengunggah sendiri
 *    ke Google Drive, mengatur akses "Anyone with the link — Viewer",
 *    lalu menempelkan ID file gambarnya di form Paslon. Aplikasi hanya
 *    menyimpan ID file tersebut di spreadsheet.
 * =====================================================================
 */

var DB = {
  Admin: 'Admin',
  Kelas: 'Kelas',
  Pemilih: 'Pemilih',
  Paslon: 'Paslon',
  Pengaturan: 'Pengaturan',
  Vote: 'Vote',
  Log: 'Log'
};

var DEFAULTS = {
  namaSekolah: 'SMA Negeri 1 Contoh',
  logo: '',
  waktuMulai: '',
  waktuTutup: ''
};

// Kolom yang sudah tidak dipakai lagi (dihapus otomatis oleh CreateDatabase
// bila masih ada di spreadsheet lama).
var OBSOLETE_KEYS = ['folderDrive'];

var SCHEMAS = {
  'Admin': ['ID', 'Username', 'Password', 'Nama Lengkap', 'Peran'],
  'Kelas': ['ID', 'Nama Kelas'],
  'Pemilih': ['ID', 'Token', 'Nama Pemilih', 'Jenis Kelamin', 'Kelas'],
  'Paslon': ['ID', 'Nomor Urut', 'Nama Ketua', 'Foto Ketua', 'Nama Wakil', 'Foto Wakil', 'Visi Misi'],
  'Pengaturan': ['ID', 'Kunci', 'Nilai'],
  'Vote': ['ID', 'ID Pemilih', 'ID Paslon', 'Waktu'],
  'Log': ['ID', 'Waktu', 'Username', 'Aksi', 'Target', 'Keterangan']
};

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Pemilihan OSIS')
    .addItem('Buat Database', 'CreateDatabase')
    .addToUi();
}

/* =====================================================================
 *  DATABASE
 * ===================================================================== */

function CreateDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var report = { dibuat: [], diperbarui: [], kolomBaru: [], kolomDihapus: [], pengaturanBaru: [] };

  Object.keys(SCHEMAS).forEach(function (name) {
    var s = ss.getSheetByName(name);
    if (!s) {
      s = ss.insertSheet(name);
      report.dibuat.push(name);
    }
    _syncSheetStructure(s, SCHEMAS[name], report);
  });

  // Pengaturan: tambahkan kunci default yang belum ada, hapus kunci usang.
  Object.keys(DEFAULTS).forEach(function (k) {
    if (_getSetting(k) === '' && _findSettingRow(k) === null) {
      _setSetting(k, DEFAULTS[k]);
      report.pengaturanBaru.push(k);
    }
  });
  OBSOLETE_KEYS.forEach(function (k) {
    var f = _findSettingRow(k);
    if (f) {
      _deleteByIds('Pengaturan', [f.ID]);
      report.kolomDihapus.push('Pengaturan.' + k);
    }
  });

  // Kolom Token & ID pemilih dipaksa berformat TEKS agar angka nol di depan
  // (token seperti 012345) tidak terpotong menjadi 5 digit.
  var pemilihSheet = ss.getSheetByName('Pemilih');
  if (pemilihSheet && pemilihSheet.getMaxRows() > 1) {
    pemilihSheet.getRange(2, 1, pemilihSheet.getMaxRows() - 1, 2).setNumberFormat('@');
  }

  if (_rows('Admin').length === 0) {
    _append('Admin', { ID: _uid(), Username: 'admin', Password: 'admin123', 'Nama Lengkap': 'Administrator', Peran: 'Admin' });
    report.dibuat.push('akun admin/admin123');
  }

  _log('system', 'Setup', 'Database', 'CreateDatabase dijalankan');

  var msg = 'Database siap. ' +
    'Sheet dibuat: ' + (report.dibuat.length ? report.dibuat.join(', ') : '-') + '. ' +
    'Struktur diperbarui: ' + (report.diperbarui.length ? report.diperbarui.join(', ') : '-') + '. ' +
    'Kolom ditambahkan: ' + (report.kolomBaru.length ? report.kolomBaru.join(', ') : '-') + '. ' +
    'Kolom/kunci dihapus: ' + (report.kolomDihapus.length ? report.kolomDihapus.join(', ') : '-') + '.';
  Logger.log(msg);
  return msg;
}

/**
 * Menyamakan header sheet dengan SCHEMAS tanpa merusak data yang sudah ada:
 * - Kolom yang belum ada ditambahkan di kanan (data lama tetap di barisnya).
 * - Kolom yang sudah ada dipertahankan beserta datanya.
 * - Kolom usang (tidak ada di skema) dihapus bila kosong seluruhnya, atau
 *   dipindah ke paling kanan bila masih berisi data (supaya tidak hilang).
 */
function _syncSheetStructure(sheet, schema, report) {
  var name = sheet.getName();
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var current = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) { return String(v).trim(); })
    : [];

  // Sheet baru / header kosong -> tulis header skema apa adanya.
  if (current.filter(String).length === 0) {
    _writeHeader(sheet, schema);
    report.diperbarui.push(name);
    return;
  }

  // 1. Kolom skema yang belum ada -> tambahkan.
  var tambahan = schema.filter(function (c) { return current.indexOf(c) === -1; });

  // 2. Kolom lama yang tidak ada di skema.
  var usang = current.filter(function (c, i) {
    if (!c) return false;
    if (schema.indexOf(c) > -1) return false;
    // jaga duplikat: hanya kolom pertama dengan nama itu yang dianggap valid
    return current.indexOf(c) === i;
  });

  var usangDipakai = [];
  if (usang.length && lastRow > 1) {
    usang.forEach(function (c) {
      var idx = current.indexOf(c);
      var vals = sheet.getRange(2, idx + 1, lastRow - 1, 1).getValues();
      var adaIsi = vals.some(function (r) { return String(r[0]).trim() !== ''; });
      if (adaIsi) usangDipakai.push(c);
    });
  }

  if (tambahan.length || usang.length) {
    var finalHead = schema.slice();
    usangDipakai.forEach(function (c) { if (finalHead.indexOf(c) === -1) finalHead.push(c); });

    // Kumpulkan data lama per nama kolom sebelum menulis ulang.
    var dataLama = {};
    if (lastRow > 1) {
      var rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      current.forEach(function (c, i) {
        if (!c) return;
        if (dataLama[c] !== undefined) return;
        dataLama[c] = rows.map(function (r) { return r[i]; });
      });
    }

    // Perluas kolom bila header baru lebih banyak.
    if (sheet.getMaxColumns() < finalHead.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), finalHead.length - sheet.getMaxColumns());
    }
    _writeHeader(sheet, finalHead);

    if (lastRow > 1) {
      var jumlahBaris = lastRow - 1;
      var out = [];
      for (var r = 0; r < jumlahBaris; r++) {
        out.push(finalHead.map(function (c) {
          return (dataLama[c] !== undefined && dataLama[c][r] !== undefined && dataLama[c][r] !== null) ? dataLama[c][r] : '';
        }));
      }
      sheet.getRange(2, 1, jumlahBaris, finalHead.length).setValues(out);
    }
    // Bersihkan sisa kolom di kanan header baru.
    var sisa = sheet.getMaxColumns() - finalHead.length;
    if (sisa > 0) {
      sheet.getRange(1, finalHead.length + 1, sheet.getMaxRows(), sisa).clearContent();
    }

    report.diperbarui.push(name);
    tambahan.forEach(function (c) { report.kolomBaru.push(name + '.' + c); });
    usang.forEach(function (c) {
      if (usangDipakai.indexOf(c) > -1) report.kolomDihapus.push(name + '.' + c + ' (dipindah ke kanan)');
      else report.kolomDihapus.push(name + '.' + c);
    });
  }
}

function _writeHeader(sheet, heads) {
  if (sheet.getMaxColumns() < heads.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), heads.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, heads.length).setValues([heads]);
  sheet.getRange(1, 1, 1, heads.length)
    .setFontWeight('bold').setBackground('#4f46e5').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
}

function _findSettingRow(k) {
  var rows = _rows('Pengaturan');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].Kunci) === String(k)) return rows[i];
  }
  return null;
}

function _ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function _sheet(name) { return _ss().getSheetByName(name); }

function _headers(name) {
  var s = _sheet(name);
  if (!s) return [];
  return s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
}

function _rows(name) {
  var s = _sheet(name);
  if (!s) return [];
  var lr = s.getLastRow();
  if (lr < 2) return [];
  var vals = s.getRange(2, 1, lr - 1, s.getLastColumn()).getValues();
  var h = _headers(name);
  return vals.map(function (r) {
    var o = {};
    h.forEach(function (c, i) { o[c] = r[i]; });
    return o;
  });
}

function _append(name, obj) {
  var s = _sheet(name);
  var h = _headers(name);
  var row = h.map(function (c) {
    return (obj[c] !== undefined && obj[c] !== null) ? obj[c] : '';
  });
  s.appendRow(row);
  return obj.ID;
}

function _update(name, id, obj) {
  var s = _sheet(name);
  var h = _headers(name);
  var lr = s.getLastRow();
  if (lr < 2) return false;
  var ids = s.getRange(2, 1, lr - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      var existing = s.getRange(i + 2, 1, 1, h.length).getValues()[0];
      var vals = h.map(function (c, j) {
        if (c === 'ID') return String(id); // jaga kolom ID selalu terisi
        if (obj[c] !== undefined && obj[c] !== null) return obj[c];
        return (existing[j] !== undefined && existing[j] !== null) ? existing[j] : '';
      });
      s.getRange(i + 2, 1, 1, h.length).setValues([vals]);
      return true;
    }
  }
  return false;
}

function _deleteByIds(name, ids) {
  var s = _sheet(name);
  if (!s) return;
  var lr = s.getLastRow();
  if (lr < 2) return;
  var heads = _headers(name);
  var all = s.getRange(2, 1, lr - 1, s.getLastColumn()).getValues();
  var idSet = {};
  ids.forEach(function (id) { idSet[String(id)] = true; });
  var keep = all.filter(function (r) { return !idSet[String(r[0])]; });
  keep = keep.map(function (r) {
    var o = r.slice();
    while (o.length < heads.length) o.push('');
    return o;
  });
  if (keep.length > 0) {
    s.getRange(2, 1, keep.length, heads.length).setValues(keep);
  }
  if (keep.length < all.length) {
    s.getRange(2 + keep.length, 1, all.length - keep.length, s.getLastColumn()).clearContent();
  }
}

function _rowById(name, id) {
  var rows = _rows(name);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].ID) === String(id)) return rows[i];
  }
  return null;
}

function _findRow(name, id) {
  var rows = _rows(name);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].ID) === String(id)) return { index: i, row: rows[i] };
  }
  return null;
}

function _uid() { return Utilities.getUuid(); }

function _normToken(x) {
  var s = String(x == null ? '' : x).trim();
  if (!/^\d+$/.test(s)) return s;
  while (s.length < 6) s = '0' + s;
  return s;
}

function _randToken() {
  // 100000..999999 — selalu 6 digit dan tidak pernah diawali angka 0.
  return String(100000 + Math.floor(Math.random() * 900000));
}

function _genToken() {
  var used = {};
  _rows('Pemilih').forEach(function (r) { used[_normToken(r.Token)] = true; });
  for (var i = 0; i < 500; i++) {
    var t = _randToken();
    if (!used[t]) return t;
  }
  throw new Error('Gagal membuat token unik. Coba lagi.');
}

function _genTokenBatch(n) {
  var used = {};
  _rows('Pemilih').forEach(function (r) { used[_normToken(r.Token)] = true; });
  var out = [];
  var guard = 0;
  while (out.length < n && guard < n * 50) {
    var t = _randToken();
    if (!used[t]) { used[t] = true; out.push(t); }
    guard++;
  }
  if (out.length < n) throw new Error('Gagal membuat token unik yang cukup untuk import.');
  return out;
}

/* =====================================================================
 *  PENGATURAN & LOG
 * ===================================================================== */

function _setSetting(k, v) {
  var rows = _rows('Pengaturan');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].Kunci) === String(k)) {
      _update('Pengaturan', rows[i].ID, { Kunci: k, Nilai: v });
      return;
    }
  }
  _append('Pengaturan', { ID: _uid(), Kunci: k, Nilai: v });
}

function _getSetting(k) {
  var rows = _rows('Pengaturan');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].Kunci) === String(k)) return rows[i].Nilai;
  }
  return '';
}

function _settings() {
  var o = {};
  Object.keys(DEFAULTS).forEach(function (k) { o[k] = _getSetting(k); });
  return o;
}

function _log(username, aksi, target, keterangan) {
  _append('Log', { ID: _uid(), Waktu: new Date(), Username: username, Aksi: aksi, Target: target, Keterangan: keterangan || '' });
}

function _fmt(d, withTime) {
  if (!d) return '';
  var dv = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dv.getTime())) return String(d);
  return Utilities.formatDate(dv, Session.getScriptTimeZone(), withTime === false ? 'yyyy-MM-dd' : 'yyyy-MM-dd HH:mm:ss');
}

function _fmtJadwal(d) {
  if (!d) return '';
  var dv = (d instanceof Date) ? d : new Date(d);
  if (isNaN(dv.getTime())) return String(d);
  return Utilities.formatDate(dv, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

function _jadwalText(set) {
  if (set.waktuMulai && set.waktuTutup) {
    return _fmtJadwal(set.waktuMulai) + ' s.d. ' + _fmtJadwal(set.waktuTutup);
  }
  return '';
}

function _voteStatus(set) {
  var now = new Date();
  if (!set.waktuMulai || !set.waktuTutup) return { label: 'Belum diatur', active: false };
  var m = new Date(set.waktuMulai), t = new Date(set.waktuTutup);
  if (isNaN(m.getTime()) || isNaN(t.getTime())) return { label: 'Jadwal tidak valid', active: false };
  if (now < m) return { label: 'Belum dibuka', active: false, buka: _fmtJadwal(m), tutup: _fmtJadwal(t) };
  if (now > t) return { label: 'Ditutup', active: false, buka: _fmtJadwal(m), tutup: _fmtJadwal(t) };
  return { label: 'Dibuka', active: true, tutup: _fmtJadwal(t) };
}

/* =====================================================================
 *  SESI (CACHE)
 * ===================================================================== */

function _createSession(u) {
  var key = 'sess_' + Utilities.getUuid();
  CacheService.getScriptCache().put(key, JSON.stringify(u), 60 * 60 * 6);
  return key;
}

function _getSession(key) {
  if (!key) return null;
  var raw = CacheService.getScriptCache().get(key);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function _destroySession(key) {
  if (key) CacheService.getScriptCache().remove(key);
}

function _requireSession(req, roles) {
  var s = _getSession(req && req.sessionId);
  if (!s) throw new Error('Sesi berakhir. Silakan login kembali.');
  if (roles) {
    var okRole = s.role === 'Admin' || roles.indexOf(s.role) > -1;
    if (!okRole) throw new Error('Anda tidak memiliki akses ke menu ini.');
  }
  return s;
}

// Pengecekan peran KETAT (tanpa pengecualian superuser Admin),
// digunakan untuk halaman yang khusus peran tertentu (Monitor/Hasil = Panitia).
function _requireRoleOnly(req, role) {
  var s = _requireSession(req);
  if (s.role !== role) throw new Error('Halaman ini hanya untuk peran ' + role + '.');
  return s;
}

function _sessionUserName(req) {
  var s = _getSession(req && req.sessionId);
  return s ? s.username : '-';
}

/* =====================================================================
 *  HANDLER HTTP (doGet / doPost)
 *  Frontend (file HTML bernama `index`) dan API dilayani dari satu
 *  URL Web App yang sama:
 *    GET  <url>            -> menampilkan aplikasi (index.html)
 *    GET  <url>?action=.. -> API mode GET (opsional, untuk debugging)
 *    POST <url>            -> API JSON {action, sessionId, payload}
 * ===================================================================== */

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  if (action) {
    var req = { action: action, sessionId: (e.parameter.sessionId || ''), payload: {} };
    if (e.parameter.payload) {
      try { req.payload = JSON.parse(e.parameter.payload); } catch (err) { req.payload = {}; }
    }
    return _jsonOut(_runHandle(req));
  }
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Pemilihan OSIS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var req = {};
  try {
    if (e && e.postData && e.postData.contents) req = JSON.parse(e.postData.contents);
  } catch (err) {
    return _jsonOut({ ok: false, error: 'Format permintaan tidak valid.' });
  }
  return _jsonOut(_runHandle(req));
}

function _runHandle(req) {
  try {
    return handle(req);
  } catch (err) {
    return { ok: false, error: (err && err.message) ? err.message : String(err) };
  }
}

function _jsonOut(out) {
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Jembatan untuk frontend yang di-host DI DALAM Apps Script.
 * index.html memanggil google.script.run.apiBridge(JSON.stringify(req))
 * dan menerima string JSON yang sama bentuknya dengan respons doPost.
 */
function apiBridge(reqJson) {
  var req = {};
  try {
    req = typeof reqJson === 'string' ? JSON.parse(reqJson) : (reqJson || {});
  } catch (err) {
    return JSON.stringify({ ok: false, error: 'Format permintaan tidak valid.' });
  }
  return JSON.stringify(_runHandle(req));
}

function handle(req) {
  var action = req.action || '';
  var p = req.payload || {};
  p.sessionId = req.sessionId || p.sessionId;
  switch (action) {
    case 'getPublicInfo': return getPublicInfo(p);
    case 'login': return login(p);
    case 'loginVoter': return loginVoter(p);
    case 'logout': return logout(p);
    case 'getProfile': return getProfile(p);
    case 'updateProfile': return updateProfile(p);
    case 'getDashboard': return getDashboard(p);
    case 'getLogs': return getLogs(p);
    case 'getAdmins': return getAdmins(p);
    case 'saveAdmin': return saveAdmin(p);
    case 'deleteAdmins': return deleteAdmins(p);
    case 'getKelas': return getKelas(p);
    case 'saveKelas': return saveKelas(p);
    case 'deleteKelas': return deleteKelas(p);
    case 'getPemilih': return getPemilih(p);
    case 'savePemilih': return savePemilih(p);
    case 'generateToken': return generateToken(p);
    case 'deletePemilih': return deletePemilih(p);
    case 'importPemilih': return importPemilih(p);
    case 'getPaslon': return getPaslon(p);
    case 'savePaslon': return savePaslon(p);
    case 'deletePaslon': return deletePaslon(p);
    case 'getVoteInfo': return getVoteInfo(p);
    case 'submitVote': return submitVote(p);
    case 'getMonitor': return getMonitor(p);
    case 'getHasil': return getHasil(p);
    case 'getPengaturan': return getPengaturan(p);
    case 'savePengaturan': return savePengaturan(p);
    case 'getKartuPeserta': return getKartuPeserta(p);
    default: throw new Error('Aksi tidak dikenal: ' + action);
  }
}

function ok(data) { return { ok: true, data: data }; }

/* =====================================================================
 *  AUTENTIKASI
 * ===================================================================== */

function getPublicInfo() {
  var set = _settings();
  _ensurePhotoPublic(set.logo);
  return ok({
    namaSekolah: set.namaSekolah,
    logo: set.logo,
    jadwalText: _jadwalText(set)
  });
}

function login(req) {
  var u = String(req.username || '').trim();
  var p = String(req.password || '');
  if (!u || !p) throw new Error('Username dan password wajib diisi.');
  var rows = _rows('Admin');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].Username).toLowerCase() === u.toLowerCase() && String(rows[i].Password) === p) {
      var key = _createSession({ kind: 'admin', username: rows[i].Username, name: rows[i]['Nama Lengkap'], role: rows[i].Peran, adminId: rows[i].ID });
      _log(rows[i].Username, 'Login', 'Admin/Panitia', 'Login berhasil sebagai ' + rows[i].Peran);
      return ok({ session: key, user: { name: rows[i]['Nama Lengkap'], role: rows[i].Peran, username: rows[i].Username } });
    }
  }
  throw new Error('Username atau password salah.');
}

function loginVoter(req) {
  var t = String(req.token || '').trim();
  if (!/^\d{6}$/.test(t)) throw new Error('Token harus 6 digit angka.');
  var rows = _rows('Pemilih');
  for (var i = 0; i < rows.length; i++) {
    if (_normToken(rows[i].Token) === t) {
      var key = _createSession({ kind: 'voter', username: t, name: rows[i]['Nama Pemilih'], role: 'Pemilih', pemilihId: rows[i].ID, token: t });
      _log(t, 'Login', 'Pemilih', 'Login berhasil');
      return ok({ session: key, user: { name: rows[i]['Nama Pemilih'], role: 'Pemilih', username: t } });
    }
  }
  throw new Error('Token tidak ditemukan.');
}

function logout(req) {
  _destroySession(req.sessionId);
  return ok(true);
}

function getProfile(req) {
  var s = _requireSession(req);
  if (s.kind === 'voter') return ok({ name: s.name, role: s.role, username: s.token });
  return ok({ name: s.name, role: s.role, username: s.username });
}

function updateProfile(req) {
  var s = _requireSession(req);
  var nama = String(req.nama || '').trim();
  if (!nama) throw new Error('Nama tidak boleh kosong.');
  if (s.kind === 'voter') {
    var f = _findRow('Pemilih', s.pemilihId);
    if (!f) throw new Error('Data pemilih tidak ditemukan.');
    _update('Pemilih', s.pemilihId, { 'Nama Pemilih': nama });
    _log(s.token, 'Edit', 'Pemilih', 'Mengubah nama sendiri');
    return ok({ name: nama });
  }
  var f2 = _findRow('Admin', s.adminId);
  if (!f2) throw new Error('Data tidak ditemukan.');
  var upd = { 'Nama Lengkap': nama };
  if (req.password) upd.Password = String(req.password);
  _update('Admin', s.adminId, upd);
  _log(s.username, 'Edit', 'Admin/Panitia', 'Mengubah profil sendiri');
  return ok({ name: nama });
}

/* =====================================================================
 *  DASBOR
 * ===================================================================== */

function getDashboard(req) {
  var s = _requireSession(req);
  var set = _settings();
  _ensurePhotoPublic(set.logo);
  var pemilih = _rows('Pemilih');
  var paslon = _rows('Paslon');
  var vote = _rows('Vote');
  var allLogs = _rows('Log').sort(function (a, b) {
    return new Date(b.Waktu) - new Date(a.Waktu);
  });
  var st = _voteStatus(set);
  return ok({
    namaSekolah: set.namaSekolah,
    logo: set.logo,
    jadwalText: _jadwalText(set),
    jumlahPemilih: pemilih.length,
    jumlahPaslon: paslon.length,
    sudahMemilih: vote.length,
    status: st.label,
    statusActive: st.active,
    role: s.role,
    name: s.name,
    logTotal: allLogs.length,
    log: allLogs.slice(0, 10).map(function (l) {
      return { Waktu: _fmt(l.Waktu), Username: l.Username, Aksi: l.Aksi, Target: l.Target, Keterangan: l.Keterangan };
    })
  });
}

function getLogs(req) {
  _requireSession(req, ['Panitia']);
  var page = Math.max(1, Number(req.page) || 1);
  var pageSize = 10;
  var all = _rows('Log').sort(function (a, b) {
    return new Date(b.Waktu) - new Date(a.Waktu);
  });
  var total = all.length;
  var start = (page - 1) * pageSize;
  return ok({
    total: total,
    page: page,
    pageSize: pageSize,
    rows: all.slice(start, start + pageSize).map(function (l) {
      return { Waktu: _fmt(l.Waktu), Username: l.Username, Aksi: l.Aksi, Target: l.Target, Keterangan: l.Keterangan };
    })
  });
}

/* =====================================================================
 *  DATA MASTER — ADMIN/PANITIA
 * ===================================================================== */

function getAdmins(req) {
  _requireSession(req, ['Admin']);
  return ok(_rows('Admin'));
}

function saveAdmin(req) {
  _requireSession(req, ['Admin']);
  var u = String(req.username || '').trim();
  var p = String(req.password || '').trim();
  var n = String(req.nama || '').trim();
  var r = req.peran;
  if (!u || !p || !n) throw new Error('Semua kolom wajib diisi.');
  if (['Admin', 'Panitia'].indexOf(r) === -1) throw new Error('Peran tidak valid.');
  var rows = _rows('Admin');
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].ID !== req.id && String(rows[i].Username).toLowerCase() === u.toLowerCase()) {
      throw new Error('Username sudah dipakai.');
    }
  }
  if (req.id) {
    if (!_findRow('Admin', req.id)) throw new Error('Data tidak ditemukan.');
    _update('Admin', req.id, { Username: u, Password: p, 'Nama Lengkap': n, Peran: r });
    _log(_sessionUserName(req), 'Edit', 'Admin/Panitia', 'Mengubah akun ' + u);
  } else {
    _append('Admin', { ID: _uid(), Username: u, Password: p, 'Nama Lengkap': n, Peran: r });
    _log(_sessionUserName(req), 'Tambah', 'Admin/Panitia', 'Menambah akun ' + u);
  }
  return ok(true);
}

function deleteAdmins(req) {
  _requireSession(req, ['Admin']);
  var ids = req.ids || [];
  if (ids.length === 0) return ok(true);
  _deleteByIds('Admin', ids);
  _log(_sessionUserName(req), 'Hapus', 'Admin/Panitia', 'Menghapus ' + ids.length + ' akun');
  return ok(true);
}

/* =====================================================================
 *  DATA MASTER — KELAS
 * ===================================================================== */

function getKelas(req) {
  _requireSession(req, ['Admin']);
  return ok(_rows('Kelas'));
}

function saveKelas(req) {
  _requireSession(req, ['Admin']);
  var n = String(req.nama || '').trim();
  if (!n) throw new Error('Nama kelas wajib diisi.');
  var rows = _rows('Kelas');
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].ID !== req.id && String(rows[i]['Nama Kelas']).toLowerCase() === n.toLowerCase()) {
      throw new Error('Nama kelas sudah ada.');
    }
  }
  if (req.id) {
    _update('Kelas', req.id, { 'Nama Kelas': n });
    _log(_sessionUserName(req), 'Edit', 'Kelas', 'Mengubah menjadi ' + n);
  } else {
    _append('Kelas', { ID: _uid(), 'Nama Kelas': n });
    _log(_sessionUserName(req), 'Tambah', 'Kelas', 'Menambah ' + n);
  }
  return ok(true);
}

function deleteKelas(req) {
  _requireSession(req, ['Admin']);
  var ids = req.ids || [];
  if (ids.length === 0) return ok(true);
  var pemilih = _rows('Pemilih');
  var kelasRows = _rows('Kelas');
  var blocked = [];
  ids.forEach(function (id) {
    var nama = null;
    for (var i = 0; i < kelasRows.length; i++) {
      if (String(kelasRows[i].ID) === String(id)) nama = kelasRows[i]['Nama Kelas'];
    }
    var used = pemilih.some(function (p) { return String(p.Kelas) === String(id); });
    if (used && nama) blocked.push(nama);
  });
  if (blocked.length) {
    throw new Error('Kelas "' + blocked.join(', ') + '" masih dipakai data Pemilih. Hapus pemilih terkait terlebih dahulu.');
  }
  _deleteByIds('Kelas', ids);
  _log(_sessionUserName(req), 'Hapus', 'Kelas', 'Menghapus ' + ids.length + ' kelas');
  return ok(true);
}

/* =====================================================================
 *  DATA MASTER — PEMILIH
 * ===================================================================== */

function getPemilih(req) {
  _requireSession(req, ['Admin']);
  return ok(_rows('Pemilih').map(function (r) {
    r.Token = _normToken(r.Token);
    return r;
  }));
}

function savePemilih(req) {
  _requireSession(req, ['Admin']);
  var n = String(req.nama || '').trim();
  var jk = req.jk;
  var kelas = req.kelas;
  if (!n || !jk || !kelas) throw new Error('Nama, jenis kelamin, dan kelas wajib diisi.');
  if (['Laki-laki', 'Perempuan'].indexOf(jk) === -1) throw new Error('Jenis kelamin tidak valid.');
  var token;
  if (req.id) {
    var f = _findRow('Pemilih', req.id);
    if (!f) throw new Error('Data tidak ditemukan.');
    token = f.row.Token;
    _update('Pemilih', req.id, { Token: token, 'Nama Pemilih': n, 'Jenis Kelamin': jk, Kelas: kelas });
    _log(_sessionUserName(req), 'Edit', 'Pemilih', 'Mengubah ' + n);
  } else {
    token = _genToken();
    _append('Pemilih', { ID: _uid(), Token: token, 'Nama Pemilih': n, 'Jenis Kelamin': jk, Kelas: kelas });
    _log(_sessionUserName(req), 'Tambah', 'Pemilih', 'Menambah ' + n + ' (Token: ' + token + ')');
  }
  return ok({ token: _normToken(token) });
}

function generateToken(req) {
  _requireSession(req, ['Admin']);
  var token = _genToken();
  if (req.id) {
    _update('Pemilih', req.id, { Token: _normToken(token) });
    _log(_sessionUserName(req), 'Edit', 'Pemilih', 'Regenerasi token');
  }
  return ok({ token: _normToken(token) });
}

function deletePemilih(req) {
  _requireSession(req, ['Admin']);
  var ids = req.ids || [];
  if (ids.length === 0) return ok(true);
  var pemilih = _rows('Pemilih');
  var voteIds = {};
  _rows('Vote').forEach(function (v) { voteIds[String(v['ID Pemilih'])] = true; });
  var blocked = [];
  ids.forEach(function (id) {
    if (voteIds[String(id)]) {
      var p = _rowById('Pemilih', id);
      if (p) blocked.push(p['Nama Pemilih'] + ' (sudah vote)');
    }
  });
  if (blocked.length) {
    throw new Error('Pemilih ' + blocked.join(', ') + ' sudah memberikan suara dan tidak dapat dihapus.');
  }
  _deleteByIds('Pemilih', ids);
  _log(_sessionUserName(req), 'Hapus', 'Pemilih', 'Menghapus ' + ids.length + ' pemilih');
  return ok(true);
}

function importPemilih(req) {
  _requireSession(req, ['Admin']);
  var rows = req.rows || [];
  if (!rows.length) throw new Error('Tidak ada data untuk diimport.');
  var kelasRows = _rows('Kelas');
  var kelasMap = {};
  kelasRows.forEach(function (k) {
    kelasMap[String(k['Nama Kelas']).toLowerCase()] = k.ID;
  });
  var tokens = _genTokenBatch(rows.length);
  var imported = 0, t = 0;
  var skippedReasons = [];
  rows.forEach(function (r) {
    var nama = String(r.Nama || '').trim();
    var jk = String(r.JenisKelamin || '').trim();
    var kelasId = kelasMap[String(r.Kelas || '').trim().toLowerCase()];
    if (!nama || !jk || !kelasId || ['Laki-laki', 'Perempuan'].indexOf(jk) === -1) {
      skippedReasons.push(nama || '(baris kosong)');
      return;
    }
    _append('Pemilih', { ID: _uid(), Token: tokens[t++], 'Nama Pemilih': nama, 'Jenis Kelamin': jk, Kelas: kelasId });
    imported++;
  });
  var skipped = rows.length - imported;
  _log(_sessionUserName(req), 'Import', 'Pemilih', 'Import CSV: ' + imported + ' berhasil, ' + skipped + ' dilewati');
  return ok({ imported: imported, skipped: skipped, skippedReasons: skippedReasons.slice(0, 20) });
}

/* =====================================================================
 *  DATA MASTER — PASLON
 * ===================================================================== */

function getPaslon(req) {
  _requireSession(req, ['Admin']);
  return ok(_rows('Paslon').map(function (p) {
    _ensurePhotoPublic(p['Foto Ketua']);
    _ensurePhotoPublic(p['Foto Wakil']);
    return p;
  }));
}

function savePaslon(req) {
  _requireSession(req, ['Admin']);
  var no = Number(req.nomor);
  var k = String(req.ketua || '').trim();
  var w = String(req.wakil || '').trim();
  var vm = String(req.visiMisi || '').trim();
  if (!no || !k || !w || !vm) throw new Error('Nomor urut, nama ketua, nama wakil, dan visi misi wajib diisi.');
  var fk = _normFotoId(req.fotoKetua);
  var fw = _normFotoId(req.fotoWakil);
  if (String(req.fotoKetua || '').trim() && !fk) {
    throw new Error('ID file foto ketua tidak valid. Tempel ID file Google Drive (contoh: 1AbCdEfGhIjKlMnOpQrS).');
  }
  if (String(req.fotoWakil || '').trim() && !fw) {
    throw new Error('ID file foto wakil tidak valid. Tempel ID file Google Drive (contoh: 1AbCdEfGhIjKlMnOpQrS).');
  }
  var rows = _rows('Paslon');
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].ID !== req.id && Number(rows[i]['Nomor Urut']) === no) {
      throw new Error('Nomor urut ' + no + ' sudah dipakai paslon lain.');
    }
  }
  var obj = {
    'Nomor Urut': no,
    'Nama Ketua': k,
    'Foto Ketua': fk,
    'Nama Wakil': w,
    'Foto Wakil': fw,
    'Visi Misi': vm
  };
  if (fk) _ensurePhotoPublic(fk);
  if (fw) _ensurePhotoPublic(fw);
  if (req.id) {
    if (!_findRow('Paslon', req.id)) throw new Error('Data tidak ditemukan.');
    _update('Paslon', req.id, obj);
    _log(_sessionUserName(req), 'Edit', 'Paslon', 'Mengubah paslon No. ' + no);
  } else {
    obj.ID = _uid();
    _append('Paslon', obj);
    _log(_sessionUserName(req), 'Tambah', 'Paslon', 'Menambah paslon No. ' + no);
  }
  return ok(true);
}

function deletePaslon(req) {
  _requireSession(req, ['Admin']);
  var ids = req.ids || [];
  if (ids.length === 0) return ok(true);
  var votes = _rows('Vote');
  var blocked = [];
  ids.forEach(function (id) {
    var used = votes.some(function (v) { return String(v['ID Paslon']) === String(id); });
    if (used) {
      var p = _rowById('Paslon', id);
      blocked.push(p ? 'No. ' + p['Nomor Urut'] + ' (' + p['Nama Ketua'] + ')' : id);
    }
  });
  if (blocked.length) {
    throw new Error('Paslon ' + blocked.join(', ') + ' sudah menerima suara dan tidak dapat dihapus.');
  }
  ids.forEach(function (id) {
    var p = _rowById('Paslon', id);
    if (p) {
      // File gambar milik admin di Drive tidak ikut dihapus — aplikasi hanya
      // menyimpan ID-nya. Admin dapat menghapus sendiri dari Drive.
      _deleteByIds('Paslon', [id]);
    }
  });
  _log(_sessionUserName(req), 'Hapus', 'Paslon', 'Menghapus ' + ids.length + ' paslon');
  return ok(true);
}

/* =====================================================================
 *  FOTO / GAMBAR (ID file Google Drive)
 *  Tidak ada upload melalui aplikasi. Admin mengunggah gambar sendiri ke
 *  Google Drive, mengatur akses "Anyone with the link — Viewer", lalu
 *  menempelkan ID file ke form. Database hanya menyimpan ID tersebut.
 * ===================================================================== */

function _isDriveId(v) {
  return !!v && !/^https?:/i.test(String(v));
}

// Ambil ID file bila yang ditempel berupa URL Drive lengkap.
function _extractDriveId(v) {
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  if (_isDriveId(s)) return s;
  var m = s.match(/\/file\/d\/([^/?#]+)/) || s.match(/[?&]id=([^&#]+)/) || s.match(/\/d\/([^/?#]+)/);
  return m ? m[1] : s;
}

// Validasi ringan: ID Drive umumnya 20-60 karakter tanpa spasi/slash.
function _normFotoId(v) {
  var s = _extractDriveId(v);
  if (!s) return '';
  if (/^https?:/i.test(s)) return '';
  if (!/^[A-Za-z0-9_-]{10,80}$/.test(s)) return '';
  return s;
}

// Pastikan file Drive bisa dibaca publik ("Anyone with the link - Viewer")
// agar foto tampil di web. Dipanggil saat dibaca, sehingga file lama pun
// otomatis diperbaiki tanpa perlu meng-share ulang secara manual.
function _ensurePhotoPublic(v) {
  if (!_isDriveId(v)) return;
  try {
    DriveApp.getFileById(String(v)).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) { /* abaikan bila file tidak ditemukan/diizinkan */ }
}

function _deleteDriveFile(v) {
  if (!_isDriveId(v)) return;
  try {
    DriveApp.getFileById(String(v)).setTrashed(true);
  } catch (e) { /* abaikan bila file tidak ditemukan */ }
}

/* =====================================================================
 *  PEMILIHAN (VOTE)
 * ===================================================================== */

function getVoteInfo(req) {
  var s = _requireSession(req, ['Pemilih']);
  var set = _settings();
  var st = _voteStatus(set);
  var already = _findVoteByPemilih(s.pemilihId);
  var data = {
    jadwalText: _jadwalText(set),
    status: st.label,
    statusActive: st.active
  };
  if (already) {
    var p = _rowById('Paslon', already['ID Paslon']);
    data.sudahVote = true;
    data.vote = {
      NomorUrut: p ? p['Nomor Urut'] : '',
      NamaKetua: p ? p['Nama Ketua'] : '',
      NamaWakil: p ? p['Nama Wakil'] : '',
      FotoKetua: p ? p['Foto Ketua'] : '',
      FotoWakil: p ? p['Foto Wakil'] : '',
      Waktu: _fmt(already.Waktu)
    };
  } else {
    data.sudahVote = false;
    data.paslon = _rows('Paslon').sort(function (a, b) {
      return Number(a['Nomor Urut']) - Number(b['Nomor Urut']);
    }).map(function (p) {
      _ensurePhotoPublic(p['Foto Ketua']);
      _ensurePhotoPublic(p['Foto Wakil']);
      return {
        ID: p.ID,
        NomorUrut: p['Nomor Urut'],
        NamaKetua: p['Nama Ketua'],
        NamaWakil: p['Nama Wakil'],
        FotoKetua: p['Foto Ketua'],
        FotoWakil: p['Foto Wakil'],
        VisiMisi: p['Visi Misi']
      };
    });
  }
  return ok(data);
}

function _findVoteByPemilih(pemilihId) {
  if (!pemilihId) return null;
  var rows = _rows('Vote');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['ID Pemilih']) === String(pemilihId)) return rows[i];
  }
  return null;
}

function submitVote(req) {
  var s = _requireSession(req, ['Pemilih']);
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var set = _settings();
    var st = _voteStatus(set);
    if (!st.active) {
      throw new Error(st.label === 'Belum dibuka' ? 'Vote belum dibuka. Tunggu hingga jadwal dimulai.' : 'Vote sudah ditutup.');
    }
    if (_findVoteByPemilih(s.pemilihId)) {
      throw new Error('Anda sudah melakukan vote.');
    }
    var p = _rowById('Paslon', req.paslonId);
    if (!p) throw new Error('Paslon tidak ditemukan.');
    _append('Vote', { ID: _uid(), 'ID Pemilih': s.pemilihId, 'ID Paslon': req.paslonId, Waktu: new Date() });
    _log(s.token, 'Vote', 'Pemilihan', 'Memilih paslon No. ' + p['Nomor Urut']);
    return ok({ paslon: p['Nomor Urut'] + ' - ' + p['Nama Ketua'] });
  } finally {
    lock.releaseLock();
  }
}

/* =====================================================================
 *  MONITOR & HASIL
 * ===================================================================== */

function _buildHasilData() {
  var set = _settings();
  var pemilih = _rows('Pemilih');
  var paslon = _rows('Paslon');
  var votes = _rows('Vote');
  var total = pemilih.length;
  var sudah = votes.length;
  var rekap = paslon.map(function (p) {
    var suara = votes.filter(function (v) { return String(v['ID Paslon']) === String(p.ID); }).length;
    return {
      NomorUrut: p['Nomor Urut'],
      NamaKetua: p['Nama Ketua'],
      NamaWakil: p['Nama Wakil'],
      suara: suara,
      persen: total ? Math.round(suara / total * 1000) / 10 : 0
    };
  }).sort(function (a, b) { return b.suara - a.suara; });
  var kelasRows = _rows('Kelas');
  var perKelas = kelasRows.map(function (k) {
    var ps = pemilih.filter(function (p) { return String(p.Kelas) === String(k.ID); });
    var su = ps.filter(function (p) { return _findVoteByPemilih(p.ID) !== null; }).length;
    return {
      NamaKelas: k['Nama Kelas'],
      total: ps.length,
      sudah: su,
      persen: ps.length ? Math.round(su / ps.length * 1000) / 10 : 0
    };
  }).filter(function (k) { return k.total > 0; });
  return {
    namaSekolah: set.namaSekolah,
    logo: set.logo,
    jadwalText: _jadwalText(set),
    totalPemilih: total,
    sudahMemilih: sudah,
    belumMemilih: total - sudah,
    persenSudah: total ? Math.round(sudah / total * 1000) / 10 : 0,
    persenBelum: total ? Math.round((total - sudah) / total * 1000) / 10 : 0,
    rekap: rekap,
    perKelas: perKelas
  };
}

function getMonitor(req) {
  _requireRoleOnly(req, 'Panitia');
  var data = _buildHasilData();
  data.status = _voteStatus(_settings()).label;
  return ok(data);
}

function getHasil(req) {
  var s = _requireRoleOnly(req, 'Panitia');
  var data = _buildHasilData();
  data.tanggalExport = _fmt(new Date());
  data.generator = { name: s.name, peran: s.role, username: s.username };
  return ok(data);
}

/* =====================================================================
 *  PENGATURAN
 * ===================================================================== */

function getPengaturan(req) {
  _requireSession(req, ['Admin']);
  return ok(_settings());
}

function savePengaturan(req) {
  _requireSession(req, ['Admin']);
  var set = req.data || {};
  Object.keys(DEFAULTS).forEach(function (k) {
    if (set[k] !== undefined) _setSetting(k, String(set[k]).trim());
  });
  _log(_sessionUserName(req), 'Edit', 'Pengaturan', 'Memperbarui pengaturan aplikasi');
  return ok(true);
}

/* =====================================================================
 *  KARTU PESERTA (cetak kartu pemilih)
 * ===================================================================== */

function getKartuPeserta(req) {
  _requireSession(req, ['Admin']);
  var set = _settings();
  _ensurePhotoPublic(set.logo);
  var kelas = _rows('Kelas');
  var kelasMap = {};
  kelas.forEach(function (k) { kelasMap[String(k.ID)] = k['Nama Kelas']; });
  var pemilih = _rows('Pemilih').map(function (p) {
    return {
      ID: p.ID,
      Token: _normToken(p.Token),
      'Nama Pemilih': p['Nama Pemilih'],
      'Jenis Kelamin': p['Jenis Kelamin'],
      Kelas: String(p.Kelas),
      NamaKelas: kelasMap[String(p.Kelas)] || '—'
    };
  }).sort(function (a, b) {
    var ka = a.NamaKelas.toLowerCase(), kb = b.NamaKelas.toLowerCase();
    if (ka !== kb) return ka < kb ? -1 : 1;
    var na = String(a['Nama Pemilih']).toLowerCase(), nb = String(b['Nama Pemilih']).toLowerCase();
    return na < nb ? -1 : na > nb ? 1 : 0;
  });
  return ok({
    namaSekolah: set.namaSekolah,
    logo: set.logo,
    jadwalText: _jadwalText(set),
    kelas: kelas.map(function (k) { return { ID: k.ID, 'Nama Kelas': k['Nama Kelas'] }; })
      .sort(function (a, b) {
        var x = a['Nama Kelas'].toLowerCase(), y = b['Nama Kelas'].toLowerCase();
        return x < y ? -1 : x > y ? 1 : 0;
      }),
    pemilih: pemilih
  });
}
