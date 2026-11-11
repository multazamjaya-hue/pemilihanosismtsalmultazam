/**
 * =====================================================================
 *  PEMILIHAN OSIS — Backend (code.gs)
 *  Dijalankan di Google Apps Script (Google Sheet container).
 *
 *  CARA DEPLOY:
 *  1. Buat project Apps Script baru lalu tempel file ini.
 *  2. Buka Google Sheet terkait (atau buat via Extensions > Apps Script,
 *     file ini otomatis terhubung ke spreadsheet aktif).
 *  3. Jalankan fungsi `CreateDatabase` sekali (via editor atau menu
 *     "Pemilihan OSIS > Buat Database").
 *  4. Deploy > New deployment > Web app:
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     Salin URL Web App, tempel di konstanta SCRIPT_URL pada index.html.
 *
 *  Catatan:
 *  - Password admin disimpan polos sesuai spesifikasi. Untuk produksi
 *    nyata, pertimbangkan hashing.
 *  - Login default awal: username `admin`, password `admin123`.
 *    Segera ubah setelah deploy pertama.
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
  namaSekolah: 'MTS AL MULTAZAM',
  logo: '',
  folderDrive: '',
  waktuMulai: '',
  waktuTutup: ''
};

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
  Object.keys(SCHEMAS).forEach(function (name) {
    var s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    var heads = SCHEMAS[name];
    s.getRange(1, 1, 1, heads.length).setValues([heads]);
    s.getRange(1, 1, 1, heads.length)
      .setFontWeight('bold').setBackground('#4f46e5').setFontColor('#ffffff');
    s.setFrozenRows(1);
  });
  Object.keys(DEFAULTS).forEach(function (k) {
    _setSetting(k, DEFAULTS[k]);
  });
  var pemilihSheet = ss.getSheetByName('Pemilih');
  if (pemilihSheet) {
    // Paksa kolom Token (kolom 2) & ID (kolom 1) berformat TEKS agar
    // angka nol di depan (token seperti 012345) tidak terpotong menjadi 5 digit.
    pemilihSheet.getRange(2, 1, Math.max(1, pemilihSheet.getMaxRows() - 1), 2).setNumberFormat('@');
  }
  if (_rows('Admin').length === 0) {
    _append('Admin', { ID: _uid(), Username: 'admin', Password: 'admin123', 'Nama Lengkap': 'Administrator', Peran: 'Admin' });
  }
  _log('system', 'Setup', 'Database', 'CreateDatabase dijalankan');
  return 'Database berhasil dibuat/diperbarui.';
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
 * ===================================================================== */

function doPost(e) {
  var out = { ok: false, error: 'Terjadi kesalahan.' };
  try {
    var req = {};
    if (e && e.postData && e.postData.contents) {
      req = JSON.parse(e.postData.contents);
    }
    out = handle(req);
  } catch (err) {
    out = { ok: false, error: (err && err.message) ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput(
    JSON.stringify({ ok: true, data: 'Pemilihan OSIS API aktif. Gunakan metode POST.' })
  ).setMimeType(ContentService.MimeType.JSON);
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
    case 'uploadFoto': return uploadFoto(p);
    case 'getVoteInfo': return getVoteInfo(p);
    case 'submitVote': return submitVote(p);
    case 'getMonitor': return getMonitor(p);
    case 'getHasil': return getHasil(p);
    case 'getPengaturan': return getPengaturan(p);
    case 'savePengaturan': return savePengaturan(p);
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
  var rows = _rows('Paslon');
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].ID !== req.id && Number(rows[i]['Nomor Urut']) === no) {
      throw new Error('Nomor urut ' + no + ' sudah dipakai paslon lain.');
    }
  }
  var obj = {
    'Nomor Urut': no,
    'Nama Ketua': k,
    'Foto Ketua': req.fotoKetua || '',
    'Nama Wakil': w,
    'Foto Wakil': req.fotoWakil || '',
    'Visi Misi': vm
  };
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
      _deleteDriveFile(p['Foto Ketua']);
      _deleteDriveFile(p['Foto Wakil']);
    }
  });
  _deleteByIds('Paslon', ids);
  _log(_sessionUserName(req), 'Hapus', 'Paslon', 'Menghapus ' + ids.length + ' paslon');
  return ok(true);
}

/* =====================================================================
 *  UPLOAD FOTO (Google Drive)
 * ===================================================================== */

function _isDriveId(v) {
  return !!v && !/^https?:/i.test(String(v));
}

// Pastikan file Drive bisa dibaca publik ("Anyone with the link - Viewer")
// agar foto tampil di web. Dipanggil saat dibaca, sehingga file lama pun
// otomatis diperbaiki tanpa perlu meng-share folder secara manual.
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

function uploadFoto(req) {
  _requireSession(req, ['Admin']);
  var folderId = String(_getSetting('folderDrive') || '').trim();
  if (!folderId) throw new Error('ID Folder belum diisi oleh Admin');
  var mime = req.mime;
  if (['image/jpeg', 'image/png'].indexOf(mime) === -1) {
    throw new Error('Format foto harus jpg/png/jpeg.');
  }
  var bytes = Utilities.base64Decode(req.data);
  if (bytes.length > 5 * 1024 * 1024) throw new Error('Ukuran foto maksimal 5MB.');
  var folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error('ID Folder Google Drive tidak ditemukan.');
  }
  var name = 'foto_' + Utilities.getUuid() + (mime === 'image/png' ? '.png' : '.jpg');
  var file = folder.createFile(Utilities.newBlob(bytes, mime, name));
  // Agar foto bisa ditampilkan di halaman web (link drive.google.com/uc),
  // setiap foto otomatis di-share "Anyone with the link - Viewer".
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return ok({ fileId: file.getId() });
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
