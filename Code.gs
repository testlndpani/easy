/** Pelacak Kerja Tim L&D — backend Apps Script. Jalankan setupSheets() sekali, lalu Deploy > Web app. */
const SHEET_ID = '1P-XbyN2GW1OMXb2bM6_mLSPL4O4iqLuXRsR621LeZHM';
const SS = () => SpreadsheetApp.openById(SHEET_ID);
const who = () => ACTOR || Session.getActiveUser().getEmail() || 'unknown';

const SCHEMA = {
  Tasks: ['task_id','title','description','status','priority','pic_id','due_date','created_at','completed_at','subtask_total','subtask_done','progress_pct','file_url','updated_at','updated_by','source_ref','start_date','category','co_pic_id','notes'],
  SubTasks: ['subtask_id','task_id','title','is_done','sort_order','done_at'],
  Users: ['user_id','full_name','short_name','position','email','role','is_active'],
  Handover: ['handover_id','employee_id','position','section','supervisor_id','backup_id','start_date','end_date','handover_date','sign_employee_at','sign_backup_at','sign_supervisor_at','created_at','updated_at','updated_by'],
  HO_Items: ['item_id','handover_id','section','sort_order','title','status','deadline','priority','pic','detail','extra1','extra2'],
  Log: ['log_id','timestamp','task_id','action','field','old_value','new_value','actor_email'],
  Link_Penting: ['no','nama','link','keterangan']
};

/** Draft awal tab Link_Penting. Kolom "no" wajib diisi, baris tanpa no tidak dibaca. Link kosong tampil sebagai "Link belum diisi". */
const SEED_LINKS = [
  [1, 'Spreadsheet Pelacak Kerja Tim L&D', 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit', 'Database tugas, sub-tugas, handover, dan log aplikasi ini'],
  [2, 'Daftar Akses Cepat (versi lama)', 'https://docs.google.com/spreadsheets/d/1GUNlLHZYIYIMcnXWTexLE_khmJ9tNWiB5l1ue-QqRiA/edit?gid=0', 'Sumber link sebelumnya, salin isinya ke tab ini lalu hapus baris ini'],
  [3, 'Annual Training Master Plan (ATMP)', '', 'Budget, realisasi, dan biaya per peserta training'],
  [4, 'PANIDOS (Pani Doco System)', '', 'Document control system'],
  [5, 'Portal Asesmen MGR', '', 'Portal asesmen untuk peserta dan trainer'],
  [6, 'Practical Assessment MGR', '', 'Portal penilaian praktik oleh asesor'],
  [7, 'Training Fast', '', 'Budget, proses, dan sertifikasi training Site Pani Gold Project']
];

/** API JSON untuk frontend di GitHub Pages. Tanpa kode akses: siapa pun yang tahu URL /exec bisa memakainya. */
const API = { getData, saveTask, setStatus, toggleSub, addSub, editSub, removeSub, removeTask, saveHandover, signHandover, removeHandover };
let ACTOR = '';
/** Status yang dikenali (sama dengan dropdown di sheet). Nilai lama 'Completed' / 'Not Started' / 'Tugas Baru' tetap dianggap valid. */
const STATUSES = ['Plan', 'In Progress', 'Finalisasi', 'Hold / Reschedule', 'Selesai'];
const isDone = s => s === 'Selesai' || s === 'Completed';
const isNew = s => s === 'Plan' || s === 'Not Started' || s === 'Tugas Baru' || s === '';
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function doGet() { return out({ ok: true, msg: 'API Pelacak Kerja Tim L&D aktif' }); }

function doPost(e) {
  try {
    const r = JSON.parse(e.postData.contents);
    if (!Object.prototype.hasOwnProperty.call(API, r.fn)) return out({ ok: false, error: 'Fungsi tidak dikenal.' });
    ACTOR = String(r.actor || '').slice(0, 60);
    return out({ ok: true, data: API[r.fn](...(r.args || [])) });
  } catch (err) { return out({ ok: false, error: String(err.message || err) }); }
}

/** Jalankan sekali: membuat sheet, header, dropdown, checkbox, dan data PIC awal. */
function setupSheets() {
  const ss = SS();
  const fresh = {};
  Object.entries(SCHEMA).forEach(([name, h]) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    // Aman untuk sheet yang sudah berisi: header lama TIDAK ditimpa, hanya kolom yang belum ada ditambahkan di kanan.
    const cur = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
    const have = cur.filter(Boolean), add = h.filter(k => !have.includes(k));
    if (!have.length) { sh.getRange(1, 1, 1, h.length).setValues([h]); fresh[name] = true; }
    else if (add.length) sh.getRange(1, cur.length + 1, 1, add.length).setValues([add]);
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#fff');
    sh.setFrozenRows(1);
  });
  const lp = ss.getSheetByName('Link_Penting');
  if (lp.getLastRow() < 2) lp.getRange(2, 1, SEED_LINKS.length, 4).setValues(SEED_LINKS);
  // Dropdown dan checkbox hanya dipasang pada sheet yang baru dibuat.
  const dv = l => SpreadsheetApp.newDataValidation().requireValueInList(l, true).build();
  if (fresh.Tasks) {
    const t = ss.getSheetByName('Tasks');
    t.getRange('D2:D').setDataValidation(dv(STATUSES));
    t.getRange('E2:E').setDataValidation(dv(['Tinggi', 'Sedang', 'Rendah']));
  }
  if (fresh.SubTasks) ss.getSheetByName('SubTasks').getRange('D2:D').insertCheckboxes();
  const u = ss.getSheetByName('Users');
  if (u.getLastRow() < 2) u.getRange(2, 1, 4, 7).setValues([
    ['yeni', 'Yeni Ikawati', 'Yeni', 'L&D Supervisor', '', 'admin', true],
    ['sarifudin', 'Sarifudin Mohi', 'Sarifudin', 'Officer Training', '', 'editor', true],
    ['ahmad', 'Ahmad Faqih Imaduddin', 'Ahmad', 'Sr. Trainer', '', 'editor', true],
    ['nfa', 'NFA Staf', 'NFA', 'Admin', '', 'editor', true]]);
}

// ---------- helpers ----------
function T(name) {
  const sh = SS().getSheetByName(name), v = sh.getDataRange().getValues(), h = v.shift(), tz = Session.getScriptTimeZone();
  const fix = x => x instanceof Date ? Utilities.formatDate(x, tz, 'yyyy-MM-dd HH:mm').replace(' 00:00', '') : x;
  return { sh, h, rows: v.filter(r => r[0] !== '').map(r => Object.fromEntries(h.map((k, i) => [k, fix(r[i])]))) };
}
function put(name, key, obj) {
  const t = T(name), i = t.rows.findIndex(r => r[key] == obj[key]);
  const m = Object.assign({}, i < 0 ? {} : t.rows[i], obj);
  const arr = t.h.map(k => m[k] === undefined ? '' : m[k]);
  i < 0 ? t.sh.appendRow(arr) : t.sh.getRange(i + 2, 1, 1, arr.length).setValues([arr]);
}
function log(taskId, action, field, oldV, newV) {
  SS().getSheetByName('Log').appendRow([Utilities.getUuid().slice(0, 8), new Date(), taskId, action, field, oldV, newV, who()]);
}
/** Kunci hanya dipegang saat MENULIS. Pembacaan ulang data (getData, lambat) dilakukan setelah kunci dilepas,
 *  supaya antrean perubahan tidak menumpuk. f boleh mengembalikan objek kecil (mis. {savedId}) yang digabung ke hasil. */
function lock(f) {
  const l = LockService.getScriptLock(); l.waitLock(30000);
  let extra;
  try { extra = f(); } finally { l.releaseLock(); }
  return Object.assign(getData(), extra || {});
}
function nextId(prefix, vals) {
  return prefix + String(Math.max(0, ...vals.map(v => +String(v).replace(/\D/g, '') || 0)) + 1).padStart(2, '0');
}

/** Hitung ulang progres; auto=true menerapkan otomasi 100% => Selesai. */
function recalc(id, auto) {
  const s = T('SubTasks').rows.filter(r => r.task_id === id), total = s.length, done = s.filter(r => r.is_done === true).length;
  const t = T('Tasks').rows.find(r => r.task_id === id); if (!t) return;
  const p = { task_id: id, subtask_total: total, subtask_done: done, progress_pct: total ? Math.round(done * 100 / total) : 0 };
  if (auto && total && done === total && !isDone(t.status)) {
    p.status = 'Selesai'; p.completed_at = new Date(); log(id, 'STATUS_CHANGE', 'status', t.status, 'Selesai (otomatis 100%)');
  } else if (auto && done < total && isDone(t.status)) {
    p.status = 'In Progress'; p.completed_at = ''; log(id, 'STATUS_CHANGE', 'status', t.status, 'In Progress (sub-tugas dibuka)');
  } else if (auto && done > 0 && isNew(t.status)) p.status = 'In Progress';
  put('Tasks', 'task_id', p);
}

// ---------- EMAIL MASUK (sheet "FROM EMAIL") ----------
const EMAIL_SHEET = 'FROM EMAIL', EMAIL_LIMIT = 200;
/** Nama kolom yang dikenali (huruf besar/kecil dan _ - spasi diabaikan). Tambah sinonim di sini bila header sheet berbeda. */
const EMAIL_COLS = {
  dari: ['from', 'sender', 'pengirim', 'dari', 'from email', 'email'],
  subjek: ['subject', 'subjek', 'judul', 'perihal', 'title'],
  tanggal: ['date', 'tanggal', 'received', 'received at', 'waktu', 'timestamp', 'datetime', 'tgl'],
  isi: ['body', 'isi', 'pesan', 'message', 'snippet', 'content', 'plain body', 'ringkasan'],
  link: ['link', 'url', 'permalink', 'thread', 'thread link', 'email link'],
  id: ['id', 'message id', 'msg id', 'thread id']
};
const normKey = s => String(s).toLowerCase().replace(/[_\-\s]+/g, ' ').trim();
/** Kata kunci longgar: header yang MENGANDUNG kata ini ikut dikenali (mis. "Email Body", "Date Received"). */
const EMAIL_FUZZY = {
  tanggal: ['date', 'tanggal', 'received', 'waktu', 'time', 'tgl'],
  isi: ['body', 'isi', 'pesan', 'message', 'snippet', 'content', 'ringkasan', 'detail', 'text'],
  subjek: ['subject', 'subjek', 'perihal', 'judul'],
  link: ['link', 'url', 'permalink'],
  dari: ['from', 'sender', 'pengirim']
};
let EMAIL_META = { headers: [], map: {} };
function emailRows() {
  const sh = SS().getSheets().find(s => normKey(s.getName()) === normKey(EMAIL_SHEET));
  if (!sh || sh.getLastRow() < 2) return [];
  const tz = Session.getScriptTimeZone(), v = sh.getDataRange().getValues(), hraw = v.shift(), h = hraw.map(normKey);
  const strict = {}, col = {}, used = new Set();
  Object.keys(EMAIL_COLS).forEach(k => { strict[k] = h.findIndex(x => EMAIL_COLS[k].includes(x)); col[k] = strict[k]; if (col[k] > -1) used.add(col[k]); });
  // 1) pencocokan longgar pada header
  ['tanggal', 'isi', 'subjek', 'link', 'dari'].forEach(k => {
    if (col[k] > -1) return;
    const i = h.findIndex((x, ix) => !used.has(ix) && EMAIL_FUZZY[k].some(w => x.includes(w)));
    if (i > -1) { col[k] = i; used.add(i); }
  });
  // 2) cadangan berdasarkan isi kolom: tanggal = kolom bernilai Date; isi = kolom teks terpanjang
  const colVals = i => v.map(r => r[i]).filter(x => x !== '' && x != null);
  if (col.tanggal < 0) {
    const i = h.findIndex((x, ix) => { if (used.has(ix)) return false; const a = colVals(ix); return a.length && a.filter(z => z instanceof Date).length >= a.length / 2; });
    if (i > -1) { col.tanggal = i; used.add(i); }
  }
  if (col.isi < 0) {
    let best = -1, bl = 20;
    h.forEach((x, ix) => { if (used.has(ix)) return; const a = colVals(ix).map(String); if (!a.length) return;
      const avg = a.reduce((s, z) => s + z.length, 0) / a.length; if (avg > bl) { bl = avg; best = ix; } });
    if (best > -1) { col.isi = best; used.add(best); }
  }
  EMAIL_META = { headers: hraw.map(String), map: col };
  const fx = x => x instanceof Date ? Utilities.formatDate(x, tz, 'yyyy-MM-dd HH:mm') : (x == null ? '' : String(x));
  const md5 = s => Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('').slice(0, 8);
  return v.map((r, i) => {
    const g = k => col[k] > -1 ? fx(r[col[k]]) : '', gs = k => strict[k] > -1 ? fx(r[strict[k]]) : '';
    const o = { dari: g('dari'), subjek: g('subjek') || r.map(fx).find(x => x.trim()) || '', tanggal: g('tanggal'), isi: g('isi').replace(/\n{3,}/g, '\n\n').trim().slice(0, 3000), link: g('link') };
    // ID tetap dihitung dengan cara lama agar email yang sudah dijadikan tugas tidak kehilangan tautannya.
    const oldSubj = gs('subjek') || r.map(fx).find(x => x.trim()) || '';
    o.id = (gs('id') || 'EM-' + md5([gs('dari'), oldSubj, gs('tanggal')].join('|'))).slice(0, 80);
    o._k = col.tanggal > -1 && r[col.tanggal] instanceof Date ? r[col.tanggal].getTime() : i;
    return o;
  }).filter(o => o.subjek || o.dari).sort((a, b) => b._k - a._k).slice(0, EMAIL_LIMIT).map(o => { delete o._k; return o; });
}
function safeEmails() { try { return emailRows(); } catch (e) { return []; } }

// ---------- API untuk dashboard ----------
function rowsOf(n) { return SS().getSheetByName(n) ? T(n).rows : []; }
function getData() {
  const emails = safeEmails();
  return { tasks: T('Tasks').rows, subs: T('SubTasks').rows, users: T('Users').rows.filter(u => u.is_active !== false),
    hos: rowsOf('Handover'), items: rowsOf('HO_Items'), links: rowsOf('Link_Penting'), emails, emailmeta: EMAIL_META };
}

function saveTask(t) {
  return lock(() => {
    const old = T('Tasks').rows.find(r => r.task_id === t.task_id);
    const subs = String(t.subtasks || '').split('\n').map(s => s.trim()).filter(Boolean);
    delete t.subtasks;
    if (!old) {
      t.task_id = nextId('TSK-', T('Tasks').rows.map(r => r.task_id)); t.created_at = new Date(); t.status = t.status || 'Plan';
      log(t.task_id, 'CREATE', '', '', t.title);
    } else Object.keys(t).forEach(k => { if (String(old[k] == null ? '' : old[k]) !== String(t[k] == null ? '' : t[k])) log(t.task_id, 'UPDATE', k, old[k], t[k]); });
    if (subs.length) {   // sub-tugas baru (saat membuat maupun mengubah tugas)
      const own = T('SubTasks').rows.filter(r => r.task_id === t.task_id);
      const base = Math.max(0, ...own.map(r => +String(r.subtask_id).split('-S')[1] || 0));
      subs.forEach((s, i) => put('SubTasks', 'subtask_id', { subtask_id: t.task_id + '-S' + (base + i + 1), task_id: t.task_id, title: s, is_done: false, sort_order: own.length + i + 1 }));
      if (old) log(t.task_id, 'UPDATE', 'sub-tugas baru', '', subs.join('; '));
    }
    if (isDone(t.status) && !(old && isDone(old.status))) t.completed_at = new Date();
    if (!isDone(t.status)) t.completed_at = '';
    ['due_date', 'start_date'].forEach(k => { if (t[k]) t[k] = new Date(t[k] + 'T00:00:00'); });
    t.updated_at = new Date(); t.updated_by = who();
    put('Tasks', 'task_id', t); recalc(t.task_id, !!old && subs.length > 0);
    return;
  });
}

function setStatus(id, status) {
  return lock(() => {
    const old = T('Tasks').rows.find(r => String(r.task_id).trim() === String(id).trim());
    if (!old) throw new Error('Tugas ' + id + ' tidak ditemukan. Cek header tab Tasks (kolom task_id harus ada), lalu muat ulang halaman.');
    put('Tasks', 'task_id', { task_id: id, status, completed_at: isDone(status) ? new Date() : '', updated_at: new Date(), updated_by: who() });
    log(id, 'STATUS_CHANGE', 'status', old.status, status);
    return;
  });
}

function toggleSub(subId, done) {
  return lock(() => {
    const s = T('SubTasks').rows.find(r => r.subtask_id === subId);
    put('SubTasks', 'subtask_id', { subtask_id: subId, is_done: !!done, done_at: done ? new Date() : '' });
    log(s.task_id, 'SUBTASK_CHECK', s.title, !!s.is_done, !!done);
    recalc(s.task_id, true);
    return;
  });
}

function addSub(taskId, title) {
  return lock(() => {
    const own = T('SubTasks').rows.filter(r => r.task_id === taskId);
    put('SubTasks', 'subtask_id', { subtask_id: taskId + '-S' + (Math.max(0, ...own.map(r => +String(r.subtask_id).split('-S')[1] || 0)) + 1), task_id: taskId, title, is_done: false, sort_order: own.length + 1 });
    log(taskId, 'UPDATE', 'sub-tugas baru', '', title);
    recalc(taskId, true);
    return;
  });
}

function editSub(subId, title) {
  return lock(() => {
    const s = T('SubTasks').rows.find(r => r.subtask_id === subId); if (!s) throw new Error('Sub-tugas tidak ditemukan.');
    put('SubTasks', 'subtask_id', { subtask_id: subId, title });
    log(s.task_id, 'UPDATE', 'sub-tugas', s.title, title);
    return;
  });
}

function removeSub(subId) {
  return lock(() => {
    const t = T('SubTasks'), i = t.rows.findIndex(r => r.subtask_id === subId); if (i < 0) return;
    const tid = t.rows[i].task_id, ttl = t.rows[i].title;
    t.sh.deleteRow(i + 2);
    log(tid, 'UPDATE', 'sub-tugas dihapus', ttl, '');
    recalc(tid, true);
    return;
  });
}

function removeTask(id) {
  return lock(() => {
    ['SubTasks', 'Tasks'].forEach(n => {
      const t = T(n), col = n === 'Tasks' ? 'task_id' : 'task_id';
      for (let i = t.rows.length - 1; i >= 0; i--) if (t.rows[i][col] === id) t.sh.deleteRow(i + 2);
    });
    log(id, 'DELETE', '', '', '');
    return;
  });
}

// ================= LEAVE HANDOVER =================
function saveHandover(h, items) {
  return lock(() => {
    const isNew = !h.handover_id;
    if (!isNew) ['sign_employee_at', 'sign_backup_at', 'sign_supervisor_at', 'created_at'].forEach(k => delete h[k]);
    if (isNew) { h.handover_id = nextId('HO-', rowsOf('Handover').map(r => r.handover_id)); h.created_at = new Date(); }
    ['start_date', 'end_date', 'handover_date'].forEach(k => { if (h[k]) h[k] = new Date(h[k] + 'T00:00:00'); });
    h.updated_at = new Date(); h.updated_by = who();
    put('Handover', 'handover_id', h);
    const t = T('HO_Items');
    for (let i = t.rows.length - 1; i >= 0; i--) if (t.rows[i].handover_id === h.handover_id) t.sh.deleteRow(i + 2);
    const rows = (items || []).map((x, i) => [h.handover_id + '-' + (i + 1), h.handover_id, x.section, i + 1, x.title || '', x.status || '',
      /^\d{4}-\d\d-\d\d$/.test(x.deadline || '') ? new Date(x.deadline + 'T00:00:00') : (x.deadline || ''), x.priority || '', x.pic || '', x.detail || '', x.extra1 || '', x.extra2 || '']);
    if (rows.length) t.sh.getRange(t.sh.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
    log(h.handover_id, isNew ? 'CREATE' : 'UPDATE', 'handover', '', h.employee_id);
    return { savedId: h.handover_id };
  });
}

/** role: employee | backup | supervisor — mencatat waktu konfirmasi (F. Handover Confirmation). */
function signHandover(id, role) {
  return lock(() => {
    put('Handover', 'handover_id', { handover_id: id, ['sign_' + role + '_at']: new Date() });
    log(id, 'SIGN', role, '', who());
    return;
  });
}

function removeHandover(id) {
  return lock(() => {
    ['HO_Items', 'Handover'].forEach(n => {
      const t = T(n);
      for (let i = t.rows.length - 1; i >= 0; i--) if (t.rows[i].handover_id === id) t.sh.deleteRow(i + 2);
    });
    log(id, 'DELETE', 'handover', '', '');
    return;
  });
}