import './styles/main.css';
import './styles/login.css';
import * as XLSX from 'xlsx';
import JsBarcode from 'jsbarcode';
import { Chart, registerables } from 'chart.js';
import { IS_PREVIEW, FIREBASE_READY, firebaseConfig, LOGO, DEF_SET } from './config.js';
import { IC, ic } from './icons.js';
import { setupGtk } from './gtk.js';
Chart.register(...registerables);

/* ============ UTIL ============ */
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => Array.from(r.querySelectorAll(s));
const esc = v => String(v == null ? '' : v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pad = (n, l = 2) => String(n).padStart(l, '0');
const ymd = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const today = () => ymd(new Date());
const addDays = (iso, n) => { const [y, m, d] = iso.split('-').map(Number); return ymd(new Date(y, m - 1, d + n)); };
const diffDays = (a, b) => { const p = a.split('-').map(Number), q = b.split('-').map(Number); return Math.round((new Date(q[0], q[1] - 1, q[2]) - new Date(p[0], p[1] - 1, p[2])) / 86400000); };
const fmtDate = iso => iso ? iso.split('-').reverse().join('/') : '-';
const fmtNum = n => new Intl.NumberFormat('id-ID').format(Number(n) || 0);
const fmtTime = ms => { const d = new Date(ms); return pad(d.getHours()) + ':' + pad(d.getMinutes()); };
const txt = v => String(v == null ? '' : v).trim();
const norm = s => txt(s).toUpperCase();
const natSort = (a, b) => String(a).localeCompare(String(b), 'id', { numeric: true, sensitivity: 'base' });
const levelOf = k => { const s = norm(k).replace(/[^A-Z0-9]/g, ''); if (/^(VIII|8)/.test(s)) return 'VIII'; if (/^(VII|7)/.test(s)) return 'VII'; if (/^(IX|9)/.test(s)) return 'IX'; return 'Lainnya'; };
const hue = s => { let h = 0; for (const c of String(s)) h = (h * 31 + c.charCodeAt(0)) % 360; return h; };
const validKode = c => {
  if (!c) return 'Kode tidak boleh kosong.';
  if (/\s/.test(c)) return 'Kode tidak boleh mengandung spasi.';
  if (/\//.test(c)) return 'Kode tidak boleh mengandung tanda /.';
  if (c === '.' || c === '..' || /^__.*__$/.test(c)) return 'Kode tidak valid.';
  if (c.length > 40) return 'Kode maksimal 40 karakter.';
  return '';
};
let actx;
function beep(ok = true) {
  try {
    actx = actx || new (window.AudioContext || window.webkitAudioContext)();
    const o = actx.createOscillator(), g = actx.createGain();
    o.type = 'sine'; o.frequency.value = ok ? 880 : 220; g.gain.value = 0.05;
    o.connect(g); g.connect(actx.destination); o.start(); o.stop(actx.currentTime + (ok ? 0.09 : 0.25));
  } catch (e) { /* tanpa suara */ }
}
const TT = { ok: ['Berhasil', 'check'], err: ['Gagal', 'alert'], info: ['Info', 'info'], warn: ['Perhatian', 'alert'] };
function toast(msg, type = 'info', ms = 3600) {
  const box = $('#lib-toast'); if (!box) return;
  const m = TT[type] || TT.info; if (type === 'err') ms = Math.max(ms, 6000);
  while (box.children.length >= 4) box.firstChild.remove();
  const el = document.createElement('div');
  el.className = 'toast ' + type; el.setAttribute('role', type === 'err' ? 'alert' : 'status');
  el.innerHTML = `<span class="tic">${ic(m[1], 18)}</span><div class="tbd"><b>${m[0]}</b><span>${esc(msg)}</span></div><button class="tx" type="button" aria-label="Tutup">${ic('close', 15)}</button><i class="tbar" style="animation-duration:${ms}ms"></i>`;
  const kill = () => { el.classList.remove('show'); setTimeout(() => el.remove(), 300); };
  $('.tx', el).addEventListener('click', kill);
  box.appendChild(el);
  requestAnimationFrame(() => el.classList.add('show'));
  setTimeout(kill, ms);
}
const sleep = ms => new Promise(r => setTimeout(r, ms));
/* progres layar penuh (simpan/unggah/cetak) dan bilah tipis (aksi cepat) */
const Prog = {
  t0: 0, id: 0,
  show(title, sub = '', pct = null) {
    this.id++; this.t0 = Date.now();
    $('#lib-prog-t').textContent = title; this.set(pct, sub);
    $('#lib-prog').classList.add('on');
  },
  set(pct, sub) {
    if (sub != null) $('#lib-prog-s').textContent = sub;
    const bar = $('#lib-prog-bar'), i = $('#lib-prog-i');
    if (pct == null) { bar.classList.add('ind'); i.style.width = ''; $('#lib-prog-p').textContent = ''; }
    else { bar.classList.remove('ind'); i.style.width = Math.max(2, Math.min(100, pct)) + '%'; $('#lib-prog-p').textContent = Math.round(pct) + '%'; }
  },
  async hide() {
    const id = this.id; await sleep(Math.max(0, 450 - (Date.now() - this.t0)));
    if (id === this.id) $('#lib-prog').classList.remove('on');
  }
};
async function withProg(title, sub, fn) {
  Prog.show(title, sub); await sleep(40);
  try { return await fn(Prog); } finally { await Prog.hide(); }
}
let topN = 0;
async function quick(fn) {
  topN++; $('#lib-topbar').classList.add('on');
  try { return await fn(); } finally { topN--; if (!topN) $('#lib-topbar').classList.remove('on'); }
}

/* ============ MODAL ============ */
let cfResolve = null, scanner = null;
function openModal(html, opt = {}) {
  const o = $('#lib-modal'), m = $('#lib-mdl');
  m.className = 'mdl' + (opt.wide ? ' wide' : '');
  m.innerHTML = html;
  o.classList.add('on');
  const f = m.querySelector('[autofocus]') || m.querySelector('input:not([type=hidden]):not([type=file]):not([type=checkbox]):not([readonly]),select');
  if (f && !opt.noFocus) setTimeout(() => f.focus(), 60);
  return m;
}
function closeModal() {
  stopScanner();
  $('#lib-modal').classList.remove('on');
  $('#lib-mdl').innerHTML = '';
  if (cfResolve) { const r = cfResolve; cfResolve = null; r(false); }
}
const mHead = t => `<div class="mh"><h3>${esc(t)}</h3><button class="ibtn" data-act="closeModal" aria-label="Tutup">${ic('close')}</button></div>`;
function confirmBox(msg, okText = 'Ya, lanjutkan', danger = false) {
  return new Promise(res => {
    openModal(`${mHead('Konfirmasi')}<div class="mb"><p>${esc(msg)}</p></div><div class="mf"><button class="btn" data-act="cfNo">Batal</button><button class="btn ${danger ? 'dng' : 'pri'}" data-act="cfYes">${esc(okText)}</button></div>`, { noFocus: true });
    cfResolve = res;
  });
}
function stopScanner() {
  if (scanner) { const s = scanner; scanner = null; try { s.stop().then(() => s.clear()).catch(() => { }); } catch (e) { /* abaikan */ } }
}

/* ============ STATE & CACHE ============ */
const S = { set: { ...DEF_SET }, anggota: [], gtk: [], gtkV: [], buku: [], pinjam: [], akt: [], aMap: new Map(), bMap: new Map(), actN: new Map(), actB: new Map(), user: null };
const views = {};
const ACT = {}, INP = {}, CHG = {}, ENT = {}, SUB = {};
const Q = { k: { q: '', kat: '', stok: '', page: 1 }, a: { q: '', kelas: '', page: 1 }, g: { q: '', jenis: '', page: 1 } };
const SEL = { buku: new Set(), anggota: new Set(), gtk: new Set() };
let cur = 'beranda', appOn = false, watchers = [];
const P1 = new Map(), P2 = new Map();

function rebuild() {
  /* GTK dipetakan ke bentuk yang sama dengan siswa (nis = No. Anggota, kelas = Jenis PTK) agar Peminjaman dan Pengembalian tetap sama. */
  S.gtkV = S.gtk.map(g => Object.assign({}, g, { nis: g.noAnggota || g.id, kelas: g.jenisPtk || 'GTK', tipe: 'gtk' }));
  S.aMap = new Map(S.gtkV.concat(S.anggota).map(a => [norm(a.nis || a.id), a]));
  S.bMap = new Map(S.buku.map(b => [norm(b.kodeBuku || b.id), b]));
  S.actN = new Map(); S.actB = new Map();
  for (const p of S.pinjam) {
    if (p.status !== 'dipinjam') continue;
    const kn = norm(p.nis), kb = norm(p.kodeBuku);
    if (!S.actN.has(kn)) S.actN.set(kn, []);
    if (!S.actB.has(kb)) S.actB.set(kb, []);
    S.actN.get(kn).push(p); S.actB.get(kb).push(p);
  }
}
const gtkOf = no => S.gtk.find(g => norm(g.noAnggota || g.id) === norm(no)) || null;
const findAnggota = c => { const k = norm(c); return S.aMap.get(k) || S.anggota.find(a => norm(a.nisn) === k) || null; };
const findBuku = c => S.bMap.get(norm(c)) || null;
const activeOf = nis => S.actN.get(norm(nis)) || [];
const activeOfBook = kode => S.actB.get(norm(kode)) || [];
/* Stok tersedia selalu dihitung dari jumlah eksemplar dikurangi pinjaman aktif, sehingga tidak bisa melebihi total. */
const availOf = b => Math.max(0, (Number(b.jumlahTotal) || 0) - activeOfBook(b.kodeBuku || b.id).length);
function mergePinjam() {
  const m = new Map(P1); for (const [k, v] of P2) m.set(k, v);
  S.pinjam = Array.from(m.values()).sort((a, b) => (b.tsPinjam || 0) - (a.tsPinjam || 0));
}
let rq = 0;
function notify() {
  cancelAnimationFrame(rq);
  rq = requestAnimationFrame(() => {
    rebuild(); applySettings();
    if (appOn && views[cur] && views[cur].update) { try { views[cur].update(); } catch (e) { console.error(e); } }
  });
}

/* ============ PENYIMPANAN DATA (Firebase / Pratinjau) ============ */
let fb = {}, fdb = null, fauth = null;
const mem = { anggota: new Map(), anggota_gtk: new Map(), buku: new Map(), peminjaman: new Map(), aktivitas: new Map(), pengaturan: new Map() };
const newId = c => IS_PREVIEW ? 'p' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8) : fb.doc(fb.collection(fdb, c)).id;
const isInc = v => v && typeof v === 'object' && !Array.isArray(v) && '$inc' in v;

async function fbInit() {
  const m = await import('./firebase.js');
  ({ fb, fdb, fauth } = m.createFirebase(firebaseConfig));
}
function memSync() {
  S.anggota = Array.from(mem.anggota.values());
  S.gtk = Array.from(mem.anggota_gtk.values());
  S.buku = Array.from(mem.buku.values());
  S.pinjam = Array.from(mem.peminjaman.values()).sort((a, b) => (b.tsPinjam || 0) - (a.tsPinjam || 0));
  S.akt = Array.from(mem.aktivitas.values()).sort((a, b) => (b.waktu || 0) - (a.waktu || 0)).slice(0, 20);
  S.set = Object.assign({}, DEF_SET, mem.pengaturan.get('umum') || {});
  notify();
}
async function commit(ops, onProg) {
  if (!ops.length) return;
  const CH = 250;
  if (IS_PREVIEW) {
    for (let oi = 0; oi < ops.length; oi++) {
      const o = ops[oi];
      if (onProg && oi % CH === 0 && oi) { onProg(oi, ops.length); await sleep(0); }
      const m = mem[o.c];
      if (o.t === 'del') { m.delete(o.id); continue; }
      const old = m.get(o.id);
      if (o.t === 'upd' && !old) throw new Error('Data tidak ditemukan: ' + o.id);
      const base = (o.t === 'set' && !o.merge) ? {} : (old || {});
      const nv = Object.assign({}, base);
      for (const k in o.d) { const v = o.d[k]; nv[k] = isInc(v) ? (Number(base[k]) || 0) + v.$inc : v; }
      nv.id = o.id; m.set(o.id, nv);
    }
    if (onProg) onProg(ops.length, ops.length);
    memSync(); return;
  }
  const conv = d => { const r = {}; for (const k in d) { const v = d[k]; if (v === undefined) continue; r[k] = isInc(v) ? fb.increment(v.$inc) : v; } return r; };
  for (let i = 0; i < ops.length; i += CH) {
    if (onProg) onProg(i, ops.length);
    const b = fb.writeBatch(fdb);
    for (const o of ops.slice(i, i + CH)) {
      const ref = fb.doc(fdb, o.c, o.id);
      if (o.t === 'del') b.delete(ref);
      else if (o.t === 'upd') b.update(ref, conv(o.d));
      else b.set(ref, conv(o.d), { merge: !!o.merge });
    }
    await b.commit();
    if (onProg) onProg(Math.min(i + CH, ops.length), ops.length);
  }
}
async function isEmpty(c) {
  if (IS_PREVIEW) return mem[c].size === 0;
  const s = await fb.getDocs(fb.query(fb.collection(fdb, c), fb.limit(1)));
  return s.empty;
}
let deniedOn = false;
function showDenied() {
  if (deniedOn) return; deniedOn = true;
  const d = document.createElement('div'); d.id = 'lib-deny'; d.setAttribute('role', 'alert');
  d.innerHTML = `<b>Akses ditolak oleh Firestore Rules</b><p>Login berhasil, tetapi Firestore menolak pembacaan data untuk akun <b>${esc(S.user || '')}</b>.</p><p>Buka Firebase Console &gt; Firestore Database &gt; Rules. Pada fungsi <b>isAdmin()</b>, isi daftar email dengan email di atas (huruf harus sama persis), klik <b>Publish</b>, lalu muat ulang halaman ini.</p><div><button class="btn" data-act="denyClose">Tutup</button><button class="btn pri" data-act="logout">Keluar</button></div>`;
  $('#lib-root').appendChild(d);
}
ACT.denyClose = () => { const d = $('#lib-deny'); if (d) d.remove(); };
function watchAll() {
  const unsub = [], ps = [];
  const listen = (key, q, onData) => {
    ps.push(new Promise(res => {
      let done = false;
      unsub.push(fb.onSnapshot(q, snap => { onData(snap); if (!done) { done = true; res(); } notify(); }, err => {
        console.error(err);
        if (err.code === 'permission-denied') showDenied(); else toast('Gagal memuat data ' + key + ': ' + (err.code || err.message), 'err', 6500);
        if (!done) { done = true; res(); }
      }));
    }));
  };
  const rows = s => s.docs.map(d => Object.assign({ id: d.id }, d.data()));
  listen('anggota', fb.collection(fdb, 'anggota'), s => { S.anggota = rows(s); });
  listen('anggota GTK', fb.collection(fdb, 'anggota_gtk'), s => { S.gtk = rows(s); });
  listen('buku', fb.collection(fdb, 'buku'), s => { S.buku = rows(s); });
  listen('peminjaman', fb.query(fb.collection(fdb, 'peminjaman'), fb.orderBy('tsPinjam', 'desc'), fb.limit(5000)), s => { P1.clear(); rows(s).forEach(r => P1.set(r.id, r)); mergePinjam(); });
  listen('peminjaman aktif', fb.query(fb.collection(fdb, 'peminjaman'), fb.where('status', '==', 'dipinjam')), s => { P2.clear(); rows(s).forEach(r => P2.set(r.id, r)); mergePinjam(); });
  listen('aktivitas', fb.query(fb.collection(fdb, 'aktivitas'), fb.orderBy('waktu', 'desc'), fb.limit(20)), s => { S.akt = rows(s); });
  listen('pengaturan', fb.doc(fdb, 'pengaturan', 'umum'), s => { S.set = Object.assign({}, DEF_SET, s.exists() ? s.data() : {}); });
  watchers = unsub;
  return Promise.all(ps);
}
function stopWatch() { watchers.forEach(f => { try { f(); } catch (e) { /* abaikan */ } }); watchers = []; P1.clear(); P2.clear(); }
const actOp = (jenis, nama, detail) => ({ t: 'set', c: 'aktivitas', id: newId('aktivitas'), d: { jenis, nama, detail, waktu: Date.now() } });

/* ============ SETUP DATABASE OTOMATIS ============ */
const SEED_SISWA = [
  ['Rina Agustina', 'VII-A', 'P'], ['Budi Santoso', 'VIII-1', 'L'], ['Siti Nurhaliza', 'VIII-2', 'P'], ['Andi Pratama', 'IX-A', 'L'],
  ['Dewi Lestari', 'VII-B', 'P'], ['Fajar Ramadhan', 'VII-A', 'L'], ['Putri Maharani', 'VIII-1', 'P'], ['Rizky Firmansyah', 'IX-B', 'L'],
  ['Nabila Azzahra', 'IX-A', 'P'], ['Dimas Saputra', 'VIII-2', 'L'], ['Anisa Rahmawati', 'VII-B', 'P'], ['Raka Wijaya', 'IX-B', 'L'],
  ['Salsabila Putri', 'VIII-1', 'P'], ['Yoga Permana', 'VII-A', 'L']
];
const SEED_BUKU = [
  ['BK-0001', 'Laskar Pelangi', 'Andrea Hirata', 'Bentang Pustaka', '2005', 'Sastra', 'R-05', 3],
  ['BK-0002', 'Bumi Manusia', 'Pramoedya Ananta Toer', 'Hasta Mitra', '1980', 'Sastra', 'R-05', 2],
  ['BK-0003', 'Atomic Habits', 'James Clear', 'Gramedia', '2019', 'Lainnya', 'R-09', 2],
  ['BK-0004', 'Filosofi Teras', 'Henry Manampiring', 'Kompas', '2018', 'Lainnya', 'R-09', 2],
  ['BK-0005', 'Matematika SMP Kelas VII', 'Tim Kemendikbud', 'Kemendikbud', '2022', 'Matematika', 'R-02', 5],
  ['BK-0006', 'Matematika SMP Kelas VIII', 'Tim Kemendikbud', 'Kemendikbud', '2022', 'Matematika', 'R-02', 5],
  ['BK-0007', 'IPA Terpadu Kelas VII', 'Tim Kemendikbud', 'Kemendikbud', '2022', 'IPA', 'R-03', 4],
  ['BK-0008', 'IPA Terpadu Kelas IX', 'Tim Kemendikbud', 'Kemendikbud', '2022', 'IPA', 'R-03', 4],
  ['BK-0009', 'IPS Terpadu Kelas VIII', 'Tim Kemendikbud', 'Kemendikbud', '2022', 'IPS', 'R-04', 4],
  ['BK-0010', 'Pendidikan Agama Islam dan Budi Pekerti VII', 'Tim Kemendikbud', 'Kemendikbud', '2021', 'Agama', 'R-01', 4],
  ['BK-0011', 'Bahasa Indonesia Kelas VIII', 'Tim Kemendikbud', 'Kemendikbud', '2022', 'Bahasa', 'R-06', 4],
  ['BK-0012', 'English for Nusantara Kelas IX', 'Tim Kemendikbud', 'Kemendikbud', '2022', 'Bahasa', 'R-06', 3],
  ['BK-0013', 'Negeri 5 Menara', 'Ahmad Fuadi', 'Gramedia', '2009', 'Sastra', 'R-05', 2],
  ['BK-0014', 'Dilan 1990', 'Pidi Baiq', 'Pastel Books', '2014', 'Sastra', 'R-05', 2],
  ['BK-0015', 'Sejarah Indonesia Kelas IX', 'Tim Kemendikbud', 'Kemendikbud', '2022', 'IPS', 'R-04', 3],
  ['BK-0016', 'Kisah Para Nabi', 'Ibnu Katsir', 'Pustaka Azzam', '2010', 'Agama', 'R-01', 2]
];
function seedData() {
  const now = Date.now(), t0 = today(), lama = DEF_SET.lamaPinjamHari;
  const siswa = SEED_SISWA.map((s, i) => ({ nis: String(260101 + i), nisn: '00' + String(61234501 + i), nama: s[0], kelas: s[1], jenisKelamin: s[2], aktif: true, contoh: true, dibuatPada: now - (i % 4 === 0 ? 6 : 80) * 86400000 }));
  const buku = SEED_BUKU.map((b, i) => ({ kodeBuku: b[0], judul: b[1], pengarang: b[2], penerbit: b[3], tahun: b[4], kategori: b[5], rak: b[6], jumlahTotal: b[7], jumlahTersedia: b[7], contoh: true, dibuatPada: now - (i % 3 === 0 ? 5 : 90) * 86400000 }));
  const pat = [0, 0, 2, 1, 0, 3, 4, 2, 0, 5, 1, 6, 12, 0, 13, 7, 2, 8, 9, 3, 10, 11, 0, 14, 15, 1, 2, 4, 0, 5];
  const perDay = [4, 5, 5, 6, 5, 7, 6];
  const act = {}, aktifPair = new Set(), pinjam = [];
  let seq = 0;
  const mk = (a, b, tglPinjam, returned, tglKembali, minute) => {
    const [y, m, d] = tglPinjam.split('-').map(Number);
    const ts = new Date(y, m - 1, d, 7, 30).getTime() + minute * 60000;
    pinjam.push({ nis: a.nis, namaSiswa: a.nama, kelas: a.kelas, kodeBuku: b.kodeBuku, judulBuku: b.judul, tglPinjam, tsPinjam: ts, jatuhTempo: addDays(tglPinjam, lama), tglKembali: returned ? tglKembali : '', tsKembali: returned ? ts + 86400000 : 0, status: returned ? 'kembali' : 'dipinjam', denda: 0, terlambat: 0, contoh: true });
    if (!returned) act[b.kodeBuku] = (act[b.kodeBuku] || 0) + 1;
  };
  for (let d = 6; d >= 0; d--) {
    for (let k = 0; k < perDay[6 - d]; k++) {
      seq++;
      const a = siswa[(seq * 3) % siswa.length], b = buku[pat[seq % pat.length]];
      const tgl = addDays(t0, -d);
      const pair = a.nis + '|' + b.kodeBuku;
      let returned = !(d < 2 && seq % 3 !== 0);
      if (!returned && ((act[b.kodeBuku] || 0) >= b.jumlahTotal || aktifPair.has(pair))) returned = true;
      if (!returned) aktifPair.add(pair);
      mk(a, b, tgl, returned, addDays(tgl, Math.min(d, 1 + (seq % 3))), seq * 13);
    }
  }
  const lateB = buku[8], lateA = siswa[5];
  mk(lateA, lateB, addDays(t0, -10), false, '', 5);
  buku.forEach(b => { b.jumlahTersedia = b.jumlahTotal - (act[b.kodeBuku] || 0); });
  const akt = pinjam.slice().sort((a, b) => b.tsPinjam - a.tsPinjam).slice(0, 6).map((p, i) => ({
    jenis: i % 3 === 1 ? 'kembali' : 'pinjam', nama: p.namaSiswa, detail: p.judulBuku, waktu: now - (i + 1) * 47 * 60000, contoh: true
  }));
  akt.push({ jenis: 'anggota', nama: siswa[2].nama, detail: 'Kelas ' + siswa[2].kelas, waktu: now - 7 * 3600000, contoh: true });
  return { siswa, buku, pinjam, akt };
}
const SEED_GTK = [
  { noAnggota: 'G001', nip: '198001012005012001', nama: 'Contoh Guru', jenisPtk: 'Guru', jenisKelamin: 'P' },
  { noAnggota: 'T001', nip: '', nama: 'Contoh Tendik', jenisPtk: 'Tendik', jenisKelamin: 'L' }
];
/* Mode produksi hanya membuat dokumen pengaturan. Data contoh hanya ada di memori pada mode pratinjau. */
async function autoSetupDatabase() {
  const ops = [];
  if (await isEmpty('pengaturan')) ops.push({ t: 'set', c: 'pengaturan', id: 'umum', d: Object.assign({}, DEF_SET, { dibuatPada: Date.now() }) });
  if (IS_PREVIEW) {
    const sd = seedData();
    sd.siswa.forEach(s => ops.push({ t: 'set', c: 'anggota', id: s.nis, d: s }));
    SEED_GTK.forEach(g => ops.push({ t: 'set', c: 'anggota_gtk', id: g.noAnggota, d: Object.assign({ aktif: true, contoh: true, dibuatPada: Date.now() }, g) }));
    sd.buku.forEach(b => ops.push({ t: 'set', c: 'buku', id: b.kodeBuku, d: b }));
    sd.pinjam.forEach(p => ops.push({ t: 'set', c: 'peminjaman', id: newId('peminjaman'), d: p }));
    sd.akt.forEach(a => ops.push({ t: 'set', c: 'aktivitas', id: newId('aktivitas'), d: a }));
  }
  await commit(ops);
}
/* ============ TAMPILAN UMUM ============ */
function applySettings() {
  const s = S.set;
  document.title = s.namaAplikasi + ' - ' + s.namaSekolah;
  const map = {
    app: s.namaAplikasi, sekolah: s.namaSekolah, sekolahUp: String(s.namaSekolah).toUpperCase(), semboyan: s.semboyan,
    sekolahTa: s.namaSekolah + ' \u2022 TA ' + s.tahunAjaran, copy: '\u00A9 ' + new Date().getFullYear() + ' ' + s.namaAplikasi + ' ' + s.namaSekolah
  };
  $$('[data-t]').forEach(el => { const v = map[el.dataset.t]; if (v != null && el.textContent !== v) el.textContent = v; });
  $$('img[data-logo]').forEach(i => {
    const want = s.logoUrl || LOGO;
    if (i.getAttribute('src') !== want) i.setAttribute('src', want);
    i.onerror = () => { if (i.getAttribute('src') !== LOGO) i.setAttribute('src', LOGO); else i.style.visibility = 'hidden'; };
  });
}
function hydrateIcons(root = document) {
  $$('[data-i]', root).forEach(el => { if (!el.firstChild) el.innerHTML = ic(el.dataset.i, +el.dataset.s || 20); });
}
const NAV = [['beranda', 'Beranda', 'home'], ['katalog', 'Katalog Buku', 'book'], ['peminjaman', 'Peminjaman', 'borrow'], ['pengembalian', 'Pengembalian', 'return'], ['anggota', 'Anggota Siswa', 'users'], ['gtk', 'Anggota GTK', 'briefcase'], ['laporan', 'Laporan', 'chart'], ['pengaturan', 'Pengaturan', 'gear']];
function renderNav() {
  $('#lib-nav').innerHTML = NAV.map(([id, l, i]) => `<button class="nv${cur === id ? ' on' : ''}" data-act="go" data-v="${id}" title="${l}">${ic(i, 22)}<span>${l}</span></button>`).join('');
  const B = [['beranda', 'Beranda', 'home'], ['katalog', 'Katalog', 'book'], ['peminjaman', 'Pinjam', 'borrow'], ['pengembalian', 'Kembali', 'return']];
  const more = ['anggota', 'gtk', 'laporan', 'pengaturan'].includes(cur);
  $('#lib-bnav').innerHTML = B.map(([id, l, i]) => `<button class="${cur === id ? 'on' : ''}" data-act="go" data-v="${id}">${ic(i, 22)}<span>${l}</span></button>`).join('') +
    `<button class="${more ? 'on' : ''}" data-act="more">${ic('dots', 22)}<span>Menu</span></button>`;
}
function go(id, fromHash) {
  if (!views[id]) id = 'beranda';
  if (views[cur] && views[cur].unmount) views[cur].unmount();
  cur = id;
  closeModal();
  const el = $('#lib-view');
  el.innerHTML = '';
  views[id].mount(el);
  el.insertAdjacentHTML('beforeend', '<div class="mfoot"><span data-t="copy"></span></div>');
  applySettings();
  hydrateIcons(el);
  try { views[id].update(); } catch (e) { console.error(e); }
  el.scrollTop = 0;
  renderNav();
  if (!fromHash) { try { history.replaceState(null, '', '#/' + id); } catch (e) { location.hash = '/' + id; } }
}
function pagerHtml(total, page, per) {
  const pages = Math.max(1, Math.ceil(total / per));
  const a = total ? (page - 1) * per + 1 : 0, b = Math.min(total, page * per);
  return `<div class="pager"><span>Menampilkan ${fmtNum(a)}\u2013${fmtNum(b)} dari ${fmtNum(total)}</span><div><button class="btn sm" data-act="pg" data-v="-1" ${page <= 1 ? 'disabled' : ''}>Sebelumnya</button><span>Hal. ${page}/${pages}</span><button class="btn sm" data-act="pg" data-v="1" ${page >= pages ? 'disabled' : ''}>Berikutnya</button></div></div>`;
}
const emptyHtml = (icon, msg) => `<div class="empty">${ic(icon, 40)}<p>${esc(msg)}</p></div>`;
const kelasList = () => Array.from(new Set(S.anggota.map(a => txt(a.kelas)).filter(Boolean))).sort(natSort);
const kategoriPengaturan = () => (S.set.kategori && S.set.kategori.length ? S.set.kategori : DEF_SET.kategori).slice();
const kategoriList = () => { const base = (S.set.kategori && S.set.kategori.length ? S.set.kategori : DEF_SET.kategori).slice(); S.buku.forEach(b => { const k = txt(b.kategori); if (k && !base.includes(k)) base.push(k); }); return base; };

/* ============ BERANDA ============ */
const BANNER_ART = `<svg class="bn-art" viewBox="0 0 560 230" preserveAspectRatio="xMaxYMax meet" aria-hidden="true">
<defs>
<linearGradient id="lb-pg" x1="0" x2="1"><stop offset="0" stop-color="#ffffff"/><stop offset="1" stop-color="#d6e6ff"/></linearGradient>
<linearGradient id="lb-pg2" x1="1" x2="0"><stop offset="0" stop-color="#ffffff"/><stop offset="1" stop-color="#cfe0ff"/></linearGradient>
<radialGradient id="lb-gl" cx=".5" cy=".5" r=".5"><stop offset="0" stop-color="#4da3ff" stop-opacity=".6"/><stop offset="1" stop-color="#4da3ff" stop-opacity="0"/></radialGradient>
<linearGradient id="lb-y" x1="0" x2="0" y1="0" y2="1"><stop offset="0" stop-color="#fed403"/><stop offset="1" stop-color="#ffb305"/></linearGradient>
</defs>
<ellipse cx="300" cy="175" rx="250" ry="62" fill="url(#lb-gl)"/>
<g opacity=".5">
<rect x="20" y="40" width="520" height="3" fill="#7fb2ff"/><rect x="20" y="92" width="520" height="3" fill="#7fb2ff"/>
<rect x="30" y="8" width="14" height="32" rx="2" fill="#1d6be0"/><rect x="46" y="14" width="12" height="26" rx="2" fill="#3b82f6"/><rect x="60" y="6" width="16" height="34" rx="2" fill="#0b3a9c"/><rect x="78" y="12" width="12" height="28" rx="2" fill="url(#lb-y)"/><rect x="92" y="9" width="15" height="31" rx="2" fill="#2563eb"/><rect x="109" y="16" width="12" height="24" rx="2" fill="#60a5fa"/><rect x="123" y="7" width="16" height="33" rx="2" fill="#1e40af"/>
<rect x="300" y="9" width="14" height="31" rx="2" fill="#3b82f6"/><rect x="316" y="5" width="16" height="35" rx="2" fill="#0b3a9c"/><rect x="334" y="13" width="12" height="27" rx="2" fill="url(#lb-y)"/><rect x="348" y="8" width="15" height="32" rx="2" fill="#2563eb"/><rect x="365" y="15" width="12" height="25" rx="2" fill="#60a5fa"/><rect x="379" y="6" width="16" height="34" rx="2" fill="#1e40af"/><rect x="397" y="11" width="13" height="29" rx="2" fill="#3b82f6"/>
<rect x="60" y="60" width="15" height="32" rx="2" fill="#2563eb"/><rect x="77" y="55" width="13" height="37" rx="2" fill="#60a5fa"/><rect x="92" y="62" width="16" height="30" rx="2" fill="url(#lb-y)"/><rect x="110" y="57" width="12" height="35" rx="2" fill="#1e40af"/><rect x="124" y="63" width="14" height="29" rx="2" fill="#3b82f6"/>
<rect x="400" y="58" width="14" height="34" rx="2" fill="#1d6be0"/><rect x="416" y="63" width="12" height="29" rx="2" fill="#60a5fa"/><rect x="430" y="55" width="16" height="37" rx="2" fill="#0b3a9c"/><rect x="448" y="61" width="13" height="31" rx="2" fill="url(#lb-y)"/><rect x="463" y="57" width="15" height="35" rx="2" fill="#2563eb"/>
</g>
<g>
<path d="M300 205c-52-30-112-34-160-18l14-72c46-14 104-10 146 16z" fill="url(#lb-pg)"/>
<path d="M300 205c52-30 112-34 160-18l-14-72c-46-14-104-10-146 16z" fill="url(#lb-pg2)"/>
<path d="M300 133v72" fill="none" stroke="#9db8e8" stroke-width="2"/>
<g fill="none" stroke="#9fb7e6" stroke-width="2.2" stroke-linecap="round" opacity=".85">
<path d="M170 135c32-8 72-4 112 14M168 150c32-8 72-4 114 14M166 165c32-8 72-4 116 14"/>
<path d="M430 135c-32-8-72-4-112 14M432 150c-32-8-72-4-114 14M434 165c-32-8-72-4-116 14"/>
</g>
<path d="M140 187l14-72" fill="none" stroke="#c4d7fa" stroke-width="3"/><path d="M460 187l-14-72" fill="none" stroke="#c4d7fa" stroke-width="3"/>
<ellipse cx="300" cy="212" rx="170" ry="9" fill="#001455" opacity=".35"/>
</g>
</svg>`;
let ch = {}, PAL = ['#1d6fe8', '#8b5cf6', '#14b8a6', '#f59e0b', '#84cc16', '#ec4899', '#64748b', '#0ea5e9'];
function destroyCharts() { Object.keys(ch).forEach(k => { try { ch[k].destroy(); } catch (e) { /* abaikan */ } }); ch = {}; }
const AKT = { pinjam: ['#2f6fe0', 'Peminjaman buku', 'book'], kembali: ['#f59e0b', 'Pengembalian buku', 'return'], anggota: ['#14b8a6', 'Anggota baru', 'users'], buku: ['#8b5cf6', 'Buku baru', 'book'] };
views.beranda = {
  mount(el) {
    destroyCharts();
    el.innerHTML = `
<section class="banner">
  ${BANNER_ART}
  <div class="bn-txt">
    <h1>Membaca Membuka Jendela<br><em>Dunia</em></h1>
    <p>Perpustakaan adalah jendela dunia. Mari tingkatkan literasi untuk masa depan yang lebih baik.</p>
    <button class="btn warn sm" data-act="go" data-v="katalog">${ic('book', 16)}Jelajahi Katalog Buku ${ic('arrow', 14)}</button>
  </div>
  <div class="bn-pts">
    <div>${ic('book', 34)}<span>Baca<br>Lebih Banyak</span></div>
    <div>${ic('bulb', 34)}<span>Tambah<br>Pengetahuan</span></div>
    <div>${ic('target', 34)}<span>Raih<br>Masa Depan</span></div>
  </div>
</section>
<section class="stats" id="d-stats"></section>
<section class="dgrid">
  <div class="card pn a-bar">
    <div class="ph"><h3>${ic('chart', 18)} Grafik Peminjaman Buku</h3><select class="sel" data-ch="chartDays" aria-label="Rentang grafik"><option value="7">7 Hari Terakhir</option><option value="14">14 Hari Terakhir</option><option value="30">30 Hari Terakhir</option></select></div>
    <div class="cbox"><canvas id="c-bar" aria-label="Grafik peminjaman"></canvas></div>
  </div>
  <div class="card pn a-cat">
    <div class="ph"><h3>${ic('book', 18)} Kategori Buku</h3></div>
    <div class="donut"><div class="dwrap"><canvas id="c-cat" aria-label="Kategori buku"></canvas><div class="dmid" id="d-mid"></div></div><ul class="leg" id="d-leg"></ul></div>
  </div>
  <div class="card pn a-act">
    <div class="ph"><h3>${ic('clock', 18)} Aktivitas Terbaru</h3><a data-act="go" data-v="laporan">Lihat Semua ${ic('arrow', 13)}</a></div>
    <ul class="alist" id="d-act"></ul>
  </div>
  <div class="card pn a-pop">
    <div class="ph"><h3>${ic('flame', 18)} Buku Terpopuler</h3><a data-act="popAll">Lihat Semua ${ic('arrow', 13)}</a></div>
    <div class="pops" id="d-pop"></div>
  </div>
  <div class="card pn a-sis">
    <div class="ph"><h3>${ic('users', 18)} Siswa Paling Sering Meminjam</h3><a data-act="sisAll">Lihat Semua ${ic('arrow', 13)}</a></div>
    <p class="sub" style="margin:-6px 0 8px">6 bulan terakhir</p>
    <ul class="alist" id="d-sis"></ul>
  </div>
  <div class="card pn a-kls">
    <div class="ph"><h3>${ic('users', 18)} Peminjaman per Kelas</h3><span class="sub" style="margin:0">30 hari</span></div>
    <div id="d-kls" class="klsgrid"></div>
  </div>
</section>`;
    const sel = $('[data-ch=chartDays]', el); if (sel) sel.value = String(S.chartDays || 7);
  },
  unmount() { destroyCharts(); },
  update() {
    const t = today(), y = addDays(t, -1), now = Date.now(), d30 = now - 30 * 86400000;
    /* Dihitung per data buku di katalog (1 baris = 1 buku), sama dengan jumlah baris di Katalog Buku. */
    let total = 0, avail = 0, newB = 0;
    S.buku.forEach(b => { total++; if (!activeOfBook(b.kodeBuku || b.id).length) avail++; if ((b.dibuatPada || 0) >= d30) newB++; });
    const newA = S.anggota.concat(S.gtk).filter(a => (a.dibuatPada || 0) >= d30).length;
    const aktif = S.anggota.concat(S.gtk).filter(a => a.aktif !== false).length;
    let pt = 0, py = 0, kt = 0, ky = 0;
    for (const p of S.pinjam) { if (p.tglPinjam === t) pt++; else if (p.tglPinjam === y) py++; if (p.tglKembali === t) kt++; else if (p.tglKembali === y) ky++; }
    const dl = (n, txtAfter) => n === 0 ? `<span class="dl">${txtAfter}</span>` : `<div class="dl"><b>${n > 0 ? '+' : ''}${fmtNum(n)}</b> ${txtAfter}</div>`;
    const cards = [
      ['t1', '#1d6fe8', 'book', 'Total Buku', fmtNum(total), `<div class="dl"><b>+${fmtNum(newB)}</b> dari bulan lalu</div>`],
      ['t2', '#16a34a', 'users', 'Anggota Aktif', fmtNum(aktif), `<div class="dl"><b>+${fmtNum(newA)}</b> dari bulan lalu</div>`],
      ['t3', '#8b5cf6', 'borrow', 'Peminjaman Hari Ini', fmtNum(pt), `<div class="dl"><b>${pt - py >= 0 ? '+' : ''}${fmtNum(pt - py)}</b> dari kemarin</div>`],
      ['t4', '#f59e0b', 'return', 'Pengembalian Hari Ini', fmtNum(kt), `<div class="dl"><b>${kt - ky >= 0 ? '+' : ''}${fmtNum(kt - ky)}</b> dari kemarin</div>`],
      ['t5', '#0ea5e9', 'file', 'Tersedia', fmtNum(avail), `<div class="dl"><b>${total ? Math.round(avail / total * 100) : 0}%</b> dari total buku</div>`]
    ];
    $('#d-stats').innerHTML = cards.map(c => `<div class="card stat ${c[0]}"><div class="ico" style="background:${c[1]}">${ic(c[2], 24)}</div><div><div class="lb">${c[3]}</div><div class="vl">${c[4]}</div>${c[5]}</div></div>`).join('');

    // grafik batang
    const days = S.chartDays || 7, WD = ['Min', 'Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab'], cnt = {};
    for (const p of S.pinjam) cnt[p.tglPinjam] = (cnt[p.tglPinjam] || 0) + 1;
    const labs = [], vals = [], cols = [];
    for (let i = days - 1; i >= 0; i--) {
      const d = addDays(t, -i), [yy, mm, dd] = d.split('-').map(Number);
      labs.push(days <= 7 ? WD[new Date(yy, mm - 1, dd).getDay()] : pad(dd) + '/' + pad(mm));
      vals.push(cnt[d] || 0); cols.push(i === 0 ? '#0147ac' : '#5aa9f7');
    }
    if (!window.Chart) {
      const mx = Math.max(1, ...vals);
      $('.cbox').innerHTML = `<div style="display:flex;align-items:flex-end;gap:6px;height:100%;padding-top:10px">${vals.map((v, i) => `<div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;height:100%;font-size:10px;color:#64748b"><span>${v}</span><i style="display:block;width:100%;max-width:30px;border-radius:6px;background:${cols[i]};height:${Math.max(3, v / mx * 78)}%"></i><span>${labs[i]}</span></div>`).join('')}</div>`;
    } else {
      if (ch.bar && ch.bar.data.labels.length === labs.length) { ch.bar.data.labels = labs; ch.bar.data.datasets[0].data = vals; ch.bar.data.datasets[0].backgroundColor = cols; ch.bar.update(); }
      else {
        if (ch.bar) ch.bar.destroy();
        ch.bar = new Chart($('#c-bar'), {
          type: 'bar', data: { labels: labs, datasets: [{ data: vals, backgroundColor: cols, borderRadius: 6, borderSkipped: false, maxBarThickness: 30 }] },
          options: { responsive: true, maintainAspectRatio: false, animation: { duration: 450 }, plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => ' ' + c.parsed.y + ' peminjaman' } } }, scales: { x: { grid: { display: false }, ticks: { color: '#64748b', font: { size: 11 }, maxRotation: 0, autoSkip: true } }, y: { beginAtZero: true, border: { display: false }, ticks: { precision: 0, color: '#94a3b8', font: { size: 11 }, maxTicksLimit: 6 }, grid: { color: '#eef2fa' } } } }
        });
      }
    }

    // donat kategori
    const kc = {}; S.buku.forEach(b => { const k = txt(b.kategori) || 'Lainnya'; kc[k] = (kc[k] || 0) + 1; });
    let ents = Object.entries(kc).sort((a, b) => b[1] - a[1]);
    if (ents.length > 7) { const rest = ents.slice(6).reduce((s, e) => s + e[1], 0); ents = ents.slice(0, 6).concat([['Lainnya', rest]]); }
    const sum = ents.reduce((s, e) => s + e[1], 0) || 1;
    $('#d-mid').innerHTML = `<b>${fmtNum(total)}</b><span>Buku</span>`;
    $('#d-leg').innerHTML = ents.length ? ents.map((e, i) => `<li><i style="background:${PAL[i % PAL.length]}"></i><span>${esc(e[0])}</span><b>${fmtNum(e[1])} (${Math.round(e[1] / sum * 100)}%)</b></li>`).join('') : '<li class="sub">Belum ada buku</li>';
    if (!window.Chart) {
      let acc = 0; const segs = ents.map((e, i) => { const a = acc / sum * 100; acc += e[1]; return PAL[i % PAL.length] + ' ' + a + '% ' + (acc / sum * 100) + '%'; });
      const dw = $('.dwrap'); dw.style.borderRadius = '50%'; dw.style.background = ents.length ? 'conic-gradient(' + segs.join(',') + ')' : '#e6eefc';
      $('#d-mid').style.cssText = 'inset:22px;background:#fff;border-radius:50%';
    } else {
      const dd = { labels: ents.map(e => e[0]), datasets: [{ data: ents.map(e => e[1]), backgroundColor: ents.map((e, i) => PAL[i % PAL.length]), borderWidth: 2, borderColor: '#fff' }] };
      if (ch.cat) { ch.cat.data = dd; ch.cat.update(); }
      else ch.cat = new Chart($('#c-cat'), { type: 'doughnut', data: dd, options: { responsive: true, maintainAspectRatio: false, cutout: '68%', animation: { duration: 450 }, plugins: { legend: { display: false } } } });
    }

    // aktivitas
    $('#d-act').innerHTML = S.akt.length ? S.akt.slice(0, 5).map(a => {
      const m = AKT[a.jenis] || AKT.pinjam;
      const dt = new Date(a.waktu || 0), same = ymd(dt) === t;
      return `<li><span class="ai" style="background:${m[0]}">${ic(m[2], 18)}</span><div class="at"><b>${m[1]}</b><span>${esc(a.nama)}${a.detail ? ' \u2013 ' + esc(a.detail) : ''}</span></div><time>${same ? fmtTime(a.waktu) : pad(dt.getDate()) + '/' + pad(dt.getMonth() + 1)}</time></li>`;
    }).join('') : '<li class="sub">Belum ada aktivitas.</li>';

    // buku terpopuler
    const pc = new Map();
    for (const p of S.pinjam) { const k = norm(p.kodeBuku); if (!pc.has(k)) pc.set(k, { kode: p.kodeBuku, judul: p.judulBuku, n: 0 }); pc.get(k).n++; }
    const top = Array.from(pc.values()).sort((a, b) => b.n - a.n).slice(0, 4);
    $('#d-pop').innerHTML = top.length ? top.map(p => {
      const b = findBuku(p.kode), jd = (b && b.judul) || p.judul;
      return `<div class="pop"><div class="spine" style="background:linear-gradient(160deg,hsl(${hue(jd)} 70% 48%),hsl(${(hue(jd) + 30) % 360} 70% 30%))">${esc(String(jd).charAt(0).toUpperCase())}</div><div><b>${esc(jd)}</b><span>${esc(b ? b.pengarang : '')}</span><em>${fmtNum(p.n)} kali dipinjam</em></div></div>`;
    }).join('') : '<p class="sub">Belum ada data peminjaman.</p>';

    // siswa paling sering meminjam (6 bulan terakhir)
    const sc = new Map(), since = sixMonthsAgo();
    for (const pj of S.pinjam) { if (pj.tglPinjam < since) continue; const k = norm(pj.nis); if (!sc.has(k)) sc.set(k, { nis: pj.nis, nama: pj.namaSiswa, kelas: pj.kelas, n: 0 }); sc.get(k).n++; }
    const tops = Array.from(sc.values()).sort((a, b) => b.n - a.n || natSort(a.nama, b.nama)).slice(0, 5);
    $('#d-sis').innerHTML = tops.length ? tops.map((x, i) => {
      const a = findAnggota(x.nis), nm = (a && a.nama) || x.nama, kl = (a && a.kelas) || x.kelas;
      return `<li><span class="ai rk" style="background:${['#f59e0b', '#94a3b8', '#c2763b', '#2f6fe0', '#2f6fe0'][i]}">${i + 1}</span><div class="at"><b>${esc(nm)}</b><span>Kelas ${esc(kl)}</span></div><span class="bdg info">${fmtNum(x.n)} kali</span></li>`;
    }).join('') : '<li class="sub">Belum ada data peminjaman.</li>';

    // per kelas
    const lv = { VII: 0, VIII: 0, IX: 0 }; let lt = 0; const lim = addDays(t, -29);
    for (const p of S.pinjam) { if (p.tglPinjam < lim) continue; const l = levelOf(p.kelas); if (l in lv) { lv[l]++; lt++; } }
    $('#d-kls').innerHTML = Object.keys(lv).map(k => { const pc2 = lt ? Math.round(lv[k] / lt * 100) : 0; return `<div class="kbar"><b>${k}</b><div class="tr"><i style="width:${pc2}%"></i></div><span>${fmtNum(lv[k])} (${pc2}%)</span></div>`; }).join('');
  }
};
CHG.chartDays = t => { S.chartDays = +t.value || 7; if (ch.bar) { ch.bar.destroy(); delete ch.bar; } views.beranda.update(); };
ACT.popAll = () => { Rpt.jenis = 'populer'; go('laporan'); };
ACT.sisAll = () => { Rpt.jenis = 'anggota'; Rpt.from = sixMonthsAgo(); Rpt.to = today(); go('laporan'); };
function sixMonthsAgo() { const [y, m, d] = today().split('-').map(Number); return ymd(new Date(y, m - 1 - 6, d)); }
/* ============ KATALOG BUKU ============ */
const PER = 25;
function bukuFiltered() {
  const q = norm(Q.k.q);
  return S.buku.filter(b => {
    if (Q.k.kat && txt(b.kategori) !== Q.k.kat) return false;
    const tv = availOf(b);
    if (Q.k.stok === 'ada' && tv <= 0) return false;
    if (Q.k.stok === 'habis' && tv > 0) return false;
    if (q) return norm(b.kodeBuku).includes(q) || norm(b.judul).includes(q) || norm(b.pengarang).includes(q) || norm(b.penerbit).includes(q) || norm(b.klasifikasi).includes(q);
    return true;
  }).sort((a, b) => ((b.dibuatPada || 0) - (a.dibuatPada || 0)) || natSort(a.kodeBuku, b.kodeBuku));
}
function bukuPage() {
  const all = bukuFiltered(), pages = Math.max(1, Math.ceil(all.length / PER));
  if (Q.k.page > pages) Q.k.page = pages;
  return { all, rows: all.slice((Q.k.page - 1) * PER, Q.k.page * PER) };
}
views.katalog = {
  mount(el) {
    el.innerHTML = `
<div class="vh"><div><h2>Katalog Buku</h2><p class="sub">Kelola koleksi buku, kode barcode, dan stiker.</p></div></div>
<div class="tools">
  <div class="srch">${ic('search', 18)}<input class="in" data-in="kq" placeholder="Cari kode, judul, pengarang\u2026" value="${esc(Q.k.q)}" aria-label="Cari buku"></div>
  <select class="in" data-ch="kkat" aria-label="Kategori"></select>
  <select class="in" data-ch="kstok" aria-label="Stok"><option value="">Semua stok</option><option value="ada">Tersedia</option><option value="habis">Sedang dipinjam semua</option></select>
</div>
<div class="tools">
  <button class="btn pri" data-act="bukuAdd">${ic('plus', 18)}Tambah Buku</button>
  <button class="btn" data-act="importBuku">${ic('upload', 18)}Impor xlsx</button>
  <button class="btn" data-act="exportBuku">${ic('download', 18)}Ekspor xlsx</button>
  <button class="btn" data-act="tplBuku">${ic('file', 18)}Template</button>
  <button class="btn warn" data-act="stikerOpen" id="k-stk">${ic('tag', 18)}Cetak Stiker</button>
  <button class="btn dng sm" data-act="bukuDelSel" id="k-del" style="display:none">${ic('trash', 16)}Hapus terpilih</button>
</div>
<div id="k-zone"></div>`;
    $('[data-ch=kstok]', el).value = Q.k.stok;
  },
  update() {
    const sel = $('[data-ch=kkat]');
    const cats = kategoriList(), key = cats.join('|');
    if (sel.dataset.key !== key) { sel.innerHTML = '<option value="">Semua kategori</option>' + cats.map(c => `<option>${esc(c)}</option>`).join(''); sel.dataset.key = key; }
    sel.value = Q.k.kat;
    const { all, rows } = bukuPage();
    const z = $('#k-zone');
    if (!all.length) { z.innerHTML = `<div class="tw">${emptyHtml('book', S.buku.length ? 'Tidak ada buku yang cocok dengan pencarian.' : 'Belum ada buku. Klik "Tambah Buku" atau impor dari file xlsx.')}</div>`; updSel(); return; }
    z.innerHTML = `<div class="tw"><table class="tbl"><thead><tr><th><input type="checkbox" data-ch="selAllBuku" aria-label="Pilih semua di halaman ini"></th><th class="kd">Kode</th><th>Judul</th><th>Klasifikasi</th><th>Pengarang</th><th>Kategori</th><th>Rak</th><th>Stok</th><th></th></tr></thead><tbody>${rows.map(b => {
      const tv = availOf(b), tt = Number(b.jumlahTotal) || 0;
      return `<tr><td data-l=""><input type="checkbox" data-ch="selBuku" data-id="${esc(b.id)}" ${SEL.buku.has(b.id) ? 'checked' : ''} aria-label="Pilih ${esc(b.judul)}"></td>
<td class="kd" data-l="Kode"><b>${esc(b.kodeBuku)}</b></td><td data-l="Judul"><div>${esc(b.judul)}<small>${esc(b.penerbit || '')}${b.tahun ? ' \u2022 ' + esc(b.tahun) : ''}</small></div></td>
<td data-l="Klasifikasi">${esc(b.klasifikasi || '-')}</td>
<td data-l="Pengarang">${esc(b.pengarang || '-')}</td><td data-l="Kategori"><span class="bdg info">${esc(b.kategori || '-')}</span></td><td data-l="Rak">${esc(b.rak || '-')}</td>
<td data-l="Stok"><span class="bdg ${tv > 0 ? 'ok' : 'bad'}">${tv} / ${tt}</span></td>
<td class="act" data-l=""><button class="ibtn" data-act="stikerOne" data-id="${esc(b.id)}" title="Cetak stiker" aria-label="Cetak stiker">${ic('tag', 18)}</button><button class="ibtn" data-act="bukuEdit" data-id="${esc(b.id)}" title="Ubah" aria-label="Ubah">${ic('edit', 18)}</button><button class="ibtn d" data-act="bukuDel" data-id="${esc(b.id)}" title="Hapus" aria-label="Hapus">${ic('trash', 18)}</button></td></tr>`;
    }).join('')}</tbody></table></div>${pagerHtml(all.length, Q.k.page, PER)}`;
    updSel();
  }
};
function updSel() {
  const kb = $('#k-del'), ks = $('#k-stk'), ab = $('#a-del'), ak = $('#a-card');
  const nb = [...SEL.buku].filter(id => S.buku.some(b => b.id === id)).length, na = [...SEL.anggota].filter(id => S.anggota.some(a => a.id === id)).length;
  if (kb) { kb.style.display = nb ? '' : 'none'; kb.innerHTML = `${ic('trash', 16)}Hapus terpilih (${nb})`; }
  if (ks) ks.innerHTML = `${ic('tag', 18)}Cetak Stiker${nb ? ' (' + nb + ')' : ''}`;
  if (ab) { ab.style.display = na ? '' : 'none'; ab.innerHTML = `${ic('trash', 16)}Hapus terpilih (${na})`; }
  if (ak) ak.innerHTML = `${ic('card', 18)}Cetak Kartu${na ? ' (' + na + ')' : ''}`;
  const hb = $('[data-ch=selAllBuku]'), ha = $('[data-ch=selAllAnggota]');
  if (hb) { const ids = bukuPage().rows.map(b => b.id); hb.checked = ids.length > 0 && ids.every(i => SEL.buku.has(i)); }
  if (ha) { const ids = anggotaPage().rows.map(a => a.id); ha.checked = ids.length > 0 && ids.every(i => SEL.anggota.has(i)); }
}
INP.kq = t => { Q.k.q = t.value; Q.k.page = 1; const h = $('#lib-q'); if (h) h.value = t.value; views.katalog.update(); };
INP.hq = t => { Q.k.q = t.value; Q.k.page = 1; if (cur !== 'katalog') go('katalog'); else { const l = $('[data-in=kq]'); if (l) l.value = t.value; views.katalog.update(); } };
CHG.kkat = t => { Q.k.kat = t.value; Q.k.page = 1; views.katalog.update(); };
CHG.kstok = t => { Q.k.stok = t.value; Q.k.page = 1; views.katalog.update(); };
CHG.selBuku = t => { t.checked ? SEL.buku.add(t.dataset.id) : SEL.buku.delete(t.dataset.id); updSel(); };
CHG.selAllBuku = t => { bukuPage().rows.forEach(b => t.checked ? SEL.buku.add(b.id) : SEL.buku.delete(b.id)); views.katalog.update(); };
ACT.pg = t => { const q = cur === 'katalog' ? Q.k : cur === 'gtk' ? Q.g : Q.a; q.page = Math.max(1, q.page + (+t.dataset.v)); views[cur].update(); $('#lib-view').scrollTo({ top: 0 }); };
const nextKode = () => { let m = 0; S.buku.forEach(b => { const x = /^BK-(\d+)$/i.exec(b.kodeBuku || ''); if (x) m = Math.max(m, parseInt(x[1], 10)); }); return 'BK-' + pad(m + 1, 4); };
function bukuForm(b) {
  const edit = !!b;
  b = b || { kodeBuku: '', jumlahTotal: 1, kategori: kategoriPengaturan()[0] || '' };
  const cats = kategoriPengaturan();
  if (edit && txt(b.kategori) && !cats.includes(txt(b.kategori))) cats.push(txt(b.kategori));
  openModal(`${mHead(edit ? 'Ubah Buku' : 'Tambah Buku')}
<form data-sub="bukuSave" data-edit="${edit ? 1 : 0}" autocomplete="off" style="display:contents">
<div class="mb"><div class="grid2">
 <div class="fld"><label>Kode buku (isi barcode)*</label>${edit ? `<input class="in" name="kode" value="${esc(b.kodeBuku)}" readonly required maxlength="40">` : `<div class="kfld"><input class="in" id="bk-kode" name="kode" value="" data-enter="bukuKode" data-ch="bukuKode" required maxlength="40" placeholder="Ketik atau scan kode buku" autocapitalize="off" spellcheck="false"><button type="button" class="btn" data-act="bukuCam" aria-label="Scan kode buku dengan kamera" title="Scan dengan kamera">${ic('camera', 20)}</button></div><div class="kreader" id="bk-reader-wrap" style="display:none"><div id="bk-reader" class="reader"></div><p class="hint" style="margin-top:8px">Arahkan kamera ke barcode buku. Kode akan terisi otomatis.</p></div>`}<small>Tanpa spasi dan tanda /. Kode ini dicetak sebagai barcode stiker.</small></div>
 <div class="fld"><label>Jumlah eksemplar*</label><input class="in" name="jumlah" type="number" min="1" max="9999" value="${esc(b.jumlahTotal)}" required></div>
 <div class="fld full"><label>Judul*</label><input class="in" name="judul" value="${esc(b.judul || '')}" required maxlength="200"></div>
 <div class="fld full"><label>Klasifikasi</label><input class="in" name="klasifikasi" value="${esc(b.klasifikasi || '')}" maxlength="60"></div>
 <div class="fld"><label>Pengarang</label><input class="in" name="pengarang" value="${esc(b.pengarang || '')}" maxlength="120"></div>
 <div class="fld"><label>Penerbit</label><input class="in" name="penerbit" value="${esc(b.penerbit || '')}" maxlength="120"></div>
 <div class="fld"><label>Tahun terbit</label><input class="in" name="tahun" value="${esc(b.tahun || '')}" maxlength="10" inputmode="numeric"></div>
 <div class="fld"><label>Kategori</label><select class="in" name="kategori" style="width:100%">${cats.map(c => `<option ${c === b.kategori ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select></div>
 <div class="fld full"><label>Rak / lokasi</label><input class="in" name="rak" value="${esc(b.rak || '')}" maxlength="30"></div>
</div></div>
<div class="mf"><button type="button" class="btn" data-act="closeModal">Batal</button><button class="btn pri" type="submit">${ic('check', 18)}Simpan</button></div></form>`);
}
// Cek kode buku saat Tambah Buku: sudah ada -> notifikasi singkat, belum ada -> fokus ke Judul
function cekKodeBuku(inp, pindah) {
  if (!inp) return;
  const kode = txt(inp.value);
  if (!kode) return;
  if (!pindah && inp.dataset.chk === norm(kode)) return;
  inp.dataset.chk = norm(kode);
  const err = validKode(kode);
  if (err) { toast(err, 'warn', 2500); return; }
  if (findBuku(kode)) { toast('Kode buku "' + kode + '" sudah ada.', 'warn', 2500); inp.focus(); inp.select(); return; }
  if (pindah) { const j = $('input[name=judul]'); if (j) j.focus(); }
}
ENT.bukuKode = t => cekKodeBuku(t, true);
CHG.bukuKode = t => cekKodeBuku(t, false);
ACT.bukuAdd = () => bukuForm();
ACT.bukuCam = async () => {
  const wrap = $('#bk-reader-wrap'), inp = $('#bk-kode');
  if (!wrap || !inp) return;
  if (scanner) { stopScanner(); wrap.style.display = 'none'; return; }
  wrap.style.display = '';
  try {
    const { Html5Qrcode, Html5QrcodeSupportedFormats } = await import('html5-qrcode');
    const F = Html5QrcodeSupportedFormats;
    const sc = new Html5Qrcode('bk-reader', { formatsToSupport: [F.CODE_128, F.CODE_39, F.EAN_13, F.QR_CODE], verbose: false });
    scanner = sc;
    await sc.start({ facingMode: 'environment' }, { fps: 10, qrbox: { width: 260, height: 120 } }, code => {
      const now = Date.now(); if (code === lastScan.c && now - lastScan.t < 2500) return;
      lastScan = { c: code, t: now };
      inp.value = txt(code); beep(true);
      stopScanner(); wrap.style.display = 'none';
      cekKodeBuku(inp, true);
    }, () => { });
  } catch (e) { scanner = null; wrap.style.display = 'none'; toast('Kamera tidak dapat dibuka: ' + (e && e.message ? e.message : e), 'err', 5500); }
};
ACT.bukuEdit = t => { const b = S.buku.find(x => x.id === t.dataset.id); if (b) bukuForm(b); };
SUB.bukuSave = async f => {
  const fd = new FormData(f), edit = f.dataset.edit === '1';
  const kode = txt(fd.get('kode')), judul = txt(fd.get('judul')), jumlah = parseInt(fd.get('jumlah'), 10);
  let err = validKode(kode);
  if (!err && !judul) err = 'Judul wajib diisi.';
  if (!err && !(jumlah >= 1 && jumlah <= 9999)) err = 'Jumlah eksemplar harus antara 1 dan 9999.';
  if (!err && !edit && findBuku(kode)) err = 'Kode buku sudah terdaftar.';
  const act = activeOfBook(kode).length;
  if (!err && edit && jumlah < act) err = 'Jumlah tidak boleh kurang dari buku yang sedang dipinjam (' + act + ').';
  if (err) { toast(err, 'err'); return; }
  const d = { kodeBuku: kode, judul, pengarang: txt(fd.get('pengarang')), klasifikasi: txt(fd.get('klasifikasi')), penerbit: txt(fd.get('penerbit')), tahun: txt(fd.get('tahun')), kategori: txt(fd.get('kategori')) || 'Lainnya', rak: txt(fd.get('rak')), jumlahTotal: jumlah, jumlahTersedia: jumlah - act };
  const ops = [];
  if (edit) { const old = findBuku(kode); if (!old) { toast('Data buku tidak ditemukan.', 'err'); return; } ops.push({ t: 'upd', c: 'buku', id: old.id, d }); }
  else { d.dibuatPada = Date.now(); ops.push({ t: 'set', c: 'buku', id: kode, d }); ops.push(actOp('buku', judul, 'Kode ' + kode)); }
  const btn = $('button[type=submit]', f); btn.disabled = true;
  try { await withProg(edit ? 'Memperbarui buku' : 'Menyimpan buku', judul, () => commit(ops)); closeModal(); toast(edit ? 'Data buku diperbarui.' : 'Buku "' + judul + '" ditambahkan.', 'ok'); }
  catch (e) { btn.disabled = false; toast('Gagal menyimpan: ' + (e.message || e), 'err', 5000); }
};
ACT.bukuDel = async t => {
  const b = S.buku.find(x => x.id === t.dataset.id); if (!b) return;
  if (activeOfBook(b.kodeBuku).length) { toast('Buku masih dipinjam, tidak bisa dihapus.', 'err'); return; }
  if (!await confirmBox(`Hapus buku "${b.judul}" (${b.kodeBuku})?`, 'Hapus', true)) return;
  try { await quick(() => commit([{ t: 'del', c: 'buku', id: b.id }])); SEL.buku.delete(b.id); toast('Buku dihapus.', 'ok'); } catch (e) { toast('Gagal menghapus: ' + (e.message || e), 'err'); }
};
ACT.bukuDelSel = async () => {
  const items = [...SEL.buku].map(id => S.buku.find(b => b.id === id)).filter(Boolean);
  const del = items.filter(b => !activeOfBook(b.kodeBuku).length), skip = items.length - del.length;
  if (!del.length) { toast('Semua buku terpilih masih dipinjam.', 'err'); return; }
  if (!await confirmBox(`Hapus ${del.length} buku terpilih?${skip ? ' ' + skip + ' buku yang sedang dipinjam dilewati.' : ''}`, 'Hapus', true)) return;
  try { await withProg('Menghapus buku', del.length + ' buku', p => commit(del.map(b => ({ t: 'del', c: 'buku', id: b.id })), (d, t) => p.set(d / t * 100, fmtNum(d) + ' dari ' + fmtNum(t)))); del.forEach(b => SEL.buku.delete(b.id)); toast(del.length + ' buku dihapus.', 'ok'); } catch (e) { toast('Gagal menghapus: ' + (e.message || e), 'err'); }
};

/* ---- cetak stiker ---- */
let STK = { ids: null };
function stikerModal(ids) {
  STK.ids = ids || null;
  const flt = bukuFiltered().length, sel = [...SEL.buku].filter(id => S.buku.some(b => b.id === id)).length;
  const opts = [];
  if (ids) opts.push(['ids', 'Buku ini saja', ids.length]);
  else { if (sel) opts.push(['sel', 'Buku yang dicentang', sel]); opts.push(['flt', 'Hasil pencarian/filter saat ini', flt]); }
  opts.push(['all', 'Semua buku', S.buku.length]);
  openModal(`${mHead('Cetak Stiker Barcode Kode Buku')}
<form data-sub="stikerPrint" style="display:contents"><div class="mb">
 <div class="fld" style="margin-bottom:14px"><label>Buku yang dicetak</label>${opts.map((o, i) => `<label class="chk"><input type="radio" name="src" value="${o[0]}" ${i === 0 ? 'checked' : ''}> ${o[1]} (${fmtNum(o[2])})</label>`).join('')}</div>
 <div class="grid2">
  <div class="fld"><label>Ukuran stiker</label><select class="in" name="size" style="width:100%"><option value="k">Kecil - 2 kolom (97 x 28 mm)</option><option value="s" selected>Sedang - 2 kolom (97 x 34 mm)</option><option value="b">Besar - 2 kolom (97 x 42 mm)</option></select></div>
  <div class="fld"><label>Salinan per kode</label><input class="in" name="copies" type="number" min="1" max="20" value="1"></div>
 </div>
 <p class="hint" style="margin-top:14px">Kiri: judul, barcode, dan kode buku. Tengah: area lipatan 0,1 cm (kosong). Kanan: nama sekolah, klasifikasi, 3 huruf awal pengarang, 1 huruf awal judul, dan kode rak. Kertas A4. Pada kotak cetak browser pilih skala 100% (Default) dan matikan "Header dan footer".</p>
</div>
<div class="mf"><button type="button" class="btn" data-act="closeModal">Batal</button><button class="btn warn" type="submit">${ic('print', 18)}Cetak Stiker</button></div></form>`);
}
ACT.stikerOpen = () => stikerModal(null);
ACT.stikerOne = t => stikerModal([t.dataset.id]);
SUB.stikerPrint = async f => {
  const fd = new FormData(f), src = fd.get('src'), copies = Math.min(20, Math.max(1, parseInt(fd.get('copies'), 10) || 1));
  let list;
  if (src === 'ids') list = STK.ids.map(id => S.buku.find(b => b.id === id)).filter(Boolean);
  else if (src === 'sel') list = [...SEL.buku].map(id => S.buku.find(b => b.id === id)).filter(Boolean);
  else if (src === 'flt') list = bukuFiltered();
  else list = S.buku.slice();
  list.sort((a, b) => natSort(a.kodeBuku, b.kodeBuku));
  if (!list.length) { toast('Tidak ada buku untuk dicetak.', 'err'); return; }
  if (list.length * copies > 3000) { toast('Maksimal 3.000 stiker sekali cetak. Kurangi pilihan buku.', 'err'); return; }
  const P = { k: [2, '28mm', 48], s: [2, '34mm', 60], b: [2, '42mm', 72] }[fd.get('size')] || [2, '34mm', 60], aw = (v, n) => Array.from(txt(v)).slice(0, n).join('');
  const items = []; list.forEach(b => { for (let i = 0; i < copies; i++) items.push(b); });
  const html = `<div class="stk-grid" style="--cols:${P[0]};--h:${P[1]}">${items.map(b => `<div class="stk"><div class="sl"><div class="jl">${esc(b.judul)}</div><svg data-bc="${esc(b.kodeBuku)}" data-h="${P[2]}" data-fs="15" data-mb="7"></svg></div><div class="sf"></div><div class="sr"><div class="sn">${esc(S.set.namaSekolah)}</div>${b.klasifikasi ? `<div class="rw"><b>${esc(b.klasifikasi)}</b></div>` : ''}${aw(b.pengarang, 3) ? `<div class="rw"><span>${esc(aw(b.pengarang, 3))}</span></div>` : ''}${aw(b.judul, 1) ? `<div class="rw"><span>${esc(aw(b.judul, 1))}</span></div>` : ''}${b.rak ? `<div class="rk"><b>${esc(b.rak)}</b></div>` : ''}</div></div>`).join('')}</div>`;
  closeModal();
  await doPrint(html, 'Stiker barcode ' + fmtNum(items.length) + ' label');
};

/* ============ ANGGOTA ============ */
function anggotaFiltered() {
  const q = norm(Q.a.q);
  return S.anggota.filter(a => {
    if (Q.a.kelas && txt(a.kelas) !== Q.a.kelas) return false;
    if (q) return norm(a.nama).includes(q) || norm(a.nis).includes(q) || norm(a.nisn).includes(q);
    return true;
  }).sort((a, b) => natSort(a.kelas, b.kelas) || natSort(a.nama, b.nama));
}
function anggotaPage() {
  const all = anggotaFiltered(), pages = Math.max(1, Math.ceil(all.length / PER));
  if (Q.a.page > pages) Q.a.page = pages;
  return { all, rows: all.slice((Q.a.page - 1) * PER, Q.a.page * PER) };
}
views.anggota = {
  mount(el) {
    el.innerHTML = `
<div class="vh"><div><h2>Anggota</h2><p class="sub">Data siswa sebagai anggota perpustakaan. NIS dipakai sebagai barcode kartu.</p></div></div>
<div class="tools">
  <div class="srch">${ic('search', 18)}<input class="in" data-in="aq" placeholder="Cari nama, NIS, atau NISN\u2026" value="${esc(Q.a.q)}" aria-label="Cari anggota"></div>
  <select class="in" data-ch="akelas" aria-label="Kelas"></select>
</div>
<div class="tools">
  <button class="btn pri" data-act="anggotaAdd">${ic('plus', 18)}Tambah Anggota</button>
  <button class="btn" data-act="importAnggota">${ic('upload', 18)}Impor xlsx</button>
  <button class="btn" data-act="exportAnggota">${ic('download', 18)}Ekspor xlsx</button>
  <button class="btn" data-act="tplAnggota">${ic('file', 18)}Template</button>
  <button class="btn warn" data-act="kartuOpen" id="a-card">${ic('card', 18)}Cetak Kartu</button>
  <button class="btn dng sm" data-act="anggotaDelSel" id="a-del" style="display:none">${ic('trash', 16)}Hapus terpilih</button>
</div>
<div id="a-zone"></div>`;
  },
  update() {
    const sel = $('[data-ch=akelas]'), ks = kelasList(), key = ks.join('|');
    if (sel.dataset.key !== key) { sel.innerHTML = '<option value="">Semua kelas</option>' + ks.map(k => `<option>${esc(k)}</option>`).join(''); sel.dataset.key = key; }
    sel.value = Q.a.kelas;
    const { all, rows } = anggotaPage(), z = $('#a-zone');
    if (!all.length) { z.innerHTML = `<div class="tw">${emptyHtml('users', S.anggota.length ? 'Tidak ada anggota yang cocok.' : 'Belum ada anggota. Klik "Tambah Anggota" atau impor dari file xlsx.')}</div>`; updSel(); return; }
    z.innerHTML = `<div class="tw"><table class="tbl"><thead><tr><th><input type="checkbox" data-ch="selAllAnggota" aria-label="Pilih semua di halaman ini"></th><th>NIS</th><th>NISN</th><th>Nama</th><th>Kelas</th><th>L/P</th><th>Dipinjam</th><th></th></tr></thead><tbody>${rows.map(a => {
      const n = activeOf(a.nis).length;
      return `<tr><td data-l=""><input type="checkbox" data-ch="selAnggota" data-id="${esc(a.id)}" ${SEL.anggota.has(a.id) ? 'checked' : ''} aria-label="Pilih ${esc(a.nama)}"></td>
<td data-l="NIS"><b>${esc(a.nis)}</b></td><td data-l="NISN">${esc(a.nisn || '-')}</td><td data-l="Nama"><div>${esc(a.nama)}${a.aktif === false ? ' <span class="bdg mut">Nonaktif</span>' : ''}</div></td>
<td data-l="Kelas">${esc(a.kelas)}</td><td data-l="L/P">${esc(a.jenisKelamin || '-')}</td><td data-l="Dipinjam">${n ? `<span class="bdg warn">${n} buku</span>` : '<span class="bdg mut">0</span>'}</td>
<td class="act" data-l=""><button class="ibtn" data-act="kartuOne" data-id="${esc(a.id)}" title="Cetak kartu" aria-label="Cetak kartu">${ic('card', 18)}</button><button class="ibtn" data-act="anggotaEdit" data-id="${esc(a.id)}" title="Ubah" aria-label="Ubah">${ic('edit', 18)}</button><button class="ibtn d" data-act="anggotaDel" data-id="${esc(a.id)}" title="Hapus" aria-label="Hapus">${ic('trash', 18)}</button></td></tr>`;
    }).join('')}</tbody></table></div>${pagerHtml(all.length, Q.a.page, PER)}`;
    updSel();
  }
};
INP.aq = t => { Q.a.q = t.value; Q.a.page = 1; views.anggota.update(); };
CHG.akelas = t => { Q.a.kelas = t.value; Q.a.page = 1; views.anggota.update(); };
CHG.selAnggota = t => { t.checked ? SEL.anggota.add(t.dataset.id) : SEL.anggota.delete(t.dataset.id); updSel(); };
CHG.selAllAnggota = t => { anggotaPage().rows.forEach(a => t.checked ? SEL.anggota.add(a.id) : SEL.anggota.delete(a.id)); views.anggota.update(); };
function anggotaForm(a) {
  const edit = !!a; a = a || { jenisKelamin: 'L', aktif: true };
  openModal(`${mHead(edit ? 'Ubah Anggota' : 'Tambah Anggota')}
<form data-sub="anggotaSave" data-edit="${edit ? 1 : 0}" autocomplete="off" style="display:contents">
<div class="mb"><div class="grid2">
 <div class="fld"><label>NIS (isi barcode kartu)*</label><input class="in" name="nis" value="${esc(a.nis || '')}" ${edit ? 'readonly' : ''} required maxlength="30" inputmode="numeric"></div>
 <div class="fld"><label>NISN</label><input class="in" name="nisn" value="${esc(a.nisn || '')}" maxlength="20" inputmode="numeric"><small>Disimpan sebagai teks, angka 0 di depan tetap ada.</small></div>
 <div class="fld full"><label>Nama lengkap*</label><input class="in" name="nama" value="${esc(a.nama || '')}" required maxlength="120"></div>
 <div class="fld"><label>Kelas*</label><input class="in" name="kelas" value="${esc(a.kelas || '')}" list="dl-kelas" required maxlength="20" placeholder="contoh: VII-A"><datalist id="dl-kelas">${kelasList().map(k => `<option value="${esc(k)}">`).join('')}</datalist></div>
 <div class="fld"><label>Jenis kelamin*</label><select class="in" name="jk" style="width:100%"><option value="L" ${a.jenisKelamin === 'L' ? 'selected' : ''}>Laki-laki</option><option value="P" ${a.jenisKelamin === 'P' ? 'selected' : ''}>Perempuan</option></select></div>
 <label class="chk full"><input type="checkbox" name="aktif" ${a.aktif !== false ? 'checked' : ''}> Anggota aktif (boleh meminjam)</label>
</div></div>
<div class="mf"><button type="button" class="btn" data-act="closeModal">Batal</button><button class="btn pri" type="submit">${ic('check', 18)}Simpan</button></div></form>`);
}
ACT.anggotaAdd = () => anggotaForm();
ACT.anggotaEdit = t => { const a = S.anggota.find(x => x.id === t.dataset.id); if (a) anggotaForm(a); };
SUB.anggotaSave = async f => {
  const fd = new FormData(f), edit = f.dataset.edit === '1';
  const nis = txt(fd.get('nis')), nisn = txt(fd.get('nisn')), nama = txt(fd.get('nama')), kelas = txt(fd.get('kelas')), jk = fd.get('jk') === 'P' ? 'P' : 'L';
  let err = validKode(nis).replace('Kode', 'NIS');
  if (!err && !nama) err = 'Nama wajib diisi.';
  if (!err && !kelas) err = 'Kelas wajib diisi.';
  if (!err && nisn && !/^\d+$/.test(nisn)) err = 'NISN hanya boleh berisi angka.';
  if (!err && !edit && S.aMap.has(norm(nis))) err = 'NIS sudah terdaftar.';
  if (!err && nisn) { const o = S.anggota.find(x => norm(x.nisn) === norm(nisn) && norm(x.nis) !== norm(nis)); if (o) err = 'NISN sudah dipakai ' + o.nama + '.'; }
  if (err) { toast(err, 'err'); return; }
  const d = { nis, nisn, nama, kelas, jenisKelamin: jk, aktif: !!fd.get('aktif') };
  const ops = [];
  if (edit) { const old = S.aMap.get(norm(nis)); if (!old) { toast('Data anggota tidak ditemukan.', 'err'); return; } ops.push({ t: 'upd', c: 'anggota', id: old.id, d }); }
  else { d.dibuatPada = Date.now(); ops.push({ t: 'set', c: 'anggota', id: nis, d }); ops.push(actOp('anggota', nama, 'Kelas ' + kelas)); }
  const btn = $('button[type=submit]', f); btn.disabled = true;
  try { await withProg(edit ? 'Memperbarui anggota' : 'Menyimpan anggota', nama, () => commit(ops)); closeModal(); toast(edit ? 'Data anggota diperbarui.' : 'Anggota ' + nama + ' ditambahkan.', 'ok'); }
  catch (e) { btn.disabled = false; toast('Gagal menyimpan: ' + (e.message || e), 'err', 5000); }
};
ACT.anggotaDel = async t => {
  const a = S.anggota.find(x => x.id === t.dataset.id); if (!a) return;
  if (activeOf(a.nis).length) { toast('Anggota masih memiliki pinjaman aktif.', 'err'); return; }
  if (!await confirmBox(`Hapus anggota ${a.nama} (${a.nis})? Riwayat peminjaman tetap tersimpan.`, 'Hapus', true)) return;
  try { await quick(() => commit([{ t: 'del', c: 'anggota', id: a.id }])); SEL.anggota.delete(a.id); toast('Anggota dihapus.', 'ok'); } catch (e) { toast('Gagal menghapus: ' + (e.message || e), 'err'); }
};
ACT.anggotaDelSel = async () => {
  const items = [...SEL.anggota].map(id => S.anggota.find(a => a.id === id)).filter(Boolean);
  const del = items.filter(a => !activeOf(a.nis).length), skip = items.length - del.length;
  if (!del.length) { toast('Semua anggota terpilih masih memiliki pinjaman aktif.', 'err'); return; }
  if (!await confirmBox(`Hapus ${del.length} anggota terpilih?${skip ? ' ' + skip + ' anggota dengan pinjaman aktif dilewati.' : ''}`, 'Hapus', true)) return;
  try { await withProg('Menghapus anggota', del.length + ' anggota', p => commit(del.map(a => ({ t: 'del', c: 'anggota', id: a.id })), (d, t) => p.set(d / t * 100, fmtNum(d) + ' dari ' + fmtNum(t)))); del.forEach(a => SEL.anggota.delete(a.id)); toast(del.length + ' anggota dihapus.', 'ok'); } catch (e) { toast('Gagal menghapus: ' + (e.message || e), 'err'); }
};

/* ---- cetak kartu perpustakaan ---- */
let KRT = { ids: null };
function kartuModal(ids) {
  KRT.ids = ids || null;
  const flt = anggotaFiltered().length, sel = [...SEL.anggota].filter(id => S.anggota.some(a => a.id === id)).length;
  const opts = [];
  if (ids) opts.push(['ids', 'Anggota ini saja', ids.length]);
  else { if (sel) opts.push(['sel', 'Anggota yang dicentang', sel]); opts.push(['flt', 'Hasil pencarian/filter saat ini', flt]); }
  opts.push(['all', 'Semua anggota', S.anggota.length]);
  openModal(`${mHead('Cetak Kartu Perpustakaan')}
<form data-sub="kartuPrint" style="display:contents"><div class="mb">
 <div class="fld"><label>Anggota yang dicetak</label>${opts.map((o, i) => `<label class="chk"><input type="radio" name="src" value="${o[0]}" ${i === 0 ? 'checked' : ''}> ${o[1]} (${fmtNum(o[2])})</label>`).join('')}</div>
 <p class="hint" style="margin-top:12px">Ukuran kartu 85,6 x 54 mm (sama dengan kartu pelajar), 10 kartu per lembar A4. Barcode berisi NIS. Pada kotak cetak browser aktifkan "Grafik latar belakang" dan skala 100%.</p>
</div>
<div class="mf"><button type="button" class="btn" data-act="closeModal">Batal</button><button class="btn warn" type="submit">${ic('print', 18)}Cetak Kartu</button></div></form>`);
}
ACT.kartuOpen = () => kartuModal(null);
ACT.kartuOne = t => kartuModal([t.dataset.id]);
SUB.kartuPrint = async f => {
  const src = new FormData(f).get('src'); let list;
  if (src === 'ids') list = KRT.ids.map(id => S.anggota.find(a => a.id === id)).filter(Boolean);
  else if (src === 'sel') list = [...SEL.anggota].map(id => S.anggota.find(a => a.id === id)).filter(Boolean);
  else if (src === 'flt') list = anggotaFiltered();
  else list = S.anggota.slice().sort((a, b) => natSort(a.kelas, b.kelas) || natSort(a.nama, b.nama));
  if (!list.length) { toast('Tidak ada anggota untuk dicetak.', 'err'); return; }
  if (list.length > 1000) { toast('Maksimal 1.000 kartu sekali cetak. Cetak per kelas.', 'err'); return; }
  const logo = S.set.logoUrl || LOGO;
  const html = `<div class="kt-grid">${list.map(a => `<div class="kt"><div class="kh"><img src="${esc(logo)}" alt=""><div><b>${esc(S.set.namaSekolah)}</b><span>KARTU PERPUSTAKAAN</span></div></div><div class="kn"><b>${esc(a.nama)}</b><span>Kelas ${esc(a.kelas)} \u2022 NIS ${esc(a.nis)}</span></div><div class="kb"><svg data-bc="${esc(a.nis)}" data-h="40" data-fs="12"></svg></div><div class="ks"></div></div>`).join('')}</div>`;
  closeModal();
  await doPrint(html, 'Kartu perpustakaan ' + fmtNum(list.length) + ' kartu');
};
/* ============ CETAK (stiker, kartu, laporan) ============ */
const PRINT_CSS = `*{box-sizing:border-box;-webkit-print-color-adjust:exact;print-color-adjust:exact}html,body{margin:0;padding:0;background:#fff}@page{size:A4;margin:8mm}
.stk-grid{display:grid;grid-template-columns:repeat(var(--cols),1fr);width:100%}
.stk{height:var(--h);border:.2mm dashed #9aa5b8;display:flex;align-items:stretch;overflow:hidden;break-inside:avoid;page-break-inside:avoid;font:6.5pt/1.2 Arial,Helvetica,sans-serif;color:#000;background:#fff}
.stk .sl{flex:1 1 0;min-width:0;padding:1.8mm 1.8mm 1.8mm 2.2mm;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;gap:.8mm}
.stk .sf{flex:none;width:1mm}
.stk .sr{flex:1 1 0;min-width:0;padding:1.8mm 2.2mm 1.8mm 1.8mm;display:flex;flex-direction:column;justify-content:center;align-items:stretch;text-align:center;gap:.9mm}
.stk .jl{max-width:100%;font-weight:700;font-size:6.8pt;line-height:1.15;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;word-break:break-word}
.stk .sl svg{width:100%;height:auto;max-height:calc(var(--h) - 12.5mm);overflow:visible;flex:none}
.stk .sn{font-weight:700;font-size:5.8pt;line-height:1.15;text-transform:uppercase;letter-spacing:.03em;padding-bottom:.8mm;border-bottom:.3mm solid #000;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.stk .rw{min-width:0}
.stk .rw b,.stk .rw span{display:block;min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.stk .rw b{font-size:11pt;line-height:1.1}
.stk .rw span{font-size:7.5pt}
.stk .rk{display:flex;align-items:center;justify-content:center;margin-top:.3mm;padding:.5mm 1.6mm;background:#000;color:#fff;border-radius:.8mm}
.stk .rk b{font-size:9pt;line-height:1.15;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;min-width:0}
.kt-grid{display:grid;grid-template-columns:repeat(2,85.6mm);gap:3mm 5mm;justify-content:center}
.kt{width:85.6mm;height:54mm;border-radius:3mm;overflow:hidden;position:relative;color:#fff;background:linear-gradient(135deg,#0147ac 0%,#001455 100%);break-inside:avoid;page-break-inside:avoid;font-family:Arial,Helvetica,sans-serif}
.kt .kh{display:flex;align-items:center;gap:2.5mm;padding:3mm 4mm 0}
.kt .kh img{width:10mm;height:10mm;object-fit:contain}
.kt .kh b{display:block;font-size:7.5pt;line-height:1.1}
.kt .kh span{display:block;font-size:6pt;letter-spacing:.08em;color:#fed403;font-weight:700;margin-top:.6mm}
.kt .kn{padding:2.2mm 4mm 0}
.kt .kn b{display:block;font-size:10.5pt;line-height:1.15;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.kt .kn span{font-size:7pt;opacity:.92}
.kt .kb{position:absolute;left:4mm;right:4mm;bottom:3mm;background:#fff;border-radius:1.8mm;padding:1.2mm 2mm;height:17mm;display:flex;align-items:center;justify-content:center}
.kt .kb svg{height:100%;width:auto;max-width:100%}
.kt .ks{position:absolute;left:0;right:0;bottom:0;height:1.6mm;background:linear-gradient(90deg,#fed403,#ffb305)}
.rp{font:10pt Arial,Helvetica,sans-serif;color:#000}
.rp h2{font-size:14pt;margin:0 0 2px}.rp p{margin:0 0 8px;color:#444;font-size:9pt}
.rp table{width:100%;border-collapse:collapse;font-size:8.5pt}
.rp th,.rp td{border:.2mm solid #888;padding:3px 5px;text-align:left}
.rp th{background:#e8eefc}`;
/* Cetak lewat iframe terisolasi: tidak terpengaruh struktur/tema Blogger. */
async function doPrint(html, label) {
  if (typeof JsBarcode === 'undefined') { toast('Pustaka barcode belum termuat. Periksa koneksi internet lalu muat ulang halaman.', 'err'); return false; }
  Prog.show('Menyiapkan cetak', label || 'Membuat barcode\u2026', 0); await sleep(60);
  let frame = null;
  try {
    const box = document.createElement('div'); box.innerHTML = html;
    const svgs = $$('svg[data-bc]', box);
    for (let i = 0; i < svgs.length; i++) {
      const s = svgs[i];
      try { JsBarcode(s, s.dataset.bc, { format: 'CODE128', width: 1.6, height: +(s.dataset.h || 34), displayValue: true, fontSize: +(s.dataset.fs || 11), margin: 0, marginBottom: +(s.dataset.mb || 0), textMargin: 1 }); }
      catch (e) { const sp = document.createElement('span'); sp.textContent = s.dataset.bc; s.replaceWith(sp); }
      if (i % 20 === 19) { Prog.set(Math.round((i + 1) / svgs.length * 70), 'Membuat barcode ' + fmtNum(i + 1) + ' dari ' + fmtNum(svgs.length)); await sleep(0); }
    }
    Prog.set(75, 'Menyusun halaman cetak\u2026'); await sleep(30);
    const old = document.getElementById('lib-pf'); if (old) old.remove();
    frame = document.createElement('iframe');
    frame.id = 'lib-pf'; frame.setAttribute('aria-hidden', 'true'); frame.title = 'Cetak';
    frame.style.cssText = 'position:fixed;left:-10000px;top:0;width:210mm;height:297mm;border:0;opacity:0;pointer-events:none';
    document.body.appendChild(frame);
    const w = frame.contentWindow, d = w.document;
    d.open();
    d.write('<!doctype html><html><head><meta charset="utf-8"><title>' + esc(label || 'Cetak') + '</title><style>' + PRINT_CSS + '</style></head><body>' + box.innerHTML + '</body></html>');
    d.close();
    Prog.set(88, 'Memuat gambar\u2026');
    await Promise.all(Array.from(d.images).map(im => im.complete ? 1 : new Promise(r => { im.onload = im.onerror = r; setTimeout(r, 3500); })));
    await sleep(200);
    Prog.set(100, 'Membuka dialog cetak\u2026'); await sleep(250);
    await Prog.hide();
    w.focus(); w.print();
    const f = frame; setTimeout(() => { if (f.parentNode) f.remove(); }, 120000);
    return true;
  } catch (e) {
    console.error(e); if (frame && frame.parentNode) frame.remove();
    await Prog.hide(); toast('Gagal menyiapkan cetak: ' + (e.message || e), 'err'); return false;
  }
}

/* ============ SCAN KAMERA ============ */
const loadScript = src => new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('Gagal memuat pustaka kamera. Periksa koneksi internet.')); document.head.appendChild(s); });
let lastScan = { c: '', t: 0 };
ACT.cam = async t => {
  const target = t.dataset.v;
  openModal(`${mHead('Scan Barcode dengan Kamera')}<div class="mb"><div id="lib-reader" class="reader"></div><p class="hint" style="margin-top:10px">Arahkan kamera ke barcode. Hasil scan langsung diproses. Tutup jendela ini jika sudah selesai.</p></div><div class="mf"><button class="btn pri" data-act="closeModal">Selesai</button></div>`, { noFocus: true });
  try {
    const { Html5Qrcode, Html5QrcodeSupportedFormats } = await import('html5-qrcode');
    const F = Html5QrcodeSupportedFormats;
    scanner = new Html5Qrcode('lib-reader', { formatsToSupport: [F.CODE_128, F.CODE_39, F.EAN_13, F.QR_CODE], verbose: false });
    await scanner.start({ facingMode: 'environment' }, { fps: 10, qrbox: { width: 260, height: 120 } }, code => {
      const now = Date.now(); if (code === lastScan.c && now - lastScan.t < 2500) return;
      lastScan = { c: code, t: now };
      const tgt = (target === 'pNis' || target === 'pBuku') ? (Pm.nis ? 'pBuku' : 'pNis') : target;
      const inp = $(`[data-enter="${tgt}"]`);
      if (inp && ENT[tgt]) { inp.value = code; ENT[tgt](inp); }
    }, () => { });
  } catch (e) { toast('Kamera tidak dapat dibuka: ' + (e && e.message ? e.message : e), 'err', 5500); }
};

/* ============ PEMINJAMAN ============ */
const Pm = { nis: null, items: [], msg: null };
const dueState = p => { const t = today(), late = diffDays(p.jatuhTempo, t); return late > 0 ? { late, cls: 'bad', txt: 'Terlambat ' + late + ' hari' } : { late: 0, cls: 'info', txt: 'Dipinjam' }; };
const scanBox = (id, enter, ph, camTarget) => `<div class="scanrow"><input class="in" id="${id}" data-enter="${enter}" placeholder="${ph}" autocomplete="off" autocapitalize="off" spellcheck="false" aria-label="${ph}"><button class="btn pri" data-act="${enter}Go" data-v="${id}" aria-label="Proses">${ic('check', 18)}</button><button class="btn" data-act="cam" data-v="${camTarget}" aria-label="Scan dengan kamera" title="Scan dengan kamera">${ic('camera', 20)}</button></div>`;
views.peminjaman = {
  mount(el) {
    el.innerHTML = `
<div class="vh"><div><h2>Peminjaman</h2><p class="sub">Scan barcode NIS pada kartu, lalu scan barcode buku. Scanner barcode USB/Bluetooth langsung bisa dipakai.</p></div></div>
<div class="scan2">
  <div class="card step"><h3><s>1</s>Scan kartu anggota (NIS / No. Anggota)</h3>${scanBox('p-nis', 'pNis', 'Scan atau ketik NIS / No. Anggota\u2026', 'pNis')}<div class="mbox" id="p-stu"></div></div>
  <div class="card step"><h3><s>2</s>Scan buku</h3>${scanBox('p-buku', 'pBuku', 'Scan atau ketik kode buku\u2026', 'pBuku')}<div class="mbox" id="p-cart"></div></div>
</div>
<div class="ph" style="margin-top:6px"><h3 style="font-size:15px;font-weight:800">Peminjaman hari ini</h3></div>
<div id="p-list"></div>`;
    setTimeout(() => { const i = $('#p-nis'); if (i) i.focus(); }, 80);
  },
  update() {
    const a = Pm.nis ? S.aMap.get(norm(Pm.nis)) : null, stu = $('#p-stu'), cart = $('#p-cart');
    if (Pm.nis && !a) { Pm.nis = null; Pm.items = []; }
    const maks = Number(S.set.maksPinjam) || 3;
    if (!a) {
      stu.className = 'mbox' + (Pm.msg ? ' er' : '');
      stu.innerHTML = Pm.msg ? `<span style="color:#be123c;font-weight:700">${esc(Pm.msg)}</span>` : '<span class="sub">Belum ada siswa. Scan kartu pelajar untuk memulai.</span>';
    } else {
      const act = activeOf(a.nis);
      stu.className = 'mbox ' + (a.aktif === false ? 'er' : 'ok');
      stu.innerHTML = `<div class="nm">${esc(a.nama)}</div><div class="sub">Kelas ${esc(a.kelas)} \u2022 NIS ${esc(a.nis)} \u2022 NISN ${esc(a.nisn || '-')}</div>
<div style="margin-top:8px"><span class="bdg ${act.length >= maks ? 'bad' : 'info'}">Dipinjam ${act.length}/${maks}</span> ${a.aktif === false ? '<span class="bdg bad">Nonaktif</span>' : ''} ${act.some(p => dueState(p).late) ? '<span class="bdg bad">Ada yang terlambat</span>' : ''}</div>
${act.length ? `<ul class="cart" style="margin-top:6px">${act.map(p => `<li><div><b>${esc(p.judulBuku)}</b><span>Jatuh tempo ${fmtDate(p.jatuhTempo)}</span></div><span class="bdg ${dueState(p).cls}">${dueState(p).txt}</span></li>`).join('')}</ul>` : ''}`;
    }
    cart.className = 'mbox';
    cart.innerHTML = Pm.items.length ? `<ul class="cart">${Pm.items.map((it, i) => `<li><div><b>${esc(it.judul)}</b><span>${esc(it.kode)}</span></div><button class="ibtn d" data-act="pRemove" data-v="${i}" aria-label="Hapus">${ic('close', 18)}</button></li>`).join('')}</ul>
<div style="display:flex;gap:8px;margin-top:10px"><button class="btn pri" data-act="pSave" style="flex:1">${ic('check', 18)}Simpan Peminjaman (${Pm.items.length})</button><button class="btn" data-act="pReset">Batal</button></div>`
      : `<span class="sub">${a ? 'Scan buku yang dipinjam. Lama pinjam ' + (Number(S.set.lamaPinjamHari) || 7) + ' hari.' : 'Scan siswa terlebih dahulu.'}</span>`;
    const t = today(), rows = S.pinjam.filter(p => p.tglPinjam === t);
    $('#p-list').innerHTML = rows.length ? `<div class="tw"><table class="tbl"><thead><tr><th>Jam</th><th>Peminjam</th><th>Buku</th><th>Jatuh tempo</th><th>Status</th></tr></thead><tbody>${rows.map(p => `<tr><td data-l="Jam">${fmtTime(p.tsPinjam)}</td><td data-l="Peminjam"><div><b>${esc(p.namaSiswa)}</b><small>${esc(p.kelas)} \u2022 ${esc(p.nis)}</small></div></td><td data-l="Buku"><div>${esc(p.judulBuku)}<small>${esc(p.kodeBuku)}</small></div></td><td data-l="Jatuh tempo">${fmtDate(p.jatuhTempo)}</td><td data-l="Status">${p.status === 'kembali' ? '<span class="bdg ok">Sudah kembali</span>' : `<span class="bdg ${dueState(p).cls}">${dueState(p).txt}</span>`}</td></tr>`).join('')}</tbody></table></div>` : `<div class="tw">${emptyHtml('borrow', 'Belum ada peminjaman hari ini.')}</div>`;
  }
};
const focusEl = id => setTimeout(() => { const i = $(id); if (i) { i.value = ''; i.focus(); } }, 30);
ENT.pNis = inp => {
  const v = txt(inp.value); if (!v) return;
  inp.value = '';
  const a = findAnggota(v);
  if (!a) { Pm.nis = null; Pm.items = []; Pm.msg = 'Nomor anggota "' + v + '" tidak terdaftar.'; beep(false); toast(Pm.msg, 'err'); views.peminjaman.update(); focusEl('#p-nis'); return; }
  Pm.msg = null;
  if (norm(Pm.nis) !== norm(a.nis)) Pm.items = [];
  Pm.nis = a.nis;
  if (a.aktif === false) { beep(false); toast(a.nama + ' berstatus nonaktif.', 'err'); } else beep(true);
  views.peminjaman.update(); focusEl('#p-buku');
};
ENT.pBuku = inp => {
  const v = txt(inp.value); if (!v) return;
  inp.value = '';
  const a = Pm.nis ? S.aMap.get(norm(Pm.nis)) : null;
  if (!a) { beep(false); toast('Scan kartu anggota (NIS / No. Anggota) terlebih dahulu.', 'err'); focusEl('#p-nis'); return; }
  if (a.aktif === false) { beep(false); toast('Anggota nonaktif tidak bisa meminjam.', 'err'); return; }
  const b = findBuku(v);
  if (!b) { beep(false); toast('Kode buku "' + v + '" tidak ditemukan.', 'err'); focusEl('#p-buku'); return; }
  const inCart = Pm.items.filter(i => norm(i.kode) === norm(b.kodeBuku)).length;
  const maks = Number(S.set.maksPinjam) || 3;
  let err = '';
  if (activeOf(a.nis).some(p => norm(p.kodeBuku) === norm(b.kodeBuku)) || inCart) err = 'Buku ini sudah dipinjam siswa tersebut.';
  else if (activeOf(a.nis).length + Pm.items.length >= maks) err = 'Batas pinjam ' + maks + ' buku per siswa tercapai.';
  else if (availOf(b) - inCart <= 0) err = 'Stok buku "' + b.judul + '" habis.';
  if (err) { beep(false); toast(err, 'err'); focusEl('#p-buku'); return; }
  Pm.items.push({ kode: b.kodeBuku, judul: b.judul }); beep(true);
  views.peminjaman.update(); focusEl('#p-buku');
};
ACT.pNisGo = () => { const i = $('#p-nis'); if (i) ENT.pNis(i); };
ACT.pBukuGo = () => { const i = $('#p-buku'); if (i) ENT.pBuku(i); };
ACT.pRemove = t => { Pm.items.splice(+t.dataset.v, 1); views.peminjaman.update(); };
ACT.pReset = () => { Pm.nis = null; Pm.items = []; Pm.msg = null; views.peminjaman.update(); focusEl('#p-nis'); };
ACT.pSave = async t => {
  const a = Pm.nis ? S.aMap.get(norm(Pm.nis)) : null;
  if (!a || !Pm.items.length) return;
  t.disabled = true;
  const tgl = today(), lama = Number(S.set.lamaPinjamHari) || 7, ops = [];
  try {
    for (const it of Pm.items) {
      const b = findBuku(it.kode);
      if (!b || availOf(b) < 1) throw new Error('Stok "' + it.judul + '" sudah habis.');
      ops.push({ t: 'set', c: 'peminjaman', id: newId('peminjaman'), d: { nis: a.nis, namaSiswa: a.nama, kelas: a.kelas, kodeBuku: b.kodeBuku, judulBuku: b.judul, tglPinjam: tgl, tsPinjam: Date.now(), jatuhTempo: addDays(tgl, lama), tglKembali: '', tsKembali: 0, status: 'dipinjam', denda: 0, terlambat: 0 } });
      ops.push({ t: 'upd', c: 'buku', id: b.id, d: { jumlahTersedia: Math.max(0, (Number(b.jumlahTotal) || 0) - activeOfBook(b.kodeBuku).length - 1) } });
      ops.push(actOp('pinjam', a.nama, b.judul));
    }
    await withProg('Menyimpan peminjaman', Pm.items.length + ' buku untuk ' + a.nama, () => commit(ops));
    beep(true); toast(Pm.items.length + ' buku dipinjam oleh ' + a.nama + '. Jatuh tempo ' + fmtDate(addDays(tgl, lama)) + '.', 'ok', 4200);
    Pm.nis = null; Pm.items = []; Pm.msg = null; views.peminjaman.update(); focusEl('#p-nis');
  } catch (e) { t.disabled = false; beep(false); toast('Gagal menyimpan: ' + (e.message || e), 'err', 5000); }
};

/* ============ PENGEMBALIAN ============ */
const Rt = { q: '', f: '', last: [] };
async function doReturn(p) {
  const cur0 = S.pinjam.find(x => x.id === p.id);
  if (!cur0 || cur0.status !== 'dipinjam') throw new Error('Peminjaman ini sudah dikembalikan.');
  const t = today(), late = Math.max(0, diffDays(p.jatuhTempo, t)), denda = late * (Number(S.set.dendaPerHari) || 0), b = findBuku(p.kodeBuku);
  const ops = [{ t: 'upd', c: 'peminjaman', id: p.id, d: { status: 'kembali', tglKembali: t, tsKembali: Date.now(), denda, terlambat: late } }];
  if (b) ops.push({ t: 'upd', c: 'buku', id: b.id, d: { jumlahTersedia: Math.min(Number(b.jumlahTotal) || 0, (Number(b.jumlahTotal) || 0) - activeOfBook(b.kodeBuku).length + 1) } });
  ops.push(actOp('kembali', p.namaSiswa, p.judulBuku));
  await commit(ops);
  Rt.last.unshift({ nama: p.namaSiswa, kelas: p.kelas, judul: p.judulBuku, late, denda, ts: Date.now() });
  Rt.last = Rt.last.slice(0, 8);
  return { late, denda };
}
const retMsg = (p, r) => p.namaSiswa + ' mengembalikan "' + p.judulBuku + '"' + (r.late ? '. Terlambat ' + r.late + ' hari' + (r.denda ? ', denda Rp ' + fmtNum(r.denda) : '') : '. Tepat waktu') + '.';
views.pengembalian = {
  mount(el) {
    el.innerHTML = `
<div class="vh"><div><h2>Pengembalian</h2><p class="sub">Scan barcode buku yang dikembalikan. Jika kode buku punya beberapa peminjam, pilih namanya.</p></div></div>
<div class="scan2">
  <div class="card step"><h3><s>1</s>Scan buku dikembalikan</h3>${scanBox('r-buku', 'rBuku', 'Scan kode buku (atau nomor anggota)\u2026', 'rBuku')}<div class="mbox" id="r-res"></div></div>
  <div class="card step"><h3>${ic('clock', 18)} Dikembalikan baru-baru ini</h3><div id="r-last"></div></div>
</div>
<div class="ph"><h3 style="font-size:15px;font-weight:800">Peminjaman aktif</h3></div>
<div class="tools"><div class="srch">${ic('search', 18)}<input class="in" data-in="rq" placeholder="Cari peminjam, nomor anggota, buku\u2026" aria-label="Cari peminjaman aktif"></div><select class="in" data-ch="rf" aria-label="Filter"><option value="">Semua</option><option value="late">Terlambat saja</option></select></div>
<div id="r-list"></div>`;
    setTimeout(() => { const i = $('#r-buku'); if (i) i.focus(); }, 80);
  },
  update() {
    const res = $('#r-res');
    if (!res.dataset.set) { res.className = 'mbox'; res.innerHTML = '<span class="sub">Menunggu scan buku\u2026</span>'; res.dataset.set = '1'; }
    $('#r-last').innerHTML = Rt.last.length ? `<ul class="cart">${Rt.last.map(l => `<li><div><b>${esc(l.judul)}</b><span>${esc(l.nama)} \u2022 ${esc(l.kelas)}</span></div><span class="bdg ${l.late ? 'bad' : 'ok'}">${l.late ? l.late + ' hari telat' : 'Tepat waktu'}</span></li>`).join('')}</ul>` : '<span class="sub">Belum ada pengembalian pada sesi ini.</span>';
    const q = norm(Rt.q), t = today();
    let rows = S.pinjam.filter(p => p.status === 'dipinjam');
    if (Rt.f === 'late') rows = rows.filter(p => diffDays(p.jatuhTempo, t) > 0);
    if (q) rows = rows.filter(p => norm(p.namaSiswa).includes(q) || norm(p.nis).includes(q) || norm(p.judulBuku).includes(q) || norm(p.kodeBuku).includes(q));
    rows.sort((a, b) => (a.jatuhTempo < b.jatuhTempo ? -1 : 1));
    const shown = rows.slice(0, 100);
    $('#r-list').innerHTML = rows.length ? `<div class="tw"><table class="tbl"><thead><tr><th>Peminjam</th><th>Buku</th><th>Dipinjam</th><th>Jatuh tempo</th><th>Status</th><th></th></tr></thead><tbody>${shown.map(p => `<tr><td data-l="Peminjam"><div><b>${esc(p.namaSiswa)}</b><small>${esc(p.kelas)} \u2022 ${esc(p.nis)}</small></div></td><td data-l="Buku"><div>${esc(p.judulBuku)}<small>${esc(p.kodeBuku)}</small></div></td><td data-l="Dipinjam">${fmtDate(p.tglPinjam)}</td><td data-l="Jatuh tempo">${fmtDate(p.jatuhTempo)}</td><td data-l="Status"><span class="bdg ${dueState(p).cls}">${dueState(p).txt}</span></td><td class="act" data-l=""><button class="btn sm" data-act="retLoan" data-id="${esc(p.id)}">${ic('return', 16)}Kembalikan</button></td></tr>`).join('')}</tbody></table></div>${rows.length > 100 ? `<p class="hint" style="margin-top:8px">Menampilkan 100 dari ${rows.length}. Gunakan pencarian untuk mempersempit.</p>` : ''}` : `<div class="tw">${emptyHtml('check', Rt.q || Rt.f ? 'Tidak ada data yang cocok.' : 'Tidak ada buku yang sedang dipinjam.')}</div>`;
  }
};
INP.rq = t => { Rt.q = t.value; views.pengembalian.update(); };
CHG.rf = t => { Rt.f = t.value; views.pengembalian.update(); };
function showRes(cls, html) { const r = $('#r-res'); if (r) { r.className = 'mbox ' + cls; r.innerHTML = html; r.dataset.set = '1'; } }
function chooseLoans(list, title) {
  openModal(`${mHead(title)}<div class="mb mlist">${list.map(p => `<button data-act="retChoose" data-id="${esc(p.id)}"><div><b>${esc(p.namaSiswa)}</b> <span class="sub">${esc(p.kelas)} \u2022 ${esc(p.nis)}</span><br><span>${esc(p.judulBuku)}</span> <span class="sub">(${esc(p.kodeBuku)})</span></div><span class="bdg ${dueState(p).cls}">${dueState(p).txt}</span></button>`).join('')}</div><div class="mf"><button class="btn" data-act="closeModal">Batal</button></div>`, { noFocus: true });
}
ENT.rBuku = async inp => {
  const v = txt(inp.value); if (!v) return;
  inp.value = '';
  const b = findBuku(v);
  if (b) {
    const act = activeOfBook(b.kodeBuku);
    if (!act.length) { beep(false); showRes('er', `<b style="color:#be123c">"${esc(b.judul)}" tidak sedang dipinjam.</b>`); toast('Buku tidak sedang dipinjam.', 'err'); focusEl('#r-buku'); return; }
    if (act.length > 1) { beep(true); chooseLoans(act, 'Pilih peminjam buku "' + b.judul + '"'); return; }
    await retOne(act[0]); return;
  }
  const a = findAnggota(v);
  if (a) {
    const act = activeOf(a.nis);
    if (!act.length) { beep(false); showRes('er', `<b style="color:#be123c">${esc(a.nama)} tidak memiliki pinjaman aktif.</b>`); focusEl('#r-buku'); return; }
    beep(true); chooseLoans(act, 'Pinjaman aktif ' + a.nama); return;
  }
  beep(false); showRes('er', `<b style="color:#be123c">Kode "${esc(v)}" tidak ditemukan.</b>`); toast('Kode tidak ditemukan.', 'err'); focusEl('#r-buku');
};
async function retOne(p) {
  try {
    const r = await quick(() => doReturn(p)); beep(true);
    showRes('ok', `<div class="nm">${esc(p.judulBuku)}</div><div class="sub">${esc(p.namaSiswa)} \u2022 ${esc(p.kelas)}</div><div style="margin-top:8px"><span class="bdg ${r.late ? 'bad' : 'ok'}">${r.late ? 'Terlambat ' + r.late + ' hari' : 'Tepat waktu'}</span> ${r.denda ? `<span class="bdg warn">Denda Rp ${fmtNum(r.denda)}</span>` : ''}</div>`);
    toast(retMsg(p, r), r.late ? 'info' : 'ok', 4200);
    if (cur === 'pengembalian') views.pengembalian.update();
  } catch (e) { beep(false); toast(e.message || String(e), 'err', 5000); }
  focusEl('#r-buku');
}
ACT.rBukuGo = () => { const i = $('#r-buku'); if (i) ENT.rBuku(i); };
ACT.retChoose = async t => { const p = S.pinjam.find(x => x.id === t.dataset.id); closeModal(); if (p) await retOne(p); };
ACT.retLoan = async t => { const p = S.pinjam.find(x => x.id === t.dataset.id); if (!p) return; t.disabled = true; await retOne(p); };
/* ============ EXCEL (impor / ekspor / template) ============ */
const needXlsx = () => { if (typeof XLSX === 'undefined') { toast('Pustaka Excel belum termuat. Periksa koneksi internet lalu muat ulang.', 'err', 5000); return false; } return true; };
function sheetFrom(aoa, textCols, widths) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const rng = XLSX.utils.decode_range(ws['!ref'] || 'A1:A1');
  for (let R = 1; R <= rng.e.r; R++) for (const C of textCols) {
    const c = ws[XLSX.utils.encode_cell({ r: R, c: C })];
    if (c) { c.t = 's'; c.v = String(c.v == null ? '' : c.v); c.z = '@'; delete c.w; }
  }
  ws['!cols'] = widths.map(w => ({ wch: w }));
  return ws;
}
function blankText(ws, from, to, cols) {
  for (let R = from; R <= to; R++) for (const C of cols) { const a = XLSX.utils.encode_cell({ r: R, c: C }); if (!ws[a]) ws[a] = { t: 's', v: '', z: '@' }; }
  const rng = XLSX.utils.decode_range(ws['!ref']); rng.e.r = Math.max(rng.e.r, to); ws['!ref'] = XLSX.utils.encode_range(rng);
}
function saveWb(wb, name) {
  try {
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob), a = document.createElement('a');
    a.href = url; a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { a.remove(); URL.revokeObjectURL(url); }, 1500);
    toast('File ' + name + ' diunduh.', 'ok');
  } catch (e) { console.error(e); try { XLSX.writeFile(wb, name); } catch (e2) { toast('Gagal membuat file xlsx: ' + (e2.message || e2), 'err', 5000); } }
}
const stamp = () => today();
async function xl(label, fn) {
  if (!needXlsx()) return;
  Prog.show(label, 'Menyusun file Excel\u2026'); await sleep(80);
  try { fn(); } catch (e) { console.error(e); toast('Gagal membuat file: ' + (e.message || e), 'err'); }
  await Prog.hide();
}
ACT.exportAnggota = () => xl('Mengekspor data anggota', () => {
  if (!needXlsx()) return;
  const rows = S.anggota.slice().sort((a, b) => natSort(a.kelas, b.kelas) || natSort(a.nama, b.nama));
  const aoa = [['NIS', 'NISN', 'Nama', 'Kelas', 'Jenis Kelamin', 'Status']].concat(rows.map(a => [txt(a.nis), txt(a.nisn), txt(a.nama), txt(a.kelas), txt(a.jenisKelamin), a.aktif === false ? 'Nonaktif' : 'Aktif']));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, sheetFrom(aoa, [0, 1, 3], [14, 16, 32, 10, 14, 10]), 'Anggota');
  saveWb(wb, 'anggota_' + stamp() + '.xlsx');
});
ACT.exportBuku = () => xl('Mengekspor data buku', () => {
  if (!needXlsx()) return;
  const rows = S.buku.slice().sort((a, b) => natSort(a.kodeBuku, b.kodeBuku));
  const aoa = [['Kode Buku', 'Judul', 'Klasifikasi', 'Pengarang', 'Penerbit', 'Tahun', 'Kategori', 'Rak', 'Jumlah', 'Tersedia']].concat(rows.map(b => [txt(b.kodeBuku), txt(b.judul), txt(b.klasifikasi), txt(b.pengarang), txt(b.penerbit), txt(b.tahun), txt(b.kategori), txt(b.rak), Number(b.jumlahTotal) || 0, availOf(b)]));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, sheetFrom(aoa, [0, 2, 5, 7], [14, 40, 16, 26, 22, 8, 14, 8, 8, 9]), 'Buku');
  saveWb(wb, 'buku_' + stamp() + '.xlsx');
});
ACT.tplAnggota = () => xl('Menyiapkan template anggota', () => {
  if (!needXlsx()) return;
  const aoa = [['NIS', 'NISN', 'Nama', 'Kelas', 'Jenis Kelamin'], ['260101', '0061234501', 'Contoh Siswa Satu', 'VII-A', 'L'], ['260102', '0061234502', 'Contoh Siswa Dua', 'VII-B', 'P']];
  const ws = sheetFrom(aoa, [0, 1, 3], [14, 16, 32, 10, 14]); blankText(ws, 3, 400, [0, 1, 3]);
  const info = XLSX.utils.aoa_to_sheet([['PETUNJUK PENGISIAN'], ['1. Isi data mulai baris ke-2 pada sheet "Anggota". Jangan mengubah judul kolom.'], ['2. NIS, NISN, dan Kelas sudah berformat Teks. Jangan diubah menjadi Angka agar angka 0 di depan tidak hilang.'], ['3. Jenis Kelamin diisi L (laki-laki) atau P (perempuan).'], ['4. NIS wajib unik. Jika NIS sudah ada di aplikasi, datanya diperbarui (bisa dinonaktifkan saat impor).'], ['5. Hapus 2 baris contoh sebelum mengimpor data asli.']]);
  info['!cols'] = [{ wch: 110 }];
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Anggota'); XLSX.utils.book_append_sheet(wb, info, 'Petunjuk');
  saveWb(wb, 'template_anggota.xlsx');
});
ACT.tplBuku = () => xl('Menyiapkan template buku', () => {
  if (!needXlsx()) return;
  const aoa = [['Kode Buku', 'Judul', 'Klasifikasi', 'Pengarang', 'Penerbit', 'Tahun', 'Kategori', 'Rak', 'Jumlah'], ['BK-1001', 'Contoh Judul Buku', '813.54', 'Nama Pengarang', 'Nama Penerbit', '2024', 'Lainnya', 'R-01', 2]];
  const ws = sheetFrom(aoa, [0, 2, 5, 7], [14, 40, 16, 26, 22, 8, 14, 8, 8]); blankText(ws, 2, 600, [0, 2, 5, 7]);
  const info = XLSX.utils.aoa_to_sheet([['PETUNJUK PENGISIAN'], ['1. Isi data mulai baris ke-2 pada sheet "Buku". Jangan mengubah judul kolom.'], ['2. Kode Buku wajib unik, tanpa spasi dan tanda /. Kode ini menjadi isi barcode stiker.'], ['3. Kode Buku, Klasifikasi, Tahun, dan Rak berformat Teks. Jangan diubah menjadi Angka.'], ['4. Jumlah = banyak eksemplar (angka 1 atau lebih). Jika kosong dianggap 1.'], ['5. Hapus baris contoh sebelum mengimpor data asli.']]);
  info['!cols'] = [{ wch: 110 }];
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Buku'); XLSX.utils.book_append_sheet(wb, info, 'Petunjuk');
  saveWb(wb, 'template_buku.xlsx');
});

const ALIAS = {
  gtk: { no: ['noanggota', 'nomoranggota', 'nomoranggotaperpustakaan', 'barcode', 'no'], nip: ['nip'], nama: ['nama', 'namalengkap', 'namagtk', 'namaguru'], jenis: ['jenisptk', 'ptk', 'jenis', 'jenisgtk', 'status'], jk: ['jeniskelamin', 'jk', 'lp', 'gender', 'kelamin'] },
  anggota: { nis: ['nis', 'nomorindukssiswa', 'nomorinduksiswa', 'nomorinduk'], nisn: ['nisn'], nama: ['nama', 'namasiswa', 'namalengkap'], kelas: ['kelas', 'rombel'], jk: ['jeniskelamin', 'jk', 'lp', 'gender', 'kelamin'] },
  buku: { kode: ['kodebuku', 'kode', 'barcode', 'nobuku'], judul: ['judul', 'judulbuku'], klasifikasi: ['klasifikasi', 'noklasifikasi', 'nomorklasifikasi', 'ddc'], pengarang: ['pengarang', 'penulis'], penerbit: ['penerbit'], tahun: ['tahun', 'tahunterbit'], kategori: ['kategori'], rak: ['rak', 'lokasi'], jumlah: ['jumlah', 'jumlahbuku', 'eksemplar', 'stok'] }
};
const REQ = { gtk: ['no', 'nama', 'jenis', 'jk'], anggota: ['nis', 'nama', 'kelas', 'jk'], buku: ['kode', 'judul'] };
const REQ_LABEL = { gtk: 'No. Anggota, Nama, Jenis PTK, Jenis Kelamin', anggota: 'NIS, Nama, Kelas, Jenis Kelamin', buku: 'Kode Buku, Judul' };
const normJk = s => { const v = txt(s).toLowerCase().replace(/[^a-z]/g, ''); if (['l', 'lk', 'laki', 'lakilaki', 'pria', 'male', 'm'].includes(v)) return 'L'; if (['p', 'pr', 'perempuan', 'wanita', 'female', 'f'].includes(v)) return 'P'; return ''; };
const normPtk = s => { const v = txt(s).toLowerCase().replace(/[^a-z]/g, ''); if (['guru', 'gurumapel', 'pendidik'].includes(v)) return 'Guru'; if (['tendik', 'tenagakependidikan', 'tu', 'tatausaha', 'staf', 'staff'].includes(v)) return 'Tendik'; return ''; };
const SCI = /^[+-]?\d+(\.\d+)?e[+-]?\d+$/i;
function parseRows(aoa, type) {
  const def = ALIAS[type]; let hi = -1, map = {};
  for (let i = 0; i < Math.min(aoa.length, 15); i++) {
    const m = {};
    (aoa[i] || []).forEach((h, ci) => { const n = txt(h).toLowerCase().replace(/[^a-z0-9]/g, ''); for (const k in def) if (def[k].includes(n) && !(k in m)) m[k] = ci; });
    if (REQ[type].every(k => k in m)) { hi = i; map = m; break; }
  }
  if (hi < 0) return { error: 'Judul kolom tidak sesuai. Kolom wajib: ' + REQ_LABEL[type] + '. Gunakan tombol Template agar format benar.' };
  const ok = [], bad = [], seen = new Set(), sciMsg = 'Angka tampil dalam format ilmiah (E+). Ubah kolom menjadi Teks lalu isi ulang.';
  const cell = (row, k) => k in map ? txt(row[map[k]]) : '';
  for (let r = hi + 1; r < aoa.length; r++) {
    const row = aoa[r] || []; if (row.every(c => txt(c) === '')) continue;
    const rn = r + 1, errs = [];
    if (type === 'gtk') {
      const no = cell(row, 'no'), nip = cell(row, 'nip'), nama = cell(row, 'nama'), jenis = normPtk(cell(row, 'jenis')), jk = normJk(cell(row, 'jk'));
      if (SCI.test(no) || SCI.test(nip)) errs.push(sciMsg);
      const ke = validKode(no); if (ke) errs.push(ke.replace('Kode', 'No. Anggota'));
      if (!nama) errs.push('Nama kosong.');
      if (!jenis) errs.push('Jenis PTK harus Guru atau Tendik.');
      if (nip && !/^\d+$/.test(nip)) errs.push('NIP harus berisi angka saja.');
      if (!jk) errs.push('Jenis kelamin harus L atau P.');
      if (!ke && seen.has(norm(no))) errs.push('No. Anggota ganda di dalam file.');
      if (errs.length) bad.push({ row: rn, msg: errs.join(' ') }); else { seen.add(norm(no)); ok.push({ row: rn, d: { noAnggota: no, nip, nama, jenisPtk: jenis, jenisKelamin: jk } }); }
    } else if (type === 'anggota') {
      const nis = cell(row, 'nis'), nisn = cell(row, 'nisn'), nama = cell(row, 'nama'), kelas = cell(row, 'kelas'), jk = normJk(cell(row, 'jk'));
      if (SCI.test(nis) || SCI.test(nisn) || SCI.test(kelas)) errs.push(sciMsg);
      const ke = validKode(nis); if (ke) errs.push(ke.replace('Kode', 'NIS'));
      if (!nama) errs.push('Nama kosong.');
      if (!kelas) errs.push('Kelas kosong.');
      if (nisn && !/^\d+$/.test(nisn)) errs.push('NISN harus berisi angka saja.');
      if (!jk) errs.push('Jenis kelamin harus L atau P.');
      if (!ke && seen.has(norm(nis))) errs.push('NIS ganda di dalam file.');
      if (errs.length) bad.push({ row: rn, msg: errs.join(' ') }); else { seen.add(norm(nis)); ok.push({ row: rn, d: { nis, nisn, nama, kelas, jenisKelamin: jk } }); }
    } else {
      const kode = cell(row, 'kode'), judul = cell(row, 'judul'), tahun = cell(row, 'tahun'), jRaw = cell(row, 'jumlah');
      if (SCI.test(kode)) errs.push(sciMsg);
      const ke = validKode(kode); if (ke) errs.push(ke);
      if (!judul) errs.push('Judul kosong.');
      let jumlah = 1;
      if (jRaw !== '') { if (/^\d+$/.test(jRaw) && +jRaw >= 1 && +jRaw <= 9999) jumlah = +jRaw; else errs.push('Jumlah harus angka 1 sampai 9999.'); }
      if (!ke && seen.has(norm(kode))) errs.push('Kode ganda di dalam file.');
      if (errs.length) bad.push({ row: rn, msg: errs.join(' ') }); else { seen.add(norm(kode)); ok.push({ row: rn, d: Object.assign({ kodeBuku: kode, judul, pengarang: cell(row, 'pengarang'), penerbit: cell(row, 'penerbit'), tahun, kategori: cell(row, 'kategori') || 'Lainnya', rak: cell(row, 'rak'), jumlahTotal: jumlah }, 'klasifikasi' in map ? { klasifikasi: cell(row, 'klasifikasi') } : {}) }); }
    }
  }
  return { ok, bad };
}
let IMP = null;
function openImport(type) {
  IMP = { type, res: null };
  const name = type === 'gtk' ? 'Data Anggota GTK' : type === 'anggota' ? 'Data Anggota Siswa' : 'Data Buku';
  openModal(`${mHead('Impor ' + name + ' (.xlsx)')}
<div class="mb">
 <p class="hint">Gunakan file dari tombol <b>Template</b> atau hasil <b>Ekspor</b>. Data diperiksa dulu sebelum disimpan; baris bermasalah dilaporkan dan tidak menghentikan impor.</p>
 <div class="drop"><input type="file" id="imp-file" data-ch="impFile" accept=".xlsx,.xls" aria-label="Pilih file xlsx"></div>
 <label class="chk"><input type="checkbox" id="imp-upd" checked> Perbarui data yang sudah ada (${type === 'gtk' ? 'No. Anggota' : type === 'anggota' ? 'NIS' : 'kode buku'} sama)</label>
 <div id="imp-res" style="margin-top:12px"></div>
</div>
<div class="mf"><button class="btn" data-act="closeModal">Tutup</button><button class="btn" id="imp-err" data-act="impErr" style="display:none">${ic('download', 18)}Unduh daftar error</button><button class="btn pri" id="imp-go" data-act="impGo" disabled>${ic('upload', 18)}Impor Sekarang</button></div>`, { wide: true, noFocus: true });
}
ACT.importAnggota = () => { if (needXlsx()) openImport('anggota'); };
ACT.importBuku = () => { if (needXlsx()) openImport('buku'); };
CHG.impFile = async inp => {
  const f = inp.files && inp.files[0], box = $('#imp-res'); if (!f || !IMP) return;
  IMP.res = null; $('#imp-go').disabled = true; $('#imp-err').style.display = 'none';
  if (!/\.(xlsx|xls)$/i.test(f.name)) { box.innerHTML = '<div class="lerr on">File harus berformat .xlsx atau .xls.</div>'; return; }
  if (f.size > 8 * 1024 * 1024) { box.innerHTML = '<div class="lerr on">Ukuran file maksimal 8 MB.</div>'; return; }
  box.innerHTML = '<p class="sub">Membaca file\u2026</p>';
  try {
    let res = null, firstErr = '';
    await withProg('Membaca file Excel', f.name, async p => {
      const buf = await f.arrayBuffer(); p.set(30, 'Membuka file\u2026'); await sleep(30);
      const wb = XLSX.read(buf, { type: 'array' }); p.set(65, 'Memeriksa isi data\u2026'); await sleep(30);
      for (const sn of wb.SheetNames) {
        const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: false, defval: '', blankrows: false });
        const r = parseRows(aoa, IMP.type);
        if (!r.error) { res = r; break; } else if (!firstErr) firstErr = r.error;
      }
      p.set(100, 'Selesai');
    });
    if (!res) { box.innerHTML = `<div class="lerr on">${esc(firstErr || 'File kosong.')}</div>`; return; }
    if (res.ok.length > 5000) { box.innerHTML = '<div class="lerr on">Maksimal 5.000 baris per impor. Pecah file menjadi beberapa bagian.</div>'; return; }
    IMP.res = res; renderImpRes();
  } catch (e) { console.error(e); box.innerHTML = `<div class="lerr on">File tidak dapat dibaca: ${esc(e.message || e)}. Pastikan file xlsx tidak rusak atau diproteksi sandi.</div>`; }
};
function impStats() {
  const R = IMP.res, upd = $('#imp-upd') && $('#imp-upd').checked;
  let baru = 0, ada = 0;
  R.ok.forEach(x => { const ex = IMP.type === 'gtk' ? !!gtkOf(x.d.noAnggota) : IMP.type === 'anggota' ? S.aMap.has(norm(x.d.nis)) : S.bMap.has(norm(x.d.kodeBuku)); ex ? ada++ : baru++; });
  return { baru, ada, upd, skip: upd ? 0 : ada };
}
function renderImpRes() {
  const R = IMP.res, st = impStats(), box = $('#imp-res');
  box.innerHTML = `<div class="chips"><span class="bdg ok">Baru: ${fmtNum(st.baru)}</span><span class="bdg info">Sudah ada: ${fmtNum(st.ada)}</span><span class="bdg ${R.bad.length ? 'bad' : 'mut'}">Bermasalah: ${fmtNum(R.bad.length)}</span></div>
${R.bad.length ? `<div class="tw" style="max-height:230px"><table class="tbl"><thead><tr><th>Baris</th><th>Masalah</th></tr></thead><tbody>${R.bad.slice(0, 100).map(b => `<tr><td data-l="Baris">${b.row}</td><td data-l="Masalah">${esc(b.msg)}</td></tr>`).join('')}</tbody></table></div>${R.bad.length > 100 ? `<p class="hint" style="margin-top:6px">Menampilkan 100 dari ${R.bad.length} baris bermasalah. Unduh daftar error untuk semuanya.</p>` : ''}` : '<p class="sub">Semua baris valid.</p>'}`;
  $('#imp-go').disabled = !(R.ok.length && (st.baru + (st.upd ? st.ada : 0)) > 0);
  $('#imp-err').style.display = R.bad.length ? '' : 'none';
}
CHG.impUpd = () => { if (IMP && IMP.res) renderImpRes(); };
document.addEventListener('change', e => { if (e.target && e.target.id === 'imp-upd') CHG.impUpd(); });
ACT.impErr = () => {
  if (!IMP || !IMP.res || !needXlsx()) return;
  const aoa = [['Baris', 'Masalah']].concat(IMP.res.bad.map(b => [b.row, b.msg]));
  const ws = XLSX.utils.aoa_to_sheet(aoa); ws['!cols'] = [{ wch: 8 }, { wch: 100 }];
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Error'); saveWb(wb, 'error_impor_' + IMP.type + '.xlsx');
};
ACT.impGo = async t => {
  if (!IMP || !IMP.res) return;
  const R = IMP.res, upd = $('#imp-upd').checked, now = Date.now(), ops = []; let baru = 0, ada = 0, skip = 0;
  t.disabled = true; t.textContent = 'Menyimpan\u2026';
  if (IMP.type === 'gtk') {
    for (const x of R.ok) {
      if (S.anggota.some(a => norm(a.nis) === norm(x.d.noAnggota))) { skip++; continue; } /* nomor sudah dipakai siswa */
      const ex = gtkOf(x.d.noAnggota);
      if (ex) { if (!upd) { skip++; continue; } ada++; ops.push({ t: 'set', c: 'anggota_gtk', id: ex.id, merge: true, d: x.d }); }
      else { baru++; ops.push({ t: 'set', c: 'anggota_gtk', id: x.d.noAnggota, merge: true, d: Object.assign({}, x.d, { aktif: true, dibuatPada: now }) }); }
    }
  } else if (IMP.type === 'anggota') {
    for (const x of R.ok) {
      const ex = S.aMap.get(norm(x.d.nis));
      if (ex && ex.tipe === 'gtk') { skip++; continue; } /* NIS bentrok dengan No. Anggota GTK */
      if (ex) { if (!upd) { skip++; continue; } ada++; ops.push({ t: 'set', c: 'anggota', id: ex.id, merge: true, d: x.d }); }
      else { baru++; ops.push({ t: 'set', c: 'anggota', id: x.d.nis, merge: true, d: Object.assign({}, x.d, { aktif: true, dibuatPada: now }) }); }
    }
  } else {
    for (const x of R.ok) {
      const ex = S.bMap.get(norm(x.d.kodeBuku));
      if (ex) {
        if (!upd) { skip++; continue; }
        const dipinjam = activeOfBook(ex.kodeBuku).length;
        if (x.d.jumlahTotal < dipinjam) { skip++; continue; }
        ada++; ops.push({ t: 'set', c: 'buku', id: ex.id, merge: true, d: Object.assign({}, x.d, { jumlahTersedia: x.d.jumlahTotal - dipinjam }) });
      } else { baru++; ops.push({ t: 'set', c: 'buku', id: x.d.kodeBuku, merge: true, d: Object.assign({}, x.d, { jumlahTersedia: x.d.jumlahTotal, dibuatPada: now }) }); }
    }
  }
  try {
    if (baru + ada) ops.push(actOp(IMP.type === 'buku' ? 'buku' : 'anggota', 'Impor ' + ((IMP.type === 'gtk' ? 'anggota GTK' : IMP.type === 'anggota' ? 'anggota siswa' : 'buku')), (baru + ada) + ' data'));
    await withProg('Mengimpor data ' + ((IMP.type === 'gtk' ? 'anggota GTK' : IMP.type === 'anggota' ? 'anggota siswa' : 'buku')), 'Menyimpan ' + fmtNum(ops.length) + ' data\u2026', p => commit(ops, (d, tt) => p.set(d / tt * 100, 'Menyimpan ' + fmtNum(d) + ' dari ' + fmtNum(tt) + ' data')));
    closeModal(); beep(true);
    toast(`Impor selesai: ${fmtNum(baru)} baru, ${fmtNum(ada)} diperbarui${skip ? ', ' + fmtNum(skip) + ' dilewati' : ''}.`, 'ok', 5000);
  } catch (e) { console.error(e); t.disabled = false; t.innerHTML = ic('upload', 18) + 'Impor Sekarang'; toast('Impor gagal: ' + (e.code || e.message || e) + '. Sebagian data mungkin sudah tersimpan; ulangi impor untuk melengkapi.', 'err', 7000); }
};

/* ============ LAPORAN ============ */
const Rpt = { jenis: 'peminjaman', from: addDays(today(), -29), to: today() };
const JENIS = { peminjaman: 'Laporan Peminjaman', pengembalian: 'Laporan Pengembalian', terlambat: 'Laporan Keterlambatan', populer: 'Buku Terpopuler', anggota: 'Anggota Paling Aktif' };
function reportData() {
  const f = Rpt.from, t = Rpt.to, now = today(), P = S.pinjam;
  const stat = p => p.status === 'kembali' ? 'Kembali' : (diffDays(p.jatuhTempo, now) > 0 ? 'Terlambat' : 'Dipinjam');
  let head, rows, text = [];
  if (Rpt.jenis === 'peminjaman') {
    head = ['Tanggal', 'NIS / No. Anggota', 'Nama', 'Kelas / PTK', 'Kode Buku', 'Judul', 'Jatuh Tempo', 'Status']; text = [1, 3, 4];
    rows = P.filter(p => p.tglPinjam >= f && p.tglPinjam <= t).sort((a, b) => (b.tsPinjam || 0) - (a.tsPinjam || 0)).map(p => [fmtDate(p.tglPinjam), p.nis, p.namaSiswa, p.kelas, p.kodeBuku, p.judulBuku, fmtDate(p.jatuhTempo), stat(p)]);
  } else if (Rpt.jenis === 'pengembalian') {
    head = ['Tgl Kembali', 'NIS / No. Anggota', 'Nama', 'Kelas / PTK', 'Kode Buku', 'Judul', 'Tgl Pinjam', 'Terlambat (hari)', 'Denda (Rp)']; text = [1, 3, 4];
    rows = P.filter(p => p.status === 'kembali' && p.tglKembali >= f && p.tglKembali <= t).sort((a, b) => (b.tsKembali || 0) - (a.tsKembali || 0)).map(p => [fmtDate(p.tglKembali), p.nis, p.namaSiswa, p.kelas, p.kodeBuku, p.judulBuku, fmtDate(p.tglPinjam), Number(p.terlambat) || 0, Number(p.denda) || 0]);
  } else if (Rpt.jenis === 'terlambat') {
    head = ['Jatuh Tempo', 'NIS / No. Anggota', 'Nama', 'Kelas / PTK', 'Kode Buku', 'Judul', 'Tgl Kembali', 'Hari Terlambat', 'Denda (Rp)']; text = [1, 3, 4];
    rows = P.filter(p => p.jatuhTempo >= f && p.jatuhTempo <= t && diffDays(p.jatuhTempo, p.status === 'kembali' ? p.tglKembali : now) > 0).sort((a, b) => (a.jatuhTempo < b.jatuhTempo ? 1 : -1)).map(p => { const l = diffDays(p.jatuhTempo, p.status === 'kembali' ? p.tglKembali : now); return [fmtDate(p.jatuhTempo), p.nis, p.namaSiswa, p.kelas, p.kodeBuku, p.judulBuku, p.status === 'kembali' ? fmtDate(p.tglKembali) : 'Belum kembali', l, l * (Number(S.set.dendaPerHari) || 0)]; });
  } else if (Rpt.jenis === 'populer') {
    head = ['Peringkat', 'Kode Buku', 'Judul', 'Pengarang', 'Kali Dipinjam']; text = [1];
    const m = new Map(); P.filter(p => p.tglPinjam >= f && p.tglPinjam <= t).forEach(p => { const k = norm(p.kodeBuku); if (!m.has(k)) m.set(k, { kode: p.kodeBuku, judul: p.judulBuku, n: 0 }); m.get(k).n++; });
    rows = Array.from(m.values()).sort((a, b) => b.n - a.n).map((x, i) => { const b = findBuku(x.kode); return [i + 1, x.kode, (b && b.judul) || x.judul, (b && b.pengarang) || '', x.n]; });
  } else {
    head = ['Peringkat', 'NIS / No. Anggota', 'Nama', 'Kelas / PTK', 'Jumlah Pinjam']; text = [1, 3];
    const m = new Map(); P.filter(p => p.tglPinjam >= f && p.tglPinjam <= t).forEach(p => { const k = norm(p.nis); if (!m.has(k)) m.set(k, { nis: p.nis, nama: p.namaSiswa, kelas: p.kelas, n: 0 }); m.get(k).n++; });
    rows = Array.from(m.values()).sort((a, b) => b.n - a.n).map((x, i) => [i + 1, x.nis, x.nama, x.kelas, x.n]);
  }
  return { head, rows, text, title: JENIS[Rpt.jenis] + ' (' + fmtDate(f) + ' s/d ' + fmtDate(t) + ')' };
}
views.laporan = {
  mount(el) {
    el.innerHTML = `
<div class="vh"><div><h2>Laporan</h2><p class="sub">Rekap peminjaman, pengembalian, keterlambatan, dan statistik. Dapat diekspor ke xlsx atau dicetak.</p></div></div>
<div class="tools">
  <select class="in" data-ch="rjenis" aria-label="Jenis laporan">${Object.keys(JENIS).map(k => `<option value="${k}" ${k === Rpt.jenis ? 'selected' : ''}>${JENIS[k]}</option>`).join('')}</select>
  <div class="fld" style="flex-direction:row;align-items:center;gap:8px"><label for="r-from">Dari</label><input id="r-from" class="in" type="date" data-ch="rfrom" value="${Rpt.from}" style="width:auto"></div>
  <div class="fld" style="flex-direction:row;align-items:center;gap:8px"><label for="r-to">Sampai</label><input id="r-to" class="in" type="date" data-ch="rto" value="${Rpt.to}" style="width:auto"></div>
</div>
<div class="tools"><button class="btn pri" data-act="rptExport">${ic('download', 18)}Ekspor xlsx</button><button class="btn" data-act="rptPrint">${ic('print', 18)}Cetak</button></div>
<div id="l-zone"></div>`;
  },
  update() {
    const d = reportData(), shown = d.rows.slice(0, 300);
    $('#l-zone').innerHTML = `<div class="chips"><span class="bdg info">${esc(d.title)}</span><span class="bdg mut">${fmtNum(d.rows.length)} baris</span></div>` +
      (d.rows.length ? `<div class="tw"><table class="tbl"><thead><tr>${d.head.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${shown.map(r => `<tr>${r.map((c, i) => `<td data-l="${esc(d.head[i])}">${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>${d.rows.length > 300 ? `<p class="hint" style="margin-top:8px">Menampilkan 300 baris pertama. Ekspor xlsx untuk semua data.</p>` : ''}` : `<div class="tw">${emptyHtml('chart', 'Tidak ada data pada rentang tanggal ini.')}</div>`);
  }
};
CHG.rjenis = t => { Rpt.jenis = t.value; views.laporan.update(); };
CHG.rfrom = t => { if (t.value) { Rpt.from = t.value; if (Rpt.to < Rpt.from) Rpt.to = Rpt.from; const o = $('#r-to'); if (o) o.value = Rpt.to; views.laporan.update(); } };
CHG.rto = t => { if (t.value) { Rpt.to = t.value; if (Rpt.from > Rpt.to) Rpt.from = Rpt.to; const o = $('#r-from'); if (o) o.value = Rpt.from; views.laporan.update(); } };
ACT.rptExport = () => xl('Mengekspor laporan', () => {
  if (!needXlsx()) return;
  const d = reportData(), ws = sheetFrom([d.head].concat(d.rows), d.text, d.head.map(h => Math.max(12, h.length + 4)));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Laporan'); saveWb(wb, 'laporan_' + Rpt.jenis + '_' + stamp() + '.xlsx');
});
ACT.rptPrint = async () => {
  const d = reportData();
  await doPrint(`<div class="rp"><h2>${esc(S.set.namaAplikasi)} - ${esc(S.set.namaSekolah)}</h2><p>${esc(d.title)} \u2022 Tahun Ajaran ${esc(S.set.tahunAjaran)}</p><table><thead><tr>${d.head.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${d.rows.map(r => `<tr>${r.map(c => `<td>${esc(c)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`, d.title);
};

/* ============ PENGATURAN ============ */
views.pengaturan = {
  mount(el) {
    const s = S.set;
    el.innerHTML = `
<div class="vh"><div><h2>Pengaturan</h2><p class="sub">Ubah identitas aplikasi dan aturan peminjaman. Perubahan langsung berlaku.</p></div></div>
<form class="card pn" data-sub="setSave" autocomplete="off" style="max-width:860px">
 <div class="grid2">
  <div class="fld"><label>Nama aplikasi*</label><input class="in" name="namaAplikasi" value="${esc(s.namaAplikasi)}" required maxlength="60"></div>
  <div class="fld"><label>Nama sekolah*</label><input class="in" name="namaSekolah" value="${esc(s.namaSekolah)}" required maxlength="80"></div>
  <div class="fld"><label>Tahun ajaran*</label><input class="in" name="tahunAjaran" value="${esc(s.tahunAjaran)}" required maxlength="20" placeholder="2026/2027"></div>
  <div class="fld"><label>Semboyan sekolah</label><input class="in" name="semboyan" value="${esc(s.semboyan)}" maxlength="80"></div>
  <div class="fld full"><label>URL logo</label><input class="in" name="logoUrl" value="${esc(s.logoUrl)}" maxlength="500" inputmode="url"><small>Kosongkan untuk memakai logo bawaan.</small></div>
  <div class="fld"><label>Lama pinjam (hari)</label><input class="in" name="lamaPinjamHari" type="number" min="1" max="60" value="${esc(s.lamaPinjamHari)}"></div>
  <div class="fld"><label>Maksimal buku per siswa</label><input class="in" name="maksPinjam" type="number" min="1" max="20" value="${esc(s.maksPinjam)}"></div>
  <div class="fld"><label>Denda per hari terlambat (Rp)</label><input class="in" name="dendaPerHari" type="number" min="0" max="1000000" step="100" value="${esc(s.dendaPerHari)}"><small>Isi 0 jika tidak ada denda.</small></div>
  <div class="fld full"><label>Kategori buku (pisahkan dengan koma)</label><textarea class="in" name="kategori" rows="2">${esc((s.kategori || []).join(', '))}</textarea></div>
 </div>
 <div style="display:flex;gap:10px;margin-top:16px;flex-wrap:wrap"><button class="btn pri" type="submit">${ic('check', 18)}Simpan Pengaturan</button></div>
</form>
<div class="card pn" style="max-width:860px;margin-top:14px">
 <div class="ph"><h3>${ic('trash', 18)} Hapus Riwayat</h3></div>
 <p class="hint">Hapus riwayat untuk merapikan data. Peminjaman yang belum dikembalikan tidak pernah ikut terhapus. Peminjaman dan pengembalian tersimpan dalam satu catatan, jadi menghapus salah satunya juga menghapus catatan lengkapnya. Data yang dihapus tidak bisa dikembalikan, sebaiknya ekspor laporan xlsx lebih dulu.</p>
 ${[['aktivitas', 'Riwayat aktivitas', 'Log pada kartu Aktivitas Terbaru'], ['peminjaman', 'Riwayat peminjaman', 'Berdasarkan tanggal pinjam'], ['pengembalian', 'Riwayat pengembalian', 'Berdasarkan tanggal kembali']].map(h => `<div class="hrow"><div><b>${h[1]}</b><span>${h[2]}</span></div><select class="in" id="h-${h[0]}" aria-label="Periode ${h[1]}" style="width:100%"><option value="m0">Bulan ini</option><option value="m3">3 bulan terakhir</option><option value="m6">6 bulan terakhir</option></select><button class="btn dng" data-act="histDel" data-v="${h[0]}">${ic('trash', 16)}Hapus</button></div>`).join('')}
</div>`;
  },
  update() { }
};
SUB.setSave = async f => {
  const fd = new FormData(f), g = k => txt(fd.get(k));
  const lama = parseInt(g('lamaPinjamHari'), 10), maks = parseInt(g('maksPinjam'), 10), denda = Math.round(Number(g('dendaPerHari') || 0));
  const kat = Array.from(new Set(g('kategori').split(/[,\n]/).map(x => txt(x)).filter(Boolean)));
  let logo = g('logoUrl'), err = '';
  if (!g('namaAplikasi') || !g('namaSekolah') || !g('tahunAjaran')) err = 'Nama aplikasi, nama sekolah, dan tahun ajaran wajib diisi.';
  else if (!(lama >= 1 && lama <= 60)) err = 'Lama pinjam harus 1 sampai 60 hari.';
  else if (!(maks >= 1 && maks <= 20)) err = 'Maksimal buku per siswa harus 1 sampai 20.';
  else if (!(denda >= 0)) err = 'Denda tidak boleh negatif.';
  else if (!kat.length) err = 'Isi minimal satu kategori buku.';
  else if (logo && !/^(https?:\/\/|data:image\/)/i.test(logo)) err = 'URL logo harus diawali http:// atau https://.';
  if (err) { toast(err, 'err'); return; }
  if (!logo) logo = LOGO;
  const btn = $('button[type=submit]', f); btn.disabled = true;
  try {
    await withProg('Menyimpan pengaturan', 'Memperbarui identitas aplikasi\u2026', () => commit([{ t: 'set', c: 'pengaturan', id: 'umum', merge: true, d: { namaAplikasi: g('namaAplikasi'), namaSekolah: g('namaSekolah'), tahunAjaran: g('tahunAjaran'), semboyan: g('semboyan'), logoUrl: logo, lamaPinjamHari: lama, maksPinjam: maks, dendaPerHari: denda, kategori: kat } }]));
    toast('Pengaturan berhasil disimpan.', 'ok');
  } catch (e) { toast('Gagal menyimpan: ' + (e.code || e.message || e), 'err', 5000); }
  btn.disabled = false;
};
const HIST = { aktivitas: ['riwayat aktivitas', 'aktivitas'], peminjaman: ['riwayat peminjaman', 'peminjaman'], pengembalian: ['riwayat pengembalian', 'pengembalian'] };
const PER_LBL = { m0: 'bulan ini', m3: '3 bulan terakhir', m6: '6 bulan terakhir' };
function histStart(per) {
  const [y, m, d] = today().split('-').map(Number);
  if (per === 'm0') return ymd(new Date(y, m - 1, 1));
  return ymd(new Date(y, m - 1 - (per === 'm3' ? 3 : 6), d));
}
async function histFind(jenis, per) {
  const start = histStart(per), [y, m, d] = start.split('-').map(Number), startMs = new Date(y, m - 1, d).getTime();
  let docs;
  if (IS_PREVIEW) {
    docs = Array.from(mem[jenis === 'aktivitas' ? 'aktivitas' : 'peminjaman'].values());
  } else {
    const col = jenis === 'aktivitas' ? 'aktivitas' : 'peminjaman';
    const fld = jenis === 'aktivitas' ? 'waktu' : (jenis === 'peminjaman' ? 'tglPinjam' : 'tglKembali');
    const snap = await fb.getDocs(fb.query(fb.collection(fdb, col), fb.where(fld, '>=', jenis === 'aktivitas' ? startMs : start)));
    docs = snap.docs.map(x => Object.assign({ id: x.id }, x.data()));
  }
  if (jenis === 'aktivitas') return { ids: docs.filter(a => (a.waktu || 0) >= startMs).map(a => a.id), skip: 0, start };
  if (jenis === 'peminjaman') {
    const inRange = docs.filter(p => p.tglPinjam >= start);
    return { ids: inRange.filter(p => p.status === 'kembali').map(p => p.id), skip: inRange.filter(p => p.status !== 'kembali').length, start };
  }
  return { ids: docs.filter(p => p.status === 'kembali' && p.tglKembali >= start).map(p => p.id), skip: 0, start };
}
ACT.histDel = async t => {
  const jenis = t.dataset.v, per = $('#h-' + jenis).value, lbl = HIST[jenis][0], pl = PER_LBL[per];
  t.disabled = true;
  try {
    const r = await withProg('Memeriksa data', 'Menghitung ' + lbl + ' ' + pl + '\u2026', () => histFind(jenis, per));
    if (!r.ids.length) { toast('Tidak ada ' + lbl + ' untuk ' + pl + (r.skip ? ' (' + fmtNum(r.skip) + ' peminjaman belum dikembalikan dilewati)' : '') + '.', 'info'); return; }
    const note = r.skip ? ' ' + fmtNum(r.skip) + ' peminjaman yang belum dikembalikan tidak ikut dihapus.' : '';
    if (!await confirmBox('Hapus ' + fmtNum(r.ids.length) + ' data ' + lbl + ' ' + pl + ' (sejak ' + fmtDate(r.start) + ')? Data tidak bisa dikembalikan.' + note, 'Ya, hapus', true)) return;
    const col = jenis === 'aktivitas' ? 'aktivitas' : 'peminjaman', ops = r.ids.map(id => ({ t: 'del', c: col, id }));
    await withProg('Menghapus ' + lbl, fmtNum(ops.length) + ' data\u2026', p => commit(ops, (dn, tt) => p.set(dn / tt * 100, fmtNum(dn) + ' dari ' + fmtNum(tt) + ' data')));
    toast(fmtNum(ops.length) + ' data ' + lbl + ' ' + pl + ' dihapus.', 'ok', 4500);
  } catch (e) { console.error(e); toast('Gagal menghapus: ' + (e.code || e.message || e), 'err', 6500); }
  finally { t.disabled = false; }
};
/* ============ AKSI UMUM ============ */
ACT.go = t => go(t.dataset.v);
ACT.closeModal = () => closeModal();
ACT.cfYes = () => { const r = cfResolve; cfResolve = null; closeModal(); if (r) r(true); };
ACT.cfNo = () => closeModal();
ACT.toggleSide = () => {
  if (window.matchMedia('(max-width:860px)').matches) { ACT.more(); return; }
  $('#lib-app').classList.toggle('collapsed');
};
ACT.more = () => {
  const items = [['anggota', 'Anggota Siswa', 'users'], ['gtk', 'Anggota GTK', 'briefcase'], ['laporan', 'Laporan', 'chart'], ['pengaturan', 'Pengaturan', 'gear']];
  openModal(`${mHead('Menu')}<div class="mb mlist">${items.map(([id, l, i]) => `<button data-act="go" data-v="${id}">${ic(i, 22)}<div><b>${l}</b></div></button>`).join('')}<button data-act="logout">${ic('logout', 22)}<div><b>Keluar</b></div></button></div>`, { noFocus: true });
};
ACT.userMenu = () => $('#lib-umenu').classList.toggle('on');
ACT.togglePw = t => { const i = $('#lg-pw'); const show = i.type === 'password'; i.type = show ? 'text' : 'password'; t.textContent = show ? 'Sembunyikan' : 'Tampilkan'; };
let deferredPrompt = null;
window.addEventListener('beforeinstallprompt', e => { e.preventDefault(); deferredPrompt = e; const b = $('#lib-install'); if (b) b.style.display = ''; });
ACT.installApp = async () => { if (!deferredPrompt) return; deferredPrompt.prompt(); try { await deferredPrompt.userChoice; } catch (e) { /* abaikan */ } deferredPrompt = null; $('#lib-install').style.display = 'none'; $('#lib-umenu').classList.remove('on'); };

/* ============ LOGIN / LOGOUT ============ */
const loginErr = m => { const e = $('#lib-lerr'); e.textContent = m; e.classList.toggle('on', !!m); };
const loginMsg = e => {
  const c = (e && e.code) || '';
  if (/invalid-credential|wrong-password|user-not-found|invalid-email|missing-password/.test(c)) return 'Email atau kata sandi salah.';
  if (/too-many-requests/.test(c)) return 'Terlalu banyak percobaan. Coba lagi beberapa menit lagi.';
  if (/network-request-failed/.test(c)) return 'Tidak ada koneksi internet.';
  if (/api-key|configuration|app-not-authorized/.test(c)) return 'Konfigurasi Firebase belum benar. Periksa Environment Variables VITE_FIREBASE_*.';
  return 'Gagal masuk' + (c ? ' (' + c + ')' : '') + '.';
};
SUB.login = async f => {
  const email = txt(f.email.value), pw = f.pw.value, btn = $('#lg-btn');
  loginErr('');
  if (!email || !pw) { loginErr('Isi email dan kata sandi.'); return; }
  btn.disabled = true; btn.textContent = 'Memeriksa\u2026';
  try {
    if (IS_PREVIEW) await startApp(email);
    else await fb.signInWithEmailAndPassword(fauth, email, pw);
  } catch (e) { console.error(e); loginErr(loginMsg(e)); }
  btn.disabled = false; btn.textContent = 'Masuk';
};
let starting = false;
async function startApp(email) {
  if (appOn || starting) return;
  starting = true;
  $('#lib-boot-t').textContent = 'Menyiapkan data\u2026';
  $('#lib-boot').classList.remove('off');
  $('#lib-login').classList.remove('on');
  S.user = email; $('#lib-email').textContent = email || 'Admin';
  try { await autoSetupDatabase(); }
  catch (e) { console.error(e); if (e.code === 'permission-denied') showDenied(); else toast('Setup database gagal: ' + (e.code || e.message) + '. Periksa Security Rules dan akun admin.', 'err', 8000); }
  try { if (IS_PREVIEW) memSync(); else await watchAll(); }
  catch (e) { console.error(e); toast('Gagal memuat data: ' + (e.code || e.message), 'err', 6000); }
  rebuild(); applySettings();
  appOn = true; starting = false;
  $('#lib-app').classList.add('on');
  $('#lib-boot').classList.add('off');
  const id = location.hash.replace(/^#\/?/, '');
  go(views[id] ? id : 'beranda');
}
function stopApp() {
  deniedOn = false; ACT.denyClose();
  stopWatch(); appOn = false; starting = false;
  Object.assign(S, { anggota: [], gtk: [], buku: [], pinjam: [], akt: [], user: null });
  rebuild(); Pm.nis = null; Pm.items = []; Rt.last = [];
  SEL.buku.clear(); SEL.anggota.clear(); SEL.gtk.clear();
  if (views[cur] && views[cur].unmount) views[cur].unmount();
  closeModal();
  $('#lib-app').classList.remove('on'); $('#lib-umenu').classList.remove('on');
  $('#lib-boot').classList.add('off'); $('#lib-login').classList.add('on');
  const pw = $('#lg-pw'); if (pw) pw.value = '';
}
ACT.logout = async () => {
  try { closeModal(); if (!IS_PREVIEW) await fb.signOut(fauth); } catch (e) { console.error(e); }
  if (IS_PREVIEW) stopApp();
};
function onUser(u) { if (u) startApp(u.email); else stopApp(); }

/* ============ PWA (manifest.webmanifest dan sw.js ada di folder public) ============ */
function setupPWA() {
  if (!('serviceWorker' in navigator) || !import.meta.env.PROD) return;
  window.addEventListener('load', () => { navigator.serviceWorker.register('/sw.js').catch(e => console.warn('Service worker gagal didaftarkan', e)); });
}

/* ============ ANGGOTA GTK ============ */
setupGtk({ S, Q, SEL, views, ACT, INP, CHG, SUB, PER, $, esc, txt, norm, ic, natSort, emptyHtml, pagerHtml, mHead, openModal, closeModal, confirmBox, toast, withProg, quick, commit, actOp, activeOf, fmtNum, validKode, doPrint, LOGO, openImport, needXlsx, xl, sheetFrom, blankText, saveWb, stamp });

/* ============ EVENT & BOOT ============ */
function bindEvents() {
  const root = $('#lib-root');
  root.addEventListener('click', e => {
    const t = e.target.closest('[data-act]');
    if (t && root.contains(t)) { const f = ACT[t.dataset.act]; if (f) { e.preventDefault(); f(t, e); } }
    else if (e.target.id === 'lib-modal') closeModal();
    if (!e.target.closest('.user')) $('#lib-umenu').classList.remove('on');
  });
  root.addEventListener('input', e => { const t = e.target.closest('[data-in]'); if (t && INP[t.dataset.in]) INP[t.dataset.in](t, e); });
  root.addEventListener('change', e => { const t = e.target.closest('[data-ch]'); if (t && CHG[t.dataset.ch]) CHG[t.dataset.ch](t, e); });
  root.addEventListener('keydown', e => {
    if (e.key === 'Enter') { const t = e.target.closest('[data-enter]'); if (t && ENT[t.dataset.enter]) { e.preventDefault(); ENT[t.dataset.enter](t, e); } }
    if (e.key === 'Escape') closeModal();
  });
  root.addEventListener('submit', e => { const f = e.target.closest('[data-sub]'); if (f && SUB[f.dataset.sub]) { e.preventDefault(); SUB[f.dataset.sub](f, e); } });
  window.addEventListener('hashchange', () => { const id = location.hash.replace(/^#\/?/, ''); if (appOn && views[id] && id !== cur) go(id, true); });
}
(async function boot() {
  document.documentElement.style.overflow = 'hidden';
  setupPWA(); hydrateIcons(); renderNav(); applySettings(); bindEvents();
  const showLogin = () => { $('#lib-boot').classList.add('off'); $('#lib-login').classList.add('on'); setTimeout(() => { const e = $('#lg-email'); if (e) e.focus(); }, 150); };
  if (IS_PREVIEW) { $('#lib-lpv').style.display = ''; showLogin(); return; }
  if (!FIREBASE_READY) { showLogin(); loginErr('Konfigurasi Firebase belum lengkap. Isi Environment Variables VITE_FIREBASE_* lalu deploy ulang.'); return; }
  try {
    await fbInit();
    let first = true;
    fb.onAuthStateChanged(fauth, u => { if (first) { first = false; if (!u) showLogin(); } onUser(u); });
  } catch (e) { console.error(e); showLogin(); loginErr('Firebase gagal dimuat. Periksa koneksi internet dan firebaseConfig.'); }
})();
