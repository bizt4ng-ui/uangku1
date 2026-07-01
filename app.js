// ============================================================
// UANGKU - GitHub Pages + Google Sheets API v4
// ============================================================

// ── KONFIGURASI ── Isi sesuai project Google Cloud Anda
const CLIENT_ID   = '641912236870-2aqj6qdrbflv0q6tb7dlpcotsak24i4j.apps.googleusercontent.com';
const API_KEY     = 'AIzaSyAArFoRAWBrXppd887w2DkinIZnwo313Ik';
const SCOPES      = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file profile email';
const SHEET_TITLE = 'UANGKU Database';

// Nama sheet tab
const TAB = { TRX: 'Transaksi', KAT: 'Kategori', BUDGET: 'Budget', CFG: 'Config' };

// ── STATE ──
let tokenClient, accessToken = null;
let spreadsheetId = null;
let kategoriData  = [];
let bulanAktif    = '';
let currentPage   = 'dashboard';
let currentJenis  = 'pemasukan';
let filterKatMode = 'semua';
let chartDonut    = null;
let chartBar      = null;

// ── COOKIE HELPER ──
function setCookie(name, val, days) {
  const exp = new Date(Date.now() + days*864e5).toUTCString();
  document.cookie = name + '=' + encodeURIComponent(val) + ';expires=' + exp + ';path=/;SameSite=Lax';
}
function getCookie(name) {
  const m = document.cookie.match('(?:^|; )' + name + '=([^;]*)');
  return m ? decodeURIComponent(m[1]) : null;
}
function delCookie(name) {
  document.cookie = name + '=;expires=Thu, 01 Jan 1970 00:00:00 GMT;path=/';
}

// ── INIT ──
window.onload = function() {
  const now = new Date();
  bulanAktif = now.getFullYear() + '-' + String(now.getMonth()+1).padStart(2,'0');
  initBulanFilter();

  // Load Google API
  gapi.load('client', async function() {
    await gapi.client.init({ apiKey: API_KEY, discoveryDocs: ['https://sheets.googleapis.com/$discovery/rest?version=v4'] });

    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID,
      scope: SCOPES,
      prompt: '',
      callback: async (resp) => {
        if (resp.error) { showToast('Login gagal: ' + resp.error, 'error'); return; }
        accessToken = resp.access_token;
        setCookie('uangku_token', accessToken, 1); // simpan 1 hari
        gapi.client.setToken({ access_token: accessToken });
        await masukAplikasi();
      }
    });

    // Cek token dari cookie
    const saved = getCookie('uangku_token');
    if (saved) {
      accessToken = saved;
      gapi.client.setToken({ access_token: saved });
      // Verifikasi token masih valid
      try {
        const test = await fetch('https://www.googleapis.com/oauth2/v3/tokeninfo?access_token=' + saved);
        const info = await test.json();
        if (info.error) throw new Error('expired');
        await masukAplikasi();
      } catch(e) {
        // Token expired, minta login ulang tanpa prompt
        delCookie('uangku_token');
        tokenClient.requestAccessToken({ prompt: '' });
      }
    }
    // Jika tidak ada token, tampilkan tombol login
  });
};

async function loginGoogle() {
  if (!tokenClient) { showToast('Google API belum siap, tunggu sebentar...', 'error'); return; }
  tokenClient.requestAccessToken({ prompt: 'select_account' });
}

async function masukAplikasi() {
  try {
    // Ambil info user
    const resp = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { Authorization: 'Bearer ' + accessToken }
    });
    if (!resp.ok) throw new Error('Token expired');
    const user = await resp.json();
    document.getElementById('userAvatar').src = user.picture || '';
    document.getElementById('userName').textContent = user.name || user.email;

    // Tampilkan UI
    document.getElementById('login-screen').style.display  = 'none';
    document.getElementById('topnav').style.display        = 'flex';
    document.getElementById('sidebar').style.display       = 'block';
    document.getElementById('app-layout').style.display    = 'flex';
    document.getElementById('bottom-nav').style.display    = 'flex';

    // Init spreadsheet
    await initSpreadsheet();
    await loadKategori();
    showPage('dashboard');
    setTanggalHariIni();

    setInterval(() => { if (currentPage === 'dashboard') loadDashboard(); }, 60000);
  } catch(e) {
    localStorage.removeItem('uangku_token');
    accessToken = null;
    document.getElementById('login-screen').style.display = 'flex';
  }
}

function logout() {
  if (!confirm('Keluar dari UANGKU?')) return;
  delCookie('uangku_token');
  delCookie('uangku_sheet_id');
  localStorage.removeItem('uangku_sheet_id');
  if (accessToken) google.accounts.oauth2.revoke(accessToken);
  location.reload();
}

// ── SPREADSHEET INIT ──
async function initSpreadsheet() {
  // Cek apakah sudah ada ID tersimpan
  spreadsheetId = getCookie('uangku_sheet_id') || localStorage.getItem('uangku_sheet_id');
  if (spreadsheetId) {
    document.getElementById('cfg-sheet-id').value = spreadsheetId;
    return;
  }

  // Cari spreadsheet UANGKU di Drive
  const search = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=name='${SHEET_TITLE}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false&fields=files(id,name)`,
    { headers: { Authorization: 'Bearer ' + accessToken } }
  );
  const result = await search.json();

  if (result.files && result.files.length > 0) {
    spreadsheetId = result.files[0].id;
  } else {
    // Buat spreadsheet baru
    spreadsheetId = await buatSpreadsheetBaru();
  }

  localStorage.setItem('uangku_sheet_id', spreadsheetId);
  setCookie('uangku_sheet_id', spreadsheetId, 365);
  document.getElementById('cfg-sheet-id').value = spreadsheetId;
}

async function buatSpreadsheetBaru() {
  const resp = await gapi.client.sheets.spreadsheets.create({
    resource: {
      properties: { title: SHEET_TITLE },
      sheets: [
        { properties: { title: TAB.TRX } },
        { properties: { title: TAB.KAT } },
        { properties: { title: TAB.BUDGET } },
        { properties: { title: TAB.CFG } }
      ]
    }
  });
  const id = resp.result.spreadsheetId;

  // Header Transaksi
  await sheetsWrite(TAB.TRX + '!A1:H1', [['ID','Tanggal','Jenis','Kategori','Keterangan','Jumlah','Saldo','Timestamp']]);
  await sheetsWrite(TAB.KAT + '!A1:D1', [['Nama','Jenis','Warna','Icon']]);
  await sheetsWrite(TAB.BUDGET + '!A1:D1', [['Bulan','Kategori','Anggaran','Terpakai']]);
  await sheetsWrite(TAB.CFG + '!A1:B1', [['Key','Value']]);

  // Seed kategori default
  const kats = [
    ['Gaji','pemasukan','#22c55e','ti-briefcase'],
    ['Bonus','pemasukan','#16a34a','ti-star'],
    ['Investasi','pemasukan','#059669','ti-trending-up'],
    ['Lain-lain Masuk','pemasukan','#10b981','ti-plus-circle'],
    ['Makan & Minum','pengeluaran','#ef4444','ti-tools-kitchen-2'],
    ['Transportasi','pengeluaran','#f97316','ti-car'],
    ['Belanja','pengeluaran','#8b5cf6','ti-shopping-cart'],
    ['Tagihan','pengeluaran','#3b82f6','ti-receipt'],
    ['Kesehatan','pengeluaran','#06b6d4','ti-heart'],
    ['Pendidikan','pengeluaran','#f59e0b','ti-school'],
    ['Hiburan','pengeluaran','#ec4899','ti-device-gamepad'],
    ['Tabungan','pengeluaran','#64748b','ti-piggy-bank'],
    ['Lain-lain Keluar','pengeluaran','#94a3b8','ti-dots-circle'],
  ];
  spreadsheetId = id;
  await sheetsAppend(TAB.KAT, kats);
  return id;
}

// ── SHEETS API HELPERS ──
async function sheetsRead(range) {
  const resp = await gapi.client.sheets.spreadsheets.values.get({ spreadsheetId, range });
  return resp.result.values || [];
}

async function sheetsWrite(range, values) {
  await gapi.client.sheets.spreadsheets.values.update({
    spreadsheetId, range,
    valueInputOption: 'USER_ENTERED',
    resource: { values }
  });
}

async function sheetsAppend(tab, values) {
  await gapi.client.sheets.spreadsheets.values.append({
    spreadsheetId,
    range: tab + '!A1',
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    resource: { values }
  });
}

async function sheetsGetAll(tab) {
  const rows = await sheetsRead(tab + '!A:Z');
  if (rows.length <= 1) return [];
  const headers = rows[0];
  return rows.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i] || '');
    return obj;
  });
}

async function sheetsFindRow(tab, keyCol, keyVal) {
  const rows = await sheetsRead(tab + '!A:Z');
  for (let i = 1; i < rows.length; i++) {
    const headers = rows[0];
    const idx = headers.indexOf(keyCol);
    if (rows[i][idx] === keyVal) return { rowIndex: i + 1, row: rows[i], headers };
  }
  return null;
}

async function sheetsDeleteRow(tab, rowIndex) {
  const sheetId = await getSheetId(tab);
  await gapi.client.sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    resource: {
      requests: [{
        deleteDimension: {
          range: { sheetId, dimension: 'ROWS', startIndex: rowIndex - 1, endIndex: rowIndex }
        }
      }]
    }
  });
}

async function getSheetId(tabName) {
  const resp = await gapi.client.sheets.spreadsheets.get({ spreadsheetId });
  const sheet = resp.result.sheets.find(s => s.properties.title === tabName);
  return sheet ? sheet.properties.sheetId : 0;
}

// ── CONFIG ──
async function getCfg(key) {
  const rows = await sheetsGetAll(TAB.CFG);
  const row  = rows.find(r => r['Key'] === key);
  return row ? row['Value'] : null;
}

async function setCfg(key, value) {
  const found = await sheetsFindRow(TAB.CFG, 'Key', key);
  if (found) {
    const col = String.fromCharCode(65 + found.headers.indexOf('Value'));
    await sheetsWrite(`${TAB.CFG}!${col}${found.rowIndex}`, [[value]]);
  } else {
    await sheetsAppend(TAB.CFG, [[key, value]]);
  }
}

// ── NAVIGASI ──
function showPage(page) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.sidebar-item').forEach(s => s.classList.remove('active'));
  document.querySelectorAll('.bottom-nav-item').forEach(b => b.classList.remove('active'));
  document.getElementById('page-' + page).classList.add('active');
  document.querySelectorAll('.sidebar-item').forEach(s => {
    if (s.getAttribute('onclick') && s.getAttribute('onclick').includes("'" + page + "'")) s.classList.add('active');
  });
  document.querySelectorAll('.bottom-nav-item').forEach(b => {
    if (b.getAttribute('onclick') && b.getAttribute('onclick').includes("'" + page + "'")) b.classList.add('active');
  });
  currentPage = page;
  if (page === 'dashboard')   loadDashboard();
  if (page === 'riwayat')     loadRiwayat();
  if (page === 'laporan')     loadLaporan();
  if (page === 'budget')      loadBudget();
  if (page === 'kategori')    renderKategori();
  if (page === 'pengaturan')  loadPengaturan();
}

function initBulanFilter() {
  const sel = document.getElementById('bulanFilter');
  const now = new Date();
  for (let i = 0; i < 12; i++) {
    const d   = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const val = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0');
    const lbl = d.toLocaleDateString('id-ID', { month:'long', year:'numeric' });
    const opt = document.createElement('option');
    opt.value = val; opt.text = lbl;
    if (val === bulanAktif) opt.selected = true;
    sel.appendChild(opt);
  }
}

function onBulanChange() {
  bulanAktif = document.getElementById('bulanFilter').value;
  if (currentPage === 'dashboard') loadDashboard();
  else if (currentPage === 'riwayat') loadRiwayat();
  else if (currentPage === 'budget') loadBudget();
}

function setTanggalHariIni() {
  const now = new Date();
  document.getElementById('trx-tgl').value =
    now.getFullYear() + '-' + String(now.getMonth()+1).padStart(2,'0') + '-' + String(now.getDate()).padStart(2,'0');
}

// ── KATEGORI ──
async function loadKategori() {
  kategoriData = await sheetsGetAll(TAB.KAT);
  updateKatSelects();
}

function updateKatSelects() {
  updateSelect('trx-kat', kategoriData.filter(k => k['Jenis'] === currentJenis));
  updateSelect('bdg-kat', kategoriData.filter(k => k['Jenis'] === 'pengeluaran'));
}

function updateSelect(id, data) {
  const sel = document.getElementById(id);
  if (!sel) return;
  sel.innerHTML = '';
  data.forEach(k => {
    const o = document.createElement('option');
    o.value = k['Nama']; o.text = k['Nama'];
    sel.appendChild(o);
  });
}

function renderKategori() {
  const container = document.getElementById('kat-list');
  let filtered = kategoriData;
  if (filterKatMode !== 'semua') filtered = kategoriData.filter(k => k['Jenis'] === filterKatMode);
  if (!filtered.length) { container.innerHTML = '<div class="empty"><i class="ti ti-tags"></i>Belum ada kategori</div>'; return; }
  container.innerHTML = '<div class="kat-grid">' + filtered.map(k =>
    `<div class="kat-chip">
      <div class="dot" style="background:${k['Warna']}"></div>
      <i class="ti ${k['Icon']||'ti-tag'}" style="color:${k['Warna']}"></i>
      ${k['Nama']}
      <span class="del" onclick="hapusKategori('${k['Nama']}')"><i class="ti ti-x"></i></span>
    </div>`
  ).join('') + '</div>';
}

function filterKat(mode, el) {
  filterKatMode = mode;
  document.querySelectorAll('.stab').forEach(s => s.classList.remove('active'));
  el.classList.add('active');
  renderKategori();
}

async function tambahKategori() {
  const nama  = document.getElementById('kat-nama').value.trim();
  const jenis = document.getElementById('kat-jenis').value;
  const warna = document.getElementById('kat-warna').value;
  if (!nama) { showToast('Nama kategori wajib diisi', 'error'); return; }
  await sheetsAppend(TAB.KAT, [[nama, jenis, warna, 'ti-tag']]);
  showToast('Kategori ditambahkan');
  document.getElementById('kat-nama').value = '';
  await loadKategori();
  renderKategori();
}

async function hapusKategori(nama) {
  if (!confirm('Hapus kategori "' + nama + '"?')) return;
  const found = await sheetsFindRow(TAB.KAT, 'Nama', nama);
  if (found) { await sheetsDeleteRow(TAB.KAT, found.rowIndex); showToast('Kategori dihapus'); await loadKategori(); renderKategori(); }
}

// ── JENIS TOGGLE ──
function setJenis(jenis) {
  currentJenis = jenis;
  document.querySelectorAll('.jenis-btn').forEach(b => b.classList.remove('active'));
  if (jenis === 'pemasukan') document.querySelector('.jenis-btn.masuk').classList.add('active');
  else document.querySelector('.jenis-btn.keluar').classList.add('active');
  updateSelect('trx-kat', kategoriData.filter(k => k['Jenis'] === jenis));
}

// ── TRANSAKSI ──
async function simpanTransaksi() {
  const tgl    = document.getElementById('trx-tgl').value;
  const jumlah = parseFloat(document.getElementById('trx-jumlah').value);
  const kat    = document.getElementById('trx-kat').value;
  const ket    = document.getElementById('trx-ket').value;
  if (!tgl || !jumlah || jumlah <= 0) { showToast('Lengkapi tanggal dan jumlah', 'error'); return; }

  const btn = document.querySelector('#page-catat .btn-primary');
  btn.disabled = true; btn.innerHTML = '<i class="ti ti-loader"></i> Menyimpan...';

  try {
    // Hitung saldo terakhir
    const rows    = await sheetsGetAll(TAB.TRX);
    let saldo     = 0;
    if (rows.length > 0) saldo = parseFloat(rows[rows.length-1]['Saldo']) || 0;
    saldo = currentJenis === 'pemasukan' ? saldo + jumlah : saldo - jumlah;

    const id = crypto.randomUUID();
    await sheetsAppend(TAB.TRX, [[id, tgl, currentJenis, kat, ket, jumlah, saldo, new Date().toISOString()]]);

    // Update budget terpakai
    if (currentJenis === 'pengeluaran') await updateBudgetTerpakai(tgl.substring(0,7), kat, jumlah);

    // Notif WA
    kirimWA(tgl, currentJenis, kat, ket, jumlah, saldo);

    document.getElementById('trx-jumlah').value = '';
    document.getElementById('trx-ket').value    = '';
    document.getElementById('navSaldoVal').textContent = 'Rp ' + formatRp(saldo);
    showToast('Transaksi berhasil disimpan ✓');
    loadDashboard();
  } catch(e) {
    showToast('Gagal: ' + e.message, 'error');
  }
  btn.disabled = false; btn.innerHTML = '<i class="ti ti-device-floppy"></i> Simpan Transaksi';
}

async function getTransaksi(bulan, jenis) {
  let rows = await sheetsGetAll(TAB.TRX);
  if (bulan) rows = rows.filter(r => r['Tanggal'].startsWith(bulan));
  if (jenis) rows = rows.filter(r => r['Jenis'] === jenis);
  return rows.reverse();
}

async function hapusTransaksi(id) {
  if (!confirm('Hapus transaksi ini?')) return;
  const found = await sheetsFindRow(TAB.TRX, 'ID', id);
  if (found) {
    await sheetsDeleteRow(TAB.TRX, found.rowIndex);
    await recalcSaldo();
    showToast('Transaksi dihapus');
    loadRiwayat();
    if (currentPage === 'dashboard') loadDashboard();
  }
}

async function recalcSaldo() {
  const rows = await sheetsGetAll(TAB.TRX);
  if (!rows.length) return;
  let saldo = 0;
  const newRows = rows.map(r => {
    const jml = parseFloat(r['Jumlah']) || 0;
    saldo = r['Jenis'] === 'pemasukan' ? saldo + jml : saldo - jml;
    return [r['ID'], r['Tanggal'], r['Jenis'], r['Kategori'], r['Keterangan'], r['Jumlah'], saldo, r['Timestamp']];
  });
  await sheetsWrite(TAB.TRX + '!A2:H' + (newRows.length + 1), newRows);
}

// ── DASHBOARD ──
async function loadDashboard() {
  const rows = await getTransaksi(bulanAktif);
  let masuk = 0, keluar = 0;
  const perKat = {};
  rows.forEach(r => {
    const jml = parseFloat(r['Jumlah']) || 0;
    if (r['Jenis'] === 'pemasukan') masuk += jml;
    else { keluar += jml; perKat[r['Kategori']] = (perKat[r['Kategori']] || 0) + jml; }
  });

  // Saldo total dari semua transaksi
  const allRows = await sheetsGetAll(TAB.TRX);
  const saldoTotal = allRows.length ? (parseFloat(allRows[allRows.length-1]['Saldo']) || 0) : 0;
  const net = masuk - keluar;

  document.getElementById('db-masuk').textContent  = 'Rp ' + formatRp(masuk);
  document.getElementById('db-keluar').textContent = 'Rp ' + formatRp(keluar);
  const netEl = document.getElementById('db-net');
  netEl.textContent = (net < 0 ? '- Rp ' : 'Rp ') + formatRp(Math.abs(net));
  netEl.className   = 'stat-val ' + (net >= 0 ? 'txt-hijau' : 'txt-merah');
  document.getElementById('db-saldo').textContent  = 'Rp ' + formatRp(saldoTotal);
  document.getElementById('navSaldoVal').textContent = 'Rp ' + formatRp(saldoTotal);

  renderDonut(perKat);
  const el = document.getElementById('db-trx-list');
  if (!rows.length) { el.innerHTML = '<div class="empty"><i class="ti ti-list"></i>Belum ada transaksi bulan ini</div>'; return; }
  el.innerHTML = rows.slice(0,6).map(t => renderTrxItem(t, false)).join('');
}

function renderDonut(perKat) {
  const labels = Object.keys(perKat);
  const values = Object.values(perKat);
  if (!labels.length) return;
  const colors = ['#ef4444','#f97316','#8b5cf6','#3b82f6','#06b6d4','#ec4899','#10b981','#64748b','#f59e0b','#22c55e'];
  if (chartDonut) chartDonut.destroy();
  const ctx = document.getElementById('chartDonut');
  if (!ctx) return;
  chartDonut = new Chart(ctx, {
    type: 'doughnut',
    data: { labels, datasets: [{ data: values, backgroundColor: colors.slice(0, labels.length), borderWidth: 2, borderColor: '#fff' }] },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } }, tooltip: { callbacks: { label: ctx => ' Rp ' + formatRp(ctx.raw) } } } }
  });
}

// ── RIWAYAT ──
async function loadRiwayat() {
  const jenis = document.getElementById('filter-jenis').value;
  const el    = document.getElementById('riwayat-list');
  el.innerHTML = '<div class="loader"><i class="ti ti-loader"></i>Memuat...</div>';
  const rows  = await getTransaksi(bulanAktif, jenis);
  if (!rows.length) { el.innerHTML = '<div class="empty"><i class="ti ti-list"></i>Tidak ada transaksi</div>'; return; }
  el.innerHTML = rows.map(t => renderTrxItem(t, true)).join('');
}

function renderTrxItem(t, showDel) {
  const kat   = kategoriData.find(k => k['Nama'] === t['Kategori']) || {};
  const warna = kat['Warna'] || (t['Jenis'] === 'pemasukan' ? '#16a34a' : '#dc2626');
  const icon  = kat['Icon']  || (t['Jenis'] === 'pemasukan' ? 'ti-trending-up' : 'ti-trending-down');
  const sign  = t['Jenis'] === 'pemasukan' ? '+' : '-';
  const col   = t['Jenis'] === 'pemasukan' ? '#16a34a' : '#dc2626';
  const del   = showDel ? `<span class="trx-del" onclick="hapusTransaksi('${t['ID']}')"><i class="ti ti-trash"></i></span>` : '';
  return `<div class="trx-item">
    <div class="trx-icon" style="background:${warna}22;color:${warna}"><i class="ti ${icon}"></i></div>
    <div class="trx-info"><div class="trx-kat">${t['Kategori']}</div><div class="trx-ket">${t['Keterangan']||'-'}</div></div>
    <div><div class="trx-jml" style="color:${col}">${sign}Rp ${formatRp(t['Jumlah'])}</div><div class="trx-tgl">${t['Tanggal']}</div></div>
    ${del}
  </div>`;
}

// ── LAPORAN ──
async function loadLaporan() {
  const allRows = await sheetsGetAll(TAB.TRX);
  const byBulan = {};
  allRows.forEach(r => {
    const bln = r['Tanggal'].substring(0,7);
    if (!byBulan[bln]) byBulan[bln] = { masuk:0, keluar:0 };
    const jml = parseFloat(r['Jumlah']) || 0;
    if (r['Jenis'] === 'pemasukan') byBulan[bln].masuk += jml;
    else byBulan[bln].keluar += jml;
  });
  const data = Object.entries(byBulan).sort((a,b) => a[0].localeCompare(b[0])).map(([bln,v]) => ({ bulan:bln, masuk:v.masuk, keluar:v.keluar, net:v.masuk-v.keluar }));

  const tbody = document.getElementById('laporan-tbl');
  if (!data.length) { tbody.innerHTML = '<tr><td colspan="4" style="text-align:center;color:var(--text-3)">Belum ada data</td></tr>'; return; }
  tbody.innerHTML = data.map(r => {
    const lbl = new Date(r.bulan+'-01').toLocaleDateString('id-ID',{month:'long',year:'numeric'});
    return `<tr><td>${lbl}</td><td style="color:var(--hijau)">Rp ${formatRp(r.masuk)}</td><td style="color:var(--merah)">Rp ${formatRp(r.keluar)}</td><td style="color:${r.net>=0?'var(--hijau)':'var(--merah)'};font-weight:600">${r.net>=0?'+':''}Rp ${formatRp(r.net)}</td></tr>`;
  }).join('');
  renderBarChart(data.slice(-6));
}

function renderBarChart(data) {
  const labels = data.map(r => new Date(r.bulan+'-01').toLocaleDateString('id-ID',{month:'short'}));
  if (chartBar) chartBar.destroy();
  const ctx = document.getElementById('chartBar');
  if (!ctx) return;
  chartBar = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [
      { label:'Pemasukan', data:data.map(r=>r.masuk), backgroundColor:'#16a34a88', borderColor:'#16a34a', borderWidth:1.5, borderRadius:6 },
      { label:'Pengeluaran', data:data.map(r=>r.keluar), backgroundColor:'#dc262688', borderColor:'#dc2626', borderWidth:1.5, borderRadius:6 }
    ]},
    options: { responsive:true, maintainAspectRatio:false, plugins:{ tooltip:{ callbacks:{ label:ctx=>' Rp '+formatRp(ctx.raw) } } }, scales:{ y:{ ticks:{ callback:v=>'Rp '+formatRp(v) } }, x:{ grid:{ display:false } } } }
  });
}

// ── BUDGET ──
async function loadBudget() {
  document.getElementById('budget-bln-lbl').textContent = new Date(bulanAktif+'-01').toLocaleDateString('id-ID',{month:'long',year:'numeric'});
  updateSelect('bdg-kat', kategoriData.filter(k => k['Jenis'] === 'pengeluaran'));
  const rows = await sheetsGetAll(TAB.BUDGET);
  const budgets = rows.filter(r => r['Bulan'] === bulanAktif);
  const el = document.getElementById('budget-list');
  if (!budgets.length) { el.innerHTML = '<div class="empty"><i class="ti ti-target"></i>Belum ada budget. Klik "Atur Budget" untuk menambahkan.</div>'; return; }
  el.innerHTML = budgets.map(b => {
    const pct = b['Anggaran'] > 0 ? Math.min(100, (b['Terpakai']/b['Anggaran'])*100) : 0;
    const col = pct >= 100 ? '#dc2626' : pct >= 80 ? '#d97706' : '#16a34a';
    return `<div class="budget-item">
      <div class="budget-header"><span style="font-weight:600;font-size:13px">${b['Kategori']}${pct>=100?' ⚠️':''}</span><span style="font-size:12px;color:var(--text-2)">Rp ${formatRp(b['Terpakai'])} / Rp ${formatRp(b['Anggaran'])}</span></div>
      <div class="budget-bar"><div class="budget-fill" style="width:${pct.toFixed(1)}%;background:${col}"></div></div>
      <div style="font-size:11px;color:var(--text-3);margin-top:3px">${pct.toFixed(1)}% terpakai</div>
    </div>`;
  }).join('');
}

async function simpanBudget() {
  const kat = document.getElementById('bdg-kat').value;
  const jml = parseFloat(document.getElementById('bdg-jml').value);
  if (!kat || !jml || jml <= 0) { showToast('Lengkapi data budget', 'error'); return; }
  const found = await sheetsFindRow(TAB.BUDGET, 'Kategori', kat);
  if (found && found.row[0] === bulanAktif) {
    const col = String.fromCharCode(65 + found.headers.indexOf('Anggaran'));
    await sheetsWrite(`${TAB.BUDGET}!${col}${found.rowIndex}`, [[jml]]);
  } else {
    await sheetsAppend(TAB.BUDGET, [[bulanAktif, kat, jml, 0]]);
  }
  showToast('Budget disimpan');
  closeModal('modal-budget');
  loadBudget();
}

async function updateBudgetTerpakai(bulan, kat, jumlah) {
  const rows = await sheetsGetAll(TAB.BUDGET);
  for (let i = 0; i < rows.length; i++) {
    if (rows[i]['Bulan'] === bulan && rows[i]['Kategori'] === kat) {
      const terpakai  = (parseFloat(rows[i]['Terpakai']) || 0) + jumlah;
      const anggaran  = parseFloat(rows[i]['Anggaran']) || 0;
      const rowIndex  = i + 2;
      await sheetsWrite(`${TAB.BUDGET}!D${rowIndex}`, [[terpakai]]);
      if (anggaran > 0 && terpakai > anggaran) {
        kirimWABudget(kat, anggaran, terpakai);
      }
      return;
    }
  }
}

// ── PENGATURAN ──
async function loadPengaturan() {
  const nama  = await getCfg('nama_pengguna');
  const wa    = await getCfg('no_wa');
  const token = await getCfg('wa_token');
  document.getElementById('cfg-nama').value  = nama  || '';
  document.getElementById('cfg-wa').value    = wa    || '';
  document.getElementById('cfg-token').value = token || '';
  document.getElementById('cfg-sheet-id').value = spreadsheetId || '';
}

async function simpanProfil() {
  await setCfg('nama_pengguna', document.getElementById('cfg-nama').value);
  showToast('Profil disimpan');
}

async function simpanConfig() {
  await setCfg('no_wa',    document.getElementById('cfg-wa').value);
  await setCfg('wa_token', document.getElementById('cfg-token').value);
  showToast('Pengaturan WA disimpan');
}

function bukaSheets() {
  window.open('https://docs.google.com/spreadsheets/d/' + spreadsheetId, '_blank');
}

// ── NOTIF WA (Fonnte) ──
async function kirimWA(tgl, jenis, kat, ket, jumlah, saldo) {
  const noWA  = await getCfg('no_wa');
  const token = await getCfg('wa_token');
  if (!noWA || !token) return;
  const pesan = `💰 *UANGKU* - Transaksi Baru\n📅 ${tgl}\n${jenis==='pemasukan'?'📈 Pemasukan':'📉 Pengeluaran'}: ${kat}\n📝 ${ket||'-'}\n💵 Rp ${formatRp(jumlah)}\n💼 Saldo: Rp ${formatRp(saldo)}`;
  fetch('https://api.fonnte.com/send', {
    method: 'POST',
    headers: { 'Authorization': token },
    body: new URLSearchParams({ target: noWA, message: pesan })
  }).catch(() => {});
}

async function kirimWABudget(kat, anggaran, terpakai) {
  const noWA  = await getCfg('no_wa');
  const token = await getCfg('wa_token');
  if (!noWA || !token) return;
  const pesan = `⚠️ *UANGKU* - Over Budget!\nKategori *${kat}* melebihi anggaran!\n🎯 Anggaran: Rp ${formatRp(anggaran)}\n💸 Terpakai: Rp ${formatRp(terpakai)}`;
  fetch('https://api.fonnte.com/send', {
    method: 'POST',
    headers: { 'Authorization': token },
    body: new URLSearchParams({ target: noWA, message: pesan })
  }).catch(() => {});
}

// ── MODAL ──
function openModal(id)  { document.getElementById(id).classList.add('open'); }
function closeModal(id) { document.getElementById(id).classList.remove('open'); }

// ── TOAST ──
let toastTimer;
function showToast(msg, type) {
  const el = document.getElementById('toast');
  document.getElementById('toast-msg').textContent = msg;
  el.className = 'toast show ' + (type || 'success');
  el.querySelector('i').className = type === 'error' ? 'ti ti-alert-circle' : 'ti ti-check';
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove('show'), 3000);
}

// ── FORMAT ──
function formatRp(n) { return Math.round(n||0).toLocaleString('id-ID'); }
