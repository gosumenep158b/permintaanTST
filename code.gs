const SHEET_SISWA = 'Siswa';
const SHEET_PERMINTAAN = 'Permintaan TST';
const SHEET_PENJADWALAN = 'Penjadwalan TST';
const KEBUTUHAN_LIST = ['PR', 'Tugas', 'UH', 'Pendalaman Materi'];
const STATUS_LIST = ['Menunggu', 'Terjadwal', 'Selesai', 'Batal'];
const JML_HARI = 3;
const WARNA = '#CE1126';
const WARNA_GELAP = '#A80E1F';
const HEADERS_PERMINTAAN = ['No', 'Tanggal Permintaan TST', 'Tingkat Kelas', 'Kelas di GO', 'Nama Siswa', 'Nomor HP', 'Mata Pelajaran', 'Nama Bab', 'Kebutuhan', 'Status', 'Keterangan'];
const HEADERS_PENJADWALAN = ['Tanggal TST', 'Jam', 'No Permintaan', 'Tingkat Kelas', 'Kelas di GO', 'Nama Siswa', 'Nomor HP', 'Mata Pelajaran', 'Kebutuhan', 'Nama Pengajar', 'Status', 'Info WA'];

function doGet(e) {
  if (e && e.parameter && e.parameter.api !== undefined) {
    let out;
    try {
      out = apiDispatch({
        fn: e.parameter.fn,
        args: e.parameter.args ? JSON.parse(e.parameter.args) : []
      });
    } catch (err) {
      out = { error: 'Terjadi kesalahan: ' + err.message };
    }
    return ContentService
      .createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('TST GO Sumenep')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function apiDispatch(payload) {
  if (!payload || !payload.fn) return { error: 'Fungsi tidak dikenali.' };
  const args = payload.args || [];
  switch (payload.fn) {
    case 'getConfig': return getConfig();
    case 'getSiswa': return getSiswa();
    case 'submitPermintaan': return submitPermintaan(args[0]);
    case 'loginAdmin': return loginAdmin(args[0], args[1]);
    case 'getPermintaan': return getPermintaan(args[0]);
    case 'getPenjadwalan': return getPenjadwalan(args[0]);
    case 'jadwalkanTST': return jadwalkanTST(args[0], args[1], args[2], args[3], args[4]);
    case 'batalkanJadwal': return batalkanJadwal(args[0], args[1]);
    case 'kirimPesanWA': return kirimPesanWA(args[0], args[1]);
    case 'gantiPassword': return gantiPassword(args[0], args[1], args[2]);
    case 'getAdminUser': return getAdminUser(args[0]);
    case 'getTelegramStatus': return getTelegramStatus(args[0]);
    case 'simpanBotConfig': return simpanBotConfig(args[0], args[1], args[2]);
    case 'tesTelegram': return tesTelegram(args[0]);
    default: return { error: 'Fungsi tidak dikenali.' };
  }
}

function doPost(e) {
  let payload = null;
  try { payload = JSON.parse(e.postData.contents); } catch (err) { payload = null; }
  let out;
  try {
    out = apiDispatch(payload);
  } catch (err) {
    out = { error: 'Terjadi kesalahan: ' + err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function getConfig() {
  return { kebutuhan: KEBUTUHAN_LIST, status: STATUS_LIST, jmlHari: JML_HARI };
}

const RENAME_HEADER = { 'Tanggal Jadwal': 'Tanggal TST' };

function syncHeaders(sheet, targetHeaders) {
  const isEmpty = sheet.getLastRow() === 0;
  if (isEmpty) {
    sheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
    sheet.getRange(1, 1, 1, targetHeaders.length)
      .setBackground(WARNA).setFontColor('#FFFFFF').setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    return;
  }
  const last = sheet.getLastColumn();
  const cur = sheet.getRange(1, 1, 1, last).getValues()[0].map(h => String(h).trim());
  let src = 0;
  for (let i = 0; i < targetHeaders.length; i++) {
    const t = targetHeaders[i];
    if (src < cur.length && cur[src] === t) { src++; continue; }
    if (src < cur.length && RENAME_HEADER[cur[src]] === t) {
      sheet.getRange(1, i + 1).setValue(t);
      cur[src] = t;
      src++;
      continue;
    }
    const existsIdx = cur.indexOf(t, src);
    if (existsIdx === -1) {
      sheet.insertColumnAfter(Math.max(1, i));
      cur.splice(i, 0, '');
      sheet.getRange(1, i + 1).setValue(t);
    } else {
      sheet.getRange(1, i + 1).setValue(t);
    }
  }
  sheet.getRange(1, 1, 1, targetHeaders.length)
    .setBackground(WARNA).setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
}

function getTZ() {
  try {
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    return tz || Session.getScriptTimeZone();
  } catch (e) {
    return Session.getScriptTimeZone();
  }
}

function formatDate(d) {
  if (!d) return '';
  if (typeof d === 'string') return d;
  return Utilities.formatDate(d, getTZ(), 'dd-MM-yyyy');
}

function normalizePhone(p) {
  if (p === null || p === undefined) return '';
  let s = String(p).trim();
  if (typeof p === 'number' || /[eE]/.test(s)) {
    if (typeof p === 'number') {
      s = p.toFixed(0);
    } else {
      const n = Number(s.replace(',', '.'));
      if (!isNaN(n)) s = n.toFixed(0);
    }
  }
  s = s.replace(/\D/g, '');
  if (!s) return '';
  if (s.indexOf('0') === 0) s = '62' + s.slice(1);
  else if (s.indexOf('8') === 0) s = '62' + s;
  return s;
}

function getAdmin() {
  const props = PropertiesService.getScriptProperties();
  let u = props.getProperty('ADMIN_USER');
  let p = props.getProperty('ADMIN_PASS');
  if (!u || !p) {
    u = 'admin';
    p = 'admin123';
    props.setProperty('ADMIN_USER', u);
    props.setProperty('ADMIN_PASS', p);
  }
  return { user: u, pass: p };
}

function simpanToken(token) {
  if (token) CacheService.getScriptCache().put('ADMIN_TOKEN', token, 21600);
  return true;
}

function cekToken(token) {
  const t = CacheService.getScriptCache().get('ADMIN_TOKEN');
  return !!t && t === token;
}

function loginAdmin(username, password) {
  const a = getAdmin();
  if (String(username) === a.user && String(password) === a.pass) {
    const token = Utilities.getUuid();
    CacheService.getScriptCache().put('ADMIN_TOKEN', token, 21600);
    return { success: true, token: token };
  }
  return { success: false, message: 'Username atau password salah.' };
}

function gantiPassword(token, lama, baru) {
  if (!cekToken(token)) return { success: false, message: 'Sesi berakhir, silakan login ulang.' };
  const a = getAdmin();
  if (String(lama) !== a.pass) return { success: false, message: 'Password lama salah.' };
  if (!baru || String(baru).length < 6) return { success: false, message: 'Password baru minimal 6 karakter.' };
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASS', String(baru));
  return { success: true, message: 'Password berhasil diganti.' };
}

function getAdminUser(token) {
  if (!cekToken(token)) return '';
  return getAdmin().user;
}

function getSiswa() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SISWA);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const idx = {};
  values[0].forEach((c, i) => { idx[String(c).trim().toLowerCase()] = i; });
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const nama = r[idx['nama_siswa']];
    if (!nama) continue;
    out.push({
      noReg: String(r[idx['no_reg']] ?? '').trim(),
      nama: String(nama).trim(),
      noHp: normalizePhone(r[idx['no_hp']]),
      kelasGo: String(r[idx['kelas_go']] ?? '').trim(),
      tingkat: String(r[idx['tingkat_kelas']] ?? '').trim()
    });
  }
  return out;
}

function ensureSheetPermintaan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_PERMINTAAN);
  const isNew = !sheet;
  if (!sheet) sheet = ss.insertSheet(SHEET_PERMINTAAN);
  syncHeaders(sheet, HEADERS_PERMINTAAN);
  sheet.setFrozenRows(1);
  sheet.getRange('A:K').setVerticalAlignment('middle');
  sheet.getRange('B:B').setNumberFormat('dd-MM-yyyy');
  sheet.setColumnWidth(2, 130);
  if (isNew) {
    const dvK = SpreadsheetApp.newDataValidation().requireValueInList(KEBUTUHAN_LIST, true).build();
    sheet.getRange(2, 9, 1000).setDataValidation(dvK);
    const dvS = SpreadsheetApp.newDataValidation().requireValueInList(STATUS_LIST, true).build();
    sheet.getRange(2, 10, 1000).setDataValidation(dvS);
  }
  return sheet;
}

function submitPermintaan(data) {
  data = data || {};
  const sheet = ensureSheetPermintaan();
  const nama = String(data.nama || '').trim();
  const noHp = normalizePhone(data.noHp);
  if (!nama) return { success: false, message: 'Nama siswa wajib diisi.' };
  if (!data.tingkat) return { success: false, message: 'Tingkat kelas wajib diisi.' };
  if (!data.kelasGo) return { success: false, message: 'Kelas di GO wajib diisi.' };
  if (!data.namaBab) return { success: false, message: 'Nama bab wajib diisi.' };
  if (!data.kebutuhan) return { success: false, message: 'Kebutuhan wajib dipilih.' };
  if (!data.tglPermintaan) return { success: false, message: 'Tanggal permintaan wajib diisi.' };
  const existing = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues() : [];
  let maxNo = 0;
  existing.forEach(row => {
    const n = Number(row[0]);
    if (!isNaN(n) && n > maxNo) maxNo = n;
  });
  const nomor = maxNo + 1;
  sheet.appendRow([
    nomor,
    new Date(data.tglPermintaan),
    String(data.tingkat),
    String(data.kelasGo),
    nama,
    noHp,
    String(data.mapel || ''),
    String(data.namaBab),
    String(data.kebutuhan),
    'Menunggu',
    String(data.keterangan || '')
  ]);
  kirimNotifTelegram(
    '<b>PERMINTAAN TST BARU</b>\n\n' +
    'No : ' + nomor + '\n' +
    'Tanggal : ' + formatDate(new Date(data.tglPermintaan)) + '\n' +
    'Nama : ' + nama + '\n' +
    'Tingkat : ' + String(data.tingkat) + '\n' +
    'Kelas GO : ' + String(data.kelasGo) + '\n' +
    'Mata Pelajaran : ' + String(data.mapel || '-') + '\n' +
    'Nama Bab : ' + String(data.namaBab) + '\n' +
    'Kebutuhan : ' + String(data.kebutuhan) + '\n' +
    'No HP : ' + noHp
  );
  return { success: true, message: 'Permintaan TST berhasil dikirim. Terima kasih!' };
}

function getBotConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    botToken: String(props.getProperty('TG_BOT_TOKEN') || ''),
    chatId: String(props.getProperty('TG_CHAT_ID') || '')
  };
}

function getTelegramStatus(token) {
  if (!cekToken(token)) return { error: 'Sesi berakhir, silakan login ulang.' };
  const c = getBotConfig();
  return {
    success: true,
    isSet: !!(c.botToken && c.chatId),
    botToken: c.botToken ? mask(c.botToken) : '',
    chatId: c.chatId || ''
  };
}

function mask(s) {
  s = String(s);
  if (s.length <= 8) return '****';
  return s.slice(0, 4) + '...' + s.slice(-4);
}

function tesTelegram(token) {
  if (!cekToken(token)) return { success: false, message: 'Sesi berakhir, silakan login ulang.' };
  const c = getBotConfig();
  if (!c.botToken || !c.chatId) return { success: false, message: 'Bot Token / Chat ID belum diatur.' };
  try {
    const resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + c.botToken + '/sendMessage', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id: c.chatId,
        text: 'Notifikasi TST GO Sumenep berhasil terhubung.',
        parse_mode: 'HTML'
      }),
      muteHttpExceptions: true
    });
    const json = JSON.parse(resp.getContentText());
    if (json.ok) return { success: true, message: 'Pesan uji berhasil dikirim ke Telegram.' };
    return { success: false, message: 'Telegram menolak: ' + (json.description || ('kode ' + json.error_code)) };
  } catch (e) {
    return { success: false, message: 'Gagal menghubungi Telegram: ' + e.message };
  }
}

function simpanBotConfig(token, botToken, chatId) {
  if (!cekToken(token)) return { success: false, message: 'Sesi berakhir, silakan login ulang.' };
  if (!botToken || !chatId) return { success: false, message: 'Bot Token dan Chat ID wajib diisi.' };
  const props = PropertiesService.getScriptProperties();
  props.setProperty('TG_BOT_TOKEN', String(botToken).trim());
  props.setProperty('TG_CHAT_ID', String(chatId).trim());
  return { success: true, message: 'Konfigurasi Telegram berhasil disimpan.' };
}

function kirimNotifTelegram(text) {
  const c = getBotConfig();
  if (!c.botToken || !c.chatId) return;
  try {
    UrlFetchApp.fetch('https://api.telegram.org/bot' + c.botToken + '/sendMessage', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        chat_id: c.chatId,
        text: text,
        parse_mode: 'HTML'
      }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log('Telegram notif gagal: ' + e.message);
  }
}

function getPermintaan(token) {
  if (!cekToken(token)) return { error: 'Sesi berakhir, silakan login ulang.' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PERMINTAAN);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[4]) continue;
    out.push({
      no: r[0],
      tglPermintaan: formatDate(r[1]),
      tingkat: String(r[2] ?? ''),
      kelasGo: String(r[3] ?? ''),
      nama: String(r[4] ?? ''),
      noHp: String(r[5] ?? ''),
      mapel: String(r[6] ?? ''),
      namaBab: String(r[7] ?? ''),
      kebutuhan: String(r[8] ?? ''),
      status: String(r[9] || 'Menunggu'),
      keterangan: String(r[10] ?? '')
    });
  }
  return out.reverse();
}

function ensureSheetPenjadwalan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_PENJADWALAN);
  const isNew = !sheet;
  if (!sheet) sheet = ss.insertSheet(SHEET_PENJADWALAN);
  syncHeaders(sheet, HEADERS_PENJADWALAN);
  sheet.setFrozenRows(1);
  sheet.getRange('A:L').setVerticalAlignment('middle');
  sheet.getRange('A:A').setNumberFormat('dd-MM-yyyy');
  sheet.setColumnWidth(1, 130);
  if (isNew) {
    const dvS = SpreadsheetApp.newDataValidation().requireValueInList(STATUS_LIST, true).build();
    sheet.getRange(2, 11, 1000).setDataValidation(dvS);
  }
  return sheet;
}

function jadwalkanTST(token, no, tglJadwal, jam, pengajar) {
  if (!cekToken(token)) return { success: false, message: 'Sesi berakhir, silakan login ulang.' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const perm = ss.getSheetByName(SHEET_PERMINTAAN);
  if (!perm) return { success: false, message: 'Sheet permintaan belum ada.' };
  if (!tglJadwal) return { success: false, message: 'Tanggal TST wajib diisi.' };
  if (!jam) return { success: false, message: 'Jam TST wajib diisi.' };
  if (!pengajar) return { success: false, message: 'Nama pengajar wajib diisi.' };
  const values = perm.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(no)) { rowIndex = i + 1; break; }
  }
  if (rowIndex === -1) return { success: false, message: 'Permintaan tidak ditemukan.' };
  if (String(values[rowIndex - 1][9]) !== 'Menunggu') return { success: false, message: 'Permintaan sudah dijadwalkan.' };
  const r = values[rowIndex - 1];
  const jadwalSheet = ensureSheetPenjadwalan();
  const tgl = new Date(tglJadwal);
  jadwalSheet.appendRow([
    tgl, String(jam), r[0], r[2], r[3], r[4], r[5], r[6], r[8], String(pengajar), 'Terjadwal', 'Belum dikirim'
  ]);
  perm.getRange(rowIndex, 10).setValue('Terjadwal');
  return { success: true, message: 'Permintaan No ' + no + ' berhasil dijadwalkan.' };
}

function getPenjadwalan(token) {
  if (!cekToken(token)) return { error: 'Sesi berakhir, silakan login ulang.' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PENJADWALAN);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[4]) continue;
    out.push({
      id: i + 1,
      tglJadwal: formatDate(r[0]),
      jam: String(r[1] ?? ''),
      no: r[2],
      tingkat: String(r[3] ?? ''),
      kelasGo: String(r[4] ?? ''),
      nama: String(r[5] ?? ''),
      noHp: String(r[6] ?? ''),
      mapel: String(r[7] ?? ''),
      kebutuhan: String(r[8] ?? ''),
      pengajar: String(r[9] ?? ''),
      status: String(r[10] || 'Terjadwal'),
      infoWa: String(r[11] ?? '')
    });
  }
  return out.reverse();
}

function batalkanJadwal(token, id) {
  if (!cekToken(token)) return { success: false, message: 'Sesi berakhir, silakan login ulang.' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const jadwalSheet = ss.getSheetByName(SHEET_PENJADWALAN);
  if (!jadwalSheet) return { success: false, message: 'Sheet jadwal belum ada.' };
  const no = jadwalSheet.getRange(Number(id), 3).getValue();
  const perm = ss.getSheetByName(SHEET_PERMINTAAN);
  if (perm) {
    const values = perm.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(no)) {
        perm.getRange(i + 1, 9).setValue('Menunggu');
        break;
      }
    }
  }
  jadwalSheet.deleteRow(Number(id));
  return { success: true, message: 'Jadwal berhasil dibatalkan, permintaan kembali ke antrian.' };
}

function kirimPesanWA(token, id) {
  if (!cekToken(token)) return { success: false, message: 'Sesi berakhir, silakan login ulang.' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PENJADWALAN);
  if (!sheet) return { success: false, message: 'Sheet jadwal belum ada.' };
  const r = sheet.getRange(Number(id), 1, 1, 12).getValues()[0];
  if (!r[5]) return { success: false, message: 'Baris jadwal tidak ditemukan.' };
  const noHp = normalizePhone(r[6]);
  if (!noHp) return { success: false, message: 'Nomor HP siswa tidak tersedia.' };
  const pesan =
    '*PENJADWALAN TST GO SUMENEP*\n\n' +
    'Halo *' + String(r[5]) + '*,\n\n' +
    'Permintaan TST kamu sudah dijadwalkan dengan detail berikut:\n\n' +
    'Tanggal TST : ' + formatDate(r[0]) + '\n' +
    'Jam : ' + String(r[1]) + '\n' +
    'Kelas di GO : ' + String(r[4]) + '\n' +
    'Tingkat : ' + String(r[3]) + '\n' +
    'Mata Pelajaran : ' + String(r[7]) + '\n' +
    'Kebutuhan : ' + String(r[8]) + '\n' +
    'Pengajar : ' + String(r[9]) + '\n\n' +
    'Mohon hadir tepat waktu. Terima kasih!\n' +
    '- Admin GO Sumenep';
  sheet.getRange(Number(id), 12).setValue('Dikirim ' + Utilities.formatDate(new Date(), getTZ(), 'dd-MM-yyyy HH:mm'));
  return { success: true, noHp: noHp, pesan: pesan };
}
