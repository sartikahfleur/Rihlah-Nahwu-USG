/**
 * Penerima nilai "Rihlah Nahwu bersama Ning Tika"
 * Tempel di Google Sheet: Ekstensi → Apps Script, lalu Deploy sebagai Aplikasi Web
 * (Jalankan sebagai: Saya; Yang memiliki akses: Siapa saja).
 */
const NAMA_SHEET = 'Nilai';
const KOLOM = ['Waktu', 'Nama', 'NIM', 'Kelas', 'Pertemuan', 'Jenis', 'Nilai', 'Benar', 'Jumlah Soal', 'Durasi (menit)', 'Soal yang salah'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(NAMA_SHEET) || ss.insertSheet(NAMA_SHEET);
    if (sh.getLastRow() === 0) {
      sh.appendRow(KOLOM);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, KOLOM.length).setFontWeight('bold').setBackground('#13285c').setFontColor('#ffffff');
    }
    sh.appendRow([
      new Date(),
      teks(d.nama), "'" + teks(d.nim), teks(d.kelas).toUpperCase(),
      teks(d.pertemuan), teks(d.jenis),
      angka(d.nilai), angka(d.benar), angka(d.jumlah), angka(d.durasi),
      teks(d.salah)
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('Penerima nilai Rihlah Nahwu aktif.');
}

// Mencegah teks diperlakukan sebagai rumus dan membatasi panjangnya
function teks(v) {
  let s = String(v == null ? '' : v).slice(0, 500).trim();
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
function angka(v) {
  const n = Number(v);
  return isFinite(n) ? n : '';
}
