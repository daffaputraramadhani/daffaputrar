/**
 * Cek Laporan - backend rekap milik pengelola.
 * 1. Buat Google Spreadsheet baru.
 * 2. Buka Extensions > Apps Script.
 * 3. Ganti isi Code.gs dengan kode ini.
 * 4. Deploy > New deployment > Web app.
 *    Execute as: Me
 *    Who has access: Anyone
 * 5. Salin URL /exec ke const SCRIPT_URL di index.html.
 *
 * Endpoint publik dapat menerima kiriman dari siapa saja yang memiliki URL.
 * Lindungi spreadsheet dengan tidak membagikannya sebagai editor.
 */
const SHEET_NAME = 'Rekap Laporan';

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents;
    if (!raw) return jsonOutput({ ok: false, error: 'Empty request' });

    const data = JSON.parse(raw);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Waktu Kirim', 'ID', 'No. Absen', 'Nama', 'Kelas', 'Nama File',
        'Nilai', 'Predikat', 'Sesuai', 'Peringatan', 'Belum Sesuai',
        'Ringkasan Pemeriksaan (JSON)'
      ]);
      sheet.setFrozenRows(1);
    }

    const required = ['id', 'absen', 'nama', 'kelas', 'file', 'nilai'];
    for (const key of required) {
      if (data[key] === undefined || data[key] === null || data[key] === '') {
        return jsonOutput({ ok: false, error: 'Missing field: ' + key });
      }
    }

    // Batas sederhana untuk mencegah sel terlalu besar.
    const summary = JSON.stringify(Array.isArray(data.ringkasan) ? data.ringkasan : []);
    const safeSummary = summary.length > 40000 ? summary.slice(0, 40000) : summary;
    const timestamp = data.ts ? new Date(Number(data.ts)) : new Date();

    // Cegah kiriman ID yang sama tercatat berulang.
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, 2, lastRow - 1, 1).getDisplayValues().flat();
      if (ids.includes(String(data.id))) {
        return jsonOutput({ ok: true, duplicate: true });
      }
    }

    sheet.appendRow([
      timestamp,
      String(data.id).slice(0, 80),
      String(data.absen).slice(0, 10),
      String(data.nama).slice(0, 100),
      String(data.kelas).slice(0, 40),
      String(data.file).slice(0, 180),
      Math.max(0, Math.min(100, Number(data.nilai) || 0)),
      String(data.predikat || '').slice(0, 5),
      Number(data.sesuai) || 0,
      Number(data.peringatan) || 0,
      Number(data.belum) || 0,
      safeSummary
    ]);
    return jsonOutput({ ok: true });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

function doGet() {
  return jsonOutput({
    ok: true,
    service: 'Cek Laporan - rekap',
    message: 'Endpoint aktif. Kiriman laporan dilakukan melalui website.'
  });
}

function jsonOutput(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
