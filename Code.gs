/****************************************************
 * BUKU TAMU / RSVP - UNDANGAN IAN & AYU
 * Backend: Google Sheets + Apps Script
 *
 * CARA PAKAI (5 menit, cukup sekali):
 * 1. Buat Google Spreadsheet baru (misal nama: Buku Tamu Ian Ayu).
 * 2. Di BARIS PERTAMA (header) tulis 5 kolom ini:
 *      A = Timestamp | B = Nama | C = Ucapan | D = Kehadiran | E = Tiket
 *    (kolom E = kode unik tiap kiriman, untuk mencegah ucapan ganda)
 * 3. Menu Extensions > Apps Script, hapus semua isi editor,
 *    lalu copy-paste SELURUH kode file ini ke sana. Save.
 * 4. Kalau nama sheet-nya bukan "Sheet1", sesuaikan SHEET_NAME di bawah.
 * 5. Deploy > New deployment > tipe "Web app":
 *      - Execute as : Me (email kamu)
 *      - Who has access : Anyone
 *    Klik Deploy > Authorize access > copy "Web app URL"
 *    (bentuknya https://script.google.com/macros/s/XXXX/exec)
 * 6. Di file index.html cari tulisan:
 *      RSVP_SCRIPT_URL = "GANTI_DENGAN_URL_WEB_APP_ANDA"
 *    ganti dengan Web app URL tadi. Save + upload ulang ke GitHub.
 * 7. Tes: buka undangan, kirim 1 ucapan, cek spreadsheet
 *    otomatis nambah 1 baris baru berisi tanggal & jam kirim.
 *
 * CATATAN:
 * - Setiap ucapan otomatis tercatat TANGGAL & JAM (kolom Timestamp).
 * - Daftar ucapan tampil di undangan dari yang TERBARU.
 * - Kalau ganti kode ini nanti: Deploy > Manage deployments > Edit >
 *   Version: New version > Deploy (URL tetap sama, tidak perlu ganti lagi).
 *   PENTING: setelah update kode WAJIB deploy sebagai New version,
 *   kalau tidak undangan tetap menjalankan kode lama.
 ****************************************************/

const SHEET_NAME = 'Sheet1'; // <-- ganti kalau nama sheet beda

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const p = (e && e.parameter) || {};

  // MODE SIMPAN: dipanggil saat tamu klik Kirim
  // contoh: ?action=add&nama=Ibu+Ani&ucapan=Selamat&kehadiran=Hadir&tiket=abc123
  // Setiap kiriman membawa kode unik (tiket). Kalau tiket yang sama
  // datang 2x (request ganda / klik ganda / retry browser), baris kedua
  // dan seterusnya DITOLAK sehingga tidak ada ucapan dobel.
  if (p.action === 'add') {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      const tiket = String(p.tiket || '');
      if (tiket !== '') {
        const lastT = sheet.getLastRow();
        if (lastT >= 2) {
          const tickets = sheet.getRange(2, 5, lastT - 1, 1).getValues();
          for (let i = 0; i < tickets.length; i++) {
            if (String(tickets[i][0]) === tiket) return json({ ok: true, duplikat: true });
          }
        }
      }
      sheet.appendRow([new Date(), p.nama || '', p.ucapan || '', p.kehadiran || '', tiket]);
      return json({ ok: true });
    } finally {
      lock.releaseLock();
    }
  }

  // MODE BACA: kembalikan daftar ucapan, terbaru dulu
  const last = sheet.getLastRow();
  const out = [];
  if (last >= 2) {
    const data = sheet.getRange(2, 1, last - 1, 4).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      out.push({
        nama: String(data[i][1] || ''),
        ucapan: String(data[i][2] || ''),
        kehadiran: String(data[i][3] || ''),
        waktu: data[i][0] ? new Date(data[i][0]).toISOString() : ''
      });
    }
  }
  return json(out);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
