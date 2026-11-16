/**
 * =========================================================================
 * BACKEND GOOGLE APPS SCRIPT - POINKU & POINKU DISIPLIN
 * =========================================================================
 * Script ini berfungsi sebagai API Web Service untuk menghubungkan 
 * aplikasi POINKU PRO / POINKU DISIPLIN dengan Google Spreadsheet.
 *
 * CARA DEPLOY:
 * 1. Simpan script ini di Apps Script (Ekstensi > Apps Script).
 * 2. Klik tombol "Terapkan (Deploy)" > "Penerapan Baru (New Deployment)".
 * 3. Pilih tipe: "Aplikasi Web (Web App)".
 * 4. Jalankan sebagai: "Saya (Me)".
 * 5. Siapa yang memiliki akses: "Siapa saja (Anyone)".
 * 6. Salin URL Web App yang berakhiran /exec.
 * =========================================================================
 */

function getSpreadsheetInstance(sheetId) {
  if (sheetId && String(sheetId).trim() !== '') {
    try {
      return SpreadsheetApp.openById(String(sheetId).trim());
    } catch (err) {
      throw new Error("Gagal membuka Spreadsheet ID (" + sheetId + "). Pastikan ID benar dan akun Google yang men-deploy script memiliki hak akses ke spreadsheet tersebut: " + err.message);
    }
  }
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}
  throw new Error("Parameter 'sheetId' wajib disertakan atau script harus terpasang di Spreadsheet aktif!");
}

function doGet(e) {
  try {
    const action = (e && e.parameter) ? e.parameter.action : '';
    const sheetId = (e && e.parameter) ? e.parameter.sheetId : '';

    const ss = getSpreadsheetInstance(sheetId);
    
    if (action === "read" || action === "getData") {
      const data = readAllData(ss);
      return createJsonResponse(data);
    }

    return createJsonResponse({ 
      status: "ready", 
      message: "API POINKU Mandiri aktif dan siap menerima data untuk Spreadsheet: " + ss.getName(),
      sheetName: ss.getName()
    });
  } catch (err) {
    return createJsonResponse({ error: err.toString() });
  }
}

function doPost(e) {
  try {
    let rawData = "";
    if (e && e.parameter && e.parameter.data) {
      rawData = e.parameter.data;
    } else if (e && e.postData && e.postData.contents) {
      rawData = e.postData.contents;
    }

    if (!rawData) {
      return createJsonResponse({ error: "Tidak ada data yang diterima (Payload kosong)!" });
    }

    const payload = JSON.parse(rawData);
    const sheetId = payload.sheetId;

    const ss = getSpreadsheetInstance(sheetId);
    saveAllData(ss, payload);

    return createJsonResponse({ success: true, message: "Data berhasil disimpan ke Google Sheet: " + ss.getName() });
  } catch (err) {
    return createJsonResponse({ error: err.toString() });
  }
}

// -------------------------------------------------------------------------
// FUNGSI MEMBACA DATA DARI SPREADSHEET
// -------------------------------------------------------------------------
function readAllData(ss) {
  const result = {
    schoolData: {},
    students: [],
    logs: [],
    behaviors: [],
    deletedLogIds: [],
    deletedStudentNis: [],
    deletedBehaviorIds: []
  };

  // 1. Baca Konfigurasi System & Profil
  const configSheet = ss.getSheetByName("_CONFIG");
  if (configSheet) {
    const configData = configSheet.getDataRange().getValues();
    configData.forEach(row => {
      const key = String(row[0] || '').trim();
      const val = row[1];
      if (key === 'schoolData' && val) {
        try { result.schoolData = JSON.parse(val); } catch(e) {}
      } else if (key === 'deletedLogIds' && val) {
        try { result.deletedLogIds = JSON.parse(val); } catch(e) {}
      } else if (key === 'deletedStudentNis' && val) {
        try { result.deletedStudentNis = JSON.parse(val); } catch(e) {}
      } else if (key === 'deletedBehaviorIds' && val) {
        try { result.deletedBehaviorIds = JSON.parse(val); } catch(e) {}
      }
    });
  }

  // 2. Baca Data Siswa
  const siswaSheet = ss.getSheetByName("Siswa");
  if (siswaSheet && siswaSheet.getLastRow() > 1) {
    const rows = siswaSheet.getRange(2, 1, siswaSheet.getLastRow() - 1, 4).getValues();
    result.students = rows.filter(r => r[0] && r[1]).map(r => ({
      nis: String(r[0]).trim(),
      name: String(r[1]).trim(),
      kelas: String(r[2] || '').trim(),
      dob: r[3] instanceof Date ? Utilities.formatDate(r[3], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[3] || '').trim(),
      photo: "",
      studentPhoto: "",
      teacherCoverPhoto: "",
      studentCoverPhoto: ""
    }));
  }

  // 3. Baca Riwayat Poin / Pelanggaran
  const logSheet = ss.getSheetByName("Riwayat_Log");
  if (logSheet && logSheet.getLastRow() > 1) {
    const rows = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 8).getValues();
    result.logs = rows.filter(r => r[0] && r[2]).map(r => ({
      id: r[0],
      date: r[1] instanceof Date ? r[1].toISOString() : String(r[1] || ''),
      studentId: String(r[2]).trim(),
      studentName: String(r[3] || '').trim(),
      behaviorLabel: String(r[4] || '').trim(),
      points: Number(r[5] || 0),
      type: String(r[6] || 'minus').trim(),
      note: String(r[7] || '').trim()
    }));
  }

  // 4. Baca Master Poin
  const behaviorSheet = ss.getSheetByName("Master_Poin");
  if (behaviorSheet && behaviorSheet.getLastRow() > 1) {
    const rows = behaviorSheet.getRange(2, 1, behaviorSheet.getLastRow() - 1, 6).getValues();
    result.behaviors = rows.filter(r => r[0] && r[1]).map(r => ({
      id: r[0],
      label: String(r[1]).trim(),
      points: Number(r[2] || 0),
      type: String(r[3] || 'minus').trim(),
      icon: String(r[4] || 'fa-star').trim(),
      category: String(r[5] || '').trim()
    }));
  }

  return result;
}

// -------------------------------------------------------------------------
// FUNGSI MENYIMPAN DATA KE SPREADSHEET
// -------------------------------------------------------------------------
function saveAllData(ss, payload) {
  // 1. Simpan Config & System State
  let configSheet = ss.getSheetByName("_CONFIG");
  if (!configSheet) {
    configSheet = ss.insertSheet("_CONFIG");
  }
  configSheet.clear();
  const configRows = [
    ["KEY", "VALUE_JSON"],
    ["schoolData", JSON.stringify(payload.schoolData || {})],
    ["deletedLogIds", JSON.stringify(payload.deletedLogIds || [])],
    ["deletedStudentNis", JSON.stringify(payload.deletedStudentNis || [])],
    ["deletedBehaviorIds", JSON.stringify(payload.deletedBehaviorIds || [])],
    ["lastSaved", new Date().toISOString()]
  ];
  configSheet.getRange(1, 1, configRows.length, 2).setValues(configRows);

  // 2. Simpan Sheet Siswa
  let siswaSheet = ss.getSheetByName("Siswa");
  if (!siswaSheet) {
    siswaSheet = ss.insertSheet("Siswa");
  }
  siswaSheet.clear();
  const siswaHeader = [["NIS", "NAMA SISWA", "KELAS", "TANGGAL LAHIR"]];
  const siswaRows = (payload.students || []).map(s => [
    String(s.nis || '').trim(),
    String(s.name || '').trim(),
    String(s.kelas || '').trim(),
    String(s.dob || '').trim()
  ]);
  const allSiswa = siswaHeader.concat(siswaRows);
  siswaSheet.getRange(1, 1, allSiswa.length, 4).setValues(allSiswa);
  formatHeaderSheet(siswaSheet, 4, "#2563EB"); // Header Biru

  // 3. Simpan Sheet Riwayat Log
  let logSheet = ss.getSheetByName("Riwayat_Log");
  if (!logSheet) {
    logSheet = ss.insertSheet("Riwayat_Log");
  }
  logSheet.clear();
  const logHeader = [["ID LOG", "TANGGAL", "NIS", "NAMA SISWA", "AKTIVITAS / PELANGGARAN", "POIN", "TIPE", "CATATAN GURU"]];
  const logRows = (payload.logs || []).map(l => [
    l.id || Date.now(),
    String(l.date || ''),
    String(l.studentId || '').trim(),
    String(l.studentName || '').trim(),
    String(l.behaviorLabel || '').trim(),
    Number(l.points || 0),
    String(l.type || 'minus').trim(),
    String(l.note || '').trim()
  ]);
  const allLogs = logHeader.concat(logRows);
  logSheet.getRange(1, 1, allLogs.length, 8).setValues(allLogs);
  formatHeaderSheet(logSheet, 8, "#E11D48"); // Header Merah / Rose

  // 4. Simpan Sheet Master Poin
  let behaviorSheet = ss.getSheetByName("Master_Poin");
  if (!behaviorSheet) {
    behaviorSheet = ss.insertSheet("Master_Poin");
  }
  behaviorSheet.clear();
  const behaviorHeader = [["ID", "LABEL AKTIVITAS", "POIN", "TIPE", "IKON", "KATEGORI"]];
  const behaviorRows = (payload.behaviors || []).map(b => [
    b.id,
    String(b.label || '').trim(),
    Number(b.points || 0),
    String(b.type || 'minus').trim(),
    String(b.icon || 'fa-star').trim(),
    String(b.category || '').trim()
  ]);
  const allBehaviors = behaviorHeader.concat(behaviorRows);
  behaviorSheet.getRange(1, 1, allBehaviors.length, 6).setValues(allBehaviors);
  formatHeaderSheet(behaviorSheet, 6, "#7C3AED"); // Header Ungu
}

// Format Header Tabel Spreadsheet
function formatHeaderSheet(sheet, numCols, hexColor) {
  const header = sheet.getRange(1, 1, 1, numCols);
  header.setFontWeight("bold");
  header.setFontColor("#FFFFFF");
  header.setBackground(hexColor);
  header.setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
