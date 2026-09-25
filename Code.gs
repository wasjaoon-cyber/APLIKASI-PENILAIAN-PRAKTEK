// ============================================================
// APLIKASI PENILAIAN PRAKTIK TSM
// Google Apps Script - TERIKAT LANGSUNG DENGAN GOOGLE SHEET
// ============================================================

const PRACTICE_SHEET = "Nilai Praktik";
const STUDENT_SHEET = "Siswa";

// ------------------------------------------------------------
// Web App
// ------------------------------------------------------------
function doGet(e) {
  try {
    const action = e && e.parameter ? (e.parameter.action || "") : "";

    if (action === "getPractice") {
      return jsonOutput_({
        ok: true,
        data: getPracticeData_()
      });
    }

    if (action === "getStudents") {
      return jsonOutput_({
        ok: true,
        data: getStudentData_()
      });
    }

    return jsonOutput_({
      ok: true,
      message: "Web App Penilaian Praktik TSM aktif.",
      actions: ["getPractice", "getStudents"]
    });

  } catch (err) {
    return jsonOutput_({
      ok: false,
      message: err.message
    });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Data POST tidak ditemukan.");
    }

    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || "";

    if (action === "savePractice") {
      if (!payload.record) {
        throw new Error("Record penilaian tidak ditemukan.");
      }

      savePractice_(payload.record);

      return jsonOutput_({
        ok: true,
        message: "Penilaian berhasil disimpan.",
        id: payload.record.id
      });
    }

    throw new Error("Action tidak dikenali: " + action);

  } catch (err) {
    return jsonOutput_({
      ok: false,
      message: err.message
    });
  }
}

// ------------------------------------------------------------
// Spreadsheet
// Karena script ini dibuat dari Extensions > Apps Script
// pada Spreadsheet, getActiveSpreadsheet() digunakan.
// ------------------------------------------------------------
function getSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      "Spreadsheet tidak ditemukan. Pastikan Apps Script ini dibuat " +
      "dari Google Spreadsheet melalui Ekstensi > Apps Script."
    );
  }

  return ss;
}

// ------------------------------------------------------------
// Sheet Nilai Praktik
// ------------------------------------------------------------
function getPracticeSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(PRACTICE_SHEET);

  if (!sh) {
    sh = ss.insertSheet(PRACTICE_SHEET);
    sh.appendRow([
      "ID",
      "Nama",
      "NIS/NISN",
      "Kelas",
      "Praktik/Jobsheet",
      "Nama Tugas",
      "Tanggal",
      "Guru Penilai",
      "Nilai Akhir",
      "Status",
      "Kelebihan",
      "Kekurangan",
      "Saran",
      "Aspek JSON",
      "Waktu Simpan"
    ]);

    sh.getRange(1, 1, 1, 15)
      .setFontWeight("bold")
      .setBackground("#d9eaff");

    sh.setFrozenRows(1);
  }

  return sh;
}

// ------------------------------------------------------------
// Sheet Siswa
// ------------------------------------------------------------
function getStudentSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(STUDENT_SHEET);

  if (!sh) {
    sh = ss.insertSheet(STUDENT_SHEET);
    sh.appendRow([
      "Nama",
      "NIS/NISN",
      "Kelas"
    ]);

    sh.getRange(1, 1, 1, 3)
      .setFontWeight("bold")
      .setBackground("#d9eaff");

    sh.setFrozenRows(1);
  }

  return sh;
}

// ------------------------------------------------------------
// Simpan nilai praktik
// ------------------------------------------------------------
function savePractice_(r) {
  const sh = getPracticeSheet_();

  const id = String(r.id || new Date().getTime());

  const values = [
    id,
    r.name || "",
    r.nis || "",
    r.kelas || "",
    r.praktik || "",
    r.tugas || "",
    r.tanggal || "",
    r.guru || "",
    Number(r.score || 0),
    Number(r.score || 0) >= 75 ? "KOMPETEN" : "BELUM KOMPETEN",
    r.lebih || "",
    r.kurang || "",
    r.saran || "",
    JSON.stringify(r.aspek || []),
    new Date()
  ];

  const lastRow = sh.getLastRow();

  // Jika ID sudah ada, update baris tersebut.
  if (lastRow >= 2) {
    const ids = sh.getRange(2, 1, lastRow - 1, 1)
      .getDisplayValues()
      .flat();

    const index = ids.findIndex(x => String(x) === id);

    if (index !== -1) {
      sh.getRange(index + 2, 1, 1, values.length)
        .setValues([values]);
      return;
    }
  }

  sh.appendRow(values);
}

// ------------------------------------------------------------
// Ambil seluruh nilai praktik
// ------------------------------------------------------------
function getPracticeData_() {
  const sh = getPracticeSheet_();
  const lastRow = sh.getLastRow();

  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, 15).getValues();

  return values
    .filter(row => row[0] !== "")
    .map(row => {
      let aspek = [];

      try {
        aspek = row[13] ? JSON.parse(row[13]) : [];
      } catch (e) {
        aspek = [];
      }

      return {
        id: row[0],
        name: row[1],
        nis: row[2],
        kelas: row[3],
        praktik: row[4],
        tugas: row[5],
        tanggal: formatDate_(row[6]),
        guru: row[7],
        score: Number(row[8] || 0),
        lebih: row[10] || "",
        kurang: row[11] || "",
        saran: row[12] || "",
        aspek: aspek
      };
    });
}

// ------------------------------------------------------------
// Ambil data siswa
// ------------------------------------------------------------
function getStudentData_() {
  const sh = getStudentSheet_();
  const lastRow = sh.getLastRow();

  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, 3).getValues();

  return values
    .filter(row => row[0] !== "")
    .map(row => ({
      nama: row[0],
      nis: row[1],
      kelas: row[2]
    }));
}

// ------------------------------------------------------------
// Format tanggal
// ------------------------------------------------------------
function formatDate_(value) {
  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone() || "Asia/Jakarta",
      "yyyy-MM-dd"
    );
  }

  return String(value);
}

// ------------------------------------------------------------
// JSON response
// ------------------------------------------------------------
function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------
// TES KONEKSI
// Jalankan fungsi ini sekali dari editor Apps Script.
// ------------------------------------------------------------
function testKoneksiSpreadsheet() {
  const ss = getSpreadsheet_();

  Logger.log("Nama Spreadsheet: " + ss.getName());
  Logger.log("ID Spreadsheet: " + ss.getId());
  Logger.log("URL Spreadsheet: " + ss.getUrl());

  getPracticeSheet_();
  getStudentSheet_();

  Logger.log("Koneksi Spreadsheet berhasil.");
  Logger.log("Sheet Nilai Praktik dan Siswa sudah siap.");
}
