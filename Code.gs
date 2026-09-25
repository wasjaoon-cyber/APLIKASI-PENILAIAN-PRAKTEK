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

  if (lastRow < 2) {
    sh.appendRow(values);
    return;
  }

  // 1) Jika ID sama, perbarui data yang sudah ada.
  const rows = sh.getRange(2, 1, lastRow - 1, 15).getValues();
  const idIndex = rows.findIndex(row => String(row[0]) === id);

  if (idIndex !== -1) {
    sh.getRange(idIndex + 2, 1, 1, values.length).setValues([values]);
    return;
  }

  // 2) Cegah duplikat berdasarkan:
  //    NIS/NISN + Kelas + Praktik + Nama Tugas + Tanggal.
  //    Jika NIS kosong, gunakan Nama + Kelas sebagai identitas.
  const newKey = makePracticeKey_(
    r.name,
    r.nis,
    r.kelas,
    r.praktik,
    r.tugas,
    r.tanggal
  );

  const duplicateIndex = rows.findIndex(row => {
    const existingKey = makePracticeKey_(
      row[1],
      row[2],
      row[3],
      row[4],
      row[5],
      row[6]
    );
    return existingKey === newKey;
  });

  if (duplicateIndex !== -1) {
    // Data yang sama ditemukan: UPDATE, jangan tambah baris baru.
    sh.getRange(duplicateIndex + 2, 1, 1, values.length)
      .setValues([values]);
    return;
  }

  // 3) Jika benar-benar penilaian baru, tambahkan baris.
  sh.appendRow(values);
}

// Membuat kunci unik penilaian.
function makePracticeKey_(name, nis, kelas, praktik, tugas, tanggal) {
  const identitas = String(nis || "").trim()
    ? String(nis).trim()
    : String(name || "").trim();

  return [
    normalizeKey_(identitas),
    normalizeKey_(kelas),
    normalizeKey_(praktik),
    normalizeKey_(tugas),
    normalizeKey_(formatDate_(tanggal))
  ].join("|");
}

function normalizeKey_(value) {
  return String(value == null ? "" : value)
    .trim()
    .toLowerCase()
    .replace(/\\s+/g, " ");
}

// Jalankan fungsi ini secara manual sekali jika ingin membersihkan
// data duplikat yang SUDAH terlanjur ada di sheet.
// Yang dipertahankan adalah baris terakhir dari setiap duplikat.
function bersihkanDuplikatPenilaian() {
  const sh = getPracticeSheet_();
  const lastRow = sh.getLastRow();

  if (lastRow < 3) {
    Logger.log("Tidak ada duplikat yang perlu dibersihkan.");
    return;
  }

  const rows = sh.getRange(2, 1, lastRow - 1, 15).getValues();
  const seen = {};
  const rowsToDelete = [];

  rows.forEach((row, i) => {
    const key = makePracticeKey_(
      row[1],
      row[2],
      row[3],
      row[4],
      row[5],
      row[6]
    );

    if (!key || key === "||||") return;

    if (seen[key] !== undefined) {
      // Hapus baris lama dan pertahankan baris yang paling baru.
      rowsToDelete.push(seen[key] + 2);
    }

    seen[key] = i;
  });

  // Hapus dari bawah ke atas agar nomor baris tidak bergeser.
  rowsToDelete.sort((a, b) => b - a).forEach(rowNumber => {
    sh.deleteRow(rowNumber);
  });

  Logger.log(
    "Pembersihan selesai. Baris duplikat yang dihapus: " +
    rowsToDelete.length
  );
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


// ============================================================
// DATABASE SISWA TSM
// Data berasal dari daftar siswa yang sudah ada di aplikasi.
// Jalankan isiDatabaseSiswa() SATU KALI untuk mengisi sheet Siswa.
// ============================================================
function isiDatabaseSiswa() {
  const sh = getStudentSheet_();

  const siswa = [{"nama":"AKBAR FARIZ AL AYUBI","nis":"'0116249039","kelas":"10 TSM 1"},{"nama":"BAGAS SAPUTRA","nis":"'0108422243","kelas":"10 TSM 1"},{"nama":"DAMAR SADEWO PRATAMA","nis":"'0095007446","kelas":"10 TSM 1"},{"nama":"DANDI NAINGGOLAN","nis":"'0111380135","kelas":"10 TSM 1"},{"nama":"FAHMY SHIHAB","nis":"'0105880815","kelas":"10 TSM 1"},{"nama":"FEBRIYANTO CHRYSTIAN","nis":"'0117485746","kelas":"10 TSM 1"},{"nama":"GIAN PUTRA ADI VIANO","nis":"'0113363008","kelas":"10 TSM 1"},{"nama":"HIJATUL ARYA PERMANA","nis":"'0102006872","kelas":"10 TSM 1"},{"nama":"ILHAM BAHRUDIN","nis":"'0101563750","kelas":"10 TSM 1"},{"nama":"LEONARDO GYAN LUIS HARTO","nis":"'0103056497","kelas":"10 TSM 1"},{"nama":"MAULANA YUSUF WIJAYANTO","nis":"'0114299937","kelas":"10 TSM 1"},{"nama":"MOCHAMAD SOHIB HUSEIN","nis":"'0115815585","kelas":"10 TSM 1"},{"nama":"MOHAMMAD FADHLI ZAIN","nis":"'0116841138","kelas":"10 TSM 1"},{"nama":"MUAHAMAD RIZKI PAIRUZ","nis":"'3104978510","kelas":"10 TSM 1"},{"nama":"MUHAMAD ALIYUDIN","nis":"'0119858313","kelas":"10 TSM 1"},{"nama":"MUHAMAD ARYA SHABILA","nis":"'0112816471","kelas":"10 TSM 1"},{"nama":"MUHAMAD DOI PUSALAM","nis":"'0102423189","kelas":"10 TSM 1"},{"nama":"MUHAMAD FARID AKBAR","nis":"'0101611766","kelas":"10 TSM 1"},{"nama":"MUHAMAD RAIHAN ALHAFIDZ","nis":"'0108090654","kelas":"10 TSM 1"},{"nama":"MUHAMAD ROFII ZAHRAN","nis":"'0101020227","kelas":"10 TSM 1"},{"nama":"MUHAMMAD ALIF ZAIDAN ALAMSYAH","nis":"'0118769324","kelas":"10 TSM 1"},{"nama":"MUHAMMAD ARSYAD","nis":"'0103271475","kelas":"10 TSM 1"},{"nama":"MUHAMMAD KHAIRUL AMRI","nis":"'0113770421","kelas":"10 TSM 1"},{"nama":"MUHAMMAD ROUF ALFAHREZI","nis":"'0112590578","kelas":"10 TSM 1"},{"nama":"OKTAVIAN SAPUTRA","nis":"'0113418458","kelas":"10 TSM 1"},{"nama":"PRAMA SURYA KUSUMA","nis":"'0116667456","kelas":"10 TSM 1"},{"nama":"REYFAN NOVEL RIZKYAWAN","nis":"'0101284720","kelas":"10 TSM 1"},{"nama":"RIZKY ADITYA PUTRA","nis":"'0119283875","kelas":"10 TSM 1"},{"nama":"SENO ARDIANSYAH","nis":"'0109952599","kelas":"10 TSM 1"},{"nama":"ADITYA NAUFAL DARY ABIYYU","nis":"'0113278366","kelas":"10 TSM 2"},{"nama":"AKHNAF ZAKY PRATAMA","nis":"'0106332999","kelas":"10 TSM 2"},{"nama":"ALBERT KRISTIAN SITORUS","nis":"'0117277486","kelas":"10 TSM 2"},{"nama":"ARZHAT ADITYA ALFAROHMI","nis":"'0115753652","kelas":"10 TSM 2"},{"nama":"BINTANG DHIMAS ADITYA","nis":"'0115630172","kelas":"10 TSM 2"},{"nama":"DIMAS ALIF FIANSYAH","nis":"'0109139729","kelas":"10 TSM 2"},{"nama":"FAHRI RAMADAN","nis":"'0116676019","kelas":"10 TSM 2"},{"nama":"FARID HARTONO","nis":"'0105395836","kelas":"10 TSM 2"},{"nama":"FAZIAN DIRGANTARA RAMADHAN","nis":"'0109201293","kelas":"10 TSM 2"},{"nama":"HAIQAL LUTFIYANA HERNAWAN","nis":"'3114909026","kelas":"10 TSM 2"},{"nama":"JIBRIL MAULANA","nis":"'0117163322","kelas":"10 TSM 2"},{"nama":"KENNY HASSAN MAXIMA","nis":"'0116021030","kelas":"10 TSM 2"},{"nama":"KENZIE DWI ALVARO","nis":"'0103862561","kelas":"10 TSM 2"},{"nama":"MUHAMAD ABDUL ROUF","nis":"'3114331621","kelas":"10 TSM 2"},{"nama":"MUHAMAD GUNTUR","nis":"'0111818102","kelas":"10 TSM 2"},{"nama":"MUHAMAD IKBAL JAMALULLAEL","nis":"'0104973826","kelas":"10 TSM 2"},{"nama":"MUHAMAD NURFAZRIL NASUTION","nis":"'0105999575","kelas":"10 TSM 2"},{"nama":"MUHAMAD RAPLI","nis":"'0118659478","kelas":"10 TSM 2"},{"nama":"MUHAMAD RASYA ALAMSYAH","nis":"'0112590232","kelas":"10 TSM 2"},{"nama":"MUHAMAD ZAQI ZAMZAMI","nis":"'0118694610","kelas":"10 TSM 2"},{"nama":"MUHAMMAD FADHIL ARKANDIRA","nis":"'0113131831","kelas":"10 TSM 2"},{"nama":"MUHAMMAD FAHMI RAMADHAN","nis":"'0109812662","kelas":"10 TSM 2"},{"nama":"MUHAMMAD RAMADANU","nis":"'0118696004","kelas":"10 TSM 2"},{"nama":"NANDA SEPTIAN","nis":"'3114191237","kelas":"10 TSM 2"},{"nama":"NAUFAL ABDILLAH SYAFIQ","nis":"'0101088989","kelas":"10 TSM 2"},{"nama":"NOVAL ILMANSAH","nis":"'0072650870","kelas":"10 TSM 2"},{"nama":"RADERO MELVIN LIMBONG","nis":"'0108454480","kelas":"10 TSM 2"},{"nama":"RASYA ALVIANSYAH GUNAWAN","nis":"'0118835997","kelas":"10 TSM 2"},{"nama":"RA`IF RAHARDIAN ANAQI","nis":"'0107880958","kelas":"10 TSM 2"},{"nama":"RIZKY RIO RAMADHAN","nis":"'3117662595","kelas":"10 TSM 2"},{"nama":"RIZQY APRIANSYAH","nis":"'3114614942","kelas":"10 TSM 2"},{"nama":"WAHYU HIDAYAT","nis":"'3113711298","kelas":"10 TSM 2"},{"nama":"ABDI ALDIANSAH","nis":"'1.25.510","kelas":"11 TSM 1"},{"nama":"ADI SETIAWAN EFENDI","nis":"'1.25.511","kelas":"11 TSM 1"},{"nama":"AHMAD NOVAL KURNIA","nis":"'1.25.512","kelas":"11 TSM 1"},{"nama":"AHMAD RAMA IBRAHIM","nis":"'1.25.513","kelas":"11 TSM 1"},{"nama":"ALEX SAPUTRA","nis":"'1.25.515","kelas":"11 TSM 1"},{"nama":"BAYU TRI WAHONO","nis":"'1.25.516","kelas":"11 TSM 1"},{"nama":"DICKY FARIS KHALWANI","nis":"'1.25.517","kelas":"11 TSM 1"},{"nama":"ELANG HAVIZ AL HABSY","nis":"'1.25.518","kelas":"11 TSM 1"},{"nama":"FACHRI TAUFIQ SAPUTRA","nis":"'1.25.519","kelas":"11 TSM 1"},{"nama":"FARRIS SHERARD ALGHIFARY","nis":"'1.25.520","kelas":"11 TSM 1"},{"nama":"HAFIZH ABDILLAH","nis":"'1.25.521","kelas":"11 TSM 1"},{"nama":"HESQIEL JAWASAM NAINGGOLAN","nis":"'1.25.522","kelas":"11 TSM 1"},{"nama":"IHSAN ADI FADHILAH","nis":"'1.25.523","kelas":"11 TSM 1"},{"nama":"IRFAN DWI ANGGARA","nis":"'1.25.524","kelas":"11 TSM 1"},{"nama":"JOLOI ALBERTO SIHOMBING","nis":"'1.25.525","kelas":"11 TSM 1"},{"nama":"MIFTAH HUDIN","nis":"'1.25.526","kelas":"11 TSM 1"},{"nama":"MUHAMAD FAHRI ERIYANSYAH","nis":"'1.25.528","kelas":"11 TSM 1"},{"nama":"MUHAMAD RIKI SAPUTRA","nis":"'1.25.529","kelas":"11 TSM 1"},{"nama":"MUHAMAD WILDAN MUZAQI","nis":"'1.25.531","kelas":"11 TSM 1"},{"nama":"MUHAMMAD AZKA AL HAFIDZ","nis":"'1.25.533","kelas":"11 TSM 1"},{"nama":"MUHAMMAD DAFFA MUSYAFA","nis":"'1.25.534","kelas":"11 TSM 1"},{"nama":"MUHAMMAD RAMADHAN","nis":"'1.25.535","kelas":"11 TSM 1"},{"nama":"MUHAMMAD RIFQI NUR AZMI","nis":"'1.25.536","kelas":"11 TSM 1"},{"nama":"MUHAMMAD YUSUF ALGINA","nis":"'1.25.537","kelas":"11 TSM 1"},{"nama":"MULKHAN DICKY KURNIAWAN","nis":"'1.25.538","kelas":"11 TSM 1"},{"nama":"NATHAN IBRAHIMOVIC PERMANA","nis":"'1.25.539","kelas":"11 TSM 1"},{"nama":"NEBULA BRIGHT IMANAR","nis":"'1.25.540","kelas":"11 TSM 1"},{"nama":"RAFA AGUSTIAN RAMADHAN","nis":"'1.25.541","kelas":"11 TSM 1"},{"nama":"RIZKI ARYA PUTRA","nis":"'1.25.542","kelas":"11 TSM 1"},{"nama":"SYIFA NURUL AUDINA","nis":"'1.25.544","kelas":"11 TSM 1"},{"nama":"ZAKWAN GHAISAN HAFIZ","nis":"'1.25.545","kelas":"11 TSM 1"},{"nama":"ADAM ZIDAN MUCHTAMAR","nis":"'1.25.546","kelas":"11 TSM 2"},{"nama":"AHMAD BAGUS HIDAYATULLAH","nis":"'1.25.548","kelas":"11 TSM 2"},{"nama":"AHMAD FAISAL","nis":"'1.25.549","kelas":"11 TSM 2"},{"nama":"AHMAD MAULANA HASANUDIN","nis":"'1.25.550","kelas":"11 TSM 2"},{"nama":"ALGHIFAHRI NUNGGAL CAHYA","nis":"'1.25.551","kelas":"11 TSM 2"},{"nama":"ARKAN ZAKI ISMAIL","nis":"'1.25.552","kelas":"11 TSM 2"},{"nama":"BAGAS TEGAR PAMUNGKAS","nis":"'1.25.553","kelas":"11 TSM 2"},{"nama":"DAFFA PRADITYA","nis":"'1.25.554","kelas":"11 TSM 2"},{"nama":"DAFFI PRADITYA","nis":"'1.25.555","kelas":"11 TSM 2"},{"nama":"DERA ARYANI","nis":"'1.25.556","kelas":"11 TSM 2"},{"nama":"FAJRI GILANG PRATAMA","nis":"'1.25.557","kelas":"11 TSM 2"},{"nama":"FARHAN GHANI","nis":"'1.25.558","kelas":"11 TSM 2"},{"nama":"HARLINO WIBOWO","nis":"'1.25.559","kelas":"11 TSM 2"},{"nama":"IRSYAD HASYIM","nis":"'1.25.560","kelas":"11 TSM 2"},{"nama":"MAD BADHAL","nis":"'1.25.561","kelas":"11 TSM 2"},{"nama":"MARIO MULYA PRAKASA","nis":"'1.25.562","kelas":"11 TSM 2"},{"nama":"MUHAMAD AL FARIDZI","nis":"'1.25.563","kelas":"11 TSM 2"},{"nama":"MUHAMAD NURRIZQI SAPUTRA","nis":"'1.25.564","kelas":"11 TSM 2"},{"nama":"MUHAMAD YUSUF EVANDA BILAL","nis":"'1.25.565","kelas":"11 TSM 2"},{"nama":"MUHAMMAD ABDUL LATHIIF","nis":"'1.25.566","kelas":"11 TSM 2"},{"nama":"MUHAMMAD ALVIN RIZQY","nis":"'1.25.567","kelas":"11 TSM 2"},{"nama":"MUHAMMAD AUFA RIZANTA","nis":"'1.25.568","kelas":"11 TSM 2"},{"nama":"MUHAMMAD HAIKAL PUTRA","nis":"'1.25.569","kelas":"11 TSM 2"},{"nama":"MUHAMMAD IRUL ARRAHMAN","nis":"'1.25.570","kelas":"11 TSM 2"},{"nama":"MUHAMMAD YODHA HAIDAR","nis":"'1.25.571","kelas":"11 TSM 2"},{"nama":"NAUFAL IBNU HIBBAN","nis":"'1.25.572","kelas":"11 TSM 2"},{"nama":"QORI SALSABILLA","nis":"'1.25.574","kelas":"11 TSM 2"},{"nama":"REGIANSYAH DWI KASBARANI","nis":"'1.25.575","kelas":"11 TSM 2"},{"nama":"REVAN SYAPUTRA","nis":"'1.25.576","kelas":"11 TSM 2"},{"nama":"RIZAL SYAHRUL PRATAMA","nis":"'1.25.577","kelas":"11 TSM 2"},{"nama":"RIZKI YANTO","nis":"'1.25.578","kelas":"11 TSM 2"},{"nama":"RIZKY JATIANSYAH","nis":"'1.25.579","kelas":"11 TSM 2"},{"nama":"RIZKY MAULANA SETIAWAN","nis":"'1.25.580","kelas":"11 TSM 2"},{"nama":"YUDA PRATAMA","nis":"'1.25.581","kelas":"11 TSM 2"},{"nama":"ABDUL HAFIZ RAFSANJANI","nis":"'1.24.576","kelas":"12 TSM 1"},{"nama":"ADE FIRMANSYAH","nis":"'1.24.577","kelas":"12 TSM 1"},{"nama":"AHMAD NURIL HUDA","nis":"'1.24.579","kelas":"12 TSM 1"},{"nama":"AHMAD SUHENDA","nis":"'1.24.580","kelas":"12 TSM 1"},{"nama":"ALDY JAYA HARTANTO","nis":"'1.24.581","kelas":"12 TSM 1"},{"nama":"ALFIANSYAH KHRISNA PRADHIKA","nis":"'1.24.582","kelas":"12 TSM 1"},{"nama":"ALIF GALIH NUSANTARA","nis":"'1.24.583","kelas":"12 TSM 1"},{"nama":"AMIR MUSLIM","nis":"'1.24.584","kelas":"12 TSM 1"},{"nama":"ANGGA ALDIAN NUGRAHA","nis":"'1.24.585","kelas":"12 TSM 1"},{"nama":"BAIKAL PAJRI","nis":"'1.24.587","kelas":"12 TSM 1"},{"nama":"DELON MUHANIF SATRIA","nis":"'1.24.588","kelas":"12 TSM 1"},{"nama":"DIAZ ANDREAN","nis":"'1.24.589","kelas":"12 TSM 1"},{"nama":"EZA HARTONO","nis":"'1.24.590","kelas":"12 TSM 1"},{"nama":"FACHRI RADITYA NAJIB","nis":"'1.24.591","kelas":"12 TSM 1"},{"nama":"FURQON AMRULLAH","nis":"'1.24.592","kelas":"12 TSM 1"},{"nama":"IVANDER TIMO","nis":"'1.24.593","kelas":"12 TSM 1"},{"nama":"LABIB NAZWAN","nis":"'1.24.594","kelas":"12 TSM 1"},{"nama":"MUHAMAD BAYU WAHYU ADAM","nis":"'1.24.596","kelas":"12 TSM 1"},{"nama":"MUHAMAD FAHRI BAIHAQI","nis":"'1.24.597","kelas":"12 TSM 1"},{"nama":"MUHAMAD IDEA TIRA MAHARDIKA","nis":"'1.24.598","kelas":"12 TSM 1"},{"nama":"MUHAMAD LUKMAN HAKIM","nis":"'1.24.599","kelas":"12 TSM 1"},{"nama":"MUHAMAD OKTA REJA","nis":"'1.24.600","kelas":"12 TSM 1"},{"nama":"MUHAMAD RAIHAN","nis":"'1.24.601","kelas":"12 TSM 1"},{"nama":"MUHAMAD SYAHRIL BADAWI","nis":"'1.24.602","kelas":"12 TSM 1"},{"nama":"MUHAMMAD FAUZAN","nis":"'1.24.603","kelas":"12 TSM 1"},{"nama":"MUHAMMAD RIDWAN ALIF","nis":"'1.24.604","kelas":"12 TSM 1"},{"nama":"MUHAMMAD TAAZUL A`RIF","nis":"'1.24.605","kelas":"12 TSM 1"},{"nama":"NAJWAN GILANG MUZAKKI","nis":"'1.24.606","kelas":"12 TSM 1"},{"nama":"RAHMAT AJI WIJAYA","nis":"'1.24.607","kelas":"12 TSM 1"},{"nama":"RESTU BAGUS ARDIYAN","nis":"'1.24.608","kelas":"12 TSM 1"},{"nama":"RIAN APRIYANSAH","nis":"'1.24.609","kelas":"12 TSM 1"},{"nama":"RISQI BAYU AJI","nis":"'1.24.610","kelas":"12 TSM 1"},{"nama":"SIRHAN MUZAFFAR","nis":"'1.24.611","kelas":"12 TSM 1"},{"nama":"YURI ARYA MAULANA","nis":"'1.24.612","kelas":"12 TSM 1"},{"nama":"ALFA RIZKI NUGRAHA","nis":"'1.24.613","kelas":"12 TSM 2"},{"nama":"ANGGA JULIA RACHMAN","nis":"'1.24.614","kelas":"12 TSM 2"},{"nama":"AQIL MUTTAQIN","nis":"'1.24.615","kelas":"12 TSM 2"},{"nama":"BAGUS HERMANSYAH","nis":"'1.24.616","kelas":"12 TSM 2"},{"nama":"DERRIEL NOVIANDRI YOKO","nis":"'1.24.617","kelas":"12 TSM 2"},{"nama":"DIKA JUNIOR","nis":"'1.24.618","kelas":"12 TSM 2"},{"nama":"DIMAS DWI PRADIFTA","nis":"'1.24.619","kelas":"12 TSM 2"},{"nama":"DIMAZ ADI PRAYOGO","nis":"'1.24.620","kelas":"12 TSM 2"},{"nama":"DWI ANGGORO","nis":"'1.24.621","kelas":"12 TSM 2"},{"nama":"EKA CIPTA WIJAYA","nis":"'1.24.622","kelas":"12 TSM 2"},{"nama":"HARRY FEBRIAN","nis":"'1.24.624","kelas":"12 TSM 2"},{"nama":"IBNU JIBRAN ABRORI","nis":"'1.24.625","kelas":"12 TSM 2"},{"nama":"KAFA GHOITS FAUZAN","nis":"'1.24.626","kelas":"12 TSM 2"},{"nama":"KEVIN APRIZI","nis":"'1.24.627","kelas":"12 TSM 2"},{"nama":"M. AQSO HAQUNA SURYA","nis":"'1.24.628","kelas":"12 TSM 2"},{"nama":"MUHAMAD ADRIAN","nis":"'1.24.629","kelas":"12 TSM 2"},{"nama":"MUHAMAD DWI PRANAYA","nis":"'1.24.631","kelas":"12 TSM 2"},{"nama":"MUHAMAD FACHRI","nis":"'1.24.632","kelas":"12 TSM 2"},{"nama":"MUHAMAD HAFID AL ASHARI","nis":"'1.24.633","kelas":"12 TSM 2"},{"nama":"MUHAMAD MURUR RIF`I","nis":"'1.24.634","kelas":"12 TSM 2"},{"nama":"MUHAMAD RAHMADANI","nis":"'1.24.635","kelas":"12 TSM 2"},{"nama":"MUHAMAD RAMADANI","nis":"'1.24.636","kelas":"12 TSM 2"},{"nama":"MUHAMAD RICKY KAKA","nis":"'1.24.637","kelas":"12 TSM 2"},{"nama":"MUHAMMAD FACHRIANTO","nis":"'1.24.638","kelas":"12 TSM 2"},{"nama":"MUHAMMAD ZENIT SULTHAN SHAQR","nis":"'1.24.639","kelas":"12 TSM 2"},{"nama":"RAFI MANDALA SAKTI","nis":"'1.24.640","kelas":"12 TSM 2"},{"nama":"RAMASYA ARYADIMAS FURQON","nis":"'1.24.641","kelas":"12 TSM 2"},{"nama":"RANDI JULIANTO","nis":"'1.24.642","kelas":"12 TSM 2"},{"nama":"RESTU MAULANA RIZKY","nis":"'1.24.643","kelas":"12 TSM 2"},{"nama":"REZZA FAHLEVI","nis":"'1.24.644","kelas":"12 TSM 2"},{"nama":"SURESSUL NURKHAQIQI","nis":"'1.24.645","kelas":"12 TSM 2"},{"nama":"SYAIFUL KARIM","nis":"'1.24.646","kelas":"12 TSM 2"},{"nama":"SYARIF FADHL","nis":"'1.24.647","kelas":"12 TSM 2"},{"nama":"TEGAR MUFTHI KHULUQI","nis":"'1.24.648","kelas":"12 TSM 2"},{"nama":"TEGUH ARYA ALDIANSYAH","nis":"'1.24.649","kelas":"12 TSM 2"}];

  // Bersihkan data lama, pertahankan header.
  sh.clearContents();
  sh.getRange(1, 1, 1, 3).setValues([["Nama", "NIS/NISN", "Kelas"]]);

  if (siswa.length > 0) {
    const values = siswa.map(s => [s.nama || "", s.nis || "", s.kelas || ""]);
    sh.getRange(2, 1, values.length, 3).setValues(values);
    // NIS/NISN diperlakukan sebagai teks agar angka 0 di depan tidak hilang.
    sh.getRange(2, 2, values.length, 1).setNumberFormat("@");
  }

  sh.getRange(1, 1, 1, 3)
    .setFontWeight("bold")
    .setBackground("#d9eaff");
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 3);

  Logger.log("Database siswa berhasil diisi: " + siswa.length + " siswa.");
  return siswa.length;
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
