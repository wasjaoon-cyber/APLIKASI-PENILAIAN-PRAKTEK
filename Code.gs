const PRACTICE_SHEET = "Nilai Praktik";
const STUDENT_SHEET = "Siswa";

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || "";
    if (action === "getPractice") return json_({ok:true, data:getPractice_()});
    if (action === "getStudents") return json_({ok:true, data:getStudents_()});
    return json_({ok:true, message:"Web App Penilaian Praktik TSM aktif."});
  } catch(err) {
    return json_({ok:false, message:String(err)});
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (body.action === "savePractice") {
      savePractice_(body.record || {});
      return json_({ok:true, message:"Nilai praktik tersimpan."});
    }
    return json_({ok:false, message:"Action tidak dikenal."});
  } catch(err) {
    return json_({ok:false, message:String(err)});
  }
}

function getPracticeSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PRACTICE_SHEET);
  if (!sh) sh = ss.insertSheet(PRACTICE_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(["ID","Nama","NIS/NISN","Kelas","Tanggal","Praktik/Jobsheet","Nama Tugas",
      "Guru Penilai","Nilai Akhir","Status","Kelebihan","Kekurangan","Saran","Aspek JSON","Timestamp"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getStudentSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(STUDENT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(STUDENT_SHEET);
    sh.getRange(1,1,1,3).setValues([["Nama Siswa","NIS/NISN","Kelas"]]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function savePractice_(r) {
  const sh = getPracticeSheet_();
  const id = String(r.id || "");
  if (!id) throw new Error("ID penilaian kosong.");
  const row = [
    id,r.name||"",r.nis||"",r.kelas||"",r.tanggal||"",r.praktik||"",r.tugas||"",
    r.guru||"",Number(r.score||0),Number(r.score||0)>=75?"KOMPETEN":"BELUM KOMPETEN",
    r.lebih||"",r.kurang||"",r.saran||"",JSON.stringify(r.aspek||[]),new Date()
  ];
  const last = sh.getLastRow();
  if (last > 1) {
    const ids = sh.getRange(2,1,last-1,1).getValues().flat().map(String);
    const idx = ids.indexOf(id);
    if (idx >= 0) {
      sh.getRange(idx+2,1,1,row.length).setValues([row]);
      return;
    }
  }
  sh.appendRow(row);
}

function getPractice_() {
  const v = getPracticeSheet_().getDataRange().getValues();
  if (v.length <= 1) return [];
  return v.slice(1).filter(r=>r[0]).map(r=>({
    id:String(r[0]), name:r[1]||"", nis:r[2]||"", kelas:r[3]||"", tanggal:formatDate_(r[4]),
    praktik:r[5]||"", tugas:r[6]||"", guru:r[7]||"", score:Number(r[8]||0),
    lebih:r[10]||"", kurang:r[11]||"", saran:r[12]||"", aspek:parseJson_(r[13],[])
  }));
}

function getStudents_() {
  const v = getStudentSheet_().getDataRange().getValues();
  if (v.length <= 1) return [];
  return v.slice(1).filter(r=>r[0]).map(r=>({
    nama:String(r[0]||"").trim(),
    nis:String(r[1]||"").trim(),
    kelas:String(r[2]||"").trim()
  }));
}

function parseJson_(v,f){try{return v?JSON.parse(v):f}catch(e){return f}}

function formatDate_(v) {
  if (Object.prototype.toString.call(v)==="[object Date]" && !isNaN(v))
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return String(v||"");
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu("📊 Penilaian Praktik TSM")
    .addItem("Buka Nilai Praktik","openPracticeSheet_")
    .addItem("Buka Daftar Siswa","openStudentSheet_")
    .addToUi();
}
function openPracticeSheet_(){getPracticeSheet_().activate()}
function openStudentSheet_(){getStudentSheet_().activate()}
