const SS_ID = "1VisdjTk21E5ufiR0UjimRUWDV5q2AVcj9ZLX4IKv00s";
const SHEET_LOGIN = "Pswd_Ext"; // kolom Username | Password (SHA256)
const SHEET_USER = "User_Ext";   // kolom Username,Nama,PT,Email


function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Web Input Data Maintenance')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ===== Hash password SHA256 =====
function hashPassword(pass) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pass)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// ===== LOGIN FAST =====
function loginUser(username, password) {

  username = username.toString().trim().toLowerCase();
  const hashedPass = hashPassword(password);

  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_LOGIN);
  const data = sh.getDataRange().getValues();

  // buat map username
  const userMap = {};

  for (let i = 1; i < data.length; i++) {

    let dbUser = data[i][0].toString().trim().toLowerCase();
    let sheetPass = data[i][1];
    let status = data[i][2] || "Nonaktif";

    // auto hash jika masih plain text
    if (sheetPass.length !== 64) {
      sheetPass = hashPassword(sheetPass);
      sh.getRange(i+1,2).setValue(sheetPass);
    }

    userMap[dbUser] = {
      password: sheetPass,
      status: status,
      originalName: data[i][0]
    };
  }

  const user = userMap[username];

  if (!user) return { status: "error" };

  if (user.password !== hashedPass)
    return { status: "error" };

  if (user.status !== "Aktif")
    return { status: "notactive", username: user.originalName };

  return { status: "success", username: user.originalName };
}


// ===== SIGN UP FAST =====
function registerUser(obj) {

  const username = obj.username.toString().trim().toLowerCase();

  const shUser = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_USER);
  const shLogin = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_LOGIN);

  const data = shLogin.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const dbUser = data[i][0].toString().trim().toLowerCase();
    if (dbUser === username) return { status: "exist" };
  }

  const hashedPass = hashPassword(obj.password);

  shUser.appendRow([username, obj.nama, obj.pt, obj.email]);
  shLogin.appendRow([username, hashedPass, "Nonaktif"]);

  return { status: "success" };
}


// ===== CHANGE PASSWORD FAST =====
function changePassword(username, oldPass, newPass) {

  username = username.toString().trim().toLowerCase();

  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_LOGIN);
  const data = sh.getDataRange().getValues();

  const oldHash = hashPassword(oldPass);
  const newHash = hashPassword(newPass);

  for (let i = 1; i < data.length; i++) {

    const dbUser = data[i][0].toString().trim().toLowerCase();
    let sheetPass = data[i][1];

    if (sheetPass.length !== 64) {
      sheetPass = hashPassword(sheetPass);
      sh.getRange(i+1,2).setValue(sheetPass);
    }

    if (dbUser === username && sheetPass === oldHash) {
      sh.getRange(i+1, 2).setValue(newHash);
      return { status: "success" };
    }
  }

  return { status: "wrong" };
}

function saveDraftFreon(data){
  const SS_ID = "1ZnxeeCyyrzkSsDnmmQ0CPELE-cr6RQk5h0dPHPYahj0";
  const SHEET_NAME = "Draft_Freon";
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_NAME);

  // Jika sheet kosong, buat header
  if(sh.getLastRow()===0){
    sh.appendRow([
      "ID","Section","Department","Tanggal Pengajuan","Tanggal Realisasi",
      "Nomor Aset","Deskripsi","Lokasi Pendinginan","Kapasitas","Freon","Maker",
      "Jenis Perubahan",
      "Ubah Freon Lama","Ubah Freon Baru",
      "Ubah Kapasitas Lama","Ubah Kapasitas Baru",
      "Ubah Lokasi Lama","Ubah Lokasi Baru"
    ]);
  }

  // ambil data perubahan
  const row = [
    data.id,
    data.section,
    data.department,
    data.tglPengajuan,
    data.tglRealisasi,
    data.noAset,
    data.deskripsi,
    data.lokasi,
    data.kapasitas,
    data.freon,
    data.maker,
    Array.isArray(data.jenis)?data.jenis.join(", "):data.jenis||"",
    data.lamaFreon||"",
    data.baruFreon||"",
    data.lamaKapasitas||"",
    data.baruKapasitas||"",
    data.lamaLokasi||"",
    data.baruLokasi||"",
    data.user||"",
    "Draft"
  ];

  sh.appendRow(row);
  return {status:"success"};
}

function ReqDataFreon(data){

  const SS_ID = "1ZnxeeCyyrzkSsDnmmQ0CPELE-cr6RQk5h0dPHPYahj0";
  const SHEET_NAME = "Req_Freon";
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_NAME);

  // Jika sheet kosong, buat header
  if (sh.getLastRow() === 0){
    sh.appendRow([
      "Seksi",
      "Alasan",
      "Req_by",
      "Log_Time",
      "Status"
    ]);
  }

  // waktu server (tidak bisa dimanipulasi user)
  const now = new Date();

  // susun row sesuai urutan header
  const row = [
    data.section,        // Seksi
    data.Bg_Req_Freon,   // Alasan
    data.user || "-",  // Req_by (fallback kalau kosong)
    now,                 // Log_Time
    "Open"               // Status default
  ];

  sh.appendRow(row);

  return { status:"success" };
}

function checkTicketExists(ticketId){
  const SS_ID = "1ZnxeeCyyrzkSsDnmmQ0CPELE-cr6RQk5h0dPHPYahj0";
  const SHEET_NAME = "Draft_Freon";
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_NAME);

  if(!sh) return false;

  const data = sh.getRange(2,1,sh.getLastRow()-1,1).getValues(); 
  // mulai baris 2, kolom 1 (ID)

  for(let i=0;i<data.length;i++){
    if(String(data[i][0]).trim() === String(ticketId).trim()){
      return true;
    }
  }

  return false;
}

function uploadFreonFile(ticketId, fileName, mimeType, base64Data){
  try {
    const FOLDER_ID = "1TzCX-cVuuC-RimDQaLbbXGCCjBRxA3YU"; // ganti dengan folder Drive tujuan
    const SS_ID = "1ZnxeeCyyrzkSsDnmmQ0CPELE-cr6RQk5h0dPHPYahj0";
    const SHEET_NAME = "Draft_Freon";

    // ===== VALIDASI TICKET ID =====
    if(!checkTicketExists(ticketId)){
      return {status:"error", message:"ID Tiket tidak ditemukan!"};
    }

    // ===== CEK FILE SUDAH ADA DI FOLDER =====
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const files = folder.getFilesByName(ticketId); // nama file = ticketId
    if(files.hasNext()){
      return {status:"error", message:"ID Tiket ini sudah pernah di-upload!"};
    }

    // ===== SIMPAN FILE KE DRIVE =====
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      mimeType,
      ticketId // nama file disamakan dengan ID Tiket
    );

    const file = folder.createFile(blob);

    // ===== BUAT LINK VIEW =====
    const viewLink = "https://drive.google.com/file/d/" + file.getId() + "/view";
    // ===== UPDATE SPREADSHEET =====
    const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_NAME);
    const data = sh.getDataRange().getValues();

    for(let i=1;i<data.length;i++){
      if(data[i][0] == ticketId){ // kolom A
        sh.getRange(i+1,20).setValue("Submit"); // kolom T
        sh.getRange(i+1,21).setValue(viewLink); // kolom U
        break;
      }
    }


    return {
      status: "success",
      message: "File berhasil di-upload",
      url: file.getUrl()
    };

  } catch(err) {
    // tangkap semua error agar JS tetap menerima response
    return {status:"error", message:"Terjadi kesalahan: "+err.message};
  }
}

function testAccess() {
  // memaksa script minta izin
  const shUser = SpreadsheetApp.openById("1VisdjTk21E5ufiR0UjimRUWDV5q2AVcj9ZLX4IKv00s").getSheetByName("User_Ext");
  const shLogin = SpreadsheetApp.openById("1VisdjTk21E5ufiR0UjimRUWDV5q2AVcj9ZLX4IKv00s").getSheetByName("Pswd_Ext");
  const shFreon = SpreadsheetApp.openById("1ZnxeeCyyrzkSsDnmmQ0CPELE-cr6RQk5h0dPHPYahj0").getSheetByName("Draft_Freon");
  const shCase1 = SpreadsheetApp.openById("12ndmYZvFVNaZCMoYxCTe-jGpovo3dz5uUidvrSflCa0").getSheetByName("Result");
  const shCase2 = SpreadsheetApp.openById("12ndmYZvFVNaZCMoYxCTe-jGpovo3dz5uUidvrSflCa0").getSheetByName("Quest");
   const shCase3 = SpreadsheetApp.openById("12ndmYZvFVNaZCMoYxCTe-jGpovo3dz5uUidvrSflCa0").getSheetByName("PDF_link");
 // Logger.log(shUser.getLastRow());
 // Logger.log(shLogin.getLastRow());
 // Logger.log(shFreon.getLastRow());
  Logger.log(shCase1.getLastRow());
  Logger.log(shCase2.getLastRow());
  Logger.log(shCase3.getLastRow());
}


function testFolderAccess() {
  const FOLDER_ID = "1TzCX-cVuuC-RimDQaLbbXGCCjBRxA3YU";
  const FOLDER_ID1 = "1JMhFYToeqQ5Jc8HckIQCi7sT_l9qeX5c";
  const folder = DriveApp.getFolderById(FOLDER_ID1);
  Logger.log(folder.getName());
}

function testCreateFile() {
  try {
    const folder = DriveApp.getFolderById("1TzCX-cVuuC-RimDQaLbbXGCCjBRxA3YU");
    const file = folder.createFile("test.txt", "Isi test file");
    Logger.log("File created: " + file.getUrl());
  } catch(e) {
    Logger.log("ERROR: " + e.toString());
  }
}

function forceAuth() {
  // Memaksa Apps Script minta izin Drive
  DriveApp.getRootFolder();
}

function testCheckTicket(){
  const ticketId = "20260304-073355-927"; // ganti dengan ID yang sudah tersimpan di sheet Draft_Freon kolom A
  const exists = checkTicketExists(ticketId);
  Logger.log("Ticket exist? " + exists);
}

function forceDriveCreateAccess() {
  const folder = DriveApp.getRootFolder(); // folder root
  const testFile = folder.createFile("test.txt", "Ini file test");
  Logger.log("File berhasil dibuat: " + testFile.getId());
}

function callAppSheetAPI(appId, tableName, action, rows) {
  const url = `https://api.appsheet.com/api/v2/apps/${appId}/tables/${tableName}/Action`;

  const payload = {
    Action: action, // "Add", "Edit", "Delete"
    Properties: {},
    Rows: rows
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      ApplicationAccessKey: "API_KEY_KAMU"
    },
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(url, options);
  return response.getContentText();
}

function getFreonListWithNotes() {

  const SS_ID = "1ZnxeeCyyrzkSsDnmmQ0CPELE-cr6RQk5h0dPHPYahj0";
  const SHEET_NAME = "Type_Freon";

  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(SHEET_NAME);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const data = sh.getRange(2, 1, lastRow - 1, 7).getValues();

  return data.map(r => {

    const safe = (v) => (v === null || v === undefined) ? "" : v.toString().trim();
    const name = safe(r[0]);
    const chem = safe(r[1]);
    const app  = safe(r[2]);
    const gwp  = safe(r[3]);
    const odp  = safe(r[4]);
    const flam = safe(r[5]);
    const note = safe(r[6]);

    return {
      name,
      note: `
Chemistry: ${chem}
Application: ${app}
GWP: ${gwp}
ODP: ${odp}
Flammability: ${flam}
Note: ${note}
      `.trim()
    };

  }).filter(x => x.name !== "");
}

function getExamDataDynamic(cfg){

  const ss = cfg.spreadsheetId
    ? SpreadsheetApp.openById(cfg.spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();

  const qSheet = ss.getSheetByName(cfg.questionSheet);
  const sSheet = ss.getSheetByName(cfg.settingSheet);

  const qData = qSheet.getDataRange().getValues();
  const sData = sSheet.getDataRange().getValues();

  let pdfUrl = "";

  sData.slice(1).forEach(r=>{
    if(r[0] === cfg.category){
      pdfUrl = r[1];
    }
  });

  const questions = qData.slice(1)
    .filter(r => r[6] === cfg.category)
    .map((r,i) => ({
      id: i,           // 🔥 WAJIB
      q: r[0],
      a: r[1],
      b: r[2],
      c: r[3],
      d: r[4],
      correct: r[5]    // 🔥 konsisten
    }));

  return { pdfUrl, questions };
}

function saveExamDynamic(data){

  const cfg = data.config;

  const ss = cfg.spreadsheetId
    ? SpreadsheetApp.openById(cfg.spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();

  const qSheet = ss.getSheetByName(cfg.questionSheet);
  const rSheet = ss.getSheetByName(cfg.resultSheet);
  const resumeSheet = ss.getSheetByName(cfg.resumeSheet); // 🔥 SHEET BARU

  const qData = qSheet.getDataRange().getValues();

  // 🔥 mapping jawaban benar
  const answerKey = {};
  qData.slice(1).forEach(r=>{
    answerKey[r[0]] = r[5];
  });

  let correct = 0;
  let wrong = 0;

  // =========================
  // 🔥 1. SIMPAN DETAIL (AS-IS)
  // =========================
  data.answers.forEach(a=>{

    if(answerKey[a.q] === a.a){
      correct++;
    } else {
      wrong++;
    }

    rSheet.appendRow([
      new Date(),
      data.user,
      cfg.category,
      a.q,
      a.a,
      data.reason
    ]);
  });

  // =========================
  // 🔥 2. HITUNG SCORE
  // =========================
  const total = data.answers.length;

  const score = total > 0
    ? Math.round((correct / total) * 100)
    : 0;

  // =========================
  // 🔥 3. SIMPAN RESUME (1 BARIS)
  // =========================
  resumeSheet.appendRow([
    new Date(),
    data.user,
    cfg.category,
    correct,
    wrong,
    total,
    score,
    data.reason
  ]);

  return {
    status: "ok",
    correct,
    wrong,
    score
  };
}

function normalizeCase(str){
  return String(str || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .trim();
}

function getUserBestScore(user, exam){

  const SPREADSHEET_ID = "12ndmYZvFVNaZCMoYxCTe-jGpovo3dz5uUidvrSflCa0";
  const SHEET_NAME = "Resume";

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);

  if(!sh){
    return {
      maxScore: -1,
      status: "NO_DATA"
    };
  }

  const values = sh.getDataRange().getValues();
  const header = values.shift();

  const colUser  = header.indexOf("User");
  const colExam  = header.indexOf("Exam");
  const colScore = header.indexOf("Skor");

  const targetExam = normalizeCase(exam);

  let maxScore = -1;

  values.forEach(row => {

    const rowUser = row[colUser];
    const rowExam = normalizeCase(row[colExam]);

    if(rowUser == user && rowExam == targetExam){

      const score = Number(row[colScore]) || 0;

      if(score > maxScore){
        maxScore = score;
      }

    }

  });

  let status =
    maxScore == -1 ? "NO_DATA"
    : maxScore >= 80 ? "LULUS"
    : "BELUM";

  return {
    maxScore,
    status
  };
}


