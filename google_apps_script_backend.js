/**
 * GOOGLE APPS SCRIPT - BACKEND API UNTUK PENCATATAN KEUANGAN PRIBADI
 * Cara Pakai:
 * 1. Buka Google Spreadsheet baru / yang sudah disiapkan.
 * 2. Klik menu Extensions (Ekstensi) -> Apps Script.
 * 3. Hapus semua kode bawaan, lalu paste seluruh kode ini.
 * 4. Klik tombol Save (ikon disket).
 * 5. Klik Deploy -> New deployment -> Pilih tipe "Web app".
 * 6. Execute as: "Me" | Who has access: "Anyone".
 * 7. Klik Deploy, izinkan akses (Authorize), lalu salin "Web app URL"-nya!
 */

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Sheet Transaksi
  let sheetTrans = ss.getSheetByName("Transaksi");
  if (!sheetTrans) {
    sheetTrans = ss.insertSheet("Transaksi");
    sheetTrans.appendRow(["ID", "Tanggal", "Tipe", "Kategori", "Nominal", "Catatan"]);
    sheetTrans.getRange("A1:F1").setFontWeight("bold").setBackground("#F3F4F6");
  }
  
  // 2. Sheet Utang_Piutang
  let sheetDebt = ss.getSheetByName("Utang_Piutang");
  if (!sheetDebt) {
    sheetDebt = ss.insertSheet("Utang_Piutang");
    sheetDebt.appendRow(["ID", "Tipe", "Nama", "Nominal", "JatuhTempo", "Status", "Catatan"]);
    sheetDebt.getRange("A1:G1").setFontWeight("bold").setBackground("#F3F4F6");
  }
  
  // 3. Sheet Tabungan_Target
  let sheetGoal = ss.getSheetByName("Tabungan_Target");
  if (!sheetGoal) {
    sheetGoal = ss.insertSheet("Tabungan_Target");
    sheetGoal.appendRow(["ID", "NamaTarget", "TargetNominal", "Terkumpul", "Emoji"]);
    sheetGoal.getRange("A1:E1").setFontWeight("bold").setBackground("#F3F4F6");
  }
}

// Mengambil seluruh data untuk ditampilkan di Aplikasi HP
function doGet(e) {
  try {
    setupSheets();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const transData = sheetToObjects(ss.getSheetByName("Transaksi"));
    const debtData = sheetToObjects(ss.getSheetByName("Utang_Piutang"));
    const goalData = sheetToObjects(ss.getSheetByName("Tabungan_Target"));
    
    const result = {
      status: "success",
      transaksi: transData,
      utangPiutang: debtData,
      tabunganTarget: goalData
    };
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Menerima data baru / update dari Aplikasi HP
function doPost(e) {
  try {
    setupSheets();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    
    if (action === "ADD_TRANSACTION") {
      const sheet = ss.getSheetByName("Transaksi");
      const d = payload.data;
      sheet.appendRow([d.id, d.tanggal, d.tipe, d.kategori, Number(d.nominal), d.catatan]);
    } 
    else if (action === "DELETE_TRANSACTION") {
      deleteRowById(ss.getSheetByName("Transaksi"), payload.id);
    }
    else if (action === "ADD_DEBT") {
      const sheet = ss.getSheetByName("Utang_Piutang");
      const d = payload.data;
      sheet.appendRow([d.id, d.tipe, d.nama, Number(d.nominal), d.jatuhTempo, "Belum Lunas", d.catatan]);
    }
    else if (action === "TOGGLE_DEBT_STATUS") {
      const sheet = ss.getSheetByName("Utang_Piutang");
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(payload.id)) {
          const currentStatus = data[i][5];
          const newStatus = currentStatus === "Lunas" ? "Belum Lunas" : "Lunas";
          sheet.getRange(i + 1, 6).setValue(newStatus);
          break;
        }
      }
    }
    else if (action === "ADD_GOAL") {
      const sheet = ss.getSheetByName("Tabungan_Target");
      const d = payload.data;
      sheet.appendRow([d.id, d.namaTarget, Number(d.targetNominal), Number(d.terkumpul || 0), d.emoji || "🎯"]);
    }
    else if (action === "TOPUP_GOAL") {
      const sheet = ss.getSheetByName("Tabungan_Target");
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(payload.id)) {
          const current = Number(data[i][3]) || 0;
          sheet.getRange(i + 1, 4).setValue(current + Number(payload.amount));
          break;
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Helper: Ubah baris spreadsheet menjadi array of objects
function sheetToObjects(sheet) {
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    let obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = rows[i][j];
    }
    result.push(obj);
  }
  return result;
}

// Helper: Hapus baris berdasarkan ID di kolom pertama
function deleteRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}