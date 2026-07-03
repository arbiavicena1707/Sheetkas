/**
 * API Google Apps Script (GAS) untuk integrasi Google Sheets.
 * Kolom Sheet: [No, Tanggal Kejadian, Tanggal Input, Deskripsi, Metode, Income, Expense, Balance]
 */

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const rows = sheet.getDataRange().getValues();
    const timeZone = Session.getScriptTimeZone();
    
    if (rows.length <= 1) {
      return makeJsonResponse({
        success: true,
        summary: { totalIncome: 0, totalExpense: 0, balance: 0, balanceQris: 0, balanceCash: 0 },
        transactions: []
      });
    }
    
    const dataRows = rows.slice(1);
    
    let totalIncome = 0;
    let totalExpense = 0;
    let balanceQris = 0;
    let balanceCash = 0;
    
    const transactions = dataRows.map((row, index) => {
      const formattedTanggalKejadian = parseDateSafely(row[1], timeZone);
      const formattedTanggalInput = parseDateSafely(row[2], timeZone);
      
      const income = Number(row[5]) || 0;
      const expense = Number(row[6]) || 0;
      const balance = Number(row[7]) || 0;
      const metode = String(row[4] || "Cash").trim();
      
      totalIncome += income;
      totalExpense += expense;
      
      if (metode.toUpperCase() === "QRIS") {
        balanceQris += (income - expense);
      } else {
        balanceCash += (income - expense);
      }
      
      return {
        id: row[0] || index + 1,
        tanggalKejadian: formattedTanggalKejadian,
        tanggalInput: formattedTanggalInput,
        description: String(row[3]),
        metode: metode,
        income: income,
        expense: expense,
        balance: balance
      };
    });
    
    const currentBalance = totalIncome - totalExpense;
    
    return makeJsonResponse({
      success: true,
      summary: {
        totalIncome: totalIncome,
        totalExpense: totalExpense,
        balance: currentBalance,
        balanceQris: balanceQris,
        balanceCash: balanceCash
      },
      transactions: transactions.reverse() // Transaksi terbaru di atas
    });
    
  } catch (error) {
    return makeJsonResponse({ success: false, error: error.toString() });
  }
}

function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) {
      throw new Error("Payload kosong.");
    }
    
    const payload = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    const dateKejadianStr = payload.tanggalKejadian; // Format "YYYY-MM-DD"
    const description = payload.description || "";
    const metode = payload.metode || "Cash"; // "QRIS" atau "Cash"
    const income = Number(payload.income) || 0;
    const expense = Number(payload.expense) || 0;
    
    // Auto-increment ID
    let newId = 1;
    if (lastRow > 1) {
      const lastIdVal = sheet.getRange(lastRow, 1).getValue();
      newId = (Number(lastIdVal) || 0) + 1;
    }
    
    // Hitung running balance
    let previousBalance = 0;
    if (lastRow > 1) {
      previousBalance = Number(sheet.getRange(lastRow, 8).getValue()) || 0;
    }
    const newBalance = previousBalance + income - expense;
    
    // Parse Tanggal Kejadian
    let dateKejadianObj;
    if (dateKejadianStr) {
      const dateParts = dateKejadianStr.split("-");
      dateKejadianObj = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);
    } else {
      dateKejadianObj = new Date();
    }
    
    // Tanggal Input (Sekarang)
    const dateInputObj = new Date();
    
    const nextRow = lastRow + 1;
    sheet.appendRow([
      newId,
      dateKejadianObj,
      dateInputObj,
      description,
      metode,
      income,
      expense,
      newBalance
    ]);
    
    // Formatting
    sheet.getRange(nextRow, 2).setNumberFormat("yyyy-mm-dd");
    sheet.getRange(nextRow, 3).setNumberFormat("yyyy-mm-dd");
    sheet.getRange(nextRow, 6).setNumberFormat('"Rp"#,##0');
    sheet.getRange(nextRow, 7).setNumberFormat('"Rp"#,##0');
    sheet.getRange(nextRow, 8).setNumberFormat('"Rp"#,##0');
    
    return makeJsonResponse({
      success: true,
      transaction: {
        id: newId,
        tanggalKejadian: Utilities.formatDate(dateKejadianObj, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        tanggalInput: Utilities.formatDate(dateInputObj, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        description: description,
        metode: metode,
        income: income,
        expense: expense,
        balance: newBalance
      }
    });
    
  } catch (error) {
    return makeJsonResponse({ success: false, error: error.toString() });
  }
}

/**
 * Parsing tanggal secara aman untuk menghindari format invalid.
 */
function parseDateSafely(dateVal, timeZone) {
  if (!dateVal) return "";
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, timeZone, "yyyy-MM-dd");
  }
  const d = new Date(dateVal);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, timeZone, "yyyy-MM-dd");
  }
  return String(dateVal);
}

function makeJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
