function telahDieditRABdanRealisasi(e) {
    // Daftar sheet yang akan didengarkan dan URL Cloud Function yang sesuai
    var sheetConfigs = {
      "WORKSHEET_NAME": "URL_CLOUD_FUNCTION",
      "WORKSHEET_NAME": "URL_CLOUD_FUNCTION"
    };
  
    try {
      // Dapatkan sheet aktif dan nama sheet
      var sheet = e.source.getActiveSheet();
      var activeSheetName = sheet.getName();
  
      // Validasi sheet yang aktif
      if (!sheetConfigs.hasOwnProperty(activeSheetName)) {
        Logger.log("Sheet tidak dalam daftar monitoring: " + activeSheetName);
        return;
      }
  
      // Validasi cell A1
      var cellA1 = sheet.getRange("A1").getValue();
      if (!cellA1 || isCellError(cellA1)) {
        Logger.log("Cell A1 kosong atau error pada sheet: " + activeSheetName);
        return;
      }
  
      // Ambil semua data dari sheet
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  
      // Buat payload dengan data lengkap
      var payload = {
        sheetName: activeSheetName,
        headers: headers,
        data: data.filter(row => row.some(cell => cell !== "")), // Filter baris kosong
        timestamp: new Date().toISOString(),
        spreadsheetId: e.source.getId(),
        spreadsheetUrl: e.source.getUrl()
      };
  
      Logger.log("Mengirim payload dari sheet: " + activeSheetName);
  
      // Konfigurasi request
      var options = {
        'method': 'post',
        'contentType': 'application/json',
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true,
        'headers': {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
        }
      };
  
      // Kirim request
      var response = UrlFetchApp.fetch(sheetConfigs[activeSheetName], options);
      var responseCode = response.getResponseCode();
      var responseText = response.getContentText();
  
      // Handle response
      if (responseCode >= 200 && responseCode < 300) {
        Logger.log("Berhasil mengirim data dari sheet: " + activeSheetName);
      } else {
        throw new Error(`HTTP Error ${responseCode}: ${responseText}`);
      }
  
    } catch (error) {
      Logger.log("Terjadi kesalahan: " + JSON.stringify({
        sheet: activeSheetName,
        error: error.toString(),
        detail: error.message
      }));
    }
  }
  
  function isCellError(value) {
    return value === null || 
           value === "" || 
           (typeof value === "object" && value instanceof Error);
  }
  
  
  