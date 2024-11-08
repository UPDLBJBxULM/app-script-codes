function generatePDFFinal() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    
    // Ambil kegiatan dari kolom S9
    var kegiatan = sheet.getRange("C7").getValue();
  
    // Hapus bagian ": " dari string dan ubah menjadi kapital semua
    kegiatan = kegiatan.replace(": ", "").toUpperCase();
  
    // Format timestamp menjadi ddMMyyyy
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "ddMMyyyy");
    
    // Tentukan nama file dengan format yang diminta
    var pdfName = sheet.getName() + "_" + "IBPPR" + "_" + kegiatan + "_" + timestamp;
    
    var targetFolderId = "TARGET_FOLDER_ID_GDRIVE";
    
    try {
      // Set print area to A1:S40
      sheet.getRange("A1:S40").activate();
      sheet.setActiveRange(sheet.getRange("A1:S40"));
      
      // URL untuk ekspor PDF dengan pengaturan kualitas tinggi
      var url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?" +
        "exportFormat=pdf" +
        "&format=pdf" +
        "&size=A4" +
        "&portrait=false" +
        "&fitw=true" +
        "&sheetnames=false" +
        "&pagenumbers=false" +
        "&gridlines=false" +
        "&printtitle=false" +
        "&scale=2" +                 // Memperbesar skala PDF
        "&fzr=true" +                // Menjaga format sel tanpa perubahan
        "&top_margin=0.75" +          // Margin untuk menjaga area cetak tetap luas
        "&bottom_margin=0.75" +
        "&left_margin=0.3" +
        "&right_margin=0.3" +
        "&range=A1:S40" +            // Tentukan area yang dicetak
        "&gid=" + sheet.getSheetId();
      
      // Ambil PDF dengan autentikasi dan pengaturan untuk kualitas maksimum
      var response = UrlFetchApp.fetch(url, {
        headers: {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        },
        muteHttpExceptions: true
      });
  
      if (response.getResponseCode() === 200) {
        var blob = response.getBlob().setName(pdfName + ".pdf");
        
        var targetFolder = DriveApp.getFolderById(targetFolderId);
        var pdfFile = targetFolder.createFile(blob);
        
        SpreadsheetApp.getUi().alert('PDF telah dibuat!\nFile disimpan sebagai: ' + pdfName + '.pdf\nLink: ' + pdfFile.getUrl());
        
        // Menyimpan URL file PDF ke dalam kolom V11
        sheet.getRange("V11").setValue(pdfFile.getUrl());
        
        // Memanggil fungsi rekapData() dari file rekapanData.gs
        rekapData();
        
      } else {
        throw new Error('Error dalam response fetch: ' + response.getContentText());
      }
  
    } catch (error) {
      SpreadsheetApp.getUi().alert('Error', 
        'Terjadi kesalahan saat membuat PDF: ' + error.toString(),
        SpreadsheetApp.getUi().ButtonSet.OK);
      Logger.log(error);
    }
  }
  
  function onOpen() {
    SpreadsheetApp.getUi().createMenu('PDF Tools')
      .addItem('Generate PDF', 'generatePDFFinal') // Ensure the correct function is called
      .addToUi();
  }
  