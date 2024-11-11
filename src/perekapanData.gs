function rekapData() {
    try {
      // Mendapatkan spreadsheet aktif
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sourceSheet = ss.getSheetByName("SOURCE_WORKSHEET_NAME");
      const targetSheet = ss.getSheetByName("TARGET_WORKSHEET_NAME");
      
      // Array sel sumber
      const sourceCells = [
        "R2", "P10", "S9", "S6", "C6", "C7", 
        "B14", "D14", "E14", "G14", "I14", "J14", "L14", "K14", "M14", "N14", "O14", "P14", "Q14", "R14", "S14",
        "B15", "D15", "E15", "G15", "I15", "J15", "L15", "K15", "M15", "N15", "O15", "P15", "Q15", "R15", "S15",
        "B16", "D16", "E16", "G16", "I16", "J16", "L16", "K16", "M16", "N16", "O16", "P16", "Q16", "R16", "S16",
        "B17", "D17", "E17", "G17", "I17", "J17", "L17", "K17", "M17", "N17", "O17", "P17", "Q17", "R17", "S17",
        "B18", "D18", "E18", "G18", "I18", "J18", "L18", "K18", "M18", "N18", "O18", "P18", "Q18", "R18", "S18",
        "B19", "D19", "E19", "G19", "I19", "J19", "L19", "K19", "M19", "N19", "O19", "P19", "Q19", "R19", "S19",
        "B20", "D20", "E20", "G20", "I20", "J20", "L20", "K20", "M20", "N20", "O20", "P20", "Q20", "R20", "S20", "V11",
      ];
      
      // Mengambil semua nilai dari sel sumber dan melakukan penghapusan teks
      const values = sourceCells.map(cell => {
        let value = sourceSheet.getRange(cell).getValue();
        
        // Menghapus teks sesuai permintaan
        if (cell === "R2" || cell === "C6" || cell === "C7") {
          value = value.replace(": ", "");  // Menghapus ": "
        } else if (cell === "P10") {
          value = value.replace("Tanggal : ", "");  // Menghapus "Tanggai : "
        }
        
        return value;
      });
      
      // Menentukan baris pertama kosong pada worksheet GeneratePDF mulai dari baris ke-2
      const firstEmptyRow = targetSheet.getLastRow() + 1;
      const startRow = firstEmptyRow > 1 ? firstEmptyRow : 2;
      
      // Mengisi data ke worksheet GeneratePDF secara berurutan dari kolom A pada baris kosong pertama
      values.forEach((value, index) => {
        targetSheet.getRange(startRow, index + 1).setValue(value);
      });
      
      // Menampilkan pesan sukses
      SpreadsheetApp.getUi().alert('Data berhasil ditambahkan ke worksheet Rekap Data!');
      
    } catch (error) {
      // Menampilkan pesan error jika terjadi masalah
      SpreadsheetApp.getUi().alert('Terjadi error: ' + error.toString());
      Logger.log('Error in rekapData: ' + error.toString());
    }
  }
  