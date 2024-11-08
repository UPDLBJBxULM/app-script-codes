function updateData(selectedValue) {
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TARGET_WORKSHEET_NAME");
  
  // Unhide rows B14 hingga B20 jika sebelumnya disembunyikan
  for (let i = 14; i <= 20; i++) {
    targetSheet.showRows(i);
  }
  
  // Buka spreadsheet sumber
  const sourceSpreadsheetId = "SOURCE_SPREADSHEET_ID";
  const sourceSheet = SpreadsheetApp.openById(sourceSpreadsheetId)
                                  .getSheetByName("SOURCE_WORKSHEET_NAME");
  
  // Ambil semua data
  const sourceData = sourceSheet.getDataRange().getValues();
  
  // Cari baris yang sesuai dengan pilihan dropdown (berdasarkan kolom A)
  const selectedRow = sourceData.find(row => row[0] === selectedValue);
  
  if (!selectedRow) {
    Logger.log("Data tidak ditemukan");
    return;
  }
  
  // Mapping data ke sel tujuan
  const dataMapping = [
    { sourceCol: 90, targetCell: "R2", addPrefix: true },
    { sourceCol: 3, targetCell: "S6" },  
    { sourceCol: 2, targetCell: "S9" }, 
    { sourceCol: 4, targetCell: "C6", addPrefix: true },  
    { sourceCol: 5, targetCell: "C7", addPrefix: true },  
    { sourceCol: 6, targetCell: "B14" }, 
    { sourceCol: 7, targetCell: "E14" }, 
    { sourceCol: 8, targetCell: "G14" }, 
    { sourceCol: 14, targetCell: "L14" },
    { sourceCol: 15, targetCell: "K14" },
    { sourceCol: 16, targetCell: "M14" },
    { sourceCol: 18, targetCell: "B15" },
    { sourceCol: 19, targetCell: "E15" },
    { sourceCol: 20, targetCell: "G15" },
    { sourceCol: 26, targetCell: "L15" },
    { sourceCol: 27, targetCell: "K15" },
    { sourceCol: 28, targetCell: "M15" },
    { sourceCol: 30, targetCell: "B16" },
    { sourceCol: 31, targetCell: "E16" },
    { sourceCol: 32, targetCell: "G16" },
    { sourceCol: 38, targetCell: "L16" },
    { sourceCol: 39, targetCell: "K16" },
    { sourceCol: 40, targetCell: "M16" },
    { sourceCol: 42, targetCell: "B17" },
    { sourceCol: 43, targetCell: "E17" },
    { sourceCol: 44, targetCell: "G17" },
    { sourceCol: 50, targetCell: "L17" },
    { sourceCol: 51, targetCell: "K17" },
    { sourceCol: 52, targetCell: "M17" },
    { sourceCol: 54, targetCell: "B18" },
    { sourceCol: 55, targetCell: "E18" },
    { sourceCol: 56, targetCell: "G18" },
    { sourceCol: 62, targetCell: "L18" },
    { sourceCol: 63, targetCell: "K18" },
    { sourceCol: 64, targetCell: "M18" },
    { sourceCol: 66, targetCell: "B19" },
    { sourceCol: 67, targetCell: "E19" },
    { sourceCol: 68, targetCell: "G19" },
    { sourceCol: 74, targetCell: "L19" },
    { sourceCol: 75, targetCell: "K19" },
    { sourceCol: 76, targetCell: "M19" },
    { sourceCol: 78, targetCell: "B20" },
    { sourceCol: 79, targetCell: "E20" },
    { sourceCol: 80, targetCell: "G20" }, 
    { sourceCol: 86, targetCell: "L20" },
    { sourceCol: 87, targetCell: "K20" },
    { sourceCol: 88, targetCell: "M20" },
  ];
  
  // Isi data ke sel tujuan
  dataMapping.forEach(mapping => {
    let value = selectedRow[mapping.sourceCol];
    if (mapping.addPrefix) {
      value = ": " + value;  // Menambahkan ": " di depan untuk C6 dan C7
    }
    targetSheet.getRange(mapping.targetCell).setValue(value);
  });

  // Panggil fungsi fillZonasi setelah updateData selesai
  fillZonasi();
  
  // Pengecekan kolom B14 hingga B20 untuk baris kosong
  for (let i = 14; i <= 20; i++) {
    const cellValue = targetSheet.getRange("B" + i).getValue();
    if (cellValue === "" || cellValue === null) {
      targetSheet.hideRows(i); // Sembunyikan baris yang kosong
    }
  }
}

// Fungsi untuk menjalankan update manual
function manualUpdate() {
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("WORKSHEET_NAME");
  const selectedValue = targetSheet.getRange("V1").getValue();
  if (selectedValue) {
    updateData(selectedValue);
  }
}

// Fungsi untuk membersihkan semua sel tujuan dan mewarnai O14 hingga O20
function clearTargetCells() {
  const targetSheet = SpreadsheetApp.getActive().getSheetByName("TARGET_WORKSHEET_NAME");
  const cellsToClear = [
    "R2", "S9", "S6", "C6", "C7", 
    "B14", "D14", "E14", "G14", "I14", "J14", "L14", "K14", "M14", "N14", "O14", "P14", "Q14", "R14", "S14",
    "B15", "D15", "E15", "G15", "I15", "J15", "L15", "K15", "M15", "N15", "O15", "P15", "Q15", "R15", "S15",
    "B16", "D16", "E16", "G16", "I16", "J16", "L16", "K16", "M16", "N16", "O16", "P16", "Q16", "R16", "S16",
    "B17", "D17", "E17", "G17", "I17", "J17", "L17", "K17", "M17", "N17", "O17", "P17", "Q17", "R17", "S17",
    "B18", "D18", "E18", "G18", "I18", "J18", "L18", "K18", "M18", "N18", "O18", "P18", "Q18", "R18", "S18",
    "B19", "D19", "E19", "G19", "I19", "J19", "L19", "K19", "M19", "N19", "O19", "P19", "Q19", "R19", "S19",
    "B20", "D20", "E20", "G20", "I20", "J20", "L20", "K20", "M20", "N20", "O20", "P20", "Q20", "R20", "S20", "V11"
  ];
  
  // Membersihkan isi sel
  cellsToClear.forEach(cell => {
    targetSheet.getRange(cell).clearContent();
  });
  
  // Mewarnai O14 hingga O20 dengan warna hex #ffffff
  targetSheet.getRange("O14:O20").setBackground("#ffffff");
}
