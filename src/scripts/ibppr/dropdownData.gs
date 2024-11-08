function updateDropdownList() {
    // Spreadsheet IDs
    const sourceSpreadsheetId = "SOURCE_SPREADSHEET_ID";
    const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("WORKSHEET_NAME");
    
    // Get source data (specifically from column A)
    const sourceSheet = SpreadsheetApp.openById(sourceSpreadsheetId)
                                    .getSheetByName("WORKSHEET_NAME");
    const sourceData = sourceSheet.getRange("A2:A") // Mulai dari A2 jika A1 adalah header
                                 .getValues()
                                 .filter(row => row[0] !== ""); // Remove empty cells
    
    // Create data validation rule for cell V1
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(sourceData.map(row => row[0]))
      .setAllowInvalid(false)
      .build();
    
    // Apply validation rule to cell V1
    targetSheet.getRange("V1").setDataValidation(rule);
  }