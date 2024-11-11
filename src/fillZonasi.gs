function fillZonasi(e) {
    // Get the active sheet
    var sheet = SpreadsheetApp.getActiveSheet();
    
    // Define the range to check (M14:M20)
    var rangeToCheck = "M14:M20";
    var values = sheet.getRange(rangeToCheck).getValues();
    
    // Define value mappings
    const valueMap = {
      "Rendah": 1,
      "Moderat": 2,
      "Tinggi": 3,
      "Sangat Tinggi": 4,
      "Ekstrem": 5
    };
    
    // Define color mappings
    const colorMap = {
      "Rendah": "#92d14f",
      "Moderat": "#01b0f1",
      "Tinggi": "#ffff00",
      "Sangat Tinggi": "#ffc000",
      "Ekstrem": "#ff0000"
    };
    
    // Find the highest value in the range
    let highestRiskLevel = "";
    let highestValue = 0;
    
    // Find the highest risk level in the range
    values.forEach(row => {
      let cell = row[0];
      if (cell && valueMap[cell] > highestValue) {
        highestValue = valueMap[cell];
        highestRiskLevel = cell;
      }
    });
  
    // Iterate through each cell in M14:M20
    for (let i = 0; i < values.length; i++) {
      let cellValue = values[i][0];
      let targetCell = sheet.getRange(14 + i, 15); // Corresponding O cell for M row
      
      if (cellValue) {
        // Set value and color if M cell is not empty
        targetCell.setValue(highestRiskLevel);
        targetCell.setBackground(colorMap[highestRiskLevel]);
      } else {
        // Clear value and color if M cell is empty
        targetCell.setValue("");
        targetCell.setBackground(null);
      }
    }
  }
  