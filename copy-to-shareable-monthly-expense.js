function copyNamedRangeWithFormatting() {
    var existingSpreadsheet = DriveApp.getFilesByName("Shareble Monthly Expense")
  
    // If 'New Spreadsheet' exists, use it; otherwise, create a new one
    var sharableMonthlyExpense
    if (existingSpreadsheet.hasNext()) {
      var existingSpreadsheet = existingSpreadsheet.next()
      sharableMonthlyExpense = SpreadsheetApp.openById(
        existingSpreadsheet.getId()
      )
    } else {
      sharableMonthlyExpense = SpreadsheetApp.create("Shareble Monthly Expense") // Change the name as desired
    }
  
    var monthlyExpense = SpreadsheetApp.getActiveSpreadsheet()
    var sourceSheet = monthlyExpense.getSheetByName("June 2024") // Source sheet with the named range
    var targetSheet = sharableMonthlyExpense.getActiveSheet() // Target sheet where you want to paste the data and formatting
    targetSheet.setName("June 2024")
  
    // Get the named range 'SharebleData' from the source sheet
    var namedRange = monthlyExpense.getRangeByName("SharebleDataJune2024")
  
    // Get the values and formatting from the named range
    var values = namedRange.getValues()
    var backgroundColors = namedRange.getBackgrounds()
    var fontColors = namedRange.getFontColors()
    var fontWeights = namedRange.getFontWeights()
    var fontStyles = namedRange.getFontStyles()
    var fontSizes = namedRange.getFontSizes()
    // var borders = namedRange.getBorders();
  
    // Calculate the target range dimensions
    var numRows = values.length
    var numCols = values[0].length
  
    // Define the top-left cell of the target range
    var targetRange = targetSheet.getRange(1, 1, numRows, numCols) // Change this to your desired target range start
  
    // Set the values and formatting to the target range
    targetRange.setValues(values)
    targetRange.setBackgrounds(backgroundColors)
    targetRange.setFontColors(fontColors)
    targetRange.setFontWeights(fontWeights)
    targetRange.setFontStyles(fontStyles)
    targetRange.setFontSizes(fontSizes)
    // targetRange.setBorder(borders.top, borders.left, borders.bottom, borders.right, borders.vertical, borders.horizontal);
  }
  