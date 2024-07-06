function copyNamedRangeWithFormatting() {
  var existingSpreadsheet = DriveApp.getFilesByName("Shareble Monthly Expense")

  var sharableMonthlyExpense
  if (existingSpreadsheet.hasNext()) {
    var existingSpreadsheet = existingSpreadsheet.next()
    sharableMonthlyExpense = SpreadsheetApp.openById(
      existingSpreadsheet.getId()
    )
  } else {
    sharableMonthlyExpense = SpreadsheetApp.create("Shareble Monthly Expense")
  }

  var monthlyExpense = SpreadsheetApp.getActiveSpreadsheet()

  var targetSheet = sharableMonthlyExpense.getSheetByName("July 2024")
  if (targetSheet == null)
    targetSheet = sharableMonthlyExpense.insertSheet("July 2024")
  else targetSheet.clear()

  var namedRange = monthlyExpense.getRangeByName("SharebleDataJuly2024")

  var values = namedRange.getValues()
  var backgroundColors = namedRange.getBackgrounds()
  var fontColors = namedRange.getFontColors()
  var fontWeights = namedRange.getFontWeights()
  var fontStyles = namedRange.getFontStyles()
  var fontSizes = namedRange.getFontSizes()

  var numRows = values.length
  var numCols = values[0].length

  var targetRange = targetSheet.getRange(1, 1, numRows, numCols)

  targetRange.setValues(values)
  targetRange.setBackgrounds(backgroundColors)
  targetRange.setFontColors(fontColors)
  targetRange.setFontWeights(fontWeights)
  targetRange.setFontStyles(fontStyles)
  targetRange.setFontSizes(fontSizes)
}
