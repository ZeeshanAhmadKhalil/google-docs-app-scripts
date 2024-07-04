function createSheet() {
  // Get active spreadsheet and create a new sheet with the current month and year as the name
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var currentMonth = new Date()
  var sheetName = Utilities.formatDate(
    currentMonth,
    Session.getScriptTimeZone(),
    "MMMM yyyy"
  )
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName)

  // Clear any existing content in the sheet
  sheet.clear()

  // Developer info
  var developerName = "Ali (React Native)"
  var monthYear = Utilities.formatDate(
    currentMonth,
    Session.getScriptTimeZone(),
    "MMMM yyyy"
  )

  // Set the developer name and current month/year at the top, increasing height to take 3 rows
  var headerText = "Developer: " + developerName + "\n" + monthYear
  sheet.getRange("A1:E3").merge().setValue(headerText)
  sheet
    .getRange("A1:E3")
    .setVerticalAlignment("middle")
    .setFontSize(14)
    .setWrap(true)

  // Set the headers for the columns
  var headers = [
    "Date",
    "Day (weekday)",
    "Project",
    "Tasks",
    "Comments/Details of the task",
  ]
  var headerRange = sheet.getRange("A4:E4")
  headerRange.setValues([headers])

  // Set the background color for the header
  headerRange.setBackground("#167AA6")
  headerRange.setFontColor("#FFFFFF")

  // Set the date and day for each row in the current month (July 2024)
  var startDate = currentMonth
  var lastDay = new Date(
    startDate.getFullYear(),
    startDate.getMonth() + 1,
    0
  ).getDate()

  var data = []
  var rowsPerDate = 4 // Number of rows per date to accommodate multiple projects

  for (var day = 1; day <= lastDay; day++) {
    var date = new Date(startDate.getFullYear(), startDate.getMonth(), day)
    var dayOfWeek = date.toLocaleString("en-US", { weekday: "long" })

    for (var i = 0; i < rowsPerDate; i++) {
      var rowData =
        i === 0 ? [date, dayOfWeek, "", "", ""] : ["", "", "", "", ""]
      data.push(rowData)
    }
  }

  // Set the data in the sheet starting from row 5
  var dataRange = sheet.getRange(5, 1, data.length, data[0].length)
  dataRange.setValues(data)

  // Merge date and day cells for each set of rowsPerDate
  for (var i = 0; i < lastDay; i++) {
    var startRow = 5 + i * rowsPerDate
    sheet.getRange(startRow, 1, rowsPerDate, 1).mergeVertically()
    sheet.getRange(startRow, 2, rowsPerDate, 1).mergeVertically()
  }

  // Format the date column to display dates correctly and align to the left
  var dateRange = sheet.getRange(5, 1, data.length, 1)
  dateRange.setNumberFormat("MM/dd/yyyy")
  dateRange.setHorizontalAlignment("left")
  dateRange.setVerticalAlignment("middle")

  var dayRange = sheet.getRange(5, 2, data.length, 1)
  dayRange.setVerticalAlignment("middle")

  // Set the background color for date cells
  dateRange.setBackground("#DEEAF7")

  // Set the width of the project, tasks, and comments columns
  sheet.setColumnWidth(3, 100) // Project column width (half of tasks)
  sheet.setColumnWidth(4, 200) // Tasks column width
  sheet.setColumnWidth(5, 400) // Comments column width

  // Optionally, you can auto-resize the columns to fit the content
  sheet.autoResizeColumns(1, 2)
}
