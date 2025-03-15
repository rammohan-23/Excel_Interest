function doPost(e) {
  if (!e || !e.postData) {
    return ContentService.createTextOutput("No data received.");
  }

  var sheet = SpreadsheetApp.openById("18d6DlxXKCPAJr7oN0f-G02WpLw6iXJk5-LktuR7FBa8").getActiveSheet();
  var data = JSON.parse(e.postData.contents);

  var lastRow = sheet.getLastRow();
  var firstEmptyRow = lastRow + 1;
  
  for (var i = 2; i <= lastRow; i++) {
    if (!sheet.getRange(i, 1).getValue()) { // Check if Column A is empty
      firstEmptyRow = i;
      break;
    }
  }

  sheet.getRange(firstEmptyRow, 1).setValue(data.date);
  sheet.getRange(firstEmptyRow, 3).setValue(data.name);
  sheet.getRange(firstEmptyRow, 17).setValue(data.amount);
  sheet.getRange(firstEmptyRow, 18).setValue(data.interest);

  return ContentService.createTextOutput("Data added successfully!");
}

function doGet() {
  var sheet = SpreadsheetApp.openById("18d6DlxXKCPAJr7oN0f-G02WpLw6iXJk5-LktuR7FBa8").getActiveSheet();
  
  // Fetch B1 value & format it
  var rawDate = sheet.getRange("B1").getValue();
  var formattedHeader = "";
  
  if (rawDate instanceof Date) {
    formattedHeader = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd-MMM-yyyy");
  } else {
    formattedHeader = rawDate; // If it's not a date, return as-is
  }

  var data = sheet.getDataRange().getValues(); // Get all data
  var filteredData = [];

  for (var i = 2; i < data.length; i++) { // Skip header row
    if (data[i][0]) { // Check if Column A is not empty
      var formattedDate = Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "dd-MMM-yyyy");

      filteredData.push([
        formattedDate, // Formatted Date (Column A)
        data[i][2],
        data[i][20],   // Column T (Duration)
        data[i][16],   // Column P (Amount)
        data[i][17],   // Column Q (Interest)
        Math.round(data[i][18] * 100) / 100,  // Total Interest (Rounded)
        Math.round(data[i][19] * 100) / 100   // Amount + Interest (Rounded)
      ]);
    }
  }

  return ContentService.createTextOutput(
    JSON.stringify({ header: formattedHeader, rows: filteredData })
  ).setMimeType(ContentService.MimeType.JSON);
}
