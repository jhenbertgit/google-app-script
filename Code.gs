const SHEET_ID = "your_google_sheet_id"

function getDataFromSheet(id) {
  var sheet = SpreadsheetApp.openById(id).getActiveSheet();
  var data = sheet.getDataRange().getValues();
  return data;
}

function convertDataToJson(data) {
  var jsonArray = [];

  if (data.length < 2) {
    return jsonArray;
  }

  var headers = data[0];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var record = {};

    for (var j = 0; j < row.length; j++) {
      record[headers[j]] = row[j];
    }

    jsonArray.push(record);
  }

  return jsonArray;
}

function doGet(e) {
  var sheetId = SHEET_ID; // Replace with your Google Sheet ID
  var data = getDataFromSheet(sheetId);
  var jsonArray = convertDataToJson(data);

  return ContentService.createTextOutput(JSON.stringify(jsonArray))
      .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return ContentService.createTextOutput("Error: No data found in the request.");
  }

  var sheetId = SHEET_ID; // Replace with your Google Sheet ID
  var data = JSON.parse(e.postData.contents);

  if (!Array.isArray(data) || data.length === 0) {
    return ContentService.createTextOutput("Error: Empty or invalid data in the request.");
  }

  var headers = Object.keys(data[0]);
  if (headers.length === 0) {
    return ContentService.createTextOutput("Error: No headers found in the data.");
  }

  writeToSheet(sheetId, data);

  return ContentService.createTextOutput("Data successfully written to Google Sheet.");
}

function writeToSheet(sheetId, data) {
  var sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
  var rows = [];
  
  // Check if the sheet is empty
  var lastRow = sheet.getLastRow();
  var headers = [];

  if (lastRow === 0) {
    // Extract headers from the first object in data
    if (data.length > 0) {
      headers = Object.keys(data[0]);
      sheet.appendRow(headers);
    }
  } else {
    // Get existing headers
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  // Write data to rows
  data.forEach(function(obj) {
    if (lastRow === 0) { // If sheet is empty, use keys as headers
      headers.forEach(function(header) {
        rows.push([obj[header] || '']);
      });
    } else { // If sheet is not empty, append values in the last row
      headers.forEach(function(header, index) {
        sheet.getRange(lastRow + 1, index + 1).setValue(obj[header] || '');
      });
    }
  });

  if (lastRow === 0) { // If sheet is empty, append all rows at once
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}