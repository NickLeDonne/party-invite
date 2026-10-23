// 1. Handle GET requests (Fetch guest list for web display)
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabName = (e && e.parameter && e.parameter.tabName) ? e.parameter.tabName : 'friendsgiving-2026';
  var sheet = ss.getSheetByName(tabName);
  
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ guests: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  var guests = [];

  // Skip header row (row 1)
  for (var i = 1; i < data.length; i++) {
    var name = data[i][1];
    var attending = data[i][2] || 'Pending';
    
    if (name) {
      guests.push({
        name: name,
        attending: attending
      });
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ guests: guests }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 2. Handle POST requests (Process RSVP form submissions)
function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = JSON.parse(e.postData.contents);
  var tabName = data.tabName || 'friendsgiving-2026';
  var sheet = ss.getSheetByName(tabName);

  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.appendRow(['Timestamp', 'Guest Name', 'Attending', 'Board Idea / Dish', 'Optional Note']);
    sheet.getRange("A1:E1").setFontWeight("bold");
  }

  var rows = sheet.getDataRange().getValues();
  var existingRowIndex = -1;

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][1].toString().trim().toLowerCase() === data.name.trim().toLowerCase()) {
      existingRowIndex = i + 1;
      break;
    }
  }

  if (existingRowIndex > 0) {
    sheet.getRange(existingRowIndex, 1).setValue(new Date());
    sheet.getRange(existingRowIndex, 3).setValue(data.attending);
    sheet.getRange(existingRowIndex, 4).setValue(data.boardIdea);
    sheet.getRange(existingRowIndex, 5).setValue(data.notes);
  } else {
    sheet.appendRow([
      new Date(),
      data.name,
      data.attending,
      data.boardIdea,
      data.notes
    ]);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ "result": "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}