// 1. Handle GET requests (Fetch guest list and category allotments for web display)
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabName = (e && e.parameter && e.parameter.tabName) ? e.parameter.tabName : 'friendsgiving-2026';
  
  // 1a. Retrieve or initialize category configurations from 'categories-config' tab
  var configSheet = ss.getSheetByName('categories-config');
  var categoryConfigs = [];

  if (!configSheet) {
    try {
      configSheet = ss.insertSheet('categories-config');
      configSheet.appendRow(['Event Tab Name', 'Category Name', 'Max Slots']);
      configSheet.getRange("A1:C1").setFontWeight("bold");
      var defaultConfigs = [
        ['friendsgiving-2026', 'Savory Sides / Mains', 6],
        ['friendsgiving-2026', 'Appetizers', 4],
        ['friendsgiving-2026', 'Desserts', 3],
        ['friendsgiving-2026', 'Drinks & Ice', 4],
        ['friendsgiving-2025', 'Savory Sides / Mains', 6],
        ['friendsgiving-2025', 'Appetizers', 4],
        ['friendsgiving-2025', 'Desserts', 3],
        ['friendsgiving-2025', 'Drinks & Ice', 4]
      ];
      for (var d = 0; d < defaultConfigs.length; d++) {
        configSheet.appendRow(defaultConfigs[d]);
      }
    } catch (err) {
      // Proceed gracefully with in-memory defaults if creation is constrained
    }
  }

  if (configSheet) {
    var configData = configSheet.getDataRange().getValues();
    for (var r = 1; r < configData.length; r++) {
      var rowEventTab = (configData[r][0] || '').toString().trim();
      var rowCatName = (configData[r][1] || '').toString().trim();
      var rowMax = parseInt(configData[r][2], 10);

      if (rowEventTab.toLowerCase() === tabName.toLowerCase() && rowCatName) {
        categoryConfigs.push({
          name: rowCatName,
          max: isNaN(rowMax) ? 0 : rowMax
        });
      }
    }
  }

  // Fallback defaults if no categories matched the event tab
  if (categoryConfigs.length === 0) {
    categoryConfigs = [
      { name: 'Savory Sides / Mains', max: 6 },
      { name: 'Appetizers', max: 4 },
      { name: 'Desserts', max: 3 },
      { name: 'Drinks & Ice', max: 4 }
    ];
  }

  // 1b. Retrieve event tab data
  var sheet = ss.getSheetByName(tabName);
  var categoriesMap = {};
  categoryConfigs.forEach(function(c) {
    categoriesMap[c.name.toLowerCase()] = {
      name: c.name,
      max: c.max,
      dishes: []
    };
  });

  function findCategoryName(rawString) {
    if (!rawString) return null;
    var lower = rawString.trim().toLowerCase();
    for (var key in categoriesMap) {
      if (key === lower) return categoriesMap[key].name;
    }
    return null;
  }

  if (!sheet) {
    var emptyCategories = categoryConfigs.map(function(c) {
      return {
        name: c.name,
        max: c.max,
        claimed: 0,
        remaining: c.max,
        dishes: []
      };
    });
    return ContentService
      .createTextOutput(JSON.stringify({ guests: [], categories: emptyCategories }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  var guests = [];

  // Skip header row (row 1)
  for (var i = 1; i < data.length; i++) {
    var name = (data[i][1] || '').toString().trim();
    var attending = (data[i][2] || 'Pending').toString().trim();
    var rawColD = (data[i][3] || '').toString().trim();
    var notes = (data[i][4] || '').toString().trim();

    if (name) {
      var parsedCat = '';
      var parsedDish = '';

      if (rawColD) {
        var colonIndex = rawColD.indexOf(':');
        if (colonIndex > -1) {
          var potentialCat = rawColD.substring(0, colonIndex).trim();
          var matchedCat = findCategoryName(potentialCat);
          if (matchedCat) {
            parsedCat = matchedCat;
            parsedDish = rawColD.substring(colonIndex + 1).trim();
          } else {
            parsedDish = rawColD;
          }
        } else {
          var matchedCatOnly = findCategoryName(rawColD);
          if (matchedCatOnly) {
            parsedCat = matchedCatOnly;
            parsedDish = '';
          } else {
            parsedDish = rawColD;
          }
        }
      }

      var guestObj = {
        name: name,
        attending: attending,
        category: parsedCat,
        dish: parsedDish,
        notes: notes
      };
      guests.push(guestObj);

      // Only count claimed slots for confirmed attendees with a category
      if (attending.toLowerCase() === 'yes' && parsedCat) {
        var catKey = parsedCat.toLowerCase();
        if (categoriesMap[catKey]) {
          categoriesMap[catKey].dishes.push({
            guest: name,
            dish: parsedDish || parsedCat
          });
        }
      }
    }
  }

  var categoriesResult = categoryConfigs.map(function(c) {
    var catKey = c.name.toLowerCase();
    var catData = categoriesMap[catKey] || { dishes: [] };
    var claimed = catData.dishes.length;
    var remaining = Math.max(0, c.max - claimed);
    return {
      name: c.name,
      max: c.max,
      claimed: claimed,
      remaining: remaining,
      dishes: catData.dishes
    };
  });

  return ContentService
    .createTextOutput(JSON.stringify({ guests: guests, categories: categoriesResult }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 2. Handle POST requests (Process RSVP form submissions & updates)
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

  // Format Column D & Column E values based on attendance
  var dishVal = "";
  var notesVal = "";

  if (data.attending && data.attending.toString().trim().toLowerCase() === 'no') {
    // If declining, release previously held category slot and clear notes
    dishVal = "";
    notesVal = "";
  } else {
    if (data.category && data.dish) {
      dishVal = data.category.trim() + ": " + data.dish.trim();
    } else if (data.category) {
      dishVal = data.category.trim();
    } else if (data.dish) {
      dishVal = data.dish.trim();
    } else if (data.boardIdea) {
      dishVal = data.boardIdea.trim();
    }
    notesVal = (data.notes || "").toString().trim();
  }

  if (existingRowIndex > 0) {
    sheet.getRange(existingRowIndex, 1).setValue(new Date());
    sheet.getRange(existingRowIndex, 3).setValue(data.attending);
    sheet.getRange(existingRowIndex, 4).setValue(dishVal);
    sheet.getRange(existingRowIndex, 5).setValue(notesVal);
  } else {
    sheet.appendRow([
      new Date(),
      data.name,
      data.attending,
      dishVal,
      notesVal
    ]);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ "result": "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}
