var DEFAULT_EVENT_ = 'friendsgiving-2026';
var DEFAULT_CATEGORIES_ = [
  { name: 'Sides / Mains', max: 6 },
  { name: 'Appetizers',    max: 4 },
  { name: 'Desserts',      max: 3 },
  { name: 'Drinks & Ice',  max: 4 }
];

// ---------- helpers ----------
function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Lowercase, trim, and collapse repeated whitespace for forgiving comparisons
function norm_(s) {
  return (s === null || s === undefined ? '' : String(s))
    .replace(/\s+/g, ' ').trim().toLowerCase();
}

function findSheet_(ss, name) {
  var exact = ss.getSheetByName(name);
  if (exact) return exact;
  var target = norm_(name);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (norm_(sheets[i].getName()) === target) return sheets[i];
  }
  return null;
}

function normalizeAttending_(v) {
  var a = norm_(v);
  if (a === 'yes') return 'Yes';
  if (a === 'no') return 'No';
  return 'Pending';
}

function getCategoryConfigs_(ss, tabName) {
  var configs = [];
  var configSheet = ss.getSheetByName('categories-config');
  if (configSheet && configSheet.getLastRow() > 1) {
    // Columns: Event Tab Name | Category Name | Max Slots
    var rows = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 3).getValues();
    rows.forEach(function (row) {
      var evt = norm_(row[0]);
      var cat = (row[1] || '').toString().trim();
      var max = parseInt(row[2], 10);
      if (evt === norm_(tabName) && cat) {
        configs.push({ name: cat, max: isNaN(max) ? 0 : max });
      }
    });
  }
  return configs.length ? configs : DEFAULT_CATEGORIES_.slice();
}

// Run this ONCE manually from the Apps Script editor to seed the config tab.
// (Seeding no longer happens inside doGet, so GET requests never write to the sheet.)
function setupConfigSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('categories-config')) return;
  var sh = ss.insertSheet('categories-config');
  sh.appendRow(['Event Tab Name', 'Category Name', 'Max Slots']);
  sh.getRange('A1:C1').setFontWeight('bold');
  ['friendsgiving-2026', 'friendsgiving-2025'].forEach(function (evt) {
    DEFAULT_CATEGORIES_.forEach(function (c) { sh.appendRow([evt, c.name, c.max]); });
  });
}

// ---------- 1. GET: guest list + category allotments ----------
function doGet(e) {
  var tabName = (e && e.parameter && e.parameter.tabName) ? e.parameter.tabName : DEFAULT_EVENT_;
  var configs = DEFAULT_CATEGORIES_;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    configs = getCategoryConfigs_(ss, tabName);

    var catMap = {};
    configs.forEach(function (c) {
      catMap[norm_(c.name)] = { name: c.name, max: c.max, dishes: [] };
    });

    var guests = [];
    var sheet = findSheet_(ss, tabName);

    if (sheet && sheet.getLastRow() > 1) {
      // A:F = Timestamp, Guest Name, Attending, Dish, Dish Category, Optional Note
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
      data.forEach(function (row) {
        var name = (row[1] || '').toString().trim();
        if (!name) return;

        var attending = normalizeAttending_(row[2]);
        var dish      = (row[3] || '').toString().trim();
        var category  = (row[4] || '').toString().trim();
        var notes     = (row[5] || '').toString().trim();

        guests.push({
          name: name,
          attending: attending,
          dish: dish,
          category: category,
          notes: notes
        });

        // Only confirmed (Yes) attendees with a configured category claim a slot
        var key = norm_(category);
        if (attending === 'Yes' && key && key !== 'undecided' && catMap[key]) {
          catMap[key].dishes.push({ guest: name, dish: dish || category });
        }
      });
    }

    var categories = configs.map(function (c) {
      var entry = catMap[norm_(c.name)];
      var claimed = entry.dishes.length;
      return {
        name: c.name,
        max: c.max,
        claimed: claimed,
        remaining: Math.max(0, c.max - claimed),
        dishes: entry.dishes
      };
    });

    return jsonOut_({ guests: guests, categories: categories });

  } catch (err) {
    // Always return the expected shape so the frontend never breaks
    return jsonOut_({
      guests: [],
      categories: configs.map(function (c) {
        return { name: c.name, max: c.max, claimed: 0, remaining: c.max, dishes: [] };
      }),
      error: String(err)
    });
  }
}

// ---------- 2. POST: upsert RSVP ----------
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var name = (data.name || '').toString().trim();
    if (!name) return jsonOut_({ result: 'error', message: 'Guest name is required.' });

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tabName = data.tabName || DEFAULT_EVENT_;
    var sheet = findSheet_(ss, tabName);
    if (!sheet) {
      sheet = ss.insertSheet(tabName);
      sheet.appendRow(['Timestamp', 'Guest Name', 'Attending', 'Dish', 'Dish Category', 'Optional Note']);
      sheet.getRange('A1:F1').setFontWeight('bold');
    }

    var attending = normalizeAttending_(data.attending);
    var notes = (data.notes || '').toString().trim();
    var dish = '';
    var category = '';

    // Dish and category are only kept for confirmed attendees; cleared for No / Pending
    if (attending === 'Yes') {
      dish = (data.dish || '').toString().trim();
      var cat = (data.category || '').toString().trim();
      category = norm_(cat) === 'undecided' ? '' : cat;
    }

    // Find existing row by Guest Name (column B)
    var existingRow = -1;
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var names = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (var i = 0; i < names.length; i++) {
        if (norm_(names[i][0]) === norm_(name)) { existingRow = i + 2; break; }
      }
    }

    var rowValues = [new Date(), name, attending, dish, category, notes];
    if (existingRow > 0) {
      // Keep the original spelling of the guest's name
      var existingName = sheet.getRange(existingRow, 2).getValue();
      rowValues[1] = existingName || name;
      sheet.getRange(existingRow, 1, 1, 6).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }

    return jsonOut_({ result: 'success' });

  } catch (err) {
    return jsonOut_({ result: 'error', message: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}