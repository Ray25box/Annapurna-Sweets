/**
 * ============================================================
 * Annapurna Sweets — Google Apps Script Backend
 * ============================================================
 * HOW TO DEPLOY:
 *   1. Open your Google Sheet.
 *   2. Click Extensions → Apps Script.
 *   3. Delete any existing code, paste this entire file.
 *   4. Click Deploy → New Deployment.
 *   5. Type: Web App
 *   6. Execute as: Me
 *   7. Who has access: Anyone
 *   8. Click Deploy → Copy the Web App URL.
 *   9. Paste that URL as SCRIPT_URL in sheets.js.
 * ============================================================
 */

// ---- Entry points ----

function doGet(e) {
  return ContentService
    .createTextOutput('Annapurna Sweets API is running ✓')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var result = dispatch(ss, data);
    return json(result);
  } catch (err) {
    return json({ status: 'error', message: err.toString() });
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- Router ----

function dispatch(ss, data) {
  switch (data.action) {
    case 'addProduct':            return addProduct(ss, data);
    case 'updateProduct':         return updateProduct(ss, data);
    case 'deleteProduct':         return deleteProduct(ss, data);
    case 'uploadImage':           return uploadImage(ss, data);
    case 'saveConfig':            return saveConfig(ss, data);
    case 'bulkImport':            return bulkImport(ss, data);
    case 'clearAll':              return clearAll(ss);
    default:
      return { status: 'error', message: 'Unknown action: ' + data.action };
  }
}

// ---- Sheet helpers ----

function getSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    // Auto-create sheet with headers on first use
    sheet = ss.insertSheet(name);
    if (name === 'Products') {
      sheet.getRange(1, 1, 1, 7)
        .setValues([['Name','Category','PricePerPc','PricePerKg','Description','ImageURL','Featured']]);
    } else if (name === 'Config') {
      sheet.getRange(1, 1, 1, 2).setValues([['Key','Value']]);
    }
  }
  return sheet;
}

/** Returns the 1-based row index of a product by name, or -1 if not found. */
function findRow(sheet, name) {
  var values = sheet.getDataRange().getValues();
  name = String(name || '').trim().toLowerCase();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === name) return i + 1;
  }
  return -1;
}

// ---- Product actions ----

function addProduct(ss, d) {
  var sheet = getSheet(ss, 'Products');
  sheet.appendRow([
    d.Name        || '',
    d.Category    || 'Sweets',
    d.PricePerPc  || '',
    d.PricePerKg  || '',
    d.Description || '',
    d.ImageURL    || '',
    d.Featured    || 'FALSE'
  ]);
  return { status: 'ok' };
}

function updateProduct(ss, d) {
  var sheet = getSheet(ss, 'Products');
  var row   = findRow(sheet, d.originalName || d.Name);
  if (row === -1) return { status: 'error', message: 'Product not found: ' + d.Name };

  // Read current row to preserve any fields not sent in the request
  var current = sheet.getRange(row, 1, 1, 7).getValues()[0];

  sheet.getRange(row, 1, 1, 7).setValues([[
    d.Name        !== undefined ? d.Name        : current[0],
    d.Category    !== undefined ? d.Category    : current[1],
    d.PricePerPc  !== undefined ? d.PricePerPc  : current[2],
    d.PricePerKg  !== undefined ? d.PricePerKg  : current[3],
    d.Description !== undefined ? d.Description : current[4],
    d.ImageURL    !== undefined ? d.ImageURL    : current[5],
    d.Featured    !== undefined ? d.Featured    : current[6]
  ]]);
  return { status: 'ok' };
}

function deleteProduct(ss, d) {
  var sheet = getSheet(ss, 'Products');
  var row   = findRow(sheet, d.Name);
  if (row === -1) return { status: 'error', message: 'Product not found' };
  sheet.deleteRow(row);
  return { status: 'ok' };
}

// ---- Image upload ----

function uploadImage(ss, d) {
  try {
    var decoded  = Utilities.base64Decode(d.base64);
    var blob     = Utilities.newBlob(decoded, d.mimeType || 'image/jpeg', d.filename || 'product.jpg');

    // Find or create an "Annapurna Sweets Images" folder in Drive
    var folderName = 'Annapurna Sweets Images';
    var iter = DriveApp.getFoldersByName(folderName);
    var folder = iter.hasNext() ? iter.next() : DriveApp.createFolder(folderName);

    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var url = 'https://drive.google.com/uc?export=view&id=' + file.getId();
    return { status: 'ok', url: url };
  } catch (err) {
    return { status: 'error', message: 'Image upload failed: ' + err.toString() };
  }
}

// ---- Config (WhatsApp numbers) ----

function saveConfig(ss, d) {
  var sheet  = getSheet(ss, 'Config');
  var values = sheet.getDataRange().getValues();

  var updates = {};
  if (d.general_wa !== undefined) updates['general_wa'] = d.general_wa;
  if (d.bulk_wa    !== undefined) updates['bulk_wa']    = d.bulk_wa;

  for (var key in updates) {
    var found = false;
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim() === key) {
        sheet.getRange(i + 1, 2).setValue(updates[key]);
        found = true;
        break;
      }
    }
    if (!found) sheet.appendRow([key, updates[key]]);
  }
  return { status: 'ok' };
}

// ---- Bulk import (from Excel upload) ----

function bulkImport(ss, d) {
  var sheet = getSheet(ss, 'Products');

  // Remember existing image URLs and featured flags so they aren't lost
  var memo = {};
  var curr = sheet.getDataRange().getValues();
  for (var i = 1; i < curr.length; i++) {
    var nm = String(curr[i][0]).trim();
    if (nm) memo[nm.toLowerCase()] = { url: curr[i][5], featured: curr[i][6] };
  }

  // Clear product rows (keep header)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).clearContent();
  }

  // Write new rows
  var products = d.products || [];
  products.forEach(function (p) {
    var prev = memo[(p.Name || '').trim().toLowerCase()] || {};
    sheet.appendRow([
      p.Name        || '',
      p.Category    || 'Sweets',
      p.PricePerPc  || '',
      p.PricePerKg  || '',
      p.Description || '',
      p.ImageURL    || prev.url || '',
      prev.featured || 'FALSE'
    ]);
  });

  return { status: 'ok', count: products.length };
}

// ---- Clear all products ----

function clearAll(ss) {
  var sheet = getSheet(ss, 'Products');
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).clearContent();
  }
  return { status: 'ok' };
}
