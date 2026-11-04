/**
 * ============================================================
 * Annapurna Sweets — Secure Google Apps Script Backend
 * ============================================================
 * Only the Store Owner can make changes using the Admin Password.
 * Requests without the correct Admin Password will be rejected.
 *
 * HOW TO DEPLOY:
 *   1. Open your Google Sheet.
 *   2. Click Extensions → Apps Script.
 *   3. Replace all code with this file.
 *   4. (Optional) Customize DEFAULT_ADMIN_PASSWORD below.
 *   5. Click Deploy → New Deployment (or Manage Deployments → Edit → New Version).
 *   6. Type: Web App
 *   7. Execute as: Me
 *   8. Who has access: Anyone
 *   9. Click Deploy → Copy the Web App URL.
 *   10. Paste that URL into SCRIPT_URL in admin.js.
 * ============================================================
 */

// 🔒 Default Admin Password (you can also change this anytime from admin.html)
var DEFAULT_ADMIN_PASSWORD = 'AnnapurnaSweets@2024';

function getAdminPassword() {
  var stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  return stored || DEFAULT_ADMIN_PASSWORD;
}

// ---- Entry points ----

function doGet(e) {
  return ContentService
    .createTextOutput('Annapurna Sweets API is running (Protected).')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents || '{}');

    // 🔒 Security Check: Every write or admin action requires the valid Admin Password token.
    var currentPassword = getAdminPassword();
    if (!data.token || String(data.token).trim() !== currentPassword) {
      return json({ status: 'error', message: 'Unauthorized: Invalid or missing admin password.' });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    case 'authCheck':             return { status: 'ok', message: 'Authenticated' };
    case 'changePassword':        return changePassword(data);
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

// ---- Password Change Action ----

function changePassword(d) {
  if (!d.newPassword || String(d.newPassword).trim().length < 6) {
    return { status: 'error', message: 'New password must be at least 6 characters.' };
  }
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', String(d.newPassword).trim());
  return { status: 'ok', message: 'Password changed successfully.' };
}

// ---- Sheet helpers ----

function getSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
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

// ---- Image upload to Google Drive ----

function uploadImage(ss, d) {
  try {
    var decoded = Utilities.base64Decode(d.base64);
    var blob    = Utilities.newBlob(decoded, d.mimeType || 'image/jpeg', d.filename || 'product.jpg');

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

  var memo = {};
  var curr = sheet.getDataRange().getValues();
  for (var i = 1; i < curr.length; i++) {
    var nm = String(curr[i][0]).trim();
    if (nm) memo[nm.toLowerCase()] = { url: curr[i][5], featured: curr[i][6] };
  }

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).clearContent();
  }

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
