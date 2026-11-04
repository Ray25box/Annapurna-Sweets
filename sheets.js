// =============================================
// Annapurna Sweets — Google Sheets Data Layer
// =============================================
//
// 🔧 STEP 1 — Paste your Google Sheet ID below.
//    It's the long string in your Sheet URL:
//    https://docs.google.com/spreadsheets/d/  ← THIS PART → /edit
//
//    Sheet must be shared as "Anyone with the link → Viewer"
//    (no "Publish to web" needed).

const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE';

// =============================================
// 🔧 STEP 2 — After deploying the Apps Script,
//    paste the Web App URL here.
// =============================================
const SCRIPT_URL = 'YOUR_APPS_SCRIPT_URL_HERE';

// =============================================
// Write / mutate data via the Apps Script
// (used only by admin.js — not public pages)
// =============================================
async function sheetWrite(payload) {
    if (!SCRIPT_URL || SCRIPT_URL === 'YOUR_APPS_SCRIPT_URL_HERE') {
        console.warn('Annapurna: Please set SCRIPT_URL in sheets.js');
        return { status: 'error', message: 'Script URL not configured. See setup guide.' };
    }
    try {
        // ⚠️ No Content-Type header here — this keeps it a "simple request"
        //    so the browser skips the CORS preflight which Apps Script doesn't support.
        const res = await fetch(SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error('HTTP ' + res.status);
        return await res.json();
    } catch (e) {
        console.error('Annapurna: sheetWrite error —', e.message);
        return { status: 'error', message: e.message };
    }
}

// =============================================
// Internal helper — fetch a named sheet tab
// and return an array of row objects.
// =============================================
async function fetchSheet(sheetName) {
    if (SHEET_ID === 'YOUR_GOOGLE_SHEET_ID_HERE') {
        console.warn('Annapurna: Please set your SHEET_ID in sheets.js');
        return [];
    }
    const url =
        'https://docs.google.com/spreadsheets/d/' + SHEET_ID +
        '/gviz/tq?tqx=out:json&sheet=' + encodeURIComponent(sheetName);
    try {
        const res = await fetch(url);
        if (!res.ok) throw new Error('HTTP ' + res.status);
        const text = await res.text();
        // Strip Google's JS wrapper and parse plain JSON
        const json = JSON.parse(text.substring(text.indexOf('{'), text.lastIndexOf('}') + 1));
        if (!json.table || !json.table.rows) return [];
        const cols = json.table.cols.map(function (c) {
            return (c.label || c.id || '').trim();
        });
        return json.table.rows
            .filter(function (row) {
                // Skip completely empty rows
                return row.c && row.c.some(function (c) {
                    return c && c.v !== null && String(c.v).trim() !== '';
                });
            })
            .map(function (row) {
                var obj = {};
                (row.c || []).forEach(function (cell, i) {
                    if (cols[i]) {
                        obj[cols[i]] = (cell && cell.v !== null) ? String(cell.v).trim() : '';
                    }
                });
                return obj;
            });
    } catch (e) {
        console.error('Annapurna: Error loading "' + sheetName + '" sheet —', e.message);
        return [];
    }
}

// =============================================
// Convert a Google Drive sharing URL into a
// direct image URL usable in <img src="…">.
//
// Supports:
//   https://drive.google.com/file/d/FILE_ID/view…
//   https://drive.google.com/open?id=FILE_ID
//   https://drive.google.com/uc?id=FILE_ID  (already direct)
// =============================================
function driveImg(url) {
    if (!url) return '';
    url = url.trim();
    // Already a direct uc link — just ensure export=view is set
    if (url.includes('drive.google.com/uc')) {
        return url.includes('export=') ? url : url + '&export=view';
    }
    // Extract file ID from /d/FILE_ID or ?id=FILE_ID patterns
    var m = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (m) return 'https://drive.google.com/uc?export=view&id=' + m[1];
    // Not a Drive URL — return as-is (supports external image URLs too)
    return url;
}

// =============================================
// Check if the "Featured" column value means
// "yes" (accepts: TRUE / Yes / 1 / true / yes)
// =============================================
function isFeatured(val) {
    return ['true', 'yes', '1'].indexOf(String(val).toLowerCase().trim()) !== -1;
}
