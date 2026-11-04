// =============================================
// Annapurna Sweets — Public Customer Data Layer (Read-Only)
// =============================================
//
// 🔧 Paste your Google Sheet ID below.
//    Found in your Google Sheet URL:
//    https://docs.google.com/spreadsheets/d/ [THIS-IS-YOUR-SHEET-ID] /edit
//
//    Ensure the sheet is shared: "Anyone with the link -> Viewer"
//    (Notice: This file is strictly READ-ONLY. It contains no write functions
//    and no admin endpoints, so customers browsing the site cannot make any changes.)

const SHEET_ID = '1adKwx5lHXEzK-6BuTHU1BTPy3bsiyYd9dMqs5VoBJMo';

// =============================================
// Fetch a named sheet tab (read-only gviz API)
// =============================================
async function fetchSheet(sheetName) {
    if (!SHEET_ID || SHEET_ID === 'YOUR_GOOGLE_SHEET_ID_HERE') {
        console.warn('Annapurna: Please configure SHEET_ID in sheets.js');
        return [];
    }
    const url =
        'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(SHEET_ID) +
        '/gviz/tq?tqx=out:json&sheet=' + encodeURIComponent(sheetName);
    try {
        const res = await fetch(url);
        if (!res.ok) throw new Error('HTTP ' + res.status);
        const text = await res.text();
        const json = JSON.parse(text.substring(text.indexOf('{'), text.lastIndexOf('}') + 1));
        if (!json.table || !json.table.rows) return [];
        const cols = json.table.cols.map(function (c) {
            return (c.label || c.id || '').trim();
        });
        return json.table.rows
            .filter(function (row) {
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
        console.error('Annapurna: Error reading "' + sheetName + '" sheet —', e.message);
        return [];
    }
}

// =============================================
// Direct Image URL Converter (Modern Google CDN)
// =============================================
function driveImg(url) {
    if (!url) return '';
    url = url.trim();
    // Extract file ID from any Google Drive sharing or direct link
    var m = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || 
            url.match(/[?&]id=([a-zA-Z0-9_-]+)/) ||
            url.match(/googleusercontent\.com\/d\/([a-zA-Z0-9_-]+)/);
    if (m) {
        // Modern Google UserContent CDN (avoids deprecated /uc blocks)
        return 'https://lh3.googleusercontent.com/d/' + m[1];
    }
    return url;
}

// Fallback to Google Drive high-resolution thumbnail if CDN is delayed
function driveImgFallback(img) {
    if (img.dataset.hasFallback) return;
    img.dataset.hasFallback = 'true';
    var m = img.src.match(/\/d\/([a-zA-Z0-9_-]+)/) || img.src.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (m) {
        img.src = 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w1000';
    }
}

// =============================================
// Featured Check
// =============================================
function isFeatured(val) {
    return ['true', 'yes', '1'].indexOf(String(val).toLowerCase().trim()) !== -1;
}
