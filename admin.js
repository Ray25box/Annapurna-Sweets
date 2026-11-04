// =============================================
// Annapurna Sweets — Admin Panel
// Google Sheets Edition
// =============================================

// --- Password: SHA-256 hash of default password ---
// Plain-text password is NEVER stored here.
const DEFAULT_PASSWORD_HASH = '6714ed55bd7f9dbb38016e24e54b5f142bf1e6d73a9eab9032ae0822969d171f';
const PASSWORD_STORAGE      = 'annapurna_admin_password_hash';

// --- In-memory product list (populated from Google Sheets after login) ---
var adminProducts     = [];
var editingProductIdx = null;   // null = adding new, number = editing existing

// =============================================
// PASSWORD SYSTEM
// =============================================

async function sha256(message) {
    const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(message));
    return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2,'0')).join('');
}

function getPasswordHash() {
    return localStorage.getItem(PASSWORD_STORAGE) || DEFAULT_PASSWORD_HASH;
}

async function checkPassword() {
    const entered     = document.getElementById('password-input').value;
    const enteredHash = await sha256(entered);
    if (enteredHash === getPasswordHash()) {
        sessionStorage.setItem('annapurna_admin_logged_in', 'true');
        document.getElementById('login-screen').style.display = 'none';
        document.getElementById('admin-panel').style.display  = 'block';
        initAdminPanel();
    } else {
        document.getElementById('login-error').style.display = 'block';
    }
}

document.addEventListener('DOMContentLoaded', function () {
    if (sessionStorage.getItem('annapurna_admin_logged_in') === 'true') {
        document.getElementById('login-screen').style.display = 'none';
        document.getElementById('admin-panel').style.display  = 'block';
        initAdminPanel();
    }
    const pw = document.getElementById('password-input');
    if (pw) pw.addEventListener('keypress', function (e) { if (e.key === 'Enter') checkPassword(); });
});

function logout() {
    sessionStorage.removeItem('annapurna_admin_logged_in');
    document.getElementById('password-input').value     = '';
    document.getElementById('admin-panel').style.display  = 'none';
    document.getElementById('login-screen').style.display = 'flex';
}

function togglePw(inputId, btn) {
    const input = document.getElementById(inputId);
    const hide  = input.type === 'password';
    input.type  = hide ? 'text' : 'password';
    btn.innerHTML = hide ? '&#128064;' : '&#128065;';
    btn.title     = hide ? 'Hide password' : 'Show password';
}

async function changePassword() {
    const current   = document.getElementById('cp-current').value;
    const newPw     = document.getElementById('cp-new').value;
    const confirmPw = document.getElementById('cp-confirm').value;
    const msgEl     = document.getElementById('cp-message');

    if (await sha256(current) !== getPasswordHash()) {
        return showCpMsg(msgEl, '❌ Current password is incorrect.', 'error');
    }
    if (newPw.length < 6) {
        return showCpMsg(msgEl, '⚠️ New password must be at least 6 characters.', 'warn');
    }
    if (newPw !== confirmPw) {
        return showCpMsg(msgEl, '⚠️ Passwords do not match.', 'warn');
    }

    localStorage.setItem(PASSWORD_STORAGE, await sha256(newPw));
    ['cp-current','cp-new','cp-confirm'].forEach(id => { document.getElementById(id).value = ''; });
    showCpMsg(msgEl, '✅ Password changed successfully!', 'ok');
    setTimeout(() => msgEl.style.display = 'none', 5000);
}

function showCpMsg(el, text, type) {
    el.textContent = text;
    el.style.cssText = type === 'ok'
        ? 'display:block;background:#f0fdf4;color:#16a34a;border-color:#bbf7d0;'
        : type === 'warn'
        ? 'display:block;background:#fffbeb;color:#d97706;border-color:#fde68a;'
        : 'display:block;background:#fef2f2;color:#dc2626;border-color:#fecaca;';
}

// =============================================
// INIT — load everything after login
// =============================================

async function initAdminPanel() {
    showMessage('⏳ Loading data from Google Sheets…', 'info');
    await refreshAllData();
    await loadContactInfo();
}

async function refreshAllData() {
    try {
        const rows = await fetchSheet('Products');
        adminProducts = rows.map(function (r) {
            return {
                name:        (r.Name        || '').trim(),
                category:    (r.Category    || '').trim(),
                pricePerPc:  (r.PricePerPc  || '').trim(),
                pricePerKg:  (r.PricePerKg  || '').trim(),
                description: (r.Description || '').trim(),
                imageURL:    (r.ImageURL    || '').trim(),
                image:        r.ImageURL ? driveImg(r.ImageURL) : null,
                isFeatured:   isFeatured(r.Featured)
            };
        });
        loadProductTable();
        loadMenuList();
        loadCategories();
        clearMessage();
    } catch (e) {
        showMessage('❌ Could not load data. Check SHEET_ID in sheets.js — ' + e.message, 'error');
    }
}

// =============================================
// EXCEL UPLOAD  (bulk import)
// =============================================

function handleExcelUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async function (e) {
        try {
            const data     = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet    = workbook.Sheets[workbook.SheetNames[0]];
            const rows     = XLSX.utils.sheet_to_json(sheet);

            if (rows.length === 0) { showMessage('⚠️ The Excel file is empty.', 'error'); return; }

            showMessage('⏳ Importing ' + rows.length + ' products to Google Sheets…', 'info');

            const products = rows.map(function (row) {
                // Preserve any image URLs already in the sheet for matching products
                const existing = adminProducts.find(function (p) {
                    return p.name.toLowerCase() === (row.Name || '').toLowerCase();
                });
                return {
                    Name:        row.Name         || 'Unnamed',
                    Category:    row.Category     || 'Sweets',
                    PricePerPc:  row.Price_Per_Pc  ? String(row.Price_Per_Pc)  : '',
                    PricePerKg:  row.Price_Per_Kg  ? String(row.Price_Per_Kg)  : '',
                    Description: row.Description  || '',
                    ImageURL:    existing          ? existing.imageURL : ''
                };
            });

            const result = await sheetWrite({ action: 'bulkImport', products: products });

            if (result.status === 'ok') {
                showMessage('✅ ' + products.length + ' products imported successfully!');
                await refreshAllData();
            } else {
                showMessage('❌ Import failed: ' + (result.message || 'Unknown error'), 'error');
            }
        } catch (err) {
            showMessage('❌ Could not read the file. Make sure it is a valid .xlsx or .csv.', 'error');
            console.error(err);
        }
    };
    reader.readAsArrayBuffer(file);
}

// =============================================
// EDIT MENU — Add / Update / Delete items
// =============================================

async function addOrUpdateProduct() {
    const name        = document.getElementById('new-name').value.trim();
    const category    = document.getElementById('new-category').value;
    const pricePerPc  = document.getElementById('new-price-pc').value.trim();
    const pricePerKg  = document.getElementById('new-price-kg').value.trim();
    const description = document.getElementById('new-description').value.trim();

    if (!name) { showMessage('⚠️ Item Name is required.', 'error'); return; }

    showMessage('⏳ Saving…', 'info');

    var result;
    if (editingProductIdx !== null) {
        const p = adminProducts[editingProductIdx];
        result = await sheetWrite({
            action:       'updateProduct',
            originalName: p.name,
            Name:         name,
            Category:     category,
            PricePerPc:   pricePerPc,
            PricePerKg:   pricePerKg,
            Description:  description
            // ImageURL and Featured are NOT sent — Apps Script preserves them from current row
        });
    } else {
        // Check for duplicate locally (fast, avoids a round-trip)
        if (adminProducts.find(function (p) { return p.name.toLowerCase() === name.toLowerCase(); })) {
            showMessage('⚠️ "' + name + '" already exists in the menu.', 'error');
            return;
        }
        result = await sheetWrite({
            action:      'addProduct',
            Name:        name,
            Category:    category,
            PricePerPc:  pricePerPc,
            PricePerKg:  pricePerKg,
            Description: description,
            ImageURL:    '',
            Featured:    'FALSE'
        });
    }

    if (result.status === 'ok') {
        showMessage(editingProductIdx !== null ? '✅ "' + name + '" updated!' : '✅ "' + name + '" added!');
        cancelEdit();
        await refreshAllData();
    } else {
        showMessage('❌ Save failed: ' + (result.message || 'Check SCRIPT_URL in sheets.js'), 'error');
    }
}

function editProduct(idx) {
    const p = adminProducts[idx];
    if (!p) return;
    editingProductIdx = idx;

    document.getElementById('new-name').value        = p.name;
    document.getElementById('new-category').value    = p.category;
    document.getElementById('new-price-pc').value    = p.pricePerPc;
    document.getElementById('new-price-kg').value    = p.pricePerKg;
    document.getElementById('new-description').value = p.description;

    const submitBtn = document.getElementById('submit-btn');
    submitBtn.innerHTML        = '&#10004; Update Item';
    submitBtn.style.background = '#16a34a';
    submitBtn.style.color      = 'white';
    document.getElementById('cancel-edit-btn').style.display = 'inline-block';

    document.getElementById('new-name').scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function cancelEdit() {
    editingProductIdx = null;
    ['new-name','new-price-pc','new-price-kg','new-description'].forEach(function (id) {
        document.getElementById(id).value = '';
    });
    const submitBtn = document.getElementById('submit-btn');
    submitBtn.innerHTML        = '&#43; Add Item to Menu';
    submitBtn.style.background = '';
    submitBtn.style.color      = '';
    document.getElementById('cancel-edit-btn').style.display = 'none';
}

async function deleteProduct(idx) {
    const p = adminProducts[idx];
    if (!p) return;
    if (!confirm('Remove "' + p.name + '" from the menu?')) return;

    showMessage('⏳ Deleting…', 'info');
    const result = await sheetWrite({ action: 'deleteProduct', Name: p.name });

    if (result.status === 'ok') {
        showMessage('🗑️ "' + p.name + '" removed.');
        await refreshAllData();
    } else {
        showMessage('❌ Delete failed: ' + (result.message || ''), 'error');
    }
}

// =============================================
// PRODUCT TABLE (Photo & Feature management)
// =============================================

function loadProductTable() {
    const tbody = document.getElementById('product-table-body');
    tbody.innerHTML = '';

    if (adminProducts.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="empty-row">No products yet. Upload an Excel sheet or add items using the Edit Menu above.</td></tr>';
        return;
    }

    adminProducts.forEach(function (p, idx) {
        const tr = document.createElement('tr');

        const photoCell = p.image
            ? '<img src="' + p.image + '" style="width:52px;height:52px;object-fit:cover;border-radius:8px;border:1px solid #ddd;" onerror="this.style.display=\'none\'">'
            : '<span style="color:#bbb;font-size:0.78rem;font-style:italic;">No image</span>';

        const removeBtn = p.image
            ? '<button class="remove-img-btn" data-idx="' + idx + '">&#128465; Remove</button>'
            : '';

        tr.innerHTML =
            '<td style="font-weight:700;color:#888;">' + (idx + 1) + '</td>' +
            '<td style="font-weight:600;">' + esc(p.name) + '</td>' +
            '<td><span style="background:#fef9ef;border:1px solid #fde68a;color:#92400e;padding:3px 10px;border-radius:20px;font-size:0.72rem;font-weight:700;">' + esc(p.category) + '</span></td>' +
            '<td>' + photoCell + '</td>' +
            '<td>' +
                '<label style="display:inline-block;background:#f0fdf4;border:1.5px dashed #86efac;padding:6px 12px;border-radius:8px;cursor:pointer;font-size:0.75rem;color:#16a34a;font-weight:600;">' +
                    '<input type="file" accept="image/*" data-idx="' + idx + '" class="img-upload-input" style="display:none;">&#128247; Upload' +
                '</label> ' + removeBtn +
            '</td>' +
            '<td style="text-align:center;">' +
                '<input type="checkbox" class="feature-toggle" data-idx="' + idx + '" ' + (p.isFeatured ? 'checked' : '') + (p.image ? '' : ' disabled title="Upload a photo first"') + ' style="width:18px;height:18px;cursor:pointer;">' +
            '</td>';

        tbody.appendChild(tr);
    });

    // Attach events via delegation (safe — no inline dynamic strings in event handlers)
    tbody.addEventListener('change', handleTableChange);
    tbody.addEventListener('click', handleTableClick);
}

function handleTableChange(e) {
    const target = e.target;
    const idx    = parseInt(target.dataset.idx, 10);
    if (isNaN(idx)) return;

    if (target.classList.contains('img-upload-input')) {
        handleImageUpload(target.files[0], idx);
    } else if (target.classList.contains('feature-toggle')) {
        toggleFeature(idx, target);
    }
}

function handleTableClick(e) {
    const btn = e.target.closest('.remove-img-btn');
    if (!btn) return;
    const idx = parseInt(btn.dataset.idx, 10);
    if (!isNaN(idx)) removeImage(idx);
}

// =============================================
// MENU LIST (Edit menu section)
// =============================================

function loadMenuList() {
    const container = document.getElementById('menu-list');
    if (!container) return;

    if (adminProducts.length === 0) {
        container.innerHTML = '<p style="padding:16px;font-style:italic;color:#aaa;">No items yet. Add one using the form above.</p>';
        return;
    }

    container.innerHTML = '';
    adminProducts.forEach(function (p, idx) {
        const row = document.createElement('div');
        row.className = 'menu-item-row';
        row.innerHTML =
            '<div class="menu-item-info">' +
                '<span class="menu-item-name">' + esc(p.name) + '</span>' +
                '<span class="menu-item-cat">'  + esc(p.category) + '</span>' +
                (p.pricePerPc  ? '<span class="menu-item-price">&#8377;' + esc(p.pricePerPc)  + '/pc</span>' : '') +
                (p.pricePerKg  ? '<span class="menu-item-price">&#8377;' + esc(p.pricePerKg)  + '/kg</span>' : '') +
                (p.description ? '<span class="menu-item-desc">'         + esc(p.description)  + '</span>'  : '') +
            '</div>' +
            '<div style="display:flex;gap:8px;">' +
                '<button class="edit-item-btn"   data-idx="' + idx + '" style="background:#eab308;color:#fff;border:none;padding:6px 10px;border-radius:4px;cursor:pointer;font-size:0.75rem;">✏️ Edit</button>' +
                '<button class="delete-item-btn" data-idx="' + idx + '">&#128465; Remove</button>' +
            '</div>';
        container.appendChild(row);
    });

    // Event delegation
    container.addEventListener('click', function (e) {
        const btn = e.target.closest('[data-idx]');
        if (!btn) return;
        const idx = parseInt(btn.dataset.idx, 10);
        if (btn.classList.contains('edit-item-btn'))   editProduct(idx);
        if (btn.classList.contains('delete-item-btn')) deleteProduct(idx);
    });
}

// =============================================
// IMAGE UPLOAD — uploads to Google Drive via Apps Script
// =============================================

async function handleImageUpload(file, idx) {
    if (!file) return;
    const p = adminProducts[idx];
    if (!p) return;

    showMessage('⏳ Compressing & uploading image to Google Drive…', 'info');

    // Compress image client-side before upload (target ≤ 600px wide, JPEG 80%)
    const compressed = await compressImage(file);
    const base64     = compressed.split(',')[1];

    // Step 1: upload the file to Drive
    const uploadResult = await sheetWrite({
        action:   'uploadImage',
        base64:   base64,
        mimeType: 'image/jpeg',
        filename: (p.name.replace(/[^a-z0-9]/gi, '_') || 'product') + '.jpg'
    });

    if (uploadResult.status !== 'ok') {
        showMessage('❌ Upload failed: ' + (uploadResult.message || ''), 'error');
        return;
    }

    // Step 2: save the Drive URL to the product row in the Sheet
    const updateResult = await sheetWrite({
        action:    'updateProduct',
        originalName: p.name,
        Name:      p.name,
        ImageURL:  uploadResult.url
    });

    if (updateResult.status === 'ok') {
        showMessage('✅ Photo uploaded for "' + p.name + '"!');
        await refreshAllData();
    } else {
        showMessage('⚠️ Photo saved to Drive but sheet update failed.', 'error');
    }
}

/** Compress an image using an offscreen canvas. Returns a JPEG data URL. */
function compressImage(file) {
    return new Promise(function (resolve) {
        const reader = new FileReader();
        reader.onload = function (e) {
            const img = new Image();
            img.onload = function () {
                const MAX = 800;
                let w = img.width, h = img.height;
                if (w > MAX || h > MAX) {
                    if (w > h) { h = Math.round(h * MAX / w); w = MAX; }
                    else       { w = Math.round(w * MAX / h); h = MAX; }
                }
                const canvas = document.createElement('canvas');
                canvas.width  = w;
                canvas.height = h;
                canvas.getContext('2d').drawImage(img, 0, 0, w, h);
                resolve(canvas.toDataURL('image/jpeg', 0.82));
            };
            img.src = e.target.result;
        };
        reader.readAsDataURL(file);
    });
}

async function removeImage(idx) {
    const p = adminProducts[idx];
    if (!p) return;

    const result = await sheetWrite({
        action:       'updateProduct',
        originalName: p.name,
        Name:         p.name,
        ImageURL:     '',
        Featured:     'FALSE'    // also un-feature when image is removed
    });

    if (result.status === 'ok') {
        showMessage('🗑️ Image removed from "' + p.name + '".');
        await refreshAllData();
    } else {
        showMessage('❌ Error removing image: ' + (result.message || ''), 'error');
    }
}

// =============================================
// FEATURE TOGGLE
// =============================================

async function toggleFeature(idx, checkbox) {
    const p = adminProducts[idx];
    if (!p) return;

    if (checkbox.checked) {
        const count = adminProducts.filter(function (x) { return x.isFeatured; }).length;
        if (count >= 12) {
            checkbox.checked = false;
            showMessage('⚠️ Maximum 12 featured items allowed on the homepage.', 'error');
            return;
        }
    }

    const result = await sheetWrite({
        action:       'updateProduct',
        originalName: p.name,
        Name:         p.name,
        Featured:     checkbox.checked ? 'TRUE' : 'FALSE'
    });

    if (result.status === 'ok') {
        p.isFeatured = checkbox.checked;
        showMessage(checkbox.checked ? '⭐ Added to Homepage!' : 'Removed from Homepage.');
    } else {
        checkbox.checked = !checkbox.checked;  // revert
        showMessage('❌ Could not update: ' + (result.message || ''), 'error');
    }
}

// =============================================
// CATEGORIES  (derived from Products sheet)
// =============================================

function loadCategories() {
    const listEl = document.getElementById('category-list');
    if (listEl) {
        const cats = getUsedCategories();
        if (cats.length === 0) {
            listEl.innerHTML = '<p style="color:#aaa;font-style:italic;font-size:0.85rem;">Categories are created automatically when you add products with that category name.</p>';
        } else {
            listEl.innerHTML = cats.map(function (c) {
                return '<div style="background:#fef3c7;border:1px solid #fde68a;color:#92400e;padding:5px 14px;border-radius:20px;font-size:0.8rem;font-weight:bold;">' + esc(c) + '</div>';
            }).join('');
        }
    }

    // Also refresh the category dropdown in the add-item form
    const select = document.getElementById('new-category');
    if (select) {
        const defaults = ['Sweets', 'Snack', 'Grocery'];
        const all      = Array.from(new Set(defaults.concat(getUsedCategories())));
        const current  = select.value;
        select.innerHTML = all.map(function (c) {
            return '<option value="' + esc(c) + '">' + esc(c) + '</option>';
        }).join('');
        if (all.indexOf(current) !== -1) select.value = current;
    }
}

function getUsedCategories() {
    return Array.from(new Set(adminProducts.map(function (p) { return p.category; }).filter(Boolean)));
}

function addCategory() {
    const input = document.getElementById('new-category-name');
    const name  = (input.value || '').trim();
    if (!name) return;

    const select = document.getElementById('new-category');
    if (select && !Array.from(select.options).find(function (o) { return o.value === name; })) {
        const opt = document.createElement('option');
        opt.value = opt.textContent = name;
        select.appendChild(opt);
        select.value = name;
    }
    input.value = '';
    showMessage('"' + name + '" added to dropdown. Use it when adding a product to save it permanently.');
}

// deleteCategory is not needed with the Sheets approach; just leave the button wired
function deleteCategory(name) {
    showMessage('To delete a category, delete all products in that category from the menu.', 'info');
}

// =============================================
// CONTACT INFO
// =============================================

async function loadContactInfo() {
    try {
        const rows = await fetchSheet('Config');
        const get  = function (key) { return (rows.find(function (r) { return r.Key === key; }) || {}).Value || ''; };
        const gen  = document.getElementById('contact-general');
        const blk  = document.getElementById('contact-bulk');
        if (gen) gen.value = get('general_wa');
        if (blk) blk.value = get('bulk_wa');
    } catch (e) {
        console.warn('Could not load contact info:', e);
    }
}

async function saveContactInfo() {
    const general = document.getElementById('contact-general').value.trim();
    const bulk    = document.getElementById('contact-bulk').value.trim();

    showMessage('⏳ Saving contact info…', 'info');
    const result = await sheetWrite({ action: 'saveConfig', general_wa: general, bulk_wa: bulk });

    if (result.status === 'ok') {
        showMessage('📞 Contact info saved successfully!');
    } else {
        showMessage('❌ Save failed: ' + (result.message || 'Check SCRIPT_URL in sheets.js'), 'error');
    }
}

// =============================================
// CLEAR ALL DATA
// =============================================

async function clearAllData() {
    if (!confirm('Are you sure? This permanently deletes ALL products from the Google Sheet.')) return;
    showMessage('⏳ Clearing all data…', 'info');
    const result = await sheetWrite({ action: 'clearAll' });
    if (result.status === 'ok') {
        adminProducts = [];
        loadProductTable();
        loadMenuList();
        showMessage('🗑️ All product data cleared from the sheet.');
    } else {
        showMessage('❌ Error: ' + (result.message || ''), 'error');
    }
}

// =============================================
// STATUS MESSAGE
// =============================================

function showMessage(msg, type) {
    const el = document.getElementById('status-message');
    if (!el) return;
    el.textContent    = msg;
    el.style.display  = 'block';
    el.style.background  = type === 'error' ? '#fef2f2' : type === 'info' ? '#eff6ff' : '#f0fdf4';
    el.style.color       = type === 'error' ? '#dc2626' : type === 'info' ? '#1d4ed8' : '#16a34a';
    el.style.borderColor = type === 'error' ? '#fecaca' : type === 'info' ? '#bfdbfe' : '#bbf7d0';
    if (type !== 'info') setTimeout(function () { el.style.display = 'none'; }, 5000);
}

function clearMessage() {
    const el = document.getElementById('status-message');
    if (el) el.style.display = 'none';
}

// =============================================
// UTILITY
// =============================================

/** HTML-escape a string for safe insertion into innerHTML. */
function esc(str) {
    return String(str || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}
