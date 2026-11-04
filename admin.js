// =============================================
// Annapurna Sweets — Secure Admin Panel Controller
// (Only loaded on admin.html — NEVER loaded on customer pages)
// =============================================

// 🔧 The Google Apps Script Web App URL for executing admin write operations:
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbx8PBszOt1QbBG-nvPQLNHlA3VEQc3GM0L4DhEuhXLKUKIKTw3mGyPDqaWFAV24HCTsaA/exec';

var adminProducts     = [];
var editingProductIdx = null;

// =============================================
// AUTHENTICATION & SECURE WRITE LAYER
// =============================================

function getAdminToken() {
    return sessionStorage.getItem('annapurna_admin_token') || '';
}

/**
 * Executes a write operation against Google Apps Script.
 * Automatically injects the active Admin Password token for authorization.
 */
async function sheetWrite(payload) {
    if (!SCRIPT_URL || SCRIPT_URL === 'YOUR_APPS_SCRIPT_URL_HERE') {
        return { status: 'error', message: 'Script URL not configured in admin.js' };
    }
    const token = payload.token || getAdminToken();
    if (!token && payload.action !== 'authCheck') {
        return { status: 'error', message: 'Unauthorized. Please login first.' };
    }
    try {
        const bodyData = Object.assign({}, payload, { token: token });
        const res = await fetch(SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify(bodyData)
        });
        if (!res.ok) throw new Error('HTTP ' + res.status);
        return await res.json();
    } catch (e) {
        console.error('Annapurna Admin Write Error:', e);
        return { status: 'error', message: e.message || 'Connection failed' };
    }
}

// --- Check Password & Log In ---
async function checkPassword() {
    const entered = (document.getElementById('password-input').value || '').trim();
    const errorEl = document.getElementById('login-error');
    errorEl.style.display = 'none';

    if (!entered) {
        errorEl.textContent = '⚠️ Please enter the admin password.';
        errorEl.style.display = 'block';
        return;
    }

    const loginBtn = document.querySelector('.login-btn');
    if (loginBtn) { loginBtn.textContent = 'Verifying…'; loginBtn.disabled = true; }

    const res = await sheetWrite({ action: 'authCheck', token: entered });

    if (loginBtn) { loginBtn.textContent = 'Login'; loginBtn.disabled = false; }

    if (res.status === 'ok') {
        sessionStorage.setItem('annapurna_admin_token', entered);
        document.getElementById('login-screen').style.display = 'none';
        document.getElementById('admin-panel').style.display  = 'block';
        initAdminPanel();
    } else {
        errorEl.textContent = '❌ ' + (res.message || 'Incorrect password.');
        errorEl.style.display = 'block';
    }
}

// --- Auto-check Session on Page Load ---
document.addEventListener('DOMContentLoaded', async function () {
    const token = getAdminToken();
    if (token) {
        document.getElementById('login-screen').style.display = 'none';
        document.getElementById('admin-panel').style.display  = 'block';
        initAdminPanel();
    }
    const pw = document.getElementById('password-input');
    if (pw) {
        pw.addEventListener('keypress', function (e) {
            if (e.key === 'Enter') checkPassword();
        });
    }
});

function logout() {
    sessionStorage.removeItem('annapurna_admin_token');
    document.getElementById('password-input').value       = '';
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

// --- Change Password from Dashboard ---
async function changePassword() {
    const current   = document.getElementById('cp-current').value.trim();
    const newPw     = document.getElementById('cp-new').value.trim();
    const confirmPw = document.getElementById('cp-confirm').value.trim();
    const msgEl     = document.getElementById('cp-message');

    if (current !== getAdminToken()) {
        return showCpMsg(msgEl, '❌ Current password does not match active session.', 'error');
    }
    if (newPw.length < 6) {
        return showCpMsg(msgEl, '⚠️ New password must be at least 6 characters.', 'warn');
    }
    if (newPw !== confirmPw) {
        return showCpMsg(msgEl, '⚠️ Passwords do not match.', 'warn');
    }

    msgEl.textContent = '⏳ Updating password on server…';
    msgEl.style.display = 'block';

    const res = await sheetWrite({
        action: 'changePassword',
        token: current,
        newPassword: newPw
    });

    if (res.status === 'ok') {
        sessionStorage.setItem('annapurna_admin_token', newPw);
        ['cp-current','cp-new','cp-confirm'].forEach(id => { document.getElementById(id).value = ''; });
        showCpMsg(msgEl, '✅ Password updated successfully!', 'ok');
        setTimeout(() => { msgEl.style.display = 'none'; }, 5000);
    } else {
        showCpMsg(msgEl, '❌ ' + (res.message || 'Failed to update password.'), 'error');
    }
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
// INIT & DATA LOADING
// =============================================

async function initAdminPanel() {
    showMessage('⏳ Connecting to Google Sheets…', 'info');
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
        showMessage('❌ Could not load products. Please check SHEET_ID in sheets.js', 'error');
    }
}

// =============================================
// EXCEL / CSV BULK IMPORT
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
            showMessage('❌ Could not read file. Ensure it is a valid .xlsx or .csv', 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

// =============================================
// ADD / EDIT / DELETE PRODUCTS
// =============================================

async function addOrUpdateProduct() {
    const name        = document.getElementById('new-name').value.trim();
    const category    = document.getElementById('new-category').value;
    const pricePerPc  = document.getElementById('new-price-pc').value.trim();
    const pricePerKg  = document.getElementById('new-price-kg').value.trim();
    const description = document.getElementById('new-description').value.trim();

    if (!name) { showMessage('⚠️ Item Name is required.', 'error'); return; }

    showMessage('⏳ Saving item to Google Sheet…', 'info');

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
        });
    } else {
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
        showMessage(editingProductIdx !== null ? '✅ "' + name + '" updated!' : '✅ "' + name + '" added to menu!');
        cancelEdit();
        await refreshAllData();
    } else {
        showMessage('❌ Save failed: ' + (result.message || 'Authorization failed'), 'error');
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

    showMessage('⏳ Removing from Google Sheet…', 'info');
    const result = await sheetWrite({ action: 'deleteProduct', Name: p.name });

    if (result.status === 'ok') {
        showMessage('🗑️ "' + p.name + '" removed from menu.');
        await refreshAllData();
    } else {
        showMessage('❌ Delete failed: ' + (result.message || ''), 'error');
    }
}

// =============================================
// PRODUCT TABLE & DRIVE IMAGE UPLOAD
// =============================================

function loadProductTable() {
    const tbody = document.getElementById('product-table-body');
    tbody.innerHTML = '';

    if (adminProducts.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="empty-row">No products yet. Add items above or import an Excel file.</td></tr>';
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

    tbody.removeEventListener('change', handleTableChange);
    tbody.removeEventListener('click', handleTableClick);
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

function loadMenuList() {
    const container = document.getElementById('menu-list');
    if (!container) return;

    if (adminProducts.length === 0) {
        container.innerHTML = '<p style="padding:16px;font-style:italic;color:#aaa;">No items yet.</p>';
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

    container.onclick = function (e) {
        const btn = e.target.closest('[data-idx]');
        if (!btn) return;
        const idx = parseInt(btn.dataset.idx, 10);
        if (btn.classList.contains('edit-item-btn'))   editProduct(idx);
        if (btn.classList.contains('delete-item-btn')) deleteProduct(idx);
    };
}

async function handleImageUpload(file, idx) {
    if (!file) return;
    const p = adminProducts[idx];
    if (!p) return;

    showMessage('⏳ Compressing and uploading photo to Google Drive…', 'info');

    const compressed = await compressImage(file);
    const base64     = compressed.split(',')[1];

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

    const updateResult = await sheetWrite({
        action:       'updateProduct',
        originalName: p.name,
        Name:         p.name,
        ImageURL:     uploadResult.url
    });

    if (updateResult.status === 'ok') {
        showMessage('✅ Photo uploaded and linked for "' + p.name + '"!');
        await refreshAllData();
    } else {
        showMessage('⚠️ Image uploaded to Drive, but updating sheet failed: ' + updateResult.message, 'error');
    }
}

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
        Featured:     'FALSE'
    });

    if (result.status === 'ok') {
        showMessage('🗑️ Photo removed from "' + p.name + '".');
        await refreshAllData();
    } else {
        showMessage('❌ Error removing image: ' + (result.message || ''), 'error');
    }
}

async function toggleFeature(idx, checkbox) {
    const p = adminProducts[idx];
    if (!p) return;

    if (checkbox.checked) {
        const count = adminProducts.filter(function (x) { return x.isFeatured; }).length;
        if (count >= 12) {
            checkbox.checked = false;
            showMessage('⚠️ Maximum 12 items can be featured on homepage.', 'error');
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
        checkbox.checked = !checkbox.checked;
        showMessage('❌ Could not update: ' + (result.message || ''), 'error');
    }
}

// =============================================
// CATEGORIES & CONTACT
// =============================================

function loadCategories() {
    const listEl = document.getElementById('category-list');
    if (listEl) {
        const cats = getUsedCategories();
        listEl.innerHTML = cats.map(function (c) {
            return '<div style="background:#fef3c7;border:1px solid #fde68a;color:#92400e;padding:5px 14px;border-radius:20px;font-size:0.8rem;font-weight:bold;">' + esc(c) + '</div>';
        }).join('');
    }

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
    showMessage('"' + name + '" added to category dropdown.');
}

async function loadContactInfo() {
    try {
        const rows = await fetchSheet('Config');
        const get  = function (key) { return (rows.find(function (r) { return r.Key === key; }) || {}).Value || ''; };
        const gen  = document.getElementById('contact-general');
        const blk  = document.getElementById('contact-bulk');
        if (gen) gen.value = get('general_wa');
        if (blk) blk.value = get('bulk_wa');
    } catch (e) {}
}

async function saveContactInfo() {
    const general = document.getElementById('contact-general').value.trim();
    const bulk    = document.getElementById('contact-bulk').value.trim();

    showMessage('⏳ Saving WhatsApp numbers to Google Sheets…', 'info');
    const result = await sheetWrite({ action: 'saveConfig', general_wa: general, bulk_wa: bulk });

    if (result.status === 'ok') {
        showMessage('📞 Contact info saved successfully!');
    } else {
        showMessage('❌ Save failed: ' + (result.message || ''), 'error');
    }
}

async function clearAllData() {
    if (!confirm('Are you sure? This permanently deletes ALL products from the Google Sheet.')) return;
    showMessage('⏳ Clearing products…', 'info');
    const result = await sheetWrite({ action: 'clearAll' });
    if (result.status === 'ok') {
        adminProducts = [];
        loadProductTable();
        loadMenuList();
        showMessage('🗑️ All products cleared from the sheet.');
    } else {
        showMessage('❌ Error: ' + (result.message || ''), 'error');
    }
}

// =============================================
// UI HELPERS
// =============================================

function showMessage(msg, type) {
    const el = document.getElementById('status-message');
    if (!el) return;
    el.textContent       = msg;
    el.style.display     = 'block';
    el.style.background  = type === 'error' ? '#fef2f2' : type === 'info' ? '#eff6ff' : '#f0fdf4';
    el.style.color       = type === 'error' ? '#dc2626' : type === 'info' ? '#1d4ed8' : '#16a34a';
    el.style.borderColor = type === 'error' ? '#fecaca' : type === 'info' ? '#bfdbfe' : '#bbf7d0';
    if (type !== 'info') setTimeout(function () { el.style.display = 'none'; }, 5000);
}

function clearMessage() {
    const el = document.getElementById('status-message');
    if (el) el.style.display = 'none';
}

function esc(str) {
    return String(str || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}
