// =============================================
// Annapurna Sweets - Admin Panel JavaScript
// =============================================

// Default SHA-256 hash of the admin password (fallback if no custom password is set)
// The plain-text password is NEVER stored here — only its hash.
const DEFAULT_PASSWORD_HASH = '6714ed55bd7f9dbb38016e24e54b5f142bf1e6d73a9eab9032ae0822969d171f';
const STORAGE_KEY      = 'annapurna_products';
const PASSWORD_STORAGE = 'annapurna_admin_password_hash';

// --- Hash a string using SHA-256 (Web Crypto API) ---
async function sha256(message) {
    const msgBuffer = new TextEncoder().encode(message);
    const hashBuffer = await crypto.subtle.digest('SHA-256', msgBuffer);
    const hashArray  = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

// --- Get the active password hash (custom takes priority over default) ---
function getPasswordHash() {
    return localStorage.getItem(PASSWORD_STORAGE) || DEFAULT_PASSWORD_HASH;
}

// --- Password Check ---
async function checkPassword() {
    const entered     = document.getElementById('password-input').value;
    const enteredHash = await sha256(entered);

    if (enteredHash === getPasswordHash()) {
        document.getElementById('login-screen').style.display = 'none';
        document.getElementById('admin-panel').style.display  = 'block';
        loadProductTable();
        loadMenuList();
    } else {
        document.getElementById('login-error').style.display = 'block';
    }
}

// --- Change Password (called from admin panel) ---
async function changePassword() {
    const current  = document.getElementById('cp-current').value;
    const newPw    = document.getElementById('cp-new').value;
    const confirmPw = document.getElementById('cp-confirm').value;
    const msgEl    = document.getElementById('cp-message');

    // Validate current password
    const currentHash = await sha256(current);
    if (currentHash !== getPasswordHash()) {
        msgEl.textContent = '❌ Current password is incorrect.';
        msgEl.style.background = '#fef2f2'; msgEl.style.color = '#dc2626'; msgEl.style.borderColor = '#fecaca';
        msgEl.style.display = 'block';
        return;
    }

    // Validate new password
    if (newPw.length < 6) {
        msgEl.textContent = '⚠️ New password must be at least 6 characters.';
        msgEl.style.background = '#fffbeb'; msgEl.style.color = '#d97706'; msgEl.style.borderColor = '#fde68a';
        msgEl.style.display = 'block';
        return;
    }

    if (newPw !== confirmPw) {
        msgEl.textContent = '⚠️ New passwords do not match.';
        msgEl.style.background = '#fffbeb'; msgEl.style.color = '#d97706'; msgEl.style.borderColor = '#fde68a';
        msgEl.style.display = 'block';
        return;
    }

    // Hash and save new password
    const newHash = await sha256(newPw);
    localStorage.setItem(PASSWORD_STORAGE, newHash);

    // Clear fields
    document.getElementById('cp-current').value  = '';
    document.getElementById('cp-new').value      = '';
    document.getElementById('cp-confirm').value  = '';

    msgEl.textContent = '✅ Password changed successfully!';
    msgEl.style.background = '#f0fdf4'; msgEl.style.color = '#16a34a'; msgEl.style.borderColor = '#bbf7d0';
    msgEl.style.display = 'block';
    setTimeout(() => { msgEl.style.display = 'none'; }, 5000);
}

// Allow Enter key on password input
document.addEventListener('DOMContentLoaded', function () {
    const pwInput = document.getElementById('password-input');
    if (pwInput) {
        pwInput.addEventListener('keypress', function (e) {
            if (e.key === 'Enter') checkPassword();
        });
    }
});

// --- Toggle password visibility (eye button) ---
function togglePw(inputId, btn) {
    const input = document.getElementById(inputId);
    const isHidden = input.type === 'password';
    input.type    = isHidden ? 'text' : 'password';
    btn.innerHTML = isHidden ? '&#128064;' : '&#128065;'; // open-eye vs eye
    btn.title     = isHidden ? 'Hide password' : 'Show password';
}

// --- Get / Save Products from localStorage ---
function getProducts() {
    try {
        const saved = localStorage.getItem(STORAGE_KEY);
        return saved ? JSON.parse(saved) : [];
    } catch (e) {
        return [];
    }
}

function saveProducts(products) {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(products));
    } catch (e) {
        showMessage('⚠️ Storage limit reached. Try using smaller image files.', 'error');
    }
}

// --- Handle Excel Upload ---
function handleExcelUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(sheet);

            if (rows.length === 0) {
                showMessage('⚠️ The Excel file appears to be empty.', 'error');
                return;
            }

            // Preserve existing images when re-uploading
            const existing = getProducts();

            const newProducts = rows.map(function (row, i) {
                const match = existing.find(function (p) {
                    return p.name === (row.Name || '');
                });
                return {
                    id:          i + 1,
                    name:        row.Name          || 'Unnamed Product',
                    category:    row.Category      || 'Sweets',
                    pricePerPc:  row.Price_Per_Pc  ? String(row.Price_Per_Pc)  : '',
                    pricePerKg:  row.Price_Per_Kg  ? String(row.Price_Per_Kg)  : '',
                    description: row.Description   || '',
                    image:       match ? match.image : null
                };
            });

            saveProducts(newProducts);
            loadProductTable();
            showMessage('✅ ' + newProducts.length + ' products imported successfully!');
        } catch (err) {
            showMessage('❌ Could not read the file. Make sure it is a valid .xlsx or .csv file.', 'error');
            console.error(err);
        }
    };
    reader.readAsArrayBuffer(file);
}

// --- Load Product Table in Admin Panel ---
function loadProductTable() {
    const products = getProducts();
    const tbody = document.getElementById('product-table-body');
    tbody.innerHTML = '';

    if (products.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" class="empty-row">No products yet. Upload an Excel sheet above to get started.</td></tr>';
        return;
    }

    products.forEach(function (product) {
        const tr = document.createElement('tr');

        const photoCell = product.image
            ? '<img src="' + product.image + '" style="width:52px;height:52px;object-fit:cover;border-radius:8px;border:1px solid #ddd;">'
            : '<span style="color:#bbb;font-size:0.78rem;font-style:italic;">No image</span>';

        const removeBtn = product.image
            ? '<button class="remove-img-btn" onclick="removeImage(' + product.id + ')">&#128465; Remove</button>'
            : '';

        tr.innerHTML = `
            <td style="font-weight:700;color:#888;">${product.id}</td>
            <td style="font-weight:600;">${product.name}</td>
            <td>
                <span style="background:#fef9ef;border:1px solid #fde68a;color:#92400e;padding:3px 10px;border-radius:20px;font-size:0.72rem;font-weight:700;">
                    ${product.category}
                </span>
            </td>
            <td>${photoCell}</td>
            <td>
                <label style="display:inline-block;background:#f0fdf4;border:1.5px dashed #86efac;padding:6px 12px;border-radius:8px;cursor:pointer;font-size:0.75rem;color:#16a34a;font-weight:600;">
                    <input type="file" accept="image/*" onchange="handleImageUpload(event, ${product.id})" style="display:none;">
                    &#128247; Upload
                </label>
                ${removeBtn}
            </td>
        `;
        tbody.appendChild(tr);
    });
}

// --- Handle Image Upload for a Product ---
function handleImageUpload(event, productId) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const products = getProducts();
        const idx = products.findIndex(function (p) { return p.id === productId; });
        if (idx !== -1) {
            products[idx].image = e.target.result;
            saveProducts(products);
            loadProductTable();
            showMessage('✅ Photo saved for "' + products[idx].name + '"!');
        }
    };
    reader.readAsDataURL(file);
}

// --- Remove Image from a Product ---
function removeImage(productId) {
    const products = getProducts();
    const idx = products.findIndex(function (p) { return p.id === productId; });
    if (idx !== -1) {
        products[idx].image = null;
        saveProducts(products);
        loadProductTable();
        showMessage('🗑️ Image removed from "' + products[idx].name + '".');
    }
}

// =============================================
// EDIT MENU — Add / Remove Items Manually
// =============================================

let editingProductId = null;

function addOrUpdateProduct() {
    const name        = document.getElementById('new-name').value.trim();
    const category    = document.getElementById('new-category').value;
    const pricePerPc  = document.getElementById('new-price-pc').value.trim();
    const pricePerKg  = document.getElementById('new-price-kg').value.trim();
    const description = document.getElementById('new-description').value.trim();

    if (!name) {
        showMessage('⚠️ Item Name is required.', 'error');
        return;
    }

    let products = getProducts();

    if (editingProductId) {
        // Update existing item
        const index = products.findIndex(p => p.id === editingProductId);
        if (index > -1) {
            // Check if name changed and conflicts with another existing item
            const conflict = products.find(p => p.name.toLowerCase() === name.toLowerCase() && p.id !== editingProductId);
            if (conflict) {
                showMessage('⚠️ "' + name + '" already exists in the menu.', 'error');
                return;
            }

            products[index].name = name;
            products[index].category = category;
            products[index].pricePerPc = pricePerPc;
            products[index].pricePerKg = pricePerKg;
            products[index].description = description;
            saveProducts(products);
            showMessage('✅ "' + name + '" updated!');
        }
        cancelEdit(); // This also clears the form and reloads the lists
    } else {
        // Add new item
        if (products.find(function(p) { return p.name.toLowerCase() === name.toLowerCase(); })) {
            showMessage('⚠️ "' + name + '" already exists in the menu.', 'error');
            return;
        }

        const newId = products.length > 0 ? Math.max.apply(null, products.map(function(p) { return p.id; })) + 1 : 1;
        products.push({
            id:          newId,
            name:        name,
            category:    category,
            pricePerPc:  pricePerPc,
            pricePerKg:  pricePerKg,
            description: description,
            image:       null
        });
        saveProducts(products);
        showMessage('✅ "' + name + '" added to the menu!');
        
        // Clear form
        document.getElementById('new-name').value        = '';
        document.getElementById('new-price-pc').value    = '';
        document.getElementById('new-price-kg').value    = '';
        document.getElementById('new-description').value = '';
        
        loadMenuList();
        loadProductTable();
    }
}

function editProduct(productId) {
    const products = getProducts();
    const product = products.find(p => p.id === productId);
    if (!product) return;
    
    editingProductId = product.id;
    
    // Populate form
    document.getElementById('new-name').value = product.name;
    document.getElementById('new-category').value = product.category;
    document.getElementById('new-price-pc').value = product.pricePerPc || product.price || '';
    document.getElementById('new-price-kg').value = product.pricePerKg || '';
    document.getElementById('new-description').value = product.description || '';
    
    // Change UI
    document.getElementById('submit-btn').innerHTML = '&#10004; Update Item';
    document.getElementById('submit-btn').style.background = '#16a34a'; // Green
    document.getElementById('submit-btn').style.color = 'white';
    document.getElementById('cancel-edit-btn').style.display = 'inline-block';
    
    // Scroll to form
    document.getElementById('new-name').scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function cancelEdit() {
    editingProductId = null;
    
    // Clear form
    document.getElementById('new-name').value        = '';
    document.getElementById('new-price-pc').value    = '';
    document.getElementById('new-price-kg').value    = '';
    document.getElementById('new-description').value = '';
    
    // Revert UI
    document.getElementById('submit-btn').innerHTML = '&#43; Add Item to Menu';
    document.getElementById('submit-btn').style.background = 'var(--gold)';
    document.getElementById('submit-btn').style.color = 'var(--brown-dark)';
    document.getElementById('cancel-edit-btn').style.display = 'none';

    loadMenuList();
    loadProductTable();
}

// --- Delete a Product ---
function deleteProduct(productId) {
    const products = getProducts();
    const product  = products.find(function(p) { return p.id === productId; });
    if (!product) return;

    if (confirm('Remove "' + product.name + '" from the menu?')) {
        const updated = products.filter(function(p) { return p.id !== productId; });
        // Re-number IDs
        updated.forEach(function(p, i) { p.id = i + 1; });
        saveProducts(updated);
        loadMenuList();
        loadProductTable();
        showMessage('🗑️ "' + product.name + '" removed from the menu.');
    }
}

// --- Load the Visual Menu List (for Edit Menu section) ---
function loadMenuList() {
    const products  = getProducts();
    const container = document.getElementById('menu-list');
    if (!container) return;

    if (products.length === 0) {
        container.innerHTML = '<p style="padding:16px;font-style:italic;color:#aaa;">No items yet. Add one using the form above.</p>';
        return;
    }

    container.innerHTML = '';
    products.forEach(function(product) {
        const pc  = product.pricePerPc || product.price || '';
        const kg  = product.pricePerKg || '';
        const row = document.createElement('div');
        row.className = 'menu-item-row';
        row.innerHTML = `
            <div class="menu-item-info">
                <span class="menu-item-name">${product.name}</span>
                <span class="menu-item-cat">${product.category}</span>
                ${pc ? '<span class="menu-item-price">&#8377;' + pc + '/pc</span>' : ''}
                ${kg ? '<span class="menu-item-price">&#8377;' + kg + '/kg</span>' : ''}
                ${product.description ? '<span class="menu-item-desc">' + product.description + '</span>' : ''}
            </div>
            <div style="display:flex; gap:8px;">
                <button class="edit-item-btn" onclick="editProduct(${product.id})" style="background:#eab308; color:#fff; border:none; padding:6px 10px; border-radius:4px; cursor:pointer; font-size:0.75rem;">✏️ Edit</button>
                <button class="delete-item-btn" onclick="deleteProduct(${product.id})">&#128465; Remove</button>
            </div>
        `;
        container.appendChild(row);
    });
}

// --- Clear All Data ---
function clearAllData() {
    if (confirm('Are you sure? This will delete ALL products and images.')) {
        try {
            localStorage.removeItem(STORAGE_KEY);
        } catch (e) {}
        loadProductTable();
        showMessage('🗑️ All data has been cleared.');
    }
}

// --- Logout ---
function logout() {
    document.getElementById('password-input').value = '';
    document.getElementById('admin-panel').style.display = 'none';
    document.getElementById('login-screen').style.display = 'flex';
    showMessage('🚪 Logged out securely.');
}

// --- Show Status Message ---
function showMessage(msg, type) {
    const el = document.getElementById('status-message');
    el.textContent = msg;
    el.style.display = 'block';
    el.style.background  = (type === 'error') ? '#fef2f2' : '#f0fdf4';
    el.style.color       = (type === 'error') ? '#dc2626' : '#16a34a';
    el.style.borderColor = (type === 'error') ? '#fecaca' : '#bbf7d0';
    setTimeout(function () { el.style.display = 'none'; }, 5000);
}
