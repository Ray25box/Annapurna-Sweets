// =============================================
// Annapurna Sweets - Admin Panel JavaScript
// =============================================

const CORRECT_PASSWORD = 'Arjya@403';
const STORAGE_KEY = 'annapurna_products';

// --- Password Check ---
function checkPassword() {
    const entered = document.getElementById('password-input').value;
    if (entered === CORRECT_PASSWORD) {
        document.getElementById('login-screen').style.display = 'none';
        document.getElementById('admin-panel').style.display = 'block';
        loadProductTable();
    } else {
        document.getElementById('login-error').style.display = 'block';
    }
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
                    id: i + 1,
                    name:        row.Name        || 'Unnamed Product',
                    category:    row.Category    || 'Sweets',
                    price:       row.Price        || '',
                    description: row.Description || '',
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
