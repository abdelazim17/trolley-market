// --- إدارة الوضع الليلي والنهاري ---
const themeToggle = document.getElementById('theme-toggle');
const currentTheme = localStorage.getItem('theme') || 'light';
document.documentElement.setAttribute('data-theme', currentTheme);
updateThemeIcon(currentTheme);

themeToggle.addEventListener('click', () => {
    let theme = document.documentElement.getAttribute('data-theme');
    let newTheme = theme === 'light' ? 'dark' : 'light';
    document.documentElement.setAttribute('data-theme', newTheme);
    localStorage.setItem('theme', newTheme);
    updateThemeIcon(newTheme);
});

function updateThemeIcon(theme) {
    themeToggle.innerHTML = theme === 'light' ? '<i class="fas fa-moon"></i>' : '<i class="fas fa-sun"></i>';
}

// --- تحميل البيانات من ملف الإكسيل ---
let allProducts = [];

async function loadExcelData() {
    try {
        const response = await fetch('products.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        
        // قراءة الملف باستخدام SheetJS
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // تحويل البيانات إلى مصفوفة جافاسكريبت
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        // تصفية المنتجات المفعلة فقط (حالة = 1)
        allProducts = jsonData.filter(item => item['الحالة'] == 1);
        
        document.getElementById('loading').style.display = 'none';
        renderProducts(allProducts);
    } catch (error) {
        document.getElementById('loading').innerHTML = "حدث خطأ أثناء تحميل المنتجات. تأكد من وجود ملف products.xlsx";
        console.error("Error loading Excel:", error);
    }
}

function renderProducts(products) {
    const grid = document.getElementById('products-grid');
    grid.innerHTML = '';
    
    products.forEach((product, index) => {
        // تأكد من مسار الصورة، لو مفيش حط صورة افتراضية
        const imgSrc = product['الصورة'] ? product['الصورة'] : 'images/placeholder.png';
        
        const card = document.createElement('div');
        card.className = 'product-card';
        card.innerHTML = `
            <img src="${imgSrc}" alt="${product['الاسم']}" class="product-img" onerror="this.src='images/placeholder.png'">
            <div class="product-title">${product['الاسم']}</div>
            <div class="product-price">${product['السعر']} جنيه / ${product['الوحدة'] || 'وحدة'}</div>
            <button class="add-to-cart-btn" onclick="addToCart('${product['الاسم']}', ${product['السعر']})">
                <i class="fas fa-plus"></i> أضف
            </button>
        `;
        grid.appendChild(card);
    });
}

// --- نظام الفلترة (الأقسام) ---
const catButtons = document.querySelectorAll('.cat-btn');
catButtons.forEach(btn => {
    btn.addEventListener('click', (e) => {
        catButtons.forEach(b => b.classList.remove('active'));
        e.target.classList.add('active');
        
        const category = e.target.getAttribute('data-category');
        if (category === 'الكل') {
            renderProducts(allProducts);
        } else {
            const filtered = allProducts.filter(p => p['القسم'] === category);
            renderProducts(filtered);
        }
    });
});

// --- نظام سلة المشتريات ---
let cart = JSON.parse(localStorage.getItem('cart')) || [];
updateCartUI();

function addToCart(name, price) {
    const existing = cart.find(item => item.name === name);
    if (existing) {
        existing.qty += 1;
    } else {
        cart.push({ name, price, qty: 1 });
    }
    saveCart();
    updateCartUI();
}

function removeFromCart(name) {
    cart = cart.filter(item => item.name !== name);
    saveCart();
    updateCartUI();
}

function saveCart() {
    localStorage.setItem('cart', JSON.stringify(cart));
}

function updateCartUI() {
    document.getElementById('cart-count').innerText = cart.reduce((sum, item) => sum + item.qty, 0);
    
    const cartItemsDiv = document.getElementById('cart-items');
    cartItemsDiv.innerHTML = '';
    let total = 0;
    
    cart.forEach(item => {
        total += item.price * item.qty;
        cartItemsDiv.innerHTML += `
            <div class="cart-item">
                <div>${item.name} (x${item.qty})</div>
                <div>${item.price * item.qty} جنيه</div>
                <button onclick="removeFromCart('${item.name}')" style="color:red; background:none; border:none; cursor:pointer;"><i class="fas fa-trash"></i></button>
            </div>
        `;
    });
    document.getElementById('total-price').innerText = total;
    
    // تجهيز رسالة الواتساب
    let whatsappMsg = "مرحباً، أود طلب الآتي:%0a";
    cart.forEach(item => {
        whatsappMsg += `- ${item.name} (الكمية: ${item.qty})%0a`;
    });
    whatsappMsg += `%0aالإجمالي: ${total} جنيه`;
    document.getElementById('checkout-btn').href = `https://wa.me/201063883209?text=${whatsappMsg}`;
}

// --- نوافذ السلة (Modal) ---
const modal = document.getElementById('cart-modal');
document.getElementById('cart-btn').onclick = () => modal.style.display = 'block';
document.querySelector('.close-btn').onclick = () => modal.style.display = 'none';
window.onclick = (e) => { if (e.target == modal) modal.style.display = 'none'; }

// تشغيل جلب البيانات عند بدء الصفحة
window.onload = loadExcelData;
