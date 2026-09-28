// Theme Setup
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

const CATEGORIES = [
    "الفواكه والخضار", "التسالي والمكسرات", "المشروبات",
    "البقالة", "الألبان", "المخبوزات",
    "التوابل والبهارات", "الحلويات والكاندي", "المنظفات"
];

// Initialize UI
function initUI() {
    const visualSlider = document.getElementById('cat-visual-slider');
    const textContainer = document.getElementById('categories-text-container');
    
    CATEGORIES.forEach((cat, index) => {
        // Visual Slider
        const div = document.createElement('div');
        div.className = 'cat-item';
        div.onclick = () => filterCategory(cat);
        div.innerHTML = `
            <img src="images/cat_${index}.png" alt="${cat}" onerror="this.src='images/placeholder.png'">
            <span>${cat}</span>
        `;
        visualSlider.appendChild(div);

        // Text Buttons
        const btn = document.createElement('button');
        btn.className = 'cat-btn';
        btn.setAttribute('data-category', cat);
        btn.innerText = cat;
        btn.onclick = (e) => {
            document.querySelectorAll('.cat-btn').forEach(b => b.classList.remove('active'));
            e.target.classList.add('active');
            filterCategory(cat);
        };
        textContainer.appendChild(btn);
    });
}

// Data Loading
let allProducts = [];
async function loadExcelData() {
    initUI();
    try {
        const response = await fetch('products.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        allProducts = jsonData.filter(item => item['الحالة'] == 1);
        
        // إعطاء كل منتج ID فريد لتسهيل التعامل مع الكميات
        allProducts.forEach((p, idx) => p.id = 'prod_' + idx);
        
        document.getElementById('loading').style.display = 'none';
        renderHome();
    } catch (error) {
        document.getElementById('loading').innerHTML = "حدث خطأ أثناء تحميل المنتجات. تأكد من وجود ملف products.xlsx";
    }
}

// Render Product Card HTML
function getProductCardHTML(product) {
    const imgSrc = product['الصورة'] ? product['الصورة'] : 'images/placeholder.png';
    return `
        <div class="product-card">
            <img src="${imgSrc}" alt="${product['الاسم']}" class="product-img" onerror="this.src='images/placeholder.png'">
            <div class="product-title">${product['الاسم']}</div>
            <div class="product-price">${product['السعر']} جنيه / ${product['الوحدة'] || 'وحدة'}</div>
            
            <div class="qty-controls">
                <button class="qty-btn" onclick="changeQty('${product.id}', 1)">+</button>
                <input type="number" id="qty-${product.id}" class="qty-input" value="1" min="1" readonly>
                <button class="qty-btn" onclick="changeQty('${product.id}', -1)">-</button>
            </div>
            
            <button class="add-to-cart-btn" onclick="addToCart('${product.id}', '${product['الاسم']}', ${product['السعر']})">
                <i class="fas fa-cart-plus"></i> أضف للسلة
            </button>
        </div>
    `;
}

// Render Homepage (Sections)
function renderHome() {
    document.querySelectorAll('.cat-btn').forEach(b => b.classList.remove('active'));
    document.querySelector('.cat-btn[data-category="الكل"]').classList.add('active');
    
    document.getElementById('filtered-grid').style.display = 'none';
    const homeSections = document.getElementById('home-sections');
    homeSections.style.display = 'block';
    homeSections.innerHTML = '';
    
    CATEGORIES.forEach(cat => {
        const catProducts = allProducts.filter(p => p['القسم'] === cat);
        if(catProducts.length > 0) {
            const displayProducts = catProducts.slice(0, 6); // عرض 6 فقط
            
            let gridHTML = displayProducts.map(p => getProductCardHTML(p)).join('');
            
            const section = document.createElement('div');
            section.className = 'category-section';
            section.innerHTML = `
                <div class="section-header">
                    <h3>${cat}</h3>
                    <button class="view-all-btn" onclick="filterCategory('${cat}')">عرض الكل &larr;</button>
                </div>
                <div class="section-grid">
                    ${gridHTML}
                </div>
            `;
            homeSections.appendChild(section);
        }
    });
}

// Filter specific category
function filterCategory(category) {
    // Update text buttons
    document.querySelectorAll('.cat-btn').forEach(b => {
        b.classList.remove('active');
        if(b.getAttribute('data-category') === category) b.classList.add('active');
    });

    document.getElementById('home-sections').style.display = 'none';
    const grid = document.getElementById('filtered-grid');
    grid.style.display = 'grid';
    grid.innerHTML = '';
    
    const filtered = allProducts.filter(p => p['القسم'] === category);
    grid.innerHTML = filtered.map(p => getProductCardHTML(p)).join('');
    
    // Scroll to products
    document.querySelector('.categories-wrapper').scrollIntoView({behavior: "smooth"});
}

document.querySelector('.cat-btn[data-category="الكل"]').onclick = renderHome;

// Quantity Change
window.changeQty = function(id, delta) {
    const input = document.getElementById(`qty-${id}`);
    if(input) {
        let val = parseInt(input.value) + delta;
        if(val < 1) val = 1;
        input.value = val;
    }
}

// Cart Logic
let cart = JSON.parse(localStorage.getItem('cart')) || [];
updateCartUI();

window.addToCart = function(id, name, price) {
    const input = document.getElementById(`qty-${id}`);
    const qtyToAdd = parseInt(input.value) || 1;
    
    const existing = cart.find(item => item.name === name);
    if (existing) {
        existing.qty += qtyToAdd;
    } else {
        cart.push({ name, price, qty: qtyToAdd });
    }
    
    // Reset input
    input.value = 1;
    
    saveCart(); 
    updateCartUI();
    
    // إشعار بصري بسيط
    const btn = document.getElementById('cart-btn');
    btn.style.transform = 'scale(1.1)';
    setTimeout(() => btn.style.transform = 'scale(1)', 200);
}

window.removeFromCart = function(name) {
    cart = cart.filter(item => item.name !== name);
    saveCart(); updateCartUI();
}

window.clearCart = function() {
    if(confirm('هل أنت متأكد من إفراغ السلة بالكامل؟')) {
        cart = [];
        saveCart(); updateCartUI();
    }
}

function saveCart() { localStorage.setItem('cart', JSON.stringify(cart)); }

function updateCartUI() {
    document.getElementById('cart-count').innerText = cart.reduce((sum, item) => sum + item.qty, 0);
    const cartItemsDiv = document.getElementById('cart-items');
    cartItemsDiv.innerHTML = '';
    let total = 0;
    
    if(cart.length === 0) {
        cartItemsDiv.innerHTML = '<p class="text-center" style="color:#777; margin: 20px 0;">السلة فارغة</p>';
    } else {
        cart.forEach(item => {
            total += item.price * item.qty;
            cartItemsDiv.innerHTML += `
                <div class="cart-item">
                    <div>
                        <strong>${item.name}</strong><br>
                        <small>الكمية: ${item.qty} &times; ${item.price} ج</small>
                    </div>
                    <div style="text-align:left;">
                        <strong>${item.price * item.qty} ج</strong><br>
                        <button onclick="removeFromCart('${item.name}')" style="color:#e74c3c; background:none; border:none; cursor:pointer; margin-top:5px;"><i class="fas fa-trash"></i> حذف</button>
                    </div>
                </div>
            `;
        });
    }
    document.getElementById('total-price').innerText = total;
    
    let whatsappMsg = "مرحباً، أود طلب الآتي:%0a%0a";
    cart.forEach(item => { whatsappMsg += `- ${item.name} (الكمية: ${item.qty}) = ${item.price * item.qty} جنيه%0a`; });
    whatsappMsg += `%0aالإجمالي: ${total} جنيه`;
    document.getElementById('checkout-btn').href = `https://wa.me/201063883209?text=${whatsappMsg}`;
}

window.copyLink = function() {
    const copyText = document.getElementById("site-url");
    copyText.select();
    document.execCommand("copy");
    alert("تم نسخ الرابط بنجاح!");
}

// Modals Setup
document.getElementById('cart-btn').onclick = () => document.getElementById('cart-modal').style.display = 'block';
window.onclick = (e) => { 
    if (e.target.classList.contains('modal')) e.target.style.display = 'none'; 
}

window.onload = loadExcelData;
