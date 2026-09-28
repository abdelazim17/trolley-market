const CATEGORIES = [
    "الفواكه والخضار", "التسالي والمكسرات", "المشروبات",
    "البقالة", "الألبان", "المخبوزات",
    "التوابل والبهارات", "الحلويات والكاندي", "المنظفات"
];

let allProducts = [];

async function loadExcelData() {
    try {
        const response = await fetch('products.xlsx');
        if (!response.ok) throw new Error('Network response was not ok');
        
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        allProducts = jsonData.filter(item => item['متوفر'] === 'نعم');
        allProducts.forEach((p, idx) => p.id = 'prod_' + idx);
        
        const loading = document.getElementById('loading');
        if(loading) loading.style.display = 'none';
        
        if(allProducts.length === 0) {
            document.getElementById('error-box').innerHTML = "ملف الإكسيل موجود ولكنه فارغ أو لا يوجد منتجات متوفرة.";
            document.getElementById('error-box').style.display = 'block';
            return;
        }

        initCatSlider();

        const isHomePage = window.location.pathname.includes('index') || window.location.pathname.endsWith('/');
        const isCategoriesPage = window.location.pathname.includes('categories');

        if(isHomePage) {
            renderTodayOffers();
        } else if(isCategoriesPage) {
            renderAllCategories();
        }

        updateCartUI();

    } catch (error) {
        console.error("خطأ في التحميل:", error);
        const loading = document.getElementById('loading');
        if(loading) loading.style.display = 'none';
        
        const errorBox = document.getElementById('error-box');
        if(errorBox) errorBox.style.display = 'block';
    }
}

// سلايدر الأقسام (صور دائرية)
function initCatSlider() {
    const slider = document.getElementById('cat-visual-slider');
    if(!slider) return;
    
    slider.innerHTML = '';
    CATEGORIES.forEach((cat, index) => {
        const item = document.createElement('a');
        item.className = 'cat-item';
        item.href = 'categories.html';
        item.innerHTML = `
            <img src="images/cat_${index}.jpg" alt="${cat}" onerror="this.src='images/placeholder.png'">
            <span>${cat}</span>
        `;
        slider.appendChild(item);
    });
}

function getProductHTML(p) {
    let hasOffer = p['السعر بعد الخصم'] && p['السعر بعد الخصم'] !== '';
    let priceHTML = hasOffer ? 
        `<span class="new-price">${p['السعر بعد الخصم']} ج</span> <span class="old-price">${p['السعر الأساسي']} ج</span>` :
        `<span class="new-price">${p['السعر الأساسي']} ج</span>`;
        
    let currentPrice = hasOffer ? p['السعر بعد الخصم'] : p['السعر الأساسي'];
    let imgSrc = p['الصورة'] && p['الصورة'] !== 'images/placeholder.png' ? p['الصورة'] : 'images/placeholder.png';

    return `
    <div class="product-card">
        <img src="${imgSrc}" class="product-img" onerror="this.src='images/placeholder.png'">
        <div class="product-title">${p['الاسم']}</div>
        <div class="price-container">${priceHTML} / ${p['الوحدة'] || 'وحدة'}</div>
        
        <div class="qty-controls">
            <button class="qty-btn" onclick="changeQty('${p.id}', 1)">+</button>
            <input type="number" id="qty-${p.id}" class="qty-input" value="1" readonly>
            <button class="qty-btn" onclick="changeQty('${p.id}', -1)">-</button>
        </div>
        
        <button class="add-to-cart-btn" onclick="addToCart('${p.id}', '${p['الاسم']}', ${currentPrice})">إضافة للسلة</button>
    </div>
    `;
}

// عرض عروض اليوم في الصفحة الرئيسية
function renderTodayOffers() {
    const container = document.getElementById('home-sections');
    if(!container) return;
    container.innerHTML = '';
    
    const offersProducts = allProducts.filter(p => p['القسم'] === 'عروض اليوم');
    
    if(offersProducts.length > 0) {
        const section = document.createElement('div');
        section.className = 'category-section';
        section.innerHTML = `
            <div class="section-header"><h3>🔥 عروض اليوم</h3></div>
            <div class="products-grid">
                ${offersProducts.map(p => getProductHTML(p)).join('')}
            </div>
        `;
        container.appendChild(section);
    } else {
        container.innerHTML = '<p style="text-align:center; font-size:1.2rem;">لا توجد عروض اليوم حالياً.</p>';
    }
}

// عرض جميع الأقسام في صفحة الأقسام
function renderAllCategories() {
    const container = document.getElementById('categories-sections');
    if(!container) return;
    container.innerHTML = '';
    
    CATEGORIES.forEach(cat => {
        const catProducts = allProducts.filter(p => p['القسم'] === cat);
        if(catProducts.length > 0) {
            const section = document.createElement('div');
            section.className = 'category-section';
            section.innerHTML = `
                <div class="section-header"><h3>${cat}</h3></div>
                <div class="products-grid">
                    ${catProducts.map(p => getProductHTML(p)).join('')}
                </div>
            `;
            container.appendChild(section);
        }
    });
}

window.changeQty = function(id, delta) {
    const input = document.getElementById(`qty-${id}`);
    if(input) {
        let val = parseInt(input.value) + delta;
        if(val < 1) val = 1;
        if(val > 25) val > 25 ? val = 25 : null;
        input.value = val;
    }
}

// السلة
let cart = JSON.parse(localStorage.getItem('cart')) || [];

window.toggleCart = function() {
    document.getElementById('cart-sidebar').classList.toggle('open');
}

window.addToCart = function(id, name, price) {
    const input = document.getElementById(`qty-${id}`);
    const qtyToAdd = parseInt(input.value) || 1;
    
    const existing = cart.find(item => item.name === name);
    if (existing) {
        existing.qty += qtyToAdd;
    } else {
        cart.push({ name, price, qty: qtyToAdd });
    }
    
    input.value = 1;
    saveCart(); 
    updateCartUI();
    document.getElementById('cart-sidebar').classList.add('open');
}

window.removeFromCart = function(name) {
    cart = cart.filter(item => item.name !== name);
    saveCart(); updateCartUI();
}

window.clearCart = function() {
    if(confirm('هل أنت متأكد من إزالة كل الطلبات من السلة؟')) {
        cart = [];
        saveCart(); updateCartUI();
    }
}

function saveCart() { localStorage.setItem('cart', JSON.stringify(cart)); }

function updateCartUI() {
    const cartCount = document.getElementById('cart-count');
    if(cartCount) cartCount.innerText = cart.reduce((sum, item) => sum + item.qty, 0);
    
    const cartItemsDiv = document.getElementById('cart-items');
    if(!cartItemsDiv) return;
    
    cartItemsDiv.innerHTML = '';
    let total = 0;
    
    if(cart.length === 0) {
        cartItemsDiv.innerHTML = '<p style="text-align:center; color:#999; margin-top:20px;">السلة فارغة</p>';
    } else {
        cart.forEach(item => {
            total += item.price * item.qty;
            cartItemsDiv.innerHTML += `
                <div class="cart-item">
                    <div>
                        <strong style="color:var(--dark);">${item.name}</strong><br>
                        <small style="color:var(--secondary);">${item.qty} &times; ${item.price} ج</small>
                    </div>
                    <div style="text-align:left;">
                        <strong style="color:var(--primary);">${item.price * item.qty} ج</strong><br>
                        <button onclick="removeFromCart('${item.name}')" style="background:none; border:none; color:red; cursor:pointer; font-size:0.9rem; margin-top:5px;"><i class="fas fa-trash"></i> حذف</button>
                    </div>
                </div>
            `;
        });
    }
    
    const totalPriceSpan = document.getElementById('total-price');
    if(totalPriceSpan) totalPriceSpan.innerText = total;
    
    let whatsappMsg = "مرحباً ترولي ماركت، أنا عاوز أطلب الحاجات دي:%0a%0a";
    cart.forEach(item => { whatsappMsg += `- ${item.name} (الكمية: ${item.qty}) = ${item.price * item.qty} ج%0a`; });
    whatsappMsg += `%0aالإجمالي: ${total} جنيه`;
    
    const checkoutBtn = document.getElementById('checkout-btn');
    if(checkoutBtn) checkoutBtn.href = `https://wa.me/201063883209?text=${whatsappMsg}`;
}

window.onload = loadExcelData;
