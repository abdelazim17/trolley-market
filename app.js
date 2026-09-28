const CATEGORIES = [
    "الفواكه والخضار", "التسالي والمكسرات", "المشروبات",
    "البقالة", "الألبان", "المخبوزات",
    "التوابل والبهارات", "الحلويات والكاندي", "المنظفات"
];

let allProducts = [];

// تحميل البيانات من الإكسيل
async function loadExcelData() {
    try {
        const response = await fetch('products.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);
        
        // فلترة المنتجات المتاحة فقط (متوفر = نعم)
        allProducts = jsonData.filter(item => item['متوفر'] === 'نعم');
        allProducts.forEach((p, idx) => p.id = 'prod_' + idx);
        
        const loading = document.getElementById('loading');
        if(loading) loading.style.display = 'none';

        const isOffersPage = window.location.pathname.includes('offers');
        
        if (isOffersPage) {
            const offers = allProducts.filter(p => p['السعر بعد الخصم'] && p['السعر بعد الخصم'] !== '');
            renderGrid(offers, 'offers-grid');
        } else {
            initCategoriesFilter();
            renderHomeSections();
        }
    } catch (error) {
        console.error(error);
        const loading = document.getElementById('loading');
        if(loading) loading.innerHTML = "يرجى رفع المشروع على سيرفر (أو جيت هاب) ليعمل الإكسيل بشكل صحيح.";
    }
}

function initCategoriesFilter() {
    const container = document.getElementById('categories-text-container');
    if(!container) return;
    
    CATEGORIES.forEach(cat => {
        const btn = document.createElement('button');
        btn.className = 'cat-btn';
        btn.setAttribute('data-category', cat);
        btn.innerText = cat;
        btn.onclick = (e) => {
            document.querySelectorAll('.cat-btn').forEach(b => b.classList.remove('active'));
            e.target.classList.add('active');
            filterHome(cat);
        };
        container.appendChild(btn);
    });
    
    // زرار الكل
    document.querySelector('.cat-btn[data-category="الكل"]').onclick = (e) => {
        document.querySelectorAll('.cat-btn').forEach(b => b.classList.remove('active'));
        e.target.classList.add('active');
        renderHomeSections();
    };
}

function getProductHTML(p) {
    let hasOffer = p['السعر بعد الخصم'] && p['السعر بعد الخصم'] !== '';
    let priceHTML = hasOffer ? 
        `<span class="new-price">${p['السعر بعد الخصم']} ج</span> <span class="old-price">${p['السعر الأساسي']} ج</span>` :
        `<span class="new-price">${p['السعر الأساسي']} ج</span>`;
        
    let currentPrice = hasOffer ? p['السعر بعد الخصم'] : p['السعر الأساسي'];
    let imgSrc = p['الصورة'] ? p['الصورة'] : 'images/placeholder.png';

    return `
    <div class="product-card">
        <img src="${imgSrc}" alt="${p['الاسم']}" class="product-img" onerror="this.outerHTML='<div style=\'height:140px; background:#f9f9f9; margin-bottom:10px; display:flex; align-items:center; justify-content:center;\'>بدون صورة</div>'">
        <div class="product-title">${p['الاسم']}</div>
        <div class="price-container">${priceHTML} / ${p['الوحدة']}</div>
        
        <div class="qty-controls">
            <button class="qty-btn" onclick="changeQty('${p.id}', 1)">+</button>
            <input type="number" id="qty-${p.id}" class="qty-input" value="1" readonly>
            <button class="qty-btn" onclick="changeQty('${p.id}', -1)">-</button>
        </div>
        
        <button class="add-to-cart-btn" onclick="addToCart('${p.id}', '${p['الاسم']}', ${currentPrice})">
            إضافة للسلة
        </button>
    </div>
    `;
}

function renderHomeSections() {
    const container = document.getElementById('home-sections');
    if(!container) return;
    
    container.innerHTML = '';
    
    CATEGORIES.forEach(cat => {
        const catProducts = allProducts.filter(p => p['القسم'] === cat);
        if(catProducts.length > 0) {
            const section = document.createElement('div');
            section.className = 'category-section';
            
            const header = document.createElement('div');
            header.className = 'section-header';
            header.innerHTML = `<h3>${cat}</h3>`;
            
            const grid = document.createElement('div');
            grid.className = 'products-grid';
            grid.innerHTML = catProducts.map(p => getProductHTML(p)).join('');
            
            section.appendChild(header);
            section.appendChild(grid);
            container.appendChild(section);
        }
    });
}

function filterHome(category) {
    const container = document.getElementById('home-sections');
    container.innerHTML = '';
    
    const catProducts = allProducts.filter(p => p['القسم'] === category);
    if(catProducts.length > 0) {
        const grid = document.createElement('div');
        grid.className = 'products-grid';
        grid.innerHTML = catProducts.map(p => getProductHTML(p)).join('');
        container.appendChild(grid);
    }
}

function renderGrid(products, gridId) {
    const grid = document.getElementById(gridId);
    if(!grid) return;
    grid.innerHTML = products.map(p => getProductHTML(p)).join('');
}

// العداد (+ و -) الحد الأقصى 25
window.changeQty = function(id, delta) {
    const input = document.getElementById(`qty-${id}`);
    if(input) {
        let val = parseInt(input.value) + delta;
        if(val < 1) val = 1;
        if(val > 25) val = 25;
        input.value = val;
    }
}

// السلة الجانبية
window.toggleCart = function() {
    document.getElementById('cart-sidebar').classList.toggle('open');
}

let cart = JSON.parse(localStorage.getItem('cart')) || [];
updateCartUI();

window.addToCart = function(id, name, price) {
    const input = document.getElementById(`qty-${id}`);
    const qtyToAdd = parseInt(input.value) || 1;
    
    const existing = cart.find(item => item.name === name);
    if (existing) {
        existing.qty += qtyToAdd;
        if(existing.qty > 50) existing.qty = 50; 
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
    cart = [];
    saveCart(); updateCartUI();
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
    
    // الرسالة المصرية المخصصة للواتساب
    let whatsappMsg = "مرحباً ترولي ماركت، أنا عاوز أطلب الحاجات دي:%0a%0a";
    cart.forEach(item => { whatsappMsg += `- ${item.name} (الكمية: ${item.qty}) = ${item.price * item.qty} ج%0a`; });
    whatsappMsg += `%0aالإجمالي: ${total} جنيه`;
    
    const checkoutBtn = document.getElementById('checkout-btn');
    if(checkoutBtn) checkoutBtn.href = `https://wa.me/201063883209?text=${whatsappMsg}`;
}

window.onload = loadExcelData;
