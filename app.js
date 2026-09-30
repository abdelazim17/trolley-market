
let allOffers = []; let allProducts = []; let CATEGORIES = [];
let sliderData = []; let bannersData = {}; let catalogData = []; let allSettings = {};
let currentSliderIndex = 0;

function initTheme() {
    if (localStorage.getItem('theme') === 'dark') document.body.classList.add('dark-mode');
    updateThemeIcon();
}

window.toggleTheme = function() {
    document.body.classList.toggle('dark-mode');
    localStorage.setItem('theme', document.body.classList.contains('dark-mode') ? 'dark' : 'light');
    updateThemeIcon();
}

function updateThemeIcon() {
    const icon = document.getElementById('theme-icon');
    if (icon) icon.className = document.body.classList.contains('dark-mode') ? 'fas fa-sun' : 'fas fa-moon';
}

async function loadExcelData() {
    try {
        const response = await fetch('data.xlsx');
        if (!response.ok) throw new Error('File not found');
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        
        try { if(workbook.Sheets['السلايدر']) sliderData = XLSX.utils.sheet_to_json(workbook.Sheets['السلايدر']); } catch(e){}
        try { 
            if(workbook.Sheets['البنرات']) {
                let b = XLSX.utils.sheet_to_json(workbook.Sheets['البنرات']);
                b.forEach(item => bannersData[item['مكان_البانر']] = item['صورة']);
            }
        } catch(e){}
        try {
            if(workbook.Sheets['عروض اليوم']) {
                allOffers = XLSX.utils.sheet_to_json(workbook.Sheets['عروض اليوم']).filter(i => i['متوفر'] !== 'لا');
                allOffers.forEach(p => p.id = 'off_' + Math.random().toString(36).substr(2, 9));
            }
        } catch(e){}
        try { if(workbook.Sheets['مجلة العروض']) catalogData = XLSX.utils.sheet_to_json(workbook.Sheets['مجلة العروض']); } catch(e){}
        try {
            if(workbook.Sheets['الإعدادات']) {
                let s = XLSX.utils.sheet_to_json(workbook.Sheets['الإعدادات']);
                s.forEach(item => allSettings[item['الإعداد']] = item['القيمة']);
            }
        } catch(e){}
        
        const excludeSheets = ['السلايدر', 'البنرات', 'عروض اليوم', 'مجلة العروض', 'الإعدادات'];
        CATEGORIES = workbook.SheetNames.filter(name => !excludeSheets.includes(name));
        
        CATEGORIES.forEach(cat => {
            try {
                let catProds = XLSX.utils.sheet_to_json(workbook.Sheets[cat]).filter(i => i['متوفر'] !== 'لا');
                catProds.forEach(p => { p['القسم'] = cat; p.id = 'prod_' + Math.random().toString(36).substr(2, 9); });
                allProducts = allProducts.concat(catProds);
            } catch(e){}
        });

        if (allSettings['أيقونة الموقع']) {
            let link = document.querySelector("link[rel~='icon']");
            if (!link) { link = document.createElement('link'); link.rel = 'icon'; document.head.appendChild(link); }
            link.href = allSettings['أيقونة الموقع'];
        }

        if(document.getElementById('loading')) document.getElementById('loading').style.display = 'none';

        const path = window.location.pathname;
        if(path.includes('index') || path.endsWith('/') || path.endsWith('.html') && !path.includes('categories') && !path.includes('about') && !path.includes('catalog')) {
            renderTodayOffers();
            startBannerSlider();
            applyBanners();
        } else if(path.includes('categories')) {
            initCategoryFilterBar();
            let targetCat = localStorage.getItem('selectedCategory') || 'الكل';
            localStorage.removeItem('selectedCategory');
            filterCategory(targetCat);
        } else if(path.includes('about')) {
            applyAboutSettings();
        } else if(path.includes('catalog')) {
            renderCatalog();
        }

        updateCartUI();
    } catch (error) {
        console.error("Error loading Excel:", error);
        if(document.getElementById('loading')) {
            document.getElementById('loading').innerHTML = '<div style="color:red; font-size:1.2rem; padding: 20px;">خطأ: تأكد من رفع ملف data.xlsx بشكل صحيح!</div>';
        }
        if(window.location.pathname.includes('about')) applyAboutSettings(); 
    }
}

function startBannerSlider() {
    const sliderLink = document.getElementById('slider-link');
    const imgElem = document.getElementById('slider-banner-img');
    const badgeElem = document.getElementById('slider-cat-name');
    
    if(!imgElem || sliderData.length === 0) {
        if(imgElem) imgElem.parentElement.style.display = 'none';
        return;
    }

    function showSlide() {
        const slide = sliderData[currentSliderIndex];
        imgElem.style.opacity = 0;
        
        setTimeout(() => {
            imgElem.src = slide['صورة'] || 'images/placeholder.png';
            if(badgeElem) {
                if(slide['القسم_المستهدف']) {
                    badgeElem.innerText = slide['القسم_المستهدف'];
                    badgeElem.style.display = 'block';
                } else {
                    badgeElem.style.display = 'none';
                }
            }
            sliderLink.onclick = (e) => {
                e.preventDefault();
                if(slide['القسم_المستهدف']) {
                    localStorage.setItem('selectedCategory', slide['القسم_المستهدف']);
                    window.location.href = 'categories.html';
                }
            };
            imgElem.style.opacity = 1;
        }, 300);
    }
    
    showSlide();
    setInterval(() => { currentSliderIndex = (currentSliderIndex + 1) % sliderData.length; showSlide(); }, 4000);
}

function applyBanners() {
    if(bannersData['رئيسي_علوي'] && document.getElementById('top-banner-img')) document.getElementById('top-banner-img').src = bannersData['رئيسي_علوي'];
    if(bannersData['مقسم_يمين'] && document.getElementById('split1-img')) document.getElementById('split1-img').src = bannersData['مقسم_يمين'];
    if(bannersData['مقسم_يسار'] && document.getElementById('split2-img')) document.getElementById('split2-img').src = bannersData['مقسم_يسار'];
    if(bannersData['رئيسي_ثابت'] && document.getElementById('main-banner-img')) document.getElementById('main-banner-img').src = bannersData['رئيسي_ثابت'];
}

function renderCatalog() {
    const container = document.getElementById('catalog-pages-container');
    if(!container) return;
    container.innerHTML = '';
    if(catalogData.length === 0) {
        container.innerHTML = '<p>لا توجد صور في مجلة العروض حالياً.</p>'; return;
    }
    catalogData.forEach(page => {
        container.innerHTML += `<div class="catalog-page"><img src="${page['صورة']}" alt="صفحة ${page['رقم_الصفحة']}" onerror="this.src='images/placeholder.png'"></div>`;
    });
}

function applyAboutSettings() {
    const phone = allSettings['رقم الواتساب'] || 'غير متوفر';
    const link = allSettings['رابط الموقع'] || 'غير متوفر';
    
    if(document.getElementById('about-main-img') && allSettings['صورة صفحة حول']) {
        document.getElementById('about-main-img').src = allSettings['صورة صفحة حول'];
    }

    if(document.getElementById('wa-number')) document.getElementById('wa-number').innerText = phone;
    if(document.getElementById('wa-link') && allSettings['رابط الواتساب']) document.getElementById('wa-link').href = allSettings['رابط الواتساب'];
    if(document.getElementById('qr-whatsapp-img') && allSettings['صورة كيو آر الواتساب']) document.getElementById('qr-whatsapp-img').src = allSettings['صورة كيو آر الواتساب'];
    
    if(document.getElementById('website-link')) {
        document.getElementById('website-link').innerText = link;
        if(link !== 'غير متوفر') document.getElementById('website-link').href = link;
    }
    if(document.getElementById('qr-website-img') && allSettings['صورة كيو آر الموقع']) document.getElementById('qr-website-img').src = allSettings['صورة كيو آر الموقع'];
    
    if(document.getElementById('fb-link') && allSettings['رابط فيسبوك']) document.getElementById('fb-link').href = allSettings['رابط فيسبوك'];
    if(document.getElementById('ig-link') && allSettings['رابط إنستجرام']) document.getElementById('ig-link').href = allSettings['رابط إنستجرام'];
    if(document.getElementById('wa-social-link') && allSettings['رابط الواتساب']) document.getElementById('wa-social-link').href = allSettings['رابط الواتساب'];
}

window.shareWebsite = function() {
    const url = allSettings['رابط الموقع'] || window.location.href;
    if (navigator.share) {
        navigator.share({ title: 'ترولي ماركت', text: 'تسوق أفضل العروض والمنتجات!', url: url }).catch(console.error);
    } else {
        alert('يمكنك نسخ هذا الرابط: ' + url);
    }
}

function getProductHTML(p) {
    let oldP = p['السعر الأساسي'];
    let newP = p['السعر بعد الخصم'];
    let hasOffer = newP && newP !== '';
    let priceHTML = ''; let badgeHTML = '';
    
    if (hasOffer && oldP > newP) {
        let discountPct = Math.round(((oldP - newP) / oldP) * 100);
        badgeHTML = `<div class="discount-badge">خصم ${discountPct}%</div>`;
        priceHTML = `<span class="new-price">${newP} ج</span> <span class="old-price">${oldP} ج</span>`;
    } else {
        priceHTML = `<span class="new-price">${oldP} ج</span>`;
    }
        
    let currentPrice = hasOffer ? newP : oldP;
    let imgSrc = p['صورة'] || 'images/placeholder.png';

    return `
    <div class="product-card">
        ${badgeHTML}
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

function renderTodayOffers() {
    const container = document.getElementById('home-sections');
    if(!container) return;
    container.innerHTML = '';
    if(allOffers.length > 0) {
        container.innerHTML = `<div class="section-header"><h3>🔥 عروض اليوم</h3></div><div class="products-grid">${allOffers.map(p => getProductHTML(p)).join('')}</div>`;
    }
}

function initCategoryFilterBar() {
    const bar = document.getElementById('category-filter-bar');
    if(!bar) return;
    bar.innerHTML = `<button class="cat-filter-btn" id="btn-cat-الكل" onclick="filterCategory('الكل')">الكل</button>`;
    CATEGORIES.forEach(cat => {
        bar.innerHTML += `<button class="cat-filter-btn" id="btn-cat-${cat}" onclick="filterCategory('${cat}')">${cat}</button>`;
    });
}

window.filterCategory = function(selectedCat) {
    document.querySelectorAll('.cat-filter-btn').forEach(b => b.classList.remove('active'));
    let btn = document.getElementById('btn-cat-' + selectedCat);
    if(btn) btn.classList.add('active');
    else if(document.getElementById('btn-cat-الكل')) document.getElementById('btn-cat-الكل').classList.add('active');

    const container = document.getElementById('categories-sections');
    if(!container) return;
    container.innerHTML = '';

    if (selectedCat === 'الكل') {
        CATEGORIES.forEach(cat => {
            const catProducts = allProducts.filter(p => p['القسم'] === cat);
            if(catProducts.length > 0) {
                container.innerHTML += `<div class="category-section"><div class="section-header"><h3>${cat}</h3></div><div class="products-grid">${catProducts.map(p => getProductHTML(p)).join('')}</div></div>`;
            }
        });
    } else {
        const catProducts = allProducts.filter(p => p['القسم'] === selectedCat);
        container.innerHTML = `<div class="category-section"><div class="section-header"><h3>${selectedCat}</h3></div><div class="products-grid">${catProducts.map(p => getProductHTML(p)).join('')}</div></div>`;
    }
}

/* Search Logic */
window.toggleSearch = function() {
    const modal = document.getElementById('search-modal');
    modal.classList.toggle('open');
    if(modal.classList.contains('open')) {
        document.getElementById('search-input').focus();
    }
}
window.performSearch = function(e) {
    const query = document.getElementById('search-input').value.trim().toLowerCase();
    const resultsContainer = document.getElementById('search-results');
    if(query.length === 0) { resultsContainer.innerHTML = ''; return; }
    
    const results = allProducts.filter(p => p['الاسم'].toLowerCase().includes(query));
    
    if(results.length > 0) {
        resultsContainer.innerHTML = `<div class="products-grid" style="grid-template-columns: repeat(auto-fill, minmax(180px, 1fr)); gap: 15px;">${results.map(p => getProductHTML(p)).join('')}</div>`;
    } else {
        let phone = allSettings['رقم الواتساب'] || '201063883209';
        let waMsg = encodeURIComponent(`طلب خاص: أحتاج إلى شراء (${query})`);
        resultsContainer.innerHTML = `
            <div style="text-align:center; padding: 40px 20px;">
                <i class="fas fa-box-open" style="font-size: 4rem; color: #ccc; margin-bottom: 20px;"></i>
                <h3 style="margin-bottom: 10px; color: var(--primary);">عذراً، المنتج غير متوفر حالياً</h3>
                <p style="font-size: 1.1rem; margin-bottom: 25px; color: var(--text-color);">يمكنك طلبه كطلب خاص وسنقوم بتوفيره لك بأسرع وقت.</p>
                <a href="https://wa.me/${phone}?text=${waMsg}" target="_blank" class="btn-primary" style="display:inline-block; width:auto; padding: 12px 30px; font-size: 1.1rem;"><i class="fab fa-whatsapp"></i> طلب خاص عبر الواتساب</a>
            </div>
        `;
    }
}

window.changeQty = function(id, delta) {
    const input = document.getElementById(`qty-${id}`);
    if(input) { let val = parseInt(input.value) + delta; if(val < 1) val = 1; if(val > 25) val = 25; input.value = val; }
}
let cart = JSON.parse(localStorage.getItem('cart')) || [];
window.toggleCart = function() { document.getElementById('cart-sidebar').classList.toggle('open'); }
window.addToCart = function(id, name, price) {
    const input = document.getElementById(`qty-${id}`);
    const qtyToAdd = input ? (parseInt(input.value) || 1) : 1;
    const existing = cart.find(i => i.name === name);
    if (existing) existing.qty += qtyToAdd; else cart.push({ name, price, qty: qtyToAdd });
    if(input) input.value = 1;
    saveCart(); updateCartUI();
    const toast = document.getElementById('toast-msg');
    if(toast) { toast.classList.add('show'); setTimeout(() => toast.classList.remove('show'), 2000); }
}
window.removeFromCart = function(name) { cart = cart.filter(i => i.name !== name); saveCart(); updateCartUI(); }
window.clearCart = function() { if(confirm('متأكد من تفريغ السلة؟')) { cart = []; saveCart(); updateCartUI(); } }
function saveCart() { localStorage.setItem('cart', JSON.stringify(cart)); }
function updateCartUI() {
    const count = document.getElementById('cart-count');
    if(count) count.innerText = cart.reduce((sum, i) => sum + i.qty, 0);
    const div = document.getElementById('cart-items');
    if(!div) return;
    div.innerHTML = ''; let total = 0;
    if(cart.length === 0) { div.innerHTML = '<p style="text-align:center;color:#999;margin-top:20px;">السلة فارغة</p>'; } 
    else {
        cart.forEach(i => {
            total += i.price * i.qty;
            div.innerHTML += `<div class="cart-item">
                <div><strong>${i.name}</strong><br><small dir="ltr" style="display:inline-block;">${i.qty} x ${i.price} ج</small></div>
                <div style="text-align:left;"><strong>${i.price * i.qty} ج</strong><br><button onclick="removeFromCart('${i.name}')" style="color:red;border:none;background:none;cursor:pointer;"><i class="fas fa-trash"></i></button></div>
            </div>`;
        });
    }
    const totalSpan = document.getElementById('total-price');
    if(totalSpan) totalSpan.innerText = total;
    let msg = "طلب جديد من ترولي ماركت:%0a%0a";
    cart.forEach(i => { msg += `- ${i.name} (الكمية: ${i.qty})%0a`; });
    msg += `%0aالإجمالي: ${total} جنيه`;
    const btn = document.getElementById('checkout-btn');
    if(btn) btn.href = `https://wa.me/${allSettings['رقم الواتساب'] || '201063883209'}?text=${msg}`;
}

document.addEventListener('DOMContentLoaded', () => { initTheme(); loadExcelData(); });
