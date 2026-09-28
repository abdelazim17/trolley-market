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
        
        // فلترة (متوفر = نعم)
        allProducts = jsonData.filter(item => item['متوفر'] === 'نعم');
        allProducts.forEach((p, idx) => p.id = 'prod_' + idx);
        
        document.getElementById('loading').style.display = 'none';
        
        if(allProducts.length === 0) {
            document.getElementById('error-box').innerHTML = "ملف الإكسيل موجود ولكنه فارغ أو لا يوجد منتجات متوفرة.";
            document.getElementById('error-box').style.display = 'block';
            return;
        }

        initCategoriesFilter();
        renderHomeSections();

    } catch (error) {
        console.error("خطأ في التحميل:", error);
        document.getElementById('loading').style.display = 'none';
        document.getElementById('error-box').style.display = 'block';
    }
}

function initCategoriesFilter() {
    const container = document.getElementById('categories-text-container');
    
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
        
    // fallback للصورة لو مش موجودة في الإكسيل
    let imgSrc = p['الصورة'] ? p['الصورة'] : 'images/placeholder.png';

    return `
    <div class="product-card">
        <img src="${imgSrc}" class="product-img" onerror="this.src='images/placeholder.png'">
        <div class="product-title">${p['الاسم']}</div>
        <div class="price-container">${priceHTML} / ${p['الوحدة']}</div>
        
        <div class="qty-controls">
            <button class="qty-btn" onclick="changeQty('${p.id}', 1)">+</button>
            <input type="number" id="qty-${p.id}" class="qty-input" value="1" readonly>
            <button class="qty-btn" onclick="changeQty('${p.id}', -1)">-</button>
        </div>
        
        <button class="add-to-cart-btn" onclick="alert('تمت الإضافة')">إضافة للسلة</button>
    </div>
    `;
}

function renderHomeSections() {
    const container = document.getElementById('home-sections');
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

function filterHome(category) {
    const container = document.getElementById('home-sections');
    container.innerHTML = '';
    
    const catProducts = allProducts.filter(p => p['القسم'] === category);
    if(catProducts.length > 0) {
        container.innerHTML = `
            <div class="products-grid">
                ${catProducts.map(p => getProductHTML(p)).join('')}
            </div>
        `;
    }
}

window.changeQty = function(id, delta) {
    const input = document.getElementById(`qty-${id}`);
    if(input) {
        let val = parseInt(input.value) + delta;
        if(val < 1) val = 1;
        if(val > 25) val = 25;
        input.value = val;
    }
}

window.onload = loadExcelData;
