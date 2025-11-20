// ===== グローバル変数 =====
const API_BASE_URL = 'https://script.google.com/macros/s/AKfycbxxpRqydeJCcPMCPm3apwmlkt6v--GDCKPeTEHW_SNWxdAoxBY8aumStZo8hXSMrpuB/exec';

let products = [];
let cart = [];
let orders = [];
let members = [];

let currentEditingProductId = null;
let currentEditingMemberId = null;

// ===== 初期化 =====
document.addEventListener('DOMContentLoaded', () => {
    initializeApp();
});

async function initializeApp() {
    loadCartFromStorage();
    setupEventListeners();

    await loadProducts();
    await loadMembers();
    await loadOrders();

    renderProducts();
    updateCartCount();
    renderOrders();
    renderMembers();
    navigateToPage('products');
}

// ===== イベントリスナー設定 =====
function setupEventListeners() {
    // ナビゲーション
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const page = e.currentTarget.dataset.page;
            navigateToPage(page);
        });
    });

    // 検索・カテゴリフィルタ
    const searchInput = document.getElementById('searchInput');
    if (searchInput) {
        searchInput.addEventListener('input', filterProducts);
    }
    const categoryFilter = document.getElementById('categoryFilter');
    if (categoryFilter) {
        categoryFilter.addEventListener('change', filterProducts);
    }

    // 商品モーダル関連
    const addProductBtn = document.getElementById('addProductBtn');
    if (addProductBtn) {
        addProductBtn.addEventListener('click', () => openProductModal());
    }

    const productForm = document.getElementById('productForm');
    if (productForm) {
        productForm.addEventListener('submit', handleProductSubmit);
    }

    const closeModalBtn = document.getElementById('closeModal');
    if (closeModalBtn) {
        closeModalBtn.addEventListener('click', closeProductModal);
    }

    const cancelProductBtn = document.getElementById('cancelBtn');
    if (cancelProductBtn) {
        cancelProductBtn.addEventListener('click', closeProductModal);
    }

    const productModal = document.getElementById('productModal');
    if (productModal) {
        productModal.addEventListener('click', (e) => {
            if (e.target.id === 'productModal') closeProductModal();
        });
    }

    // メンバーモーダル関連
    const addMemberBtn = document.getElementById('addMemberBtn');
    if (addMemberBtn) {
        addMemberBtn.addEventListener('click', () => openMemberModal());
    }

    const memberForm = document.getElementById('memberForm');
    if (memberForm) {
        memberForm.addEventListener('submit', handleMemberSubmit);
    }

    const closeMemberModalBtn = document.getElementById('closeMemberModal');
    if (closeMemberModalBtn) {
        closeMemberModalBtn.addEventListener('click', closeMemberModal);
    }

    const cancelMemberBtn = document.getElementById('cancelMemberBtn');
    if (cancelMemberBtn) {
        cancelMemberBtn.addEventListener('click', closeMemberModal);
    }

    const memberModal = document.getElementById('memberModal');
    if (memberModal) {
        memberModal.addEventListener('click', (e) => {
            if (e.target.id === 'memberModal') closeMemberModal();
        });
    }

    // カート → 発注
    const createOrderBtn = document.getElementById('createOrderBtn');
    if (createOrderBtn) {
        createOrderBtn.addEventListener('click', createOrder);
    }

    // 発注書プレビュー関連
    const closePreviewBtn = document.getElementById('closePreviewBtn');
    if (closePreviewBtn) {
        closePreviewBtn.addEventListener('click', closePreviewModal);
    }

    const printOrderBtn = document.getElementById('printOrderBtn');
    if (printOrderBtn) {
        printOrderBtn.addEventListener('click', printOrder);
    }

    const orderPreviewModal = document.getElementById('orderPreviewModal');
    if (orderPreviewModal) {
        orderPreviewModal.addEventListener('click', (e) => {
            if (e.target.id === 'orderPreviewModal') closePreviewModal();
        });
    }
}

// ===== ページ遷移 =====
function navigateToPage(pageName) {
    document.querySelectorAll('.page').forEach(page => {
        page.classList.remove('active');
    });
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.classList.remove('active');
    });

    const pageEl = document.getElementById(`${pageName}Page`);
    if (pageEl) pageEl.classList.add('active');
    const navBtn = document.querySelector(`[data-page="${pageName}"]`);
    if (navBtn) navBtn.classList.add('active');

    if (pageName === 'cart') {
        renderCart();
    } else if (pageName === 'orders') {
        renderOrders();
    } else if (pageName === 'admin') {
        renderAdminProducts();
    } else if (pageName === 'members') {
        renderMembers();
    }
}

// ===== 商品関連 =====
async function loadProducts() {
    try {
        const response = await fetch(`${API_BASE_URL}?action=getProducts`);
        const result = await response.json();
        if (!result.success && result.error) {
            console.error('getProducts error:', result.error);
        }
        products = result.data || [];
        updateCategoryFilter();
    } catch (error) {
        console.error('商品の読み込みに失敗しました:', error);
        showNotification('商品の読み込みに失敗しました', 'error');
        products = [];
    }
}

function updateCategoryFilter() {
    const categoryFilter = document.getElementById('categoryFilter');
    if (!categoryFilter) return;

    const categories = [...new Set(products.map(p => p.category).filter(Boolean))];
    categoryFilter.innerHTML = '<option value="">すべてのカテゴリ</option>';

    categories.forEach(category => {
        const option = document.createElement('option');
        option.value = category;
        option.textContent = category;
        categoryFilter.appendChild(option);
    });
}

function renderProducts() {
    const container = document.getElementById('productsList');
    if (!container) return;

    if (products.length === 0) {
        container.innerHTML = `
            <div style="text-align: center; padding: 60px 20px;">
                <i class="fas fa-box-open" style="font-size: 64px; color: var(--border-color); margin-bottom: 16px;"></i>
                <p style="color: var(--text-secondary); font-size: 16px;">登録されている商品がありません</p>
                <button class="btn btn-primary" style="margin-top: 16px;" onclick="navigateToPage('admin')">
                    <i class="fas fa-plus"></i> 最初の商品を登録
                </button>
            </div>
        `;
        return;
    }

    container.innerHTML = products.map(product => `
        <div class="product-card">
            <div class="product-image">
                ${product.image_url ?
                    `<img src="${product.image_url}" alt="${product.name}">` :
                    '<i class="fas fa-box"></i>'
                }
            </div>
            <div class="product-info">
                <div class="product-category">${product.category || '未分類'}</div>
                ${product.short_name ? `<div class="product-shortname-badge">${product.short_name}</div>` : ''}
                <div class="product-name">${product.name}</div>
                <div class="product-manufacturer">${product.manufacturer}</div>
                <div class="product-catalog">商品コード: ${product.catalog_number}</div>
                <div class="product-capacity">内容量: ${product.capacity}</div>
                <div class="product-usage">使用場所: ${product.usage_place}</div>
                <div class="product-actions" style="margin-top: 12px;">
                    <button class="btn btn-primary btn-full" onclick="addToCart('${product.id}')">
                        <i class="fas fa-cart-plus"></i> カートに追加
                    </button>
                </div>
            </div>
        </div>
    `).join('');
}

function filterProducts() {
    const container = document.getElementById('productsList');
    if (!container) return;

    const searchTerm = (document.getElementById('searchInput').value || '').toLowerCase();
    const category = document.getElementById('categoryFilter').value;

    const filtered = products.filter(product => {
        const name = (product.name || '').toLowerCase();
        const shortName = (product.short_name || '').toLowerCase();
        const manufacturer = (product.manufacturer || '').toLowerCase();
        const catalog = (product.catalog_number || '').toLowerCase();

        const matchesSearch =
            name.includes(searchTerm) ||
            shortName.includes(searchTerm) ||
            manufacturer.includes(searchTerm) ||
            catalog.includes(searchTerm);

        const matchesCategory = !category || product.category === category;

        return matchesSearch && matchesCategory;
    });

    if (filtered.length === 0) {
        container.innerHTML = `
            <div style="text-align: center; padding: 60px 20px;">
                <i class="fas fa-search" style="font-size: 48px; color: var(--border-color); margin-bottom: 16px;"></i>
                <p style="color: var(--text-secondary); font-size: 16px;">条件に一致する商品がありません</p>
            </div>
        `;
        return;
    }

    container.innerHTML = filtered.map(product => `
        <div class="product-card">
            <div class="product-image">
                ${product.image_url ?
                    `<img src="${product.image_url}" alt="${product.name}">` :
                    '<i class="fas fa-box"></i>'
                }
            </div>
            <div class="product-info">
                <div class="product-category">${product.category || '未分類'}</div>
                ${product.short_name ? `<div class="product-shortname-badge">${product.short_name}</div>` : ''}
                <div class="product-name">${product.name}</div>
                <div class="product-manufacturer">${product.manufacturer}</div>
                <div class="product-catalog">商品コード: ${product.catalog_number}</div>
                <div class="product-capacity">内容量: ${product.capacity}</div>
                <div class="product-usage">使用場所: ${product.usage_place}</div>
                <div class="product-actions" style="margin-top: 12px;">
                    <button class="btn btn-primary btn-full" onclick="addToCart('${product.id}')">
                        <i class="fas fa-cart-plus"></i> カートに追加
                    </button>
                </div>
            </div>
        </div>
    `).join('');
}

// ===== カート機能 =====
function loadCartFromStorage() {
    try {
        const storedCart = localStorage.getItem('labOrderCart');
        cart = storedCart ? JSON.parse(storedCart) : [];
    } catch (error) {
        console.error('カートの読み込みに失敗しました:', error);
        cart = [];
    }
}

function saveCartToStorage() {
    localStorage.setItem('labOrderCart', JSON.stringify(cart));
}

function addToCart(productId) {
    const product = products.find(p => p.id === productId);
    if (!product) return;

    const existingItem = cart.find(item => item.productId === productId);

    if (existingItem) {
        existingItem.quantity++;
    } else {
        cart.push({
            productId: productId,
            quantity: 1,
            checked: true
        });
    }

    saveCartToStorage();
    updateCartCount();
    showNotification('カートに追加しました', 'success');
}

function removeFromCart(productId) {
    cart = cart.filter(item => item.productId !== productId);
    saveCartToStorage();
    updateCartCount();
    renderCart();
}

function updateCartCount() {
    const count = cart.reduce((sum, item) => sum + item.quantity, 0);
    document.getElementById('cartCount').textContent = count;
}

function toggleCartItem(productId) {
    const item = cart.find(item => item.productId === productId);
    if (!item) return;

    item.checked = !item.checked;
    saveCartToStorage();
    renderCart();
}

function updateCartQuantity(productId, newQuantity) {
    const item = cart.find(item => item.productId === productId);
    if (!item) return;

    if (newQuantity < 1 || isNaN(newQuantity)) {
        removeFromCart(productId);
    } else {
        item.quantity = newQuantity;
        saveCartToStorage();
        renderCart();
        updateCartCount();
    }
}

function renderCart() {
    const container = document.getElementById('cartItems');
    const summary = document.getElementById('cartSummary');
    if (!container || !summary) return;

    if (cart.length === 0) {
        container.innerHTML = `
            <div style="text-align: center; padding: 60px 20px;">
                <i class="fas fa-shopping-cart" style="font-size: 64px; color: var(--border-color); margin-bottom: 16px;"></i>
                <p style="color: var(--text-secondary); font-size: 16px;">カートに商品がありません</p>
                <button class="btn btn-primary" style="margin-top: 16px;" onclick="navigateToPage('products')">
                    <i class="fas fa-boxes"></i> 商品を探す
                </button>
            </div>
        `;
        summary.style.display = 'none';
        return;
    }

    let totalCheckedItems = 0;
    let totalQuantity = 0;

    container.innerHTML = cart.map(item => {
        const product = products.find(p => p.id === item.productId);
        if (!product) return '';

        if (item.checked) {
            totalCheckedItems += 1;
            totalQuantity += item.quantity;
        }

        return `
            <div class="cart-item ${item.checked ? '' : 'unchecked'}">
                <div class="cart-item-check">
                    <input type="checkbox" ${item.checked ? 'checked' : ''} onchange="toggleCartItem('${item.productId}')">
                </div>
                <div class="cart-item-details">
                    <div class="cart-item-name">
                        ${product.short_name ? `[${product.short_name}] ` : ''}${product.name}
                    </div>
                    <div class="cart-item-manufacturer">
                        ${product.manufacturer} | ${product.catalog_number}
                    </div>
                    <div class="cart-item-catalog">
                        内容量: ${product.capacity} / 使用場所: ${product.usage_place}
                    </div>
                </div>
                <div class="cart-item-quantity">
                    <button class="quantity-btn" onclick="updateCartQuantity('${item.productId}', ${item.quantity - 1})">
                        <i class="fas fa-minus"></i>
                    </button>
                    <input type="number" class="quantity-input" value="${item.quantity}" min="1"
                           onchange="updateCartQuantity('${item.productId}', parseInt(this.value))">
                    <button class="quantity-btn" onclick="updateCartQuantity('${item.productId}', ${item.quantity + 1})">
                        <i class="fas fa-plus"></i>
                    </button>
                </div>
                <button class="cart-item-remove" onclick="removeFromCart('${item.productId}')">
                    <i class="fas fa-trash"></i>
                </button>
            </div>
        `;
    }).join('');

    const checkedItems = cart.filter(item => item.checked);
    if (checkedItems.length > 0) {
        document.getElementById('subtotal').textContent = `${totalCheckedItems}件`;
        document.getElementById('total').textContent = `${totalQuantity} 個`;
        summary.style.display = 'block';
    } else {
        summary.style.display = 'none';
    }
}

// ===== メンバー関連 =====
async function loadMembers() {
    try {
        const response = await fetch(`${API_BASE_URL}?action=getMembers`);
        const result = await response.json();
        if (!result.success && result.error) {
            console.error('getMembers error:', result.error);
        }
        members = result.data || [];
        updateMemberSelect();
    } catch (error) {
        console.error('メンバーの読み込みに失敗しました:', error);
        members = [];
        updateMemberSelect();
    }
}

function updateMemberSelect() {
    const select = document.getElementById('orderMemberSelect');
    if (!select) return;

    select.innerHTML = '<option value="">選択してください</option>';

    members.forEach(member => {
        const option = document.createElement('option');
        option.value = member.id;
        option.textContent = member.name;
        select.appendChild(option);
    });
}

function renderMembers() {
    const container = document.getElementById('membersList');
    if (!container) return;

    if (members.length === 0) {
        container.innerHTML = `
            <div style="text-align: center; padding: 40px 20px;">
                <p style="color: var(--text-secondary);">登録されているメンバーがありません。</p>
            </div>
        `;
        return;
    }

    container.innerHTML = `
        <div class="admin-table">
            <table>
                <thead>
                    <tr>
                        <th>名前</th>
                        <th>メールアドレス</th>
                        <th>操作</th>
                    </tr>
                </thead>
                <tbody>
                    ${members.map(member => `
                        <tr>
                            <td>${member.name}</td>
                            <td>${member.email || ''}</td>
                            <td>
                                <button class="btn btn-secondary btn-sm" onclick="editMember('${member.id}')">
                                    <i class="fas fa-edit"></i> 編集
                                </button>
                                <button class="btn btn-danger btn-sm" onclick="deleteMember('${member.id}')">
                                    <i class="fas fa-trash"></i> 削除
                                </button>
                            </td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;
}

function openMemberModal(member = null) {
    currentEditingMemberId = member ? member.id : null;

    const modal = document.getElementById('memberModal');
    const title = document.getElementById('memberModalTitle');
    const form = document.getElementById('memberForm');

    title.textContent = member ? 'メンバーを編集' : 'メンバーを追加';

    if (member) {
        document.getElementById('memberName').value = member.name || '';
        document.getElementById('memberEmail').value = member.email || '';
        document.getElementById('memberId').value = member.id;
    } else {
        form.reset();
        document.getElementById('memberId').value = '';
    }

    modal.classList.add('active');
}

function closeMemberModal() {
    document.getElementById('memberModal').classList.remove('active');
}

async function handleMemberSubmit(e) {
    e.preventDefault();

    const memberData = {
        action: 'saveMember',
        id: currentEditingMemberId,
        name: document.getElementById('memberName').value,
        email: document.getElementById('memberEmail').value
    };

    const response = await fetch(API_BASE_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json;charset=utf-8' },
        body: JSON.stringify(memberData)
    });

    const result = await response.json();

    if (result.success) {
        showNotification('メンバーを保存しました');
        closeMemberModal();
        loadMembers();
    } else {
        showNotification('メンバーの保存に失敗しました', 'error');
    }
}

async function editMember(memberId) {
    const member = members.find(m => m.id === memberId);
    if (member) {
        openMemberModal(member);
    }
}

async function deleteMember(memberId) {
    if (!confirm('このメンバーを削除してもよろしいですか?')) {
        return;
    }

    const body = {
        action: 'deleteMember',
        id: memberId
    };

    try {
        const response = await fetch(API_BASE_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json;charset=utf-8' },
            body: JSON.stringify(body)
        });
        const result = await response.json();

        if (result && result.success) {
            showNotification('メンバーを削除しました', 'success');
            await loadMembers();
            renderMembers();
        } else {
            console.error('メンバーの削除APIエラー:', result);
            showNotification('メンバーの削除に失敗しました', 'error');
        }
    } catch (error) {
        console.error('メンバーの削除に失敗しました:', error);
        showNotification('メンバーの削除に失敗しました', 'error');
    }
}

// ===== 発注書作成 =====
async function createOrder() {
    const checkedItems = cart.filter(item => item.checked);
    if (checkedItems.length === 0) {
        showNotification('発注する商品を選択してください', 'error');
        return;
    }

    const memberSelect = document.getElementById('orderMemberSelect');
    const memberId = memberSelect ? memberSelect.value : '';
    if (!memberId) {
        showNotification('注文者を選択してください', 'error');
        return;
    }

    const member = members.find(m => m.id === memberId);
    const memberName = member ? member.name : '';
    const memberEmail = member ? member.email : '';

    const orderItems = checkedItems.map(item => {
        const product = products.find(p => p.id === item.productId);
        return {
            productId: product.id,
            short_name: product.short_name,
            name: product.name,
            manufacturer: product.manufacturer,
            catalog_number: product.catalog_number,
            capacity: product.capacity,
            usage_place: product.usage_place,
            quantity: item.quantity
        };
    });

    const notes = document.getElementById('orderNotes').value;

    const orderData = {
        action: 'createOrder',
        order_number: generateOrderNumber(),
        order_date: new Date().toISOString().split('T')[0],
        member_name: memberName,
        member_email: memberEmail,
        items: JSON.stringify(orderItems),
        notes: notes || '',
        status: 'draft'
    };

    try {
        const response = await fetch(API_BASE_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json;charset=utf-8' },
            body: JSON.stringify(orderData)
        });
        const result = await response.json();

        if (result && result.success) {
            const newOrder = result.data;

            cart = cart.filter(item => !item.checked);
            saveCartToStorage();
            updateCartCount();

            showOrderPreview(newOrder);
            showNotification('発注書を作成しました', 'success');

            await loadOrders();
            renderOrders();
        } else {
            console.error('発注書の作成APIエラー:', result);
            showNotification('発注書の作成に失敗しました', 'error');
        }
    } catch (error) {
        console.error('発注書の作成に失敗しました:', error);
        showNotification('発注書の作成に失敗しました', 'error');
    }
}

function generateOrderNumber() {
    const date = new Date();
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    const rand = String(Math.floor(Math.random() * 10000)).padStart(4, '0');
    return `ORD-${y}${m}${d}-${rand}`;
}

// ===== 発注履歴 =====
async function loadOrders() {
    try {
        const response = await fetch(`${API_BASE_URL}?action=getOrders`);
        const result = await response.json();
        if (!result.success && result.error) {
            console.error('getOrders error:', result.error);
        }
        orders = result.data || [];
    } catch (error) {
        console.error('発注履歴の読み込みに失敗しました:', error);
        showNotification('発注履歴の読み込みに失敗しました', 'error');
        orders = [];
    }
}

function renderOrders() {
    const container = document.getElementById('ordersList');
    if (!container) return;

    if (orders.length === 0) {
        container.innerHTML = `
            <div style="text-align: center; padding: 60px 20px;">
                <i class="fas fa-file-invoice" style="font-size: 64px; color: var(--border-color); margin-bottom: 16px;"></i>
                <p style="color: var(--text-secondary); font-size: 16px;">発注履歴がありません</p>
            </div>
        `;
        return;
    }

    const statusLabels = {
        draft: '下書き',
        submitted: '発注済み',
        completed: '完了'
    };

    container.innerHTML = orders.map(order => {
        let items = [];
        try {
            items = JSON.parse(order.items || '[]');
        } catch (e) {
            items = [];
        }

        return `
            <div class="order-card">
                <div class="order-header">
                    <div class="order-info">
                        <h3>${order.order_number}</h3>
                        <div class="order-date">
                            発注日: ${order.order_date || ''}<br>
                            注文者: ${order.member_name || ''}
                        </div>
                    </div>
                    <div class="order-status ${order.status}">
                        ${statusLabels[order.status] || order.status || '下書き'}
                    </div>
                </div>
                <div class="order-items">
                    ${items.map(item => `
                        <div class="order-item">
                            <div>
                                <div class="order-item-name">
                                    ${item.short_name ? `[${item.short_name}] ` : ''}${item.name}
                                </div>
                                <div class="order-item-details">
                                    ${item.manufacturer} | ${item.catalog_number}<br>
                                    内容量: ${item.capacity} / 使用場所: ${item.usage_place} / 数量: ${item.quantity}
                                </div>
                            </div>
                        </div>
                    `).join('')}
                </div>
                ${order.notes ? `
                    <div class="order-notes">
                        <strong>備考:</strong> ${order.notes}
                    </div>
                ` : ''}
                <div class="order-actions">
                    <button class="btn btn-primary" onclick="showOrderPreview(${JSON.stringify(order).replace(/"/g, '&quot;')})">
                        <i class="fas fa-eye"></i> プレビュー
                    </button>
                    <button class="btn btn-secondary" onclick="printOrderDirect(${JSON.stringify(order).replace(/"/g, '&quot;')})">
                        <i class="fas fa-print"></i> 印刷
                    </button>
                    <button class="btn btn-danger" onclick="deleteOrder('${order.id}')">
                        <i class="fas fa-trash"></i> 削除
                    </button>
                </div>
            </div>
        `;
    }).join('');
}

// ===== 発注書プレビュー =====
function showOrderPreview(order) {
    const items = (() => {
        try {
            return JSON.parse(order.items || '[]');
        } catch (e) {
            return [];
        }
    })();

    const previewHtml = `
        <div class="preview-header">
            <div>
                <h2>発注書</h2>
                <div>発注番号: ${order.order_number}</div>
                <div>発注日: ${order.order_date}</div>
                <div>注文者: ${order.member_name || ''}</div>
            </div>
        </div>
        <table class="preview-table">
            <thead>
                <tr>
                    <th style="width: 5%;">No.</th>
                    <th style="width: 30%;">商品名</th>
                    <th style="width: 15%;">メーカー</th>
                    <th style="width: 15%;">商品コード</th>
                    <th style="width: 15%;">内容量</th>
                    <th style="width: 10%;">使用場所</th>
                    <th style="width: 10%; text-align: center;">数量</th>
                </tr>
            </thead>
            <tbody>
                ${items.map((item, idx) => `
                    <tr>
                        <td style="text-align: center;">${idx + 1}</td>
                        <td>${item.name}</td>
                        <td>${item.manufacturer}</td>
                        <td>${item.catalog_number}</td>
                        <td>${item.capacity}</td>
                        <td>${item.usage_place}</td>
                        <td style="text-align: center;">${item.quantity}</td>
                    </tr>
                `).join('')}
            </tbody>
        </table>
        ${order.notes ? `
            <div class="preview-notes">
                <strong>備考:</strong><br>
                ${order.notes}
            </div>
        ` : ''}
    `;

    document.getElementById('orderPreview').innerHTML = previewHtml;
    document.getElementById('orderPreviewModal').classList.add('active');
}

function closePreviewModal() {
    document.getElementById('orderPreviewModal').classList.remove('active');
}

function printOrder() {
    const previewContent = document.getElementById('orderPreview').innerHTML;
    const originalContent = document.body.innerHTML;

    document.body.innerHTML = previewContent;
    window.print();
    document.body.innerHTML = originalContent;
    window.location.reload();
}

function printOrderDirect(order) {
    showOrderPreview(order);
    setTimeout(printOrder, 300);
}

// ===== 発注書削除（案内のみ） =====
async function deleteOrder(orderId) {
    if (!confirm('この発注書を削除してもよろしいですか?')) {
        return;
    }
    showNotification('発注書の削除は現在スプレッドシートからのみ行ってください', 'error');
}

// ===== 商品管理（管理画面） =====
function renderAdminProducts() {
    const container = document.getElementById('adminProductsList');
    if (!container) return;

    if (products.length === 0) {
        container.innerHTML = `
            <div style="text-align: center; padding: 40px 20px;">
                <p style="color: var(--text-secondary);">登録されている商品がありません。</p>
            </div>
        `;
        return;
    }

    container.innerHTML = `
        <div class="admin-table">
            <table>
                <thead>
                    <tr>
                        <th>略称</th>
                        <th>商品名</th>
                        <th>メーカー</th>
                        <th>商品コード</th>
                        <th>内容量</th>
                        <th>使用場所</th>
                        <th>カテゴリ</th>
                        <th>操作</th>
                    </tr>
                </thead>
                <tbody>
                    ${products.map(product => `
                        <tr>
                            <td>${product.short_name || ''}</td>
                            <td>${product.name}</td>
                            <td>${product.manufacturer}</td>
                            <td>${product.catalog_number}</td>
                            <td>${product.capacity}</td>
                            <td>${product.usage_place}</td>
                            <td>${product.category || ''}</td>
                            <td>
                                <button class="btn btn-secondary btn-sm" onclick="editProduct('${product.id}')">
                                    <i class="fas fa-edit"></i> 編集
                                </button>
                                <button class="btn btn-danger btn-sm" onclick="deleteProduct('${product.id}')">
                                    <i class="fas fa-trash"></i> 削除
                                </button>
                            </td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;
}

function openProductModal(product = null) {
    currentEditingProductId = product ? product.id : null;

    const modal = document.getElementById('productModal');
    const title = document.getElementById('modalTitle');
    const form = document.getElementById('productForm');

    title.textContent = product ? '商品を編集' : '商品を追加';

    if (product) {
        document.getElementById('productImage').value = product.image_url || '';
        document.getElementById('productShortName').value = product.short_name || '';
        document.getElementById('productName').value = product.name || '';
        document.getElementById('productManufacturer').value = product.manufacturer || '';
        document.getElementById('productCatalog').value = product.catalog_number || '';
        document.getElementById('productCapacity').value = product.capacity || '';
        document.getElementById('productUsagePlace').value = product.usage_place || '';
        document.getElementById('productCategory').value = product.category || '';
        document.getElementById('productId').value = product.id;
    } else {
        form.reset();
        document.getElementById('productId').value = '';
    }

    modal.classList.add('active');
}

function closeProductModal() {
    document.getElementById('productModal').classList.remove('active');
}

async function handleProductSubmit(e) {
    e.preventDefault();

    const productData = {
        action: 'saveProduct',
        id: currentEditingProductId,
        image_url: document.getElementById('productImage').value,
        short_name: document.getElementById('productShortName').value,
        name: document.getElementById('productName').value,
        manufacturer: document.getElementById('productManufacturer').value,
        catalog_number: document.getElementById('productCatalog').value,
        capacity: document.getElementById('productCapacity').value,
        usage_place: document.getElementById('productUsagePlace').value,
        category: document.getElementById('productCategory').value
    };

    const response = await fetch(API_BASE_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json;charset=utf-8' },
        body: JSON.stringify(productData)
    });

    const result = await response.json();

    if (result.success) {
        showNotification('商品を保存しました');
        closeProductModal();
        loadProducts();
    } else {
        showNotification('商品の保存に失敗しました', 'error');
    }
}

async function editProduct(productId) {
    const product = products.find(p => p.id === productId);
    if (product) {
        openProductModal(product);
    }
}

async function deleteProduct(productId) {
    if (!confirm('この商品を削除してもよろしいですか?')) {
        return;
    }

    const body = {
        action: 'deleteProduct',
        id: productId
    };

    try {
        const response = await fetch(API_BASE_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json;charset=utf-8' },
            body: JSON.stringify(body)
        });
        const result = await response.json();

        if (result && result.success) {
            showNotification('商品を削除しました', 'success');
            await loadProducts();
            renderProducts();
            renderAdminProducts();
        } else {
            console.error('商品の削除APIエラー:', result);
            showNotification('商品の削除に失敗しました', 'error');
        }
    } catch (error) {
        console.error('商品の削除に失敗しました:', error);
        showNotification('商品の削除に失敗しました', 'error');
    }
}

// ===== 通知 =====
function showNotification(message, type = 'success') {
    const notification = document.getElementById('notification');
    notification.textContent = message;
    notification.className = `notification ${type} show`;

    setTimeout(() => {
        notification.classList.remove('show');
    }, 3000);
}
