// ========== 데이터 저장/로드 (localStorage) ==========
function loadData(key) {
    return JSON.parse(localStorage.getItem(key) || '[]');
}

function saveData(key, data) {
    localStorage.setItem(key, JSON.stringify(data));
}

let products = loadData('kenvue_products');
let productions = loadData('kenvue_productions');

// ========== 탭 전환 ==========
document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
        btn.classList.add('active');
        document.getElementById(btn.dataset.tab).classList.add('active');

        if (btn.dataset.tab === 'manage') refreshProductSelect();
        if (btn.dataset.tab === 'sales') {
            refreshFilterProductSelect();
            updateSales();
        }
    });
});

// ========== 숫자 포맷 ==========
function formatNumber(n) {
    return Number(n).toLocaleString('ko-KR');
}

// ========== 배치번호 → 생산일자 ==========
function parseBatchDate(batch) {
    const parts = batch.split('-');
    if (parts.length < 1) return '';
    const datePart = parts[0];
    if (datePart.length !== 6) return '';
    const yy = datePart.substring(0, 2);
    const mm = datePart.substring(2, 4);
    const dd = datePart.substring(4, 6);
    return `20${yy}${mm}${dd}`;
}

function formatDateDisplay(dateStr) {
    if (dateStr.length !== 8) return dateStr;
    return `${dateStr.substring(0, 4)}-${dateStr.substring(4, 6)}-${dateStr.substring(6, 8)}`;
}

// ========== 마스터파일 (품목 관리) ==========
let editingIndex = -1;

function renderProducts() {
    const tbody = document.getElementById('productTableBody');
    const emptyMsg = document.getElementById('productEmptyMsg');

    if (products.length === 0) {
        tbody.innerHTML = '';
        emptyMsg.style.display = 'block';
        return;
    }

    emptyMsg.style.display = 'none';
    tbody.innerHTML = products.map((p, i) => {
        if (editingIndex === i) {
            return `
            <tr class="editing-row">
                <td><input type="text" id="editCode" value="${p.code}" class="inline-input"></td>
                <td><input type="text" id="editName" value="${p.name}" class="inline-input"></td>
                <td><input type="text" id="editSpec" value="${p.spec || ''}" class="inline-input"></td>
                <td><input type="number" id="editPrice" value="${p.price}" class="inline-input"></td>
                <td>
                    <button class="btn btn-save" onclick="saveEdit(${i})">저장</button>
                    <button class="btn btn-cancel" onclick="cancelEdit()">취소</button>
                </td>
            </tr>`;
        }
        return `
        <tr>
            <td>${p.code}</td>
            <td>${p.name}</td>
            <td>${p.spec || '-'}</td>
            <td class="text-right">${formatNumber(p.price)}</td>
            <td>
                <button class="btn btn-edit" onclick="startEdit(${i})">수정</button>
                <button class="btn btn-danger" onclick="deleteProduct(${i})">삭제</button>
            </td>
        </tr>`;
    }).join('');
}

window.startEdit = function(index) {
    editingIndex = index;
    renderProducts();
};

window.cancelEdit = function() {
    editingIndex = -1;
    renderProducts();
};

window.saveEdit = function(index) {
    const code = document.getElementById('editCode').value.trim();
    const name = document.getElementById('editName').value.trim();
    const spec = document.getElementById('editSpec').value.trim();
    const price = Number(document.getElementById('editPrice').value);

    if (!code || !name || !price) {
        alert('품목코드, 품목명, 단가는 필수입니다.');
        return;
    }

    // 코드 변경 시 중복 체크
    const duplicate = products.findIndex((p, i) => p.code === code && i !== index);
    if (duplicate >= 0) {
        alert('동일한 품목코드가 이미 존재합니다.');
        return;
    }

    products[index] = { code, name, spec, price };
    editingIndex = -1;
    saveData('kenvue_products', products);
    renderProducts();
    refreshProductSelect();
};

document.getElementById('productForm').addEventListener('submit', e => {
    e.preventDefault();
    const code = document.getElementById('productCode').value.trim();
    const name = document.getElementById('productName').value.trim();
    const spec = document.getElementById('productSpec').value.trim();
    const price = Number(document.getElementById('productPrice').value);

    const existing = products.findIndex(p => p.code === code);
    if (existing >= 0) {
        if (confirm(`품목코드 "${code}"가 이미 존재합니다. 덮어쓰시겠습니까?`)) {
            products[existing] = { code, name, spec, price };
        } else {
            return;
        }
    } else {
        products.push({ code, name, spec, price });
    }

    saveData('kenvue_products', products);
    renderProducts();
    refreshProductSelect();
    e.target.reset();
});

window.deleteProduct = function(index) {
    if (confirm('이 품목을 삭제하시겠습니까?')) {
        products.splice(index, 1);
        saveData('kenvue_products', products);
        renderProducts();
        refreshProductSelect();
    }
};

// ========== 생산실적 관리 ==========
function refreshProductSelect() {
    const select = document.getElementById('prodProductCode');
    const currentVal = select.value;
    select.innerHTML = '<option value="">선택하세요</option>' +
        products.map(p => `<option value="${p.code}">${p.code} - ${p.name}</option>`).join('');
    select.value = currentVal;
}

function getProductByCode(code) {
    return products.find(p => p.code === code);
}

// 품목코드 선택 시 존슨코드/품목명/단가 자동 표시
document.getElementById('prodProductCode').addEventListener('change', function() {
    const product = getProductByCode(this.value);
    if (product) {
        document.getElementById('prodJohnsonCode').value = product.spec || '';
        document.getElementById('prodProductName').value = product.name;
        document.getElementById('prodUnitPrice').value = formatNumber(product.price) + ' 원';
    } else {
        document.getElementById('prodJohnsonCode').value = '';
        document.getElementById('prodProductName').value = '';
        document.getElementById('prodUnitPrice').value = '';
    }
    updatePreviewSales();
});

// 배치번호 입력 시 생산일자 자동 추출
document.getElementById('batchNumber').addEventListener('input', e => {
    const batch = e.target.value.trim();
    const dateStr = parseBatchDate(batch);
    document.getElementById('prodDate').value = dateStr || '';
    updatePreviewSales();
});

document.getElementById('prodQuantity').addEventListener('input', updatePreviewSales);

function updatePreviewSales() {
    const code = document.getElementById('prodProductCode').value;
    const qty = Number(document.getElementById('prodQuantity').value);
    const product = getProductByCode(code);

    if (product && qty > 0) {
        document.getElementById('prodSales').value = formatNumber(product.price * qty) + ' 원';
    } else {
        document.getElementById('prodSales').value = '';
    }
}

function renderProductions() {
    const tbody = document.getElementById('productionTableBody');
    const emptyMsg = document.getElementById('productionEmptyMsg');

    if (productions.length === 0) {
        tbody.innerHTML = '';
        emptyMsg.style.display = 'block';
        return;
    }

    emptyMsg.style.display = 'none';
    tbody.innerHTML = productions.map((p, i) => {
        const product = getProductByCode(p.code);
        const productName = product ? product.name : '(삭제된 품목)';
        const johnsonCode = product ? (product.spec || '-') : '-';
        const price = product ? product.price : p.price;
        const sales = price * p.quantity;
        return `
        <tr>
            <td>${p.code}</td>
            <td>${johnsonCode}</td>
            <td>${productName}</td>
            <td class="text-right">${formatNumber(p.quantity)}</td>
            <td>${p.batch}</td>
            <td>${p.date}</td>
            <td class="text-right">${formatNumber(price)}</td>
            <td class="text-right">${formatNumber(sales)}</td>
            <td><button class="btn btn-danger" onclick="deleteProduction(${i})">삭제</button></td>
        </tr>
        `;
    }).join('');
}

document.getElementById('productionForm').addEventListener('submit', e => {
    e.preventDefault();
    const code = document.getElementById('prodProductCode').value;
    const batch = document.getElementById('batchNumber').value.trim();
    const quantity = Number(document.getElementById('prodQuantity').value);
    const dateStr = parseBatchDate(batch);

    if (!code) { alert('품목코드를 선택하세요.'); return; }
    if (!dateStr) { alert('배치번호 형식이 올바르지 않습니다. (예: 260309-021)'); return; }

    const product = getProductByCode(code);
    if (!product) { alert('마스터파일에 등록되지 않은 품목입니다.'); return; }

    if (productions.some(p => p.code === code && p.batch === batch)) {
        if (!confirm(`동일 품목의 배치 "${batch}"가 이미 존재합니다. 추가하시겠습니까?`)) return;
    }

    productions.push({
        code,
        batch,
        date: dateStr,
        quantity,
        price: product.price
    });

    saveData('kenvue_productions', productions);
    renderProductions();
    e.target.reset();
    document.getElementById('prodJohnsonCode').value = '';
    document.getElementById('prodProductName').value = '';
    document.getElementById('prodUnitPrice').value = '';
    document.getElementById('prodDate').value = '';
    document.getElementById('prodSales').value = '';
});

window.deleteProduction = function(index) {
    if (confirm('이 생산실적을 삭제하시겠습니까?')) {
        productions.splice(index, 1);
        saveData('kenvue_productions', productions);
        renderProductions();
    }
};

// ========== 매출 현황 ==========
function refreshFilterProductSelect() {
    const select = document.getElementById('filterProduct');
    select.innerHTML = '<option value="">전체</option>' +
        products.map(p => `<option value="${p.code}">${p.code} - ${p.name}</option>`).join('');
}

function updateSales() {
    const startStr = document.getElementById('filterStart').value;
    const endStr = document.getElementById('filterEnd').value;
    const filterCode = document.getElementById('filterProduct').value;

    const start = startStr ? startStr.replace(/-/g, '') : '';
    const end = endStr ? endStr.replace(/-/g, '') : '';

    let filtered = productions.filter(p => {
        if (filterCode && p.code !== filterCode) return false;
        if (start && p.date < start) return false;
        if (end && p.date > end) return false;
        return true;
    });

    const summary = {};
    let totalSales = 0;
    let totalQty = 0;

    filtered.forEach(p => {
        const product = getProductByCode(p.code);
        const price = product ? product.price : p.price;
        const sales = price * p.quantity;

        if (!summary[p.code]) {
            summary[p.code] = {
                code: p.code,
                name: product ? product.name : '(삭제된 품목)',
                quantity: 0,
                price: price,
                sales: 0
            };
        }
        summary[p.code].quantity += p.quantity;
        summary[p.code].sales += sales;
        totalSales += sales;
        totalQty += p.quantity;
    });

    document.getElementById('totalSales').textContent = formatNumber(totalSales) + ' 원';
    document.getElementById('totalQuantity').textContent = formatNumber(totalQty) + ' 개';
    document.getElementById('totalProducts').textContent = Object.keys(summary).length + ' 건';

    const summaryBody = document.getElementById('salesSummaryBody');
    summaryBody.innerHTML = Object.values(summary).map(s => `
        <tr>
            <td>${s.code}</td>
            <td>${s.name}</td>
            <td class="text-right">${formatNumber(s.quantity)}</td>
            <td class="text-right">${formatNumber(s.price)}</td>
            <td class="text-right">${formatNumber(s.sales)}</td>
        </tr>
    `).join('');

    const detailBody = document.getElementById('salesDetailBody');
    detailBody.innerHTML = filtered.map(p => {
        const product = getProductByCode(p.code);
        const price = product ? product.price : p.price;
        return `
        <tr>
            <td>${p.code}</td>
            <td>${product ? product.name : '(삭제된 품목)'}</td>
            <td>${p.batch}</td>
            <td>${formatDateDisplay(p.date)}</td>
            <td class="text-right">${formatNumber(p.quantity)}</td>
            <td class="text-right">${formatNumber(price)}</td>
            <td class="text-right">${formatNumber(price * p.quantity)}</td>
        </tr>
        `;
    }).join('');
}

document.getElementById('filterBtn').addEventListener('click', updateSales);

// ========== 엑셀 가져오기/내보내기 ==========

document.getElementById('importProducts').addEventListener('change', e => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
        const wb = XLSX.read(evt.target.result, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws);

        let count = 0;
        data.forEach(row => {
            const code = String(row['품목코드'] || row['code'] || '').trim();
            const name = String(row['품목명'] || row['name'] || '').trim();
            const spec = String(row['존슨코드'] || row['spec'] || '').trim();
            const price = Number(row['단가'] || row['price'] || 0);

            if (!code || !name || !price) return;

            const existing = products.findIndex(p => p.code === code);
            if (existing >= 0) {
                products[existing] = { code, name, spec, price };
            } else {
                products.push({ code, name, spec, price });
            }
            count++;
        });

        saveData('kenvue_products', products);
        renderProducts();
        refreshProductSelect();
        alert(`${count}건의 품목이 불러와졌습니다.`);
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
});

document.getElementById('exportProducts').addEventListener('click', () => {
    const data = products.map(p => ({
        '품목코드': p.code,
        '품목명': p.name,
        '존슨코드': p.spec,
        '단가': p.price
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '마스터파일');
    XLSX.writeFile(wb, '켄뷰_마스터파일.xlsx');
});

document.getElementById('importProduction').addEventListener('change', e => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
        const wb = XLSX.read(evt.target.result, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws);

        let count = 0;
        data.forEach(row => {
            const code = String(row['품목코드'] || row['code'] || '').trim();
            const batch = String(row['배치번호'] || row['batch'] || '').trim();
            const quantity = Number(row['생산수량'] || row['quantity'] || 0);

            if (!code || !batch || !quantity) return;

            const dateStr = parseBatchDate(batch);
            if (!dateStr) return;

            const product = getProductByCode(code);
            const price = product ? product.price : 0;

            productions.push({ code, batch, date: dateStr, quantity, price });
            count++;
        });

        saveData('kenvue_productions', productions);
        renderProductions();
        alert(`${count}건의 생산실적이 불러와졌습니다.`);
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
});

document.getElementById('exportProduction').addEventListener('click', () => {
    const data = productions.map(p => {
        const product = getProductByCode(p.code);
        const price = product ? product.price : p.price;
        return {
            '품목코드': p.code,
            '품목명': product ? product.name : '',
            '배치번호': p.batch,
            '생산일자': formatDateDisplay(p.date),
            '생산수량': p.quantity,
            '단가': price,
            '매출액': price * p.quantity
        };
    });
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '생산실적');
    XLSX.writeFile(wb, '켄뷰_생산실적.xlsx');
});

document.getElementById('exportSales').addEventListener('click', () => {
    const rows = document.querySelectorAll('#salesDetailBody tr');
    const data = [];
    rows.forEach(tr => {
        const tds = tr.querySelectorAll('td');
        data.push({
            '품목코드': tds[0].textContent,
            '품목명': tds[1].textContent,
            '배치번호': tds[2].textContent,
            '생산일자': tds[3].textContent,
            '생산수량': tds[4].textContent,
            '단가': tds[5].textContent,
            '매출액': tds[6].textContent
        });
    });
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '매출현황');
    XLSX.writeFile(wb, '켄뷰_매출현황.xlsx');
});

// ========== 초기 렌더링 ==========
renderProducts();
refreshProductSelect();
renderProductions();
