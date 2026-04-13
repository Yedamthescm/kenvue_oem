// ========== Firebase 초기화 ==========
const firebaseConfig = {
    apiKey: "AIzaSyA_Km_RuNidOBA5DEOfC0WJ4c3qH0doYTc",
    authDomain: "kenvue-oem-app.firebaseapp.com",
    projectId: "kenvue-oem-app",
    storageBucket: "kenvue-oem-app.firebasestorage.app",
    messagingSenderId: "783771297702",
    appId: "1:783771297702:web:bbd42f6fabe4ebad791167"
};

firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();

// ========== 데이터 저장/로드 (Firestore + localStorage 백업) ==========
function loadData(key) {
    return JSON.parse(localStorage.getItem(key) || '[]');
}

function saveData(key, data) {
    // localStorage 백업 (즉시 반영용)
    localStorage.setItem(key, JSON.stringify(data));
    // Firestore 저장
    db.collection('appData').doc(key).set({ items: data })
        .catch(err => console.error('Firestore 저장 실패:', key, err));
}

async function loadFromFirestore(key) {
    try {
        const doc = await db.collection('appData').doc(key).get();
        if (doc.exists && doc.data().items) {
            const data = doc.data().items;
            localStorage.setItem(key, JSON.stringify(data));
            return data;
        }
    } catch (err) {
        console.error('Firestore 로드 실패:', key, err);
    }
    return loadData(key);
}

let products = loadData('kenvue_products');
let productions = loadData('kenvue_productions');
let orders = loadData('kenvue_orders');
let performanceData = loadData('kenvue_performance');
let asnData = loadData('kenvue_asn');

// Firestore에서 최신 데이터 불러와서 화면 갱신
async function initFromFirestore() {
    products = await loadFromFirestore('kenvue_products');
    productions = await loadFromFirestore('kenvue_productions');
    orders = await loadFromFirestore('kenvue_orders');
    performanceData = await loadFromFirestore('kenvue_performance');
    asnData = await loadFromFirestore('kenvue_asn');
    renderProducts();
    renderOrders();
    renderPerformance();
    renderAsn();
}

initFromFirestore();

// ========== 탭 전환 ==========
document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
        btn.classList.add('active');
        document.getElementById(btn.dataset.tab).classList.add('active');

        if (btn.dataset.tab === 'asn') renderAsn();
        if (btn.dataset.tab === 'sales') {
            refreshFilterProductSelect();
            // 오늘 날짜 기준 매출월을 기본 선택
            const today = getKoreanDate();
            const currentMonth = today.substring(2, 4) + '년 ' + Number(today.substring(5, 7)) + '월';
            const monthSelect = document.getElementById('filterSalesMonth');
            if ([...monthSelect.options].some(o => o.value === currentMonth)) {
                monthSelect.value = currentMonth;
            }
            updateSales();
        }
    });
});

// ========== 한국 시간 ==========
function getKoreanDate() {
    const now = new Date();
    const kr = new Date(now.toLocaleString('en-US', { timeZone: 'Asia/Seoul' }));
    const yyyy = kr.getFullYear();
    const mm = String(kr.getMonth() + 1).padStart(2, '0');
    const dd = String(kr.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
}

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

// ========== 수주 마스터 ==========
function renderOrders() {
    const tbody = document.getElementById('orderTableBody');
    const emptyMsg = document.getElementById('orderEmptyMsg');

    if (orders.length === 0) {
        tbody.innerHTML = '';
        emptyMsg.style.display = 'block';
        return;
    }

    emptyMsg.style.display = 'none';
    tbody.innerHTML = orders.map(o => `
        <tr>
            <td>${o.purchaseOrder || '-'}</td>
            <td>${o.salesDoc}</td>
            <td>${o.material}</td>
            <td>${o.description}</td>
            <td class="text-right">${formatNumber(o.totalQty)}</td>
        </tr>
    `).join('');
}

document.getElementById('importOrders').addEventListener('change', e => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
        const wb = XLSX.read(evt.target.result, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws);

        // 디버깅: 시트 범위 및 첫 5행 raw 데이터 확인
        alert('시트명: ' + wb.SheetNames[0] + '\n범위: ' + (ws['!ref'] || '없음'));
        const rawRows = XLSX.utils.sheet_to_json(ws, { header: 1 });
        alert('첫 5행:\n' + rawRows.slice(0, 5).map((r, i) => i + ': ' + JSON.stringify(r)).join('\n'));

        let count = 0;
        const newOrders = [];
        data.forEach(row => {
            const purchaseOrder = String(row['구매오더번호'] || row['Purchase Order'] || '').trim();
            const salesDoc = String(row['판매문서'] || row['판매 문서'] || row['Sales Document'] || '').trim();
            const material = String(row['자재'] || row['Material'] || '').trim();
            const description = String(row['내역'] || row['Description'] || '').trim();
            const totalQty = Number(row['총오더수량'] || row['Total Order Qty'] || 0);

            if (!salesDoc && !material) return;

            newOrders.push({ purchaseOrder, salesDoc, material, description, totalQty });
            count++;
        });

        orders = newOrders;
        saveData('kenvue_orders', orders);
        renderOrders();
        alert(`${count}건의 수주 데이터가 불러와졌습니다.`);
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
});

document.getElementById('clearOrders').addEventListener('click', () => {
    if (confirm('수주 데이터를 모두 초기화하시겠습니까?')) {
        orders = [];
        saveData('kenvue_orders', orders);
        renderOrders();
    }
});

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
                <td><input type="text" id="editKenvueCode" value="${p.kenvueCode || ''}" class="inline-input"></td>
                <td><input type="text" id="editDescription" value="${p.description || ''}" class="inline-input"></td>
                <td><input type="text" id="editName" value="${p.name}" class="inline-input"></td>
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
            <td>${p.kenvueCode || '-'}</td>
            <td>${p.description || '-'}</td>
            <td>${p.name}</td>
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
    const kenvueCode = document.getElementById('editKenvueCode').value.trim();
    const description = document.getElementById('editDescription').value.trim();
    const name = document.getElementById('editName').value.trim();
    const price = Number(document.getElementById('editPrice').value);

    if (!code || !name || !price) {
        alert('Cosmax Code, 품목명, 단가는 필수입니다.');
        return;
    }

    // 코드 변경 시 중복 체크
    const duplicate = products.findIndex((p, i) => p.code === code && i !== index);
    if (duplicate >= 0) {
        alert('동일한 Cosmax Code가 이미 존재합니다.');
        return;
    }

    products[index] = { code, kenvueCode, description, name, price };
    editingIndex = -1;
    saveData('kenvue_products', products);
    renderProducts();
};

document.getElementById('productForm').addEventListener('submit', e => {
    e.preventDefault();
    const code = document.getElementById('productCode').value.trim();
    const kenvueCode = document.getElementById('productKenvueCode').value.trim();
    const description = document.getElementById('productDescription').value.trim();
    const name = document.getElementById('productName').value.trim();
    const price = Number(document.getElementById('productPrice').value);

    const existing = products.findIndex(p => p.code === code);
    if (existing >= 0) {
        if (confirm(`Cosmax Code "${code}"가 이미 존재합니다. 덮어쓰시겠습니까?`)) {
            products[existing] = { code, kenvueCode, description, name, price };
        } else {
            return;
        }
    } else {
        products.push({ code, kenvueCode, description, name, price });
    }

    saveData('kenvue_products', products);
    renderProducts();
    e.target.reset();
});

window.deleteProduct = function(index) {
    if (confirm('이 품목을 삭제하시겠습니까?')) {
        products.splice(index, 1);
        saveData('kenvue_products', products);
        renderProducts();

    }
};

// ========== 실적 입력 ==========
function getProductByCode(code) {
    return products.find(p => p.code === code);
}

function extractSalesDoc(specialStock) {
    // "특별 재고 번호"에서 "/ 10"을 제외한 숫자만 추출
    const str = String(specialStock || '');
    const cleaned = str.replace(/\/\s*10/, '').trim();
    return cleaned;
}

function renderPerformance() {
    const tbody = document.getElementById('performanceTableBody');
    const emptyMsg = document.getElementById('performanceEmptyMsg');

    if (performanceData.length === 0) {
        tbody.innerHTML = '';
        emptyMsg.style.display = 'block';
        return;
    }

    emptyMsg.style.display = 'none';
    tbody.innerHTML = performanceData.map(p => `
        <tr>
            <td>${p.salesDoc}</td>
            <td>${p.material}</td>
            <td>${p.materialDesc}</td>
            <td>${p.batch}</td>
            <td class="text-right">${formatNumber(p.available)}</td>
            <td>${p.qualityInspection}</td>
        </tr>
    `).join('');
}

document.getElementById('importPerformance').addEventListener('change', e => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
        const wb = XLSX.read(evt.target.result, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws);

        let count = 0;
        const newData = [];
        data.forEach(row => {
            // A열: 특별 재고 번호 → 판매문서 (숫자만, "/ 10" 제외)
            const specialStock = row['특별 재고 번호'] || row['Special Stock Number'] || '';
            const salesDoc = extractSalesDoc(specialStock);

            const material = String(row['자재'] || row['Material'] || '').trim();
            const materialDesc = String(row['자재내역'] || row['Material Description'] || row['자재 내역'] || '').trim();
            const batch = String(row['배치'] || row['Batch'] || '').trim();
            const available = Number(row['가용'] || row['Available'] || 0);
            const qualityInspection = String(row['품질 검사'] || row['Quality Inspection'] || row['품질검사'] || '').trim();

            if (!salesDoc && !material) return;

            newData.push({ salesDoc, material, materialDesc, batch, available, qualityInspection });
            count++;
        });

        performanceData = newData;
        saveData('kenvue_performance', performanceData);
        renderPerformance();

        // ASN 자동 생성 및 누적
        const uploadDate = getKoreanDate();
        const newAsnRows = [];
        newData.forEach(perf => {
            const matchedOrders = findOrdersBySalesDoc(perf.salesDoc);
            const po = matchedOrders.length > 0 ? matchedOrders[0].purchaseOrder : '-';
            const totalQty = matchedOrders.reduce((sum, o) => sum + o.totalQty, 0);

            const product = findProductByMaterial(perf.material);
            const cosmaxCode = product ? product.code : '-';
            const kenvueCode = product ? (product.kenvueCode || '-') : '-';
            const description = product ? (product.description || '-') : '-';
            const price = product ? product.price : 0;

            const qualityQty = Number(perf.qualityInspection) || 0;
            const qty = perf.available + qualityQty;
            const vendorBatch = normalizeBatch(perf.batch, perf.material);
            let mfgDateRaw = '';
            if (vendorBatch !== 'N/A') {
                if (perf.material && YMX_BATCH_MATERIALS.includes(perf.material)) {
                    mfgDateRaw = parseYmxBatchDate(vendorBatch);
                } else {
                    mfgDateRaw = parseBatchDate(vendorBatch);
                }
            }
            const mfgDate = mfgDateRaw ? formatDateDisplay(mfgDateRaw) : 'N/A';
            const sales = price * qty;

            const salesMonth = uploadDate.substring(2, 4) + '년 ' + Number(uploadDate.substring(5, 7)) + '월';

            newAsnRows.push({
                salesMonth, uploadDate, cosmaxCode, po, kenvueCode, description,
                qty, vendorBatch, mfgDate, totalQty, price, sales
            });
        });

        // 배치 기준 중복 처리
        let addedCount = 0;
        let updatedCount = 0;
        let skippedCount = 0;

        newAsnRows.forEach(newRow => {
            if (newRow.vendorBatch === 'N/A') {
                // N/A도 같은 cosmaxCode + qty가 이미 있으면 스킵
                const dupNa = asnData.some(existing =>
                    existing.cosmaxCode === newRow.cosmaxCode &&
                    existing.qty === newRow.qty &&
                    existing.vendorBatch === 'N/A'
                );
                if (dupNa) {
                    skippedCount++;
                } else {
                    asnData.push(newRow);
                    addedCount++;
                }
            } else {
                // 같은 cosmaxCode + vendorBatch가 이미 존재하면 스킵
                const duplicate = asnData.some(existing =>
                    existing.vendorBatch === newRow.vendorBatch &&
                    existing.cosmaxCode === newRow.cosmaxCode
                );

                if (duplicate) {
                    skippedCount++;
                } else {
                    // 기존 N/A 행 중 같은 cosmaxCode + qty가 있으면 배치 정보 업데이트
                    const naIndex = asnData.findIndex(existing =>
                        existing.vendorBatch === 'N/A' &&
                        existing.cosmaxCode === newRow.cosmaxCode &&
                        existing.qty === newRow.qty
                    );

                    if (naIndex >= 0) {
                        asnData[naIndex].vendorBatch = newRow.vendorBatch;
                        asnData[naIndex].mfgDate = newRow.mfgDate;
                        updatedCount++;
                    } else {
                        asnData.push(newRow);
                        addedCount++;
                    }
                }
            }
        });

        saveData('kenvue_asn', asnData);
        renderAsn();
        alert(`${count}건의 실적 데이터가 불러와졌습니다.\nASN: ${addedCount}건 추가, ${updatedCount}건 업데이트, ${skippedCount}건 중복 스킵`);
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
});

document.getElementById('clearPerformance').addEventListener('click', () => {
    if (confirm('실적 데이터를 모두 초기화하시겠습니까?')) {
        performanceData = [];
        saveData('kenvue_performance', performanceData);
        renderPerformance();
    }
});

// ========== 배치번호 정규화 ==========
const SPECIAL_BATCH_MATERIALS = ['9NTG0051110', '9NTG0051020', '9NTG0051230', '9JNJ0020510', '9JNJ0020910'];

// YMX 형식 배치 품목 (Y=연도끝자리, M=월알파벳A~L, X=충진순서)
const YMX_BATCH_MATERIALS = ['9JNJ0020810'];

function normalizeBatch(batch, material) {
    const str = String(batch || '').trim();

    // YMX 형식 품목 (예: 6C1 = 2026년 3월 첫 충진)
    if (material && YMX_BATCH_MATERIALS.includes(material)) {
        if (/^\d[A-L][0-9A-C]$/i.test(str)) return str.toUpperCase();
        return 'N/A';
    }

    // 특수 품목: DDMMYYYYNN (10자리) → YYMMDD-NNN
    if (material && SPECIAL_BATCH_MATERIALS.includes(material)) {
        // 이미 YYMMDD-NNN 형태
        if (/^\d{6}-\d{3}$/.test(str)) return str;

        if (/^\d{10}$/.test(str)) {
            const dd = str.substring(0, 2);
            const mm = str.substring(2, 4);
            const yy = str.substring(6, 8);
            const nn = str.substring(8, 10);
            return yy + mm + dd + '-' + nn.padStart(3, '0');
        }

        return 'N/A';
    }

    // 기본 로직: 이미 YYMMDD-NNN 형태
    if (/^\d{6}-\d{3}$/.test(str)) return str;

    // 9자리 숫자 → YYMMDD-NNN (예: 260309021 → 260309-021)
    if (/^\d{9}$/.test(str)) {
        return str.substring(0, 6) + '-' + str.substring(6, 9);
    }

    return 'N/A';
}

// YMX 배치 → 제조일자 (YYYYMM01)
function parseYmxBatchDate(batch) {
    const str = String(batch || '').trim().toUpperCase();
    if (!/^\d[A-L][0-9A-C]$/.test(str)) return '';
    const y = parseInt(str[0]);
    const baseYear = 2020 + y; // 0=2020, 1=2021, ..., 9=2029 (이후 순환)
    const monthMap = { A:'01', B:'02', C:'03', D:'04', E:'05', F:'06', G:'07', H:'08', I:'09', J:'10', K:'11', L:'12' };
    const mm = monthMap[str[1]];
    if (!mm) return '';
    return String(baseYear) + mm + '01';
}

// ========== ASN ==========
function findProductByMaterial(materialCode) {
    // 실적 입력의 자재를 기준으로 품목 마스터에서 매칭 (Kenvue Code 또는 Cosmax Code)
    return products.find(p => p.kenvueCode === materialCode || p.code === materialCode);
}

function findOrdersBySalesDoc(salesDoc) {
    // 판매문서를 기준으로 수주 마스터에서 모든 매칭 행 반환
    return orders.filter(o => o.salesDoc === salesDoc);
}

function deduplicateAsn() {
    const seen = new Set();
    const before = asnData.length;
    asnData = asnData.filter(row => {
        if (row.vendorBatch === 'N/A') return true;
        const key = row.cosmaxCode + '|' + row.vendorBatch;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    });
    if (asnData.length < before) {
        saveData('kenvue_asn', asnData);
    }
}

function renderAsn() {
    const tbody = document.getElementById('asnTableBody');
    const emptyMsg = document.getElementById('asnEmptyMsg');

    // 기존 데이터 중복 제거 (N/A 제외)
    deduplicateAsn();

    if (asnData.length === 0) {
        tbody.innerHTML = '';
        emptyMsg.style.display = 'block';
        return;
    }

    // 매출월이 없는 기존 데이터는 uploadDate 기준으로 자동 채움
    let needsSave = false;
    asnData.forEach(row => {
        if (!row.salesMonth && row.uploadDate) {
            row.salesMonth = row.uploadDate.substring(2, 4) + '년 ' + Number(row.uploadDate.substring(5, 7)) + '월';
            needsSave = true;
        }
    });
    if (needsSave) saveData('kenvue_asn', asnData);

    emptyMsg.style.display = 'none';
    tbody.innerHTML = asnData.map((row, i) => `
        <tr>
            <td><button class="btn btn-delete-light btn-sm" onclick="deleteAsnRow(${i})">삭제</button></td>
            <td class="editable-cell" onclick="editSalesMonth(${i})" title="클릭하여 수정">${row.salesMonth || ''}</td>
            <td>${row.uploadDate}</td>
            <td>${row.cosmaxCode}</td>
            <td>${row.po}</td>
            <td>${row.kenvueCode}</td>
            <td>${row.description}</td>
            <td class="text-right">${formatNumber(row.qty)}</td>
            <td>${row.vendorBatch}</td>
            <td>${row.mfgDate}</td>
            <td class="text-right">${formatNumber(row.totalQty)}</td>
            <td class="text-right">${formatNumber(row.price)}</td>
            <td class="text-right">${formatNumber(row.sales)}</td>
        </tr>`).join('');
}

window.deleteAsnRow = function(index) {
    if (confirm('이 행을 삭제하시겠습니까?')) {
        asnData.splice(index, 1);
        saveData('kenvue_asn', asnData);
        renderAsn();
    }
};

window.editSalesMonth = function(index) {
    const row = asnData[index];
    const td = document.querySelectorAll('#asnTableBody tr')[index].children[0];
    const currentValue = row.salesMonth || '';
    const input = document.createElement('input');
    input.type = 'text';
    input.value = currentValue;
    input.placeholder = '예: 26년 1월';
    input.className = 'inline-input';
    input.style.width = '100px';
    td.textContent = '';
    td.appendChild(input);
    input.focus();

    function save() {
        const newValue = input.value.trim();
        asnData[index].salesMonth = newValue;
        saveData('kenvue_asn', asnData);
        renderAsn();
    }

    input.addEventListener('blur', save);
    input.addEventListener('keydown', e => {
        if (e.key === 'Enter') { e.preventDefault(); save(); }
        if (e.key === 'Escape') renderAsn();
    });
};

document.getElementById('clearAsn').addEventListener('click', () => {
    if (confirm('ASN 데이터를 모두 초기화하시겠습니까?')) {
        asnData = [];
        saveData('kenvue_asn', asnData);
        renderAsn();
    }
});

document.getElementById('exportAsn').addEventListener('click', () => {
    // 내보내기 시 cosmaxCode + vendorBatch 기준 중복 제거
    const seen = new Set();
    const dedupedAsn = asnData.filter(row => {
        if (row.vendorBatch === 'N/A') return true;
        const key = row.cosmaxCode + '|' + row.vendorBatch;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    });
    const data = dedupedAsn.map(row => ({
        '매출월': row.salesMonth || '',
        '업로드 일자': row.uploadDate,
        'Cosmax Code': row.cosmaxCode,
        'PO': row.po,
        'Kenvue Code': row.kenvueCode,
        'Description (Ready to release)': row.description,
        "Q'ty": row.qty,
        'vendor batch no.': row.vendorBatch,
        'manufacturing date': row.mfgDate,
        "Total Q'ty for RM Usage": row.totalQty,
        '가격': row.price,
        '매출': row.sales
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'ASN');
    XLSX.writeFile(wb, '켄뷰_ASN.xlsx');
});

// ========== 매출 현황 ==========
function refreshFilterProductSelect() {
    const select = document.getElementById('filterProduct');
    select.innerHTML = '<option value="">전체</option>' +
        products.map(p => `<option value="${p.code}">${p.code} - ${p.name}</option>`).join('');

    // 매출월 필터 드롭다운 갱신
    const monthSelect = document.getElementById('filterSalesMonth');
    const months = [...new Set(asnData.map(r => r.salesMonth).filter(m => m))];
    monthSelect.innerHTML = '<option value="">전체</option>' +
        months.map(m => `<option value="${m}">${m}</option>`).join('');
}

function updateSales() {
    const filterMonth = document.getElementById('filterSalesMonth').value;
    const filterCode = document.getElementById('filterProduct').value;

    // ASN 데이터 기준으로 필터링 (매출월 기준)
    let filtered = asnData.filter(row => {
        if (filterCode && row.cosmaxCode !== filterCode) return false;
        if (filterMonth && row.salesMonth !== filterMonth) return false;
        return true;
    });

    const summary = {};
    let totalSalesAmt = 0;
    let totalQty = 0;

    filtered.forEach(row => {
        if (!summary[row.cosmaxCode]) {
            summary[row.cosmaxCode] = {
                code: row.cosmaxCode,
                kenvueCode: row.kenvueCode,
                description: row.description,
                quantity: 0,
                price: row.price,
                sales: 0
            };
        }
        summary[row.cosmaxCode].quantity += row.qty;
        summary[row.cosmaxCode].sales += row.sales;
        totalSalesAmt += row.sales;
        totalQty += row.qty;
    });

    document.getElementById('totalSales').textContent = formatNumber(totalSalesAmt) + ' 원';
    document.getElementById('totalQuantity').textContent = formatNumber(totalQty) + ' 개';
    document.getElementById('totalProducts').textContent = Object.keys(summary).length + ' 건';

    const summaryBody = document.getElementById('salesSummaryBody');
    summaryBody.innerHTML = Object.values(summary).map(s => `
        <tr>
            <td>${s.code}</td>
            <td>${s.description}</td>
            <td class="text-right">${formatNumber(s.quantity)}</td>
            <td class="text-right">${formatNumber(s.price)}</td>
            <td class="text-right">${formatNumber(s.sales)}</td>
        </tr>
    `).join('');

    // 배치 기준 중복 제거 (ASN 탭과 동일하게)
    const seen = new Set();
    const deduped = filtered.filter(row => {
        if (row.vendorBatch === 'N/A') return true;
        const key = row.cosmaxCode + '|' + row.vendorBatch;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    });

    const detailBody = document.getElementById('salesDetailBody');
    detailBody.innerHTML = deduped.map(row => `
        <tr>
            <td>${row.salesMonth || ''}</td>
            <td>${row.uploadDate}</td>
            <td>${row.cosmaxCode}</td>
            <td>${row.kenvueCode}</td>
            <td>${row.description}</td>
            <td>${row.vendorBatch || '-'}</td>
            <td class="text-right">${formatNumber(row.qty)}</td>
            <td class="text-right">${formatNumber(row.price)}</td>
            <td class="text-right">${formatNumber(row.sales)}</td>
        </tr>
    `).join('');
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
            const code = String(row['Cosmax Code'] || row['품목코드'] || row['code'] || '').trim();
            const kenvueCode = String(row['Kenvue Code'] || row['존슨코드'] || '').trim();
            const description = String(row['Description (Ready to release)'] || row['Description'] || '').trim();
            const name = String(row['품목명'] || row['name'] || '').trim();
            const price = Number(row['단가'] || row['price'] || 0);

            if (!code || !name || !price) return;

            const existing = products.findIndex(p => p.code === code);
            if (existing >= 0) {
                products[existing] = { code, kenvueCode, description, name, price };
            } else {
                products.push({ code, kenvueCode, description, name, price });
            }
            count++;
        });

        saveData('kenvue_products', products);
        renderProducts();

        alert(`${count}건의 품목이 불러와졌습니다.`);
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
});

document.getElementById('exportProducts').addEventListener('click', () => {
    const data = products.map(p => ({
        'Cosmax Code': p.code,
        'Kenvue Code': p.kenvueCode || '',
        'Description (Ready to release)': p.description || '',
        '품목명': p.name,
        '단가': p.price
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '품목마스터');
    XLSX.writeFile(wb, '켄뷰_품목마스터.xlsx');
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
renderOrders();
renderPerformance();
renderAsn();
