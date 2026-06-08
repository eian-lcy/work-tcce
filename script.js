// ==========================================
// 1. 設定分類資料
// ==========================================
const categoryData = {
    "鑑定": ["現況鑑定", "都市更新", "安全鑑定", "損害修復", "未報勘驗", "工程鑑估", "鄰房鑑定", "鑑定", "其他"],
    "施工": ["安全評估", "擋土支撐", "施工計畫", "監測報告", "其他"],
    "房屋": ["複審費", "初評", "詳評", "詳評審查", "危老初評", "危老詳評", "公安初評", "公安詳評", "其他"],
    "地敏": ["地下水", "山崩地滑", "地質敏感", "活動斷層", "其他"],
    "水保": ["水保審查", "其他"],
    "耐震": ["耐震標章", "其他"],
    "結構": ["結構外審", "結構變更", "其他"],
    "津貼": ["生育津貼", "結婚津貼", "住院津貼", "喪葬津貼", "其他"],
    "其他": ["其他"]
};

let allBatches = [];
let currentRightClickInput = null; // 暫存目前被按右鍵的輸入框

// ==========================================
// 2. 初始化與介面連動
// ==========================================
window.onload = function () {
    const dateInput = document.getElementById('dateInput');
    if (dateInput) dateInput.valueAsDate = new Date();
    initCategorySelects();
    addRow();

    document.addEventListener('click', function () {
        document.getElementById('payeeContextMenu').style.display = 'none';
    });

    window.addEventListener('beforeunload', function (e) {
        const currentInput = document.querySelector('.payee-name');
        const hasTyping = currentInput && currentInput.value.trim() !== "";

        if (allBatches.length > 0 || hasTyping) {
            e.preventDefault();
            e.returnValue = '';
        }
    });
};

function initCategorySelects() {
    const cat1Select = document.getElementById('cat1Input');
    cat1Select.innerHTML = '';
    for (let key in categoryData) {
        let option = document.createElement('option');
        option.value = key;
        option.textContent = key;
        cat1Select.appendChild(option);
    }
    cat1Select.value = "鑑定";
    handleCat1Change();
}

function handleCat1Change() {
    updateCat2Options();
    updateUIState();
    updateUsagePreview();
}

function handleCat2Change() {
    updateUIState();
    updateUsagePreview();
}

function updateCat2Options() {
    const cat1Value = document.getElementById('cat1Input').value;
    const cat2Select = document.getElementById('cat2Input');
    cat2Select.innerHTML = '';

    const subCategories = categoryData[cat1Value] || [];
    subCategories.forEach(sub => {
        let option = document.createElement('option');
        option.value = sub;
        option.textContent = sub;
        cat2Select.appendChild(option);
    });
}

function updateUIState() {
    const cat1Value = document.getElementById('cat1Input').value;
    const cat2Value = document.getElementById('cat2Input').value;
    const page1Div = document.getElementById('page1');
    const extraNoteDiv = document.getElementById('extraNoteDiv');
    const extraNoteLabel = document.getElementById('extraNoteLabel');
    const extraNoteInput = document.getElementById('extraNoteInput');

    if (cat1Value === "鑑定") {
        page1Div.classList.add('show-scan');
    } else {
        page1Div.classList.remove('show-scan');
        const scanInputs = document.querySelectorAll('.payee-scan');
        scanInputs.forEach(input => input.value = '');
    }

    if (cat1Value === "其他" || cat2Value === "其他") {
        extraNoteDiv.style.display = "block";
        if (cat1Value === "其他") {
            extraNoteLabel.textContent = "備註說明";
            extraNoteInput.placeholder = "請輸入雜支說明...";
        } else {
            extraNoteLabel.textContent = "輸入分類二";
            extraNoteInput.placeholder = "請輸入分類二細項...";
        }
    } else {
        extraNoteDiv.style.display = "none";
        extraNoteInput.value = "";
    }

    const caseNoInput = document.getElementById('caseNoInput');
    const deptInput = document.getElementById('deptInput');
    const caseRow = caseNoInput.closest('p');
    const deptRow = deptInput.closest('p');

    if (cat1Value === "津貼") {
        caseRow.style.display = 'none';
        deptRow.style.display = 'none';
        caseNoInput.value = '';
        deptInput.value = '';
    } else {
        caseRow.style.display = 'flex';
        deptRow.style.display = 'flex';
    }
}

// ==========================================
// 3. 受款人自動完成與右鍵選單邏輯
// ==========================================
function autoExpandPayee(input) {
    const val = input.value.trim();
    if (payeeRegistry[val]) {
        input.setAttribute('data-short-name', val);
        input.value = payeeRegistry[val];
    }
}

function showContextMenu(e, input) {
    if (!input.getAttribute('data-short-name') && input.value === "") {
        return;
    }

    e.preventDefault();
    currentRightClickInput = input;

    const menu = document.getElementById('payeeContextMenu');
    menu.style.display = 'block';
    menu.style.left = e.pageX + 'px';
    menu.style.top = e.pageY + 'px';
}

function revertToShortName() {
    if (currentRightClickInput) {
        const shortName = currentRightClickInput.getAttribute('data-short-name');
        if (shortName) {
            currentRightClickInput.value = shortName;
        }
    }
    document.getElementById('payeeContextMenu').style.display = 'none';
}

// ==========================================
// 4. 核心邏輯
// ==========================================
function getFormattedCaseNo(caseNo, mainCat) {
    if (caseNo && caseNo.includes("-")) {
        const parts = caseNo.split("-");
        if (parts.length === 2) {
            let middle = "";
            if (mainCat === "鑑定") middle = "鑑";
            else if (mainCat === "施工") middle = "施";
            else if (mainCat === "耐震") middle = "耐";
            else if (mainCat === "地敏") middle = "地";
            else if (mainCat === "水保") middle = "水";
            else if (mainCat === "房屋") middle = "房";
            else if (mainCat === "結構") middle = "結";

            if (middle) {
                return `${parts[0]}${middle}${parts[1]}號`;
            }
        }
    }
    return caseNo;
}

function getSuffix(subCat) {
    const target = (subCat || "").trim();
    if (subCat === "耐震標章") return "(耐震標章)";
    const standardSuffixes = [
        "現況鑑定", "都市更新", "安全鑑定", "損害修復", "未報勘驗", "工程鑑估",
        "鄰房鑑定", "安全評估", "擋土支撐", "施工計畫", "監測報告",
        "初評", "危老初評", "詳評", "詳評審查", "危老初評", "危老詳評", "公安詳評",
        "地下水", "山崩地滑", "地質敏感", "活動斷層", "水保審查", "結構外審", "結構變更", "鑑定"
    ];
    if (standardSuffixes.includes(subCat)) {
        return `(${subCat})`;
    }
    if (subCat === "其他") return "(其他)";
    return "";
}

function updateUsagePreview() {
    const cat1 = document.getElementById('cat1Input').value;
    const cat2 = document.getElementById('cat2Input').value;
    const rawCaseNo = document.getElementById('caseNoInput').value.trim();
    const applicant = document.getElementById('deptInput').value.trim();
    const extraNote = document.getElementById('extraNoteInput').value.trim();

    let finalUsage = "";

    if (cat1 === "津貼") {
        if (cat2 === "喪葬津貼") {
            finalUsage = "奠儀";
        } else if (cat2 === "其他") {
            // 當選擇津貼的「其他」時，優先使用備註細項(extraNote)
            // 如果還沒輸入，就先暫時顯示「其他津貼」防呆
            finalUsage = extraNote ? extraNote : "其他津貼"; 
        } else {
            finalUsage = cat2;
        }
    }
    else if (cat1 === "其他") {
        // 主分類為「其他」時的邏輯維持不變
        finalUsage = extraNote ? extraNote : "雜支";
    }
    else {
        const fmtCaseNo = getFormattedCaseNo(rawCaseNo, cat1);
        if (cat1 === "房屋" && cat2 === "複審費") {
            finalUsage = `${fmtCaseNo}複審費`;
        }
        else {
            let suffix = getSuffix(cat2);
            if (cat2 === "其他" && extraNote) {
                suffix = `(${extraNote})`;
            }
            else if (cat1 === "鑑定" && extraNote) {
                if (suffix.endsWith(")")) {
                    suffix = suffix.slice(0, -1) + "-" + extraNote + ")";
                } else {
                    suffix = `(${extraNote})`;
                }
            }

            const applicantPart = applicant ? `${applicant}` : "";
            finalUsage = `${fmtCaseNo}執行所得-${applicantPart}${suffix}`;
        }
    }

    document.getElementById('usagePreview').value = finalUsage;
}

// ==========================================
// 5. 表單與操作邏輯
// ==========================================
function handleEnter(event) {
    if (event.key === "Enter") {
        event.preventDefault();
        addRow();
        setTimeout(() => {
            const rows = document.querySelectorAll('.row-inputs');
            rows[rows.length - 1].querySelector('.payee-name').focus();
        }, 50);
    }
}

document.addEventListener('keydown', function(event) {
    // 檢查是否按下 Alt + Enter
    if (event.altKey && event.key === 'Enter') {
        // 防止 Enter 鍵在 textarea 或表單中觸發預設的換行或提交動作
        event.preventDefault(); 
        
        console.log("偵測到 Alt + Enter！");
        
        // 呼叫函式
        saveCurrentBatch();
    }
});

function addRow(nameVal = "", amtVal = "", scanVal = "") {
    const container = document.getElementById('dynamicRows');
    const div = document.createElement('div');
    div.className = 'row-inputs';
    div.innerHTML = `
<input type="text" placeholder="撥款人姓名" class="payee-name" value="${nameVal}" 
       onblur="autoExpandPayee(this)" oncontextmenu="showContextMenu(event, this)">
<input type="text" placeholder="金額" class="payee-amount" value="${amtVal}" 
       oninput="calculateTotal()" onkeydown="handleEnter(event)">
<input type="text" placeholder="掃描費" class="payee-scan" value="${scanVal}" 
       oninput="calculateTotal()" onkeydown="handleEnter(event)">
<button type="button" class="btn-delete" onclick="removeRow(this)">刪除</button>
`;
    container.appendChild(div);
}

function calculateTotal() {
    let amountTotal = 0; // 金額總計
    let scanTotal = 0;   // 掃描總計

    // 抓取畫面上所有的金額與掃描費輸入框
    const amtInputs = document.querySelectorAll('.payee-amount');
    const scanInputs = document.querySelectorAll('.payee-scan');

    // 解析輸入內容的防呆小工具
    const parseValue = (input) => {
        let val = input.value.trim();
        if (!val) return 0;

        // 如果發現是以 = 開頭，嘗試拆掉並計算公式
        if (val.startsWith('=')) {
            try {
                const formula = val.substring(1);
                // 使用 new Function 安全地計算 "100+200" 這樣的算式
                const result = new Function(`return ${formula}`)();
                return (!isNaN(result) && isFinite(result)) ? parseFloat(result) : 0;
            } catch (e) {
                return 0; // 公式未打完或格式錯誤時先當 0，不卡死程式
            }
        }
        return parseFloat(val) || 0; // 一般數字正常轉換
    };

    // 分別累加金額與掃描費
    amtInputs.forEach(input => amountTotal += parseValue(input));
    scanInputs.forEach(input => scanTotal += parseValue(input));

    // 將計算結果即時寫入網頁對應的 <span> 區塊
    const totalAmountDisplay = document.getElementById('totalAmount');
    const totalScanDisplay = document.getElementById('totalScan');

    if (totalAmountDisplay) {
        totalAmountDisplay.innerText = amountTotal.toLocaleString(); // 加上千分位
    }
    if (totalScanDisplay) {
        totalScanDisplay.innerText = scanTotal.toLocaleString(); // 加上千分位
    }
}

function removeRow(btn) {
    const row = btn.parentNode;
    const container = document.getElementById('dynamicRows');
    if (container.children.length > 1) {
        row.remove();
    } else {
        row.querySelector('.payee-name').value = '';
        row.querySelector('.payee-amount').value = '';
        row.querySelector('.payee-scan').value = '';
    }
}

function saveCurrentBatch() {
    const date = document.getElementById('dateInput').value;
    const rawCaseNo = document.getElementById('caseNoInput').value.trim();
    const cat1 = document.getElementById('cat1Input').value;
    const cat2 = document.getElementById('cat2Input').value;
    const dept = document.getElementById('deptInput').value.trim();
    const extraNote = document.getElementById('extraNoteInput').value.trim();
    const usage = document.getElementById('usagePreview').value.trim();
    const editingId = document.getElementById('editingBatchId').value;

    if (!date) { alert("請選擇日期"); return; }

    if (!editingId) {
        if (cat1 !== "津貼" && cat1 !== "其他" && !rawCaseNo) {
            alert("請輸入案號"); return;
        }
    }

    if (!usage) { alert("用途欄位空白"); return; }

    const names = document.querySelectorAll('.payee-name');
    const amounts = document.querySelectorAll('.payee-amount');
    const scans = document.querySelectorAll('.payee-scan');

    let batchPayees = [];
    let batchTotal = 0;
    const fmtCaseNo = getFormattedCaseNo(rawCaseNo, cat1);

    for (let i = 0; i < names.length; i++) {
        let n = names[i].value.trim();
        let a = parseFloat(amounts[i].value) || 0;
        let s = parseFloat(scans[i].value) || 0;

        if (n) {
            let autoTax = false;
            let autoHealth = false;

            if (cat1 === "津貼") {
                autoTax = false;
                autoHealth = false;
            } else if (n.includes("公司") || n.includes("郵局") || n === "麗邦企業社") {
                autoTax = false;
                autoHealth = false;
            } else if (n.includes("事務所")) {
                autoTax = true;
                autoHealth = false;
            } else {
                autoTax = true;
                autoHealth = (a >= 20000); 
            }

            batchPayees.push({
                name: n,
                amount: a,
                scanFee: s,
                incomeUsage: usage,
                scanUsage: `${fmtCaseNo}掃描費`,
                isTax: autoTax,
                isHealth: autoHealth
            });
            batchTotal += a;
            if (s > 0) batchTotal -= s;
        }
    }

    if (batchPayees.length === 0) {
        alert("請至少輸入一筆撥款人資料");
        return;
    }

    if (editingId) {
        const index = allBatches.findIndex(b => b.id == editingId);
        if (index !== -1) {
            allBatches[index] = {
                id: parseInt(editingId),
                date: date,
                cat1: cat1,
                cat2: cat2,
                rawCaseNo: rawCaseNo,
                dept: dept,
                extraNote: extraNote,
                usage: usage,
                payees: batchPayees,
                total: batchTotal
            };
        }
        document.getElementById('editingBatchId').value = "";
        const btnSave = document.getElementById('btnSaveBatch');
        btnSave.textContent = "⬇ 加入暫存列表 (輸入下一案)";
        btnSave.classList.remove('editing');
    } else {
        allBatches.push({
            id: Date.now(),
            date: date,
            cat1: cat1,
            cat2: cat2,
            rawCaseNo: rawCaseNo,
            dept: dept,
            extraNote: extraNote,
            usage: usage,
            payees: batchPayees,
            total: batchTotal
        });
    }

    renderSavedList();
    clearInputs();

    const deptField = document.getElementById('deptInput');
    if (deptField) {
        deptField.value = '';
        // 如果你的「用途預覽」是根據申請單位連動的，需要觸發這一行：
        if (typeof updateUsagePreview === 'function') {
            updateUsagePreview();
        }
    }
}

function clearInputs() {
    document.getElementById('caseNoInput').value = '';
    document.getElementById('extraNoteInput').value = '';
    document.getElementById('usagePreview').value = '';

    const container = document.getElementById('dynamicRows');
    container.innerHTML = '';
    addRow();

    document.getElementById('caseNoInput').focus();
}

function editBatch(id) {
    const batch = allBatches.find(b => b.id === id);
    if (!batch) return;

    document.getElementById('dateInput').value = batch.date;
    document.getElementById('cat1Input').value = batch.cat1;
    updateCat2Options(); 
    document.getElementById('cat2Input').value = batch.cat2;

    updateUIState();

    document.getElementById('caseNoInput').value = batch.rawCaseNo || "";
    document.getElementById('deptInput').value = batch.dept || "";
    document.getElementById('extraNoteInput').value = batch.extraNote || "";
    document.getElementById('usagePreview').value = batch.usage;

    const container = document.getElementById('dynamicRows');
    container.innerHTML = '';
    batch.payees.forEach(p => {
        addRow(p.name, p.amount, p.scanFee);
    });

    document.getElementById('editingBatchId').value = batch.id;
    const btnSave = document.getElementById('btnSaveBatch');
    btnSave.textContent = "💾 更新此案件資料";
    btnSave.classList.add('editing');

    window.scrollTo({ top: 0, behavior: 'smooth' });

    const dateInput = document.getElementById('dateInput');
    dateInput.style.backgroundColor = "#fff3cd";
    setTimeout(() => { dateInput.style.backgroundColor = ""; }, 1000);
}

function jumpToEditBatch(id) {
    document.getElementById('page2').classList.remove('active');
    document.getElementById('page1').classList.add('active');
    editBatch(id);
    alert("已載入案件資料，請修改「撥款日期」或「內容」後，按下「更新」按鈕。");
}

function deleteBatch(id) {
    if (confirm("確定要刪除此筆暫存資料嗎？")) {
        allBatches = allBatches.filter(b => b.id !== id);
        renderSavedList();

        const editingId = document.getElementById('editingBatchId').value;
        if (editingId == id) {
            clearInputs();
            document.getElementById('editingBatchId').value = "";
            const btnSave = document.getElementById('btnSaveBatch');
            btnSave.textContent = "⬇ 加入暫存列表 (輸入下一案)";
            btnSave.classList.remove('editing');
        }
    }
}

function renderSavedList() {
    const listDiv = document.getElementById('savedList');
    const container = document.getElementById('savedListContainer');
    document.getElementById('savedCount').innerText = allBatches.length;

    if (allBatches.length > 0) {
        container.style.display = 'block';
    } else {
        container.style.display = 'none';
    }

    listDiv.innerHTML = '';
    allBatches.forEach((batch, index) => {
        let item = document.createElement('div');
        item.className = 'saved-item';
        item.innerHTML = `
    <div style="flex:1">
        <span><strong>#${index + 1}</strong> ${batch.date} - ${batch.usage}</span><br>
        <span style="font-size:0.85em; color:#666;">$${batch.total.toLocaleString()} (${batch.payees.length}人)</span>
    </div>
    <div class="saved-item-actions">
        <button class="btn-edit-item" onclick="editBatch(${batch.id})">編輯</button>
        <button class="btn-delete-item" onclick="deleteBatch(${batch.id})">刪除</button>
    </div>
`;
        listDiv.appendChild(item);
    });
}

// ==========================================
// 6. 產生報表
// ==========================================
function generateFinalReport() {
    const currentName = document.querySelector('.payee-name').value;
    if (currentName && (allBatches.length === 0 || document.getElementById('editingBatchId').value)) {
        if (confirm("您目前的輸入尚未加入/更新暫存，是否執行？")) {
            saveCurrentBatch();
        } else {
            return;
        }
    } else if (allBatches.length === 0) {
        alert("目前沒有任何案件資料！");
        return;
    }

    updateReportDateSelect();
    renderReportForSelectedDate();

    document.getElementById('page1').classList.remove('active');
    document.getElementById('page2').classList.add('active');
}

function updateReportDateSelect() {
    const dates = [...new Set(allBatches.map(b => b.date))].sort();
    const select = document.getElementById('reportDateSelect');
    select.innerHTML = '';
    dates.forEach(d => {
        const option = document.createElement('option');
        option.value = d;
        option.text = d;
        select.appendChild(option);
    });
    if (dates.length > 0) {
        select.value = dates[dates.length - 1];
    }
}

function renderReportForSelectedDate() {
    const selectedDate = document.getElementById('reportDateSelect').value;
    const reportDiv = document.getElementById('reportContent');
    reportDiv.innerHTML = '';

    const targetBatches = allBatches.filter(b => b.date === selectedDate);
    const peopleMap = {};

    targetBatches.forEach(batch => {
        batch.payees.forEach((p, pIndex) => {
            if (p.amount !== 0) {
                if (!peopleMap[p.name]) peopleMap[p.name] = [];
                peopleMap[p.name].push({
                    usage: p.incomeUsage,
                    amount: p.amount,
                    type: 'income',
                    isTax: p.isTax,
                    isHealth: p.isHealth,
                    refBatchId: batch.id,
                    refPayeeIndex: pIndex
                });
            }
            if (p.scanFee > 0) {
                if (!peopleMap[p.name]) peopleMap[p.name] = [];
                peopleMap[p.name].push({
                    usage: p.scanUsage,
                    amount: -p.scanFee,
                    type: 'scan',
                    isTax: false,
                    isHealth: false,
                    refBatchId: batch.id,
                    refPayeeIndex: pIndex
                });
            }
        });
    });

    let tableHtml = `
<div class="batch-container">
    <table>
        <thead>
            <tr>
                <th width="15%">受款人</th>
                <th width="15%">總實付金額</th>
                <th width="35%">用途</th>
                <th width="12%">案件金額</th>
                <th width="10%">10% 扣繳</th>
                <th width="13%">健保費</th>
            </tr>
        </thead>
        <tbody id="reportTableBody">
`;

    for (let [name, entries] of Object.entries(peopleMap)) {
        let initialPersonNet = 0;
        entries.forEach(e => {
            let tax = e.isTax ? Math.round(e.amount * 0.1) : 0;
            let health = e.isHealth ? Math.round(e.amount * 0.0211) : 0;
            if (e.amount < 0) { tax = 0; health = 0; }
            initialPersonNet += (e.amount - tax - health);
        });

        const rowCount = entries.length;

        entries.forEach((entry, index) => {
            const isScan = entry.type === 'scan';
            const rowClass = isScan ? 'scan-fee-row' : '';
            const disabledAttr = isScan ? 'disabled' : '';
            const taxChecked = entry.isTax ? 'checked' : '';
            const healthChecked = entry.isHealth ? 'checked' : '';

            tableHtml += `<tr class="calc-row ${rowClass}" 
        data-name="${name}" 
        data-batch-id="${entry.refBatchId}" 
        data-payee-index="${entry.refPayeeIndex}"
        data-type="${entry.type}">`;

            if (index === 0) {
                tableHtml += `
            <td rowspan="${rowCount}" class="merged-cell name-cell">${name}</td>
            <td rowspan="${rowCount}" class="merged-cell net-cell" id="merged-net-${name}">
                ${initialPersonNet.toLocaleString()}
            </td>
        `;
            }

            tableHtml += `
            <td>
                <div class="usage-cell-container">
                    <input type="text" class="edit-input usage-input" value="${entry.usage}">
                    <button class="btn-mini-edit" onclick="jumpToEditBatch(${entry.refBatchId})" title="修改此筆案件資料(含日期)">✏️</button>
                </div>
            </td>
            <td>
                <input type="number" class="edit-input amount-input" value="${entry.amount}" oninput="calculateRow(this)">
            </td>
            <td>
                <div class="check-cell">
                    <input type="checkbox" class="chk-tax" onchange="calculateRow(this)" ${disabledAttr} ${taxChecked}>
                    <span class="val-tax">0</span>
                </div>
            </td>
            <td>
                <div class="check-cell">
                    <input type="checkbox" class="chk-health" onchange="calculateRow(this)" ${disabledAttr} ${healthChecked}>
                    <span class="val-health">0</span>
                </div>
                <span class="val-net" style="display:none;">0</span> 
            </td>
        </tr>
    `;
        });

        tableHtml += `
    <tr class="subtotal-row" id="subtotal-${name}">
        <td colspan="2" style="background-color:#fff;"></td>
        <td class="heji">合計</td>
        <td class="heji sub-listed">0</td>
        <td class="heji sub-tax">0</td>
        <td class="heji sub-health">0</td>
    </tr>
`;
    }

    tableHtml += `
        <tr class="grand-total-row">
            <td colspan="2"></td>
            <td style="text-align:right">總計</td>
            <td id="grandList">0</td>
            <td id="grandTax">0</td>
            <td id="grandHealth">0</td>
        </tr>
    </tbody>
</table>
</div>
`;

    reportDiv.innerHTML = tableHtml;
    recalculateAll();
}

function updateUsageData(input) {
    const row = input.closest('tr');
    const batchId = parseInt(row.getAttribute('data-batch-id'));
    const payeeIndex = parseInt(row.getAttribute('data-payee-index'));
    const type = row.getAttribute('data-type');
    const newVal = input.value;

    const batch = allBatches.find(b => b.id === batchId);
    if (batch && batch.payees[payeeIndex]) {
        if (type === 'income') {
            batch.payees[payeeIndex].incomeUsage = newVal;
        } else if (type === 'scan') {
            batch.payees[payeeIndex].scanUsage = newVal;
        }
    }
}

function calculateRow(element) {
    const row = element.closest('tr');
    const name = row.getAttribute('data-name');
    
    // 抓出這個人的所有資料列
    const personRows = document.querySelectorAll(`.calc-row[data-name="${name}"]`);
    
    // 自動判斷身分 (沿用你在 saveCurrentBatch 裡的邏輯)
    let payeeType = '個人';
    if (name.includes("公司") || name.includes("郵局") || name.includes("企業") || name === "麗邦企業社") {
        payeeType = '公司';
    } else if (name.includes("事務所")) {
        payeeType = '事務所';
    }

    // 呼叫群組計算函數
    calculatePayeeGroup(name, payeeType, personRows);
}

function calculatePayeeGroup(name, payeeType, rows) {
    let eligibleForNHI = 0;

    // --- 第一階段：若是個人，先計算「二代健保有效累計金額」 ---
    if (payeeType === '個人') {
        rows.forEach(row => {
            let amountInput = row.querySelector('.amount-input');
            let usageInput = row.querySelector('.usage-input'); // 用途欄位
            
            let amount = parseFloat(amountInput.value) || 0;
            let usage = usageInput ? usageInput.value : '';

            // 判斷是否為排除項目 (負項、複審費、津貼)
            let isExcluded = usage.includes('複審費') || usage.includes('津貼') || amount < 0;

            if (!isExcluded) {
                eligibleForNHI += amount;
            }
        });
    }

    let sumListed = 0;
    let sumTax = 0;
    let sumHealth = 0;
    let sumNet = 0;

    // --- 第二階段：逐筆計算並更新每一列的畫面與資料 ---
    rows.forEach(row => {
        let amountInput = row.querySelector('.amount-input');
        let usageInput = row.querySelector('.usage-input');
        let chkTax = row.querySelector('.chk-tax');
        let chkHealth = row.querySelector('.chk-health');
        let spanTax = row.querySelector('.val-tax');
        let spanHealth = row.querySelector('.val-health');
        let spanNet = row.querySelector('.val-net');

        let amount = parseFloat(amountInput.value) || 0;
        let usage = usageInput ? usageInput.value : '';
        
        let tax = 0;
        let nhi = 0;

        // 規則 1: 公司 (無扣繳、無健保)
        if (payeeType === '公司') {
            tax = 0;
            nhi = 0;
            if(chkTax) chkTax.checked = false;
            if(chkHealth) chkHealth.checked = false;
        } 
        // 規則 2: 事務所 (扣 10% 稅，無健保)
        else if (payeeType === '事務所') {
            tax = Math.round(amount * 0.1); 
            nhi = 0;
            if(chkTax) chkTax.checked = true;
            if(chkHealth) chkHealth.checked = false;
        } 
        // 規則 3: 個人 (累計超兩萬扣健保，排除特定項目)
        else if (payeeType === '個人') {
            let isExcluded = usage.includes('複審費') || usage.includes('津貼') || amount < 0;

            // 個人的稅額依據 checkbox 是否被手動勾選來決定
            if (chkTax && chkTax.checked && amount > 0) {
                tax = Math.round(amount * 0.1);
            }

            // 健保費判斷：非排除項目 且 累計達兩萬
            if (!isExcluded && eligibleForNHI >= 20000) {
                nhi = Math.round(amount * 0.0211);
                if(chkHealth) chkHealth.checked = true; // 自動打勾
            } else {
                nhi = 0;
                if(chkHealth) chkHealth.checked = false; // 自動取消打勾
            }
        }

        // 更新單列 HTML 顯示數值
        if (spanTax) spanTax.textContent = tax > 0 ? tax : "0";
        if (spanHealth) spanHealth.textContent = nhi > 0 ? nhi : "0";
        
        let net = amount - tax - nhi;
        if (spanNet) spanNet.textContent = net;

        // 累加該受款人的小計
        sumListed += amount;
        sumTax += tax;
        sumHealth += nhi;
        sumNet += net;

        // 同步回寫到 JavaScript 記憶體 (allBatches) 中，以便 Excel 匯出正確
        const batchId = parseInt(row.getAttribute('data-batch-id'));
        const payeeIndex = parseInt(row.getAttribute('data-payee-index'));
        const type = row.getAttribute('data-type');
        const batch = allBatches.find(b => b.id === batchId);
        
        if (batch && batch.payees[payeeIndex]) {
            if (type === 'income') {
                batch.payees[payeeIndex].amount = amount;
                batch.payees[payeeIndex].isTax = (tax > 0);
                batch.payees[payeeIndex].isHealth = (nhi > 0);
            } else if (type === 'scan') {
                batch.payees[payeeIndex].scanFee = Math.abs(amount);
            }
        }
    });

    // --- 第三階段：更新報表上的小計列與合併儲存格 ---
    const subRow = document.getElementById(`subtotal-${name}`);
    if (subRow) {
        subRow.querySelector('.sub-listed').textContent = sumListed.toLocaleString();
        subRow.querySelector('.sub-tax').textContent = sumTax.toLocaleString();
        subRow.querySelector('.sub-health').textContent = sumHealth.toLocaleString();
    }

    const mergedNetCell = document.getElementById(`merged-net-${name}`);
    if (mergedNetCell) {
        mergedNetCell.textContent = sumNet.toLocaleString();
    }

    // 觸發最下方的總計更新
    updateGrandTotal();
}

function updateGrandTotal() {
    let grandList = 0;
    let grandTax = 0;
    let grandHealth = 0;

    document.querySelectorAll('.subtotal-row').forEach(row => {
        grandList += parseFloat(row.querySelector('.sub-listed').textContent.replace(/,/g, '')) || 0;
        grandTax += parseFloat(row.querySelector('.sub-tax').textContent.replace(/,/g, '')) || 0;
        grandHealth += parseFloat(row.querySelector('.sub-health').textContent.replace(/,/g, '')) || 0;
    });

    document.getElementById('grandList').textContent = grandList.toLocaleString();
    document.getElementById('grandTax').textContent = grandTax.toLocaleString();
    document.getElementById('grandHealth').textContent = grandHealth.toLocaleString();
}

function recalculateAll() {
    // 1. 恢復這段被註解的程式碼，讓畫面初始載入時觸發計算
    document.querySelectorAll('.amount-input').forEach(input => {
        calculateRow(input);
    });

    // 2. 原本的 addEventListener 可以移除
    // 因為你在 renderReportForSelectedDate 的 HTML 模板裡
    // 已經寫了 oninput="calculateRow(this)"，重複綁定可能會造成預期外的衝突
}

function goBack() {
    document.getElementById('page2').classList.remove('active');
    document.getElementById('page1').classList.add('active');
}

function importFromExcel(input) {
    const file = input.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        if (!workbook.Sheets["RawData"]) {
            alert("錯誤：找不到 'RawData' 分頁，無法還原資料。");
            return;
        }

        const rawJson = XLSX.utils.sheet_to_json(workbook.Sheets["RawData"]);
        const batchesMap = {};

        rawJson.forEach(row => {
            const id = row.BatchID;
            if (!batchesMap[id]) {
                batchesMap[id] = {
                    id: id,
                    date: row.Date,
                    cat1: row.Cat1,
                    cat2: row.Cat2,
                    rawCaseNo: row.RawCaseNo,
                    dept: row.Dept,
                    extraNote: row.ExtraNote,
                    usage: row.IncomeUsage,
                    payees: [],
                    total: 0
                };
            }

            const amount = parseFloat(row.Amount) || 0;
            const scanFee = parseFloat(row.ScanFee) || 0;

            const isTax = (String(row.IsTax).toUpperCase() === "TRUE");
            const isHealth = (String(row.IsHealth).toUpperCase() === "TRUE");

            batchesMap[id].payees.push({
                name: row.PayeeName,
                amount: amount,
                scanFee: scanFee,
                incomeUsage: row.IncomeUsage,
                scanUsage: row.ScanUsage,
                isTax: isTax,
                isHealth: isHealth
            });

            batchesMap[id].total += amount;
            if (scanFee > 0) batchesMap[id].total -= scanFee;
        });

        allBatches = Object.values(batchesMap);

        alert(`成功匯入 ${allBatches.length} 筆案件資料！`);
        generateFinalReport();
        input.value = '';
    };
    reader.readAsArrayBuffer(file);
}

// ==========================================
// 7. 彈出視窗控制邏輯
// ==========================================
function openRegistryModal() {
    const modal = document.getElementById('registryModal');
    const tbody = document.getElementById('registryTableBody');
    const searchInput = document.getElementById('registrySearchInput');

    searchInput.value = '';
    tbody.innerHTML = '';

    const entries = Object.entries(rawPayeeRegistry);

    if (entries.length === 0) {
        tbody.innerHTML = '<tr><td colspan="2" style="text-align:center; color:#999;">目前尚無設定對照表</td></tr>';
    } else {
        entries.forEach(([short, full]) => {
            const displayShort = short.replace(/,/g, '<br><span style="color:#aaa; font-size:0.9em;">或者是</span> ');

            const tr = document.createElement('tr');
            tr.innerHTML = `
        <td style="font-weight:bold; color:#2c3e50;">${displayShort}</td> <td>${full}</td>
    `;
            tbody.appendChild(tr);
        });
    }

    modal.style.display = "block";

    setTimeout(() => {
        searchInput.focus();
    }, 100);
}

function closeRegistryModal() {
    document.getElementById('registryModal').style.display = "none";
}

function filterRegistryTable() {
    const input = document.getElementById('registrySearchInput');
    const filter = input.value.toUpperCase();
    const tbody = document.getElementById('registryTableBody');
    const tr = tbody.getElementsByTagName('tr');

    for (let i = 0; i < tr.length; i++) {
        const tdShort = tr[i].getElementsByTagName("td")[0];
        const tdFull = tr[i].getElementsByTagName("td")[1];

        if (tdShort && tdFull) {
            const txtShort = tdShort.textContent || tdShort.innerText;
            const txtFull = tdFull.textContent || tdFull.innerText;

            if (txtShort.toUpperCase().indexOf(filter) > -1 || txtFull.toUpperCase().indexOf(filter) > -1) {
                tr[i].style.display = ""; 
            } else {
                tr[i].style.display = "none"; 
            }
        }
    }
}

window.onclick = function (event) {
    const modal = document.getElementById('registryModal');
    if (event.target == modal) {
        modal.style.display = "none";
    }
}

// --- 1. 計算機基礎功能 ---
function toggleCalculator() {
    const calc = document.getElementById('calculator');
    calc.style.display = (calc.style.display === 'block') ? 'none' : 'block';
}

function calcAppend(val) {
    document.getElementById('calcDisplay').value += val;
}

function calcClear() {
    document.getElementById('calcDisplay').value = '';
}

function calcBackspace() {
    const display = document.getElementById('calcDisplay');
    display.value = display.value.slice(0, -1);
}

function calcCalculate() {
    const display = document.getElementById('calcDisplay');
    try {
        // 使用 Function 代替 eval 較安全
        const result = new Function('return ' + display.value)();
        display.value = Number.isInteger(result) ? result : result.toFixed(2);
    } catch (e) {
        display.value = "Error";
        setTimeout(calcClear, 1000);
    }
}

// --- 2. 鍵盤支援與千分位邏輯 ---
document.addEventListener('keydown', function(e) {
    // 只有當計算機顯示時才監聽
    if (document.getElementById('calculator').style.display === 'block') {
        if (/[0-9\+\-\*\/\.]/.test(e.key)) calcAppend(e.key);
        if (e.key === 'Enter') calcCalculate();
        if (e.key === 'Backspace') calcBackspace();
        if (e.key === 'Escape') calcClear();
    }
});

// 金額千分位格式化（補強現有邏輯）
function formatNumber(num) {
    return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

// --- 3. 拖曳邏輯 ---
dragElement(document.getElementById("calculator"));

function dragElement(elmnt) {
    var pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
    const header = document.getElementById("calcHeader");
    if (header) {
        header.onmousedown = dragMouseDown;
    }

    function dragMouseDown(e) {
        e.preventDefault();
        pos3 = e.clientX;
        pos4 = e.clientY;
        document.onmouseup = closeDragElement;
        document.onmousemove = elementDrag;
    }

    function elementDrag(e) {
        e.preventDefault();
        pos1 = pos3 - e.clientX;
        pos2 = pos4 - e.clientY;
        pos3 = e.clientX;
        pos4 = e.clientY;
        elmnt.style.top = (elmnt.offsetTop - pos2) + "px";
        elmnt.style.left = (elmnt.offsetLeft - pos1) + "px";
    }

    function closeDragElement() {
        document.onmouseup = null;
        document.onmousemove = null;
    }
}

// 處理算式輸入
function handleCalculation(input) {
    let val = input.value.trim();
    if (val.startsWith('=')) {
        try {
            // 移除 '=' 並計算剩下的數學表達式
            const expression = val.substring(1);
            // 僅允許數字與運算符號，防止惡意代碼
            if (/^[0-9+\-*/().\s]+$/.test(expression)) {
                input.value = new Function(`return ${expression}`)();
            }
        } catch (e) {
            console.error("算式格式錯誤");
        }
    }
    updateTotals(); // 計算完後更新上方小計
}

// 更新小計的功能
function updateTotals() {
    let total = 0;
    document.querySelectorAll('.amount-input').forEach(input => {
        total += Number(input.value) || 0;
    });
    document.getElementById('total-display').innerText = total;
}

function updateAllTotals() {
    let grandTotal = 0;

    // 取得所有金額欄位與掃描費欄位
    const amountInputs = document.querySelectorAll('.payee-amount');
    const scanInputs = document.querySelectorAll('.payee-scan');

    // 定義一個處理公式的內部函式
    const parseValue = (input) => {
        let rawValue = input.value.trim();
        if (!rawValue) return 0;

        // 如果以 = 開頭，嘗試計算
        if (rawValue.startsWith('=')) {
            try {
                // 移除 = 並計算結果 (例如 =500+20)
                // 為了安全與方便，使用簡單的 Function constructor 替代 eval
                const result = new Function(`return ${rawValue.substring(1)}`)();
                return parseFloat(result) || 0;
            } catch (e) {
                return 0; // 公式錯誤時暫不處理
            }
        }
        return parseFloat(rawValue) || 0;
    };

    // 加總金額
    amountInputs.forEach(input => grandTotal += parseValue(input));
    // 加總掃描費
    scanInputs.forEach(input => grandTotal += parseValue(input));

    // 更新到頁面上的總額欄位 (請確認你總額顯示區域的 ID 是什麼，這裡假設是 totalAmount)
    const totalDisplay = document.getElementById('totalAmount');
    if (totalDisplay) {
        totalDisplay.innerText = grandTotal.toLocaleString(); // 加上千分位
    }
}
// 監聽整個網頁的失焦事件
document.addEventListener('blur', function (e) {
    if (e.target.classList.contains('payee-amount') || e.target.classList.contains('payee-scan')) {
        let val = e.target.value.trim();
        if (val.startsWith('=')) {
            try {
                const formula = val.substring(1);
                const result = new Function(`return ${formula}`)();
                if (!isNaN(result) && isFinite(result)) {
                    e.target.value = result; // 欄位直接顯示計算結果
                    calculateTotal();        // 再次確保總額同步更新
                }
            } catch (err) {
                // 公式不完整則不處理
            }
        }
    }
}, true); // 使用 Capture 階段以正確監聽 blur 事件
