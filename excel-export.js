// excel-export.js

function syncUiStateToDataModel() {
    const rows = document.querySelectorAll('.calc-row');
    rows.forEach(tr => {
        const batchId = parseInt(tr.getAttribute('data-batch-id'));
        const payeeIndex = parseInt(tr.getAttribute('data-payee-index'));
        const type = tr.getAttribute('data-type');
        
        const chkTax = tr.querySelector('.chk-tax');
        const chkHealth = tr.querySelector('.chk-health');
        
        const isTax = chkTax ? chkTax.checked : false;
        const isHealth = chkHealth ? chkHealth.checked : false;

        const batch = allBatches.find(b => b.id === batchId);
        if (batch && batch.payees[payeeIndex]) {
            if (type === 'income') {
                batch.payees[payeeIndex].isTax = isTax;
                batch.payees[payeeIndex].isHealth = isHealth;
            }
        }
    });
}

function exportToExcel() {
    const dateStr = document.getElementById('reportDateSelect').value;
    if (!dateStr) { alert("無資料可匯出"); return; }

    syncUiStateToDataModel();

    const dateObj = new Date(dateStr);
    const rocYear = dateObj.getFullYear() - 1911;
    const titleText = `${rocYear}年${dateObj.getMonth() + 1}月${dateObj.getDate()}日撥款`;

    const ws_data = [];
    ws_data.push([titleText, null, null, null, null, null]);

    const targetBatches = allBatches.filter(b => b.date === dateStr);
    const peopleMap = {};

    targetBatches.forEach(batch => {
        batch.payees.forEach(p => {
            const name = p.name ? p.name.trim() : "未知";
            if (p.amount !== 0) {
                if (!peopleMap[name]) peopleMap[name] = [];
                let tax = p.isTax ? Math.round(p.amount * 0.1) : 0;
                let health = p.isHealth ? Math.round(p.amount * 0.0211) : 0;

                peopleMap[name].push({
                    usage: p.incomeUsage,
                    listed: p.amount,
                    tax: tax,
                    health: health,
                    date: batch.date,
                    isTax: p.isTax,
                    isHealth: p.isHealth
                });
            }

            if (p.scanFee > 0) {
                if (!peopleMap[name]) peopleMap[name] = [];
                peopleMap[name].push({
                    usage: p.scanUsage,
                    listed: -p.scanFee,
                    tax: 0,
                    health: 0,
                    date: batch.date,
                    isTax: false,
                    isHealth: false
                });
            }
        });
    });

    const merges = [];
    merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } });

    let currentDataIndex = 0; 
    const subtotalCellsD = [];
    const subtotalCellsE = [];
    const subtotalCellsF = [];

    for (let [name, entries] of Object.entries(peopleMap)) {
        ws_data.push(["明細資料", "金額", "用途", "案件金額", "10%扣繳", "健保費"]);
        currentDataIndex++;

        const startDataIndex = currentDataIndex + 1; 
        const startExcelRow = startDataIndex + 1;    

        entries.forEach((e, idx) => {
            currentDataIndex++; 
            const excelRow = currentDataIndex + 1; 

            let taxCell = e.isTax ? { t: 'n', f: `ROUND(D${excelRow}*0.1, 0)`, v: e.tax } : 0;
            let healthCell = e.isHealth ? { t: 'n', f: `ROUND(D${excelRow}*0.0211, 0)`, v: e.health } : 0;

            if (idx === 0) {
                ws_data.push([name, null, e.usage, e.listed, taxCell, healthCell]);
            } else {
                ws_data.push([null, null, e.usage, e.listed, taxCell, healthCell]);
            }
        });

        const endExcelRow = currentDataIndex + 1; 

        ws_data[startDataIndex][1] = { 
            t: 'n', 
            f: `SUM(D${startExcelRow}:D${endExcelRow})-SUM(E${startExcelRow}:F${endExcelRow})` 
        };

        if (entries.length > 1) {
            merges.push({ s: { r: startDataIndex, c: 0 }, e: { r: currentDataIndex, c: 0 } });
            merges.push({ s: { r: startDataIndex, c: 1 }, e: { r: currentDataIndex, c: 1 } });
        }

        ws_data.push([
            null, null, "合計", 
            { t: 'n', f: `SUM(D${startExcelRow}:D${endExcelRow})` }, 
            { t: 'n', f: `SUM(E${startExcelRow}:E${endExcelRow})` }, 
            { t: 'n', f: `SUM(F${startExcelRow}:F${endExcelRow})` }
        ]);
        currentDataIndex++; 

        subtotalCellsD.push(`D${currentDataIndex + 1}`);
        subtotalCellsE.push(`E${currentDataIndex + 1}`);
        subtotalCellsF.push(`F${currentDataIndex + 1}`);
    }

    const grandFormulaD = subtotalCellsD.length > 0 ? subtotalCellsD.join('+') : "0";
    const grandFormulaE = subtotalCellsE.length > 0 ? subtotalCellsE.join('+') : "0";
    const grandFormulaF = subtotalCellsF.length > 0 ? subtotalCellsF.join('+') : "0";

    ws_data.push([
        null, null, "總計", 
        { t: 'n', f: grandFormulaD }, 
        { t: 'n', f: grandFormulaE }, 
        { t: 'n', f: grandFormulaF }
    ]);

    const ws = XLSX.utils.aoa_to_sheet(ws_data);

    // --- 樣式設定 (包含空格修正) ---
    const fontStyle16 = { name: "標楷體", sz: 16, bold: true };
    const fontStyle12 = { name: "標楷體", sz: 12, bold: true };
    const currencyFmt = '"$"#,##0_ ;"-$"#,##0_ ; ';
    const thinBorder = {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" }
    };

    ws['!cols'] = [{ wch: 20.13 }, { wch: 21.25 }, { wch: 43.38 }, { wch: 17.5 }, { wch: 14.25 }, { wch: 11.88 }];
    const range = XLSX.utils.decode_range(ws['!ref']);

    for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = 0; C <= 5; ++C) {
            const cell_address = XLSX.utils.encode_cell({ r: R, c: C });
            
            // 重要：判斷是否為被合併掉的「配角格」
            const isMergedSecondary = merges.some(m => 
                R >= m.s.r && R <= m.e.r && 
                C >= m.s.c && C <= m.e.c && 
                !(R === m.s.r && C === m.s.c)
            );

            // 如果是配角格，不要賦值，避免複製時出現空格
            if (isMergedSecondary) continue;

            if (!ws[cell_address]) ws[cell_address] = { t: 's', v: "" };
            if (!ws[cell_address].s) ws[cell_address].s = {};

            ws[cell_address].s.border = thinBorder;
            ws[cell_address].s.font = fontStyle16;
            ws[cell_address].s.alignment = { vertical: "center" };

            const val = ws[cell_address].v;
            if (C === 0) ws[cell_address].s.alignment.horizontal = "center";
            else if (C === 1) {
                ws[cell_address].s.alignment.horizontal = "center";
                if (ws[cell_address].t === 'n' || ws[cell_address].f) ws[cell_address].z = currencyFmt;
            }
            else if (C === 2) {
                ws[cell_address].s.alignment.wrapText = true;
                if (val !== "用途" && val !== "合計" && val !== "總計" && val !== "") ws[cell_address].s.font = fontStyle12;
                if (val === "合計" || val === "總計") ws[cell_address].s.alignment.horizontal = "right";
            }
            else if (C >= 3) {
                if (ws[cell_address].t === 'n' || ws[cell_address].f) ws[cell_address].z = currencyFmt;
                ws[cell_address].s.alignment.horizontal = "right";
            }
        }
    }

    ws['!merges'] = merges;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "撥款明細");

    // --- Sheet 2: 扣繳稅額表 ---
    const taxDataRaw = [];
    for (let [name, entries] of Object.entries(peopleMap)) {
        entries.forEach(e => {
            if (e.tax > 0) {
                const dObj = new Date(e.date);
                const dateText = `${dObj.getMonth() + 1}月${dObj.getDate()}日`;
                taxDataRaw.push({
                    usage: e.usage,
                    amount: e.listed,
                    tax: e.tax,
                    subtracted: e.listed - e.tax,
                    dateText: dateText,
                    name: name
                });
            }
        });
    }
    taxDataRaw.sort((a, b) => a.name.localeCompare(b.name, "zh-Hant"));
    const ws_tax_data = [];
    ws_tax_data.push(["用途", "金額", "10%扣繳", "小計", "撥款日期", "受款人姓名"]);
    taxDataRaw.forEach(item => {
        ws_tax_data.push([item.usage, item.amount, item.tax, item.subtracted, item.dateText, item.name]);
    });
    const ws_tax = XLSX.utils.aoa_to_sheet(ws_tax_data);
    ws_tax['!cols'] = [{ wch: 40 }, { wch: 15 }, { wch: 12 }, { wch: 15 }, { wch: 15 }, { wch: 20 }];

    // --- Sheet 3: 二代健保 ---
    const healthMap = {};
    for (let [name, entries] of Object.entries(peopleMap)) {
        entries.forEach(e => {
            if (e.health > 0) {
                if (!healthMap[name]) healthMap[name] = { totalAmount: 0, totalHealth: 0 };
                healthMap[name].totalAmount += e.listed;
                healthMap[name].totalHealth += e.health;
            }
        });
    }
    const ws_health_data = [["收款人", "金額加總", "健保費加總"]];
    for (let [name, val] of Object.entries(healthMap)) {
        ws_health_data.push([name, val.totalAmount, val.totalHealth]);
    }
    const ws_health = XLSX.utils.aoa_to_sheet(ws_health_data);
    ws_health['!cols'] = [{ wch: 20 }, { wch: 20 }, { wch: 20 }];

    // --- Sheet 4: Raw Data ---
    const rawData = [["BatchID", "Date", "Cat1", "Cat2", "RawCaseNo", "Dept", "ExtraNote", "PayeeName", "Amount", "ScanFee", "IncomeUsage", "ScanUsage", "IsTax", "IsHealth"]];
    allBatches.forEach(batch => {
        batch.payees.forEach(p => {
            rawData.push([
                batch.id, batch.date, batch.cat1, batch.cat2, batch.rawCaseNo || "", batch.dept || "", batch.extraNote || "",
                p.name, p.amount, p.scanFee, p.incomeUsage, p.scanUsage,
                p.isTax ? "TRUE" : "FALSE", p.isHealth ? "TRUE" : "FALSE"
            ]);
        });
    });
    const wsRaw = XLSX.utils.aoa_to_sheet(rawData);

    // --- 組合與存檔 ---
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "撥款明細");
    XLSX.utils.book_append_sheet(wb, ws_tax, "扣繳稅額表");
    XLSX.utils.book_append_sheet(wb, ws_health, "二代健保表");
    XLSX.utils.book_append_sheet(wb, wsRaw, "RawData");

    const month = String(dateObj.getMonth() + 1).padStart(2, '0');
    const day = String(dateObj.getDate()).padStart(2, '0');
    let defaultFileName = `${rocYear}${month}${day}撥款明細.xlsx`;
    let userFileName = prompt("請確認檔案名稱 (直接按 Enter 即可存檔)：", defaultFileName);

    if (userFileName) {
        if (!userFileName.toLowerCase().endsWith(".xlsx")) {
            userFileName += ".xlsx";
        }
        XLSX.writeFile(wb, userFileName);
    }
}
