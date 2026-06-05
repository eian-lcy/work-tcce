// payee-registry.js (至尊全量下載與絕緣防護版)
console.log("%c[Payee-Registry] 啟動至尊級全量引擎與絕緣防護...", "color: #007bff; font-weight: bold;");

const SUPABASE_URL = 'https://eudnzrhqolvyijmgpjbz.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_r92k3Hbr-MGm_AugHi6SuQ_HFwUyLba';

let supabaseClient;
try {
    supabaseClient = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
} catch (e) {
    console.error("[Payee-Registry] Supabase 初始化失敗", e);
}

// 記憶體快取
window.allPayeeData = [];    
window.rawPayeeRegistry = {}; 
window.payeeRegistry = {};

// 核心功能：使用無限迴圈分頁下載，直到把 Supabase 幾千筆資料全部抽乾下載完畢
async function loadCompanyData() {
    if (!supabaseClient) return;
    console.log("[Payee-Registry] 開始進行突破上限之分批全量下載...");

    let allData = [];
    let from = 0;
    let to = 999;
    let hasMore = true;
    let loops = 0;

    // 💡 核心修正：用迴圈每次抓 1000 筆，直到某次抓取少於 1000 筆，代表抓完了
    while (hasMore && loops < 10) { // 設定 loops < 10 防範萬一死迴圈，上限 10000 筆
        const { data, error } = await supabaseClient
            .from('payee')
            .select('*')
            .range(from, to)
            .order('id', { ascending: false });

        if (error) {
            console.error(`[Payee-Registry] 在區間 ${from}-${to} 抓取失敗:`, error.message);
            break;
        }

        if (data && data.length > 0) {
            allData = allData.concat(data);
            console.log(`[Payee-Registry] 已成功下載第 ${from} ~ ${from + data.length - 1} 筆資料...`);
            
            if (data.length < 1000) {
                hasMore = false; // 抓到最後一頁了
            } else {
                from += 1000;
                to += 1000;
            }
        } else {
            hasMore = false;
        }
        loops++;
    }

    console.log(`[Payee-Registry] 🚀 雲端資料庫「徹底抽乾」下載完畢！共拿到 ${allData.length} 筆資料。`);

    window.allPayeeData = [];
    window.rawPayeeRegistry = {};
    window.payeeRegistry = {};

    // 進行源頭資料清洗
    allData.forEach(item => {
        if (!item.short_name || !item.full_name) return;

        const cleanArray = parseShortNamesToCleanArray(item.short_name);
        
        cleanArray.forEach(short => {
            if (short) {
                window.rawPayeeRegistry[short] = item.full_name;
                window.payeeRegistry[short] = item.full_name;
            }
        });

        window.allPayeeData.push({
            ...item,
            short_name: cleanArray.join(', ') // 儲存為統一純文字格式
        });
    });

    // 初始渲染最新 50 筆
    renderRegistryTable(window.allPayeeData.slice(0, 50));
    
    // 強制把搜尋事件霸佔過來
    bindSearchInputSecurely();
}

// 強效解析器
function parseShortNamesToCleanArray(shortNameStr) {
    let rawShort = String(shortNameStr).trim();
    let result = [];
    try {
        if (rawShort.startsWith('[') && rawShort.endsWith(']')) {
            const validJsonStr = rawShort.replace(/'/g, '"');
            result = JSON.parse(validJsonStr);
        } else if (rawShort.includes(',') || rawShort.includes('，') || rawShort.includes('、')) {
            result = rawShort.split(/,|，|、/);
        } else {
            result = [rawShort];
        }
    } catch (e) {
        const cleanStr = rawShort.replace(/[\[\]'"]/g, ''); 
        result = cleanStr.split(/,|，|、/);
    }
    return result.map(name => String(name).replace(/[\[\]'"]/g, '').trim()).filter(Boolean);
}

// 核心功能：繪製表格
function renderRegistryTable(itemsToRender) {
    const tbody = document.getElementById('registryTableBody') || document.querySelector('#payee-modal table tbody') || document.querySelector('.registry-table tbody');
    if (!tbody) return;

    tbody.innerHTML = ''; 
    const fragment = document.createDocumentFragment();

    itemsToRender.forEach(item => {
        const shortNamesArray = String(item.short_name).split(', ').filter(Boolean);
        const tr = document.createElement('tr');
        tr.style.borderBottom = "1px solid #dee2e6";

        const badgesHTML = shortNamesArray.map(name => 
            `<span style="background-color: #e7f5ff; color: #0d6efd; padding: 4px 8px; border-radius: 4px; font-size: 13px; font-weight: 500; margin-right: 4px; display: inline-block; margin-bottom: 4px;">${name}</span>`
        ).join('');

        tr.innerHTML = `
            <td style="padding: 12px 15px; vertical-align: middle; color: #333 !important; min-width: 120px;">${badgesHTML}</td>
            <td style="padding: 12px 15px; color: #333 !important; vertical-align: middle; font-weight: 500; word-break: break-all;">${item.full_name}</td>
            <td style="padding: 12px 15px; text-align: center; vertical-align: middle; width: 60px;">
                <button class="delete-btn" style="color: #dc3545; background: none; border: none; cursor: pointer; font-size: 16px; padding: 6px 10px; border-radius: 4px;">🗑️</button>
            </td>
        `;

        const deleteBtn = tr.querySelector('.delete-btn');
        if (deleteBtn) {
            deleteBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                handleDeleteCompany(item.id);
            });
        }

        fragment.appendChild(tr);
    });

    tbody.appendChild(fragment);
}

// 核心修正：強制安全隔離監聽搜尋輸入框，阻斷原本網頁舊函數的干擾
function bindSearchInputSecurely() {
    const input = document.getElementById('registrySearchInput');
    if (!input) return;

    // 拔除 HTML 上寫死的 onkeyup，改由我們全權控制
    input.removeAttribute('onkeyup');
    
    // 移除舊的監聽，重新綁定唯一的、具備斷路保護的監聽事件
    input.onkeyup = function(e) {
        if(e) {
            e.stopImmediatePropagation(); // 🛑 核心防禦：阻止網頁上其他腳本監聽此事件
            e.stopPropagation();
        }
        
        const keyword = input.value.toLowerCase().trim();

        if (keyword === '') {
            renderRegistryTable(window.allPayeeData.slice(0, 50));
            return;
        }

        // 精準過濾
        const filteredResults = window.allPayeeData.filter(item => {
            const shortMatch = String(item.short_name).toLowerCase().includes(keyword);
            const fullMatch = String(item.full_name).toLowerCase().includes(keyword);
            return shortMatch || fullMatch;
        });

        renderRegistryTable(filteredResults);
    };
    
    // 同步給 window 確保不報錯
    window.filterRegistryTable = input.onkeyup;
}

// 新增功能
async function handleAddCompany() {
    const shortNameInput = document.getElementById('newShortName');
    const fullNameInput = document.getElementById('newFullName');
    if (!shortNameInput || !fullNameInput) return;

    let rawShortText = shortNameInput.value.trim();
    const fullName = fullNameInput.value.trim();

    if (!rawShortText || !fullName) {
        alert('請填寫完整簡稱與全稱！');
        return;
    }

    if (rawShortText.includes(',') || rawShortText.includes('，') || rawShortText.includes('、')) {
        const items = rawShortText.split(/,|，|、/).map(i => i.trim()).filter(Boolean);
        rawShortText = `['${items.join("', '")}']`; 
    }

    const { error } = await supabaseClient
        .from('payee')
        .insert([{ short_name: rawShortText, full_name: fullName }]);

    if (error) {
        alert('雲端儲存失敗：' + error.message);
    } else {
        shortNameInput.value = '';
        fullNameInput.value = '';
        await loadCompanyData(); 
        alert('✨ 成功儲存至雲端對照表！');
    }
}

// 刪除功能
async function handleDeleteCompany(id) {
    if (!confirm('確定要刪除此筆對照資料嗎？')) return;

    const { error } = await supabaseClient
        .from('payee')
        .delete()
        .eq('id', id);

    if (error) {
        alert('刪除失敗：' + error.message);
    } else {
        await loadCompanyData(); 
    }
}

// 打開 Modal
async function openRegistryModal() {
    const modal = document.getElementById('payee-modal') || document.getElementById('registryModal');
    if (modal) modal.style.display = 'block';
    await loadCompanyData();
}

// 網頁載入啟動
document.addEventListener('DOMContentLoaded', loadCompanyData);
loadCompanyData();
