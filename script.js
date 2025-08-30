// ===== 倉別對應 Apps Script URL =====
const scriptURLs = {
"TAO1": "https://script.google.com/macros/s/AKfycbyMdIIg_BjzqOWvn4Mzfgo7hCOx8nYNry1Xg-EFpEo9Gg8ECD8x0NRPwFvQZj4agM6u/exec",
"TAO3": "https://script.google.com/macros/s/AKfycbyCXqK6d8tu91LCfiK5Iwn175GwC0b8xS3to20v3lOvs41r9isiUCnr87KOKBcFI0Q/exec",
"TAO4": "https://script.google.com/macros/s/AKfycbxOcL5pq_bO8mUn4ANSD2CupGQyz-0aqqHJepyui_0TjWQ7ije8dn9FB6FIxObu-Ro/exec",
"TAO5": "https://script.google.com/macros/s/AKfycbx3PC0hwH4_OdUxC8zMgut_ZxA8KtERvu8IkH-xQLCK5-khbn7MMu6w3xLLmZxjuvOs/exec",
"TAO6": "https://script.google.com/macros/s/AKfycbxGoISWpRxJ3Da2wYSFkPuM_PjHyidl1l-Pe2xWR0xDO1oku6M63wvRwOL4FsOOUgs/exec",
"TAO7": "https://script.google.com/macros/s/AKfycbzGwJMTeYTBq_9ZgH_oGroVC_isk-qnCxFRE0DxmwzTabZXoezw4Gexa9WGp-94V1I/exec",
"TAO10": "https://script.google.com/macros/s/AKfycbyZUMsx-sfV_F5KroTtBfWK8G9CPs-MGF6c-fj7gIj3hH0gJr8lKGPMycEgzB1_2vnk/exec"
};

// ===== 倉別對應 Google 試算表 =====
const spreadsheetLinks = {
"TAO1": "https://docs.google.com/spreadsheets/d/1_bhGQdx0YH7lsqPFEq5___6_Nwq_gbelJmIHv0bmaIE/edit",
"TAO3": "https://docs.google.com/spreadsheets/d/1cffI2jIVZA1uSiAyaLLXXgPzDByhy87xznaN85O7wEE/edit",
"TAO4": "https://docs.google.com/spreadsheets/d/1tVxQbV0298fn2OXWAF0UqZa7FLbypsatciatxs4YVTU/edit",
"TAO5": "https://docs.google.com/spreadsheets/d/1jzVXC6gt36hJtlUHoxtTzZLMNj4EtTsd4k8eNB1bdiA/edit",
"TAO6": "https://docs.google.com/spreadsheets/d/1wwPLSLjl2abfM_OMdTNI9PoiPKo3waCV_y0wmx2DxAE/edit",
"TAO7": "https://docs.google.com/spreadsheets/d/16nGCqRO8DYDm0PbXFbdt-fiEFZCXxXjlOWjKU67p4LY/edit",
"TAO10": "https://docs.google.com/spreadsheets/d/1y0w49xdFlHvcVtgtG8fq6zdrF26y8j7HMFh5ujzUyR4/edit"
};

// ===== 自動判斷倉別設定 =====
const WAREHOUSE_PRIORITY = ['TAO10','TAO1','TAO3','TAO4','TAO5','TAO6','TAO7'];
const SHEET_KEYWORDS = ['出勤','班表','名冊','名單','在職','員工'];
const NAME_CACHE_KEY = 'nameWarehouseSheetCache.v1';

// 快取
function loadCache() {
try { return JSON.parse(localStorage.getItem(NAME_CACHE_KEY) || '{}'); }
catch { return {}; }
}
function saveCache(map) {
localStorage.setItem(NAME_CACHE_KEY, JSON.stringify(map || {}));
}

// 取得倉別的資料表
async function getSheets(warehouse) {
const url = scriptURLs[warehouse];
const res = await fetch(`${url}?mode=getSheets`);
if (!res.ok) throw new Error("getSheets 失敗");
return res.json();
}

// 測試某表是否有這個人
async function testSheetHasName(warehouse, sheet, name) {
const url = `${scriptURLs[warehouse]}?sheet=${encodeURIComponent(sheet)}&name=${encodeURIComponent(name)}&from=glitch`;
const res = await fetch(url);
const text = await res.text();
return !text.includes("查無資料") && !text.includes("查無此工作表") && !text.includes("查無此分頁設定");
}

// 幫 select 填表單
async function loadSheets(warehouse, preselect) {
const sheetSel = document.getElementById("sheet");
sheetSel.innerHTML = '<option>載入中...</option>';
try {
const names = await getSheets(warehouse);
sheetSel.innerHTML = '<option value="">請選擇資料表</option>';
names.forEach(n => {
const opt = document.createElement("option");
opt.value = n; opt.text = n;
sheetSel.appendChild(opt);
});
if (preselect && names.includes(preselect)) {
sheetSel.value = preselect;
} else {
const picked = names.find(n => SHEET_KEYWORDS.some(k => n.includes(k)));
if (picked) sheetSel.value = picked;
}
sheetSel.disabled = false;
} catch {
sheetSel.innerHTML = '<option value="">載入失敗</option>';
}
}

// ⭐ 輸入姓名 → 自動帶倉別
async function autoPickByName() {
const name = document.getElementById("name").value.trim();
const hint = document.getElementById("nameHint");
if (!name) return;

const cache = loadCache();
if (cache[name]) {
document.getElementById("warehouse").value = cache[name].warehouse;
await loadSheets(cache[name].warehouse, cache[name].sheet);
hint.textContent = `已自動帶入倉別：${cache[name].warehouse}`;
return;
}

// 沒快取 → 順序掃倉別
for (const wh of WAREHOUSE_PRIORITY) {
try {
const sheets = await getSheets(wh);
const candidate = sheets.find(s => SHEET_KEYWORDS.some(k => s.includes(k))) || sheets[0];
if (!candidate) continue;
const ok = await testSheetHasName(wh, candidate, name);
if (ok) {
document.getElementById("warehouse").value = wh;
await loadSheets(wh, candidate);
cache[name] = { warehouse: wh, sheet: candidate };
saveCache(cache);
hint.textContent = `已自動帶入倉別：${wh}`;
return;
}
} catch (e) { console.warn("檢查失敗", wh, e); }
}
hint.textContent = "❌ 沒找到對應倉別，請檢查姓名是否正確";
}

// 綁定事件：姓名輸入完就觸發
document.getElementById("name").addEventListener("blur", autoPickByName);
document.getElementById("name").addEventListener("keydown", e => {
if (e.key === "Enter") autoPickByName();
});

// 查詢按鈕
document.getElementById("queryBtn").addEventListener("click", function () {
const warehouse = document.getElementById("warehouse").value;
const sheet = document.getElementById("sheet").value;
const name = document.getElementById("name").value.trim();
const phone = document.getElementById("phone").value.trim();

if (!warehouse || !sheet || !name) {
alert("請輸入姓名，系統會自動帶入倉別與資料表！");
return;
}

let url = `${scriptURLs[warehouse]}?sheet=${encodeURIComponent(sheet)}&name=${encodeURIComponent(name)}&from=glitch`;
if (phone) url += `&phone=${encodeURIComponent(phone)}`;

// 存快取
const cache = loadCache();
cache[name] = { warehouse, sheet };
saveCache(cache);

window.open(url, "_blank");
});

// 開啟試算表按鈕
document.getElementById("openSheetBtn").addEventListener("click", function () {
const warehouse = document.getElementById("spreadsheetWarehouse").value;
const link = spreadsheetLinks[warehouse];
if (link) window.open(link, "_blank");
else alert("請選擇倉別！");
});
