// === Google Sheets API Key ===
const API_KEY = "AIzaSyBfhunDmZjbbWxZKYqEjCPe5qZLU_mbtUI";

// === 每倉對應的試算表 ID ===
const SHEET_IDS = {
"TAO1": "1_bhGQdx0YH7lsqPFEq5___6_Nwq_gbelJmIHv0bmaIE",
"TAO3": "1cffI2jIVZA1uSiAyaLLXXgPzDByhy87xznaN85O7wEE",
"TAO4": "1tVxQbV0298fn2OXWAF0UqZa7FLbypsatciatxs4YVTU",
"TAO5": "1jzVXC6gt36hJtlUHoxtTzZLMNj4EtTsd4k8eNB1bdiA",
"TAO6": "1wwPLSLjl2abfM_OMdTNI9PoiPKo3waCV_y0wmx2DxAE",
"TAO7": "16nGCqRO8DYDm0PbXFbdt-fiEFZCXxXjlOWjKU67p4LY",
"TAO10": "1y0w49xdFlHvcVtgtG8fq6zdrF26y8j7HMFh5ujzUyR4"
};

// === DOM 物件 ===
const warehouseSel = document.getElementById("warehouse");
const sheetSel = document.getElementById("sheet");
const nameInput = document.getElementById("name");
const phoneInput = document.getElementById("phone");
const queryBtn = document.getElementById("queryBtn");

// === 取得分頁清單 ===
async function loadSheets(warehouse) {
const id = SHEET_IDS[warehouse];
if (!id) {
sheetSel.innerHTML = `<option value="">找不到此倉別 ID</option>`;
return;
}
sheetSel.innerHTML = `<option value="">載入中...</option>`;

try {
const url = `https://sheets.googleapis.com/v4/spreadsheets/${id}?fields=sheets(properties(title))&key=${API_KEY}`;
const res = await fetch(url);
if (!res.ok) throw new Error("讀取分頁失敗");
const data = await res.json();
const titles = (data.sheets || []).map(s => s.properties.title);

if (!titles.length) {
sheetSel.innerHTML = `<option value="">此試算表沒有分頁</option>`;
return;
}

sheetSel.innerHTML = `<option value="">請選擇分頁</option>`;
titles.forEach(t => {
const opt = document.createElement("option");
opt.value = t;
opt.text = t;
sheetSel.appendChild(opt);
});
} catch (err) {
console.error(err);
sheetSel.innerHTML = `<option value="">載入失敗</option>`;
}
}

// === 切換倉別時載入分頁 ===
warehouseSel.addEventListener("change", () => {
const wh = warehouseSel.value;
if (wh) loadSheets(wh);
});

// === 查詢（這裡先簡化：直接開 Google Sheet 的該分頁） ===
queryBtn.addEventListener("click", () => {
const warehouse = warehouseSel.value;
const sheet = sheetSel.value;
const name = nameInput.value.trim();
const phone = phoneInput.value.trim();

if (!warehouse || !sheet || !name) {
alert("請選擇倉別、分頁並輸入姓名！");
return;
}

const id = SHEET_IDS[warehouse];
const url = `https://docs.google.com/spreadsheets/d/${id}/edit`;
window.open(url, "_blank");
});
