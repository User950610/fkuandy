const fileInput = document.getElementById("fileInput");
const dropZone = document.getElementById("dropZone");
const fileStatus = document.getElementById("fileStatus");
const dataPanel = document.getElementById("dataPanel");
const resultPanel = document.getElementById("resultPanel");
const sheetSelect = document.getElementById("sheetSelect");
const columnSelect = document.getElementById("columnSelect");
const previewTable = document.getElementById("previewTable");
const previewCount = document.getElementById("previewCount");
const statsGrid = document.getElementById("statsGrid");
const resultDescription = document.getElementById("resultDescription");
const recalculateButton = document.getElementById("recalculateButton");

let workbook = null;
let currentRows = [];
let currentHeaders = [];

fileInput.addEventListener("change", (event) => {
  const file = event.target.files[0];
  if (file) readExcelFile(file);
});

// 支援拖曳檔案到上傳區
["dragenter", "dragover"].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.add("dragover");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.remove("dragover");
  });
});

dropZone.addEventListener("drop", (event) => {
  const file = event.dataTransfer.files[0];
  if (file) readExcelFile(file);
});

sheetSelect.addEventListener("change", () => {
  showSelectedSheet();
});

columnSelect.addEventListener("change", calculateStatistics);
recalculateButton.addEventListener("click", calculateStatistics);

async function readExcelFile(file) {
  clearError();

  if (typeof XLSX === "undefined") {
    showError("Excel 讀取函式庫載入失敗。請確認網路連線後重新整理網頁。");
    return;
  }

  const allowedExtensions = /\.(xlsx|xls|xlsm|xlsb|csv)$/i;
  if (!allowedExtensions.test(file.name)) {
    showError("檔案格式不支援，請選擇 Excel 或 CSV 檔案。");
    return;
  }

  fileStatus.textContent = `正在讀取：${file.name}`;
  fileStatus.classList.remove("error");

  try {
    const buffer = await file.arrayBuffer();
    workbook = XLSX.read(buffer, { type: "array" });

    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
      throw new Error("檔案中找不到工作表。");
    }

    sheetSelect.innerHTML = "";
    workbook.SheetNames.forEach((sheetName) => {
      const option = document.createElement("option");
      option.value = sheetName;
      option.textContent = sheetName;
      sheetSelect.appendChild(option);
    });

    fileStatus.textContent = `已載入：${file.name}`;
    showSelectedSheet();
  } catch (error) {
    console.error(error);
    showError("無法讀取此檔案。請確認檔案沒有損壞，且第一列包含欄位名稱。");
    dataPanel.hidden = true;
    resultPanel.hidden = true;
  }
}

function showSelectedSheet() {
  if (!workbook) return;

  const sheet = workbook.Sheets[sheetSelect.value];
  // header: 1 會以二維陣列取得資料；defval 保留空白儲存格
  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: true
  });

  if (!matrix.length) {
    currentHeaders = [];
    currentRows = [];
    dataPanel.hidden = true;
    resultPanel.hidden = true;
    showError("這個工作表沒有資料。請選擇其他工作表。");
    return;
  }

  const rawHeaders = matrix[0].map((value, index) => {
    const name = String(value ?? "").trim();
    return name || `未命名欄位 ${index + 1}`;
  });

  // 避免重複欄位名稱造成欄位選擇混淆
  const seen = {};
  currentHeaders = rawHeaders.map((name) => {
    seen[name] = (seen[name] || 0) + 1;
    return seen[name] > 1 ? `${name} (${seen[name]})` : name;
  });

  currentRows = matrix.slice(1).map((row) => {
    return currentHeaders.map((_, index) => row[index] ?? "");
  });

  renderPreview();
  populateColumnOptions();
  dataPanel.hidden = false;
  resultPanel.hidden = false;
  clearError();
  calculateStatistics();
}

function renderPreview() {
  const thead = previewTable.querySelector("thead");
  const tbody = previewTable.querySelector("tbody");
  thead.replaceChildren();
  tbody.replaceChildren();

  const headerRow = document.createElement("tr");
  currentHeaders.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  const previewRows = currentRows.slice(0, 10);
  previewRows.forEach((row) => {
    const tr = document.createElement("tr");
    currentHeaders.forEach((_, index) => {
      const td = document.createElement("td");
      td.textContent = formatCell(row[index]);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  previewCount.textContent = `共 ${currentRows.length} 列資料`;
}

function formatCell(value) {
  if (value === null || value === undefined || value === "") return "—";
  if (value instanceof Date) return value.toLocaleDateString("zh-TW");
  return String(value);
}

function populateColumnOptions() {
  columnSelect.replaceChildren();

  currentHeaders.forEach((header, index) => {
    const option = document.createElement("option");
    option.value = String(index);
    option.textContent = header;
    columnSelect.appendChild(option);
  });

  // 優先選擇包含數值的第一個欄位
  const firstNumericIndex = currentHeaders.findIndex((_, index) => {
    return currentRows.some((row) => toNumber(row[index]) !== null);
  });

  if (firstNumericIndex >= 0) {
    columnSelect.value = String(firstNumericIndex);
  }
}

function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return null;

  const trimmed = value.trim();
  if (!trimmed) return null;

  // 僅接受一般數字格式及千分位逗號，不把任意文字強制轉成數字
  const normalized = trimmed.replace(/,/g, "");
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(normalized)) {
    return null;
  }

  const number = Number(normalized);
  return Number.isFinite(number) ? number : null;
}

function calculateStatistics() {
  if (!currentHeaders.length) return;

  const columnIndex = Number(columnSelect.value);
  const values = currentRows
    .map((row) => toNumber(row[columnIndex]))
    .filter((value) => value !== null);

  statsGrid.replaceChildren();

  if (values.length === 0) {
    resultDescription.textContent = `欄位：${currentHeaders[columnIndex] ?? ""}`;
    const message = document.createElement("p");
    message.className = "muted";
    message.textContent = "此欄位沒有可計算的數值。請選擇其他欄位。";
    statsGrid.appendChild(message);
    return;
  }

  const sorted = [...values].sort((a, b) => a - b);
  const n = values.length;
  const sum = values.reduce((total, value) => total + value, 0);
  const mean = sum / n;
  const median = n % 2 === 1
    ? sorted[Math.floor(n / 2)]
    : (sorted[n / 2 - 1] + sorted[n / 2]) / 2;

  const frequencies = new Map();
  values.forEach((value) => {
    frequencies.set(value, (frequencies.get(value) || 0) + 1);
  });

  const highestFrequency = Math.max(...frequencies.values());
  const modes = [...frequencies.entries()]
    .filter(([, frequency]) => frequency === highestFrequency)
    .map(([value]) => value)
    .sort((a, b) => a - b);

  const modeText = highestFrequency === 1
    ? "無眾數"
    : modes.map(formatNumber).join("、");

  const min = sorted[0];
  const max = sorted[n - 1];
  const range = max - min;
  const squaredDeviationSum = values.reduce(
    (total, value) => total + (value - mean) ** 2,
    0
  );
  const sampleVariance = n > 1 ? squaredDeviationSum / (n - 1) : null;
  const populationVariance = squaredDeviationSum / n;

  const statistics = [
    ["有效樣本數", String(n)],
    ["總和", formatNumber(sum)],
    ["平均數", formatNumber(mean)],
    ["中位數", formatNumber(median)],
    ["眾數", modeText],
    ["最大值", formatNumber(max)],
    ["最小值", formatNumber(min)],
    ["全距", formatNumber(range)],
    ["樣本變異數", sampleVariance === null ? "至少需要 2 筆" : formatNumber(sampleVariance)],
    ["樣本標準差", sampleVariance === null ? "至少需要 2 筆" : formatNumber(Math.sqrt(sampleVariance))],
    ["母體變異數", formatNumber(populationVariance)],
    ["母體標準差", formatNumber(Math.sqrt(populationVariance))]
  ];

  statistics.forEach(([label, value]) => {
    const card = document.createElement("div");
    card.className = "stat-card";

    const labelElement = document.createElement("div");
    labelElement.className = "stat-label";
    labelElement.textContent = label;

    const valueElement = document.createElement("div");
    valueElement.className = "stat-value";
    valueElement.textContent = value;

    card.append(labelElement, valueElement);
    statsGrid.appendChild(card);
  });

  resultDescription.textContent =
    `工作表：${sheetSelect.value}　｜　欄位：${currentHeaders[columnIndex]}　｜　有效數值：${n} 筆`;
}

function formatNumber(value) {
  if (!Number.isFinite(value)) return "無法計算";
  return new Intl.NumberFormat("zh-TW", {
    maximumFractionDigits: 6
  }).format(value);
}

function showError(message) {
  fileStatus.textContent = message;
  fileStatus.classList.add("error");
}

function clearError() {
  fileStatus.classList.remove("error");
}
