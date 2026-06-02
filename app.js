const state = {
  sheetIndex: 0,
  query: "",
  selectedRow: null,
};

const workbook = window.WORKBOOK_DATA;
const tabs = document.getElementById("tabs");
const table = document.getElementById("sheet-table");
const search = document.getElementById("search");
const count = document.getElementById("count");
const editStatus = document.getElementById("edit-status");
const deleteRowButton = document.getElementById("delete-row");
const storageKey = `logistics-tracker-edits:${workbook.title}:${workbook.generatedAt}`;

document.getElementById("workbook-title").textContent = workbook.title;

function cloneCell(cell = {}) {
  return {
    ...cell,
    style: cell.style ? { ...cell.style } : undefined,
  };
}

const originalRows = workbook.sheets.map(sheet => sheet.rows.map(row => row.map(cloneCell)));

function ensureSheetShape(sheet) {
  sheet.rows.forEach(row => {
    while (row.length < sheet.colCount) row.push({ text: "" });
  });
}

function loadSavedEdits() {
  try {
    const saved = JSON.parse(localStorage.getItem(storageKey) || "null");
    if (!saved || !Array.isArray(saved.sheets)) return;
    saved.sheets.forEach((savedSheet, sheetIndex) => {
      const sheet = workbook.sheets[sheetIndex];
      if (!sheet || !Array.isArray(savedSheet.rows)) return;
      sheet.rows = savedSheet.rows.map(row => row.map(cloneCell));
      sheet.rowCount = sheet.rows.length;
      ensureSheetShape(sheet);
    });
  } catch {
    localStorage.removeItem(storageKey);
  }
}

function saveEdits() {
  const payload = {
    savedAt: new Date().toISOString(),
    sheets: workbook.sheets.map(sheet => ({
      name: sheet.name,
      rows: sheet.rows,
    })),
  };
  localStorage.setItem(storageKey, JSON.stringify(payload));
  updateMeta();
}

function escapeText(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function escapeXml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function safeSheetName(name) {
  return String(name || "Sheet").replace(/[\[\]:*?/\\]/g, " ").slice(0, 31);
}

function styleAttr(style) {
  if (!style) return "";
  const rules = [];
  if (style.fill) rules.push(`background:${style.fill}`);
  if (style.color) rules.push(`color:${style.color}`);
  if (style.bold) rules.push("font-weight:700");
  if (style.italic) rules.push("font-style:italic");
  if (style.align) rules.push(`text-align:${style.align}`);
  if (style.valign) rules.push(`vertical-align:${style.valign}`);
  return rules.length ? ` style="${rules.join(";")}"` : "";
}

function getOriginalText(sheetIndex, rowIndex, colIndex) {
  return originalRows[sheetIndex]?.[rowIndex]?.[colIndex]?.text ?? "";
}

function isChanged(sheetIndex, rowIndex, colIndex) {
  const sheet = workbook.sheets[sheetIndex];
  return (sheet.rows[rowIndex]?.[colIndex]?.text ?? "") !== getOriginalText(sheetIndex, rowIndex, colIndex);
}

function countChangedCells() {
  let changed = 0;
  workbook.sheets.forEach((sheet, sheetIndex) => {
    sheet.rows.forEach((row, rowIndex) => {
      row.forEach((_, colIndex) => {
        if (isChanged(sheetIndex, rowIndex, colIndex)) changed += 1;
      });
    });
  });
  return changed;
}

function updateMeta() {
  const changed = countChangedCells();
  const saved = localStorage.getItem(storageKey) ? "已自动保存在本浏览器" : "未保存本地改动";
  document.getElementById("meta").textContent = `${workbook.sheets.length} 个工作表 · ${new Date(workbook.generatedAt).toLocaleString("zh-CN")}`;
  editStatus.textContent = changed ? `${changed} 个单元格已修改 · ${saved}` : "没有未还原的修改";
}

function renderTabs() {
  tabs.innerHTML = workbook.sheets.map((sheet, index) => {
    const active = index === state.sheetIndex ? " active" : "";
    return `<button class="tab${active}" data-index="${index}" type="button">${escapeText(sheet.name)}</button>`;
  }).join("");
}

function rowMatches(row, query) {
  if (!query) return true;
  return row.some(cell => String(cell.text ?? cell.value ?? "").toLowerCase().includes(query));
}

function renderTable() {
  const sheet = workbook.sheets[state.sheetIndex];
  ensureSheetShape(sheet);
  const query = state.query.trim().toLowerCase();
  const rows = sheet.rows
    .map((row, index) => ({ row, index }))
    .filter(({ row }) => rowMatches(row, query));

  const colHeaders = ["<th class=\"corner\"></th>", ...sheet.columns.map((col, index) => {
    const width = sheet.widths[index] ? Math.min(Math.max(sheet.widths[index] * 8, 72), 220) : 120;
    return `<th class="col-head" style="min-width:${width}px">${col}</th>`;
  })].join("");

  const body = rows.map(({ row, index }) => {
    const selected = state.selectedRow === index ? " selected-row" : "";
    const cells = row.map((cell, colIndex) => {
      const classes = ["cell"];
      if (!cell.text) classes.push("empty");
      if (isChanged(state.sheetIndex, index, colIndex)) classes.push("changed");
      const title = cell.formula ? ` title="${escapeText(cell.formula)}"` : "";
      return `<td class="${classes.join(" ")}" tabindex="0" data-row="${index}" data-col="${colIndex}"${title}${styleAttr(cell.style)}>${escapeText(cell.text)}</td>`;
    }).join("");
    return `<tr class="${selected}"><th class="row-head" data-row="${index}" title="点击选择整行">${index + 1}</th>${cells}</tr>`;
  }).join("");

  table.innerHTML = `<thead><tr>${colHeaders}</tr></thead><tbody>${body}</tbody>`;
  count.textContent = `${rows.length} / ${sheet.rowCount} 行`;
  deleteRowButton.disabled = state.selectedRow === null;
  updateMeta();
}

function render() {
  renderTabs();
  renderTable();
}

function updateCell(rowIndex, colIndex, text) {
  const sheet = workbook.sheets[state.sheetIndex];
  const cell = sheet.rows[rowIndex][colIndex] || { text: "" };
  cell.text = text;
  cell.value = text;
  if (cell.formula && text !== getOriginalText(state.sheetIndex, rowIndex, colIndex)) {
    delete cell.formula;
  }
  sheet.rows[rowIndex][colIndex] = cell;
  saveEdits();
}

function startCellEdit(cellElement) {
  if (cellElement.classList.contains("editing")) return;
  const rowIndex = Number(cellElement.dataset.row);
  const colIndex = Number(cellElement.dataset.col);
  const sheet = workbook.sheets[state.sheetIndex];
  const current = sheet.rows[rowIndex]?.[colIndex]?.text ?? "";
  cellElement.classList.add("editing");
  cellElement.innerHTML = "";

  const editor = document.createElement("textarea");
  editor.className = "cell-editor";
  editor.value = current;
  editor.rows = Math.min(6, Math.max(1, String(current).split(/\r\n|\r|\n/).length));
  cellElement.appendChild(editor);
  editor.focus();
  editor.select();

  const commit = () => {
    if (!cellElement.isConnected) return;
    updateCell(rowIndex, colIndex, editor.value.trim());
    renderTable();
  };

  editor.addEventListener("blur", commit, { once: true });
  editor.addEventListener("keydown", event => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      editor.blur();
    }
    if (event.key === "Escape") {
      event.preventDefault();
      renderTable();
    }
  });
}

function downloadBlob(content, fileName, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function exportCsv() {
  const sheet = workbook.sheets[state.sheetIndex];
  const csv = sheet.rows.map(row => row.map(cell => {
    const value = String(cell.text ?? "");
    return /[",\r\n]/.test(value) ? `"${value.replaceAll('"', '""')}"` : value;
  }).join(",")).join("\r\n");
  downloadBlob(`\ufeff${csv}`, `${sheet.name}.csv`, "text/csv;charset=utf-8");
}

function exportExcelXml() {
  const worksheets = workbook.sheets.map(sheet => {
    const rows = sheet.rows.map(row => {
      const cells = row.map(cell => `<Cell><Data ss:Type="String">${escapeXml(cell.text ?? "")}</Data></Cell>`).join("");
      return `<Row>${cells}</Row>`;
    }).join("");
    return `<Worksheet ss:Name="${escapeXml(safeSheetName(sheet.name))}"><Table>${rows}</Table></Worksheet>`;
  }).join("");
  const xml = `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
${worksheets}
</Workbook>`;
  downloadBlob(xml, `${workbook.title}-编辑版.xls`, "application/vnd.ms-excel;charset=utf-8");
}

function addRow() {
  const sheet = workbook.sheets[state.sheetIndex];
  const row = Array.from({ length: sheet.colCount }, () => ({ text: "" }));
  sheet.rows.push(row);
  sheet.rowCount = sheet.rows.length;
  state.selectedRow = sheet.rows.length - 1;
  saveEdits();
  renderTable();
}

function deleteSelectedRow() {
  if (state.selectedRow === null) return;
  const sheet = workbook.sheets[state.sheetIndex];
  sheet.rows.splice(state.selectedRow, 1);
  sheet.rowCount = sheet.rows.length;
  state.selectedRow = null;
  saveEdits();
  renderTable();
}

function resetSheet() {
  const sheet = workbook.sheets[state.sheetIndex];
  sheet.rows = originalRows[state.sheetIndex].map(row => row.map(cloneCell));
  sheet.rowCount = sheet.rows.length;
  state.selectedRow = null;
  saveEdits();
  renderTable();
}

function resetAll() {
  workbook.sheets.forEach((sheet, index) => {
    sheet.rows = originalRows[index].map(row => row.map(cloneCell));
    sheet.rowCount = sheet.rows.length;
  });
  state.selectedRow = null;
  localStorage.removeItem(storageKey);
  render();
}

tabs.addEventListener("click", event => {
  const button = event.target.closest("button[data-index]");
  if (!button) return;
  state.sheetIndex = Number(button.dataset.index);
  state.query = "";
  state.selectedRow = null;
  search.value = "";
  render();
});

search.addEventListener("input", event => {
  state.query = event.target.value;
  state.selectedRow = null;
  renderTable();
});

table.addEventListener("click", event => {
  const rowHead = event.target.closest(".row-head[data-row]");
  if (rowHead) {
    state.selectedRow = Number(rowHead.dataset.row);
    renderTable();
    return;
  }
  const cell = event.target.closest("td.cell[data-row][data-col]");
  if (cell) startCellEdit(cell);
});

document.getElementById("export-xls").addEventListener("click", exportExcelXml);
document.getElementById("export-csv").addEventListener("click", exportCsv);
document.getElementById("add-row").addEventListener("click", addRow);
document.getElementById("delete-row").addEventListener("click", deleteSelectedRow);
document.getElementById("reset-sheet").addEventListener("click", resetSheet);
document.getElementById("reset-all").addEventListener("click", resetAll);

loadSavedEdits();
render();
