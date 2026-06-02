const state = {
  sheetIndex: 0,
  query: "",
  selectedRow: null,
  selectedRange: null,
};

const history = {
  undo: [],
  redo: [],
  limit: 60,
};

const workbook = window.WORKBOOK_DATA;
const tabs = document.getElementById("tabs");
const table = document.getElementById("sheet-table");
const search = document.getElementById("search");
const count = document.getElementById("count");
const editStatus = document.getElementById("edit-status");
const undoButton = document.getElementById("undo");
const redoButton = document.getElementById("redo");
const deleteRowButton = document.getElementById("delete-row");
const mergeButton = document.getElementById("merge-cells");
const unmergeButton = document.getElementById("unmerge-cells");
const storageKey = `logistics-tracker-edits:${workbook.title}:${workbook.generatedAt}`;

document.getElementById("workbook-title").textContent = workbook.title;

function cloneCell(cell = {}) {
  return {
    ...cell,
    style: cell.style ? { ...cell.style } : undefined,
  };
}

function cloneRows(rows) {
  return rows.map(row => row.map(cloneCell));
}

function cloneSheetState(sheet) {
  return {
    rows: cloneRows(sheet.rows),
    merges: [...(sheet.merges || [])],
    rowCount: sheet.rowCount,
  };
}

const originalSheets = workbook.sheets.map(cloneSheetState);

function ensureSheetShape(sheet) {
  if (!Array.isArray(sheet.merges)) sheet.merges = [];
  sheet.rows.forEach(row => {
    while (row.length < sheet.colCount) row.push({ text: "" });
  });
  sheet.rowCount = sheet.rows.length;
}

function snapshot() {
  return workbook.sheets.map(cloneSheetState);
}

function restoreSnapshot(saved) {
  saved.forEach((savedSheet, index) => {
    const sheet = workbook.sheets[index];
    if (!sheet) return;
    sheet.rows = cloneRows(savedSheet.rows);
    sheet.merges = [...(savedSheet.merges || [])];
    sheet.rowCount = savedSheet.rowCount;
    ensureSheetShape(sheet);
  });
  state.selectedRow = null;
  state.selectedRange = null;
  saveEdits(false);
  render();
}

function pushHistory() {
  history.undo.push(snapshot());
  if (history.undo.length > history.limit) history.undo.shift();
  history.redo = [];
  updateButtons();
}

function undo() {
  if (!history.undo.length) return;
  history.redo.push(snapshot());
  restoreSnapshot(history.undo.pop());
}

function redo() {
  if (!history.redo.length) return;
  history.undo.push(snapshot());
  restoreSnapshot(history.redo.pop());
}

function loadSavedEdits() {
  try {
    const saved = JSON.parse(localStorage.getItem(storageKey) || "null");
    if (!saved || !Array.isArray(saved.sheets)) return;
    saved.sheets.forEach((savedSheet, sheetIndex) => {
      const sheet = workbook.sheets[sheetIndex];
      if (!sheet || !Array.isArray(savedSheet.rows)) return;
      sheet.rows = cloneRows(savedSheet.rows);
      sheet.merges = Array.isArray(savedSheet.merges) ? [...savedSheet.merges] : [...(sheet.merges || [])];
      ensureSheetShape(sheet);
    });
  } catch {
    localStorage.removeItem(storageKey);
  }
}

function saveEdits(update = true) {
  const payload = {
    savedAt: new Date().toISOString(),
    sheets: workbook.sheets.map(sheet => ({
      name: sheet.name,
      rows: sheet.rows,
      merges: sheet.merges || [],
    })),
  };
  localStorage.setItem(storageKey, JSON.stringify(payload));
  if (update) updateMeta();
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

function indexToColumn(index) {
  let n = index + 1;
  let value = "";
  while (n > 0) {
    const remainder = (n - 1) % 26;
    value = String.fromCharCode(65 + remainder) + value;
    n = Math.floor((n - 1) / 26);
  }
  return value;
}

function columnToIndex(column) {
  return column.split("").reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0) - 1;
}

function parseCellRef(ref) {
  const match = String(ref).match(/^([A-Z]+)(\d+)$/i);
  if (!match) return null;
  return { row: Number(match[2]) - 1, col: columnToIndex(match[1].toUpperCase()) };
}

function normalizeRange(range) {
  if (!range) return null;
  return {
    sr: Math.min(range.sr, range.er),
    sc: Math.min(range.sc, range.ec),
    er: Math.max(range.sr, range.er),
    ec: Math.max(range.sc, range.ec),
  };
}

function parseMergeRange(value) {
  const [start, end = start] = String(value).split(":");
  const a = parseCellRef(start);
  const b = parseCellRef(end);
  if (!a || !b) return null;
  return normalizeRange({ sr: a.row, sc: a.col, er: b.row, ec: b.col });
}

function rangeToA1(range) {
  const r = normalizeRange(range);
  const start = `${indexToColumn(r.sc)}${r.sr + 1}`;
  const end = `${indexToColumn(r.ec)}${r.er + 1}`;
  return start === end ? start : `${start}:${end}`;
}

function rangesOverlap(a, b) {
  const left = normalizeRange(a);
  const right = normalizeRange(b);
  return left.sr <= right.er && left.er >= right.sr && left.sc <= right.ec && left.ec >= right.sc;
}

function styleAttr(style) {
  if (!style) return "";
  const rules = [];
  if (style.fill) rules.push(`background:${style.fill}`);
  if (style.color) rules.push(`color:${style.color}`);
  if (style.bold) rules.push("font-weight:700");
  if (style.italic) rules.push("font-style:italic");
  return rules.length ? ` style="${rules.join(";")}"` : "";
}

function getOriginalText(sheetIndex, rowIndex, colIndex) {
  return originalSheets[sheetIndex]?.rows?.[rowIndex]?.[colIndex]?.text ?? "";
}

function isChanged(sheetIndex, rowIndex, colIndex) {
  const sheet = workbook.sheets[sheetIndex];
  return (sheet.rows[rowIndex]?.[colIndex]?.text ?? "") !== getOriginalText(sheetIndex, rowIndex, colIndex);
}

function isCellSelected(rowIndex, colIndex) {
  const range = normalizeRange(state.selectedRange);
  return Boolean(range && rowIndex >= range.sr && rowIndex <= range.er && colIndex >= range.sc && colIndex <= range.ec);
}

function getMergeInfo(sheet, useMerges) {
  const skip = new Set();
  const spans = new Map();
  if (!useMerges) return { skip, spans };
  (sheet.merges || []).forEach(value => {
    const range = parseMergeRange(value);
    if (!range) return;
    const bounded = {
      sr: Math.max(0, range.sr),
      sc: Math.max(0, range.sc),
      er: Math.min(sheet.rows.length - 1, range.er),
      ec: Math.min(sheet.colCount - 1, range.ec),
    };
    if (bounded.sr > bounded.er || bounded.sc > bounded.ec) return;
    spans.set(`${bounded.sr}:${bounded.sc}`, {
      rowspan: bounded.er - bounded.sr + 1,
      colspan: bounded.ec - bounded.sc + 1,
    });
    for (let row = bounded.sr; row <= bounded.er; row += 1) {
      for (let col = bounded.sc; col <= bounded.ec; col += 1) {
        if (row !== bounded.sr || col !== bounded.sc) skip.add(`${row}:${col}`);
      }
    }
  });
  return { skip, spans };
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
  const range = normalizeRange(state.selectedRange);
  const rangeText = range ? ` · 选区 ${rangeToA1(range)}` : "";
  const saved = localStorage.getItem(storageKey) ? "已自动保存在本浏览器" : "未保存本地改动";
  document.getElementById("meta").textContent = `${workbook.sheets.length} 个工作表 · ${new Date(workbook.generatedAt).toLocaleString("zh-CN")}`;
  editStatus.textContent = changed ? `${changed} 个单元格已修改 · ${saved}${rangeText}` : `没有未还原的修改${rangeText}`;
  updateButtons();
}

function updateButtons() {
  const range = normalizeRange(state.selectedRange);
  const area = range ? (range.er - range.sr + 1) * (range.ec - range.sc + 1) : 0;
  undoButton.disabled = history.undo.length === 0;
  redoButton.disabled = history.redo.length === 0;
  deleteRowButton.disabled = state.selectedRow === null;
  mergeButton.disabled = area < 2 || Boolean(state.query);
  unmergeButton.disabled = !range || Boolean(state.query);
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
  const useMerges = !query;
  const mergeInfo = getMergeInfo(sheet, useMerges);
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
      if (mergeInfo.skip.has(`${index}:${colIndex}`)) return "";
      const span = mergeInfo.spans.get(`${index}:${colIndex}`);
      const classes = ["cell"];
      if (!cell.text) classes.push("empty");
      if (isChanged(state.sheetIndex, index, colIndex)) classes.push("changed");
      if (isCellSelected(index, colIndex)) classes.push("selected-cell");
      if (span) classes.push("merged-cell");
      const title = cell.formula ? ` title="${escapeText(cell.formula)}"` : "";
      const spanAttrs = span ? ` rowspan="${span.rowspan}" colspan="${span.colspan}"` : "";
      return `<td class="${classes.join(" ")}" tabindex="0" data-row="${index}" data-col="${colIndex}"${spanAttrs}${title}${styleAttr(cell.style)}>${escapeText(cell.text)}</td>`;
    }).join("");
    return `<tr class="${selected}"><th class="row-head" data-row="${index}" title="点击选择整行">${index + 1}</th>${cells}</tr>`;
  }).join("");

  table.innerHTML = `<thead><tr>${colHeaders}</tr></thead><tbody>${body}</tbody>`;
  count.textContent = `${rows.length} / ${sheet.rowCount} 行`;
  updateMeta();
}

function render() {
  renderTabs();
  renderTable();
}

function updateCell(rowIndex, colIndex, text) {
  const sheet = workbook.sheets[state.sheetIndex];
  const cell = sheet.rows[rowIndex][colIndex] || { text: "" };
  if ((cell.text ?? "") === text) return false;
  pushHistory();
  cell.text = text;
  cell.value = text;
  cell.style = { ...(cell.style || {}), align: "center", valign: "middle" };
  if (cell.formula && text !== getOriginalText(state.sheetIndex, rowIndex, colIndex)) {
    delete cell.formula;
  }
  sheet.rows[rowIndex][colIndex] = cell;
  saveEdits();
  return true;
}

function startCellEdit(cellElement) {
  if (cellElement.classList.contains("editing")) return;
  const rowIndex = Number(cellElement.dataset.row);
  const colIndex = Number(cellElement.dataset.col);
  const sheet = workbook.sheets[state.sheetIndex];
  const current = sheet.rows[rowIndex]?.[colIndex]?.text ?? "";
  state.selectedRange = { sr: rowIndex, sc: colIndex, er: rowIndex, ec: colIndex };
  state.selectedRow = null;
  updateMeta();
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
    const mergeInfo = getMergeInfo(sheet, true);
    const rows = sheet.rows.map((row, rowIndex) => {
      const cells = row.map((cell, colIndex) => {
        if (mergeInfo.skip.has(`${rowIndex}:${colIndex}`)) return "";
        const span = mergeInfo.spans.get(`${rowIndex}:${colIndex}`);
        const mergeAttrs = span ? ` ss:MergeAcross="${span.colspan - 1}" ss:MergeDown="${span.rowspan - 1}"` : "";
        return `<Cell ss:StyleID="Center"${mergeAttrs}><Data ss:Type="String">${escapeXml(cell.text ?? "")}</Data></Cell>`;
      }).join("");
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
<Styles>
  <Style ss:ID="Center"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/></Style>
</Styles>
${worksheets}
</Workbook>`;
  downloadBlob(xml, `${workbook.title}-编辑版.xls`, "application/vnd.ms-excel;charset=utf-8");
}

function addRow() {
  pushHistory();
  const sheet = workbook.sheets[state.sheetIndex];
  const row = Array.from({ length: sheet.colCount }, () => ({ text: "", style: { align: "center", valign: "middle" } }));
  sheet.rows.push(row);
  sheet.rowCount = sheet.rows.length;
  state.selectedRow = sheet.rows.length - 1;
  state.selectedRange = null;
  saveEdits();
  renderTable();
}

function deleteSelectedRow() {
  if (state.selectedRow === null) return;
  pushHistory();
  const sheet = workbook.sheets[state.sheetIndex];
  sheet.rows.splice(state.selectedRow, 1);
  sheet.merges = (sheet.merges || []).filter(value => {
    const range = parseMergeRange(value);
    return range && range.er < state.selectedRow || range && range.sr > state.selectedRow;
  });
  sheet.rowCount = sheet.rows.length;
  state.selectedRow = null;
  state.selectedRange = null;
  saveEdits();
  renderTable();
}

function resetSheet() {
  pushHistory();
  const sheet = workbook.sheets[state.sheetIndex];
  const original = originalSheets[state.sheetIndex];
  sheet.rows = cloneRows(original.rows);
  sheet.merges = [...(original.merges || [])];
  sheet.rowCount = sheet.rows.length;
  state.selectedRow = null;
  state.selectedRange = null;
  saveEdits();
  renderTable();
}

function resetAll() {
  pushHistory();
  workbook.sheets.forEach((sheet, index) => {
    const original = originalSheets[index];
    sheet.rows = cloneRows(original.rows);
    sheet.merges = [...(original.merges || [])];
    sheet.rowCount = sheet.rows.length;
  });
  state.selectedRow = null;
  state.selectedRange = null;
  localStorage.removeItem(storageKey);
  render();
}

function mergeSelectedCells() {
  const sheet = workbook.sheets[state.sheetIndex];
  const range = normalizeRange(state.selectedRange);
  if (!range || (range.er === range.sr && range.ec === range.sc) || state.query) return;
  pushHistory();
  sheet.merges = (sheet.merges || []).filter(value => {
    const existing = parseMergeRange(value);
    return existing && !rangesOverlap(existing, range);
  });
  sheet.merges.push(rangeToA1(range));
  for (let row = range.sr; row <= range.er; row += 1) {
    for (let col = range.sc; col <= range.ec; col += 1) {
      const cell = sheet.rows[row][col] || { text: "" };
      cell.style = { ...(cell.style || {}), align: "center", valign: "middle" };
      if (row !== range.sr || col !== range.sc) {
        cell.text = "";
        cell.value = "";
        delete cell.formula;
      }
      sheet.rows[row][col] = cell;
    }
  }
  saveEdits();
  renderTable();
}

function unmergeSelectedCells() {
  const sheet = workbook.sheets[state.sheetIndex];
  const range = normalizeRange(state.selectedRange);
  if (!range || state.query) return;
  const next = (sheet.merges || []).filter(value => {
    const existing = parseMergeRange(value);
    return existing && !rangesOverlap(existing, range);
  });
  if (next.length === (sheet.merges || []).length) return;
  pushHistory();
  sheet.merges = next;
  saveEdits();
  renderTable();
}

tabs.addEventListener("click", event => {
  const button = event.target.closest("button[data-index]");
  if (!button) return;
  state.sheetIndex = Number(button.dataset.index);
  state.query = "";
  state.selectedRow = null;
  state.selectedRange = null;
  search.value = "";
  render();
});

search.addEventListener("input", event => {
  state.query = event.target.value;
  state.selectedRow = null;
  state.selectedRange = null;
  renderTable();
});

table.addEventListener("click", event => {
  const rowHead = event.target.closest(".row-head[data-row]");
  if (rowHead) {
    state.selectedRow = Number(rowHead.dataset.row);
    state.selectedRange = null;
    renderTable();
    return;
  }
  const cell = event.target.closest("td.cell[data-row][data-col]");
  if (!cell) return;
  const rowIndex = Number(cell.dataset.row);
  const colIndex = Number(cell.dataset.col);
  if (event.shiftKey && state.selectedRange) {
    state.selectedRange = {
      sr: state.selectedRange.sr,
      sc: state.selectedRange.sc,
      er: rowIndex,
      ec: colIndex,
    };
    state.selectedRow = null;
    renderTable();
    return;
  }
  startCellEdit(cell);
});

table.addEventListener("keydown", event => {
  const cell = event.target.closest("td.cell[data-row][data-col]");
  if (!cell) return;
  if (event.key === "Enter") {
    event.preventDefault();
    startCellEdit(cell);
  }
});

document.addEventListener("keydown", event => {
  if (!event.ctrlKey && !event.metaKey) return;
  const key = event.key.toLowerCase();
  if (key === "z") {
    event.preventDefault();
    undo();
  }
  if (key === "y") {
    event.preventDefault();
    redo();
  }
});

undoButton.addEventListener("click", undo);
redoButton.addEventListener("click", redo);
document.getElementById("export-xls").addEventListener("click", exportExcelXml);
document.getElementById("export-csv").addEventListener("click", exportCsv);
document.getElementById("add-row").addEventListener("click", addRow);
deleteRowButton.addEventListener("click", deleteSelectedRow);
mergeButton.addEventListener("click", mergeSelectedCells);
unmergeButton.addEventListener("click", unmergeSelectedCells);
document.getElementById("reset-sheet").addEventListener("click", resetSheet);
document.getElementById("reset-all").addEventListener("click", resetAll);

loadSavedEdits();
workbook.sheets.forEach(ensureSheetShape);
render();
