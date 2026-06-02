const state = {
  sheetIndex: 0,
  query: "",
};

const workbook = window.WORKBOOK_DATA;
const tabs = document.getElementById("tabs");
const table = document.getElementById("sheet-table");
const search = document.getElementById("search");
const count = document.getElementById("count");

document.getElementById("workbook-title").textContent = workbook.title;
document.getElementById("meta").textContent = `${workbook.sheets.length} 个工作表 · ${new Date(workbook.generatedAt).toLocaleString("zh-CN")}`;

function escapeText(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
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
  const query = state.query.trim().toLowerCase();
  const rows = sheet.rows
    .map((row, index) => ({ row, index }))
    .filter(({ row }) => rowMatches(row, query));

  const colHeaders = ["<th class=\"corner\"></th>", ...sheet.columns.map((col, index) => {
    const width = sheet.widths[index] ? Math.min(Math.max(sheet.widths[index] * 8, 72), 220) : 120;
    return `<th class="col-head" style="min-width:${width}px">${col}</th>`;
  })].join("");

  const body = rows.map(({ row, index }) => {
    const cells = row.map(cell => {
      const classes = ["cell"];
      if (!cell.text) classes.push("empty");
      const title = cell.formula ? ` title="${escapeText(cell.formula)}"` : "";
      return `<td class="${classes.join(" ")}"${title}${styleAttr(cell.style)}>${escapeText(cell.text)}</td>`;
    }).join("");
    return `<tr><th class="row-head">${index + 1}</th>${cells}</tr>`;
  }).join("");

  table.innerHTML = `<thead><tr>${colHeaders}</tr></thead><tbody>${body}</tbody>`;
  count.textContent = `${rows.length} / ${sheet.rowCount} 行`;
}

function render() {
  renderTabs();
  renderTable();
}

tabs.addEventListener("click", event => {
  const button = event.target.closest("button[data-index]");
  if (!button) return;
  state.sheetIndex = Number(button.dataset.index);
  state.query = "";
  search.value = "";
  render();
});

search.addEventListener("input", event => {
  state.query = event.target.value;
  renderTable();
});

render();
