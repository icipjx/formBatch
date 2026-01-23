/**
 * FormBatcher MVP Side Panel
 * - Import CSV/XLSX
 * - Field mapping by "pick element on page"
 * - Fill rows one by one (no submit/click automation)
 */

const $ = (sel) => document.querySelector(sel);

const ui = {
  app: document.querySelector(".app"),
  fileInput: $("#fileInput"),
  fileNameText: $("#fileNameText"),
  sheetSelect: $("#sheetSelect"),
  rowStart: $("#rowStart"),
  rowEnd: $("#rowEnd"),
  pillCols: $("#pillCols"),
  pillRows: $("#pillRows"),
  colSelect: $("#colSelect"),
  btnPick: $("#btnPick"),
  pickHint: $("#pickHint"),
  mapList: $("#mapList"),
  btnSaveMap: $("#btnSaveMap"),
  btnClearMap: $("#btnClearMap"),
  btnLoadSite: $("#btnLoadSite"),
  btnViewSites: $("#btnViewSites"),
  siteList: $("#siteList"),
  btnStep: $("#btnStep"),
  btnNextRow: $("#btnNextRow"),
  stState: $("#stState"),
  stProg: $("#stProg"),
  stRow: $("#stRow"),
  stMsg: $("#stMsg"),
  btnExport: $("#btnExport"),
  logBox: $("#logBox"),
  btnSettings: $("#btnSettings"),
  btnHelp: $("#btnHelp"),
  helpCard: $("#helpCard"),
  btnClearFile: $("#btnClearFile"),
  rowPreviewHead: $("#rowPreviewHead"),
  rowPreviewBody: $("#rowPreviewBody"),
  autoAdvance: $("#autoAdvance"),
  filterField: $("#filterField"),
  filterKeyword: $("#filterKeyword"),
  btnFilter: $("#btnFilter"),
  filterTitle: $("#filterTitle"),
  filterList: $("#filterList"),
  settingsPanel: $("#settingsPanel"),
  identityColSelect: $("#identityColSelect"),
  btnPickIdentity: $("#btnPickIdentity"),
  identityStrategy: $("#identityStrategy"),
  btnMatchIdentity: $("#btnMatchIdentity"),
  identityStatus: $("#identityStatus"),
  identityMapList: $("#identityMapList"),
  identityMatchTitle: $("#identityMatchTitle"),
  identityMatchList: $("#identityMatchList"),
};

const state = {
  data: {
    headers: [],
    rows: [],
    filename: "",
    sheets: [],
    activeSheet: "",
    workbook: null,
  },
  mapping: [], // {col, colIndex, selector, meta}
  runner: {
    running: false,
    results: [], // {i, status, msg}
    siteKey: null,
  },
  pick: {
    active: false,
    mode: null,
  },
  selection: {
    currentRowIndex: null, // Excel row number
  },
  lang: "zh",
  identityMapping: null, // {colIndex, colName, selector, matchStrategy, normalize}
};

const i18n = {
  zh: {
    "brand.sub": "Batch-fill any web form · CSV / XLSX",
    "settings.title": "设置",
    "settings.theme": "主题",
    "settings.theme.dark": "暗色",
    "settings.theme.light": "亮色",
    "settings.lang": "语言",
    "settings.lang.zh": "中文",
    "settings.lang.en": "English",
    "help.title": "使用说明（MVP）",
    "help.desc": "不需要系统开发权限，模拟你在网页上输入与点击。",
    "help.step1": "打开目标系统的“新增/编辑”表单页面。",
    "help.step2": "打开浏览器侧边栏：点击扩展图标→打开侧边栏。",
    "help.step3": "选择 CSV / XLSX 文件（第一行是表头）。",
    "help.step4": "在“选择列”里选中列名，点击“开始点选网页字段”，然后去网页上点对应输入框完成绑定。",
    "help.step5": "点击“填充当前行”，逐行完成录入。",
    "help.tipLabel": "提示：",
    "help.tip": "保存/提交由用户手动完成，避免误操作。",
    "log.title": "日志",
    "log.desc": "运行过程与异常记录。",
    "data.title": "1) 数据源",
    "data.desc": "导入表格文件，使用第一行作为表头。",
    "data.upload": "上传数据文件",
    "data.noFile": "未选择",
    "data.sheet": "Sheet（仅 XLSX）",
    "data.rowRange": "行范围",
    "map.title": "2) 字段映射",
    "map.desc": "绑定数据列映射关系和数据行的映射关系（身份映射）。",
    "map.selectCol": "数据列映射",
    "map.pick": "开始点选网页数据列字段",
    "map.hint": "提示：点击网页中的输入框/下拉框/文本区域，即可绑定到当前列。",
    "map.list": "数据列映射列表",
    "map.fileCols": "文件数据列(默认第一行)",
    "map.save": "保存到当前站点",
    "map.clear": "清空映射",
    "map.load": "载入站点配置",
    "map.view": "查看站点配置",
    "identity.title": "数据行映射(身份映射)",
    "identity.desc": "建议使用唯一 ID 列，名称可能重复。",
    "identity.col": "锚点列（建议使用唯一 ID 列，名称可能重复）",
    "identity.pick": "开始点选网页身份列字段",
    "identity.list": "身份映射列表",
    "identity.hint": "提示：点击网页任意元素，绑定数据身份列，实现当前页面自动选择填充数据行。",
    "identity.strategy": "匹配策略",
    "identity.autoTitle": "自动匹配填充行",
    "identity.matches": "匹配结果",
    "identity.autoTip": "提示：需完成数据行映射(身份映射)",
    "identity.strategy.exact": "完全一致",
    "identity.strategy.contains": "包含匹配",
    "identity.match": "自动匹配当前页面行",
    "fill.title": "3) 网页填充",
    "fill.desc": "逐行填充，用户手动确认保存/提交。",
    "fill.rowSelect": "手动选择填充行",
    "fill.filterField": "筛选字段",
    "fill.keyword": "关键词（模糊匹配）",
    "fill.keyword.ph": "例如：张三 / EMP-0001",
    "fill.search": "查询",
    "fill.result": "查询结果",
    "fill.block": "填充",
    "fill.preview": "预览当前行号 — 数据",
    "fill.auto": "自动步进下一行数据",
    "fill.step": "填充当前行",
    "fill.next": "下一条",
    "footer.text": "v0.1 MVP · Local-first · No cloud",
    "footer.siteKey": "站点配置键",
  },
  en: {
    "brand.sub": "Batch-fill any web form · CSV / XLSX",
    "settings.title": "Settings",
    "settings.theme": "Theme",
    "settings.theme.dark": "Dark",
    "settings.theme.light": "Light",
    "settings.lang": "Language",
    "settings.lang.zh": "中文",
    "settings.lang.en": "English",
    "help.title": "Quick Guide (MVP)",
    "help.desc": "No system privileges needed. Simulates input and clicks on the page.",
    "help.step1": "Open the target system’s new/edit form page.",
    "help.step2": "Open the side panel: click the extension icon → open side panel.",
    "help.step3": "Choose a CSV/XLSX file (first row is headers).",
    "help.step4": "Select a column, click “Pick field on page”, then click the matching input on the page.",
    "help.step5": "Click “Fill current row” to fill row by row.",
    "help.tipLabel": "Tip:",
    "help.tip": "Save/submit manually to avoid mistakes.",
    "log.title": "Logs",
    "log.desc": "Runtime events and errors.",
    "data.title": "1) Data Source",
    "data.desc": "Import a sheet. Use the first row as headers.",
    "data.upload": "Upload data file",
    "data.noFile": "No file",
    "data.sheet": "Sheet (XLSX only)",
    "data.rowRange": "Row range",
    "map.title": "2) Field Mapping",
    "map.desc": "Bind data column mappings and identity row mapping.",
    "map.selectCol": "Data column mapping",
    "map.pick": "Pick data field on page",
    "map.hint": "Tip: click an input/select/textarea to bind it to current column.",
    "map.list": "Data mapping list",
    "map.fileCols": "File data columns (default: first row)",
    "map.save": "Save to current site",
    "map.clear": "Clear mapping",
    "map.load": "Load site config",
    "map.view": "View site configs",
    "identity.title": "Row Mapping (Identity)",
    "identity.desc": "Use a unique ID when possible; names may repeat.",
    "identity.col": "Anchor column (prefer a unique ID; names may repeat)",
    "identity.pick": "Pick identity column on page",
    "identity.list": "Identity mapping",
    "identity.hint": "Tip: click any element to bind the identity column and auto-locate the row.",
    "identity.strategy": "Match strategy",
    "identity.autoTitle": "Auto match row",
    "identity.matches": "Match results",
    "identity.autoTip": "Tip: complete row mapping (identity) first.",
    "identity.strategy.exact": "Exact",
    "identity.strategy.contains": "Contains",
    "identity.match": "Match row on current page",
    "fill.title": "3) Web Fill",
    "fill.desc": "Fill row by row. User confirms save/submit.",
    "fill.rowSelect": "Manual row selection",
    "fill.filterField": "Filter field",
    "fill.keyword": "Keyword (fuzzy)",
    "fill.keyword.ph": "e.g. John / EMP-0001",
    "fill.search": "Search",
    "fill.result": "Search results",
    "fill.block": "Fill",
    "fill.preview": "Preview row — data",
    "fill.auto": "Auto-advance to next row",
    "fill.step": "Fill current row",
    "fill.next": "Next",
    "footer.text": "v0.1 MVP · Local-first · No cloud",
    "footer.siteKey": "Site config key",
  },
};

function applyLang(lang) {
  const dict = i18n[lang] || i18n.zh;
  document.querySelectorAll("[data-i18n]").forEach((el) => {
    const key = el.getAttribute("data-i18n");
    if (dict[key]) el.textContent = dict[key];
  });
  document.querySelectorAll("[data-i18n-title]").forEach((el) => {
    const key = el.getAttribute("data-i18n-title");
    if (dict[key]) el.setAttribute("title", dict[key]);
  });
  document.querySelectorAll("[data-i18n-placeholder]").forEach((el) => {
    const key = el.getAttribute("data-i18n-placeholder");
    if (dict[key]) el.setAttribute("placeholder", dict[key]);
  });
  state.lang = lang;
}

function setTheme(theme) {
  ui.app.setAttribute("data-theme", theme);
  chrome.storage.local.set({ "ui:theme": theme });
  document.querySelectorAll(".seg-btn[data-theme-val]").forEach((btn) => {
    btn.classList.toggle("is-active", btn.getAttribute("data-theme-val") === theme);
  });
}

function setLang(lang) {
  applyLang(lang);
  chrome.storage.local.set({ "ui:lang": lang });
  document.querySelectorAll(".seg-btn[data-lang]").forEach((btn) => {
    btn.classList.toggle("is-active", btn.getAttribute("data-lang") === lang);
  });
  renderIdentityMappingUI();
}

function log(line) {
  const t = new Date().toLocaleTimeString();
  ui.logBox.textContent = `[${t}] ${line}\n` + ui.logBox.textContent;
}

function setStatus({ stateText, progText, rowText, msg }) {
  if (stateText && ui.stState) ui.stState.textContent = stateText;
  if (progText && ui.stProg) ui.stProg.textContent = progText;
  if (rowText && ui.stRow) ui.stRow.textContent = rowText;
  if (msg && ui.stMsg) ui.stMsg.textContent = msg;
}

async function getActiveTab() {
  const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  return tabs?.[0];
}

async function safeSendToActiveTab(message) {
  const tab = await getActiveTab();
  if (!tab?.id) throw new Error("未找到活动标签页");
  const url = tab.url || "";
  if (url.startsWith("chrome://") || url.startsWith("edge://") || url.startsWith("chrome-extension://")) {
    throw new Error("当前页面不允许注入脚本，请切换到普通网页。");
  }
  try {
    return await chrome.tabs.sendMessage(tab.id, message);
  } catch (e) {
    try {
      await chrome.scripting.executeScript({ target: { tabId: tab.id }, files: ["content.js"] });
      return await chrome.tabs.sendMessage(tab.id, message);
    } catch (_) {
      throw new Error("未连接到页面，请刷新页面或确保该页面可注入。");
    }
  }
}

function refreshColsUI() {
  ui.colSelect.innerHTML = "";
  ui.filterField.innerHTML = "";
  if (ui.identityColSelect) ui.identityColSelect.innerHTML = "";
  state.data.headers.forEach((h, idx) => {
    const label = h || `列${idx + 1}`;
    const opt1 = document.createElement("option");
    opt1.value = String(idx);
    opt1.textContent = label;
    ui.colSelect.appendChild(opt1);

    const opt2 = document.createElement("option");
    opt2.value = String(idx);
    opt2.textContent = label;
    ui.filterField.appendChild(opt2);

    if (ui.identityColSelect) {
      const opt3 = document.createElement("option");
      opt3.value = String(idx);
      opt3.textContent = label;
      ui.identityColSelect.appendChild(opt3);
    }
  });
  ui.filterField.value = "0";
  ui.pillCols.textContent = `${state.data.headers.length} 列`;
  ui.pillRows.textContent = `${state.data.rows.length} 行`;
  updatePickIndicator();
  updateIdentityIndicator();
  updateFillIndicator();
}

function updatePickIndicator() {
  const hasSelection = state.data.headers.length > 0 && String(ui.colSelect.value || "").length > 0;
  const active = state.pick.active && state.pick.mode === "field";
  ui.btnPick.classList.toggle("is-ready", active);
  ui.btnPick.disabled = !hasSelection;
  const label = ui.btnPick.querySelector("[data-i18n=\"map.pick\"]");
  if (label) {
    label.textContent = active
      ? (state.lang === "en" ? "Cancel pick" : "取消点选")
      : (state.lang === "en" ? "Pick data field on page" : "开始点选网页数据列字段");
  }
}

function updateIdentityIndicator() {
  if (!ui.btnPickIdentity || !ui.identityColSelect) return;
  const hasSelection = state.data.headers.length > 0 && String(ui.identityColSelect.value || "").length > 0;
  const active = state.pick.active && state.pick.mode === "identity";
  ui.btnPickIdentity.classList.toggle("is-ready", active);
  ui.btnPickIdentity.disabled = !hasSelection;
  const label = ui.btnPickIdentity.querySelector("[data-i18n=\"identity.pick\"]");
  if (label) {
    label.textContent = active
      ? (state.lang === "en" ? "Cancel pick" : "取消点选")
      : (state.lang === "en" ? "Pick identity column on page" : "开始点选网页身份列字段");
  }
}

function updateIdentityUI() {
  if (!state.identityMapping) {
    if (ui.identityStatus) ui.identityStatus.textContent = "—";
    renderIdentityMappingUI();
    return;
  }
  if (ui.identityColSelect) {
    if (typeof state.identityMapping.colIndex === "number") {
      ui.identityColSelect.value = String(state.identityMapping.colIndex);
    } else if (state.identityMapping.colName) {
      const idx = state.data.headers.findIndex((h) => h === state.identityMapping.colName);
      if (idx >= 0) {
        state.identityMapping.colIndex = idx;
        ui.identityColSelect.value = String(idx);
      }
    }
  }
  if (ui.identityStrategy) ui.identityStrategy.value = state.identityMapping.matchStrategy || "exact";
  renderIdentityMappingUI();
}

function renderIdentityMappingUI() {
  if (!ui.identityMapList) return;
  ui.identityMapList.innerHTML = "";
  if (!state.identityMapping?.selector) {
    const empty = document.createElement("div");
    empty.style.padding = "10px";
    empty.style.color = "var(--muted)";
    empty.style.fontSize = "12px";
    empty.textContent = state.lang === "en"
      ? "No identity mapping yet. Select a column and pick an element."
      : "暂无身份映射。选择一列并点选页面身份字段。";
    ui.identityMapList.appendChild(empty);
    return;
  }

  const wrap = document.createElement("div");
  wrap.className = "mapitem";

  const badge = document.createElement("div");
  badge.className = "badge";
  badge.textContent = `${state.identityMapping.colName || (state.lang === "en" ? "Unnamed" : "未命名列")}`;

  const meta = document.createElement("div");
  meta.className = "mapmeta";
  const t = document.createElement("div");
  t.className = "t";
  t.textContent = state.lang === "en"
    ? `${state.identityMapping.colName || "Column"} → Page`
    : `${state.identityMapping.colName || "锚点列"} → 页面元素`;
  const d = document.createElement("div");
  d.className = "d";
  d.textContent = state.identityMapping.selector;

  meta.appendChild(t);
  meta.appendChild(d);
  wrap.appendChild(badge);
  wrap.appendChild(meta);

  const x = document.createElement("button");
  x.className = "x";
  x.textContent = state.lang === "en" ? "Remove" : "移除";
  x.addEventListener("click", () => {
    state.identityMapping = null;
    updateIdentityUI();
  });
  wrap.appendChild(x);
  ui.identityMapList.appendChild(wrap);
}

function updateFillIndicator() {
  const hasRows = state.data.rows.length > 0;
  const hasMapping = state.mapping.length > 0;
  const ready = hasRows && hasMapping;
  ui.btnStep.classList.toggle("is-ready", ready);
}

function setupCardToggles() {
  document.querySelectorAll(".card-hd").forEach((hd) => {
    const btn = document.createElement("button");
    btn.className = "card-toggle";
    btn.type = "button";
    btn.textContent = "v";
    btn.setAttribute("aria-expanded", "true");
    hd.appendChild(btn);

    const card = hd.closest(".card");
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      if (!card) return;
      const collapsed = card.classList.toggle("collapsed");
      btn.setAttribute("aria-expanded", collapsed ? "false" : "true");
      btn.textContent = collapsed ? ">" : "v";
    });
  });
}

function setupGroupToggles() {
  document.querySelectorAll(".group-hd").forEach((hd) => {
    const btn = hd.querySelector(".group-toggle");
    if (!btn) return;
    const group = hd.closest(".group");
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      if (!group) return;
      const collapsed = group.classList.toggle("collapsed");
      btn.setAttribute("aria-expanded", collapsed ? "false" : "true");
      btn.textContent = collapsed ? ">" : "v";
    });
  });
}

function setupModuleLayout() {
  const logCard = ui.logBox?.closest(".card");
  const helpCard = ui.helpCard;
  const content = document.querySelector(".content");
  if (logCard && content) {
    logCard.classList.add("is-log");
    logCard.classList.add("collapsed");
    content.prepend(logCard);
  }
  if (helpCard && content) {
    content.prepend(helpCard);
  }
}

function updateRowRangeDefaults() {
  const total = state.data.rows.length;
  if (total > 0) {
    ui.rowStart.value = "2";
    ui.rowEnd.value = String(total + 1);
    state.selection.currentRowIndex = 2;
  } else {
    ui.rowStart.value = "2";
    ui.rowEnd.value = "2";
    state.selection.currentRowIndex = null;
  }
  updateFillIndicator();
}

function refreshSheetUI() {
  ui.sheetSelect.innerHTML = "";
  if (state.data.sheets.length === 0) {
    const opt = document.createElement("option");
    opt.value = "";
    opt.textContent = "—";
    ui.sheetSelect.appendChild(opt);
    ui.sheetSelect.disabled = true;
    return;
  }
  ui.sheetSelect.disabled = false;
  state.data.sheets.forEach((s) => {
    const opt = document.createElement("option");
    opt.value = s;
    opt.textContent = s;
    ui.sheetSelect.appendChild(opt);
  });
  ui.sheetSelect.value = state.data.activeSheet || state.data.sheets[0];
}

function refreshMappingUI() {
  ui.mapList.innerHTML = "";
  if (state.mapping.length === 0) {
    const empty = document.createElement("div");
    empty.style.padding = "10px";
    empty.style.color = "var(--muted)";
    empty.style.fontSize = "12px";
    empty.textContent = state.lang === "en"
      ? "No mapping yet. Select a column and pick a field on the page."
      : "暂无映射。选择一列并点选网页中的输入框。";
    ui.mapList.appendChild(empty);
    updateFillIndicator();
    return;
  }

  state.mapping.forEach((m, idx) => {
    const wrap = document.createElement("div");
    wrap.className = "mapitem";

    const badge = document.createElement("div");
    badge.className = "badge";
    badge.textContent = `${m.col || (state.lang === "en" ? "Unnamed" : "未命名列")}`;

    const meta = document.createElement("div");
    meta.className = "mapmeta";
    const t = document.createElement("div");
    t.className = "t";
    t.textContent = `${m.meta?.label || m.meta?.placeholder || m.meta?.name || m.meta?.tag || (state.lang === "en" ? "Field" : "字段")} → ${m.col}`;
    const d = document.createElement("div");
    d.className = "d";
    d.textContent = m.selector;

    meta.appendChild(t);
    meta.appendChild(d);

    const x = document.createElement("button");
    x.className = "x";
    x.textContent = state.lang === "en" ? "Remove" : "移除";
    x.addEventListener("click", () => {
      state.mapping.splice(idx, 1);
      refreshMappingUI();
    });

    wrap.appendChild(badge);
    wrap.appendChild(meta);
    wrap.appendChild(x);
    ui.mapList.appendChild(wrap);
  });

  updateFillIndicator();
}

function getRowSlice() {
  const rawStart = parseInt(ui.rowStart.value || "2", 10);
  const rawEnd = parseInt(ui.rowEnd.value || "999999", 10);
  const start = Math.max(2, rawStart || 2);
  const maxEnd = state.data.rows.length + 1;
  const end = Math.max(start, Math.min(rawEnd || maxEnd, maxEnd));
  const startIndex = start - 2;
  const endIndex = end - 1;
  const rows = state.data.rows.slice(startIndex, endIndex);
  return { start, end, rows };
}

function renderRowPreview(row, rowIndex) {
  if (!row || !row.length) {
    ui.rowPreviewHead.textContent = "—";
    ui.rowPreviewBody.textContent = "—";
    const preview = document.querySelector(".row-preview");
    if (preview) preview.classList.remove("is-pending", "is-done");
    return;
  }
  const rowNum = Number(rowIndex);
  const headers = [state.lang === "en" ? "Row" : "行号"].concat(
    state.data.headers.map((h, i) => h || `列${i + 1}`)
  );
  const body = row.map((v) => String(v ?? "").trim());
  const cols = Math.max(headers.length, body.length + 1);

  const rowLabel = Number.isFinite(rowNum) && rowNum > 0 ? String(rowNum) : "—";
  const bodyWithRow = [rowLabel].concat(body);
  ui.rowPreviewHead.innerHTML = headers.map((h) => `<span class="row-preview-cell">${h}</span>`).join("");
  ui.rowPreviewBody.innerHTML = bodyWithRow.map((v) => `<span class="row-preview-cell">${v || "—"}</span>`).join("");

  const headCells = Array.from(ui.rowPreviewHead.children);
  const bodyCells = Array.from(ui.rowPreviewBody.children);
  for (let i = 0; i < cols; i++) {
    const h = headCells[i];
    const b = bodyCells[i];
    if (!h || !b) continue;
    const width = Math.max(h.scrollWidth, b.scrollWidth);
    const w = Math.min(Math.max(width, 80), 260);
    h.style.flex = `0 0 ${w}px`;
    b.style.flex = `0 0 ${w}px`;
  }

  const preview = document.querySelector(".row-preview");
  if (preview) {
    if (!preview.classList.contains("is-done")) {
      preview.classList.add("is-pending");
    }
  }
}

async function matchIdentityRow() {
  if (!state.identityMapping?.selector) {
    if (ui.identityStatus) ui.identityStatus.textContent = state.lang === "en" ? "No identity mapping yet." : "尚未绑定页面身份字段。";
    return;
  }
  if (!state.data.rows.length) {
    if (ui.identityStatus) ui.identityStatus.textContent = state.lang === "en" ? "Please load data first." : "请先载入数据。";
    return;
  }
  const colIndex = parseInt(ui.identityColSelect?.value || state.identityMapping.colIndex || "0", 10);
  const strategy = ui.identityStrategy?.value || state.identityMapping.matchStrategy || "exact";
  state.identityMapping.colIndex = colIndex;
  state.identityMapping.matchStrategy = strategy;

  let pageValue = "";
  try {
    const res = await safeSendToActiveTab({
      type: "FB_READ_TEXT",
      payload: { selector: state.identityMapping.selector },
    });
    pageValue = String(res?.text || "").trim();
  } catch (e) {
    if (ui.identityStatus) ui.identityStatus.textContent = (state.lang === "en" ? "Read failed: " : "读取失败：") + String(e?.message || e);
    return;
  }

  if (!pageValue) {
    if (ui.identityStatus) ui.identityStatus.textContent = state.lang === "en" ? "No text found on page." : "页面未读取到身份文本。";
    return;
  }

  const matches = [];
  state.data.rows.forEach((row, idx) => {
    const cell = row?.[colIndex];
    if (cell === undefined || cell === null) return;
    const cellStr = String(cell);
    if (strategy === "exact") {
      if (cellStr === pageValue) matches.push(idx);
      return;
    }
    if (strategy === "contains") {
      if (cellStr.includes(pageValue) || pageValue.includes(cellStr)) matches.push(idx);
      return;
    }
  });

  if (matches.length === 0) {
    if (ui.identityStatus) ui.identityStatus.textContent = state.lang === "en" ? "No match found." : "未匹配到，请检查映射或数据。";
    renderRowPreview([], null);
    if (ui.identityMatchTitle) ui.identityMatchTitle.style.display = "none";
    if (ui.identityMatchList) ui.identityMatchList.style.display = "none";
    return;
  }
  if (matches.length > 1) {
    if (ui.identityStatus) {
      ui.identityStatus.textContent = state.lang === "en"
        ? `Matched ${matches.length} rows. Please choose one.`
        : `匹配到 ${matches.length} 条，请选择一条。`;
    }
    renderRowPreview([], null);
    if (ui.identityMatchTitle) ui.identityMatchTitle.style.display = "block";
    if (ui.identityMatchList) {
      ui.identityMatchList.innerHTML = "";
      const { start } = getRowSlice();
      matches.slice(0, 50).forEach((idx) => {
        const excelRow = idx + 2;
        const item = document.createElement("div");
        item.className = "filter-item";
        item.addEventListener("click", () => {
          state.selection.currentRowIndex = excelRow;
          renderRowPreview(state.data.rows[idx] || [], excelRow);
          ui.identityMatchList.style.display = "none";
        });

        const num = document.createElement("div");
        num.className = "filter-rownum";
        num.textContent = state.lang === "en" ? `Row ${excelRow}` : `第${excelRow}行`;

        const text = document.createElement("div");
        text.className = "filter-text";
        const rowVals = (state.data.rows[idx] || []).map((v) => String(v ?? "").trim()).filter((v) => v !== "");
        text.textContent = rowVals.join(" | ");

        item.appendChild(num);
        item.appendChild(text);
        ui.identityMatchList.appendChild(item);
      });
      ui.identityMatchList.style.display = "block";
    }
    return;
  }

  const rowIndex = matches[0] + 2;
  state.selection.currentRowIndex = rowIndex;
  renderRowPreview(state.data.rows[matches[0]] || [], rowIndex);
  if (ui.identityStatus) ui.identityStatus.textContent = state.lang === "en" ? `Matched row ${rowIndex}: ${pageValue}` : `已匹配到第 ${rowIndex} 行：${pageValue}`;
  if (ui.identityMatchTitle) ui.identityMatchTitle.style.display = "none";
  if (ui.identityMatchList) ui.identityMatchList.style.display = "none";
}

// -------- Data parsing --------
function parseCSV(text) {
  const rows = [];
  let cur = [];
  let val = "";
  let inQ = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    const n = text[i + 1];
    if (inQ) {
      if (c === '"' && n === '"') { val += '"'; i++; continue; }
      if (c === '"') { inQ = false; continue; }
      val += c; continue;
    }
    if (c === '"') { inQ = true; continue; }
    if (c === ",") { cur.push(val); val = ""; continue; }
    if (c === "\r" && n === "\n") { cur.push(val); rows.push(cur); cur = []; val = ""; i++; continue; }
    if (c === "\n") { cur.push(val); rows.push(cur); cur = []; val = ""; continue; }
    val += c;
  }
  cur.push(val);
  rows.push(cur);
  while (rows.length && rows[rows.length - 1].every((x) => (x ?? "").trim() === "")) rows.pop();
  return rows;
}

async function loadFile(file) {
  state.data = {
    headers: [],
    rows: [],
    filename: file.name,
    sheets: [],
    activeSheet: "",
    workbook: null,
  };
  ui.fileNameText.textContent = file.name;
  ui.fileNameText.parentElement?.classList.add("has-file");

  const ext = (file.name.split(".").pop() || "").toLowerCase();
  if (ext === "csv") {
    const text = await file.text();
    const all = parseCSV(text);
    const headers = all.shift() || [];
    state.data.headers = headers.map((h) => String(h ?? "").trim());
    state.data.rows = all;
    state.data.sheets = [];
    state.data.activeSheet = "";
    refreshSheetUI();
    refreshColsUI();
    updateRowRangeDefaults();
    renderRowPreview(state.data.rows[0] || [], 2);
    log(state.lang === "en"
      ? `Loaded CSV: ${state.data.headers.length} cols, ${state.data.rows.length} rows`
      : `已加载CSV：${state.data.headers.length} 列，${state.data.rows.length} 行`);
    return;
  }

  if (ext === "xlsx") {
    if (!window.XLSX) throw new Error(state.lang === "en" ? "Missing XLSX parser. Reload extension." : "缺少 XLSX 解析库，请重新加载扩展。" );
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    state.data.workbook = workbook;
    state.data.sheets = workbook.SheetNames || [];
    state.data.activeSheet = state.data.sheets[0] || "";
    refreshSheetUI();
    await reloadActiveSheetFromCurrentFile();
    renderRowPreview(state.data.rows[0] || [], 2);
    log(state.lang === "en"
      ? `Loaded XLSX: ${state.data.sheets.length} sheets`
      : `已加载XLSX：${state.data.sheets.length} 个 Sheet`);
    return;
  }

  throw new Error(state.lang === "en" ? "Only .csv or .xlsx" : "仅支持 .csv 或 .xlsx");
}

function clearFileSelection() {
  state.data = {
    headers: [],
    rows: [],
    filename: "",
    sheets: [],
    activeSheet: "",
    workbook: null,
  };
  ui.fileNameText.textContent = i18n[state.lang]["data.noFile"] || "未选择";
  ui.fileNameText.parentElement?.classList.remove("has-file");
  ui.fileInput.value = "";
  refreshSheetUI();
  refreshColsUI();
  renderRowPreview([], null);
  updateFillIndicator();
}

async function reloadActiveSheetFromCurrentFile() {
  if (!state.data.workbook) return;
  const sheetName = ui.sheetSelect.value || state.data.activeSheet;
  state.data.activeSheet = sheetName;
  const sheet = state.data.workbook.Sheets?.[sheetName];
  if (!sheet) {
    state.data.headers = [];
    state.data.rows = [];
    refreshColsUI();
    renderRowPreview([], null);
    updateFillIndicator();
    return;
  }
  const all = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const headers = all.shift() || [];
  state.data.headers = headers.map((h) => String(h ?? "").trim());
  state.data.rows = all;
  refreshColsUI();
  updateRowRangeDefaults();
  renderRowPreview(state.data.rows[0] || [], 2);
}

// -------- Site config persistence --------
async function computeSiteKey() {
  const tab = await getActiveTab();
  const url = new URL(tab.url);
  return `${url.origin}${url.pathname}`;
}

async function saveSiteConfig() {
  const key = await computeSiteKey();
  state.runner.siteKey = key;
  const payload = { mapping: state.mapping, identityMapping: state.identityMapping };
  await chrome.storage.local.set({ ["site:" + key]: payload });
  log(state.lang === "en" ? "Saved site config: " + key : "已保存站点配置：" + key);
}

async function loadSiteConfig() {
  const key = await computeSiteKey();
  state.runner.siteKey = key;
  const data = await chrome.storage.local.get("site:" + key);
  const payload = data?.["site:" + key];
  if (!payload) {
    log(state.lang === "en" ? "No config for this site: " + key : "该站点暂无保存配置：" + key);
    return;
  }
  state.mapping = payload.mapping || [];
  state.identityMapping = payload.identityMapping || null;
  refreshMappingUI();
  updateIdentityUI();
  log(state.lang === "en" ? "Loaded site config: " + key : "已载入站点配置：" + key);
}

async function loadSiteConfigByKey(key) {
  const data = await chrome.storage.local.get("site:" + key);
  const payload = data?.["site:" + key];
  if (!payload) {
    log(state.lang === "en" ? "Config not found: " + key : "未找到站点配置：" + key);
    return;
  }
  state.mapping = payload.mapping || [];
  state.identityMapping = payload.identityMapping || null;
  refreshMappingUI();
  updateIdentityUI();
  log(state.lang === "en" ? "Loaded site config: " + key : "已载入站点配置：" + key);
}

async function renderSiteList() {
  const data = await chrome.storage.local.get(null);
  const keys = Object.keys(data).filter((k) => k.startsWith("site:"));
  ui.siteList.innerHTML = "";

  if (keys.length === 0) {
    const empty = document.createElement("div");
    empty.className = "site-meta";
    empty.textContent = state.lang === "en" ? "No saved site configs." : "暂无保存的站点配置。";
    ui.siteList.appendChild(empty);
    return;
  }

  keys.sort();
  keys.forEach((k) => {
    const key = k.replace(/^site:/, "");

    const row = document.createElement("div");
    row.className = "site-item";
    row.title = state.lang === "en" ? "Click to load" : "点击载入";
    row.addEventListener("click", () => loadSiteConfigByKey(key));

    const label = document.createElement("div");
    label.textContent = key;

    const del = document.createElement("button");
    del.className = "site-del";
    del.type = "button";
    del.setAttribute("aria-label", state.lang === "en" ? "Delete config" : "删除配置");
    del.title = state.lang === "en" ? "Delete" : "删除";
    del.textContent = "";
    del.addEventListener("click", async (e) => {
      e.stopPropagation();
      await chrome.storage.local.remove("site:" + key);
      await renderSiteList();
    });

    row.appendChild(label);
    row.appendChild(del);
    ui.siteList.appendChild(row);
  });
}

// -------- Picking elements on the page --------
async function beginPick(mode) {
  const tab = await getActiveTab();
  if (!tab?.id) return;

  if (state.pick.active) {
    await cancelPick();
    return;
  }

  const colIndex = mode === "identity"
    ? parseInt(ui.identityColSelect?.value || "0", 10)
    : parseInt(ui.colSelect.value || "0", 10);
  const colName = state.data.headers[colIndex] ?? `列${colIndex + 1}`;

  if (mode === "field") {
    const exists = state.mapping.find((m) => m.colIndex === colIndex);
    if (exists) {
      const msg = state.lang === "en"
        ? `${colName} already has a mapping. Pick again?`
        : `${colName}列已存在映射关系，是否重新点选？`;
      if (!window.confirm(msg)) return;
    }
  }

  setStatus({ stateText: state.lang === "en" ? "Picking" : "点选中", msg: state.lang === "en" ? "Click the target element on the page." : "请到网页上点击目标元素。" });
  log(state.lang === "en" ? `Pick mode (${mode}): click an element.` : `进入点选模式(${mode})：请点击网页元素。`);

  state.pick.active = true;
  state.pick.mode = mode;
  updatePickIndicator();
  updateIdentityIndicator();

  try {
    await safeSendToActiveTab({
      type: "FB_BEGIN_PICK",
      mode,
      colIndex,
      colName,
    });
  } catch (e) {
    const msg = String(e?.message || e);
    log(`${state.lang === "en" ? "Pick failed" : "点选启动失败"}：${msg}`);
    setStatus({ stateText: state.lang === "en" ? "Ready" : "就绪", msg });
    state.pick.active = false;
    state.pick.mode = null;
    updatePickIndicator();
    updateIdentityIndicator();
  }
}

async function cancelPick() {
  try {
    await safeSendToActiveTab({ type: "FB_CANCEL_PICK" });
  } catch (_) {
    // ignore cancel errors
  }
  state.pick.active = false;
  state.pick.mode = null;
  updatePickIndicator();
  updateIdentityIndicator();
  setStatus({
    stateText: state.lang === "en" ? "Ready" : "就绪",
    msg: state.lang === "en" ? "Pick cancelled." : "已取消点选。",
  });
  log(state.lang === "en" ? "Pick cancelled." : "点选已取消。");
}

chrome.runtime.onMessage.addListener((msg) => {
  if (msg?.type === "FB_PICK_RESULT") {
    const { mode, colIndex, colName, selector, meta } = msg.payload;
    state.pick.active = false;
    state.pick.mode = null;
    updatePickIndicator();
    updateIdentityIndicator();

    if (mode === "field") {
      state.mapping = state.mapping.filter((m) => m.colIndex !== colIndex);
      state.mapping.push({ col: colName, colIndex, selector, meta });
      refreshMappingUI();
      if (state.data.headers.length > 0) {
        const nextIndex = Math.min(colIndex + 1, state.data.headers.length - 1);
        ui.colSelect.value = String(nextIndex);
        updatePickIndicator();
      }
      log(state.lang === "en" ? `Bound: ${colName} -> ${selector}` : `已绑定：${colName} -> ${selector}`);
      setStatus({ stateText: state.lang === "en" ? "Ready" : "就绪", msg: state.lang === "en" ? "Field bound." : "字段已绑定。" });
      return;
    }

    if (mode === "identity") {
      state.identityMapping = {
        colIndex,
        colName,
        selector,
        meta,
        matchStrategy: ui.identityStrategy?.value || "exact",
      };
      updateIdentityUI();
      if (ui.identityStatus) ui.identityStatus.textContent = state.lang === "en" ? "Identity field bound." : "身份字段已绑定。";
      log(state.lang === "en" ? `Identity bound: ${colName} -> ${selector}` : `身份已绑定：${colName} -> ${selector}`);
    }
  }

  if (msg?.type === "FB_PICK_CANCELLED") {
    setStatus({ stateText: state.lang === "en" ? "Ready" : "就绪", msg: state.lang === "en" ? "Pick cancelled." : "已取消点选。" });
    log(state.lang === "en" ? "Pick cancelled." : "点选已取消。");
    state.pick.active = false;
    state.pick.mode = null;
    updatePickIndicator();
    updateIdentityIndicator();
  }

  if (msg?.type === "FB_RUN_EVENT") {
    log(msg.payload?.line || (state.lang === "en" ? "Run event" : "运行事件"));
  }
});

// -------- Runner --------
function normalizeCell(v) {
  if (v === null || v === undefined) return "";
  return String(v);
}

async function runBatch({ stepOnly = false }) {
  if (!state.data.rows.length) {
    log(state.lang === "en" ? "Please load a CSV/XLSX file." : "请先载入 CSV/XLSX。" );
    return;
  }
  if (!state.mapping.length) {
    log(state.lang === "en" ? "Please configure field mapping." : "请先配置字段映射。" );
    return;
  }

  state.runner.running = true;

  const { start, rows } = getRowSlice();
  const total = rows.length;

  const current = state.selection.currentRowIndex ?? start;
  if (current < start || current > start + total - 1) {
    state.selection.currentRowIndex = start;
  }
  const idx = (state.selection.currentRowIndex ?? start) - start;
  const globalRowIndex = start + idx;
  const row = rows[idx];

  setStatus({
    stateText: state.lang === "en" ? "Step" : "单步",
    progText: `${idx + 1} / ${total}`,
    rowText: state.lang === "en" ? `Row ${globalRowIndex}` : `第${globalRowIndex}行`,
    msg: state.lang === "en" ? "Filling..." : "填充中…",
  });
  renderRowPreview(row, globalRowIndex);
  const preview = document.querySelector(".row-preview");
  if (preview) {
    preview.classList.add("is-pending");
    preview.classList.remove("is-done");
  }

  const rowData = {};
  state.data.headers.forEach((_, i) => (rowData[i] = normalizeCell(row[i] ?? "")));

  const payload = {
    mapping: state.mapping,
    rowDataByColIndex: rowData,
  };

  try {
    const res = await safeSendToActiveTab({ type: "FB_FILL_ONE", payload });
    if (res?.ok) {
      state.runner.results.push({ i: globalRowIndex, status: "OK", msg: res.message || "" });
      state.runner.lastFill = { i: globalRowIndex, status: "OK" };
      if (preview) {
        preview.classList.add("is-done");
        preview.classList.remove("is-pending");
      }
      setStatus({ msg: state.lang === "en" ? "Success" : "成功" });
      log(state.lang === "en" ? `✓ Row ${globalRowIndex} success` : `✓ 第${globalRowIndex}行成功`);
    } else {
      const err = res?.error || (state.lang === "en" ? "Unknown error" : "未知错误");
      state.runner.results.push({ i: globalRowIndex, status: "FAIL", msg: err });
      state.runner.lastFill = { i: globalRowIndex, status: "FAIL" };
      if (preview) {
        preview.classList.add("is-pending");
        preview.classList.remove("is-done");
      }
      setStatus({ msg: (state.lang === "en" ? "Failed: " : "失败：") + err });
      log(state.lang === "en" ? `✗ Row ${globalRowIndex} failed: ${err}` : `✗ 第${globalRowIndex}行失败：${err}`);
    }
  } catch (e) {
    const err = String(e?.message || e);
    state.runner.results.push({ i: globalRowIndex, status: "FAIL", msg: err });
    state.runner.lastFill = { i: globalRowIndex, status: "FAIL" };
    if (preview) {
      preview.classList.add("is-pending");
      preview.classList.remove("is-done");
    }
    setStatus({ msg: (state.lang === "en" ? "Failed: " : "失败：") + err });
    log(state.lang === "en" ? `✗ Row ${globalRowIndex} failed (comm error): ${err}` : `✗ 第${globalRowIndex}行失败（通信异常）：${err}`);
  }

  state.runner.running = false;
  setStatus({ stateText: state.lang === "en" ? "Ready" : "就绪", msg: state.lang === "en" ? "Done" : "完成" });

  if (stepOnly && ui.autoAdvance.checked) {
    const next = globalRowIndex + 1;
    if (next <= start + total - 1) {
      state.selection.currentRowIndex = next;
      renderRowPreview(rows[idx + 1] || [], next);
      const preview = document.querySelector(".row-preview");
      if (preview) {
        preview.classList.add("is-pending");
        preview.classList.remove("is-done");
      }
    }
  }
}

function goNextRowPreview() {
  if (!state.data.rows.length) return;
  const { start, rows } = getRowSlice();
  const total = rows.length;
  if (!total) return;
  const current = state.selection.currentRowIndex ?? start;
  let next = current + 1;
  const maxRow = start + total - 1;
  if (next > maxRow) next = maxRow;
  state.selection.currentRowIndex = next;
  const idx = next - start;
  renderRowPreview(rows[idx] || [], next);
  const preview = document.querySelector(".row-preview");
  if (preview) {
    preview.classList.add("is-pending");
    preview.classList.remove("is-done");
  }
}

async function init() {
  setupModuleLayout();
  setupCardToggles();
  setupGroupToggles();

  const stored = await chrome.storage.local.get(["ui:theme", "ui:lang"]);
  const theme = stored["ui:theme"] || "dark";
  const lang = stored["ui:lang"] || "zh";
  setTheme(theme);
  setLang(lang);

  refreshSheetUI();
  refreshMappingUI();
  setStatus({ stateText: lang === "en" ? "Ready" : "就绪", progText: "0 / 0", rowText: "—", msg: "—" });
  updatePickIndicator();
  updateFillIndicator();
  updateIdentityUI();
  updateIdentityIndicator();

  ui.fileInput.addEventListener("change", async () => {
    const f = ui.fileInput.files?.[0];
    if (!f) return;
    try {
      if (state.data.filename) clearFileSelection();
      await loadFile(f);
    } catch (e) {
      log((state.lang === "en" ? "Load failed: " : "载入失败：") + String(e?.message || e));
    }
  });

  ui.fileInput.addEventListener("click", () => {
    ui.fileInput.value = "";
  });

  ui.btnClearFile.addEventListener("click", () => {
    clearFileSelection();
  });

  ui.sheetSelect.addEventListener("change", async () => {
    try {
      await reloadActiveSheetFromCurrentFile();
    } catch (e) {
      log((state.lang === "en" ? "Switch sheet failed: " : "切换 Sheet 失败：") + String(e?.message || e));
    }
  });

  ui.colSelect.addEventListener("change", () => {
    updatePickIndicator();
  });
  if (ui.identityColSelect) {
    ui.identityColSelect.addEventListener("change", () => {
      updateIdentityIndicator();
      if (state.identityMapping) {
        state.identityMapping.colIndex = parseInt(ui.identityColSelect.value || "0", 10);
        state.identityMapping.colName = state.data.headers[state.identityMapping.colIndex] || state.identityMapping.colName;
      }
    });
  }

  ui.btnPick.addEventListener("click", () => beginPick("field"));
  if (ui.btnPickIdentity) ui.btnPickIdentity.addEventListener("click", () => beginPick("identity"));

  ui.btnSaveMap.addEventListener("click", saveSiteConfig);
  ui.btnClearMap.addEventListener("click", () => {
    state.mapping = [];
    refreshMappingUI();
    log(state.lang === "en" ? "Mapping cleared." : "已清空映射。");
  });

  ui.btnLoadSite.addEventListener("click", loadSiteConfig);
  ui.btnViewSites.addEventListener("click", async () => {
    const show = ui.siteList.style.display === "none";
    if (show) await renderSiteList();
    ui.siteList.style.display = show ? "block" : "none";
    ui.btnViewSites.classList.toggle("is-active", show);
  });

  if (ui.identityStrategy) {
    ui.identityStrategy.addEventListener("change", () => {
      if (!state.identityMapping) return;
      state.identityMapping.matchStrategy = ui.identityStrategy.value;
    });
  }
  if (ui.btnMatchIdentity) ui.btnMatchIdentity.addEventListener("click", matchIdentityRow);

  ui.btnFilter.addEventListener("click", () => {
    ui.filterTitle.style.display = "block";
    const fieldIndex = parseInt(ui.filterField.value || "0", 10);
    const keyword = String(ui.filterKeyword.value || "").trim().toLowerCase();
    const { start, rows } = getRowSlice();
    const hits = [];
    rows.forEach((row, i) => {
      const cell = String(row[fieldIndex] ?? "").toLowerCase();
      if (!keyword || cell.includes(keyword)) {
        hits.push({ excelRow: start + i, row });
      }
    });

    ui.filterList.innerHTML = "";
    if (hits.length === 0) {
      const empty = document.createElement("div");
      empty.className = "site-meta";
      empty.textContent = state.lang === "en" ? "No matches." : "未找到匹配行。";
      ui.filterList.appendChild(empty);
      ui.filterList.style.display = "block";
      return;
    }

    hits.slice(0, 50).forEach((hit) => {
      const item = document.createElement("div");
      item.className = "filter-item";
      item.addEventListener("click", () => {
        state.selection.currentRowIndex = hit.excelRow;
        renderRowPreview(hit.row, hit.excelRow);
        ui.filterList.style.display = "none";
      });

      const num = document.createElement("div");
      num.className = "filter-rownum";
      num.textContent = state.lang === "en" ? `Row ${hit.excelRow}` : `第${hit.excelRow}行`;

      const text = document.createElement("div");
      text.className = "filter-text";
      text.textContent = String(hit.row[fieldIndex] ?? "");

      item.appendChild(num);
      item.appendChild(text);
      ui.filterList.appendChild(item);
    });

    ui.filterList.style.display = "block";
  });

  ui.btnStep.addEventListener("click", () => runBatch({ stepOnly: true }));
  if (ui.btnNextRow) ui.btnNextRow.addEventListener("click", goNextRowPreview);

  ui.btnHelp.addEventListener("click", () => {
    const show = ui.helpCard.style.display === "none";
    ui.helpCard.style.display = show ? "block" : "none";
  });

  ui.btnSettings.addEventListener("click", (e) => {
    e.stopPropagation();
    const show = ui.settingsPanel.style.display === "none";
    ui.settingsPanel.style.display = show ? "block" : "none";
  });

  document.addEventListener("click", (e) => {
    if (!ui.settingsPanel || ui.settingsPanel.style.display === "none") return;
    if (ui.settingsPanel.contains(e.target) || ui.btnSettings.contains(e.target)) return;
    ui.settingsPanel.style.display = "none";
  });

  document.querySelectorAll(".seg-btn[data-theme-val]").forEach((btn) => {
    btn.addEventListener("click", () => setTheme(btn.getAttribute("data-theme-val")));
  });
  document.querySelectorAll(".seg-btn[data-lang]").forEach((btn) => {
    btn.addEventListener("click", () => {
      setLang(btn.getAttribute("data-lang"));
      refreshMappingUI();
    });
  });

  $("#openOptions").addEventListener("click", async (e) => {
    e.preventDefault();
    const key = await computeSiteKey();
    log((state.lang === "en" ? "Current site key: " : "当前站点键：") + key);
  });

  log(state.lang === "en" ? "FormBatcher MVP ready. Open target form page." : "FormBatcher MVP 已就绪。打开目标表单页面后开始配置。");
}

init();
