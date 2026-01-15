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
  stState: $("#stState"),
  stProg: $("#stProg"),
  stRow: $("#stRow"),
  stMsg: $("#stMsg"),
  btnExport: $("#btnExport"),
  logBox: $("#logBox"),
  rowPreview: $("#rowPreview"),
  rowPreviewLabel: $("#rowPreviewLabel"),
  rowPreviewHead: $("#rowPreviewHead"),
  rowPreviewBody: $("#rowPreviewBody"),
  btnTheme: $("#btnTheme"),
  btnHelp: $("#btnHelp"),
  helpCard: $("#helpCard"),
  btnClearFile: $("#btnClearFile"),
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
    paused: false,
    stop: false,
    cursor: 0,
    results: [], // {i, status, msg}
    siteKey: null,
  },
  pick: {
    active: false,
  },
};

function log(line) {
  const t = new Date().toLocaleTimeString();
  ui.logBox.textContent = `[${t}] ${line}\n` + ui.logBox.textContent;
}

function setStatus({ stateText, progText, rowText, msg }) {
  if (stateText) ui.stState.textContent = stateText;
  if (progText) ui.stProg.textContent = progText;
  if (rowText) ui.stRow.textContent = rowText;
  if (msg) ui.stMsg.textContent = msg;
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
  state.data.headers.forEach((h, idx) => {
    const opt = document.createElement("option");
    opt.value = String(idx);
    opt.textContent = h || `列${idx + 1}`;
    ui.colSelect.appendChild(opt);
  });
  ui.pillCols.textContent = `${state.data.headers.length} 列`;
  ui.pillRows.textContent = `${state.data.rows.length} 行`;
  updatePickIndicator();
}

function updatePickIndicator() {
  const hasSelection = state.data.headers.length > 0 && String(ui.colSelect.value || "").length > 0;
  const ready = hasSelection && !state.pick.active;
  ui.btnPick.classList.toggle("pick-ready", ready);
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
  } else {
    ui.rowStart.value = "2";
    ui.rowEnd.value = "2";
  }
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
    empty.textContent = "暂无映射。选择一列并点选网页中的输入框。";
    ui.mapList.appendChild(empty);
    return;
  }

  state.mapping.forEach((m, idx) => {
    const wrap = document.createElement("div");
    wrap.className = "mapitem";

    const badge = document.createElement("div");
    badge.className = "badge";
    badge.textContent = `${m.col || "未命名列"}`;

    const meta = document.createElement("div");
    meta.className = "mapmeta";
    const t = document.createElement("div");
    t.className = "t";
    t.textContent = `${m.meta?.label || m.meta?.placeholder || m.meta?.name || m.meta?.tag || "字段"}  ←  ${m.col}`;
    const d = document.createElement("div");
    d.className = "d";
    d.textContent = m.selector;

    meta.appendChild(t);
    meta.appendChild(d);

    const x = document.createElement("button");
    x.className = "x";
    x.textContent = "移除";
    x.addEventListener("click", () => {
      state.mapping.splice(idx, 1);
      refreshMappingUI();
    });

    wrap.appendChild(badge);
    wrap.appendChild(meta);
    wrap.appendChild(x);
    ui.mapList.appendChild(wrap);
  });
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
  if (rowIndex) {
    ui.rowPreviewLabel.textContent = `当前行号 ${rowIndex}，预览`;
  } else {
    ui.rowPreviewLabel.textContent = "当前行号 —，预览";
  }
  if (!row || !row.length) {
    ui.rowPreviewHead.textContent = "—";
    ui.rowPreviewBody.textContent = "—";
    return;
  }
  const headers = state.data.headers.map((h, i) => h || `列${i + 1}`);
  const body = row.map((v) => String(v ?? "").trim());
  const cols = Math.max(headers.length, body.length);

  ui.rowPreviewHead.innerHTML = headers.map((h) => `<span class="row-preview-cell">${h}</span>`).join("");
  ui.rowPreviewBody.innerHTML = body.map((v) => `<span class="row-preview-cell">${v || "—"}</span>`).join("");

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
    log(`已加载 CSV：${state.data.headers.length} 列，${state.data.rows.length} 行`);
    return;
  }

  if (ext === "xlsx") {
    if (!window.XLSX) throw new Error("缺少 XLSX 解析库，请重新加载扩展。");
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    state.data.workbook = workbook;
    state.data.sheets = workbook.SheetNames || [];
    state.data.activeSheet = state.data.sheets[0] || "";
    refreshSheetUI();
    await reloadActiveSheetFromCurrentFile();
    renderRowPreview(state.data.rows[0] || [], 2);
    log(`已加载 XLSX：${state.data.sheets.length} 个 Sheet`);
    return;
  }

  throw new Error("仅支持 .csv 或 .xlsx");
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
  ui.fileNameText.textContent = "未选择";
  ui.fileNameText.parentElement?.classList.remove("has-file");
  ui.fileInput.value = "";
  refreshSheetUI();
  refreshColsUI();
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
  const payload = { mapping: state.mapping };
  await chrome.storage.local.set({ ["site:" + key]: payload });
  log("已保存站点配置：" + key);
}

async function loadSiteConfig() {
  const key = await computeSiteKey();
  state.runner.siteKey = key;
  const data = await chrome.storage.local.get("site:" + key);
  const payload = data?.["site:" + key];
  if (!payload) {
    log("该站点暂无保存配置：" + key);
    return;
  }
  state.mapping = payload.mapping || [];
  refreshMappingUI();
  log("已载入站点配置：" + key);
}

async function loadSiteConfigByKey(key) {
  const data = await chrome.storage.local.get("site:" + key);
  const payload = data?.["site:" + key];
  if (!payload) {
    log("未找到站点配置：" + key);
    return;
  }
  state.mapping = payload.mapping || [];
  refreshMappingUI();
  log("已载入站点配置：" + key);
}

async function renderSiteList() {
  const data = await chrome.storage.local.get(null);
  const keys = Object.keys(data).filter((k) => k.startsWith("site:"));
  ui.siteList.innerHTML = "";

  if (keys.length === 0) {
    const empty = document.createElement("div");
    empty.className = "site-meta";
    empty.textContent = "暂无保存的站点配置。";
    ui.siteList.appendChild(empty);
    return;
  }

  keys.sort();
  keys.forEach((k) => {
    const key = k.replace(/^site:/, "");
    const payload = data[k] || {};
    const mappingCount = payload.mapping?.length || 0;

    const row = document.createElement("div");
    row.className = "site-item";

    const label = document.createElement("div");
    label.textContent = key;

    const meta = document.createElement("div");
    meta.className = "site-meta";
    meta.textContent = `映射 ${mappingCount}`;

    const actions = document.createElement("div");
    actions.className = "site-actions";

    const btnLoad = document.createElement("button");
    btnLoad.className = "btn ghost";
    btnLoad.textContent = "载入";
    btnLoad.addEventListener("click", () => loadSiteConfigByKey(key));

    actions.appendChild(btnLoad);
    row.appendChild(label);
    row.appendChild(meta);
    row.appendChild(actions);
    ui.siteList.appendChild(row);
  });
}

// -------- Picking elements on the page --------
async function beginPick(mode) {
  const tab = await getActiveTab();
  if (!tab?.id) return;

  const colIndex = parseInt(ui.colSelect.value || "0", 10);
  const colName = state.data.headers[colIndex] ?? `列${colIndex + 1}`;

  setStatus({ stateText: "点选中", msg: "请到网页上点击目标元素…" });
  log(`进入点选模式 (${mode})：请点击网页元素。`);

  state.pick.active = true;
  updatePickIndicator();

  try {
    await safeSendToActiveTab({
      type: "FB_BEGIN_PICK",
      mode,
      colIndex,
      colName,
    });
  } catch (e) {
    const msg = String(e?.message || e);
    log(`点选启动失败：${msg}`);
    setStatus({ stateText: "就绪", msg });
    state.pick.active = false;
    updatePickIndicator();
  }
}

chrome.runtime.onMessage.addListener((msg) => {
  if (msg?.type === "FB_PICK_RESULT") {
    const { mode, colIndex, colName, selector, meta } = msg.payload;
    if (mode !== "field") return;

    state.pick.active = false;
    updatePickIndicator();

    state.mapping = state.mapping.filter((m) => m.colIndex !== colIndex);
    state.mapping.push({ col: colName, colIndex, selector, meta });
    refreshMappingUI();
    log(`已绑定：${colName} -> ${selector}`);
    setStatus({ stateText: "就绪", msg: "字段已绑定。" });
  }

  if (msg?.type === "FB_PICK_CANCELLED") {
    setStatus({ stateText: "就绪", msg: "已取消点选。" });
    log("点选已取消。");
    state.pick.active = false;
    updatePickIndicator();
  }

  if (msg?.type === "FB_RUN_EVENT") {
    log(msg.payload?.line || "运行事件");
  }
});

// -------- Runner --------
function normalizeCell(v) {
  if (v === null || v === undefined) return "";
  return String(v);
}

async function runBatch({ stepOnly = false }) {
  if (!state.data.rows.length) {
    log("请先载入 CSV/XLSX。");
    return;
  }
  if (!state.mapping.length) {
    log("请先配置字段映射。");
    return;
  }

  state.runner.stop = false;
  state.runner.paused = false;
  state.runner.running = true;

  const { start, rows } = getRowSlice();
  const total = rows.length;

  renderRowPreview(rows[0] || [], start);

  setStatus({
    stateText: "单步",
    progText: `0 / ${total}`,
    msg: "开始执行…",
  });
  log(`单步填充（从第${start}行，共${total}行）`);

  const tab = await getActiveTab();
  if (!tab?.id) {
    log("未找到活动标签页");
    return;
  }

  for (let i = 0; i < total; i++) {
    const globalRowIndex = start - 1 + i;
    const row = rows[i];

    renderRowPreview(row, globalRowIndex + 1);

    setStatus({
      progText: `${i + 1} / ${total}`,
      rowText: `第${globalRowIndex + 1}行`,
      msg: "填充中…",
    });

    const rowData = {};
    state.data.headers.forEach((_, idx) => (rowData[idx] = normalizeCell(row[idx] ?? "")));

    const payload = {
      mapping: state.mapping,
      rowDataByColIndex: rowData,
    };

    try {
      const res = await safeSendToActiveTab({ type: "FB_FILL_ONE", payload });
      if (res?.ok) {
        state.runner.results.push({ i: globalRowIndex + 1, status: "OK", msg: res.message || "" });
        setStatus({ msg: "成功" });
        log(`✓ 第${globalRowIndex + 1}行成功`);
      } else {
        state.runner.results.push({ i: globalRowIndex + 1, status: "FAIL", msg: res?.error || "未知错误" });
        setStatus({ msg: "失败：" + (res?.error || "未知错误") });
        log(`✗ 第${globalRowIndex + 1}行失败：${res?.error || "未知错误"}`);
      }
    } catch (e) {
      state.runner.results.push({ i: globalRowIndex + 1, status: "FAIL", msg: String(e?.message || e) });
      setStatus({ msg: "失败：" + String(e?.message || e) });
      log(`✗ 第${globalRowIndex + 1}行失败（通信异常）：${String(e?.message || e)}`);
    }

    if (stepOnly) break;
  }

  state.runner.running = false;
  setStatus({ stateText: "就绪", msg: "完成" });
  log("完成。");
}

function exportResultsCSV() {
  const lines = [];
  lines.push(["row", "status", "message"].join(","));
  state.runner.results.forEach((r) => {
    const safe = (s) => `"${String(s ?? "").replaceAll('"', '""')}"`;
    lines.push([r.i, r.status, safe(r.msg)].join(","));
  });
  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "formbatcher_results.csv";
  a.click();
  URL.revokeObjectURL(url);
  log("已导出结果 CSV。");
}

function toggleTheme() {
  const cur = ui.app.getAttribute("data-theme") || "dark";
  const next = cur === "dark" ? "light" : "dark";
  ui.app.setAttribute("data-theme", next);
  ui.btnTheme.querySelector(".ico").textContent = next === "dark" ? "☾" : "☀";
  chrome.storage.local.set({ "ui:theme": next });
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
    btn.addEventListener("click", () => {
      if (!card) return;
      const collapsed = card.classList.toggle("collapsed");
      btn.setAttribute("aria-expanded", collapsed ? "false" : "true");
      btn.textContent = collapsed ? ">" : "v";
    });
  });
}

async function init() {
  setupModuleLayout();
  setupCardToggles();

  const t = await chrome.storage.local.get("ui:theme");
  const theme = t["ui:theme"] || "dark";
  ui.app.setAttribute("data-theme", theme);
  ui.btnTheme.querySelector(".ico").textContent = theme === "dark" ? "☾" : "☀";

  refreshSheetUI();
  refreshMappingUI();
  setStatus({ stateText: "就绪", progText: "0 / 0", rowText: "—", msg: "—" });

  ui.fileInput.addEventListener("change", async () => {
    const f = ui.fileInput.files?.[0];
    if (!f) return;
    try {
      await loadFile(f);
    } catch (e) {
      log("载入失败：" + String(e?.message || e));
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
      log("切换 Sheet 失败：" + String(e?.message || e));
    }
  });

  ui.colSelect.addEventListener("change", () => {
    updatePickIndicator();
  });

  ui.btnPick.addEventListener("click", () => beginPick("field"));

  ui.btnSaveMap.addEventListener("click", saveSiteConfig);
  ui.btnClearMap.addEventListener("click", () => {
    state.mapping = [];
    refreshMappingUI();
    log("已清空映射。");
  });

  ui.btnLoadSite.addEventListener("click", loadSiteConfig);
  ui.btnViewSites.addEventListener("click", async () => {
    const show = ui.siteList.style.display === "none";
    if (show) await renderSiteList();
    ui.siteList.style.display = show ? "block" : "none";
  });

  ui.btnStep.addEventListener("click", () => runBatch({ stepOnly: true }));
  ui.btnExport.addEventListener("click", exportResultsCSV);
  ui.btnTheme.addEventListener("click", toggleTheme);

  ui.btnHelp.addEventListener("click", () => {
    const show = ui.helpCard.style.display === "none";
    ui.helpCard.style.display = show ? "block" : "none";
  });

  $("#openOptions").addEventListener("click", async (e) => {
    e.preventDefault();
    const key = await computeSiteKey();
    log("当前站点键：" + key);
  });

  log("FormBatcher MVP 已就绪。打开目标表单页面后开始配置。");
}

init();
