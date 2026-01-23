/**
 * FormBatcher MVP Content Script
 * - Pick mode: user clicks an element, we compute a stable selector and return it.
 * - Fill mode: receive mapping + row values, fill DOM elements, trigger events.
 */

let pickSession = null; // {mode, colIndex, colName}
let hoverBox = null;
let pickLabel = null;
let hoverRAF = null;
let lastHoverTarget = null;
const choiceRootCache = new WeakMap();
let filledHighlights = new Set();
let filledClearHandler = null;
let filledStyleTag = null;

function cssEscapeIdent(ident) {
  return ident.replace(/([ !"#$%&'()*+,./:;<=>?@[\]^`{|}~\\])/g, "\\$1");
}

function isFillable(el) {
  if (!el) return false;
  const tag = el.tagName?.toLowerCase();
  if (tag === "input" || tag === "textarea" || tag === "select") return true;
  if (el.isContentEditable) return true;
  const role = el.getAttribute?.("role");
  if (role === "textbox" || role === "combobox") return true;
  return false;
}

function isChoiceInput(el) {
  const tag = el?.tagName?.toLowerCase();
  if (tag !== "input") return false;
  const type = (el.getAttribute?.("type") || "").toLowerCase();
  return type === "checkbox" || type === "radio";
}

function findChoiceInputFromTarget(target) {
  if (!target) return null;
  if (isChoiceInput(target)) return target;
  const inLabel = target.closest?.("label");
  if (inLabel) {
    const input = inLabel.querySelector('input[type="radio"], input[type="checkbox"]');
    if (input) return input;
  }
  const nearby = target.closest?.('input[type="radio"], input[type="checkbox"]');
  if (nearby) return nearby;
  const parent = target.parentElement;
  if (parent) {
    const input = parent.querySelector?.('input[type="radio"], input[type="checkbox"]');
    if (input) return input;
  }
  return null;
}

function findLabelText(el) {
  try {
    const id = el.getAttribute?.("id");
    if (id) {
      const lab = document.querySelector(`label[for="${CSS.escape(id)}"]`);
      if (lab) return lab.innerText.trim();
    }
    let cur = el;
    for (let i = 0; i < 4 && cur; i++) {
      const lab = cur.closest?.("label");
      if (lab && lab.innerText) return lab.innerText.trim();
      cur = cur.parentElement;
    }
  } catch (_) {}
  return "";
}

function getOptionLabel(el) {
  const label = findLabelText(el);
  if (label) return label;
  const txt = (el.closest("label")?.innerText || "").trim();
  if (txt) return txt;
  return (el.getAttribute?.("aria-label") || "").trim();
}

function findChoiceGroupRoot(input) {
  if (!input) return null;
  const cached = choiceRootCache.get(input);
  if (cached) return cached;
  const type = (input.getAttribute?.("type") || "").toLowerCase();
  const name = input.getAttribute?.("name");
  const candidates = [
    "fieldset",
    ".radio-group",
    ".checkbox-group",
    ".form-group",
    ".form-row",
    ".field",
    ".group",
    ".item",
    ".row",
  ];
  for (const sel of candidates) {
    const root = input.closest(sel);
    if (!root) continue;
    const items = root.querySelectorAll(`input[type="${type}"]${name ? `[name="${CSS.escape(name)}"]` : ""}`);
    if (items.length > 1) {
      choiceRootCache.set(input, root);
      return root;
    }
  }
  const fallback = input.parentElement || input;
  choiceRootCache.set(input, fallback);
  return fallback;
}

function buildSelector(el) {
  if (!el || el === document.body) return null;
  const tag = el.tagName.toLowerCase();

  const id = el.getAttribute("id");
  if (id && id.length <= 64) return `${tag}#${CSS.escape(id)}`;

  const name = el.getAttribute("name");
  if (name && name.length <= 80) return `${tag}[name="${CSS.escape(name)}"]`;

  const aria = el.getAttribute("aria-label");
  if (aria && aria.length <= 80) return `${tag}[aria-label="${CSS.escape(aria)}"]`;

  const parts = [];
  let cur = el;
  for (let depth = 0; depth < 5 && cur && cur !== document.documentElement; depth++) {
    const t = cur.tagName.toLowerCase();
    let part = t;

    const cls = (cur.getAttribute("class") || "").trim().split(/\s+/).filter(Boolean)
      .filter((c) => !c.startsWith("css-") && c.length < 24)
      .slice(0, 2);
    if (cls.length) part += "." + cls.map((c) => cssEscapeIdent(c)).join(".");

    const parent = cur.parentElement;
    if (parent) {
      const sibs = Array.from(parent.children).filter((x) => x.tagName.toLowerCase() === t);
      if (sibs.length > 1) {
        const idx = sibs.indexOf(cur) + 1;
        part += `:nth-of-type(${idx})`;
      }
    }
    parts.unshift(part);
    const sel = parts.join(" > ");
    try {
      if (document.querySelectorAll(sel).length === 1) return sel;
    } catch (_) {}
    cur = parent;
  }
  return parts.join(" > ");
}

function highlight(el) {
  if (!el) return;
  const r = el.getBoundingClientRect();
  const box = document.createElement("div");
  box.style.position = "fixed";
  box.style.left = `${r.left}px`;
  box.style.top = `${r.top}px`;
  box.style.width = `${r.width}px`;
  box.style.height = `${r.height}px`;
  box.style.zIndex = "2147483647";
  box.style.pointerEvents = "none";
  box.style.borderRadius = "10px";
  box.style.border = "2px solid rgba(122,162,255,.95)";
  box.style.boxShadow = "0 0 0 6px rgba(122,162,255,.18)";
  box.style.background = "rgba(122,162,255,.06)";
  document.documentElement.appendChild(box);
  setTimeout(() => box.remove(), 800);
}

function ensureFilledStyle() {
  if (filledStyleTag) return;
  filledStyleTag = document.createElement("style");
  filledStyleTag.textContent = `
    .fb-filled-highlight{
      outline: 2px solid rgba(122,162,255,.95) !important;
      outline-offset: 2px;
      border-radius: 10px;
    }
  `;
  document.documentElement.appendChild(filledStyleTag);
}

function clearFilledHighlights() {
  if (!filledHighlights.size) return;
  filledHighlights.forEach((el) => el.classList.remove("fb-filled-highlight"));
  filledHighlights = new Set();
  if (filledClearHandler) {
    document.removeEventListener("mousedown", filledClearHandler, true);
    filledClearHandler = null;
  }
}

function markFilled(el) {
  if (!el) return;
  ensureFilledStyle();
  el.classList.add("fb-filled-highlight");
  filledHighlights.add(el);
  if (!filledClearHandler) {
    filledClearHandler = (e) => {
      if (pickSession) return;
      for (const item of filledHighlights) {
        if (item.contains(e.target)) return;
      }
      clearFilledHighlights();
    };
    document.addEventListener("mousedown", filledClearHandler, true);
  }
}

function ensureHoverBox() {
  if (hoverBox) return hoverBox;
  hoverBox = document.createElement("div");
  hoverBox.style.position = "fixed";
  hoverBox.style.zIndex = "2147483646";
  hoverBox.style.pointerEvents = "none";
  hoverBox.style.borderRadius = "10px";
  hoverBox.style.border = "2px solid rgba(94,154,255,.9)";
  hoverBox.style.boxShadow = "0 0 0 6px rgba(94,154,255,.18)";
  hoverBox.style.background = "rgba(94,154,255,.08)";
  hoverBox.style.transition = "all 60ms ease";
  document.documentElement.appendChild(hoverBox);
  return hoverBox;
}

function updateHoverBox(el) {
  if (!el) return;
  const r = el.getBoundingClientRect();
  const box = ensureHoverBox();
  box.style.left = `${r.left}px`;
  box.style.top = `${r.top}px`;
  box.style.width = `${r.width}px`;
  box.style.height = `${r.height}px`;
}

function clearHoverBox() {
  if (!hoverBox) return;
  hoverBox.remove();
  hoverBox = null;
}

function setValue(el, value) {
  const tag = el.tagName?.toLowerCase();
  const type = (el.getAttribute?.("type") || "").toLowerCase();

  if (el.isContentEditable) {
    el.focus();
    document.execCommand("selectAll", false, null);
    document.execCommand("insertText", false, value);
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    el.blur();
    return;
  }

  if (tag === "input") {
    if (type === "checkbox") {
      const normalized = String(value).trim().toLowerCase();
      const truthyValues = ["1", "true", "yes", "y", "on", "对", "是"];
      const truthy = truthyValues.includes(normalized);
      el.checked = truthy;
      el.dispatchEvent(new Event("change", { bubbles: true }));
      return;
    }
    if (type === "radio") {
      if (String(el.value) === String(value)) {
        el.checked = true;
        el.dispatchEvent(new Event("change", { bubbles: true }));
      }
      return;
    }
    el.focus();
    const proto = Object.getPrototypeOf(el);
    const desc = Object.getOwnPropertyDescriptor(proto, "value");
    if (desc && desc.set) desc.set.call(el, value);
    else el.value = value;
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    el.blur();
    return;
  }

  if (tag === "textarea") {
    el.focus();
    el.value = value;
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    el.blur();
    return;
  }

  if (tag === "select") {
    const valStr = String(value);
    const opts = Array.from(el.options || []);
    const byVal = opts.find((o) => String(o.value) === valStr);
    const byText = opts.find((o) => String(o.textContent || "").trim() === valStr.trim());
    if (byVal) el.value = byVal.value;
    else if (byText) el.value = byText.value;
    else el.value = valStr;
    el.dispatchEvent(new Event("change", { bubbles: true }));
    return;
  }

  if (el.getAttribute?.("role") === "textbox" || el.getAttribute?.("role") === "combobox") {
    el.focus();
    el.textContent = value;
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    el.blur();
  }
}

function setChoiceGroup(container, meta, rawValue) {
  if (!container) return;
  const type = meta?.groupType;
  const name = meta?.groupName;
  let inputs = [];
  if (container.tagName?.toLowerCase() === "input" && isChoiceInput(container)) {
    const q = `input[type="${type}"]${name ? `[name="${CSS.escape(name)}"]` : ""}`;
    inputs = Array.from(document.querySelectorAll(q));
  } else {
    const q = `input[type="${type}"]${name ? `[name="${CSS.escape(name)}"]` : ""}`;
    inputs = Array.from(container.querySelectorAll(q));
  }
  if (!inputs.length) return;

  const value = String(rawValue ?? "").trim();
  const tokens = type === "checkbox"
    ? value.split(/[|,，、]/).map((v) => v.trim()).filter(Boolean)
    : [value];
  const truthyValues = new Set(["1", "true", "yes", "y", "on", "对", "是"]);
  const hasTruthy = tokens.some((t) => truthyValues.has(t.toLowerCase()));

  inputs.forEach((input) => {
    const optVal = String(input.value ?? "").trim();
    const optLabel = getOptionLabel(input);
    const matches = tokens.some((t) => t === optVal || t === optLabel);
    if (type === "checkbox") {
      const useTruthy = !matches && inputs.length === 1 && hasTruthy;
      input.checked = matches || useTruthy;
      input.dispatchEvent(new Event("change", { bubbles: true }));
    } else if (type === "radio") {
      if (matches) {
        input.checked = true;
        input.dispatchEvent(new Event("change", { bubbles: true }));
      }
    }
  });
}

function fillOne(payload) {
  const { mapping, rowDataByColIndex } = payload;
  for (const m of mapping) {
    const val = rowDataByColIndex?.[m.colIndex] ?? "";
    const selector = m.selector;
    let el = null;
    try {
      el = document.querySelector(selector);
    } catch (e) {
      return { ok: false, error: `无效选择器：${selector}` };
    }
    if (!el) return { ok: false, error: `未找到元素：${selector}` };
    try {
      if (m.meta?.groupType) {
        setChoiceGroup(el, m.meta, val);
      } else {
        setValue(el, val);
      }
      markFilled(el);
    } catch (e) {
      return { ok: false, error: `填充失败(${selector})：${String(e?.message || e)}` };
    }
  }
  return { ok: true, message: "done" };
}

function ensurePickLabel() {
  if (pickLabel) return pickLabel;
  pickLabel = document.createElement("div");
  pickLabel.style.position = "fixed";
  pickLabel.style.zIndex = "2147483647";
  pickLabel.style.pointerEvents = "none";
  pickLabel.style.padding = "4px 8px";
  pickLabel.style.borderRadius = "8px";
  pickLabel.style.fontSize = "12px";
  pickLabel.style.fontFamily = "Segoe UI, Arial, sans-serif";
  pickLabel.style.background = "rgba(12,18,40,.85)";
  pickLabel.style.color = "#e9eefc";
  pickLabel.style.boxShadow = "0 6px 16px rgba(0,0,0,.25)";
  pickLabel.style.transform = "translate(12px, 12px)";
  document.documentElement.appendChild(pickLabel);
  return pickLabel;
}

function updatePickLabel(e) {
  if (!pickSession) return;
  const label = ensurePickLabel();
  const name = pickSession.colName || "";
  label.textContent = name || "";
  label.style.left = `${e.clientX}px`;
  label.style.top = `${e.clientY}px`;
}

function clearPickLabel() {
  if (!pickLabel) return;
  pickLabel.remove();
  pickLabel = null;
}

function onPickClick(e) {
  if (!pickSession) return;
  let target = e.target;
  if (pickSession.mode !== "identity") {
    for (let i = 0; i < 4 && target && !isFillable(target); i++) {
      target = target.parentElement;
    }
    if (!target || !isFillable(target)) {
      const choice = findChoiceInputFromTarget(e.target);
      if (!choice) return;
      target = choice;
    }
  }
  if (!target) return;
  e.preventDefault();
  e.stopPropagation();

  let pickTarget = target;
  if (isChoiceInput(target) && pickSession.mode !== "identity") {
    const root = findChoiceGroupRoot(target);
    if (root) pickTarget = root;
  }
  const selector = buildSelector(pickTarget) || buildSelector(target);
  if (!selector) return;
  const meta = {
    tag: target.tagName?.toLowerCase(),
    type: (target.getAttribute?.("type") || "").toLowerCase(),
    name: target.getAttribute?.("name") || "",
    id: target.getAttribute?.("id") || "",
    placeholder: target.getAttribute?.("placeholder") || "",
    label: findLabelText(target),
    ariaLabel: target.getAttribute?.("aria-label") || "",
    text: (target.textContent || "").trim(),
  };
  if (isChoiceInput(target)) {
    meta.groupType = (target.getAttribute?.("type") || "").toLowerCase();
    meta.groupName = target.getAttribute?.("name") || "";
  }

  highlight(pickTarget);
  chrome.runtime.sendMessage({
    type: "FB_PICK_RESULT",
    payload: {
      mode: pickSession.mode,
      colIndex: pickSession.colIndex,
      colName: pickSession.colName,
      selector,
      meta,
    },
  });

  endPick();
}

function onPickMove(e) {
  if (!pickSession) return;
  updatePickLabel(e);
  let target = e.target;
  if (pickSession.mode !== "identity") {
    for (let i = 0; i < 4 && target && !isFillable(target); i++) {
      target = target.parentElement;
    }
    if (!target || !isFillable(target)) {
      const choice = findChoiceInputFromTarget(e.target);
      if (!choice) return;
      target = choice;
    }
  }
  if (!target) return;
  let hoverTarget = target;
  if (isChoiceInput(target) && pickSession.mode !== "identity") {
    const root = findChoiceGroupRoot(target);
    if (root) hoverTarget = root;
  }
  lastHoverTarget = hoverTarget;
  if (hoverRAF) return;
  hoverRAF = window.requestAnimationFrame(() => {
    hoverRAF = null;
    if (lastHoverTarget) updateHoverBox(lastHoverTarget);
  });
}

function endPick() {
  pickSession = null;
  if (hoverRAF) {
    window.cancelAnimationFrame(hoverRAF);
    hoverRAF = null;
  }
  lastHoverTarget = null;
  window.removeEventListener("click", onPickClick, true);
  window.removeEventListener("mousemove", onPickMove, true);
  window.removeEventListener("keydown", onPickEsc, true);
  document.documentElement.style.cursor = "";
  clearHoverBox();
  clearPickLabel();
}

function onPickEsc(e) {
  if (e.key === "Escape") {
    chrome.runtime.sendMessage({ type: "FB_PICK_CANCELLED" });
    endPick();
  }
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg?.type === "FB_BEGIN_PICK") {
    pickSession = { mode: msg.mode, colIndex: msg.colIndex, colName: msg.colName };
    window.addEventListener("click", onPickClick, true);
    window.addEventListener("mousemove", onPickMove, true);
    window.addEventListener("keydown", onPickEsc, true);
    document.documentElement.style.cursor = "crosshair";
    chrome.runtime.sendMessage({
      type: "FB_RUN_EVENT",
      payload: { line: "进入点选模式：请点击输入框/下拉框/文本域" },
    });
    sendResponse({ ok: true });
    return true;
  }

  if (msg?.type === "FB_CANCEL_PICK") {
    endPick();
    sendResponse({ ok: true });
    return true;
  }

  if (msg?.type === "FB_READ_TEXT") {
    const selector = msg.payload?.selector;
    try {
      const el = document.querySelector(selector);
      if (!el) {
        sendResponse({ ok: false, text: "" });
        return true;
      }
      const tag = el.tagName?.toLowerCase();
      let text = "";
      if (tag === "input" || tag === "textarea") text = el.value || "";
      else if (tag === "select") {
        const opt = el.selectedOptions?.[0];
        text = opt?.textContent || el.value || "";
      } else text = el.textContent || "";
      sendResponse({ ok: true, text: text.trim() });
      return true;
    } catch (e) {
      sendResponse({ ok: false, text: "" });
      return true;
    }
  }

  if (msg?.type === "FB_FILL_ONE") {
    const res = fillOne(msg.payload);
    sendResponse(res);
    return true;
  }
});
