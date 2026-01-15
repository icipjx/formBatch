/**
 * FormBatcher MVP Content Script
 * - Pick mode: user clicks an element, we compute a stable selector and return it.
 * - Fill mode: receive mapping + row values, fill DOM elements, trigger events.
 */

let pickSession = null; // {mode, colIndex, colName}

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
      const truthy = ["1", "true", "yes", "y", "on"].includes(String(value).trim().toLowerCase());
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
    highlight(el);
    try {
      setValue(el, val);
    } catch (e) {
      return { ok: false, error: `填充失败(${selector})：${String(e?.message || e)}` };
    }
  }
  return { ok: true, message: "done" };
}

function onPickClick(e) {
  if (!pickSession) return;
  let target = e.target;
  for (let i = 0; i < 4 && target && !isFillable(target); i++) {
    target = target.parentElement;
  }
  if (!target || !isFillable(target)) return;
  e.preventDefault();
  e.stopPropagation();

  const selector = buildSelector(target);
  const meta = {
    tag: target.tagName?.toLowerCase(),
    type: (target.getAttribute?.("type") || "").toLowerCase(),
    name: target.getAttribute?.("name") || "",
    id: target.getAttribute?.("id") || "",
    placeholder: target.getAttribute?.("placeholder") || "",
    label: findLabelText(target),
    ariaLabel: target.getAttribute?.("aria-label") || "",
  };

  highlight(target);
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

function endPick() {
  pickSession = null;
  window.removeEventListener("click", onPickClick, true);
  window.removeEventListener("keydown", onPickEsc, true);
  document.documentElement.style.cursor = "";
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
    window.addEventListener("keydown", onPickEsc, true);
    document.documentElement.style.cursor = "crosshair";
    chrome.runtime.sendMessage({
      type: "FB_RUN_EVENT",
      payload: { line: "进入点选模式：请点击输入框/下拉框/文本域" },
    });
    sendResponse({ ok: true });
    return true;
  }

  if (msg?.type === "FB_FILL_ONE") {
    const res = fillOne(msg.payload);
    sendResponse(res);
    return true;
  }
});
