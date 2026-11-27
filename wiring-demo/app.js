"use strict";

const $ = (id) => document.getElementById(id);
const form = $("review");
const fields = form.elements;
const workbookDrawer = $("workbook-drawer");
const workbookToggle = $("workbook-toggle");
const employeeInput = $("employee-id");
const statuses = { pending: "待核對", complete: "已確認", revise: "需修改" };
const blankReview = { lengthMode: "", adjustment: "", content: "pending", revise: false, note: "" };
const contentResults = {
  pending: { label: "接線待核對", tone: "pending" },
  correct: { label: "接線正確", tone: "complete" },
  wrong: { label: "接線有誤", tone: "revise" },
};
const lengthModes = { "": "待核對", same: "不需調整", increase: "需增加", decrease: "需縮短" };
const auditLabels = {
  lengthMode: "線長調整", adjustment: "調整量（mm）", content: "接線內容",
  revise: "其他修改事項", note: "現場備註",
};
const taipeiTime = new Intl.DateTimeFormat("sv-SE", { dateStyle: "short", timeStyle: "medium", timeZone: "Asia/Taipei" });
const storagePrefix = "wiring-review-v1:";
const trackingPrefix = "wiring-tracking-outbox:";
const trackingEnabled = location.hostname === "localhost" && location.protocol === "http:";
const number = new Intl.NumberFormat("zh-TW", { maximumFractionDigits: 6 });
const normalized = (value) => String(value ?? "").normalize("NFKC").replace(/\s+/g, "").toLowerCase();

let project;
let sheet;
let selected;
let reviews = new Map();
const unsavedReviews = new Map();
let workbooks = [];
let activeWorkbook = -1;
let operator = "";
let ongoingEdit = null;
let pendingReviewLog = null;
let trackingInFlight = false;
let trackingRetry;
const trackingOutbox = new Map();

function node(tag, text, className) {
  const element = document.createElement(tag);
  element.textContent = text;
  if (className) element.className = className;
  return element;
}

function readStorage(key) {
  try {
    return JSON.parse(localStorage.getItem(storagePrefix + key));
  } catch {
    return null;
  }
}

function writeStorage(key, value) {
  try {
    localStorage.setItem(storagePrefix + key, JSON.stringify(value));
    return true;
  } catch {
    return false;
  }
}

function restoreTrackingOutbox() {
  try {
    for (const key of Object.keys(localStorage)) {
      if (!key.startsWith(trackingPrefix)) continue;
      const event = JSON.parse(localStorage.getItem(key));
      if (event?.id) trackingOutbox.set(event.id, event);
    }
  } catch {
    // Tracking remains available for this page even when browser storage is unavailable.
  }
}

async function flushTracking() {
  if (trackingInFlight || !trackingOutbox.size) return;
  trackingInFlight = true;
  clearTimeout(trackingRetry);
  try {
    for (const [id, event] of trackingOutbox) {
      const response = await fetch("/_tracking", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify(event), keepalive: true,
      });
      if (!response.ok) break;
      trackingOutbox.delete(id);
      try { localStorage.removeItem(trackingPrefix + id); } catch {}
    }
  } catch {
    // Retry queued actions if the local server is temporarily stopped.
  } finally {
    trackingInFlight = false;
    if (trackingOutbox.size) trackingRetry = setTimeout(flushTracking, 5000);
  }
}

function track(action, details = {}, employeeId = operator, time = new Date().toISOString()) {
  if (!trackingEnabled || !employeeId) return;
  const event = { id: crypto.randomUUID(), time, action, employeeId, details };
  trackingOutbox.set(event.id, event);
  try { localStorage.setItem(trackingPrefix + event.id, JSON.stringify(event)); } catch {}
  flushTracking();
}

function commitReviewLog() {
  if (!pendingReviewLog) return;
  const { entry, details } = pendingReviewLog;
  pendingReviewLog = null;
  track("review_changed", details, entry.employeeId, entry.at);
}

function savedReviewEntries() {
  return Object.keys(localStorage)
    .filter((key) => key.startsWith(storagePrefix))
    .map((key) => [key, localStorage.getItem(key)]);
}

function downloadReviewBackup() {
  const entries = savedReviewEntries();
  if (!entries.length) return "此檔案頁面沒有已儲存的線表紀錄。";
  const backup = JSON.stringify({ format: "wiring-review-backup-v1", entries });
  const url = URL.createObjectURL(new Blob([backup], { type: "application/json" }));
  const link = node("a", "");
  link.href = url;
  link.download = "wiring-review-backup.json";
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  return "備份已下載。請在 localhost 登入畫面選擇此備份。";
}

function restoreReviewBackup(backup) {
  if (backup?.format !== "wiring-review-backup-v1" || !Array.isArray(backup.entries) || !backup.entries.length) {
    throw new Error("請選擇從舊版備線核對台下載的備份。");
  }
  if (savedReviewEntries().length) {
    throw new Error("localhost 已有儲存資料，無法直接匯入這份備份。");
  }
  const keys = new Set();
  for (const entry of backup.entries) {
    const [key, value] = Array.isArray(entry) ? entry : [];
    if (typeof key !== "string" || !key.startsWith(storagePrefix) || typeof value !== "string" || keys.has(key)) {
      throw new Error("備份格式有誤。");
    }
    try {
      if (!JSON.parse(value)) throw new Error();
    } catch {
      throw new Error("備份格式有誤。");
    }
    keys.add(key);
  }
  const restored = [];
  try {
    for (const [key, value] of backup.entries) {
      localStorage.setItem(key, value);
      restored.push(key);
    }
  } catch {
    for (const key of restored) localStorage.removeItem(key);
    throw new Error("瀏覽器無法儲存備份，請確認剩餘儲存空間。");
  }
  try {
    initializeDashboard();
  } catch {
    for (const key of restored) localStorage.removeItem(key);
    initializeDashboard();
    throw new Error("備份中的線表資料無法讀取。");
  }
}

function readProject(name, bytes) {
  const workbook = XLSX.read(bytes, { cellHTML: false });
  const sheets = workbook.SheetNames.map((sheetName) => {
    const source = workbook.Sheets[sheetName];
    if (!source["!ref"]) return null;
    const rows = XLSX.utils.sheet_to_json(source, { header: 1, range: 0, raw: false, defval: "" });
    const headerRow = rows.findIndex((row) => {
      const labels = row.map(normalized);
      return labels.some((label) => /^(no\.?|編號|序號)$/.test(label))
        && labels.includes("線名") && labels.includes("前接頭")
        && labels.some((label) => /^(長度|線長|設計長度|線長\(.+\)|長度\(.+\))$/.test(label));
    });
    if (headerRow < 0) return null;

    const headers = rows[headerRow].map((label) => label.trim().replace(/\s+/g, ""));
    const labels = headers.map(normalized);
    const serial = labels.findIndex((label) => /^(no\.?|編號|序號)$/.test(label));
    const wireName = labels.indexOf("線名");
    const front = labels.indexOf("前接頭");
    const length = labels.findIndex((label) => /^(長度|線長|設計長度|線長\(.+\)|長度\(.+\))$/.test(label));
    const back = labels.indexOf("後接頭");
    const remark = labels.findIndex((label) => /^(備註|備考)$/.test(label));
    const specification = labels.findIndex((label) => /線材|^規格$/.test(label));
    const notes = [];
    const wires = [];
    let current = null;
    let inNotes = false;
    rows.slice(headerRow + 1).forEach((row, offset) => {
      const cells = headers.map((_, column) => row[column]?.trim() || "");
      const populated = row.map((value) => value.trim()).filter(Boolean);
      if (!populated.length) return;
      if (/^備註\s*[:：]/.test(populated[0])) inNotes = true;
      if (inNotes) {
        notes.push(populated.join(" ").replace(/^備註\s*[:：]\s*/, ""));
      } else if (cells[serial] && (cells[wireName] || cells[length])) {
        current = { row: headerRow + offset + 2, cells };
        wires.push(current);
      } else if (!cells[serial] && current) {
        cells.forEach((value, column) => {
          if (value) current.cells[column] = [current.cells[column], value].filter(Boolean).join("\n");
        });
      }
    });
    if (!wires.length) return null;
    const metadata = rows.slice(0, headerRow).flat().map((value) => value.trim());
    const meta = (label) => metadata.find((value) => value.startsWith(label))?.replace(/^[^:：]+[:：]\s*/, "") || "";
    return {
      name: sheetName, headers, serial, wireName, front, length, back, remark, specification,
      model: meta("機型"), part: meta("料號"), revision: meta("版次"),
      wires, notes: notes.filter(Boolean),
    };
  }).filter(Boolean);
  if (!sheets.length) throw new Error("找不到備線清單。請使用含 NO、線名、前接頭及長度欄位的工作表。");
  return { name, sheets };
}

function reviewKey(wire, worksheet = sheet) {
  const key = [project.name, worksheet.name, worksheet.model, worksheet.part, worksheet.revision, worksheet.headers, wire.row, wire.cells];
  const version = workbooks[activeWorkbook].version;
  if (version > 1) key.push(version);
  return JSON.stringify(key);
}

function reviewFor(wire) {
  return reviews.get(wire.row);
}

function designLength(wire, worksheet = sheet) {
  const value = normalized(wire.cells[worksheet.length]);
  const match = value.match(/^(\d+(?:\.\d+)?|\.\d+)(mm|cm|m|毫米|公厘|公分|公尺)?$/);
  if (!match) return null;
  const unit = match[2] || normalized(worksheet.headers[worksheet.length]).match(/\(([^)]+)\)$/)?.[1];
  const multipliers = { mm: 1, 毫米: 1, 公厘: 1, cm: 10, 公分: 10, m: 1000, 公尺: 1000 };
  const converted = Number(match[1]) * multipliers[unit];
  return Number.isFinite(converted) ? converted : null;
}

function legacyLengthStatus(wire, review, worksheet) {
  const reference = designLength(wire, worksheet);
  if (reference === null || !review.actual || review.tolerance === "") return "pending";
  const actual = Number(review.actual);
  const tolerance = Number(review.tolerance ?? 0);
  const roundingError = 2 * Number.EPSILON * Math.max(actual, reference, tolerance);
  return Math.abs(actual - reference) <= tolerance + roundingError ? "complete" : "revise";
}

function loadReview(wire, worksheet = sheet) {
  const key = reviewKey(wire, worksheet);
  const stored = unsavedReviews.get(key) ?? readStorage(key);
  if (!stored || Object.hasOwn(stored, "lengthMode")) return { ...blankReview, ...stored };
  const previousStatus = legacyLengthStatus(wire, stored, worksheet);
  return {
    ...blankReview,
    ...stored,
    lengthMode: previousStatus === "complete" ? "same" : "",
    priorLengthIssue: previousStatus === "revise",
  };
}

function needsAdjustment(mode) {
  return mode === "increase" || mode === "decrease";
}

function adjustedLength(reference, mode, amount) {
  if (reference === null || !mode) return null;
  if (mode === "same") return reference;
  if (!amount) return null;
  return mode === "increase" ? reference + Number(amount) : reference - Number(amount);
}

function statusOf(wire, review = reviewFor(wire)) {
  if (review.revise || review.content === "wrong" || needsAdjustment(review.lengthMode) || review.priorLengthIssue) return "revise";
  if (review.content === "correct" && review.lengthMode === "same") return "complete";
  return "pending";
}

function badge(status) {
  return node("span", statuses[status], "badge " + status);
}

function drawSummary() {
  const counts = { pending: 0, complete: 0, revise: 0 };
  sheet.wires.forEach((wire) => counts[statusOf(wire)]++);
  $("count-total").textContent = sheet.wires.length;
  Object.keys(counts).forEach((status) => { $("count-" + status).textContent = counts[status]; });
  $("progress").max = sheet.wires.length;
  $("progress").value = counts.complete;
  $("progress-label").textContent = `已確認 ${counts.complete} / ${sheet.wires.length} 筆`;
}

function listColumns() {
  return [...new Set([sheet.serial, sheet.wireName, sheet.front, sheet.specification, sheet.length])].filter((column) => column >= 0).sort((a, b) => a - b);
}

function isSecondaryColumn(column) {
  return column === sheet.front || column === sheet.specification;
}

function drawTable() {
  if (!sheet) return;
  const search = normalized($("search").value);
  const filter = $("filter").value;
  const columns = listColumns();
  const matches = sheet.wires.filter((wire) => {
    return (filter === "all" || statusOf(wire) === filter) && wire.cells.some((cell) => normalized(cell).includes(search));
  });
  const body = $("table-body");
  body.replaceChildren();
  for (const wire of matches) {
    const row = body.insertRow();
    row.classList.toggle("selected", wire === selected);
    for (const column of columns) {
      const cell = row.insertCell();
      cell.classList.toggle("secondary-column", isSecondaryColumn(column));
      if (column === sheet.wireName) {
        const button = node("button", wire.cells[column] || "查看項目", "wire-button");
        button.type = "button";
        button.setAttribute("aria-pressed", String(wire === selected));
        button.addEventListener("click", () => {
          drawDetail(wire);
          drawTable();
          $("detail-title").focus();
        });
        cell.append(button);
      } else {
        const value = column === sheet.front ? wire.cells[column].split("\n")[0] : wire.cells[column];
        cell.textContent = value || "—";
        if (column === sheet.serial || column === sheet.length) cell.className = "code";
      }
    }
    const results = node("div", "", "row-results");
    const content = contentResults[reviewFor(wire).content];
    results.append(badge(statusOf(wire)), node("span", content.label, "row-content " + content.tone));
    row.insertCell().append(results);
  }
  if (!matches.length) {
    const cell = body.insertRow().insertCell();
    cell.colSpan = columns.length + 1;
    cell.className = "empty";
    cell.textContent = "沒有符合條件的備線項目。";
  }
  $("row-count").textContent = `顯示 ${matches.length} / ${sheet.wires.length} 筆`;
}

function updateDetail() {
  const review = reviewFor(selected);
  const last = review.history?.at(-1);
  $("review-operator").textContent = last ? `最近核對工號 ${last.employeeId}` : "";
  const reference = designLength(selected);
  const result = adjustedLength(reference, review.lengthMode, review.adjustment);
  $("detail-status").replaceChildren(badge(statusOf(selected)));
  const content = contentResults[review.content];
  $("content-result").replaceChildren(node("span", content.label, "badge " + content.tone));
  $("adjusted-length").textContent = result === null ? "—" : `${number.format(result)} mm`;
  const feedback = $("length-feedback");
  feedback.className = "length-feedback";
  if (fields.adjustment.getAttribute("aria-invalid")) {
    feedback.classList.add("revise");
    feedback.textContent = fields.adjustment.validationMessage;
  } else if (needsAdjustment(review.lengthMode)) {
    feedback.classList.add("revise");
    const direction = review.lengthMode === "increase" ? "增加" : "縮短";
    if (!review.adjustment) feedback.textContent = `請輸入要${direction}的毫米數。`;
    else if (reference === null) feedback.textContent = `需${direction} ${number.format(Number(review.adjustment))} mm。原表線長缺少可辨識的數值或單位。`;
    else feedback.textContent = `需${direction} ${number.format(Number(review.adjustment))} mm。`;
  } else if (review.lengthMode === "same") {
    feedback.classList.add("complete");
    feedback.textContent = "線長已核對，無需調整。";
  } else if (review.actual) {
    feedback.textContent = `上次實測 ${number.format(Number(review.actual))} mm。請確認是否需調整。`;
  } else {
    feedback.textContent = "請確認線長是否需要調整。";
  }
}

function drawWireDiagram(wire) {
  const labels = sheet.headers.map(normalized);
  const stripLength = /^剝線長(?:\(.+\))?$/;
  const frontStrip = labels.findIndex((label, i) => i > sheet.front && i < sheet.length && stripLength.test(label));
  const backName = labels.findIndex((label, i) => i > sheet.length && label === "線名");
  const backStrip = labels.findIndex((label, i) => i > sheet.length && stripLength.test(label));
  const values = {
    "wire-front-name": sheet.wireName,
    "wire-front-connector": sheet.front,
    "wire-front-strip": frontStrip,
    "wire-back-name": backName,
    "wire-back-connector": sheet.back,
    "wire-back-strip": backStrip,
    "wire-specification": sheet.specification,
    "wire-design-length": sheet.length,
  };
  for (const [id, column] of Object.entries(values)) {
    $(id).textContent = wire.cells[column] || "—";
  }
  const middle = $("wire-middle");
  middle.replaceChildren();
  const rightStart = Math.min(...[backName, backStrip, sheet.back].filter((column) => column > sheet.length));
  if (Number.isFinite(rightStart)) {
    for (let column = sheet.length + 1; column < rightStart; column++) {
      const value = wire.cells[column];
      if (!value) continue;
      const field = node("div", "");
      const letter = XLSX.utils.encode_col(column);
      const label = sheet.headers[column];
      field.append(node("dt", label ? `${label}（${letter}欄）` : `${letter} 欄`), node("dd", value));
      middle.append(field);
    }
  }
  middle.hidden = !middle.childElementCount;
  const remark = wire.cells[sheet.remark] || "";
  $("wire-remark-text").textContent = remark;
  $("wire-remarks").hidden = !remark;
}

function drawDetail(wire) {
  commitReviewLog();
  selected = wire;
  ongoingEdit = null;
  $("detail-title").tabIndex = -1;
  $("detail-name").textContent = `原表第 ${wire.row} 列 · ${wire.cells[sheet.wireName]}`;
  form.hidden = false;
  $("empty-detail").hidden = true;
  drawWireDiagram(wire);
  const all = $("source-all");
  all.replaceChildren();
  sheet.headers.forEach((label, column) => {
    if (!label && !wire.cells[column]) return;
    const repeated = sheet.headers.filter((heading) => heading === label).length > 1;
    const heading = repeated ? `${label}（${XLSX.utils.encode_col(column)}欄）` : label;
    all.append(node("dt", heading || `${XLSX.utils.encode_col(column)} 欄`), node("dd", wire.cells[column] || "—"));
  });
  const reference = designLength(wire);
  const original = wire.cells[sheet.length] || "—";
  $("reference").textContent = reference === null ? original : `${number.format(reference)} mm`;
  $("reference").title = `原表 ${original}`;

  const review = reviewFor(wire);
  fields.lengthMode.value = review.lengthMode;
  fields.adjustment.value = review.adjustment;
  fields.content.value = review.content;
  fields.revise.checked = review.revise;
  fields.note.value = review.note;
  syncValidation();
  $("save-status").textContent = "核對紀錄儲存在此瀏覽器";
  updateDetail();
}

function syncValidation() {
  const mode = fields.lengthMode.value;
  const enabled = needsAdjustment(mode);
  const input = fields.adjustment;
  const reference = designLength(selected);
  $("adjustment-field").hidden = !enabled;
  input.disabled = !enabled;
  input.setCustomValidity("");
  if (enabled && (input.validity.badInput || (input.value !== "" && !(input.valueAsNumber > 0)))) {
    input.setCustomValidity("請輸入大於 0 的毫米數。");
  } else if (enabled && input.value && reference !== null) {
    const result = adjustedLength(reference, mode, input.value);
    if (result <= 0) input.setCustomValidity("縮短量必須小於設計線長。");
    else if (!Number.isFinite(result)) input.setCustomValidity("請輸入有效的調整量。");
  }
  if (input.validity.valid) input.removeAttribute("aria-invalid");
  else input.setAttribute("aria-invalid", "true");
}

function openSheet(index) {
  sheet = project.sheets[index];
  $("model").textContent = sheet.model || sheet.name;
  $("part").textContent = sheet.part ? "料號 " + sheet.part : "";
  $("revision").textContent = sheet.revision ? "版次 " + sheet.revision : "";
  $("sheet-name").textContent = `工作表 ${sheet.name}`;
  $("sheet").value = String(index);
  $("search").value = "";
  $("filter").value = "all";
  reviews = new Map(sheet.wires.map((wire) => [wire.row, loadReview(wire)]));
  const header = node("tr", "");
  $("table-head").replaceChildren(header);
  for (const column of [...listColumns(), -1]) {
    const cell = node("th", sheet.headers[column] || "核對狀態");
    if (column >= 0 && isSecondaryColumn(column)) cell.classList.add("secondary-column");
    cell.scope = "col";
    header.append(cell);
  }
  $("source-notes").hidden = !sheet.notes.length;
  $("notes-list").replaceChildren(...sheet.notes.map((note) => node("li", note)));
  drawDetail(sheet.wires[0]);
  drawSummary();
  drawTable();
}

function openProject(data, index = 0) {
  project = data;
  $("file-name").textContent = data.name;
  $("sheet").replaceChildren(...data.sheets.map((item, i) => {
    const option = node("option", item.name);
    option.value = i;
    return option;
  }));
  $("sheet-control").hidden = data.sheets.length < 2;
  $("export-report").disabled = false;
  openSheet(index);
}

function addWorkbook(data, sheetIndex = 0) {
  const matches = workbooks.filter((item) => item.data.name === data.name);
  if (matches.length) {
    const sheets = JSON.stringify(data.sheets);
    const existing = matches.find((item) => JSON.stringify(item.data.sheets) === sheets);
    if (existing) return workbooks.indexOf(existing);
  }
  workbooks.push({ data, sheetIndex, version: matches.length + 1 });
  return workbooks.length - 1;
}

function saveWorkbooks() {
  return writeStorage("workbooks", { workbooks, activeWorkbook });
}

function selectWorkbook(index) {
  activeWorkbook = index;
  const workbook = workbooks[index];
  openProject(workbook.data, workbook.sheetIndex);
  [...$("workbook-list").querySelectorAll("button")].forEach((button, i) => {
    button.setAttribute("aria-pressed", String(i === index));
  });
}

function drawWorkbookList() {
  $("workbook-count").textContent = `${workbooks.length} 個`;
  $("workbook-empty").hidden = workbooks.length > 0;
  $("workbook-list").replaceChildren(...workbooks.map(({ data, version }, index) => {
    const item = node("li", "");
    const button = node("button", "");
    button.type = "button";
    const wires = data.sheets.reduce((count, item) => count + item.wires.length, 0);
    const meta = version > 1 ? `版本 ${version} · ${wires} 筆線材` : `${wires} 筆線材`;
    button.append(node("strong", data.name), node("span", meta));
    button.addEventListener("click", () => {
      if (index !== activeWorkbook) {
        selectWorkbook(index);
        const saved = saveWorkbooks();
        message(saved ? "" : "已切換檔案。瀏覽器無法保留這次切換。", !saved);
      }
      workbookDrawer.close();
    });
    item.append(button);
    return item;
  }));
}

async function readWorkbookFile(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const compound = bytes[0] === 0xd0 && bytes[1] === 0xcf;
  const biff = bytes[0] === 0x09 && [0, 2, 4, 8].includes(bytes[1]);
  const zip = bytes[0] === 0x50 && bytes[1] === 0x4b;
  if (!/\.xlsx?$/i.test(file.name) || !(compound || biff || zip)) throw new Error("請選擇有效的 .xls 或 .xlsx 工作簿。");
  return readProject(file.name, bytes);
}

function message(text, isError = false) {
  const target = $("message");
  target.textContent = text;
  target.hidden = !text;
  target.classList.toggle("error", isError);
}

function recordReview(previous, next, target) {
  let history = previous.history ?? [];
  const changes = Object.keys(auditLabels)
    .filter((field) => previous[field] !== next[field])
    .map((field) => ({ field, before: previous[field], after: next[field] }));
  if (!changes.length) return;

  // Keep one change per focused text or number edit, while still saving each input.
  const draftField = target === fields.note ? "note" : target === fields.adjustment ? "adjustment" : null;
  const singleDraft = draftField && changes.length === 1 && changes[0].field === draftField;
  const last = history.at(-1);
  if (singleDraft && ongoingEdit === target && last?.employeeId === operator
    && last.changes.length === 1 && last.changes[0].field === draftField) {
    history = history.slice(0, -1);
    changes[0].before = last.changes[0].before;
    if (changes[0].before === changes[0].after) {
      ongoingEdit = null;
      next.history = history;
      return null;
    }
  }
  ongoingEdit = singleDraft ? target : null;
  const entry = { employeeId: operator, at: new Date().toISOString(), changes };
  next.history = [...history, entry];
  return entry;
}

function reportValue(field, value) {
  if (field === "lengthMode") return lengthModes[value];
  if (field === "content") return contentResults[value].label;
  if (field === "revise") return value ? "是" : "否";
  if (field === "adjustment" && value !== "") return Number(value);
  return value;
}

function trackReviewEdit(entry, previous, review, target) {
  if (!trackingEnabled || entry === undefined) return;
  const isDraft = target === fields.note || target === fields.adjustment;
  if (!entry) {
    if (isDraft) pendingReviewLog = null;
    return;
  }
  const valueForLog = (field, value) => {
    const displayValue = reportValue(field, value);
    return typeof displayValue === "string" && displayValue.length > 1000
      ? displayValue.slice(0, 1000) + "…" : displayValue;
  };
  const details = {
    workbook: project.name, version: workbooks[activeWorkbook].version, worksheet: sheet.name,
    row: selected.row, wire: selected.cells[sheet.serial]?.slice(0, 250),
    wireName: selected.cells[sheet.wireName]?.slice(0, 250),
    statusBefore: pendingReviewLog?.details.statusBefore ?? statuses[statusOf(selected, previous)],
    statusAfter: statuses[statusOf(selected, review)],
    changes: entry.changes.map(({ field, before, after }) => ({
      field: auditLabels[field], before: valueForLog(field, before), after: valueForLog(field, after),
    })),
  };
  pendingReviewLog = { target, entry, details };
  if (!isDraft) commitReviewLog();
}

function exportReviewReport() {
  if (!operator || !project) return;
  const date = taipeiTime.format(new Date());
  const version = workbooks[activeWorkbook].version;
  const summary = [
    ["備線核對報告"],
    ["來源線表", project.name],
    ["線表版本", version],
    ["產出者工號", operator],
    ["產出時間（台灣）", date],
    ["備線工作表數", project.sheets.length],
  ];
  const results = [[
    "工作表", "原表列", "線號", "線名", "前接頭", "後接頭", "線材規格", "原設計線長",
    "線長調整", "調整量（mm）", "調整後線長（mm）", "接線核對", "其他修改事項",
    "整體狀態", "現場備註", "最近操作工號", "最近操作時間（台灣）",
  ]];
  const historyRows = [["工作表", "原表列", "線號", "線名", "操作工號", "操作時間（台灣）", "欄位", "修改前", "修改後"]];
  for (const worksheet of project.sheets) {
    for (const wire of worksheet.wires) {
      const review = loadReview(wire, worksheet);
      const history = review.history ?? [];
      const latest = history.at(-1);
      const adjusted = adjustedLength(designLength(wire, worksheet), review.lengthMode, review.adjustment);
      const identity = [worksheet.name, wire.row, wire.cells[worksheet.serial], wire.cells[worksheet.wireName]];
      results.push([
        ...identity, wire.cells[worksheet.front], wire.cells[worksheet.back], wire.cells[worksheet.specification],
        wire.cells[worksheet.length], lengthModes[review.lengthMode], reportValue("adjustment", review.adjustment),
        adjusted ?? "", contentResults[review.content].label, reportValue("revise", review.revise),
        statuses[statusOf(wire, review)], review.note, latest?.employeeId ?? "",
        latest ? taipeiTime.format(new Date(latest.at)) : "",
      ]);
      for (const { employeeId, at, changes } of history) {
        for (const { field, before, after } of changes) {
          historyRows.push([...identity, employeeId, taipeiTime.format(new Date(at)), auditLabels[field],
            reportValue(field, before), reportValue(field, after)]);
        }
      }
    }
  }
  const report = XLSX.utils.book_new();
  for (const [name, data, widths] of [
    ["報告資訊", summary, [24, 45]],
    ["核對結果", results, [18, 10, 10, 20, 25, 20, 30, 16, 16, 16, 20, 16, 16, 14, 36, 18, 26]],
    ["修改紀錄", historyRows, [18, 10, 10, 20, 18, 26, 18, 36, 36]],
  ]) {
    const reportSheet = XLSX.utils.aoa_to_sheet(data);
    reportSheet["!cols"] = widths.map((wch) => ({ wch }));
    XLSX.utils.book_append_sheet(report, reportSheet, name);
  }
  const source = project.name.replace(/\.xlsx?$/i, "");
  const versionLabel = version > 1 ? `_版本${version}` : "";
  const filename = `${source}${versionLabel}_核對報告_${date.slice(0, 10)}.xlsx`;
  XLSX.writeFile(report, filename);
  track("report_export", { workbook: project.name, version, filename });
}

const filePage = location.protocol === "file:";
if (filePage || location.hostname === "localhost") {
  $("migration-tools").hidden = false;
  $("migration-hint").textContent = filePage
    ? "下載此檔案頁面已儲存的線表與核對紀錄。"
    : "首次使用 localhost 時，可匯入從舊檔案頁面下載的紀錄。";
  $("migration-action").textContent = filePage ? "下載舊版紀錄" : "選擇備份";
}
function migrationFeedback(text) {
  $("migration-status").textContent = text;
  $("migration-status").hidden = false;
}
$("migration-action").addEventListener("click", () => {
  if (!filePage) {
    $("migration-file").click();
    return;
  }
  try {
    migrationFeedback(downloadReviewBackup());
  } catch {
    migrationFeedback("無法讀取此瀏覽器的舊版紀錄。");
  }
});
$("migration-file").addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file || filePage || operator) return;
  event.target.value = "";
  let backup;
  try {
    backup = JSON.parse(await file.text());
  } catch {
    migrationFeedback("無法讀取備份，請確認是有效的 JSON 檔案。");
    return;
  }
  if (operator) return;
  try {
    restoreReviewBackup(backup);
    migrationFeedback("已移轉線表和核對紀錄，請輸入工號進入。");
  } catch (error) {
    migrationFeedback(error.message);
  }
});

$("operator-form").addEventListener("submit", (event) => {
  event.preventDefault();
  const employeeId = employeeInput.value.trim();
  if (!employeeId) {
    employeeInput.setCustomValidity("請輸入工號。");
    employeeInput.reportValidity();
    return;
  }
  operator = employeeId;
  track("login");
  $("current-operator").textContent = `工號 ${operator}`;
  $("operator-gate").hidden = true;
  $("dashboard-header").hidden = false;
  $("dashboard").hidden = false;
  workbookToggle.focus();
});
employeeInput.addEventListener("input", () => employeeInput.setCustomValidity(""));
$("switch-operator").addEventListener("click", () => {
  commitReviewLog();
  track("operator_switch");
  operator = "";
  ongoingEdit = null;
  employeeInput.value = "";
  employeeInput.setCustomValidity("");
  $("dashboard").hidden = true;
  $("dashboard-header").hidden = true;
  $("operator-gate").hidden = false;
  employeeInput.focus();
});
$("export-report").addEventListener("click", () => {
  try {
    exportReviewReport();
  } catch {
    message("無法匯出核對報告，請重試。", true);
  }
});

form.addEventListener("submit", (event) => event.preventDefault());
form.addEventListener("focusout", () => {
  commitReviewLog();
  ongoingEdit = null;
});
form.addEventListener("input", (event) => {
  if (!operator) return;
  if (pendingReviewLog && pendingReviewLog.target !== event.target) commitReviewLog();
  const lengthMode = fields.lengthMode.value;
  const previous = reviewFor(selected);
  if (event.target === fields.lengthMode && !needsAdjustment(lengthMode)) fields.adjustment.value = "";
  syncValidation();
  const review = {
    ...previous,
    lengthMode,
    adjustment: needsAdjustment(lengthMode) && fields.adjustment.validity.valid ? fields.adjustment.value : "",
    priorLengthIssue: event.target === fields.lengthMode ? false : previous.priorLengthIssue,
    content: fields.content.value,
    revise: fields.revise.checked,
    note: fields.note.value,
  };
  const entry = recordReview(previous, review, event.target);
  reviews.set(selected.row, review);
  const key = reviewKey(selected);
  const saved = writeStorage(key, review);
  if (saved) unsavedReviews.delete(key);
  else unsavedReviews.set(key, review);
  trackReviewEdit(entry, previous, review, event.target);
  $("save-status").textContent = saved ? "已儲存至此瀏覽器" : "瀏覽器無法儲存，請保留此頁。";
  updateDetail();
  drawSummary();
  drawTable();
});

$("search").addEventListener("input", drawTable);
$("filter").addEventListener("change", drawTable);
workbookToggle.addEventListener("click", () => {
  workbookDrawer.showModal();
  workbookToggle.setAttribute("aria-expanded", "true");
  workbookDrawer.querySelector('[aria-pressed="true"]')?.focus();
});
$("workbook-close").addEventListener("click", () => workbookDrawer.close());
workbookDrawer.addEventListener("click", (event) => {
  if (event.target !== workbookDrawer) return;
  const bounds = workbookDrawer.getBoundingClientRect();
  if (event.clientX < bounds.left || event.clientX >= bounds.right) workbookDrawer.close();
});
workbookDrawer.addEventListener("close", () => {
  workbookToggle.setAttribute("aria-expanded", "false");
});
$("sheet").addEventListener("change", (event) => {
  const index = Number(event.target.value);
  workbooks[activeWorkbook].sheetIndex = index;
  openSheet(index);
  if (!saveWorkbooks()) message("已切換工作表。瀏覽器無法保留這次切換。", true);
});
$("import").addEventListener("click", () => $("file").click());
$("file").addEventListener("change", async (event) => {
  const files = [...event.target.files];
  if (!files.length) return;
  const importingOperator = operator;
  event.target.value = "";
  $("import").disabled = true;
  try {
    const failures = [];
    let imported = 0;
    let lastWorkbook = activeWorkbook;
    for (const file of files) {
      try {
        const data = await readWorkbookFile(file);
        lastWorkbook = addWorkbook(data);
        track("workbook_import", {
          workbook: data.name, version: workbooks[lastWorkbook].version, worksheetCount: data.sheets.length,
        }, importingOperator);
        imported++;
      } catch (error) {
        const detail = /^找不到|^請選擇/.test(error.message) ? error.message : "無法讀取這份 Excel。";
        failures.push(`${file.name}：${detail}`);
        track("workbook_import_failed", { workbook: file.name, message: detail }, importingOperator);
      }
    }
    const notices = [];
    let saved = true;
    if (imported) {
      drawWorkbookList();
      selectWorkbook(lastWorkbook);
      saved = saveWorkbooks();
      notices.push(`已匯入 ${imported} 個檔案，清單共 ${workbooks.length} 個。`);
      if (!saved) notices.push("瀏覽器無法儲存這次匯入，請保留此頁。");
    }
    message([...notices, ...failures].join(" "), failures.length > 0 || !saved);
  } finally {
    $("import").disabled = false;
  }
});

if (trackingEnabled) {
  restoreTrackingOutbox();
  flushTracking();
  window.addEventListener("online", flushTracking);
  window.addEventListener("pagehide", commitReviewLog);
}

function initializeDashboard() {
  const saved = readStorage("workbooks");
  let index;
  if (saved?.workbooks?.[saved.activeWorkbook]?.data?.sheets?.length) {
    workbooks = saved.workbooks;
    index = saved.activeWorkbook;
  } else {
    workbooks = [];
    const { name, base64 } = window.demoWorkbook;
    const bytes = Uint8Array.from(atob(base64), (character) => character.charCodeAt(0));
    index = addWorkbook(readProject(name, bytes));
    const previous = readStorage("active");
    if (previous?.project?.sheets?.[previous.index]?.wires?.length) {
      index = addWorkbook(previous.project, previous.index);
      workbooks[index].sheetIndex = previous.index;
    }
  }
  drawWorkbookList();
  selectWorkbook(index);
}

try {
  initializeDashboard();
} catch {
  drawWorkbookList();
  message("載入展示資料失敗，請重新匯入 Excel。", true);
  $("empty-detail").hidden = false;
}
