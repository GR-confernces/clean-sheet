"use strict";

const EMAIL_FIND = /[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}/gi;
const EMAIL_TEST = /[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}/i;
const HEADER_EMAIL = /e[\s\-_.]*mail|email|mail(?:[\s\-_]*(?:id|address|addr))?|gmail|ई\s*[\-.]?\s*मेल|इ\s*[\-.]?\s*मेल|मेल(?:\s*आई?\s*डी)?/i;

const state = {
  buffer: null,
  fileName: "",
  workbook: null,
  sheetName: "",
  texts: [],
  sources: [],
  emailCol: 0,
  hasHeader: true,
  headerRow: 0,
  detection: null,
  mode: "keep-first",
  preview: "kept",
  result: null,
};

const $ = (id) => document.getElementById(id);

function esc(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function colLetter(index) {
  let n = index + 1;
  let label = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

function emailKey(value) {
  const text = String(value ?? "").trim().toLowerCase();
  if (!text) return "";
  EMAIL_FIND.lastIndex = 0;
  const matches = text.match(EMAIL_FIND);
  if (!matches) return text;
  return matches.join("|");
}

function cellParts(cell) {
  if (!cell) return { text: "", source: "" };
  let text = "";
  if (cell.w != null) text = String(cell.w);
  else if (cell.v == null) text = "";
  else if (cell.v instanceof Date) text = cell.v.toLocaleDateString("en-GB");
  else if (typeof cell.v === "object") text = "";
  else text = String(cell.v);

  let source = text;
  const target = cell.l && cell.l.Target ? String(cell.l.Target) : "";
  const mailto = target.match(/mailto:([^?]+)/i);
  if (mailto) {
    try {
      source = decodeURIComponent(mailto[1]);
    } catch (err) {
      source = mailto[1];
    }
  }
  return { text, source };
}

function sheetToRows(sheet) {
  const texts = [];
  const sources = [];
  if (!sheet || !sheet["!ref"]) return { texts, sources };
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  for (let r = 0; r <= range.e.r; r += 1) {
    const textRow = [];
    const sourceRow = [];
    for (let c = 0; c <= range.e.c; c += 1) {
      const part = cellParts(sheet[XLSX.utils.encode_cell({ r, c })]);
      textRow.push(part.text);
      sourceRow.push(part.source);
    }
    texts.push(textRow);
    sources.push(sourceRow);
  }
  while (
    texts.length > 1 &&
    texts[texts.length - 1].every((cell) => String(cell).trim() === "")
  ) {
    texts.pop();
    sources.pop();
  }
  return { texts, sources };
}

function findLayout(texts) {
  const scan = Math.min(texts.length, 15);
  for (let r = 0; r < scan; r += 1) {
    const cells = (texts[r] || []).map((cell) => String(cell || "").trim()).filter(Boolean);
    if (!cells.length) continue;
    const named = cells.some((cell) => HEADER_EMAIL.test(cell));
    const emailValues = cells.filter((cell) => EMAIL_TEST.test(cell)).length;
    if (named && emailValues === 0) return { hasHeader: true, headerRow: r };
  }
  if (!texts.length) return { hasHeader: true, headerRow: 0 };
  const first = texts[0] || [];
  if (first.some((cell) => EMAIL_TEST.test(String(cell || "")))) {
    return { hasHeader: false, headerRow: -1 };
  }
  const labels = first.filter((cell) => {
    const text = String(cell || "").trim();
    return text && text.length <= 40 && !EMAIL_TEST.test(text) && !/^\d+$/.test(text);
  }).length;
  if (labels >= 1 && texts.length > 1) return { hasHeader: true, headerRow: 0 };
  return { hasHeader: false, headerRow: -1 };
}

function columnStats(sources, texts, col, headerRow) {
  const start = headerRow + 1;
  const header = headerRow >= 0 ? String(texts[headerRow]?.[col] || "") : "";
  let hits = 0;
  let total = 0;
  for (let r = start; r < sources.length; r += 1) {
    const raw = String(sources[r]?.[col] || "").trim();
    if (!raw) continue;
    total += 1;
    if (EMAIL_TEST.test(raw)) hits += 1;
  }
  return {
    index: col,
    hits,
    total,
    headerHit: HEADER_EMAIL.test(header),
  };
}

function detectEmailColumn(sources, texts, headerRow) {
  const width = texts.reduce((max, row) => Math.max(max, row.length), 0);
  let best = { index: 0, hits: 0, total: 0, headerHit: false, score: -1 };
  for (let c = 0; c < Math.max(width, 1); c += 1) {
    const stats = columnStats(sources, texts, c, headerRow);
    const ratio = stats.total ? stats.hits / stats.total : 0;
    const score = stats.hits * 100 + ratio * 20 + (stats.headerHit ? 40 : 0);
    if (score > best.score) best = { ...stats, score };
  }
  return best;
}

function dedupe(texts, sources, emailCol, headerRow, mode) {
  const start = headerRow + 1;
  const kept = [];
  const removed = [];
  const seen = new Map();
  const counts = new Map();
  let blank = 0;

  for (let i = 0; i < start && i < texts.length; i += 1) kept.push(texts[i]);

  for (let i = start; i < texts.length; i += 1) {
    const key = emailKey(sources[i]?.[emailCol]);
    if (!key) continue;
    counts.set(key, (counts.get(key) || 0) + 1);
  }

  for (let i = start; i < texts.length; i += 1) {
    const key = emailKey(sources[i]?.[emailCol]);
    if (!key) {
      blank += 1;
      kept.push(texts[i]);
      continue;
    }
    const duplicate = mode === "remove-all" ? counts.get(key) > 1 : seen.has(key);
    if (duplicate) {
      removed.push({
        excelRow: i + 1,
        key,
        texts: texts[i],
        copies: counts.get(key),
      });
    } else {
      seen.set(key, i + 1);
      kept.push(texts[i]);
    }
  }

  return {
    kept,
    removed,
    blank,
    dataRows: Math.max(texts.length - start, 0),
  };
}

function compute() {
  return dedupe(
    state.texts,
    state.sources,
    state.emailCol,
    state.hasHeader ? state.headerRow : -1,
    state.mode
  );
}

function readWorkbook(buffer, fileName) {
  if (/\.csv$/i.test(fileName)) {
    const bytes = buffer instanceof ArrayBuffer ? new Uint8Array(buffer) : buffer;
    const text = new TextDecoder("utf-8").decode(bytes).replace(/^\uFEFF/, "");
    return XLSX.read(text, { type: "string" });
  }
  return XLSX.read(buffer, { type: "array", cellDates: true });
}

function cleanedFileName(name) {
  const base = String(name || "cleaned").replace(/\.(xlsx|xls|csv)$/i, "");
  return `${base}_cleaned.xlsx`;
}

function buildCleanWorkbook() {
  const result = compute();
  const workbook = readWorkbook(cloneBuffer(state.buffer), state.fileName);
  workbook.Sheets[state.sheetName] = XLSX.utils.aoa_to_sheet(result.kept);
  return workbook;
}

function cloneBuffer(buffer) {
  if (buffer instanceof ArrayBuffer) return buffer.slice(0);
  return buffer.slice();
}

function showError(message) {
  const error = $("error");
  error.hidden = !message;
  error.textContent = message || "";
}

function flash(message) {
  const toast = $("toast");
  toast.hidden = false;
  toast.textContent = message;
  clearTimeout(flash.timer);
  flash.timer = setTimeout(() => {
    toast.hidden = true;
  }, 4200);
}

function headerLabels() {
  const width = state.texts.reduce((max, row) => Math.max(max, row.length), 0);
  const labels = [];
  for (let c = 0; c < width; c += 1) {
    const letter = colLetter(c);
    if (state.headerRow >= 0) {
      const name = String(state.texts[state.headerRow]?.[c] || "").trim();
      labels.push(name ? `${letter} — ${name}` : letter);
    } else {
      labels.push(letter);
    }
  }
  return labels;
}

function inspectSheet(workbook, name) {
  const rows = sheetToRows(workbook.Sheets[name]);
  const layout = findLayout(rows.texts);
  const headerRow = layout.hasHeader ? layout.headerRow : -1;
  return {
    name,
    texts: rows.texts,
    sources: rows.sources,
    hasHeader: layout.hasHeader,
    headerRow,
    detection: detectEmailColumn(rows.sources, rows.texts, headerRow),
  };
}

function renderTable(result) {
  const wrap = $("table-wrap");
  const labels = headerLabels();
  const headerRow = state.hasHeader ? state.headerRow : -1;
  const rows = state.preview === "removed"
    ? result.removed.map((row) => ({ excelRow: row.excelRow, cells: row.texts }))
    : result.kept.slice(headerRow + 1).map((cells, index) => ({
        excelRow: excelRowForKept(result, index),
        cells,
      }));

  if (!rows.length) {
    wrap.innerHTML = `<p class="empty">${
      state.preview === "removed" ? "No duplicate rows found." : "No rows left to show."
    }</p>`;
    return;
  }

  const shown = rows.slice(0, state.preview === "removed" ? 200 : 50);
  const body = shown
    .map((row) => {
      const cells = labels
        .map((_, index) => {
          const value = row.cells?.[index] ?? "";
          const mark = index === state.emailCol ? "email" : "";
          return `<td class="${mark}" title="${esc(value)}">${esc(value)}</td>`;
        })
        .join("");
      return `<tr class="${state.preview === "removed" ? "removed" : ""}"><td>${row.excelRow}</td>${cells}</tr>`;
    })
    .join("");
  const head = labels
    .map((label, index) => `<th class="${index === state.emailCol ? "email" : ""}">${esc(label)}</th>`)
    .join("");
  const more = rows.length > shown.length
    ? `<p class="summary">Showing ${shown.length} of ${rows.length} rows.</p>`
    : "";
  wrap.innerHTML = `<div class="table-scroll"><table><thead><tr><th>Source row</th>${head}</tr></thead><tbody>${body}</tbody></table></div>${more}`;
}

function excelRowForKept(result, keptDataIndex) {
  const removedRows = new Set(result.removed.map((row) => row.excelRow));
  const headerRow = state.hasHeader ? state.headerRow : -1;
  const start = headerRow + 2;
  let seen = -1;
  for (let excelRow = start; excelRow <= state.texts.length; excelRow += 1) {
    if (removedRows.has(excelRow)) continue;
    seen += 1;
    if (seen === keptDataIndex) return excelRow;
  }
  return keptDataIndex + start;
}

function duplicateList(result) {
  const grouped = new Map();
  result.removed.forEach((row) => {
    grouped.set(row.key, (grouped.get(row.key) || 0) + 1);
  });
  return [...grouped.entries()]
    .slice(0, 8)
    .map(([key, count]) => `${key} (${count})`)
    .join(", ");
}

function renderNote(result, stats) {
  const box = $("note-file");
  const list = $("note-points");
  const columnName = headerLabels()[state.emailCol] || colLetter(state.emailCol);
  const keptData = result.dataRows - result.removed.length;
  const headerText = state.hasHeader ? `Row ${state.headerRow + 1}` : "No header row found";
  const samples = [];
  const start = (state.hasHeader ? state.headerRow : -1) + 1;
  for (let r = start; r < state.sources.length && samples.length < 3; r += 1) {
    const key = emailKey(state.sources[r]?.[state.emailCol]);
    if (key.includes("@") && !samples.includes(key)) samples.push(key);
  }
  const removedLabel = result.removed.length === 1 ? "1 row will be removed" : `${result.removed.length} rows will be removed`;
  const keptLabel = keptData === 1 ? "1 row will remain" : `${keptData} rows will remain`;
  const rows = [
    ["File", state.fileName],
    ["Sheet", state.sheetName],
    ["Email column", stats.hits ? `${columnName} · ${stats.hits} emails` : "No clear email column found"],
    ["Header", headerText],
    ["Rule", "The first row of each email stays. Later duplicate rows are removed in full."],
    ["Result", `${removedLabel}. ${keptLabel}.`],
  ];
  if (samples.length) rows.push(["Sample", samples.join(", ")]);
  const duplicates = duplicateList(result);
  if (duplicates) rows.push(["Emails to remove", duplicates]);
  if (result.blank) {
    rows.push(["Blank emails", result.blank === 1 ? "1 row kept." : `${result.blank} rows kept.`]);
  }
  if (state.workbook.SheetNames.length > 1) {
    rows.push(["Other sheets", "They stay unchanged in the download."]);
  }
  list.innerHTML = rows
    .map(([label, value]) => `<div><dt>${esc(label)}</dt><dd>${esc(value)}</dd></div>`)
    .join("");
  box.hidden = false;
}

function refresh() {
  const result = compute();
  state.result = result;
  const keptData = result.dataRows - result.removed.length;
  const headerRow = state.hasHeader ? state.headerRow : -1;
  const stats = columnStats(state.sources, state.texts, state.emailCol, headerRow);
  const columnName = headerLabels()[state.emailCol] || colLetter(state.emailCol);
  $("file-name").textContent = state.fileName;
  $("stat-total").textContent = String(result.dataRows);
  $("stat-removed").textContent = String(result.removed.length);
  $("stat-kept").textContent = String(keptData);
  $("tab-removed").textContent = `Removed rows (${result.removed.length})`;
  $("tab-kept").classList.toggle("active", state.preview === "kept");
  $("tab-removed").classList.toggle("active", state.preview === "removed");
  if (stats.hits === 0) {
    $("summary").textContent = "No clear email column in this sheet. See the note below.";
  } else if (result.removed.length) {
    $("summary").textContent = `${state.sheetName} · ${columnName}. ${result.removed.length} duplicate rows will be removed. ${keptData} will remain.`;
  } else {
    $("summary").textContent = `${state.sheetName} · ${columnName}. No duplicate emails found.`;
  }
  renderNote(result, stats);
  renderTable(result);
}

function openWorkbook(buffer, fileName) {
  const workbook = readWorkbook(buffer, fileName);
  if (!workbook.SheetNames.length) throw new Error("empty");
  let best = null;
  workbook.SheetNames.forEach((name) => {
    const info = inspectSheet(workbook, name);
    if (!best || info.detection.score > best.detection.score) best = info;
  });
  state.buffer = buffer;
  state.fileName = fileName;
  state.workbook = workbook;
  state.sheetName = best.name;
  state.texts = best.texts;
  state.sources = best.sources;
  state.hasHeader = best.hasHeader;
  state.headerRow = best.headerRow;
  state.emailCol = best.detection.index;
  state.detection = best.detection;
  state.preview = "kept";
  $("change-file").textContent = "Change file";
  $("workspace").hidden = false;
  showError("");
  refresh();
  $("workspace").scrollIntoView({ behavior: "smooth", block: "start" });
}

function handleFile(file) {
  if (!file) return;
  if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
    showError("Choose an .xlsx, .xls, or .csv file.");
    return;
  }
  showError("");
  file.arrayBuffer()
    .then((buffer) => openWorkbook(buffer, file.name))
    .catch(() => {
      $("workspace").hidden = true;
      showError("This file could not be opened. Use .xlsx, .xls, or .csv. Password-protected files are not supported.");
    })
    .finally(() => {
      $("file-input").value = "";
    });
}

function buildExampleBuffer() {
  const contacts = [
    ["Customer list", "", ""],
    ["Name", "City", "Mail Id"],
    ["Amit", "Delhi", "amit@test.com"],
    ["Amit Sharma", "Mumbai", "amit@test.com"],
    ["Neha", "Pune", "NEHA@Test.com"],
    ["Neha Singh", "Nagpur", "neha@test.com"],
    ["Ravi", "Jaipur", "ravi@test.com"],
    ["Priya", "Kota", ""],
    ["Priya", "Agra", ""],
    ["Amit", "Goa", "amit@other.com"],
    ["Karan", "Lucknow", "  amit@test.com  "],
    ["Sana", "Indore", "SANA@hello.com"],
    ["Amit Joshi", "Noida", "amit@test.com"],
  ];
  const notes = [["Note"], ["This sheet should stay unchanged."]];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(contacts), "Contacts");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(notes), "Notes");
  return XLSX.write(workbook, { type: "array", bookType: "xlsx" });
}

function setPreview(next) {
  state.preview = next;
  $("tab-kept").classList.toggle("active", next === "kept");
  $("tab-removed").classList.toggle("active", next === "removed");
  $("tab-kept").setAttribute("aria-selected", String(next === "kept"));
  $("tab-removed").setAttribute("aria-selected", String(next === "removed"));
  if (state.result) renderTable(state.result);
}

function resetWorkspace() {
  state.buffer = null;
  state.workbook = null;
  state.texts = [];
  state.sources = [];
  state.result = null;
  $("workspace").hidden = true;
  $("note-file").hidden = true;
  $("change-file").textContent = "Choose file";
  $("file-input").value = "";
  showError("");
}

async function loadGoogleSheet(url) {
  const button = $("sheet-submit");
  button.disabled = true;
  button.textContent = "Reading...";
  showError("");
  try {
    const response = await fetch(`/api/google-sheet?url=${encodeURIComponent(url)}`);
    if (!response.ok) {
      throw new Error((await response.text()) || "The Google Sheet could not be opened.");
    }
    const encoded = response.headers.get("X-Filename");
    const name = encoded ? decodeURIComponent(encoded) : "google-sheet.xlsx";
    openWorkbook(await response.arrayBuffer(), name);
  } catch (err) {
    const offline = !err.message || err.message === "Failed to fetch";
    showError(offline
      ? "Could not reach the sheet service. Check your connection and try again."
      : err.message);
  } finally {
    button.disabled = false;
    button.textContent = "Load sheet";
  }
}

function bind() {
  const input = $("file-input");
  const dropzone = $("dropzone");

  input.addEventListener("change", () => handleFile(input.files && input.files[0]));
  dropzone.addEventListener("click", (event) => {
    if (event.target === input) return;
    input.click();
  });
  dropzone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      input.click();
    }
  });
  dropzone.addEventListener("dragover", (event) => {
    event.preventDefault();
    dropzone.classList.add("dragover");
  });
  dropzone.addEventListener("dragleave", () => dropzone.classList.remove("dragover"));
  dropzone.addEventListener("drop", (event) => {
    event.preventDefault();
    dropzone.classList.remove("dragover");
    handleFile(event.dataTransfer.files && event.dataTransfer.files[0]);
  });

  $("example").addEventListener("click", () => {
    try {
      openWorkbook(buildExampleBuffer(), "example.xlsx");
    } catch (err) {
      showError("The example file could not be created.");
    }
  });

  $("sheet-form").addEventListener("submit", (event) => {
    event.preventDefault();
    const url = $("sheet-url").value.trim();
    if (!url) {
      showError("Paste a Google Sheet link.");
      return;
    }
    loadGoogleSheet(url);
  });

  $("change-file").addEventListener("click", () => {
    $("file-input").click();
  });

  $("download").addEventListener("click", () => {
    if (!state.buffer) return;
    const name = cleanedFileName(state.fileName);
    XLSX.writeFile(buildCleanWorkbook(), name);
    flash(`Download started: ${name}`);
  });

  $("reset").addEventListener("click", resetWorkspace);
  $("tab-kept").addEventListener("click", () => setPreview("kept"));
  $("tab-removed").addEventListener("click", () => setPreview("removed"));
}

bind();

window.CleanSheet = {
  buildCleanWorkbook,
  openWorkbook,
  getState: () => state,
};
