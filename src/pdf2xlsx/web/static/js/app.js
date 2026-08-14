/* PDF → Excel Converter & Merger — web app client. */

"use strict";

const $ = (id) => document.getElementById(id);

const els = {
  dropzone: $("dropzone"),
  fileInput: $("file-input"),
  queue: $("queue"),
  password: $("password"),
  convertBtn: $("convert-btn"),
  clearFiles: $("clear-files"),
  uploadProgress: $("upload-progress"),
  convertStatus: $("convert-status"),
  outputsList: $("outputs-list"),
  outputsEmpty: $("outputs-empty"),
  mergeFiles: $("merge-files"),
  mergeMode: $("merge-mode"),
  mergeBtn: $("merge-btn"),
  mergeStatus: $("merge-status"),
  fieldChecks: $("field-checks"),
  extractBtn: $("extract-btn"),
  extractStatus: $("extract-status"),
  resetBtn: $("reset-btn"),
  toast: $("toast"),
};

const queue = new Map(); // id -> {file, name, size, status}

let toastTimer = null;

function formatSize(bytes) {
  if (bytes >= 1048576) return (bytes / 1048576).toFixed(1) + " MB";
  if (bytes >= 1024) return (bytes / 1024).toFixed(1) + " KB";
  return bytes + " B";
}

function toast(message, isError) {
  els.toast.textContent = message;
  els.toast.classList.toggle("err", !!isError);
  els.toast.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { els.toast.hidden = true; }, 4000);
}

function setConvertEnabled() {
  els.convertBtn.disabled = queue.size === 0;
  els.clearFiles.disabled = queue.size === 0;
}

/* ------------------------------------------------------------ file queue */
function addFiles(fileList) {
  for (const file of fileList) {
    if (!/\.pdf$/i.test(file.name)) {
      toast(`Skipped "${file.name}" — only PDF files are accepted.`, true);
      continue;
    }
    const id = file.name + "_" + file.size + "_" + file.lastModified;
    if (queue.has(id)) continue;
    queue.set(id, { file, name: file.name, size: file.size, status: "pending" });
  }
  renderQueue();
}

function renderQueue() {
  if (queue.size === 0) {
    els.queue.hidden = true;
    setConvertEnabled();
    return;
  }
  els.queue.hidden = false;
  els.queue.innerHTML = "";
  queue.forEach((item, id) => {
    const row = document.createElement("div");
    row.className = "file-item";

    const name = document.createElement("span");
    name.className = "fname";
    name.title = item.name;
    name.textContent = item.name;

    const size = document.createElement("span");
    size.className = "fsize";
    size.textContent = formatSize(item.size);

    const status = document.createElement("span");
    status.className = "fstatus " + item.status;
    const statusText = {
      pending: "queued",
      uploading: "uploading…",
      done: "converted ✓",
      error: "failed",
    }[item.status] || item.status;
    status.textContent = statusText;

    const rem = document.createElement("button");
    rem.className = "rem";
    rem.title = "Remove";
    rem.textContent = "✕";
    rem.addEventListener("click", () => {
      queue.delete(id);
      renderQueue();
    });

    row.append(name, size, status, rem);
    els.queue.appendChild(row);
  });
  setConvertEnabled();
}

/* ---------------------------------------------------------------- upload */
function uploadAll() {
  const form = new FormData();
  queue.forEach((item) => form.append("pdfs", item.file));
  if (els.password.value) form.append("password", els.password.value);

  const xhr = new XMLHttpRequest();
  xhr.open("POST", "/api/upload");

  xhr.upload.addEventListener("progress", (e) => {
    if (e.lengthComputable) {
      els.uploadProgress.hidden = false;
      const pct = Math.round((e.loaded / e.total) * 100);
      els.uploadProgress.textContent = `Uploading… ${pct}%`;
    }
  });

  xhr.addEventListener("load", () => {
    els.uploadProgress.hidden = true;
    els.convertStatus.hidden = false;
    let payload;
    try {
      payload = JSON.parse(xhr.responseText);
    } catch {
      els.convertStatus.className = "status error";
      els.convertStatus.textContent = "Server returned an invalid response.";
      return;
    }
    if (!payload.ok) {
      els.convertStatus.className = "status error";
      els.convertStatus.textContent = payload.error || "Conversion failed.";
      return;
    }
    // mark queue items
    let i = 0;
    queue.forEach((item) => {
      const result = payload.files[i++] || { ok: false };
      item.status = result.ok ? "done" : "error";
      item.detail = result;
    });
    renderQueue();
    els.convertStatus.className = "status info";
    const failed = payload.files.filter((f) => !f.ok);
    const okCount = payload.files.length - failed.length;
    els.convertStatus.textContent =
      `${okCount} file${okCount === 1 ? "" : "s"} converted. ` +
      (failed.length ? `${failed.length} failed — check the list.` : "");
    if (payload.outputs) renderOutputs(payload.outputs);
    refreshMergeOptions();
    refreshExtractEnabled();
  });

  xhr.addEventListener("error", () => {
    els.uploadProgress.hidden = true;
    els.convertStatus.className = "status error";
    els.convertStatus.textContent = "Network error — could not reach the server.";
  });

  queue.forEach((item) => { item.status = "uploading"; });
  renderQueue();
  els.convertStatus.hidden = false;
  els.convertStatus.className = "status info";
  els.convertStatus.textContent = "Uploading and converting…";
  xhr.send(form);
}

/* --------------------------------------------------------------- outputs */
function renderOutputs(outputs) {
  els.outputsList.innerHTML = "";
  els.outputsEmpty.hidden = outputs.length > 0;
  els.mergeFiles.innerHTML = "";
  els.mergeFiles.disabled = outputs.length === 0;
  els.mergeBtn.disabled = outputs.length === 0;

  for (const out of outputs) {
    const card = document.createElement("div");
    card.className = "output-card";

    const name = document.createElement("div");
    name.className = "oname";
    name.textContent = out.name;
    name.title = out.name;

    const size = document.createElement("div");
    size.className = "osize";
    size.textContent = formatSize(out.size);

    const dl = document.createElement("a");
    dl.className = "btn btn-primary btn-sm";
    dl.textContent = "Download";
    dl.href = "/api/download/" + encodeURIComponent(out.name);

    card.append(name, size, dl);
    els.outputsList.appendChild(card);

    const opt = document.createElement("option");
    opt.value = out.name;
    opt.textContent = out.name;
    els.mergeFiles.appendChild(opt);
  }
}

async function refreshOutputs() {
  try {
    const res = await fetch("/api/outputs");
    const payload = await res.json();
    if (payload.ok) {
      renderOutputs(payload.outputs);
      refreshMergeOptions();
      refreshExtractEnabled();
    }
  } catch {
    /* transient — ignore */
  }
}

function refreshMergeOptions() {
  const count = els.mergeFiles.options.length;
  els.mergeBtn.disabled = count === 0;
}

function refreshExtractEnabled() {
  els.extractBtn.disabled = els.mergeFiles.options.length === 0;
}

/* ----------------------------------------------------------------- merge */
async function doMerge() {
  const selected = Array.from(els.mergeFiles.selectedOptions).map((o) => o.value);
  const payload = { mode: els.mergeMode.value, files: selected };
  els.mergeStatus.hidden = false;
  els.mergeStatus.textContent = "Merging…";
  try {
    const res = await fetch("/api/merge", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || "Merge failed.");
    els.mergeStatus.textContent =
      `Merged ${data.result.sheets.length} sheets → ${data.result.name}`;
    if (data.outputs) renderOutputs(data.outputs);
  } catch (err) {
    els.mergeStatus.textContent = err.message;
    toast(err.message, true);
  }
}

/* -------------------------------------------------------------- extract */
function renderFieldChecks(fields) {
  els.fieldChecks.innerHTML = "";
  for (const field of fields) {
    const label = document.createElement("label");
    label.className = "check";
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.value = field.key;
    label.appendChild(cb);
    label.appendChild(document.createTextNode(field.label));
    els.fieldChecks.appendChild(label);
  }
}

async function doExtract() {
  const fields = Array.from(els.fieldChecks.querySelectorAll("input:checked"))
    .map((cb) => cb.value);
  if (fields.length === 0) {
    toast("Select at least one field to extract.", true);
    return;
  }
  const payload = { fields };
  els.extractStatus.hidden = false;
  els.extractStatus.textContent = "Extracting…";
  try {
    const res = await fetch("/api/extract", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (!data.ok) throw new Error(data.error || "Extraction failed.");
    const notes = (data.result.notes || []).length;
    els.extractStatus.textContent =
      `Extracted ${data.result.rows} row${data.result.rows === 1 ? "" : "s"} across ` +
      `${data.result.columns.length} columns → ${data.result.name}` +
      (notes ? ` (${notes} note${notes === 1 ? "" : "s"})` : "");
    if (data.outputs) renderOutputs(data.outputs);
  } catch (err) {
    els.extractStatus.textContent = err.message;
    toast(err.message, true);
  }
}

/* ----------------------------------------------------------------- reset */
async function doReset() {
  if (!confirm("Delete all uploaded files and generated workbooks for this session?")) return;
  try {
    await fetch("/api/reset", { method: "POST" });
  } catch { /* ignore */ }
  queue.clear();
  els.password.value = "";
  els.queue.hidden = true;
  els.convertStatus.hidden = true;
  els.mergeStatus.hidden = true;
  els.extractStatus.hidden = true;
  els.mergeFiles.innerHTML = "";
  els.outputsList.innerHTML = "";
  els.outputsEmpty.hidden = false;
  setConvertEnabled();
  refreshMergeOptions();
  toast("Session reset.");
  window.location.hash = "#convert";
}

/* ------------------------------------------------------------- bindings */
els.dropzone.addEventListener("click", () => els.fileInput.click());
els.dropzone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    els.fileInput.click();
  }
});
els.dropzone.addEventListener("dragover", (e) => {
  e.preventDefault();
  els.dropzone.classList.add("dragover");
});
els.dropzone.addEventListener("dragleave", () => els.dropzone.classList.remove("dragover"));
els.dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  els.dropzone.classList.remove("dragover");
  addFiles(e.dataTransfer.files);
});
els.fileInput.addEventListener("change", () => {
  addFiles(els.fileInput.files);
  els.fileInput.value = "";
});
els.convertBtn.addEventListener("click", uploadAll);
els.clearFiles.addEventListener("click", () => {
  queue.clear();
  renderQueue();
});
els.mergeBtn.addEventListener("click", doMerge);
els.extractBtn.addEventListener("click", doExtract);
els.resetBtn.addEventListener("click", doReset);

/* ---------------------------------------------------------------- init */
(async function init() {
  try {
    const res = await fetch("/api/fields");
    const data = await res.json();
    renderFieldChecks(data.fields);
  } catch {
    toast("Could not load extraction fields.", true);
  }
  refreshOutputs();
  setConvertEnabled();
})();
