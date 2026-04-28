const $ = (id) => document.getElementById(id);

const state = {
  labels: JSON.parse(localStorage.getItem("lineLabels") || "[]"),
  checks: JSON.parse(localStorage.getItem("lineChecks") || "[]")
};

function save() {
  localStorage.setItem("lineLabels", JSON.stringify(state.labels));
  localStorage.setItem("lineChecks", JSON.stringify(state.checks));
}

function todayKey(date = new Date()) {
  return date.toISOString().slice(0, 10);
}

function makeLineId(po, item) {
  const stamp = Date.now().toString(36).toUpperCase();
  const safePo = String(po || "NOPO").replace(/\s+/g, "");
  const safeItem = String(item || "NOITEM").replace(/\s+/g, "");
  return `${safePo}-${safeItem}-${stamp}`;
}

function labelPayload(label) {
  return JSON.stringify({
    type: "line_check",
    lineId: label.lineId,
    po: label.po,
    item: label.item,
    qty: label.qty,
    description: label.description,
    createdAt: label.createdAt
  });
}

function addLabel({ po, item, qty, description, copies = 1 }) {
  for (let i = 0; i < copies; i++) {
    state.labels.push({
      lineId: makeLineId(po, item),
      po,
      item,
      qty,
      description,
      createdAt: new Date().toISOString()
    });
  }
  save();
  renderLabels();
}

function renderLabels() {
  const sheet = $("labelSheet");
  const template = $("labelTemplate");
  sheet.innerHTML = "";

  state.labels.forEach((label) => {
    const node = template.content.cloneNode(true);
    node.querySelector(".po").textContent = label.po || "";
    node.querySelector(".item").textContent = label.item || "";
    node.querySelector(".qty").textContent = label.qty || "";
    node.querySelector(".desc").textContent = label.description || "";
    node.querySelector(".lineid").textContent = label.lineId || "";

    const canvas = node.querySelector(".qr");
    QRCode.toCanvas(canvas, labelPayload(label), {
      width: 120,
      margin: 1,
      errorCorrectionLevel: "M"
    });

    sheet.appendChild(node);
  });
}

function parseScanData(raw) {
  const clean = raw.trim();
  if (!clean) throw new Error("Nothing was scanned.");

  try {
    const data = JSON.parse(clean);
    if (data.type !== "line_check" || !data.lineId) {
      throw new Error("This QR is not a line check label.");
    }
    return data;
  } catch (err) {
    throw new Error("QR data was not readable. Try scanning again.");
  }
}

function markChecked(data) {
  const worker = $("workerName").value.trim() || "Unknown Worker";
  const alreadyChecked = state.checks.some((row) => row.lineId === data.lineId);

  if (alreadyChecked) {
    alert("This line was already checked. No duplicate count added.");
    return;
  }

  state.checks.unshift({
    checkedAt: new Date().toISOString(),
    date: todayKey(),
    worker,
    lineId: data.lineId,
    po: data.po || "",
    item: data.item || "",
    qty: data.qty || "",
    description: data.description || ""
  });

  save();
  renderLog();
  $("scanInput").value = "";
}

function renderLog() {
  const body = $("checkLogBody");
  body.innerHTML = "";

  state.checks.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${new Date(row.checkedAt).toLocaleString()}</td>
      <td>${escapeHtml(row.worker)}</td>
      <td>${escapeHtml(row.po)}</td>
      <td>${escapeHtml(row.item)}</td>
      <td>${escapeHtml(row.qty)}</td>
      <td>${escapeHtml(row.description)}</td>
      <td>${escapeHtml(row.lineId)}</td>
    `;
    body.appendChild(tr);
  });

  const today = todayKey();
  const worker = $("workerName").value.trim();
  const todayRows = state.checks.filter((row) => row.date === today);
  const workerRows = worker
    ? todayRows.filter((row) => row.worker.toLowerCase() === worker.toLowerCase())
    : [];

  $("todayTotal").textContent = todayRows.length;
  $("workerTotal").textContent = worker ? workerRows.length : 0;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function exportCsv() {
  const headers = ["Checked At", "Date", "Worker", "PO/Transfer", "Item", "Qty", "Description", "Line ID"];
  const rows = state.checks.map((r) => [
    new Date(r.checkedAt).toLocaleString(),
    r.date,
    r.worker,
    r.po,
    r.item,
    r.qty,
    r.description,
    r.lineId
  ]);

  const csv = [headers, ...rows]
    .map((row) => row.map((cell) => `"${String(cell ?? "").replaceAll('"', '""')}"`).join(","))
    .join("\n");

  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `line-check-log-${todayKey()}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

$("addLabelBtn").addEventListener("click", () => {
  addLabel({
    po: $("poNumber").value.trim(),
    item: $("itemNumber").value.trim(),
    qty: $("quantity").value.trim(),
    description: $("description").value.trim(),
    copies: Number($("labelCopies").value || 1)
  });
});

$("bulkAddBtn").addEventListener("click", () => {
  const lines = $("bulkText").value.split("\n").map((x) => x.trim()).filter(Boolean);

  lines.forEach((line) => {
    const [po, item, qty, ...descParts] = line.split(",");
    addLabel({
      po: (po || "").trim(),
      item: (item || "").trim(),
      qty: (qty || "").trim(),
      description: descParts.join(",").trim(),
      copies: 1
    });
  });

  $("bulkText").value = "";
});

$("checkLineBtn").addEventListener("click", () => {
  try {
    const data = parseScanData($("scanInput").value);
    markChecked(data);
  } catch (err) {
    alert(err.message);
  }
});

$("printLabelsBtn").addEventListener("click", () => window.print());

$("clearLabelsBtn").addEventListener("click", () => {
  if (confirm("Clear all printable labels?")) {
    state.labels = [];
    save();
    renderLabels();
  }
});

$("resetChecksBtn").addEventListener("click", () => {
  if (confirm("Reset all checked line history on this device?")) {
    state.checks = [];
    save();
    renderLog();
  }
});

$("exportCsvBtn").addEventListener("click", exportCsv);
$("workerName").addEventListener("input", renderLog);

renderLabels();
renderLog();
