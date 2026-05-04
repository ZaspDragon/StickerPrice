document.addEventListener("DOMContentLoaded", () => {
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
    return `${po || "NOPO"}-${item || "NOITEM"}-${stamp}`.replace(/\s+/g, "");
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
    copies = Math.max(1, Number(copies) || 1);

    if (!po || !item) {
      alert("Enter at least PO / Transfer # and Item #.");
      return;
    }

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
    alert(`${copies} sticker label(s) added.`);
  }

  function renderLabels() {
    const sheet = $("labelSheet");
    const template = $("labelTemplate");

    if (!sheet || !template) {
      console.error("Missing labelSheet or labelTemplate.");
      return;
    }

    sheet.innerHTML = "";

    state.labels.forEach((label) => {
      const node = template.content.cloneNode(true);

      node.querySelector(".po").textContent = label.po || "";
      node.querySelector(".item").textContent = label.item || "";
      node.querySelector(".qty").textContent = label.qty || "";
      node.querySelector(".desc").textContent = label.description || "";
      node.querySelector(".lineid").textContent = label.lineId || "";

      const canvas = node.querySelector(".qr");

      if (window.QRCode) {
        QRCode.toCanvas(canvas, labelPayload(label), {
          width: 120,
          margin: 1,
          errorCorrectionLevel: "M"
        });
      } else {
        canvas.replaceWith(document.createTextNode("QR library failed"));
      }

      sheet.appendChild(node);
    });
  }

  function addBulkLabels() {
    const raw = $("bulkText").value.trim();

    if (!raw) {
      alert("Paste at least one bulk line first.");
      return;
    }

    const lines = raw.split("\n").map(x => x.trim()).filter(Boolean);
    let added = 0;

    lines.forEach((line) => {
      const [po, item, qty, ...descParts] = line.split(",");

      if (!po || !item) return;

      state.labels.push({
        lineId: makeLineId(po.trim(), item.trim()),
        po: po.trim(),
        item: item.trim(),
        qty: (qty || "").trim(),
        description: descParts.join(",").trim(),
        createdAt: new Date().toISOString()
      });

      added++;
    });

    save();
    renderLabels();
    $("bulkText").value = "";

    alert(`${added} bulk sticker label(s) added.`);
  }

  function renderLog() {
    const body = $("checkLogBody");
    if (!body) return;

    body.innerHTML = "";

    state.checks.forEach((row) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${new Date(row.checkedAt).toLocaleString()}</td>
        <td>${row.worker || ""}</td>
        <td>${row.po || ""}</td>
        <td>${row.item || ""}</td>
        <td>${row.qty || ""}</td>
        <td>${row.description || ""}</td>
        <td>${row.lineId || ""}</td>
      `;
      body.appendChild(tr);
    });

    const today = todayKey();
    const worker = $("workerName").value.trim();

    const todayRows = state.checks.filter(row => row.date === today);
    const workerRows = worker
      ? todayRows.filter(row => row.worker.toLowerCase() === worker.toLowerCase())
      : [];

    $("todayTotal").textContent = todayRows.length;
    $("workerTotal").textContent = worker ? workerRows.length : 0;
  }

  function parseScanData(raw) {
    const clean = raw.trim();
    if (!clean) throw new Error("Nothing was scanned.");

    const data = JSON.parse(clean);

    if (data.type !== "line_check" || !data.lineId) {
      throw new Error("This QR is not a line check label.");
    }

    return data;
  }

  function markChecked(data) {
    const worker = $("workerName").value.trim() || "Unknown Worker";

    if (state.checks.some(row => row.lineId === data.lineId)) {
      alert("This line was already checked.");
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

  function exportCsv() {
    const headers = ["Checked At", "Date", "Worker", "PO/Transfer", "Item", "Qty", "Description", "Line ID"];

    const rows = state.checks.map(r => [
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
      .map(row => row.map(cell => `"${String(cell ?? "").replaceAll('"', '""')}"`).join(","))
      .join("\n");

    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = `line-check-log-${todayKey()}.csv`;
    a.click();

    URL.revokeObjectURL(url);
  }

  $("addLabelBtn")?.addEventListener("click", () => {
    addLabel({
      po: $("poNumber").value.trim(),
      item: $("itemNumber").value.trim(),
      qty: $("quantity").value.trim(),
      description: $("description").value.trim(),
      copies: $("labelCopies").value
    });
  });

  $("bulkAddBtn")?.addEventListener("click", addBulkLabels);

  $("checkLineBtn")?.addEventListener("click", () => {
    try {
      markChecked(parseScanData($("scanInput").value));
    } catch (err) {
      alert("QR data was not readable. Scan again.");
    }
  });

  $("printLabelsBtn")?.addEventListener("click", () => window.print());

  $("clearLabelsBtn")?.addEventListener("click", () => {
    if (confirm("Clear all printable labels?")) {
      state.labels = [];
      save();
      renderLabels();
    }
  });

  $("resetChecksBtn")?.addEventListener("click", () => {
    if (confirm("Reset all checked line history?")) {
      state.checks = [];
      save();
      renderLog();
    }
  });

  $("exportCsvBtn")?.addEventListener("click", exportCsv);
  $("workerName")?.addEventListener("input", renderLog);

  renderLabels();
  renderLog();
});
