document.addEventListener("DOMContentLoaded", () => {
  const $ = (id) => document.getElementById(id);

  const state = {
    labels: JSON.parse(localStorage.getItem("lineLabels") || "[]"),
    checks: JSON.parse(localStorage.getItem("lineChecks") || "[]")
  };

  let importedPoRows = [];

  function save() {
    localStorage.setItem("lineLabels", JSON.stringify(state.labels));
    localStorage.setItem("lineChecks", JSON.stringify(state.checks));
  }

  function todayKey(date = new Date()) {
    return date.toISOString().slice(0, 10);
  }

  function makeLineId(po, item, location = "") {
    const stamp = Date.now().toString(36).toUpperCase();
    const safePo = String(po || "NOPO").replace(/\s+/g, "");
    const safeItem = String(item || "NOITEM").replace(/\s+/g, "");
    const safeLoc = String(location || "").replace(/\s+/g, "");
    return `${safePo}-${safeItem}-${safeLoc}-${stamp}`;
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function labelPayload(label) {
    return JSON.stringify({
      type: "line_check",
      lineId: label.lineId,
      po: label.po,
      item: label.item,
      qty: label.qty,
      description: label.description,
      location: label.location || "",
      createdAt: label.createdAt
    });
  }

  function addLabel({ po, item, qty, description, location, copies = 1 }) {
    copies = Math.max(1, Number(copies) || 1);

    if (!po || !item) {
      alert("Enter at least PO / Transfer # and Item #.");
      return;
    }

    for (let i = 0; i < copies; i++) {
      state.labels.push({
        lineId: makeLineId(po, item, location),
        po,
        item,
        qty,
        description,
        location,
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

    if (!sheet || !template) return;

    sheet.innerHTML = "";

    state.labels.forEach((label) => {
      const node = template.content.cloneNode(true);

      node.querySelector(".po").textContent = label.po || "";
      node.querySelector(".item").textContent = label.item || "";
      node.querySelector(".qty").textContent = label.qty || "";
      node.querySelector(".desc").textContent = label.description || "";

      const locEl = node.querySelector(".location");
      if (locEl) locEl.textContent = label.location || "";

      node.querySelector(".lineid").textContent = label.lineId || "";

      const canvas = node.querySelector(".qr");

      if (window.QRCode) {
        QRCode.toCanvas(canvas, labelPayload(label), {
          width: 120,
          margin: 1,
          errorCorrectionLevel: "M"
        }).catch(() => {
          canvas.replaceWith(document.createTextNode("QR failed"));
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

    const lines = raw.split("\n").map((x) => x.trim()).filter(Boolean);
    let added = 0;
    let skipped = 0;

    lines.forEach((line) => {
      const [po, item, qty, description, location] = line.split(",");

      if (!po || !item) {
        skipped++;
        return;
      }

      state.labels.push({
        lineId: makeLineId(po.trim(), item.trim(), location?.trim() || ""),
        po: po.trim(),
        item: item.trim(),
        qty: (qty || "").trim(),
        description: (description || "").trim(),
        location: (location || "").trim(),
        createdAt: new Date().toISOString()
      });

      added++;
    });

    save();
    renderLabels();
    $("bulkText").value = "";

    alert(`Added ${added} bulk label(s). Skipped ${skipped}.`);
  }

  function parseScanData(raw) {
    const clean = raw.trim();

    if (!clean) {
      throw new Error("Nothing was scanned.");
    }

    const data = JSON.parse(clean);

    if (data.type !== "line_check" || !data.lineId) {
      throw new Error("This QR is not a line check label.");
    }

    return data;
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
      description: data.description || "",
      location: data.location || ""
    });

    save();
    renderLog();
    $("scanInput").value = "";
  }

  function renderLog() {
    const body = $("checkLogBody");
    if (!body) return;

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
        <td>${escapeHtml(row.location || "")}</td>
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

  function exportCsv() {
    const headers = [
      "Checked At",
      "Date",
      "Worker",
      "PO/Transfer",
      "Item",
      "Qty",
      "Description",
      "Location",
      "Line ID"
    ];

    const rows = state.checks.map((r) => [
      new Date(r.checkedAt).toLocaleString(),
      r.date,
      r.worker,
      r.po,
      r.item,
      r.qty,
      r.description,
      r.location || "",
      r.lineId
    ]);

    const csv = [headers, ...rows]
      .map((row) =>
        row.map((cell) => `"${String(cell ?? "").replaceAll('"', '""')}"`).join(",")
      )
      .join("\n");

    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = `line-check-log-${todayKey()}.csv`;
    a.click();

    URL.revokeObjectURL(url);
  }

  function cleanKey(value) {
    return String(value || "")
      .toLowerCase()
      .replace(/[^a-z0-9]/g, "");
  }

  function pickValue(row, possibleNames) {
    const keys = Object.keys(row);

    for (const name of possibleNames) {
      const target = cleanKey(name);
      const foundKey = keys.find((key) => cleanKey(key) === target);
      if (foundKey) return row[foundKey];
    }

    return "";
  }

  function normalizePoRow(row) {
    return {
      po: String(
        pickValue(row, [
          "PO",
          "PO #",
          "PO Number",
          "Purchase Order",
          "Transfer",
          "Transfer #",
          "Transfer Number"
        ]) || ""
      ).trim(),

      item: String(
        pickValue(row, [
          "Item",
          "Item #",
          "Item Number",
          "SKU",
          "Part Number",
          "Product Number"
        ]) || ""
      ).trim(),

      qty: String(
        pickValue(row, [
          "Qty",
          "Quantity",
          "QTY Received",
          "Received Qty",
          "Order Qty",
          "Ordered Qty"
        ]) || ""
      ).trim(),

      description: String(
        pickValue(row, [
          "Description",
          "Item Description",
          "Desc",
          "Product Description"
        ]) || ""
      ).trim(),

      location: String(
        pickValue(row, [
          "Location",
          "Bin",
          "Bin Location",
          "Loc",
          "Putaway Location",
          "Primary Location"
        ]) || ""
      ).trim()
    };
  }

  async function readPoFile() {
    const fileInput = $("poFileInput");
    const file = fileInput?.files?.[0];

    if (!file) {
      alert("Choose a CSV or Excel file first.");
      return [];
    }

    if (!window.XLSX) {
      alert("Excel/CSV reader failed to load. Check internet connection.");
      return [];
    }

    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    return rawRows
      .map(normalizePoRow)
      .filter((row) => row.po || row.item || row.qty || row.description || row.location);
  }

  function renderPoPreview(rows) {
    const body = $("poPreviewBody");
    const summary = $("importSummary");

    if (!body) return;

    body.innerHTML = "";

    rows.forEach((row) => {
      const tr = document.createElement("tr");

      tr.innerHTML = `
        <td>${escapeHtml(row.po)}</td>
        <td>${escapeHtml(row.item)}</td>
        <td>${escapeHtml(row.qty)}</td>
        <td>${escapeHtml(row.description)}</td>
        <td>${escapeHtml(row.location)}</td>
      `;

      body.appendChild(tr);
    });

    if (summary) {
      summary.textContent = `${rows.length} PO line(s) ready to import.`;
    }
  }

  $("addLabelBtn")?.addEventListener("click", () => {
    addLabel({
      po: $("poNumber").value.trim(),
      item: $("itemNumber").value.trim(),
      qty: $("quantity").value.trim(),
      description: $("description").value.trim(),
      location: $("location").value.trim(),
      copies: $("labelCopies").value
    });
  });

  $("bulkAddBtn")?.addEventListener("click", addBulkLabels);

  $("checkLineBtn")?.addEventListener("click", () => {
    try {
      const data = parseScanData($("scanInput").value);
      markChecked(data);
    } catch (err) {
      alert("QR data was not readable. Scan again.");
    }
  });

  $("printLabelsBtn")?.addEventListener("click", () => {
    const labels = document.querySelectorAll(".zebra-label");

    if (!labels.length) {
      alert("Add at least one label before printing.");
      return;
    }

    window.print();
  });

  $("clearLabelsBtn")?.addEventListener("click", () => {
    if (confirm("Clear all printable labels?")) {
      state.labels = [];
      save();
      renderLabels();
    }
  });

  $("resetChecksBtn")?.addEventListener("click", () => {
    if (confirm("Reset all checked line history on this device?")) {
      state.checks = [];
      save();
      renderLog();
    }
  });

  $("exportCsvBtn")?.addEventListener("click", exportCsv);

  $("workerName")?.addEventListener("input", renderLog);

  $("previewPoBtn")?.addEventListener("click", async () => {
    try {
      importedPoRows = await readPoFile();
      renderPoPreview(importedPoRows);
    } catch (err) {
      console.error(err);
      alert("Could not read this file. Make sure it is CSV, XLS, or XLSX.");
    }
  });

  $("importPoBtn")?.addEventListener("click", async () => {
    try {
      if (!importedPoRows.length) {
        importedPoRows = await readPoFile();
      }

      if (!importedPoRows.length) {
        alert("No PO lines found.");
        return;
      }

      let added = 0;
      let skipped = 0;

      importedPoRows.forEach((row) => {
        if (!row.po || !row.item) {
          skipped++;
          return;
        }

        const duplicate = state.labels.some((label) =>
          String(label.po).toLowerCase() === row.po.toLowerCase() &&
          String(label.item).toLowerCase() === row.item.toLowerCase() &&
          String(label.location || "").toLowerCase() === row.location.toLowerCase()
        );

        if (duplicate) {
          skipped++;
          return;
        }

        state.labels.push({
          lineId: makeLineId(row.po, row.item, row.location),
          po: row.po,
          item: row.item,
          qty: row.qty,
          description: row.description,
          location: row.location,
          createdAt: new Date().toISOString()
        });

        added++;
      });

      save();
      renderLabels();

      const summary = $("importSummary");
      if (summary) {
        summary.textContent = `Imported ${added} label(s). Skipped ${skipped}.`;
      }

      alert(`Imported ${added} label(s). Skipped ${skipped}.`);
    } catch (err) {
      console.error(err);
      alert("Import failed. Check the file columns and try again.");
    }
  });

  renderLabels();
  renderLog();
});
