document.addEventListener("DOMContentLoaded", () => {
  const $ = (id) => document.getElementById(id);

  const state = {
    labels: JSON.parse(localStorage.getItem("lineLabels") || "[]"),
    checks: JSON.parse(localStorage.getItem("lineChecks") || "[]")
  };

  let importedRows = [];

  function save() {
    localStorage.setItem("lineLabels", JSON.stringify(state.labels));
    localStorage.setItem("lineChecks", JSON.stringify(state.checks));
  }

  function todayKey(date = new Date()) {
    return date.toISOString().slice(0, 10);
  }

  function cleanText(value) {
    return String(value || "").trim();
  }

  function makeLineId(docType, docNumber, item, location = "") {
    const stamp = Date.now().toString(36).toUpperCase();
    return `${docType || "DOC"}-${docNumber || "NODOC"}-${item || "NOITEM"}-${location || "NOLOC"}-${stamp}`
      .replace(/\s+/g, "")
      .toUpperCase();
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
      docType: label.docType,
      docNumber: label.docNumber,
      item: label.item,
      qty: label.qty,
      description: label.description,
      location: label.location || "",
      createdAt: label.createdAt
    });
  }

  function addLabel({ docType, docNumber, item, qty, description, location, copies = 1 }) {
    copies = Math.max(1, Number(copies) || 1);

    if (!docNumber || !item) {
      alert("Enter at least Document # and Item #.");
      return;
    }

    for (let i = 0; i < copies; i++) {
      state.labels.push({
        lineId: makeLineId(docType, docNumber, item, location),
        docType,
        docNumber,
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

      node.querySelector(".docTypeText").textContent = label.docType || "";
      node.querySelector(".docNumber").textContent = label.docNumber || "";
      node.querySelector(".item").textContent = label.item || "";
      node.querySelector(".qty").textContent = label.qty || "";
      node.querySelector(".desc").textContent = label.description || "";
      node.querySelector(".location").textContent = label.location || "";
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
      const parts = line.split(",").map((x) => x.trim());

      let docType = parts[0] || $("docType").value || "PO";
      let docNumber = parts[1] || "";
      let item = parts[2] || "";
      let qty = parts[3] || "";
      let description = parts[4] || "";
      let location = parts[5] || "";

      docType = docType.toUpperCase();

      if (!["PO", "SPO", "SXFR", "XFR"].includes(docType)) {
        location = description;
        description = qty;
        qty = item;
        item = docNumber;
        docNumber = docType;
        docType = $("docType").value || "PO";
      }

      if (!docNumber || !item) {
        skipped++;
        return;
      }

      state.labels.push({
        lineId: makeLineId(docType, docNumber, item, location),
        docType,
        docNumber,
        item,
        qty,
        description,
        location,
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
    if (!clean) throw new Error("Nothing was scanned.");

    const data = JSON.parse(clean);

    if (data.type !== "line_check" || !data.lineId) {
      throw new Error("This QR is not a line check label.");
    }

    return data;
  }

  function markChecked(data) {
    const worker = $("workerName").value.trim() || "Unknown Worker";

    if (state.checks.some((row) => row.lineId === data.lineId)) {
      alert("This line was already checked. No duplicate count added.");
      return;
    }

    state.checks.unshift({
      checkedAt: new Date().toISOString(),
      date: todayKey(),
      worker,
      lineId: data.lineId,
      docType: data.docType || "",
      docNumber: data.docNumber || "",
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
        <td>${escapeHtml(row.docType)}</td>
        <td>${escapeHtml(row.docNumber)}</td>
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
      "Type",
      "Document Number",
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
      r.docType,
      r.docNumber,
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

  function detectDocType(row, docNumber) {
    const defaultDocType = $("defaultDocType")?.value || "PO";

    const directType = cleanText(
      pickValue(row, [
        "Type",
        "Doc Type",
        "Document Type",
        "Order Type",
        "Transfer Type"
      ])
    ).toUpperCase();

    if (["PO", "SPO", "SXFR", "XFR"].includes(directType)) {
      return directType;
    }

    const docString = cleanText(docNumber).toUpperCase();

    if (docString.startsWith("SPO")) return "SPO";
    if (docString.startsWith("SXFR")) return "SXFR";
    if (docString.startsWith("XFR")) return "XFR";
    if (docString.startsWith("PO")) return "PO";

    return defaultDocType;
  }

  function normalizeUploadedRow(row) {
    const docNumber = cleanText(
      pickValue(row, [
        "PO",
        "PO #",
        "PO Number",
        "Purchase Order",
        "SPO",
        "SPO #",
        "SPO Number",
        "SXFR",
        "SXFR #",
        "SXFR Number",
        "XFR",
        "XFR #",
        "XFR Number",
        "Transfer",
        "Transfer #",
        "Transfer Number",
        "Document",
        "Document #",
        "Document Number",
        "Order Number",
        "Order #",
        "Receipt Number",
        "Reference Number"
      ])
    );

    const docType = detectDocType(row, docNumber);

    return {
      docType,
      docNumber,
      item: cleanText(
        pickValue(row, [
          "Item",
          "Item #",
          "Item Number",
          "SKU",
          "Part Number",
          "Product Number",
          "Item ID"
        ])
      ),
      qty: cleanText(
        pickValue(row, [
          "Qty",
          "Quantity",
          "QTY Received",
          "Received Qty",
          "Order Qty",
          "Ordered Qty",
          "Qty Ordered",
          "Qty To Receive",
          "QTY To Receive"
        ])
      ),
      description: cleanText(
        pickValue(row, [
          "Description",
          "Item Description",
          "Desc",
          "Product Description",
          "Item Desc"
        ])
      ),
      location: cleanText(
        pickValue(row, [
          "Location",
          "Bin",
          "Bin Location",
          "Loc",
          "Putaway Location",
          "Primary Location",
          "To Bin",
          "From Bin"
        ])
      )
    };
  }

  async function readUploadedFile() {
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
      .map(normalizeUploadedRow)
      .filter((row) => row.docNumber || row.item || row.qty || row.description || row.location);
  }

  function renderUploadPreview(rows) {
    const body = $("poPreviewBody");
    const summary = $("importSummary");

    if (!body) return;

    body.innerHTML = "";

    rows.forEach((row) => {
      const tr = document.createElement("tr");

      tr.innerHTML = `
        <td>${escapeHtml(row.docType)}</td>
        <td>${escapeHtml(row.docNumber)}</td>
        <td>${escapeHtml(row.item)}</td>
        <td>${escapeHtml(row.qty)}</td>
        <td>${escapeHtml(row.description)}</td>
        <td>${escapeHtml(row.location)}</td>
      `;

      body.appendChild(tr);
    });

    if (summary) {
      summary.textContent = `${rows.length} line(s) ready to import.`;
    }
  }

  $("addLabelBtn")?.addEventListener("click", () => {
    addLabel({
      docType: $("docType").value,
      docNumber: $("poNumber").value.trim(),
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
      markChecked(parseScanData($("scanInput").value));
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
      importedRows = await readUploadedFile();
      renderUploadPreview(importedRows);
    } catch (err) {
      console.error(err);
      alert("Could not read this file. Make sure it is CSV, XLS, or XLSX.");
    }
  });

  $("importPoBtn")?.addEventListener("click", async () => {
    try {
      if (!importedRows.length) {
        importedRows = await readUploadedFile();
      }

      if (!importedRows.length) {
        alert("No lines found.");
        return;
      }

      let added = 0;
      let skipped = 0;

      importedRows.forEach((row) => {
        if (!row.docNumber || !row.item) {
          skipped++;
          return;
        }

        const duplicate = state.labels.some((label) =>
          String(label.docType).toLowerCase() === row.docType.toLowerCase() &&
          String(label.docNumber).toLowerCase() === row.docNumber.toLowerCase() &&
          String(label.item).toLowerCase() === row.item.toLowerCase() &&
          String(label.location || "").toLowerCase() === row.location.toLowerCase()
        );

        if (duplicate) {
          skipped++;
          return;
        }

        state.labels.push({
          lineId: makeLineId(row.docType, row.docNumber, row.item, row.location),
          docType: row.docType,
          docNumber: row.docNumber,
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
