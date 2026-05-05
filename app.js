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
    return String(value ?? "").trim();
  }

  function cleanKey(value) {
    return String(value ?? "")
      .toLowerCase()
      .replace(/[^a-z0-9]/g, "");
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function makeLineId(docType, docNumber, item, location = "") {
    const stamp = Date.now().toString(36).toUpperCase();
    return `${docType || "DOC"}-${docNumber || "NODOC"}-${item || "NOITEM"}-${location || "NOLOC"}-${stamp}`
      .replace(/\s+/g, "")
      .toUpperCase();
  }

  function formatDocDisplay(label) {
    const type = cleanText(label.docType).toUpperCase();
    const number = cleanText(label.docNumber).toUpperCase();

    if (!number && type) return type;
    if (!type) return number;

    if (number === type) return "";
    if (number.startsWith(type)) return number;

    return `${type}${number}`;
  }

  function labelPayload(label) {
    const payload = {
      type: "line_check",
      lineId: label.lineId,
      docType: label.docType,
      docNumber: label.docNumber,
      branch: label.branch || "",
      item: label.item,
      qty: label.qty,
      description: label.description,
      location: label.location || "",
      createdAt: label.createdAt
    };

    const encoded = encodeURIComponent(JSON.stringify(payload));
    return `${window.location.origin}${window.location.pathname}?scan=${encoded}`;
  }

  function handleUrlScan() {
    const params = new URLSearchParams(window.location.search);
    const scan = params.get("scan");

    if (!scan) return;

    try {
      const data = JSON.parse(decodeURIComponent(scan));

      if ($("scanInput")) {
        $("scanInput").value = JSON.stringify(data);
      }

      window.history.replaceState({}, document.title, window.location.pathname);
    } catch (err) {
      console.error("Could not read QR scan from URL", err);
    }
  }

  function pickValue(row, names) {
    const keys = Object.keys(row);

    for (const name of names) {
      const target = cleanKey(name);
      const foundKey = keys.find((key) => cleanKey(key) === target);
      if (foundKey && cleanText(row[foundKey]) !== "") return row[foundKey];
    }

    for (const name of names) {
      const target = cleanKey(name);
      const foundKey = keys.find((key) => {
        const c = cleanKey(key);
        return c.includes(target) || target.includes(c);
      });

      if (foundKey && cleanText(row[foundKey]) !== "") return row[foundKey];
    }

    return "";
  }

  function detectDocType(row, docNumber) {
    const def = $("defaultDocType")?.value || "PO";

    const direct = cleanText(
      pickValue(row, ["Type", "Doc Type", "Document Type", "Order Type", "Transfer Type"])
    ).toUpperCase();

    if (["PO", "SPO", "SXFR", "XFR"].includes(direct)) return direct;

    const d = cleanText(docNumber).toUpperCase();

    if (d.includes("SXFR")) return "SXFR";
    if (d.includes("SPO")) return "SPO";
    if (d.includes("XFR")) return "XFR";
    if (d.includes("PO")) return "PO";

    return def;
  }

  function findDocNumberFromRow(row) {
    for (const value of Object.values(row).map(cleanText)) {
      const upper = value.toUpperCase();
      const match = upper.match(/\b(SPO|SXFR|XFR|PO)[-:\s#]*([A-Z0-9-]+)\b/);
      if (match) return `${match[1]}${match[2]}`;
    }

    return "";
  }

  function fixDocNumber(row, docNumber, docType) {
    let number = cleanText(docNumber);
    const type = cleanText(docType).toUpperCase();

    if (!number || number.toUpperCase() === type) {
      const found = findDocNumberFromRow(row);
      if (found && found.toUpperCase() !== type) number = found;
    }

    return number;
  }

  function normalizeUploadedRow(row) {
    let docNumber = cleanText(
      pickValue(row, [
        "P.O.#",
        "P.O. #",
        "PO #",
        "PO Number",
        "Purchase Order",
        "SPO #",
        "SPO Number",
        "SXFR #",
        "SXFR Number",
        "XFR #",
        "XFR Number",
        "Transfer #",
        "Transfer Number",
        "Document #",
        "Document Number",
        "Doc #",
        "Doc Number",
        "Order Number",
        "Order #",
        "Receipt Number",
        "Receiver Number",
        "Reference Number",
        "PO",
        "SPO",
        "SXFR",
        "XFR",
        "Transfer",
        "Document",
        "Doc",
        "Receiver"
      ])
    );

    let docType = detectDocType(row, docNumber);
    docNumber = fixDocNumber(row, docNumber, docType);
    docType = detectDocType(row, docNumber);

    const branch =
      cleanText(
        pickValue(row, [
          "Branch",
          "Branch #",
          "Branch Number",
          "Warehouse",
          "Warehouse Branch",
          "Site",
          "Facility",
          "Location Branch"
        ])
      ) || cleanText($("defaultBranch")?.value || "");

    const item = cleanText(
      pickValue(row, [
        "Item",
        "Item #",
        "Item Number",
        "Item No",
        "Item ID",
        "SKU",
        "Part Number",
        "Product Number",
        "Material",
        "Product"
      ])
    );

    const qty = cleanText(
      pickValue(row, [
        "Qty",
        "QTY",
        "Quantity",
        "Qty Received",
        "QTY Received",
        "Received Qty",
        "Order Qty",
        "Ordered Qty",
        "Qty Ordered",
        "Qty To Receive",
        "QTY To Receive",
        "Transfer Qty",
        "Ship Qty"
      ])
    );

    const description = cleanText(
      pickValue(row, [
        "Description",
        "Item Description",
        "Desc",
        "Product Description",
        "Item Desc",
        "Name"
      ])
    );

    const location = cleanText(
      pickValue(row, [
        "Location",
        "Bin",
        "Bin Location",
        "Loc",
        "Putaway Location",
        "Primary Location",
        "To Bin",
        "From Bin",
        "Slot"
      ])
    );

    return {
      docType,
      docNumber,
      branch,
      item,
      qty,
      description,
      location
    };
  }

  function findBestHeaderRow(rows) {
    let bestIndex = 0;
    let bestScore = -1;

    rows.slice(0, 30).forEach((row, index) => {
      const joined = row.map(cleanKey).join(" ");
      let score = 0;

      if (joined.includes("item")) score += 4;
      if (joined.includes("qty") || joined.includes("quantity")) score += 4;
      if (joined.includes("description") || joined.includes("desc")) score += 2;
      if (joined.includes("location") || joined.includes("bin")) score += 2;
      if (joined.includes("branch") || joined.includes("warehouse")) score += 2;

      if (
        joined.includes("po") ||
        joined.includes("spo") ||
        joined.includes("xfr") ||
        joined.includes("document") ||
        joined.includes("receiver")
      ) {
        score += 3;
      }

      if (score > bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    });

    return bestIndex;
  }

  async function readUploadedFile() {
    const input = $("poFileInput");
    const file = input?.files?.[0];

    if (!file) {
      alert("Choose a CSV, XLS, or XLSX file first.");
      return [];
    }

    if (!window.XLSX) {
      alert("Excel/CSV reader failed to load. Check your internet connection.");
      return [];
    }

    const data = await file.arrayBuffer();

    const workbook = XLSX.read(data, {
      type: "array",
      cellDates: false,
      raw: false
    });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      raw: false
    });

    if (!rows.length) return [];

    const headerIndex = findBestHeaderRow(rows);
    const headers = rows[headerIndex] || [];
    const dataRows = rows.slice(headerIndex + 1);

    const objects = dataRows.map((row) => {
      const obj = {};
      headers.forEach((header, index) => {
        const name = cleanText(header) || `Column${index + 1}`;
        obj[name] = row[index] ?? "";
      });
      return obj;
    });

    return objects
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
        <td>${escapeHtml(formatDocDisplay(row))}</td>
        <td>${escapeHtml(row.branch || "")}</td>
        <td>${escapeHtml(row.item)}</td>
        <td>${escapeHtml(row.qty)}</td>
        <td>${escapeHtml(row.description)}</td>
        <td>${escapeHtml(row.location)}</td>
      `;

      body.appendChild(tr);
    });

    if (summary) summary.textContent = `${rows.length} line(s) ready to import.`;
  }

  function addLabel({ docType, docNumber, branch, item, qty, description, location, copies = 1 }) {
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
        branch,
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

      const typeEl = node.querySelector(".docTypeText");
      const branchEl = node.querySelector(".branch");

      if (typeEl) typeEl.textContent = label.docType || "";
      if (branchEl) branchEl.textContent = label.branch || "";

      node.querySelector(".docNumber").textContent = formatDocDisplay(label);
      node.querySelector(".item").textContent = label.item || "";
      node.querySelector(".qty").textContent = label.qty || "";
      node.querySelector(".desc").textContent = label.description || "";
      node.querySelector(".location").textContent = label.location || "";

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
    const raw = $("bulkText")?.value.trim();

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
      let branch = parts[2] || "";
      let item = parts[3] || "";
      let qty = parts[4] || "";
      let description = parts[5] || "";
      let location = parts[6] || "";

      docType = docType.toUpperCase();

      if (!["PO", "SPO", "SXFR", "XFR"].includes(docType)) {
        location = description;
        description = qty;
        qty = item;
        item = branch;
        branch = "";
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
        branch,
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

    if (clean.startsWith("http")) {
      const url = new URL(clean);
      const scan = url.searchParams.get("scan");
      if (!scan) throw new Error("No scan data in URL.");
      return JSON.parse(decodeURIComponent(scan));
    }

    const data = JSON.parse(clean);

    if (data.type !== "line_check" || !data.lineId) {
      throw new Error("This QR is not a line check label.");
    }

    return data;
  }

  function markChecked(data) {
    const worker = $("workerName")?.value.trim() || "Unknown Worker";

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
      branch: data.branch || "",
      item: data.item || "",
      qty: data.qty || "",
      description: data.description || "",
      location: data.location || ""
    });

    save();
    renderLog();

    if ($("scanInput")) $("scanInput").value = "";
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
        <td>${escapeHtml(formatDocDisplay(row))}</td>
        <td>${escapeHtml(row.branch || "")}</td>
        <td>${escapeHtml(row.item)}</td>
        <td>${escapeHtml(row.qty)}</td>
        <td>${escapeHtml(row.description)}</td>
        <td>${escapeHtml(row.location || "")}</td>
        <td>${escapeHtml(row.lineId)}</td>
      `;

      body.appendChild(tr);
    });

    const today = todayKey();
    const worker = $("workerName")?.value.trim();

    const todayRows = state.checks.filter((row) => row.date === today);
    const workerRows = worker
      ? todayRows.filter((row) => row.worker.toLowerCase() === worker.toLowerCase())
      : [];

    if ($("todayTotal")) $("todayTotal").textContent = todayRows.length;
    if ($("workerTotal")) $("workerTotal").textContent = worker ? workerRows.length : 0;
  }

  function exportCsv() {
    const headers = [
      "Checked At",
      "Date",
      "Worker",
      "Type",
      "Document Number",
      "Branch",
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
      formatDocDisplay(r),
      r.branch || "",
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

  $("addLabelBtn")?.addEventListener("click", () => {
    addLabel({
      docType: $("docType")?.value || "PO",
      docNumber: $("poNumber")?.value.trim() || "",
      branch: $("branch")?.value.trim() || "",
      item: $("itemNumber")?.value.trim() || "",
      qty: $("quantity")?.value.trim() || "",
      description: $("description")?.value.trim() || "",
      location: $("location")?.value.trim() || "",
      copies: $("labelCopies")?.value || 1
    });
  });

  $("bulkAddBtn")?.addEventListener("click", addBulkLabels);

  $("checkLineBtn")?.addEventListener("click", () => {
    try {
      markChecked(parseScanData($("scanInput").value));
    } catch (err) {
      console.error(err);
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

      if (!importedRows.length) {
        alert("No PO / SPO / SXFR / XFR lines found. The file opened, but item/qty/document columns were not detected.");
      }
    } catch (err) {
      console.error(err);
      alert("Could not read this file. Make sure it is CSV, XLS, or XLSX.");
    }
  });

  $("importPoBtn")?.addEventListener("click", async () => {
    try {
      if (!importedRows.length) importedRows = await readUploadedFile();

      if (!importedRows.length) {
        alert("No PO / SPO / SXFR / XFR lines found.");
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
          branch: row.branch || "",
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
      if (summary) summary.textContent = `Imported ${added} label(s). Skipped ${skipped}.`;

      alert(`Imported ${added} label(s). Skipped ${skipped}.`);
    } catch (err) {
      console.error(err);
      alert("Import failed. Check the file columns and try again.");
    }
  });

  handleUrlScan();
  renderLabels();
  renderLog();
});
