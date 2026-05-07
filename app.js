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
    return String(value ?? "")
      .replace(/\$/g, "")
      .replace(/\s+/g, " ")
      .trim();
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

  function isDocType(value) {
    return ["PO", "SPO", "SXFR", "XFR"].includes(cleanText(value).toUpperCase());
  }

  function normalizeDocNumber(value, defaultType = "PO") {
    let text = cleanText(value).toUpperCase();
    text = text.replace(/P\.O\./g, "PO");
    text = text.replace(/[^A-Z0-9-]/g, "");

    if (!text) return "";
    if (/^(PO|SPO|SXFR|XFR)/.test(text)) return text;

    return `${defaultType}${text}`;
  }

  function formatDocDisplay(label) {
    const type = cleanText(label.docType).toUpperCase();
    const number = cleanText(label.docNumber).toUpperCase();

    if (!number && type) return type;
    if (!type) return number;
    if (number.startsWith(type)) return number;

    return `${type}${number}`;
  }

  function labelPayload(label) {
    const item = encodeURIComponent(cleanText(label.item));
    return `https://www.chadwellsupply.com/search/?q=${item}`;
  }

  function looksLikeItemNumber(value) {
    return /^[0-9]{5,8}$/.test(cleanText(value));
  }

  function looksLikeLocation(value) {
    const text = cleanText(value).toUpperCase();
    return /^[A-Z]{1,4}-[0-9]{1,3}-[0-9]{1,3}(-[0-9]{1,3})?$/.test(text);
  }

  function rowIsBlank(row) {
    return row.map(cleanText).every((cell) => !cell);
  }

  function extractTopDocumentMeta(rows) {
    const fallbackType = $("defaultDocType")?.value || "SPO";
    let docType = fallbackType;
    let docNumber = cleanText($("defaultDocNumber")?.value || "");
    let branch = cleanText($("defaultBranch")?.value || "");

    const topText = rows
      .slice(0, 25)
      .map((row) => row.map(cleanText).join(" "))
      .join(" ")
      .toUpperCase();

    const fullDocMatch = topText.match(/\b(SPO|SXFR|XFR|PO)[\s#:\-]*([0-9]{4,12})\b/i);

    if (!docNumber && fullDocMatch) {
      docType = fullDocMatch[1].toUpperCase();
      docNumber = `${docType}${fullDocMatch[2]}`;
    }

    const branchMatch = topText.match(/\bBRANCH\s*[-:#]?\s*([A-Z]{2}\d{2})\b/i);

    if (!branch && branchMatch) {
      branch = branchMatch[1].toUpperCase();
    }

    docNumber = normalizeDocNumber(docNumber, docType);

    return { docType, docNumber, branch };
  }

  function detectReceivingReportRows(rows, meta) {
    const fixedRows = [];

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i] || [];
      const nextRow = rows[i + 1] || [];

      if (rowIsBlank(row)) continue;

      const lineNumber = cleanText(row[0]);
      const itemNumber = cleanText(row[4]);      // Column E = ITEM #
      const description = cleanText(nextRow[0]); // Next row Column A = DESCRIPTION
      const location = cleanText(nextRow[6]);    // Next row Column G = BIN
      const qty = cleanText(nextRow[12]);        // Next row Column M = QTY ORD

      const isRealReceivingLine =
        /^[0-9]{1,4}$/.test(lineNumber) &&
        looksLikeItemNumber(itemNumber);

      if (!isRealReceivingLine) continue;

      fixedRows.push({
        docType: meta.docType || $("defaultDocType")?.value || "SPO",
        docNumber: meta.docNumber || $("defaultDocNumber")?.value || "",
        branch: meta.branch || $("defaultBranch")?.value || "",
        item: itemNumber,
        qty: qty ? String(Number(qty)) : "",
        description,
        location
      });

      i++;
    }

    return fixedRows;
  }

  function findHeaderRow(rows) {
    let bestIndex = 0;
    let bestScore = -1;

    rows.forEach((row, index) => {
      if (index > 70) return;

      const keys = row.map(cleanKey);
      let score = 0;

      if (keys.includes("item") || keys.includes("itemnumber") || keys.includes("itemno")) score += 10;
      if (keys.includes("description") || keys.includes("desc") || keys.includes("itemdescription")) score += 10;
      if (keys.includes("qty") || keys.includes("quantity") || keys.includes("qtyord")) score += 5;
      if (keys.includes("bin") || keys.includes("location") || keys.includes("loc")) score += 5;

      if (score > bestScore) {
        bestScore = score;
        bestIndex = index;
      }
    });

    return bestIndex;
  }

  function mapHeaders(headers) {
    const map = {};

    headers.forEach((header, index) => {
      const key = cleanKey(header);

      if (["item", "itemnumber", "itemno", "itemnum", "itemid", "sku"].includes(key)) {
        map.item = index;
      }

      if (["description", "desc", "itemdescription", "productdescription", "itemdesc"].includes(key)) {
        map.description = index;
      }

      if (["qty", "quantity", "qtyord", "qtyordered", "orderedqty", "qtyreceived", "receivedqty"].includes(key)) {
        map.qty = index;
      }

      if (["bin", "location", "loc", "binlocation", "putawaylocation"].includes(key)) {
        map.location = index;
      }

      if (["branch", "warehouse", "site", "facility"].includes(key)) {
        map.branch = index;
      }

      if (["po", "ponumber", "spo", "sponumber", "sxfr", "sxfrnumber", "xfr", "xfrnumber", "documentnumber", "docnumber"].includes(key)) {
        map.docNumber = index;
      }

      if (["type", "doctype", "documenttype"].includes(key)) {
        map.docType = index;
      }
    });

    return map;
  }

  function getCell(row, index) {
    if (index === undefined || index === null) return "";
    return cleanText(row[index]);
  }

  function readCleanTableRows(rows, meta) {
    const headerIndex = findHeaderRow(rows);
    const headers = rows[headerIndex] || [];
    const map = mapHeaders(headers);
    const dataRows = rows.slice(headerIndex + 1);

    return dataRows
      .filter((row) => !rowIsBlank(row))
      .map((row) => {
        let docType = getCell(row, map.docType) || meta.docType || $("defaultDocType")?.value || "SPO";
        docType = cleanText(docType).toUpperCase();

        if (!isDocType(docType)) docType = meta.docType || "SPO";

        const docNumber = normalizeDocNumber(
          getCell(row, map.docNumber) || meta.docNumber || $("defaultDocNumber")?.value || "",
          docType
        );

        const branch = getCell(row, map.branch) || meta.branch || $("defaultBranch")?.value || "";
        const item = getCell(row, map.item);
        const qty = getCell(row, map.qty);
        const description = getCell(row, map.description);
        const location = getCell(row, map.location);

        return {
          docType,
          docNumber,
          branch,
          item,
          qty,
          description,
          location
        };
      })
      .filter((row) => looksLikeItemNumber(row.item));
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

    const meta = extractTopDocumentMeta(rows);

    const receivingRows = detectReceivingReportRows(rows, meta);

    if (receivingRows.length) {
      return receivingRows;
    }

    return readCleanTableRows(rows, meta);
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

      node.querySelector(".docNumber").textContent = formatDocDisplay(label);
      node.querySelector(".branch").textContent = label.branch || "";
      const typeText = node.querySelector(".docTypeText");
if (typeText) typeText.textContent = label.docType || "";
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
      const parts = line.split(",").map((x) => cleanText(x));

      let docType = parts[0] || $("docType").value || "SPO";
      let docNumber = parts[1] || "";
      let branch = parts[2] || "";
      let item = parts[3] || "";
      let qty = parts[4] || "";
      let description = parts[5] || "";
      let location = parts[6] || "";

      docType = docType.toUpperCase();

      if (!isDocType(docType)) {
        location = description;
        description = qty;
        qty = item;
        item = branch;
        branch = "";
        docNumber = docType;
        docType = $("docType").value || "SPO";
      }

      docNumber = normalizeDocNumber(docNumber, docType);

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
      window.open(clean, "_blank");
      throw new Error("Opened item search.");
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
    const docType = $("docType")?.value || "SPO";

    addLabel({
      docType,
      docNumber: normalizeDocNumber($("poNumber")?.value.trim() || "", docType),
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
      if (err.message !== "Opened item search.") {
        alert("QR data was not readable. Scan again.");
      }
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
        alert("No valid receiving lines found. Make sure the file has real item numbers in the ITEM # column.");
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
        alert("No valid PO / SPO / SXFR / XFR lines found.");
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
          String(label.docType).toLowerCase() === String(row.docType).toLowerCase() &&
          String(label.docNumber).toLowerCase() === String(row.docNumber).toLowerCase() &&
          String(label.item).toLowerCase() === String(row.item).toLowerCase() &&
          String(label.location || "").toLowerCase() === String(row.location || "").toLowerCase()
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

  renderLabels();
  renderLog();
});
