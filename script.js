let backgroundImage = "Letter.jpg";
let parsedData = [];
let skippedEntries = [];
let currentBatchIndex = 0;
let totalValidLetters = 0;
let duplicateEmailCount = 0;
let duplicateAddressCount = 0;
let todayString = "";

const seenEmails = new Set();
const seen = new Set();
const batchSize = 100;

seenEmails.clear();
seen.clear();

document.getElementById("bgUpload").addEventListener("change", function (e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (event) {
    backgroundImage = event.target.result;
    console.log("✅ Background image updated!");
  };
  reader.readAsDataURL(file);
});

document.getElementById("fontSize").addEventListener("input", function (e) {
  document.getElementById("fontSizeValue").textContent = e.target.value;
  applyLiveStyles();
});

document
  .getElementById("fontSelect")
  .addEventListener("change", applyLiveStyles);
document.getElementById("lineHeight").addEventListener("input", function (e) {
  document.getElementById("lineHeightValue").textContent = e.target.value;
  applyLiveStyles();
});
document
  .getElementById("textAlign")
  .addEventListener("change", applyLiveStyles);

document.getElementById("dataUpload").addEventListener("change", function (e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (evt) {
    let rows;
    duplicateEmailCount = 0;
    duplicateAddressCount = 0;

    if (file.name.endsWith(".csv")) {
      rows = [];
      const raw = evt.target.result.trim().split("\n");
      const headers = raw[0].split(",").map((h) => h.trim());
      raw.slice(1).forEach((line) => {
        const values = line.split(",");
        const row = Object.fromEntries(
          headers.map((h, i) => [h, values[i]?.trim() || ""])
        );
        rows.push(row);
      });
    } else {
      const workbook = XLSX.read(evt.target.result, { type: "binary" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    }

    const emailSet = new Set();
    const addressSet = new Set();
    parsedData = [];

    rows.forEach((row) => {
      const keys = Object.keys(row).reduce((acc, key) => {
        const k = key.toLowerCase().trim();
        if (k.includes("salutation")) acc.title = row[key];
        if (
          k.includes("contact name") ||
          k.includes("full name") ||
          k === "name"
        )
          acc.name = row[key];
        if (k === "street 1") acc.street1 = row[key];
        if (k === "street 2") acc.street2 = row[key];
        if (k === "street 3") acc.street3 = row[key];
        if (k.includes("postal")) acc.postal = row[key];
        return acc;
      }, {});

      const email =
        row.email ||
        row.Email ||
        row["Email Address"] ||
        Object.values(row).find((_, k) =>
          String(Object.keys(row)[k] || "")
            .toLowerCase()
            .includes("email")
        ) ||
        "";

      const addressKey =
        `${keys.street1}|${keys.street2}|${keys.street3}|${keys.postal}`
          .toLowerCase()
          .trim();
      const emailKey = email.trim().toLowerCase();

      if (emailKey && emailSet.has(emailKey)) {
        duplicateEmailCount++;
        return;
      }
      if (addressSet.has(addressKey)) {
        duplicateAddressCount++;
        return;
      }

      if (emailKey) emailSet.add(emailKey);
      addressSet.add(addressKey);

      parsedData.push({
        title: keys.title || "",
        name: keys.name || "",
        street1: keys.street1 || "",
        street2: keys.street2 || "",
        street3: keys.street3 || "",
        postal: keys.postal || "",
        email,
      });
    });
    if (parsedData.length === 0) {
      alert(
        "⚠️ No valid data found. Please check your file format and required fields."
      );
    } else {
      generate();
    }
  };

  if (file.name.endsWith(".csv")) reader.readAsText(file);
  else reader.readAsBinaryString(file);
});

document
  .getElementById("greetingOffset")
  .addEventListener("input", function (e) {
    document.getElementById("greetingOffsetValue").textContent = e.target.value;
    applyLiveStyles();
  });
document.getElementById("marginLeft").addEventListener("input", function (e) {
  document.getElementById("marginLeftValue").textContent = e.target.value;
  applyLiveStyles();
});
document.getElementById("marginTop").addEventListener("input", function (e) {
  document.getElementById("marginTopValue").textContent = e.target.value;
  applyLiveStyles();
});

function presetMinimum() {
  document.getElementById("fontSize").value = 10;
  document.getElementById("lineHeight").value = 1.0;
  document.getElementById("greetingOffset").value = 0;
  document.getElementById("marginLeft").value = 0;
  document.getElementById("marginTop").value = 0;

  document.getElementById("fontSizeValue").textContent = 10;
  document.getElementById("lineHeightValue").textContent = 1.0;
  document.getElementById("greetingOffsetValue").textContent = 0;
  document.getElementById("marginLeftValue").textContent = 0;
  document.getElementById("marginTopValue").textContent = 0;

  applyLiveStyles();
}

function getFormattedToday() {
  totalValidLetters = 0;
  const now = new Date();
  return `${now.getDate()} ${now.toLocaleString("default", {
    month: "long",
  })} ${now.getFullYear()}`;
}

function renderBatch() {
  const start = currentBatchIndex * batchSize;
  const end = Math.min(start + batchSize, parsedData.length);
  const batch = parsedData.slice(start, end);

  batch.forEach((row, i) => {
    const rowIndex = start + i;
    renderLetter(row, rowIndex);
  });

  currentBatchIndex++;
}

function generate() {
  const output = document.getElementById("output");
  if (!output) {
    alert("❌ 'output' container not found. Please check your HTML.");
    return;
  }
  output.innerHTML = "";

  skippedEntries = [];
  seenEmails.clear();
  seen.clear();
  currentBatchIndex = 0;

  todayString = getFormattedToday();
  if (parsedData.length === 0) {
    alert("⚠️ No data found! Please upload a valid CSV or XLSX file.");
    return;
  }
  renderBatch();
  setTimeout(
    () => output.scrollIntoView({ behavior: "smooth", block: "start" }),
    100
  );

  document.getElementById(
    "processSummary"
  ).innerHTML = `📄 <strong>Total Letters Processed:</strong> ${parsedData.length}<br>
     ✅ <strong>Successfully Generated:</strong> ${totalValidLetters}<br>
     ⛔ <strong>Skipped:</strong> ${skippedEntries.length}
     (${duplicateEmailCount} duplicate emails, ${duplicateAddressCount} duplicate addresses)`;

  renderSkippedTable();
}

function renderLetter(row, i) {
  const { title, name, street1, street2, street3, postal, email } = row;
  const rowNum = i + 2;
  const output = document.getElementById("output");

  if (!name && !street1 && !postal) {
    skippedEntries.push({
      row: rowNum,
      name: "",
      reason: "Blank or missing essential data",
    });
    return;
  }
  if (email?.trim()) {
    const key = email.trim().toLowerCase();
    if (seenEmails.has(key)) {
      skippedEntries.push({ row: rowNum, name, reason: "Duplicate email" });
      duplicateEmailCount++;
      return;
    }
    seenEmails.add(key);
  }
  const addrKey = `${street1}|${street2}|${street3}|${postal}`
    .toLowerCase()
    .trim();
  if (seen.has(addrKey)) {
    skippedEntries.push({ row: rowNum, name, reason: "Duplicate address" });
    duplicateAddressCount++;
    return;
  }
  seen.add(addrKey);

  const missing = [];
  if (!name) missing.push("Contact Name");
  if (!street1) missing.push("Street 1");
  if (!postal) missing.push("Postal Code");
  if (missing.length) {
    skippedEntries.push({
      row: rowNum,
      name,
      reason: `Missing ${missing.join(", ")}`,
    });
    return;
  }

  const fullName = `${title ? title + " " : ""}${name}`;

  const addressLines = [
    todayString,
    "",
    fullName,
    street1,
    street2,
    street3,
    `Singapore ${postal}`,
  ]
    .filter(
      (line, idx) =>
        idx === 1 ||
        (line &&
          String(line).trim().toUpperCase() !== "NIL" &&
          String(line).trim() !== "")
    )
    .map((line, idx) => (idx === 0 || idx === 1 ? line : line.toUpperCase()));
  const addressText = addressLines;

  const container = document.createElement("div");

  const bgImg = document.createElement("img");
  bgImg.src = backgroundImage;
  bgImg.crossOrigin = "anonymous";
  Object.assign(bgImg.style, {
    position: "absolute",
    top: 0,
    left: 0,
    width: "100%",
    height: "100%",
    zIndex: 0,
    objectFit: "cover",
  });
  container.appendChild(bgImg);

  container.className = "page";

  const block = document.createElement("div");

  block.innerHTML = addressLines.join("<br>");

  block.className = "text-block";
  const fontSize = document.getElementById("fontSize").value;
  const lineHeight = document.getElementById("lineHeight").value;
  const textAlign = document.getElementById("textAlign").value;
  const fontFamily = document.getElementById("fontSelect").value;

  Object.assign(block.style, {
    fontFamily,
    fontSize: `${fontSize}px`,
    lineHeight,
    textAlign,
    whiteSpace: "pre-line",
    zIndex: 1,
  });

  container.appendChild(block);

  const greeting = document.createElement("div");
  greeting.className = "greeting";
  greeting.textContent = `Dear ${fullName},`;

  const lines = addressText.length;
  const lineHeightPx = parseFloat(lineHeight) * parseFloat(fontSize);
  const textTop = parseInt(document.getElementById("marginTop").value);
  const addressBottom = textTop + lines * lineHeightPx;
  const greetOff = parseInt(document.getElementById("greetingOffset").value);

  Object.assign(greeting.style, {
    position: "absolute",
    left: `${document.getElementById("marginLeft").value}px`,
    fontFamily,
    fontSize: `${fontSize}px`,
    lineHeight,
    textAlign,
    whiteSpace: "pre-line",
    zIndex: 1,
    top: `${addressBottom + greetOff}px`,
  });

  container.appendChild(greeting);

  const scaled = document.createElement("div");
  scaled.className = "scaled-container";
  const wrapper = document.createElement("div");
  wrapper.className = "preview-wrapper";
  wrapper.appendChild(container);
  scaled.appendChild(wrapper);
  output.appendChild(scaled);

  totalValidLetters++;
}

function renderSkippedTable() {
  const table = document.getElementById("skipped-table");

  if (parsedData.length === 0) {
    table.innerHTML =
      "<p>⚠️ Please upload a contact file before generating.</p>";
    return;
  }

  if (!skippedEntries.length) {
    table.innerHTML = "<p>✅ All contacts processed successfully!</p>";
    return;
  }

  let html = `<h2>⚠️ Skipped Rows</h2><table border="1" cellpadding="6" cellspacing="0">
  <thead><tr><th>Row</th><th>Name</th><th>Reason</th></tr></thead><tbody>`;

  skippedEntries.forEach((entry) => {
    html += `<tr><td>${entry.row}</td><td>${entry.name || "-"}</td><td>${
      entry.reason
    }</td></tr>`;
  });

  html += "</tbody></table>";
  table.innerHTML = html;
}

function createProgressBar(total) {
  const bar = document.createElement("div");
  bar.id = "progress-bar";
  bar.style.position = "fixed";
  bar.style.top = "95px";
  bar.style.left = "50%";
  bar.style.transform = "translateX(-50%)";
  bar.style.width = "300px";
  bar.style.height = "25px";
  bar.style.background = "#ddd";
  bar.style.border = "1px solid #aaa";
  bar.style.zIndex = "9999";
  bar.style.borderRadius = "4px";

  const fill = document.createElement("div");
  fill.id = "progress-fill";
  fill.style.height = "100%";
  fill.style.width = "0%";
  fill.style.background = "#4caf50";
  fill.style.transition = "width 0.2s";
  bar.appendChild(fill);

  document.body.appendChild(bar);
  return bar;
}

function updateProgressBar(bar, current, total) {
  const fill = bar.querySelector("#progress-fill");
  fill.style.width = `${(current / total) * 100}%`;
  fill.textContent = `${current}/${total}`;
}

function applyLiveStyles() {
  const fontSize = parseFloat(document.getElementById("fontSize").value);
  const lineHeight = parseFloat(document.getElementById("lineHeight").value);
  const textAlign = document.getElementById("textAlign").value;
  const greetingOffset = parseInt(
    document.getElementById("greetingOffset").value
  );
  const marginLeft = parseInt(document.getElementById("marginLeft").value);
  const marginTop = parseInt(document.getElementById("marginTop").value);
  const fontFamily = document.getElementById("fontSelect").value;

  document.querySelectorAll(".text-block").forEach((block) => {
    block.style.fontFamily = fontFamily;
    block.style.fontSize = `${fontSize}px`;
    block.style.lineHeight = lineHeight;
    block.style.textAlign = textAlign;
    block.style.left = `${marginLeft}px`;
    block.style.top = `${marginTop}px`;

    const lineCount = block.innerHTML.split("<br>").length;
    const addressBottom = marginTop + lineCount * fontSize * lineHeight;

    const greeting = block.parentElement.querySelector(".greeting");
    if (greeting) {
      greeting.style.fontFamily = fontFamily;
      greeting.style.fontSize = `${fontSize}px`;
      greeting.style.lineHeight = lineHeight;
      greeting.style.textAlign = textAlign;
      greeting.style.top = `${addressBottom + greetingOffset}px`;
      greeting.style.left = `${marginLeft}px`;
    }
  });
}

async function exportAllLettersToPDFInChunks(batchSize = 500) {
  document.querySelectorAll("button").forEach((btn) => (btn.disabled = true));

  if (!parsedData.length) {
    alert("⚠️ Please upload a contact list first.");
    document
      .querySelectorAll("button")
      .forEach((btn) => (btn.disabled = false));
    return;
  }

  let total = parsedData.length;
  let totalPagesSaved = 0;
  let pdfIndex = 1;
  let skippedCount = 0;

  const progressBar = createProgressBar(total);
  let pdf = new jsPDF({ unit: "px", format: [595, 842] });
  let currentBatchCount = 0;

  for (let i = 0; i < parsedData.length; i++) {
    const row = parsedData[i];
    const { name, street1, postal, email, street2, street3, title } = row;

    if (!name && !street1 && !postal) {
      skippedCount++;
      continue;
    }

    if (email?.trim()) {
      const emailKey = email.trim().toLowerCase();
      if (seenEmails.has(emailKey)) {
        skippedCount++;
        duplicateEmailCount++;
        continue;
      }
      seenEmails.add(emailKey);
    }

    const addrKey = `${street1}|${street2}|${street3}|${postal}`
      .toLowerCase()
      .trim();
    if (seen.has(addrKey)) {
      skippedCount++;
      duplicateAddressCount++;
      continue;
    }
    seen.add(addrKey);

    const container = document.createElement("div");
    container.style.width = "595px";
    container.style.height = "842px";
    container.style.position = "relative";
    container.style.fontFamily = document.getElementById("fontSelect").value;
    container.style.fontSize = document.getElementById("fontSize").value + "px";
    container.style.lineHeight = document.getElementById("lineHeight").value;
    container.style.textAlign = document.getElementById("textAlign").value;

    const today = getFormattedToday();
    const fullName = `${title ? title + " " : ""}${name}`;

    const addressLines = [
      todayString,
      "",
      fullName,
      street1,
      street2,
      street3,
      `Singapore ${postal}`,
    ]
      .filter(
        (line, idx) =>
          idx === 1 ||
          (line &&
            String(line).trim().toUpperCase() !== "NIL" &&
            String(line).trim() !== "")
      )
      .map((line, idx) => (idx === 0 || idx === 1 ? line : line.toUpperCase()));

    const bgImg = document.createElement("img");
    bgImg.src = backgroundImage;
    bgImg.crossOrigin = "anonymous";
    Object.assign(bgImg.style, {
      position: "absolute",
      top: 0,
      left: 0,
      width: "100%",
      height: "100%",
      zIndex: 0,
      objectFit: "cover",
    });
    container.appendChild(bgImg);

    const textDiv = document.createElement("div");
    textDiv.className = "text-block";

    const fontSizeValue = parseFloat(document.getElementById("fontSize").value);
    const fontSize = fontSizeValue + "px";
    const lineHeight = document.getElementById("lineHeight").value;
    const fontFamily = document.getElementById("fontSelect").value;
    const textAlign = document.getElementById("textAlign").value;
    const marginTop = document.getElementById("marginTop").value + "px";
    const marginLeft = document.getElementById("marginLeft").value + "px";

    textDiv.style.position = "absolute";
    textDiv.style.top = marginTop;
    textDiv.style.left = marginLeft;
    textDiv.style.whiteSpace = "pre-line";
    textDiv.style.fontFamily = fontFamily;
    textDiv.style.fontSize = fontSize;
    textDiv.style.lineHeight = lineHeight;
    textDiv.style.textAlign = textAlign;
    textDiv.style.zIndex = "1";
    textDiv.innerHTML = addressLines.join("<br>");
    container.appendChild(textDiv);

    const greeting = document.createElement("div");
    greeting.textContent = `Dear ${fullName},`;
    greeting.className = "greeting";

    greeting.style.position = "absolute";
    greeting.style.zIndex = "1";
    greeting.style.fontFamily = fontFamily;
    greeting.style.fontSize = fontSize;
    greeting.style.lineHeight = lineHeight;
    greeting.style.textAlign = textAlign;
    greeting.style.left = marginLeft;

    const lineHeightPx = parseFloat(lineHeight) * fontSizeValue;
    const addressBottom =
      parseInt(marginTop) + addressLines.length * lineHeightPx;
    greeting.style.top = `${
      addressBottom + parseInt(document.getElementById("greetingOffset").value)
    }px`;

    container.appendChild(greeting);

    document.body.appendChild(container);

    await new Promise((resolve) => setTimeout(resolve, 0));

    await new Promise((resolve) => setTimeout(resolve, 0));

    const canvas = await html2canvas(container, {
      scale: 2,
      width: 595,
      height: 842,
      useCORS: true,
    });

    const imageData = canvas.toDataURL("image/jpeg");

    pdf.addImage(imageData, "JPEG", 0, 0, 595, 842);
    document.body.removeChild(container);
    totalPagesSaved++;
    currentBatchCount++;
    updateProgressBar(progressBar, i + 1, total);

    if (currentBatchCount >= batchSize || i === parsedData.length - 1) {
      const filename = `Letters_Part_${pdfIndex}.pdf`;
      pdf.save(filename);
      pdf = new jsPDF({ unit: "px", format: [595, 842] });
      currentBatchCount = 0;
      pdfIndex++;
    }
  }
  document.title = "Mailer Merge Pro";

  document.querySelectorAll("button").forEach((btn) => (btn.disabled = false));

  seenEmails.clear();
  seen.clear();

  progressBar.remove();
}

async function startExport() {
  await exportAllLettersToPDFInChunks();
}

window.addEventListener("DOMContentLoaded", () => {
  presetMinimum();
});
