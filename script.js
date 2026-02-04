const CONFIG = {
  cols: 26, // A-Z
  rows: 100,
  defaultColWidth: 100,
  defaultRowHeight: 22,
  headerHeight: 25,
  rowHeaderWidth: 50,
};

const state = {
  sheets: {
    Sheet1: {
      data: {
        "0-0": "二维码名称",
        "0-1": "二维码内容",
      },
      // We can store column widths here if we wanted interactive resizing
    },
    Sheet2: {
      data: {
        "0-0": "二维码名称",
        "0-1": "二维码内容",
      },
    },
  },
  activeSheet: "Sheet1",
  selection: { r: 0, c: 0 },
};

// DOM Elements
const container = document.getElementById("spreadsheet-container");
const nameBox = document.getElementById("name-box");
const formulaInput = document.getElementById("formula-input");
const sheetTabsList = document.getElementById("sheet-tabs-list");

// Initialize
// Initialize
function init() {
  renderSheetTabs();
  setupGrid();

  // Default load
  loadSheet(state.activeSheet);
  selectCell(0, 0);

  // Try to load fy.json
  fetch("fy.json")
    .then((response) => {
      if (!response.ok) throw new Error("Network response was not ok");
      return response.json();
    })
    .then((json) => {
      const keys = Object.keys(json);
      if (keys.length > 0) {
        // Apply default headers
        Object.values(json).forEach((sheet) => {
          if (!sheet.data) sheet.data = {};
          sheet.data["0-0"] = "二维码名称";
          sheet.data["0-1"] = "二维码内容";
        });
        state.sheets = json;
        state.activeSheet = keys[0];
        renderSheetTabs();
        loadSheet(state.activeSheet);
        selectCell(0, 0);
      }
    })
    .catch((err) => {
      console.log(
        "Could not load fy.json (likely due to local file restrictions or missing file), keeping default state.",
        err,
      );
    });

  // Event Listeners
  setupEvents();
}

function setupGrid() {
  // 1. Set CSS Grid Template
  // We adjust the first two columns to be wider for the QR content
  // But CSS repeat is rigid. Let's make it inline styles or just uniform for simplicity first,
  // maybe 150px for all or custom.
  // Let's create the DOM elements first.

  container.innerHTML = "";

  // 0. Corner Header
  const corner = document.createElement("div");
  corner.className = "corner-header";
  corner.style.gridColumn = "1 / 2";
  corner.style.gridRow = "1 / 2";
  container.appendChild(corner);

  // 1. Column Headers (A-Z)
  for (let c = 0; c < CONFIG.cols; c++) {
    const char = String.fromCharCode(65 + c);
    const header = document.createElement("div");
    header.className = "header-col";
    header.textContent = char;
    header.dataset.col = c;
    header.style.gridColumn = `${c + 2} / ${c + 3}`;
    header.style.gridRow = "1 / 2";
    container.appendChild(header);
  }

  // 2. Row Headers (1-100)
  for (let r = 0; r < CONFIG.rows; r++) {
    const header = document.createElement("div");
    header.className = "header-row";
    header.textContent = r + 1;
    header.dataset.row = r;
    header.style.gridColumn = "1 / 2";
    header.style.gridRow = `${r + 2} / ${r + 3}`;
    container.appendChild(header);
  }

  // 3. Cells
  // We create them once. We verify performance. 2600 divs is okay.
  // If we wanted virtual scrolling we'd do that, but for this size it's fine.
  for (let r = 0; r < CONFIG.rows; r++) {
    for (let c = 0; c < CONFIG.cols; c++) {
      const cell = document.createElement("div");
      cell.className = "cell";
      cell.contentEditable = true;
      cell.dataset.r = r;
      cell.dataset.c = c;
      cell.style.gridColumn = `${c + 2} / ${c + 3}`;
      cell.style.gridRow = `${r + 2} / ${r + 3}`;
      // Specific styling for QR columns in headers? No, data.
      container.appendChild(cell);
    }
  }

  // 4. Selection Overlay
  const selectionBox = document.createElement("div");
  selectionBox.className = "selection-box";
  selectionBox.id = "selection-box";
  selectionBox.innerHTML = '<div class="selection-handle"></div>';
  container.appendChild(selectionBox);
}

function loadSheet(sheetName) {
  state.activeSheet = sheetName;
  const data = state.sheets[sheetName].data;

  // Update all cells
  const cells = container.querySelectorAll(".cell");
  cells.forEach((cell) => {
    const r = cell.dataset.r;
    const c = cell.dataset.c;
    const key = `${r}-${c}`;
    cell.textContent = data[key] || "";
  });

  updateTabsUI();
}

function renderSheetTabs() {
  sheetTabsList.innerHTML = "";
  Object.keys(state.sheets).forEach((name) => {
    const tab = document.createElement("div");
    tab.className = `sheet-tab ${name === state.activeSheet ? "active" : ""}`;

    // Wrap text in span to easily replace with input later or just target text
    const span = document.createElement("span");
    span.textContent = name;
    tab.appendChild(span);

    tab.onclick = (e) => {
      // If editing, don't trigger sheet load
      if (e.target.tagName === "INPUT") return;

      loadSheet(name);
      selectCell(0, 0);
    };

    tab.ondblclick = () => {
      startRename(tab, name);
    };

    sheetTabsList.appendChild(tab);
  });
}

function startRename(tabElement, oldName) {
  const input = document.createElement("input");
  input.type = "text";
  input.value = oldName;
  // Styling to match tab
  input.style.width = "100px";
  input.style.fontFamily = "inherit";
  input.style.fontSize = "inherit";
  input.style.border = "none";
  input.style.outline = "2px solid #217346";
  input.style.padding = "2px";

  tabElement.innerHTML = "";
  tabElement.appendChild(input);
  input.focus();
  input.select();

  let committed = false;
  const commit = () => {
    if (committed) return;
    committed = true;

    const newName = input.value.trim();
    if (newName && newName !== oldName) {
      // Check for duplicates
      if (state.sheets[newName]) {
        alert("Sheet name already exists");
        renderSheetTabs(); // Revert
        return;
      }
      performRename(oldName, newName);
    } else {
      renderSheetTabs(); // Revert
    }
  };

  input.addEventListener("blur", commit);
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      commit();
    } else if (e.key === "Escape") {
      committed = true;
      renderSheetTabs();
    }
  });
}

function performRename(oldName, newName) {
  const newSheets = {};
  // Preserve order by iterating old keys
  Object.keys(state.sheets).forEach((key) => {
    if (key === oldName) {
      newSheets[newName] = state.sheets[oldName];
    } else {
      newSheets[key] = state.sheets[key];
    }
  });

  state.sheets = newSheets;

  // Update active sheet reference if needed
  if (state.activeSheet === oldName) {
    state.activeSheet = newName;
  }

  renderSheetTabs();
}

function updateTabsUI() {
  const tabs = sheetTabsList.querySelectorAll(".sheet-tab");
  tabs.forEach((tab) => {
    if (tab.textContent === state.activeSheet) {
      tab.classList.add("active");
    } else {
      tab.classList.remove("active");
    }
  });
}

function selectCell(r, c) {
  r = parseInt(r);
  c = parseInt(c);
  state.selection = { r, c };

  // Update variables
  const cell = getCellEl(r, c);
  if (!cell) return;

  // 1. Move Selection Box
  const box = document.getElementById("selection-box");

  // We can use offsetLeft/Top but only if layout is stable.
  // Since it matches the grid lines, we can use the same rendering logic?
  // Actually box is absolute.
  // The cell position:
  box.style.left = cell.offsetLeft + "px";
  box.style.top = cell.offsetTop + "px";
  box.style.width = cell.offsetWidth + "px";
  box.style.height = cell.offsetHeight + "px";

  // 2. Update Name Box
  const colName = String.fromCharCode(65 + c);
  nameBox.textContent = `${colName}${r + 1}`;

  // 3. Update Formula Bar
  const key = `${r}-${c}`;
  formulaInput.textContent = state.sheets[state.activeSheet].data[key] || "";

  // 4. Highlight Headers
  document
    .querySelectorAll(".header-col")
    .forEach((h) => h.classList.remove("active"));
  document
    .querySelectorAll(".header-row")
    .forEach((h) => h.classList.remove("active"));

  const colHeader = container.querySelector(`.header-col[data-col="${c}"]`);
  const rowHeader = container.querySelector(`.header-row[data-row="${r}"]`);
  if (colHeader) colHeader.classList.add("active");
  if (rowHeader) rowHeader.classList.add("active");
}

function getCellEl(r, c) {
  // A bit inefficient selector, but fine for prototype
  return container.querySelector(`.cell[data-r="${r}"][data-c="${c}"]`);
}

function setupEvents() {
  // Grid Clicks
  container.addEventListener("mousedown", (e) => {
    const cell = e.target.closest(".cell");
    if (cell) {
      selectCell(cell.dataset.r, cell.dataset.c);
    }
  });

  // Content Edits
  container.addEventListener("input", (e) => {
    if (e.target.classList.contains("cell")) {
      const r = e.target.dataset.r;
      const c = e.target.dataset.c;
      const val = e.target.textContent;

      // Save to state
      state.sheets[state.activeSheet].data[`${r}-${c}`] = val;

      // Update formula bar if active
      if (state.selection.r == r && state.selection.c == c) {
        formulaInput.textContent = val;
      }
    }
  });

  // Formula Bar Edits
  formulaInput.addEventListener("input", () => {
    const val = formulaInput.textContent;
    const { r, c } = state.selection;
    const cell = getCellEl(r, c);
    if (cell) {
      cell.textContent = val;
      state.sheets[state.activeSheet].data[`${r}-${c}`] = val;
    }
  });

  // New Sheet
  document.getElementById("new-sheet-btn").addEventListener("click", () => {
    const count = Object.keys(state.sheets).length + 1;
    const newName = `Sheet${count}`;
    state.sheets[newName] = {
      data: {
        "0-0": "二维码名称",
        "0-1": "二维码内容",
      },
    };
    renderSheetTabs();
    loadSheet(newName);
  });

  // Export Data (<)
  document.getElementById("btn-export").addEventListener("click", () => {
    // Clone state to filter out header row (0-x) and sort keys
    const sheetsToExport = JSON.parse(JSON.stringify(state.sheets));
    Object.values(sheetsToExport).forEach((sheet) => {
      if (sheet.data) {
        const sortedData = {};
        Object.keys(sheet.data)
          .filter((key) => {
            // Filter out headers (row 0)
            const [r] = key.split("-").map(Number);
            return r !== 0;
          })
          .sort((a, b) => {
            // Sort by Row then Column
            const [r1, c1] = a.split("-").map(Number);
            const [r2, c2] = b.split("-").map(Number);
            if (r1 !== r2) return r1 - r2;
            return c1 - c2;
          })
          .forEach((key) => {
            sortedData[key] = sheet.data[key];
          });
        sheet.data = sortedData;
      }
    });

    const dataStr =
      "data:text/json;charset=utf-8," +
      encodeURIComponent(JSON.stringify(sheetsToExport, null, 2));
    const downloadAnchorNode = document.createElement("a");
    downloadAnchorNode.setAttribute("href", dataStr);
    downloadAnchorNode.setAttribute("download", "excel_data.json");
    document.body.appendChild(downloadAnchorNode);
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
  });

  // Import Data Trigger (>)
  document.getElementById("btn-import").addEventListener("click", () => {
    document.getElementById("import-file").click();
  });

  // Handle Import File
  document.getElementById("import-file").addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const jsonObj = JSON.parse(event.target.result);
        // Basic validation
        const keys = Object.keys(jsonObj);
        if (keys.length > 0) {
          // Enforce default headers for all imported sheets
          Object.values(jsonObj).forEach((sheet) => {
            if (!sheet.data) sheet.data = {};
            sheet.data["0-0"] = "二维码名称";
            sheet.data["0-1"] = "二维码内容";
          });

          state.sheets = jsonObj;
          state.activeSheet = keys[0];
          renderSheetTabs();
          loadSheet(state.activeSheet);
          selectCell(0, 0);
        } else {
          alert("Invalid data format or empty file");
        }
      } catch (err) {
        alert("Error parsing JSON: " + err.message);
      }
    };
    reader.readAsText(file);
    e.target.value = ""; // Reset
  });

  // ===== QR Selector Modal =====
  setupQRSelector();
}

// QR Selector Module
function setupQRSelector() {
  const qrModalOverlay = document.getElementById("qr-modal-overlay");
  const qrModalClose = document.getElementById("qr-modal-close");
  const qrMainNav = document.getElementById("qr-main-nav");
  const qrSubList = document.getElementById("qr-sub-list");
  const qrTarget = document.getElementById("qr-target");
  const qrCardTitle = document.getElementById("qr-card-title");
  const qrCardSubtitle = document.getElementById("qr-card-subtitle");
  const qrZoomBtn = document.getElementById("qr-zoom-btn");
  const qrFullscreenOverlay = document.getElementById("qr-fullscreen-overlay");
  const qrFullscreenClose = document.getElementById("qr-fullscreen-close");
  const qrFullscreenTarget = document.getElementById("qr-fullscreen-target");
  const qrFullscreenTitle = document.getElementById("qr-fullscreen-title");
  const qrFullscreenSubtitle = document.getElementById(
    "qr-fullscreen-subtitle",
  );

  let qrCurrentCategory = "";
  let qrCurrentItem = null;

  // Open modal when fx is clicked
  document.getElementById("fx-btn").addEventListener("click", () => {
    openQRModal();
  });

  // Close modal
  qrModalClose.addEventListener("click", closeQRModal);
  qrModalOverlay.addEventListener("click", (e) => {
    if (e.target === qrModalOverlay) closeQRModal();
  });

  // Fullscreen handlers
  qrZoomBtn.addEventListener("click", openQRFullscreen);
  qrFullscreenClose.addEventListener("click", closeQRFullscreen);
  qrFullscreenOverlay.addEventListener("click", (e) => {
    if (e.target === qrFullscreenOverlay) closeQRFullscreen();
  });

  // ESC key handling
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      if (qrFullscreenOverlay.classList.contains("active")) {
        closeQRFullscreen();
      } else if (qrModalOverlay.classList.contains("active")) {
        closeQRModal();
      }
    }
  });

  function openQRModal() {
    qrModalOverlay.classList.add("active");
    renderQRMainNav();
    // Select first category
    const categories = Object.keys(state.sheets);
    if (categories.length > 0) {
      selectQRCategory(categories[0]);
    }
  }

  function closeQRModal() {
    qrModalOverlay.classList.remove("active");
  }

  function renderQRMainNav() {
    qrMainNav.innerHTML = "";
    Object.keys(state.sheets).forEach((sheetName) => {
      const navItem = document.createElement("div");
      navItem.className = `qr-nav-item ${sheetName === qrCurrentCategory ? "active" : ""}`;
      navItem.onclick = () => selectQRCategory(sheetName);
      navItem.innerHTML = `<div class="qr-nav-circle"></div><div class="qr-nav-text">${sheetName}</div>`;
      qrMainNav.appendChild(navItem);
    });
  }

  function selectQRCategory(category) {
    qrCurrentCategory = category;
    // Update nav active state
    qrMainNav.querySelectorAll(".qr-nav-item").forEach((item) => {
      const text = item.querySelector(".qr-nav-text").innerText;
      item.classList.toggle("active", text === category);
    });
    renderQRSubList(category);
  }

  function getItemsFromSheetData(sheetData) {
    if (!sheetData) return [];
    const rows = {};
    Object.keys(sheetData).forEach((key) => {
      const parts = key.split("-");
      if (parts.length < 2) return;
      const r = parseInt(parts[0]);
      const c = parseInt(parts[1]);
      // Skip header row (row 0)
      if (r === 0) return;
      if (!rows[r]) rows[r] = {};
      if (c === 0) rows[r].name = sheetData[key];
      if (c === 1) rows[r].content = sheetData[key];
    });
    return Object.keys(rows)
      .sort((a, b) => parseInt(a) - parseInt(b))
      .map((r) => rows[r])
      .filter((item) => item.name || item.content);
  }

  function renderQRSubList(category) {
    qrSubList.innerHTML = "";
    const sheet = state.sheets[category];
    const items = sheet && sheet.data ? getItemsFromSheetData(sheet.data) : [];

    // Get displayField from sheet config, default to "name"
    // "name" = 显示二维码名称 (column 0)
    // "content" = 显示二维码内容 (column 1)
    const displayField =
      sheet && sheet.displayField ? sheet.displayField : "name";

    if (items.length > 0) {
      items.forEach((item, index) => {
        const label = document.createElement("label");
        label.className = "qr-radio-item";
        const input = document.createElement("input");
        input.type = "radio";
        input.name = "qr-device";
        if (index === 0) input.checked = true;
        input.onchange = () => updateQRCard(item);
        const spanRadio = document.createElement("span");
        spanRadio.className = "qr-custom-radio";
        const spanText = document.createElement("span");
        // Display based on displayField setting
        if (displayField === "content") {
          spanText.innerText = item.content || "(无内容)";
        } else {
          spanText.innerText = item.name || "(无名称)";
        }
        label.appendChild(input);
        label.appendChild(spanRadio);
        label.appendChild(spanText);
        qrSubList.appendChild(label);
        if (index === 0) updateQRCard(item);
      });
    } else {
      qrCurrentItem = null;
      qrCardTitle.innerText = "暂无数据";
      qrCardSubtitle.innerText = "-";
      qrTarget.innerHTML = "";
    }
  }

  function updateQRCard(item) {
    qrCurrentItem = item;
    qrCardTitle.innerText = item.name || "(无名称)";
    qrCardSubtitle.innerText = item.content || "-";
    qrTarget.innerHTML = "";
    if (item.content) {
      new QRCode(qrTarget, {
        text: item.content,
        width: 180,
        height: 180,
        colorDark: "#000000",
        colorLight: "#ffffff",
        correctLevel: QRCode.CorrectLevel.H,
      });
    }
  }

  function openQRFullscreen() {
    if (!qrCurrentItem || !qrCurrentItem.content) return;
    qrFullscreenOverlay.classList.add("active");
    qrFullscreenTitle.innerText = qrCurrentItem.name || "(无名称)";
    qrFullscreenSubtitle.innerText = qrCurrentItem.content || "-";
    qrFullscreenTarget.innerHTML = "";
    new QRCode(qrFullscreenTarget, {
      text: qrCurrentItem.content,
      width: 400,
      height: 400,
      colorDark: "#000000",
      colorLight: "#ffffff",
      correctLevel: QRCode.CorrectLevel.H,
    });
  }

  function closeQRFullscreen() {
    qrFullscreenOverlay.classList.remove("active");
    setTimeout(() => {
      qrFullscreenTarget.innerHTML = "";
    }, 300);
  }
}

init();
