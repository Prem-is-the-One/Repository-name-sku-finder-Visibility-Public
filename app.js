const DATA_URL = "data/data.json";
const INITIAL_RESULTS = 100;
const MAX_SEARCH_RESULTS = 250;

let skuData = [];

const searchInput = document.getElementById("searchInput");
const clearButton = document.getElementById("clearButton");
const results = document.getElementById("results");
const recordCount = document.getElementById("recordCount");
const statusText = document.getElementById("statusText");
const lastUpdated = document.getElementById("lastUpdated");

function normalize(value) {
  return String(value ?? "").toLowerCase().trim();
}

function escapeHTML(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function validURL(value) {
  const url = String(value ?? "").trim();

  if (!url) return "";
  if (url.startsWith("http://") || url.startsWith("https://")) return url;

  return `https://${url}`;
}

function setClearButtonState() {
  clearButton.disabled = searchInput.value.trim() === "";
}

function emptyState(title, message, symbol = "?") {
  results.innerHTML = `
    <div class="empty">
      <div class="empty-inner">
        <div class="empty-icon" aria-hidden="true">${escapeHTML(symbol)}</div>
        <strong>${escapeHTML(title)}</strong>
        <p>${escapeHTML(message)}</p>
      </div>
    </div>
  `;
}

function renderResults(rows, query = "") {
  results.innerHTML = "";
  results.setAttribute("aria-busy", "false");

  if (rows.length === 0) {
    emptyState(
      "No matching SKU found",
      query
        ? `Try checking the SKU number or searching with fewer digits. Search: ${query}`
        : "No SKU data is currently available.",
      "×"
    );

    statusText.textContent = query
      ? `No result found for “${query}”`
      : "No SKU data found.";

    return;
  }

  const table = document.createElement("div");
  table.className = "sku-table";

  table.innerHTML = `
    <div class="sku-table-header">
      <div>SKU Number</div>
      <div>Product Document</div>
    </div>
  `;

  rows.forEach((row, index) => {
    const sku = row["SKU Number"] ?? "";
    const rawURL = row["Product URL"] ?? "";
    const url = validURL(rawURL);

    const rowElement = document.createElement("div");
    rowElement.className = "sku-row";

    const buttonHTML = url
      ? `
        <a
          href="${escapeHTML(url)}"
          target="_blank"
          rel="noopener noreferrer"
          class="view-button"
          aria-label="Open product document for SKU ${escapeHTML(sku)}"
        >
          Open Product
          <svg viewBox="0 0 24 24" aria-hidden="true">
            <path d="M7 17 17 7M9 7h8v8"/>
          </svg>
        </a>
      `
      : `<span class="no-url">Product link unavailable</span>`;

    rowElement.innerHTML = `
      <div class="sku-number-wrap">
        <span class="sku-index" aria-hidden="true">
          ${String(index + 1).padStart(2, "0")}
        </span>

        <div class="sku-number">${escapeHTML(sku)}</div>
      </div>

      <div class="sku-action">
        ${buttonHTML}
      </div>
    `;

    table.appendChild(rowElement);
  });

  results.appendChild(table);

  if (query) {
    statusText.textContent =
      `${rows.length.toLocaleString()} result${rows.length === 1 ? "" : "s"} found`;
  } else {
    statusText.textContent =
      `Showing the first ${rows.length.toLocaleString()} SKU${rows.length === 1 ? "" : "s"}`;
  }
}

function searchSKU() {
  const query = normalize(searchInput.value);

  setClearButtonState();

  if (!query) {
    renderResults(skuData.slice(0, INITIAL_RESULTS));
    return;
  }

  const matched = skuData.filter((row) => {
    const sku = normalize(row["SKU Number"]);
    return sku.includes(query);
  });

  renderResults(
    matched.slice(0, MAX_SEARCH_RESULTS),
    searchInput.value.trim()
  );
}

function formatUpdatedDate(value) {
  const updated = new Date(value);

  if (Number.isNaN(updated.getTime())) {
    return "";
  }

  return updated.toLocaleString(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit"
  });
}

async function loadData() {
  try {
    results.setAttribute("aria-busy", "true");

    const response = await fetch(
      `${DATA_URL}?v=${Date.now()}`,
      {
        cache: "no-store"
      }
    );

    if (!response.ok) {
      throw new Error(`HTTP Error ${response.status}`);
    }

    const payload = await response.json();

    skuData = Array.isArray(payload)
      ? payload
      : payload.rows || [];

    skuData = skuData.filter((row) => {
      return String(row["SKU Number"] ?? "").trim() !== "";
    });

    recordCount.textContent =
      skuData.length.toLocaleString();

    if (payload.meta?.generated_at) {
      const formatted =
        formatUpdatedDate(payload.meta.generated_at);

      lastUpdated.textContent =
        formatted
          ? `Last updated ${formatted}`
          : "";
    }

    renderResults(
      skuData.slice(0, INITIAL_RESULTS)
    );

  } catch (error) {
    console.error(error);

    recordCount.textContent = "—";
    statusText.textContent =
      "Could not load Excel data.";

    results.setAttribute(
      "aria-busy",
      "false"
    );

    emptyState(
      "Database could not be loaded",
      "Run tools/export_excel.py and make sure data/data.json exists before opening the website.",
      "!"
    );
  }
}

searchInput.addEventListener(
  "input",
  searchSKU
);

clearButton.addEventListener(
  "click",
  () => {
    searchInput.value = "";
    searchInput.focus();
    searchSKU();
  }
);

document.addEventListener(
  "keydown",
  (event) => {

    if (
      event.key === "Escape" &&
      document.activeElement === searchInput
    ) {
      searchInput.value = "";
      searchSKU();
    }

    if (
      event.key === "/" &&
      document.activeElement !== searchInput
    ) {
      event.preventDefault();
      searchInput.focus();
      searchInput.select();
    }
  }
);

setClearButtonState();

loadData();