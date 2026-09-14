const DATA_URL = "data/data.json";

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

    if (!url) {
        return "";
    }

    if (
        url.startsWith("http://") ||
        url.startsWith("https://")
    ) {
        return url;
    }

    return "https://" + url;
}


function renderResults(rows, query = "") {

    results.innerHTML = "";

    if (rows.length === 0) {

        results.innerHTML = `
            <div class="empty">
                No SKU found.
            </div>
        `;

        statusText.textContent = query
            ? `No result found for "${query}"`
            : "No SKU data found.";

        return;
    }


    const table = document.createElement("div");

    table.className = "sku-table";


    table.innerHTML = `
        <div class="sku-table-header">
            <div>SKU Number</div>
            <div>Product</div>
        </div>
    `;


    rows.forEach(row => {

        const sku = row["SKU Number"] ?? "";
        const rawURL = row["Product URL"] ?? "";
        const url = validURL(rawURL);


        const rowElement = document.createElement("div");

        rowElement.className = "sku-row";


        let buttonHTML = `
            <span class="no-url">
                No URL
            </span>
        `;


        if (url) {

            buttonHTML = `
                <a
                    href="${escapeHTML(url)}"
                    target="_blank"
                    rel="noopener noreferrer"
                    class="view-button"
                >
                    View Product
                </a>
            `;
        }


        rowElement.innerHTML = `
            <div class="sku-number">
                ${escapeHTML(sku)}
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
            `Showing ${rows.length.toLocaleString()} SKU${rows.length === 1 ? "" : "s"}`;
    }
}


function searchSKU() {

    const query = normalize(searchInput.value);


    if (!query) {

        renderResults(skuData.slice(0, 100));

        return;
    }


    const matched = skuData.filter(row => {

        const sku = normalize(row["SKU Number"]);

        return sku.includes(query);

    });


    renderResults(
        matched.slice(0, 250),
        searchInput.value.trim()
    );
}


async function loadData() {

    try {

        const response = await fetch(
            `${DATA_URL}?v=${Date.now()}`,
            {
                cache: "no-store"
            }
        );


        if (!response.ok) {

            throw new Error(
                `HTTP Error ${response.status}`
            );
        }


        const payload = await response.json();


        skuData = Array.isArray(payload)
            ? payload
            : payload.rows || [];


        skuData = skuData.filter(row => {

            return String(
                row["SKU Number"] ?? ""
            ).trim() !== "";

        });


        recordCount.textContent =
            `${skuData.length.toLocaleString()} SKUs`;


        if (
            payload.meta &&
            payload.meta.generated_at
        ) {

            const updated =
                new Date(
                    payload.meta.generated_at
                );


            lastUpdated.textContent =
                `Updated ${updated.toLocaleString()}`;

        }


        renderResults(
            skuData.slice(0, 100)
        );


    } catch (error) {

        console.error(error);


        recordCount.textContent =
            "Load Error";


        statusText.textContent =
            "Could not load Excel data.";


        results.innerHTML = `
            <div class="empty">
                Data could not be loaded.
                Please run export_excel.py first.
            </div>
        `;

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


loadData();