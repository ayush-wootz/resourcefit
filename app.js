(function () {
    "use strict";

    const elements = {
        status: document.getElementById("status"),
        webContainer: document.getElementById("webContainer"),
        embeddedLink: document.getElementById("embeddedLink"),
        imageWrapper: document.getElementById("imageWrapper"),
        directImage: document.getElementById("directImage"),
        pdfViewer: document.getElementById("pdfViewer"),
        workbookContainer: document.getElementById("workbookContainer"),
        workbookTitle: document.getElementById("workbookTitle"),
        sheetTabs: document.getElementById("sheetTabs"),
        sheetViewport: document.getElementById("sheetViewport")
    };

    const query = new URLSearchParams(window.location.search);
    const sourceLink = query.get("link");

    function hideAll() {
        elements.status.style.display = "none";
        elements.webContainer.style.display = "none";
        elements.imageWrapper.style.display = "none";
        elements.pdfViewer.style.display = "none";
        elements.workbookContainer.style.display = "none";
        const sharePointPdf = document.getElementById("sharePointPdf");
        if (sharePointPdf) sharePointPdf.style.display = "none";
    }

    function escapeHtml(value) {
        return String(value)
            .replaceAll("&", "&amp;")
            .replaceAll("<", "&lt;")
            .replaceAll(">", "&gt;")
            .replaceAll('"', "&quot;")
            .replaceAll("'", "&#039;");
    }

    function showStatus(title, message, loading) {
        hideAll();
        elements.status.style.display = "flex";
        elements.status.innerHTML = `
            <div class="status-card">
                ${loading ? '<div class="spinner" aria-hidden="true"></div>' : ""}
                <h1>${escapeHtml(title)}</h1>
                <p>${escapeHtml(message)}</p>
                ${sourceLink ? `<div class="actions"><a class="button-link" href="${escapeHtml(sourceLink)}" target="_blank" rel="noopener noreferrer">Download original</a></div>` : ""}
            </div>`;
    }

    function getExtension(url) {
        try {
            const parsed = new URL(url);
            const pathExtension = parsed.pathname.split(".").pop().toLowerCase();
            if (/^(xlsx|xls|xlsm|xlsb|pdf|doc|docx|ppt|pptx|jpg|jpeg|png|gif|bmp|webp|svg)$/.test(pathExtension)) return pathExtension;
            const fileName = parsed.searchParams.get("file") || parsed.searchParams.get("filename") || "";
            return fileName.split(".").pop().toLowerCase();
        } catch (_) {
            return "";
        }
    }

    function detectContentType(url) {
        const extension = getExtension(url);
        if (["xlsx", "xls", "xlsm", "xlsb"].includes(extension)) return "workbook";
        if (["jpg", "jpeg", "png", "gif", "bmp", "webp", "svg"].includes(extension)) return "image";
        // Preserve the original document iframe route for non-Excel documents.
        if (["pdf", "doc", "docx", "ppt", "pptx"].includes(extension)) return "pdf";
        return "webpage";
    }

    function fileNameFromUrl(url) {
        try {
            const parsed = new URL(url);
            const queryName = parsed.searchParams.get("file") || parsed.searchParams.get("filename");
            const pathName = decodeURIComponent(parsed.pathname.split("/").pop() || "");
            return queryName || pathName || "Excel workbook";
        } catch (_) {
            return "Excel workbook";
        }
    }

    function showImage(url, onError) {
        hideAll();
        elements.imageWrapper.style.display = "block";
        elements.directImage.onerror = onError || null;
        elements.directImage.src = url;
    }

    function showPdf(url) {
        hideAll();
        elements.pdfViewer.style.display = "block";
        elements.pdfViewer.src = url;
    }

    function showWebpage(url) {
        hideAll();
        elements.webContainer.style.display = "block";
        elements.embeddedLink.src = url;
    }

    function renderSheet(workbook, sheetName, activeButton) {
        elements.sheetTabs.querySelectorAll("button").forEach((button) => {
            const selected = button === activeButton;
            button.classList.toggle("active", selected);
            button.setAttribute("aria-selected", String(selected));
            button.tabIndex = selected ? 0 : -1;
        });

        const worksheet = workbook.Sheets[sheetName];
        elements.sheetViewport.innerHTML = window.XLSX.utils.sheet_to_html(worksheet, {
            id: "sheetTable",
            editable: false
        });
        elements.sheetViewport.scrollTo({ top: 0, left: 0 });
    }

    function showWorkbook(workbook, title) {
        if (!workbook.SheetNames.length) throw new Error("The workbook does not contain any worksheets.");

        hideAll();
        elements.workbookContainer.style.display = "flex";
        elements.workbookTitle.textContent = title || fileNameFromUrl(sourceLink);
        elements.sheetTabs.replaceChildren();

        workbook.SheetNames.forEach((sheetName, index) => {
            const button = document.createElement("button");
            button.type = "button";
            button.className = "sheet-tab";
            button.textContent = sheetName;
            button.setAttribute("role", "tab");
            button.setAttribute("aria-selected", "false");
            button.addEventListener("click", () => renderSheet(workbook, sheetName, button));
            elements.sheetTabs.appendChild(button);
            if (index === 0) renderSheet(workbook, sheetName, button);
        });
    }

    async function loadWorkbook(url, title) {
        showStatus("Loading workbook", "Downloading and preparing the read-only preview…", true);
        if (!window.XLSX) throw new Error("The Excel preview library could not be loaded.");

        let response;
        try {
            response = await fetch(url, { mode: "cors", credentials: "omit" });
        } catch (_) {
            throw new Error("The file could not be downloaded. Use a direct Glide file URL that permits browser access.");
        }
        if (!response.ok) throw new Error(`The file download failed (HTTP ${response.status}).`);

        const data = await response.arrayBuffer();
        if (!data.byteLength) throw new Error("The downloaded file is empty.");
        const workbook = window.XLSX.read(data, { type: "array", cellDates: true });
        showWorkbook(workbook, title);
    }

    async function initialize() {
        if (!sourceLink) {
            showStatus("No file selected", "Provide a link parameter to preview a file.");
            return;
        }
        try {
            new URL(sourceLink);
        } catch (_) {
            showStatus("Invalid link", "The supplied file link is not a valid URL.");
            return;
        }

        try {
            const apiBase = window.RESOURCEFIT_CONFIG?.sharePointApiBase;
            if (apiBase && new URL(sourceLink).hostname.endsWith(".sharepoint.com")) {
                const { previewSharePoint } = await import("./sharepoint-preview.mjs?v=1");
                await previewSharePoint({ sourceLink, apiBase, showStatus, showImage, loadWorkbook, hideAll });
                return;
            }
            switch (detectContentType(sourceLink)) {
                case "workbook": await loadWorkbook(sourceLink); break;
                case "image": showImage(sourceLink); break;
                case "pdf": showPdf(sourceLink); break;
                default: showWebpage(sourceLink);
            }
        } catch (error) {
            showStatus("Preview unavailable", error && error.message ? error.message : "The file could not be previewed.");
        }
    }

    initialize();
}());
