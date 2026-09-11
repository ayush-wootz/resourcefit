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
        sheetViewport: document.getElementById("sheetViewport"),
        driveContainer: document.getElementById("driveContainer"),
        driveTitle: document.getElementById("driveTitle"),
        driveOpen: document.getElementById("driveOpen"),
        driveFrame: document.getElementById("driveFrame"),
        driveTabs: document.getElementById("driveTabs"),
        viewerControls: document.getElementById("viewerControls"),
        zoomControls: document.getElementById("zoomControls"),
        zoomIn: document.getElementById("zoomIn"),
        zoomOut: document.getElementById("zoomOut"),
        zoomReset: document.getElementById("zoomReset"),
        fullscreenToggle: document.getElementById("fullscreenToggle"),
        openTab: document.getElementById("openTab")
    };

    const query = new URLSearchParams(window.location.search);
    const sourceLink = query.get("link");
    const driveParam = query.get("driveId") || query.get("driveIds") || sourceLink;
    const driveNames = (query.get("names") || "").split(",").map((name) => name.trim());
    // Google Drive previews accept bare file ids, so the "open original" link is resolved separately.
    let originalLink = sourceLink;

    // --- Zoom and fullscreen controls ------------------------------------------
    let zoomController = null;

    function setZoomTarget(viewport, content) {
        zoomController = null;
        if (elements.zoomControls) elements.zoomControls.style.display = viewport ? "flex" : "none";
        if (!viewport || !window.RESOURCEFIT_ZOOM) return;
        zoomController = window.RESOURCEFIT_ZOOM.attach(viewport, content, (scale) => {
            if (elements.zoomReset) elements.zoomReset.textContent = `${Math.round(scale * 100)}%`;
        });
    }

    function fullscreenElement() {
        return document.fullscreenElement || document.webkitFullscreenElement || null;
    }

    async function toggleFullscreen() {
        const root = document.getElementById("container");
        if (fullscreenElement()) {
            const exit = document.exitFullscreen || document.webkitExitFullscreen;
            if (exit) await exit.call(document);
            return;
        }
        const request = root.requestFullscreen || root.webkitRequestFullscreen;
        try {
            if (!request) throw new Error("unsupported");
            await request.call(root);
        } catch (_) {
            // Glide and other hosts can embed this page without allow="fullscreen",
            // in which case the API rejects and only the in-embed expansion is possible.
            document.body.classList.toggle("maximized");
        }
        if (zoomController) zoomController.refresh();
    }

    function initializeControls() {
        if (!elements.fullscreenToggle) return;
        elements.fullscreenToggle.addEventListener("click", toggleFullscreen);
        elements.zoomIn.addEventListener("click", () => zoomController && zoomController.zoomBy(1.35));
        elements.zoomOut.addEventListener("click", () => zoomController && zoomController.zoomBy(1 / 1.35));
        elements.zoomReset.addEventListener("click", () => zoomController && zoomController.reset());
        if (elements.openTab) elements.openTab.href = window.location.href;
        window.addEventListener("resize", () => zoomController && zoomController.refresh());
    }

    function hideAll() {
        elements.status.style.display = "none";
        elements.webContainer.style.display = "none";
        elements.imageWrapper.style.display = "none";
        elements.pdfViewer.style.display = "none";
        elements.workbookContainer.style.display = "none";
        elements.driveContainer.style.display = "none";
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
        setZoomTarget(null);
        elements.status.style.display = "flex";
        elements.status.innerHTML = `
            <div class="status-card">
                ${loading ? '<div class="spinner" aria-hidden="true"></div>' : ""}
                <h1>${escapeHtml(title)}</h1>
                <p>${escapeHtml(message)}</p>
                ${originalLink ? `<div class="actions"><a class="button-link" href="${escapeHtml(originalLink)}" target="_blank" rel="noopener noreferrer">Download original</a></div>` : ""}
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
        setZoomTarget(elements.imageWrapper, elements.directImage);
        // Bounds depend on the decoded size, so re-clamp once the image lands.
        elements.directImage.onload = () => zoomController && zoomController.refresh();
    }

    function showPdf(url) {
        hideAll();
        setZoomTarget(null);
        elements.pdfViewer.style.display = "block";
        elements.pdfViewer.src = url;
    }

    function showWebpage(url) {
        hideAll();
        setZoomTarget(null);
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
        const surface = document.createElement("div");
        surface.id = "sheetSurface";
        surface.innerHTML = window.XLSX.utils.sheet_to_html(worksheet, {
            id: "sheetTable",
            editable: false
        });
        elements.sheetViewport.replaceChildren(surface);
        setZoomTarget(elements.sheetViewport, surface);
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

    // --- Google Drive previews -------------------------------------------------
    // Drive renders uploaded PDFs, Office files and images server-side, so a bare
    // file id can be embedded without CORS downloads, SheetJS or a backend.
    const DRIVE_ID_PATTERN = /^[A-Za-z0-9_-]{25,45}$/;
    const DRIVE_EDITOR_KINDS = ["spreadsheets", "document", "presentation", "drawings"];

    function driveTargetFrom(token) {
        const value = String(token || "").trim();
        if (!value) return null;
        if (DRIVE_ID_PATTERN.test(value)) return { id: value, kind: "file" };

        let parsed;
        try {
            parsed = new URL(value);
        } catch (_) {
            return null;
        }
        const host = parsed.hostname.toLowerCase();
        if (host !== "drive.google.com" && host !== "docs.google.com") return null;

        const segments = parsed.pathname.split("/").filter(Boolean);
        const marker = segments.indexOf("d");
        const id = (marker !== -1 ? segments[marker + 1] : "") || parsed.searchParams.get("id") || "";
        if (!DRIVE_ID_PATTERN.test(id)) return null;

        const kind = DRIVE_EDITOR_KINDS.includes(segments[0]) ? segments[0] : "file";
        return { id, kind };
    }

    function driveTargets(input) {
        const tokens = String(input || "").split(/[,\s]+/).filter(Boolean);
        if (!tokens.length) return [];
        const targets = tokens.map(driveTargetFrom);
        // Mixed input stays on the existing routes rather than half-rendering.
        return targets.every(Boolean) ? targets : [];
    }

    function drivePreviewUrl(target) {
        if (target.kind === "file") return `https://drive.google.com/file/d/${target.id}/preview`;
        return `https://docs.google.com/${target.kind}/d/${target.id}/preview`;
    }

    function driveOpenUrl(target) {
        if (target.kind === "file") return `https://drive.google.com/file/d/${target.id}/view`;
        return `https://docs.google.com/${target.kind}/d/${target.id}/view`;
    }

    function driveLabel(index) {
        return driveNames[index] || `File ${index + 1}`;
    }

    function showDriveTarget(targets, index, activeButton) {
        elements.driveTabs.querySelectorAll("button").forEach((button) => {
            const selected = button === activeButton;
            button.classList.toggle("active", selected);
            button.setAttribute("aria-selected", String(selected));
            button.tabIndex = selected ? 0 : -1;
        });

        const target = targets[index];
        elements.driveTitle.textContent = driveLabel(index);
        elements.driveOpen.href = driveOpenUrl(target);
        elements.driveFrame.src = drivePreviewUrl(target);
    }

    function showDrive(targets) {
        originalLink = driveOpenUrl(targets[0]);
        if (targets.length === 1) {
            showPdf(drivePreviewUrl(targets[0]));
            return;
        }

        hideAll();
        setZoomTarget(null);
        elements.driveContainer.style.display = "flex";
        elements.driveTabs.replaceChildren();
        targets.forEach((target, index) => {
            const button = document.createElement("button");
            button.type = "button";
            button.className = "sheet-tab";
            button.textContent = driveLabel(index);
            button.setAttribute("role", "tab");
            button.setAttribute("aria-selected", "false");
            button.addEventListener("click", () => showDriveTarget(targets, index, button));
            elements.driveTabs.appendChild(button);
            if (index === 0) showDriveTarget(targets, index, button);
        });
    }

    async function initialize() {
        initializeControls();
        const targets = driveTargets(driveParam);
        if (targets.length) {
            showDrive(targets);
            return;
        }

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
