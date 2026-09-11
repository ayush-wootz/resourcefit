// Loaded only for SharePoint URLs when a backend has been configured.
const PDF_JS = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/";

export async function previewSharePoint({ sourceLink, apiBase, showStatus, showImage, loadWorkbook, hideAll }) {
    const base = new URL(apiBase);
    if (base.protocol !== "https:" || base.username || base.password || base.search || base.hash || base.pathname !== "/") {
        throw new Error("The SharePoint preview service URL is not configured correctly.");
    }
    showStatus("Loading preview", "Retrieving the shared file…", true);
    const endpoint = new URL("/resolve", base);
    endpoint.searchParams.set("link", sourceLink);
    const response = await fetch(endpoint, { mode: "cors", credentials: "omit", signal: AbortSignal.timeout(120_000) });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || "The shared file could not be loaded.");
    const content = new URL(result.contentUrl, base);
    if (content.origin !== base.origin || content.pathname !== "/content") throw new Error("The preview service returned an invalid file URL.");
    switch (result.type) {
        case "workbook": await loadWorkbook(content.href, result.name); break;
        case "image":
            showImage(content.href, () => showStatus("Preview unavailable", "The shared image could not be loaded. Please open the original file."));
            break;
        case "pdf": await renderPdf(content.href, result.name, sourceLink, hideAll, showStatus); break;
        default: showStatus("Preview not supported", "This file type can be opened using Download original.");
    }
}

async function renderPdf(url, name, original, hideAll, showStatus) {
    // Canvas avoids the browser PDF plugin, which can be blocked inside Glide embeds.
    const pdfjs = await import(`${PDF_JS}pdf.min.mjs`);
    pdfjs.GlobalWorkerOptions.workerSrc = `${PDF_JS}pdf.worker.min.mjs`;
    const pdf = await pdfjs.getDocument({ url, withCredentials: false, isEvalSupported: false,
        cMapUrl: "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/cmaps/", cMapPacked: true,
        standardFontDataUrl: "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/standard_fonts/" }).promise;
    hideAll();
    let section = document.getElementById("sharePointPdf");
    if (!section) {
        section = document.createElement("section");
        section.id = "sharePointPdf";
        document.getElementById("container").appendChild(section);
    }
    section.setAttribute("aria-label", `PDF preview: ${name}`);
    section.style.cssText = "display:flex;flex-direction:column;height:100%;background:#eef1f4";
    const bar = document.createElement("nav");
    bar.setAttribute("aria-label", "PDF pages");
    bar.style.cssText = "display:flex;align-items:center;gap:12px;flex-wrap:wrap;padding:10px;background:white;border-bottom:1px solid #dce2ea";
    const previous = document.createElement("button");
    previous.textContent = "Previous";
    const next = document.createElement("button");
    next.textContent = "Next";
    for (const button of [previous, next]) { button.type = "button"; button.className = "button-link"; }
    const label = document.createElement("span");
    label.setAttribute("aria-live", "polite");
    const open = document.createElement("a");
    open.textContent = "Open original";
    open.href = original;
    open.target = "_blank";
    open.rel = "noopener noreferrer";
    open.style.marginLeft = "auto";
    bar.append(previous, label, next, open);
    const viewport = document.createElement("div");
    viewport.style.cssText = "flex:1;min-height:0;overflow:auto;padding:12px;text-align:center";
    const canvas = document.createElement("canvas");
    canvas.setAttribute("role", "img");
    canvas.style.cssText = "max-width:100%;height:auto;background:white";
    viewport.appendChild(canvas);
    section.replaceChildren(bar, viewport);
    let pageNumber = 1;
    async function draw() {
        previous.disabled = next.disabled = true;
        const page = await pdf.getPage(pageNumber);
        const normal = page.getViewport({ scale: 1 });
        const width = Math.max(240, viewport.clientWidth - 24);
        const scale = Math.min(width / normal.width * Math.min(window.devicePixelRatio || 1, 2),
            Math.sqrt(12_000_000 / (normal.width * normal.height)));
        const size = page.getViewport({ scale });
        canvas.width = Math.ceil(size.width);
        canvas.height = Math.ceil(size.height);
        canvas.setAttribute("aria-label", `${name}, page ${pageNumber} of ${pdf.numPages}. Open original for accessible text.`);
        await page.render({ canvasContext: canvas.getContext("2d"), viewport: size }).promise;
        label.textContent = `${pageNumber} / ${pdf.numPages}`;
        previous.disabled = pageNumber === 1;
        next.disabled = pageNumber === pdf.numPages;
        viewport.scrollTop = 0;
        page.cleanup();
    }
    for (const [button, step] of [[previous, -1], [next, 1]]) {
        button.addEventListener("click", async () => {
            pageNumber += step;
            try { await draw(); } catch { showStatus("Preview unavailable", "This PDF page could not be rendered. Please open the original file."); }
        });
    }
    await draw();
}
