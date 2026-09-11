import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { previewSharePoint } from "../../sharepoint-preview.mjs";

const source = await readFile(new URL("../../app.js", import.meta.url), "utf8");
async function render(link, apiBase = "") {
    const elements = new Map();
    const element = () => ({ style: {}, classList: { toggle() {} }, setAttribute() {}, addEventListener() {},
        appendChild() {}, replaceChildren() {}, querySelectorAll: () => [], scrollTo() {} });
    const document = { getElementById(id) { if (!elements.has(id)) elements.set(id, element()); return elements.get(id); }, createElement: element };
    const requests = [];
    const context = { document, URL, URLSearchParams,
        window: { location: { search: `?${new URLSearchParams({ link })}` }, RESOURCEFIT_CONFIG: { sharePointApiBase: apiBase },
            XLSX: { read: () => ({ SheetNames: ["Sheet1"], Sheets: { Sheet1: {} } }), utils: { sheet_to_html: () => "<table></table>" } } },
        fetch: async url => { requests.push(url); return { ok: true, arrayBuffer: async () => new ArrayBuffer(4) }; }
    };
    vm.runInNewContext(source, context);
    await new Promise(resolve => setImmediate(resolve));
    return { elements, requests };
}

test("existing direct-file routing remains identical with backend disabled or enabled", async () => {
    const routes = [
        ...["xlsx", "xls", "xlsm", "xlsb"].map(ext => [ext, "workbookContainer"]),
        ...["png", "jpg", "jpeg", "gif", "webp", "bmp", "svg"].map(ext => [ext, "imageWrapper"]),
        ...["pdf", "doc", "docx", "ppt", "pptx"].map(ext => [ext, "pdfViewer"]),
        ["html", "webContainer"], ["unknown", "webContainer"]
    ];
    for (const base of ["", "https://preview.example"]) {
        for (const [ext, target] of routes) {
            for (const url of [`https://storage.googleapis.com/files/RFQ%20N-6.${ext}`, `https://files.example/download?file=test.${ext.toUpperCase()}`]) {
                const { elements, requests } = await render(url, base);
                assert.equal(elements.get(target).style.display, target === "workbookContainer" ? "flex" : "block", `${base}: ${url}`);
                if (target === "workbookContainer") assert.deepEqual(requests, [url]);
                else assert.equal(requests.length, 0);
            }
        }
    }
});

test("unconfigured SharePoint links retain the previous route", async () => {
    const link = "https://tenant-my.sharepoint.com/:x:/g/public-token";
    const result = await render(link);
    assert.equal(result.elements.get("embeddedLink").src, link);
    assert.equal(result.requests.length, 0);
});

test("SharePoint resolver preserves URL encoding, resolved file name and backend errors", async t => {
    const originalFetch = globalThis.fetch;
    t.after(() => { globalThis.fetch = originalFetch; });
    const sourceLink = "https://tenant-my.sharepoint.com/:x:/g/token?e=x&file=RFQ%20Washer.xlsx";
    const calls = [];
    const argumentsForPreview = { sourceLink, apiBase: "https://preview.example", showStatus() {}, hideAll() {},
        showImage: url => calls.push(["image", url]), loadWorkbook: async (url, name) => calls.push(["workbook", url, name]) };
    globalThis.fetch = async (url, options) => {
        assert.equal(url.searchParams.get("link"), sourceLink);
        assert.equal(options.credentials, "omit");
        return { ok: true, json: async () => ({ type: "workbook", name: "RFQ Washer.xlsx", contentUrl: "/content?link=encoded" }) };
    };
    await previewSharePoint(argumentsForPreview);
    assert.deepEqual(calls, [["workbook", "https://preview.example/content?link=encoded", "RFQ Washer.xlsx"]]);
    globalThis.fetch = async () => ({ ok: false, json: async () => ({ error: "Link revoked" }) });
    await assert.rejects(previewSharePoint(argumentsForPreview), /Link revoked/);
    globalThis.fetch = async () => ({ ok: true, json: async () => ({ type: "image", contentUrl: "https://evil.example/content" }) });
    await assert.rejects(previewSharePoint(argumentsForPreview), /invalid file URL/);
    await assert.rejects(previewSharePoint({ ...argumentsForPreview, apiBase: "http://preview.example" }), /not configured correctly/);
});
