import http from "node:http";
import { pathToFileURL } from "node:url";

class PreviewError extends Error {
    constructor(status, message) { super(message); this.status = status; }
}
const fail = (status, message) => { throw new PreviewError(status, message); };
const MB = 1024 * 1024;

function httpsUrl(value) {
    let url;
    try { url = new URL(value); } catch { fail(400, "Invalid file link."); }
    if (url.protocol !== "https:" || url.username || url.password || url.port || url.hash) {
        fail(400, "An HTTPS link without credentials or a fragment is required.");
    }
    return url;
}

export function configuration(env = process.env) {
    const required = ["MS_TENANT_ID", "MS_CLIENT_ID", "MS_CLIENT_SECRET", "SHAREPOINT_HOST", "ALLOWED_ORIGINS"];
    for (const name of required) if (!env[name]) throw new Error(`Set ${name} before starting the backend.`);
    const guid = /^[a-f\d]{8}(?:-[a-f\d]{4}){3}-[a-f\d]{12}$/i;
    if (!guid.test(env.MS_TENANT_ID) || !guid.test(env.MS_CLIENT_ID)) throw new Error("Microsoft tenant and client IDs must be GUIDs.");
    const host = env.SHAREPOINT_HOST.toLowerCase();
    if (!/^[a-z0-9-]+\.sharepoint\.com$/.test(host)) throw new Error("Set SHAREPOINT_HOST to your exact SharePoint hostname.");
    if (env.ALLOW_DYNAMIC_PUBLIC_LINKS && !["true", "false"].includes(env.ALLOW_DYNAMIC_PUBLIC_LINKS)) {
        throw new Error("ALLOW_DYNAMIC_PUBLIC_LINKS must be true or false.");
    }
    const allowDynamic = env.ALLOW_DYNAMIC_PUBLIC_LINKS === "true";
    const links = JSON.parse(env.PUBLIC_SHARE_LINKS?.trim() || "[]");
    if (!Array.isArray(links) || (!links.length && !allowDynamic)) {
        throw new Error("Set ALLOW_DYNAMIC_PUBLIC_LINKS=true or register public links in PUBLIC_SHARE_LINKS.");
    }
    const publicLinks = new Set(links.map(value => {
        const url = httpsUrl(value);
        if (url.hostname !== host) throw new Error("Every registered link must use SHAREPOINT_HOST.");
        return url.href;
    }));
    const origins = new Set(env.ALLOWED_ORIGINS.split(",").map(value => {
        const url = httpsUrl(value.trim());
        if (url.href !== `${url.origin}/`) throw new Error("ALLOWED_ORIGINS must contain HTTPS origins only.");
        return url.origin;
    }));
    const positive = (key, fallback, maximum) => {
        const value = Number(env[key] ?? fallback);
        if (!Number.isSafeInteger(value) || value < 1 || value > maximum) throw new Error(`Invalid ${key}.`);
        return value;
    };
    const maxFileBytes = positive("MAX_FILE_MB", 20, 100) * MB;
    const maxCacheBytes = positive("MAX_CACHE_MB", 128, 1024) * MB;
    if (maxCacheBytes < maxFileBytes) throw new Error("MAX_CACHE_MB must be at least MAX_FILE_MB.");
    return {
        tenant: env.MS_TENANT_ID, client: env.MS_CLIENT_ID, secret: env.MS_CLIENT_SECRET,
        host, publicLinks, allowDynamic, origins, maxFileBytes, maxCacheBytes,
        maxCacheEntries: positive("MAX_CACHE_ENTRIES", 200, 2000),
        ttl: positive("CACHE_TTL_SECONDS", 300, 300) * 1000,
        maxConcurrent: positive("MAX_CONCURRENT_FETCHES", 4, 16),
        port: positive("PORT", 8080, 65535)
    };
}

export function fileType(name, mime) {
    const ext = name.split(".").pop().toLowerCase();
    if (["xlsx", "xls", "xlsm", "xlsb"].includes(ext)) return "workbook";
    if (mime === "application/pdf" || ext === "pdf") return "pdf";
    if (/^image\/(png|jpeg|gif|webp|bmp|avif)$/.test(mime)) return "image";
    return "download";
}

export function createPreviewService(config, { fetchImpl = fetch, now = Date.now } = {}) {
    const cache = new Map();
    const pending = new Map();
    let cacheBytes = 0;
    let token;
    let tokenPending;

    function drop(key) {
        const entry = cache.get(key);
        if (entry) cacheBytes -= entry.bytes.length;
        cache.delete(key);
    }
    function remember(key, entry) {
        drop(key);
        while (cache.size && (cacheBytes + entry.bytes.length > config.maxCacheBytes || cache.size >= (config.maxCacheEntries ?? 200))) {
            drop(cache.keys().next().value);
        }
        cache.set(key, entry);
        cacheBytes += entry.bytes.length;
    }
    async function request(url, options = {}) {
        return fetchImpl(url, { ...options, redirect: "manual", signal: AbortSignal.timeout(30_000) });
    }
    async function accessToken() {
        if (token && token.until > now()) return token.value;
        if (!tokenPending) {
            tokenPending = (async () => {
                const response = await request(`https://login.microsoftonline.com/${config.tenant}/oauth2/v2.0/token`, {
                    method: "POST",
                    body: new URLSearchParams({ client_id: config.client, client_secret: config.secret,
                        scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" })
                });
                if (!response.ok) fail(502, "Microsoft server authentication failed. Contact the viewer administrator.");
                const body = await response.json();
                if (!body.access_token || !Number.isFinite(Number(body.expires_in))) fail(502, "Invalid Microsoft authentication response.");
                token = { value: body.access_token, until: now() + Math.max(0, Number(body.expires_in) - 60) * 1000 };
                return token.value;
            })().finally(() => { tokenPending = undefined; });
        }
        return tokenPending;
    }
    async function graph(path, retry = true) {
        const response = await request(`https://graph.microsoft.com/v1.0${path}`, {
            headers: { Authorization: `Bearer ${await accessToken()}` }
        });
        if (response.status === 401 && retry) { token = undefined; return graph(path, false); }
        if ([403, 404, 410].includes(response.status)) fail(403, "This sharing link is unavailable or access has been removed.");
        if (response.status === 429) fail(503, "Microsoft is busy. Please try again shortly.");
        if (!response.ok) fail(502, "Microsoft could not resolve this sharing link. Contact the viewer administrator.");
        return response.json();
    }
    function publicPermission(permission) {
        // Check the permission for THIS link, never an arbitrary public permission on the item.
        if (permission?.link?.scope !== "anonymous" || !["view", "edit"].includes(permission.link.type) || permission.hasPassword === true) {
            fail(403, "Only password-free Anyone sharing links can be previewed.");
        }
        const expiration = permission.expirationDateTime;
        const expiresAt = !expiration || expiration.startsWith("0001-01-01") ? Infinity : Date.parse(expiration);
        if (!(expiresAt > now())) fail(403, "This sharing link has expired.");
        return expiresAt;
    }
    async function download(value, anonymous = false) {
        for (let redirects = 0; redirects <= 3; redirects++) {
            const url = httpsUrl(value);
            // No arbitrary proxy, and no bearer token is sent to the download endpoint.
            if (url.hostname !== config.host) {
                fail(anonymous ? 403 : 502, anonymous
                    ? "Public download could not be verified. Use an Anyone sharing link that opens without sign-in or a password."
                    : "Microsoft returned a download host that is not configured.");
            }
            const response = await request(url.href, { credentials: "omit" });
            if ([301, 302, 303, 307, 308].includes(response.status)) {
                await response.body?.cancel();
                const location = response.headers.get("location");
                if (!location) fail(502, "Microsoft returned an invalid download redirect.");
                value = new URL(location, url).href;
                continue;
            }
            if (!response.ok) fail(anonymous && [401, 403, 404].includes(response.status) ? 403 : 502, "The SharePoint file could not be downloaded.");
            const mime = (response.headers.get("content-type") || "").split(";")[0].trim().toLowerCase();
            if (anonymous && (!mime || /html|json|xml/.test(mime) && !/officedocument/.test(mime))) {
                await response.body?.cancel();
                fail(403, "SharePoint returned a viewing or sign-in page instead of a public file download.");
            }
            if (Number(response.headers.get("content-length")) > config.maxFileBytes) {
                await response.body?.cancel();
                fail(413, "This file exceeds the preview size limit.");
            }
            const chunks = [];
            let length = 0;
            for await (const chunk of response.body) {
                length += chunk.length;
                if (length > config.maxFileBytes) fail(413, "This file exceeds the preview size limit.");
                chunks.push(Buffer.from(chunk));
            }
            if (!length) fail(422, "The file is empty.");
            const bytes = Buffer.concat(chunks, length);
            if (anonymous && /^\s*(?:<!doctype\s+html|<html|<head|<body|<script)/i.test(bytes.subarray(0, 1024).toString("utf8"))) {
                fail(403, "Public file access could not be verified.");
            }
            return bytes;
        }
        fail(502, "Too many SharePoint download redirects.");
    }
    async function refresh(key, old) {
        const startedAt = now();
        const share = `u!${Buffer.from(key).toString("base64url")}`;
        const permission = await graph(`/shares/${share}/permission`);
        const expiresAt = publicPermission(permission);
        const item = await graph(`/shares/${share}/driveItem`);
        if (!item.file || !item.id || typeof item.name !== "string") fail(422, "This link must point to a single file.");
        if (!Number.isSafeInteger(item.size) || item.size < 1 || item.size > config.maxFileBytes) fail(413, "This file is empty or exceeds the preview size limit.");
        const version = item.eTag || item.cTag;
        const identity = `${item.parentReference?.driveId || ""}/${item.id}`;
        const unchanged = old && version && old.version === version && old.identity === identity;
        let bytes;
        if (config.allowDynamic) {
            // Graph may omit SharePoint password metadata. Prove public access by fetching
            // the ORIGINAL sharing link without app tokens, cookies or Graph's signed URL.
            // This check runs even when the version is unchanged at cache revalidation.
            const publicDownload = new URL(key);
            publicDownload.searchParams.set("download", "1");
            bytes = await download(publicDownload.href, true);
            if (bytes.length !== item.size) fail(502, "The public download did not match the file metadata. Please try again.");
            const extension = item.name.split(".").pop().toLowerCase();
            const zip = bytes[0] === 0x50 && bytes[1] === 0x4b;
            const ole = bytes.subarray(0, 8).equals(Buffer.from("d0cf11e0a1b11ae1", "hex"));
            if ((["xlsx", "xlsm", "xlsb"].includes(extension) && !zip && !ole)
                || (extension === "pdf" && !bytes.subarray(0, 1024).includes(Buffer.from("%PDF-")))) {
                fail(403, "SharePoint did not return the expected public file. Open the original link to check access.");
            }
        } else {
            bytes = unchanged ? old.bytes : await download(item["@microsoft.graph.downloadUrl"]);
        }
        const validUntil = Math.min(startedAt + config.ttl, expiresAt);
        if (validUntil <= now()) fail(403, "The sharing link expired while loading. Please try again.");
        const mime = typeof item.file.mimeType === "string" ? item.file.mimeType.toLowerCase() : "application/octet-stream";
        const entry = { name: item.name, mime, type: fileType(item.name, mime), bytes, version, identity, validUntil };
        remember(key, entry);
        return entry;
    }
    async function get(value) {
        const url = httpsUrl(value);
        const key = url.href;
        if (url.hostname !== config.host) fail(403, "This link is outside the configured SharePoint domain.");
        if (!config.allowDynamic && !config.publicLinks.has(key)) fail(403, "This public sharing link has not been registered for preview.");
        const old = cache.get(key);
        if (old && old.validUntil > now()) { cache.delete(key); cache.set(key, old); return old; }
        if (pending.has(key)) return pending.get(key);
        if (pending.size >= config.maxConcurrent) fail(503, "The preview service is busy. Please try again shortly.");
        const task = refresh(key, old).catch(error => { drop(key); throw error; }).finally(() => pending.delete(key));
        pending.set(key, task);
        return task;
    }
    return { get };
}

export function createServer(config, dependencies = {}) {
    const service = createPreviewService(config, dependencies);
    return http.createServer({ maxHeaderSize: 16_384, requestTimeout: 30_000 }, async (req, res) => {
        res.setHeader("Cache-Control", "no-store");
        res.setHeader("X-Content-Type-Options", "nosniff");
        res.setHeader("Referrer-Policy", "no-referrer");
        res.setHeader("Vary", "Origin");
        try {
            const origin = req.headers.origin;
            if (origin && !config.origins.has(origin)) fail(403, "This viewer origin is not allowed.");
            if (origin) res.setHeader("Access-Control-Allow-Origin", origin);
            res.setHeader("Access-Control-Allow-Methods", "GET, HEAD, OPTIONS");
            res.setHeader("Access-Control-Allow-Headers", "Range");
            res.setHeader("Access-Control-Expose-Headers", "Content-Length, Content-Range, Accept-Ranges");
            if (req.method === "OPTIONS") { res.writeHead(204); res.end(); return; }
            if (!["GET", "HEAD"].includes(req.method)) { res.setHeader("Allow", "GET, HEAD, OPTIONS"); fail(405, "Method not allowed."); }
            const url = new URL(req.url, "http://backend.local");
            if (url.pathname === "/healthz") {
                res.setHeader("Content-Type", "application/json"); res.end('{"ok":true}'); return;
            }
            if (!["/resolve", "/content"].includes(url.pathname)) fail(404, "Not found.");
            const link = url.searchParams.get("link");
            const entry = await service.get(link);
            if (url.pathname === "/resolve") {
                res.setHeader("Content-Type", "application/json");
                res.end(JSON.stringify({ name: entry.name, type: entry.type,
                    contentUrl: `/content?${new URLSearchParams({ link })}` }));
                return;
            }
            const safeMime = entry.type === "pdf" ? "application/pdf" : entry.type === "image" ? entry.mime : "application/octet-stream";
            res.setHeader("Content-Type", safeMime);
            res.setHeader("Content-Security-Policy", "default-src 'none'; sandbox");
            res.setHeader("Content-Disposition", `${["image", "pdf"].includes(entry.type) ? "inline" : "attachment"}; filename*=UTF-8''${encodeURIComponent(entry.name).replaceAll("'", "%27")}`);
            res.setHeader("Accept-Ranges", "bytes");
            let start = 0;
            let end = entry.bytes.length - 1;
            if (req.headers.range) {
                const match = /^bytes=(\d*)-(\d*)$/.exec(req.headers.range);
                if (!match || (!match[1] && !match[2])) { res.setHeader("Content-Range", `bytes */${entry.bytes.length}`); fail(416, "Invalid byte range."); }
                if (match[1]) {
                    start = Number(match[1]);
                    if (match[2]) end = Math.min(Number(match[2]), end);
                } else start = Math.max(0, entry.bytes.length - Number(match[2]));
                if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start > end) { res.setHeader("Content-Range", `bytes */${entry.bytes.length}`); fail(416, "Invalid byte range."); }
                res.statusCode = 206;
                res.setHeader("Content-Range", `bytes ${start}-${end}/${entry.bytes.length}`);
            }
            res.setHeader("Content-Length", end - start + 1);
            res.end(req.method === "HEAD" ? undefined : entry.bytes.subarray(start, end + 1));
        } catch (error) {
            res.removeHeader("Content-Length");
            res.removeHeader("Content-Disposition");
            res.statusCode = error instanceof PreviewError ? error.status : 502;
            res.setHeader("Content-Type", "application/json");
            // Never log upstream bodies, sharing URLs, tokens or signed download URLs.
            res.end(JSON.stringify({ error: error instanceof PreviewError ? error.message : "The preview service could not reach Microsoft. Please try again." }));
        }
    });
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
    const config = configuration();
    createServer(config).listen(config.port, "0.0.0.0", () => console.log(`ResourceFit backend listening on port ${config.port}`));
}
