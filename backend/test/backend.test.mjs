import test from "node:test";
import assert from "node:assert/strict";
import { once } from "node:events";
import { configuration, createPreviewService, createServer, fileType } from "../server.mjs";

const host = "example-my.sharepoint.com";
const link = `https://${host}/:x:/g/personal/user/public-token?e=abc`;
const otherLink = `https://${host}/:x:/g/personal/user/other-token`;
const origin = "https://ayush-wootz.github.io";
const config = {
    tenant: "tenant", client: "client", secret: "server-only-secret", host,
    publicLinks: new Set([link, otherLink]), origins: new Set([origin]),
    maxFileBytes: 1024, maxCacheBytes: 2048, ttl: 300_000, maxConcurrent: 4
};
const json = (body, status = 200) => new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } });

function fixture(overrides = {}) {
    const state = {
        time: 1_800_000_000_000, requests: [], permission: { link: { scope: "anonymous", type: "view" } },
        status: 200, bytes: "workbook-v1", version: "v1", download: `https://${host}/download?secret=signed`,
        ...overrides
    };
    const fetchImpl = async (value, options) => {
        const url = String(value);
        state.requests.push({ url, options });
        assert.equal(options.redirect, "manual");
        if (url.includes("login.microsoftonline.com")) return json({ access_token: "graph-token", expires_in: 3600 });
        if (url.startsWith("https://graph.microsoft.com/")) {
            assert.equal(options.headers.Authorization, "Bearer graph-token");
            if (state.status !== 200) return json({ error: { code: state.errorCode, message: "private upstream message must not leak" } }, state.status);
            if (url.endsWith("/permission")) {
                if (state.permissionStatus) return json({ error: { code: "notSupported" } }, state.permissionStatus);
                return json(state.permission);
            }
            if (new URL(url).pathname.endsWith("/permissions")) return json(state.permissionPages?.[url] || { value: state.permissions || [] });
            return json({ id: "item", parentReference: { driveId: "drive" }, name: "RFQ Washer.xlsx", size: state.bytes.length,
                eTag: state.version, file: { mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
                "@microsoft.graph.downloadUrl": state.download });
        }
        assert.equal(new URL(url).hostname, host);
        assert.equal(options.headers?.Authorization, undefined);
        if (state.redirect) return new Response(null, { status: 302, headers: { location: state.redirect } });
        return new Response(state.bytes, { headers: { "content-type": state.mime || "application/octet-stream" } });
    };
    const dependencies = { fetchImpl, now: () => state.time };
    return { state, dependencies, service: createPreviewService(config, dependencies) };
}

test("first click downloads once, cache avoids Microsoft, unchanged version reuses bytes", async () => {
    const { state, service } = fixture();
    const first = await service.get(link);
    assert.equal(first.name, "RFQ Washer.xlsx");
    assert.equal(state.requests.length, 4);
    assert.equal(await service.get(link), first);
    assert.equal(state.requests.length, 4);
    state.time += 300_001;
    const refreshed = await service.get(link);
    assert.equal(refreshed.bytes, first.bytes);
    assert.equal(state.requests.length, 6);
    state.time += 300_001;
    state.version = "v2";
    state.bytes = "workbook-v2";
    assert.equal((await service.get(link)).bytes.toString(), "workbook-v2");
    assert.equal(state.requests.length, 9);
});

test("concurrent opens of one file share a single fetch", async () => {
    const { service, state } = fixture();
    const values = await Promise.all(Array.from({ length: 10 }, () => service.get(link)));
    assert.ok(values.every(value => value === values[0]));
    assert.equal(state.requests.length, 4);
});

test("unregistered, wrong-host, insecure and credential-bearing URLs never reach Microsoft", async () => {
    const { service, state } = fixture();
    for (const value of [otherLink + "?not=registered", "https://evil.example/file", link.replace("https:", "http:"), link.replace("https://", "https://user:pass@"), link + "#x", "invalid"]) {
        await assert.rejects(service.get(value));
    }
    assert.equal(state.requests.length, 0);
});

test("private, password, blocked-download, missing and expired permissions are denied", async () => {
    for (const permission of [
        {}, { link: { scope: "organization", type: "view" } },
        { link: { scope: "anonymous", type: "blocksDownload" } },
        { link: { scope: "anonymous", type: "view" }, hasPassword: true },
        { link: { scope: "anonymous", type: "view" }, expirationDateTime: "2020-01-01T00:00:00Z" },
        { link: { scope: "anonymous", type: "view" }, expirationDateTime: "invalid" }
    ]) {
        const { service, state } = fixture({ permission });
        await assert.rejects(service.get(link), { status: 403 });
        assert.equal(state.requests.length, 2);
    }
});

test("expiry caps cache lifetime; revoked and upstream-error entries are not served stale", async () => {
    const { service, state } = fixture();
    state.permission.expirationDateTime = new Date(state.time + 1000).toISOString();
    await service.get(link);
    state.time += 1001;
    await assert.rejects(service.get(link), { status: 403 });
    delete state.permission.expirationDateTime;
    await service.get(link);
    for (const status of [403, 404, 429, 500]) {
        state.time += 300_001;
        state.status = status;
        await assert.rejects(service.get(link));
    }
    state.status = 200;
    const before = state.requests.length;
    await service.get(link);
    assert.equal(state.requests.length - before, 3); // evicted bytes require a new download
});

test("download URLs and redirects cannot escape the configured tenant host", async () => {
    for (const options of [
        { download: "https://127.0.0.1/private" },
        { download: "https://example-my.sharepoint.com.evil.example/file" },
        { redirect: "http://169.254.169.254/latest/meta-data/" },
        { redirect: "https://evil.example/file" }
    ]) {
        const { service } = fixture(options);
        await assert.rejects(service.get(link));
    }
});

test("file size and concurrent work are bounded", async () => {
    const large = fixture({ bytes: "x".repeat(1025) });
    await assert.rejects(large.service.get(link), { status: 413 });
    const base = fixture();
    const limited = createPreviewService({ ...config, maxConcurrent: 1 }, base.dependencies);
    const first = limited.get(link);
    await assert.rejects(limited.get(otherLink), { status: 503 });
    await first;
});

test("LRU cache evicts file bytes at the configured budget", async () => {
    const { state, dependencies } = fixture();
    const service = createPreviewService({ ...config, maxCacheBytes: 11 }, dependencies);
    await service.get(link);
    await service.get(otherLink);
    const before = state.requests.length;
    await service.get(link);
    assert.equal(state.requests.length - before, 3);
});

test("HTTP resolve, content, ranges, HEAD, CORS and errors do not expose credentials", async t => {
    const { dependencies } = fixture();
    const server = createServer(config, dependencies);
    server.listen(0, "127.0.0.1");
    await once(server, "listening");
    t.after(() => { server.closeAllConnections(); server.close(); });
    const base = `http://127.0.0.1:${server.address().port}`;
    const query = new URLSearchParams({ link });
    const resolve = await fetch(`${base}/resolve?${query}`, { headers: { Origin: origin } });
    assert.equal(resolve.status, 200);
    assert.equal(resolve.headers.get("access-control-allow-origin"), origin);
    assert.equal(resolve.headers.get("cache-control"), "no-store");
    const metadata = await resolve.json();
    assert.equal(metadata.type, "workbook");
    assert.ok(!JSON.stringify(metadata).includes("signed"));
    const content = await fetch(base + metadata.contentUrl);
    assert.equal(content.headers.get("content-type"), "application/octet-stream");
    assert.match(content.headers.get("content-disposition"), /^attachment/);
    assert.equal(await content.text(), "workbook-v1");
    for (const [range, expected] of [["bytes=0-3", "work"], ["bytes=-2", "v1"], ["bytes=9-", "v1"]]) {
        const response = await fetch(base + metadata.contentUrl, { headers: { Range: range } });
        assert.equal(response.status, 206);
        assert.equal(await response.text(), expected);
    }
    for (const range of ["bytes=999-", "bytes=-0", "bytes=0-1,3-4", "bytes=-"]) {
        const response = await fetch(base + metadata.contentUrl, { headers: { Range: range } });
        assert.equal(response.status, 416);
    }
    const head = await fetch(base + metadata.contentUrl, { method: "HEAD" });
    assert.equal(head.headers.get("content-length"), "11");
    assert.equal(await head.text(), "");
    assert.equal((await fetch(`${base}/resolve?${query}`, { headers: { Origin: "https://evil.example" } })).status, 403);
    assert.equal((await fetch(`${base}/resolve?${query}`, { method: "POST" })).status, 405);
});

test("file types support Excel, PDF, raster images, and a safe unsupported fallback", () => {
    assert.equal(fileType("a.XLSX", "application/octet-stream"), "workbook");
    assert.equal(fileType("a", "application/pdf"), "pdf");
    assert.equal(fileType("a", "image/png"), "image");
    for (const mime of ["text/html", "image/svg+xml", "application/javascript"]) assert.equal(fileType("a", mime), "download");
});

test("startup fails closed for absent config and invalid registrations", () => {
    assert.throws(() => configuration({}), /MS_TENANT_ID/);
    const env = { MS_TENANT_ID: "00000000-0000-0000-0000-000000000000", MS_CLIENT_ID: "00000000-0000-0000-0000-000000000001",
        MS_CLIENT_SECRET: "secret", SHAREPOINT_HOST: host, PUBLIC_SHARE_LINKS: JSON.stringify([link]), ALLOWED_ORIGINS: origin };
    assert.equal(configuration(env).ttl, 300000);
    assert.throws(() => configuration({ ...env, PUBLIC_SHARE_LINKS: "[]" }));
    assert.throws(() => configuration({ ...env, PUBLIC_SHARE_LINKS: '["https://evil.example/file"]' }));
    assert.throws(() => configuration({ ...env, CACHE_TTL_SECONDS: "301" }));
    const dynamic = configuration({ ...env, ALLOW_DYNAMIC_PUBLIC_LINKS: "true", PUBLIC_SHARE_LINKS: "" });
    assert.equal(dynamic.allowDynamic, true);
    assert.equal(dynamic.publicLinks.size, 0);
    assert.throws(() => configuration({ ...env, ALLOW_DYNAMIC_PUBLIC_LINKS: "yes" }));
});

test("dynamic URLs require no registration and download anonymously, never using the app-signed URL", async () => {
    const { state, dependencies } = fixture({ bytes: "PKworkbook-v1" });
    const service = createPreviewService({ ...config, allowDynamic: true, publicLinks: new Set() }, dependencies);
    const dynamicLink = `https://${host}/:x:/g/personal/user/new-token?e=xyz&nav=sheet`;
    const entry = await service.get(dynamicLink);
    assert.equal(entry.bytes.toString(), "PKworkbook-v1");
    const request = state.requests.at(-1);
    const url = new URL(request.url);
    assert.equal(url.pathname, new URL(dynamicLink).pathname);
    assert.equal(url.searchParams.get("e"), "xyz");
    assert.equal(url.searchParams.get("nav"), "sheet");
    assert.equal(url.searchParams.get("download"), "1");
    assert.equal(request.options.credentials, "omit");
    assert.equal(request.options.headers?.Authorization, undefined);
    assert.equal(request.options.headers?.Cookie, undefined);
    assert.ok(!state.requests.some(request => request.url.includes("secret=signed")));
    const count = state.requests.length;
    assert.equal(await service.get(dynamicLink), entry);
    assert.equal(state.requests.length, count);
    state.time += 300001;
    await service.get(dynamicLink);
    assert.equal(state.requests.length, count + 3); // permission, metadata, anonymous public download
});

test("dynamic links fail closed on login/password pages even if Graph says anonymous", async () => {
    for (const overrides of [
        { redirect: "https://login.microsoftonline.com/signin" },
        { bytes: "<html>Password required</html>", mime: "text/html" },
        { bytes: "<html>Password required</html>", mime: "application/octet-stream" },
        { bytes: "not-a-workbook", mime: "application/octet-stream" }
    ]) {
        const { dependencies, state } = fixture(overrides);
        const service = createPreviewService({ ...config, allowDynamic: true, publicLinks: new Set() }, dependencies);
        await assert.rejects(service.get(link), { status: 403 });
        assert.ok(!state.requests.some(request => request.url.includes("secret=signed")));
    }
});

test("dynamic access still rejects non-public permissions and the wrong tenant", async () => {
    for (const permission of [{}, { link: { scope: "organization", type: "view" } },
        { link: { scope: "anonymous", type: "view" }, hasPassword: true }]) {
        const { dependencies, state } = fixture({ permission });
        const service = createPreviewService({ ...config, allowDynamic: true }, dependencies);
        await assert.rejects(service.get(link), { status: 403 });
        assert.equal(state.requests.length, 2);
        await assert.rejects(service.get(link.replace(host, "other.sharepoint.com")), { status: 403 });
        assert.equal(state.requests.length, 2);
    }
});

test("dynamic cache entries are bounded even for many small files", async () => {
    const { state, dependencies } = fixture({ bytes: "PKsmall" });
    const service = createPreviewService({ ...config, allowDynamic: true, maxCacheEntries: 1 }, dependencies);
    await service.get(link);
    await service.get(otherLink);
    const count = state.requests.length;
    await service.get(link);
    assert.equal(state.requests.length, count + 3);
});

test("dynamic cache is evicted if a link starts requiring a password", async () => {
    const { state, dependencies } = fixture({ bytes: "PKworkbook" });
    const service = createPreviewService({ ...config, allowDynamic: true }, dependencies);
    await service.get(link);
    state.time += 300001;
    state.bytes = "<html>Enter password</html>";
    state.mime = "text/html";
    await assert.rejects(service.get(link), { status: 403 });
    await assert.rejects(service.get(link), { status: 403 });
});

test("unsupported permission navigation falls back to the matching file permission", async () => {
    for (const permissionStatus of [400, 405, 501]) {
        const { dependencies, state } = fixture({ permissionStatus, bytes: "PKworkbook", permissions: [
            { link: { scope: "anonymous", type: "view", webUrl: otherLink } },
            { link: { scope: "anonymous", type: "view", webUrl: link.split("?")[0] } }
        ] });
        const service = createPreviewService({ ...config, allowDynamic: true }, dependencies);
        assert.equal((await service.get(link)).bytes.toString(), "PKworkbook");
        assert.ok(state.requests.some(request => request.url === "https://graph.microsoft.com/v1.0/drives/drive/items/item/permissions"));
        assert.ok(!state.requests.some(request => request.url.includes("secret=signed")));
    }
});

test("fallback never substitutes another public link or ignores the matched link's restrictions", async () => {
    for (const permissions of [
        [{ link: { scope: "anonymous", type: "view", webUrl: otherLink } }],
        [{ link: { scope: "anonymous", type: "view", webUrl: link.replace(host, "evil.example") } }],
        [{ link: { scope: "anonymous", type: "view" } }],
        [{ link: { scope: "organization", type: "view", webUrl: link } }, { link: { scope: "anonymous", type: "view", webUrl: otherLink } }],
        [{ link: { scope: "anonymous", type: "view", webUrl: link }, hasPassword: true }],
        [{ link: { scope: "anonymous", type: "view", webUrl: link }, expirationDateTime: "2020-01-01T00:00:00Z" }]
    ]) {
        const { service, state } = fixture({ permissionStatus: 400, permissions });
        await assert.rejects(service.get(link), { status: 403 });
        assert.ok(state.requests.every(request => !request.url.startsWith(`https://${host}/`)));
    }
});

test("permission pagination stays on the exact Graph collection", async () => {
    const base = "https://graph.microsoft.com/v1.0/drives/drive/items/item/permissions";
    const second = base + "?$skiptoken=next";
    const { service } = fixture({ permissionStatus: 400, permissionPages: {
        [base]: { value: [], "@odata.nextLink": second },
        [second]: { value: [{ link: { scope: "anonymous", type: "view", webUrl: link } }] }
    } });
    assert.equal((await service.get(link)).bytes.toString(), "workbook-v1");
    for (const next of ["https://evil.example/permissions", "https://graph.microsoft.com/v1.0/users"]) {
        const bad = fixture({ permissionStatus: 400, permissionPages: { [base]: { value: [], "@odata.nextLink": next } } });
        await assert.rejects(bad.service.get(link), { status: 502 });
        assert.ok(!bad.state.requests.some(request => request.url === next));
    }
});

test("denied or failed permission calls do not trigger the fallback", async () => {
    for (const permissionStatus of [401, 403, 404, 429, 500]) {
        const { service, state } = fixture({ permissionStatus });
        await assert.rejects(service.get(link));
        assert.ok(!state.requests.some(request => request.url.endsWith("/driveItem") || request.url.endsWith("/permissions")));
    }
});

test("Graph errors expose only a safe operation, HTTP status and bounded code", async () => {
    const result = fixture({ status: 500, errorCode: "invalidRequest" });
    await assert.rejects(result.service.get(link), error => {
        assert.match(error.message, /link permission: HTTP 500, invalidRequest/);
        assert.ok(!error.message.includes("private upstream message"));
        assert.ok(!error.message.includes(link));
        return true;
    });
    const unsafe = fixture({ status: 500, errorCode: "https://secret.example/token" });
    await assert.rejects(unsafe.service.get(link), error => {
        assert.match(error.message, /HTTP 500, unknown/);
        assert.ok(!error.message.includes("secret.example"));
        return true;
    });
});
