# Cached SharePoint preview service

This service retrieves file bytes through Microsoft Graph and serves them to the ResourceFit viewer. GitHub Pages continues to host the frontend. GitHub Pages cannot execute this backend: deploy it separately on a Node/container host with HTTPS and outbound access to Microsoft.

The code is disabled in the frontend until `viewer-config.js` contains the deployed backend origin. Merging alone does not activate SharePoint previews.

## 1. Configure Microsoft server access

Use a dedicated Microsoft Entra application registration for this backend. In **API permissions → Add a permission → Microsoft Graph → Application permissions**, Microsoft's documented minimum for the sharing-link API is **Files.ReadWrite.All**. An administrator must grant consent. This is broad tenant access, including write capability, even though this service only performs read requests to Graph. If your administrator cannot approve that permission, do not deploy this sharing-link implementation; a different integration using explicitly selected sites/items is required.

Create a client secret under **Certificates & secrets**, then put its **Value** in your hosting provider's secret environment settings as `MS_CLIENT_SECRET`. Set `MS_TENANT_ID` and `MS_CLIENT_ID` from the backend app registration overview. Never put the secret in GitHub, Glide, `viewer-config.js`, or a URL. The previous SPA registration and its delegated permissions alone are not sufficient. This backend uses client credentials and needs no redirect URI or browser login.

Reference: [Microsoft Graph sharing-link API permissions](https://learn.microsoft.com/en-us/graph/api/shares-get?view=graph-rest-1.0).

## 2. Enable dynamic public links

Set these server environment variables (see `.env.example`):

| Variable | Value |
| --- | --- |
| `MS_TENANT_ID` | Backend app directory/tenant GUID |
| `MS_CLIENT_ID` | Backend app application/client GUID |
| `MS_CLIENT_SECRET` | Secret value, stored only on the backend host |
| `SHAREPOINT_HOST` | `netorgft12597064-my.sharepoint.com` |
| `ALLOW_DYNAMIC_PUBLIC_LINKS` | `true` to accept changing public links from the configured domain |
| `PUBLIC_SHARE_LINKS` | Optional in dynamic mode; omit it or use `[]` |
| `ALLOWED_ORIGINS` | `https://ayush-wootz.github.io` |
| `CACHE_TTL_SECONDS` | `300` by default; may be reduced to 1–300 |
| `MAX_FILE_MB` | `20` by default |
| `MAX_CACHE_MB` | `128` by default, at least `MAX_FILE_MB` |
| `MAX_CACHE_ENTRIES` | `200` by default; limits the number of cached files as well as bytes |
| `MAX_CONCURRENT_FETCHES` | `4` by default |
| `PORT` | Hosting provider's port, default `8080` |

For changing Glide URLs, set `ALLOW_DYNAMIC_PUBLIC_LINKS=true` in Render's Environment settings. Remove `PUBLIC_SHARE_LINKS` or set it to `[]`. Keep the other Microsoft and domain settings. You do not need to update Render when a new file is added: Glide sends each file's current URL in the viewer's `link` parameter.

Get the source URL from **Share → Copy link** with **Anyone** access. It must open without a password in a signed-out browser. Pass the complete URL, including query parameters. A browser's `Doc.aspx` address is not a substitute for an Anyone sharing link.

On each cache miss/revalidation, the service checks the **permission associated with that exact sharing link**, requiring anonymous view/edit access and rejecting expired, password-marked, restricted, missing, or download-blocking permissions. It does not search for a different public link on the same file.

Microsoft does not consistently report password metadata for SharePoint/OneDrive for Business. Dynamic mode therefore also requests the original sharing URL with `download=1`, without Microsoft application tokens or browser cookies. It only serves bytes obtained by that anonymous download; it never substitutes Graph's application-authorized signed download URL. Sign-in redirects, HTML viewer/password pages, unexpected download hosts, or mismatched file sizes fail closed. Some public SharePoint links may not provide a direct anonymous download via this method; they will show an error and an original-link fallback, rather than use private application access. Test your tenant's links after deployment.

Dynamic mode refreshes the public download after the cache window, even for unchanged versions, to reverify actual public access. Repeated opens within the window still reuse cached bytes. CORS is not authentication: anyone holding a supported public sharing link on your configured domain can use the service.

### Optional registered-link mode

The previous behavior remains available by leaving `ALLOW_DYNAMIC_PUBLIC_LINKS` unset or setting it to `false`, and setting `PUBLIC_SHARE_LINKS` to a non-empty JSON array of exact, approved sharing URLs. Example (replace placeholders):

```json
["https://YOUR-TENANT-my.sharepoint.com/:x:/g/personal/USER/TOKEN","https://YOUR-TENANT-my.sharepoint.com/:i:/g/personal/USER/TOKEN"]
```

In registered-link mode, each new file must be added to the list and the service restarted/redeployed. The administrator must verify each registered link is password-free. This mode uses Graph's signed file download URL after registration and permission validation; it can reuse unchanged bytes on revalidation. An empty list fails startup unless dynamic mode is explicitly enabled. In either mode, if the tenant does not expose the link permission to this app, the request fails closed.

References: [sharedDriveItem permission relationship](https://learn.microsoft.com/en-us/graph/api/resources/shareddriveitem?view=graph-rest-1.0), [permission properties](https://learn.microsoft.com/en-us/graph/api/resources/permission?view=graph-rest-1.0).

## 3. Deploy the backend

No npm dependencies are needed. With Node 22 or later, from `backend/`:

```sh
cp .env.example .env
# Fill in .env locally using your editor. It is ignored by git.
node --env-file=.env server.mjs
```

For a container host, set the build context to `backend/` and the Dockerfile to `backend/Dockerfile` (or `Dockerfile` relative to that context). Configure the environment variables in the host, enable HTTPS, and use `/healthz` as the health check. Give the service at least 512 MB RAM with the default cache and concurrency limits; adjust after measuring actual files and traffic.

Local container build/run from the repository root:

```sh
docker build -t resourcefit-sharepoint backend
docker run --rm -p 8080:8080 --env-file backend/.env resourcefit-sharepoint
```

The health check confirms the process is running; it does **not** validate Microsoft consent or file access. Verify `/resolve?link=ENCODED_SHARING_URL` on the deployed host before enabling the viewer. A successful response includes the file name, type and a relative `/content` URL, never a Graph token or Microsoft download URL.

Keep request query strings out of hosting/access logs because sharing links are bearer links. The application itself does not log URLs, secrets, or upstream error bodies. The download host must match `SHAREPOINT_HOST`; an unexpected Microsoft CDN/redirect host fails closed and must be reviewed before adapting the allowlist.

## 4. Enable the existing viewer

Set the deployed HTTPS **origin** in the repository's `viewer-config.js`, then deploy that frontend change:

```js
window.RESOURCEFIT_CONFIG = Object.freeze({
    sharePointApiBase: "https://YOUR-BACKEND-HOST"
});
```

No backend secret belongs in this file. No trailing path, query or token is needed.

Glide still uses the same viewer format:

```text
https://ayush-wootz.github.io/resourcefit/?link=ENCODED_ORIGINAL_SHAREPOINT_SHARING_URL
```

Use Glide's Construct URL query parameter `link` with the original file URL. Let Construct URL encode the parameter once; do not pre-encode it. If using JavaScript, construct it with `new URL(...).searchParams.set("link", originalUrl)`. In Glide's protocol field use `https`, without `://`.

The frontend origin in CORS is `https://ayush-wootz.github.io`, including when the viewer is inside Glide. No Microsoft login iframe or popup is used. Existing non-SharePoint image/PDF/Excel/webpage routes are untouched. Setting `sharePointApiBase` back to an empty string disables the new route.

## Cache behavior and performance

- First open downloads the entire file into the backend's bounded memory cache, then sends bytes to the visitor. `/resolve` warms the cache so the following `/content` request normally does not download again.
- Repeat opens within five minutes reuse bytes without calling Microsoft. Simultaneous requests for the same link share one fetch.
- The first request after expiry checks the link permission and file metadata. Dynamic mode downloads anonymously again to reverify access. In registered-link mode, an unchanged item/version reuses bytes; a changed or missing version tag causes a fresh download.
- Removed access or Microsoft errors after expiry evict the entry and return an error. Stale content is not served on failure. A revoked link can remain usable for the existing cache window, at most five minutes. Already downloaded copies cannot be revoked.
- Least-recently-used files are evicted when the byte budget or entry limit is reached. The cache is per process and resets on restart. Multiple instances each fetch/cache independently; use shared storage only if scale justifies it.
- Browser responses use `Cache-Control: no-store` to avoid a second long-lived authorization cache. Visitors still transfer the file bytes from your backend on each open, so bandwidth grows with file size and opens; caching primarily saves SharePoint downloads and latency. For example, 100 opens of a 5 MB file transfer roughly 500 MB to browsers even if SharePoint was fetched only once.
- Each upstream request has a 30-second timeout; concurrency and file size are bounded. PDF range requests are supported, but the backend still downloads a whole source file on a miss.

SharePoint Excel previews reuse SheetJS's read-only worksheet view; complex Excel formatting is not reproduced exactly. SharePoint PDFs use a lazily loaded, pinned PDF.js library and render one page at a time with Previous/Next controls to avoid Chrome's embedded PDF plugin. Canvas previews have no text selection/search layer; use Open original for those features and accessible document text. Images support PNG, JPEG, GIF, WebP, BMP and AVIF. SVG, HTML, Word, PowerPoint and other formats offer the original link rather than executing file content or promising a conversion.

## Verification

From the repository root:

```sh
node --test backend/test/*.test.mjs
```

Tests use mocked Microsoft responses and real local HTTP requests. They cover cache reuse/revalidation, file changes, revocation, expiry, registration, dynamic links, anonymous download verification, download host validation, memory/concurrency limits, byte ranges, CORS, server-only credentials and existing viewer routing. A real tenant test is still required after credentials and a host are configured: open a public image, PDF and workbook directly and inside Glide, then revoke a test link and confirm access fails after the configured cache window. In dynamic mode, also verify a new link works without an environment update, and a password-protected link is refused even if Graph reports anonymous scope.
