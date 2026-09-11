# ResourceFit

ResourceFit previews images, PDFs, web pages, Excel workbooks, and Google Drive files from a `link` query parameter.

## Google Drive previews

Pass a Google Drive file id, a Drive share link, or a Google Docs/Sheets/Slides link:

```text
https://ayush-wootz.github.io/resourcefit/?link=1H8e8ldZDcvKo5RRi89b49vaDYJeE0Nb0
https://ayush-wootz.github.io/resourcefit/?link=ENCODED_DRIVE_SHARE_LINK
```

Drive renders the file itself, so uploaded Excel workbooks, PDFs, Word, PowerPoint, and images
all preview without a CORS download, SheetJS, or the backend. Recognised inputs are a bare id,
`/file/d/<id>/view`, `open?id=<id>`, `uc?id=<id>`, and `docs.google.com/<spreadsheets|document|presentation>/d/<id>/...`.

Two requirements:

- Sharing must be **Anyone with the link → Viewer**. Otherwise the embed shows a sign-in wall,
  which usually fails inside a webview even for people who do have access.
- The viewer always embeds `/preview`, never `/view`. Drive sends `X-Frame-Options: SAMEORIGIN`
  on `/view`, so that URL renders blank in an iframe.

Comma-separate ids to show several files as tabs, with optional labels:

```text
?link=ID_ONE,ID_TWO,ID_THREE&names=Quote,Drawing,Spec
```

Ids may also be passed as `driveId` instead of `link`. Mixed input (a Drive id plus a non-Drive
URL) falls through to the existing routes rather than rendering half of it.

Drive workbooks use Google's viewer rather than the SheetJS worksheet tabs, because
`drive.google.com` download URLs send no CORS headers. Serving Drive bytes through the backend
to reuse the SheetJS renderer would be a separate change.

## Zoom and fullscreen

A floating control bar offers fullscreen and, where the page renders the content
itself, zoom. Fullscreen uses the Fullscreen API on the whole viewer, so it also
covers the Drive and SharePoint frames. Hosts that embed this page without
`allow="fullscreen"` reject that call, and the button then expands the viewer
within the embed instead; the adjacent button opens the viewer in a new tab.

Pinch, drag, double-tap, and ctrl+wheel zoom apply to Excel worksheets and images.
They do not apply to the Google Drive, SharePoint, PDF, or web frames: touch events
inside a cross-origin iframe never reach this page, so those viewers own their own
gestures and cannot be extended from here. The zoom buttons hide on those routes
rather than appearing inert.

## Read-only Excel previews

Excel files are downloaded and parsed in the browser with SheetJS. The viewer shows worksheet tabs and cell values without enabling edits or requiring a Microsoft account.

Pass a URL-encoded, directly downloadable workbook URL:

```text
https://ayush-wootz.github.io/resourcefit/?link=ENCODED_XLSX_URL
```

The file host must permit cross-origin browser downloads. A Glide file-column URL or another CORS-enabled static or signed URL should work. A SharePoint `Doc.aspx` viewing page is not a direct workbook download. Public SharePoint sharing links can use the optional backend below.

Complex Excel styling is not reproduced exactly. SheetJS Community Edition focuses on worksheet data and number formatting.

## Optional cached SharePoint previews

The [backend setup guide](backend/README.md) explains how to deploy the Node service and connect it using `viewer-config.js`. Existing direct-file routes are unchanged; the additional SharePoint route activates only when `sharePointApiBase` is configured.

Supported SharePoint previews: Excel worksheet data, PDFs with page navigation, and raster images. Other types offer the original link. Set `ALLOW_DYNAMIC_PUBLIC_LINKS=true` for changing Glide URLs from your configured SharePoint domain; no per-file list is required. The backend checks the link's **Anyone** permission and verifies an anonymous file download. Links requiring sign-in or a password are rejected. It uses server-side Microsoft authentication and a bounded five-minute cache, so visitors do not sign into Microsoft in the embed. The original registered-link mode remains available.

Run the automated backend and routing checks with Node 22 or later:

```sh
node --test backend/test/*.test.mjs
```
