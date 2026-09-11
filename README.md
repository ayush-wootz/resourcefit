# ResourceFit

ResourceFit previews images, PDFs, web pages, and Excel workbooks from a `link` query parameter.

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

Supported SharePoint previews: Excel worksheet data, PDFs with page navigation, and raster images. Other types offer the original link. The backend requires registered, password-free **Anyone** sharing links; it is not a private-file authentication gateway. It uses server-side Microsoft authentication and a bounded five-minute cache, so visitors do not sign into Microsoft in the embed.

Run the automated backend and routing checks with Node 22 or later:

```sh
node --test backend/test/*.test.mjs
```
