# ResourceFit

ResourceFit previews images, PDFs, web pages, and Excel workbooks from a `link` query parameter.

## Read-only Excel previews

Excel files are downloaded and parsed in the browser with SheetJS. The viewer shows worksheet tabs and cell values without enabling edits or requiring a Microsoft account.

Pass a URL-encoded, directly downloadable workbook URL:

```text
https://ayush-wootz.github.io/resourcefit/?link=ENCODED_XLSX_URL
```

The file host must permit cross-origin browser downloads. A Glide file-column URL or another CORS-enabled static or signed URL should work. A private SharePoint `Doc.aspx` viewing page is not a direct workbook download and cannot be read by this client-side viewer.

Complex Excel styling is not reproduced exactly. SheetJS Community Edition focuses on worksheet data and number formatting.
