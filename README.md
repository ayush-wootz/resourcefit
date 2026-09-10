# ResourceFit

ResourceFit previews images, PDFs, web pages, and Microsoft Excel workbooks from a `link` query parameter.

## Microsoft Excel preview setup

Create a single-tenant app registration in Microsoft Entra ID and configure:

- Platform: Single-page application (SPA)
- Redirect URI: `https://ayush-wootz.github.io/resourcefit/auth.html`
- Microsoft Graph delegated permission: `Files.Read`

Put the Directory (tenant) ID and Application (client) ID in `config.js`. These identifiers are public application configuration; do not add a client secret to this repository.

For the best silent sign-in result, pass the signed-in user's company email as `login_hint`:

```text
https://ayush-wootz.github.io/resourcefit/?link=ENCODED_SHAREPOINT_URL&login_hint=ENCODED_USER_EMAIL
```

The viewer tries silent SSO first. If Microsoft requires consent, MFA, or browser cookies are unavailable in the iframe, it shows a **Sign in to preview** button and uses a popup. An **Open in Excel** fallback is always available from status and error screens.

For configuration testing, `tenant_id` and `client_id` can be passed as query parameters. Configure `config.js` before production use so Glide only needs to supply the file URL and optional login hint.
