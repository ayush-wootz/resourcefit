(function () {
    "use strict";

    const elements = {
        status: document.getElementById("status"),
        webContainer: document.getElementById("webContainer"),
        embeddedLink: document.getElementById("embeddedLink"),
        imageWrapper: document.getElementById("imageWrapper"),
        directImage: document.getElementById("directImage"),
        pdfViewer: document.getElementById("pdfViewer"),
        officeContainer: document.getElementById("officeContainer"),
        officePreview: document.getElementById("officePreview")
    };

    const query = new URLSearchParams(window.location.search);
    const sourceLink = query.get("link");
    const loginHint = query.get("login_hint") || undefined;
    const suppliedTenantId = query.get("tenant_id");
    const suppliedClientId = query.get("client_id");
    const config = window.RESOURCEFIT_CONFIG || {};
    const tenantId = suppliedTenantId || config.tenantId;
    const clientId = suppliedClientId || config.clientId;
    let msalClient;

    function hideAll() {
        elements.status.style.display = "none";
        elements.webContainer.style.display = "none";
        elements.imageWrapper.style.display = "none";
        elements.pdfViewer.style.display = "none";
        elements.officeContainer.style.display = "none";
    }

    function escapeHtml(value) {
        return String(value)
            .replaceAll("&", "&amp;")
            .replaceAll("<", "&lt;")
            .replaceAll(">", "&gt;")
            .replaceAll('"', "&quot;")
            .replaceAll("'", "&#039;");
    }

    function showStatus(title, message, options) {
        const settings = options || {};
        hideAll();
        elements.status.style.display = "flex";
        elements.status.innerHTML = `
            <div class="status-card">
                ${settings.loading ? '<div class="spinner" aria-hidden="true"></div>' : ""}
                <h1>${escapeHtml(title)}</h1>
                <p>${escapeHtml(message)}</p>
                <div class="actions">
                    ${settings.signIn ? '<button id="signInButton" type="button">Sign in to preview</button>' : ""}
                    ${sourceLink ? '<a class="button-link" href="' + escapeHtml(sourceLink) + '" target="_blank" rel="noopener noreferrer">Open in Excel</a>' : ""}
                </div>
            </div>`;

        const signInButton = document.getElementById("signInButton");
        if (signInButton) signInButton.addEventListener("click", signInAndPreview);
    }

    function getExtension(url) {
        try { return new URL(url).pathname.split(".").pop().toLowerCase(); }
        catch (_) { return ""; }
    }

    function detectContentType(url) {
        const extension = getExtension(url);
        const imageExtensions = ["jpg", "jpeg", "png", "gif", "bmp", "webp", "svg"];
        const documentExtensions = ["pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx"];
        if (imageExtensions.includes(extension)) return "image";
        if (documentExtensions.includes(extension)) return "document";
        return "webpage";
    }

    function isMicrosoftWorkbook(url) {
        try {
            const parsed = new URL(url);
            const microsoftHost = parsed.hostname.endsWith(".sharepoint.com") ||
                parsed.hostname === "onedrive.live.com" || parsed.hostname.endsWith(".onedrive.com");
            const workbookPath = /\.xls(x|m|b)?$/i.test(parsed.pathname) || /\/\:x\:\//i.test(parsed.pathname) ||
                /\/doc2\.aspx$/i.test(parsed.pathname) || /sourcedoc=/i.test(parsed.search);
            return microsoftHost && workbookPath;
        } catch (_) { return false; }
    }

    function toShareId(url) {
        const bytes = new TextEncoder().encode(url);
        let binary = "";
        bytes.forEach((byte) => { binary += String.fromCharCode(byte); });
        return "u!" + btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
    }

    function showImage(url) {
        hideAll();
        elements.imageWrapper.style.display = "block";
        elements.directImage.src = url;
    }

    function showPdf(url) {
        hideAll();
        elements.pdfViewer.style.display = "block";
        elements.pdfViewer.src = url;
    }

    function showWebpage(url) {
        hideAll();
        elements.webContainer.style.display = "block";
        elements.embeddedLink.src = url;
    }

    function showOfficePreview(preview) {
        hideAll();
        elements.officeContainer.style.display = "block";
        if (preview.getUrl) {
            elements.officePreview.src = preview.getUrl;
            return;
        }
        if (preview.postUrl && preview.postParameters) {
            const form = document.createElement("form");
            form.method = "post";
            form.action = preview.postUrl;
            form.target = elements.officePreview.name;
            form.hidden = true;
            new URLSearchParams(preview.postParameters).forEach((value, key) => {
                const input = document.createElement("input");
                input.type = "hidden";
                input.name = key;
                input.value = value;
                form.appendChild(input);
            });
            document.body.appendChild(form);
            form.submit();
            form.remove();
            return;
        }
        throw new Error("Microsoft did not return an embeddable preview URL.");
    }

    async function initializeMsal() {
        if (!tenantId || !clientId) throw new Error("Microsoft preview has not been configured yet.");
        if (!window.msal) throw new Error("Microsoft sign-in could not be loaded.");
        msalClient = new window.msal.PublicClientApplication({
            auth: {
                clientId,
                authority: `https://login.microsoftonline.com/${tenantId}`,
                redirectUri: new URL("./auth.html", window.location.href).href,
                navigateToLoginRequestUrl: false
            },
            cache: { cacheLocation: "localStorage", storeAuthStateInCookie: true }
        });
        if (typeof msalClient.initialize === "function") await msalClient.initialize();
    }

    async function getToken(interactive) {
        const scopes = ["Files.Read"];
        let account = msalClient.getAllAccounts()[0];
        if (!account && !interactive) {
            const response = await msalClient.ssoSilent({ scopes, loginHint });
            return response.accessToken;
        }
        if (!account && interactive) {
            const response = await msalClient.loginPopup({ scopes, loginHint });
            return response.accessToken;
        }
        try {
            const response = await msalClient.acquireTokenSilent({ scopes, account });
            return response.accessToken;
        } catch (error) {
            if (!interactive) throw error;
            const response = await msalClient.acquireTokenPopup({ scopes, account });
            return response.accessToken;
        }
    }

    async function requestPreview(accessToken) {
        const response = await fetch(`https://graph.microsoft.com/v1.0/shares/${toShareId(sourceLink)}/driveItem/preview`, {
            method: "POST",
            headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" },
            body: "{}"
        });
        if (!response.ok) {
            let detail = "";
            try {
                const data = await response.json();
                detail = data.error && data.error.message ? ` ${data.error.message}` : "";
            } catch (_) {}
            throw new Error(`Microsoft could not create the preview.${detail}`);
        }
        return response.json();
    }

    async function loadMicrosoftPreview(interactive) {
        showStatus("Loading Excel preview", "Checking your Microsoft access…", { loading: true });
        try {
            if (!msalClient) await initializeMsal();
            const preview = await requestPreview(await getToken(interactive));
            showOfficePreview(preview);
        } catch (error) {
            if (!tenantId || !clientId) {
                showStatus(
                    "Microsoft preview setup required",
                    "The Microsoft application and tenant identifiers still need to be configured."
                );
                return;
            }
            if (!interactive) {
                showStatus("Microsoft sign-in required", "Sign in once to verify your access to this workbook.", { signIn: true });
                return;
            }
            showStatus("Preview unavailable", error && error.message ? error.message : "Open the workbook in Excel to continue.");
        }
    }

    async function signInAndPreview() { await loadMicrosoftPreview(true); }

    function initialize() {
        if (!sourceLink) {
            showStatus("No file selected", "Provide a link parameter to preview a file.");
            return;
        }
        try { new URL(sourceLink); }
        catch (_) {
            showStatus("Invalid link", "The supplied file link is not a valid URL.");
            return;
        }
        if (isMicrosoftWorkbook(sourceLink)) {
            loadMicrosoftPreview(false);
            return;
        }
        switch (detectContentType(sourceLink)) {
            case "image": showImage(sourceLink); break;
            case "document": showPdf(sourceLink); break;
            default: showWebpage(sourceLink);
        }
    }

    initialize();
}());
