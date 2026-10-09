
const MSAL_CONFIG = {
    auth: {
        clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",
        authority:
            "https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1",
        redirectUri: window.location.hostname === "localhost"
            ? "http://localhost:5500/blank.html"
            : "https://alferzati.github.io/live-office/blank.html"
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
    }
};

const LOGIN_REQUEST = {
    scopes: [
        "User.Read",
        "Presence.Read.All",
        "Sites.Read.All"
    ]
};