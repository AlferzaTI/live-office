/* ============================================================
   SEGURIDAD - ALFERZA LIVE OFFICE
   Microsoft Graph
   ============================================================ */

const msalConfig = {
    auth: {
        clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",
        authority:
            "https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1",
        redirectUri:
            "https://alferzati.github.io/live-office/blank.html"
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
    }
};

const scopes = [
    "User.Read",
    "User.Read.All",
    "Presence.Read.All",
    "Sites.Read.All",
    "AuditLog.Read.All",
    "DeviceManagementManagedDevices.Read.All"
];

const msalInstance =
    new msal.PublicClientApplication(msalConfig);


/* ============================================================
   VARIABLES
   ============================================================ */

let currentAccount = null;

let securityData = {
    users: [],
    inactiveUsers: [],
    blockedUsers: [],
    mfaUsers: [],
    devices: [],
    signIns: []
};


/* ============================================================
   INICIO
   ============================================================ */

document.addEventListener("DOMContentLoaded", async () => {

    try {

        showLoading();

        await initializeAuthentication();

        await loadSecurityData();

        renderDashboard();

        hideLoading();

    } catch (error) {

        console.error("Error cargando seguridad:", error);

        hideLoading();

        showError(
            "No se pudieron cargar los datos de seguridad."
        );
    }

});


/* ============================================================
   AUTENTICACIÓN
   ============================================================ */

async function initializeAuthentication() {

    const response = await msalInstance.handleRedirectPromise();

    if (response?.account) {
        currentAccount = response.account;
        msalInstance.setActiveAccount(currentAccount);
    }

    const accounts = msalInstance.getAllAccounts();

    if (!currentAccount && accounts.length > 0) {
        currentAccount = accounts[0];
        msalInstance.setActiveAccount(currentAccount);
    }

    if (!currentAccount) {

        window.location.href = "login.html";

        throw new Error("Usuario no autenticado.");
    }
}


/* ============================================================
   OBTENER TOKEN
   ============================================================ */

async function getAccessToken() {

    try {

        const response =
            await msalInstance.acquireTokenSilent({
                scopes,
                account: currentAccount
            });

        return response.accessToken;

    } catch (error) {

        console.warn(
            "No se pudo obtener token silenciosamente.",
            error
        );

        const response =
            await msalInstance.acquireTokenPopup({
                scopes
            });

        return response.accessToken;
    }
}


/* ============================================================
   PETICIÓN GRAPH
   ============================================================ */

async function graphRequest(endpoint) {

    const token = await getAccessToken();

    const response = await fetch(
        `https://graph.microsoft.com/v1.0${endpoint}`,
        {
            method: "GET",

            headers: {
                Authorization: `Bearer ${token}`,
                "Content-Type": "application/json"
            }
        }
    );

    if (!response.ok) {

        const errorText = await response.text();

        console.error(
            "Microsoft Graph:",
            response.status,
            errorText
        );

        throw new Error(
            `Graph API ${response.status}`
        );
    }

    return response.json();
}


/* ============================================================
   CARGAR TODOS LOS DATOS
   ============================================================ */

async function loadSecurityData() {

    console.log(
        "Cargando información de seguridad..."
    );

    await Promise.all([
        loadUsers(),
        loadMfa(),
        loadDevices(),
        loadSignIns()
    ]);

    console.log(
        "Información de seguridad cargada."
    );
}


/* ============================================================
   USUARIOS
   ============================================================ */

async function loadUsers() {

    let users = [];

    let url =
        "/users?$select=id,displayName,userPrincipalName,accountEnabled,department,signInActivity";

    while (url) {

        const data = await graphRequest(
            url.replace(
                "https://graph.microsoft.com/v1.0",
                ""
            )
        );

        if (data.value) {
            users.push(...data.value);
        }

        url = data["@odata.nextLink"] || null;
    }

    securityData.users = users;

    securityData.blockedUsers =
        users.filter(
            user => user.accountEnabled === false
        );

    securityData.inactiveUsers =
        users.filter(user =>
            isInactive(user.signInActivity)
        );

    console.log(
        "Usuarios:",
        users.length
    );

    console.log(
        "Bloqueados:",
        securityData.blockedUsers.length
    );

    console.log(
        "Inactivos:",
        securityData.inactiveUsers.length
    );
}


/* ============================================================
   DETECTAR USUARIOS INACTIVOS
   ============================================================ */

function isInactive(signInActivity) {

    if (!signInActivity) {
        return true;
    }

    const lastLogin =
        signInActivity.lastSignInDateTime;

    if (!lastLogin) {
        return true;
    }

    const lastDate =
        new Date(lastLogin);

    const now =
        new Date();

    const difference =
        now - lastDate;

    const days =
        difference /
        (1000 * 60 * 60 * 24);

    return days >= 30;
}


/* ============================================================
   DÍAS DE INACTIVIDAD
   ============================================================ */

function getInactiveDays(signInActivity) {

    if (
        !signInActivity ||
        !signInActivity.lastSignInDateTime
    ) {
        return null;
    }

    const lastDate =
        new Date(
            signInActivity.lastSignInDateTime
        );

    const now =
        new Date();

    return Math.floor(
        (now - lastDate) /
        (1000 * 60 * 60 * 24)
    );
}


/* ============================================================
   MFA
   ============================================================ */

async function loadMfa() {

    const token =
        await getAccessToken();

    const response =
        await fetch(
            "https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails",
            {
                headers: {
                    Authorization:
                        `Bearer ${token}`
                }
            }
        );

    if (!response.ok) {

        const error =
            await response.text();

        console.error(
            "Error MFA:",
            error
        );

        throw new Error(
            "No se pudo obtener información MFA."
        );
    }

    let users = [];

    let data =
        await response.json();

    users.push(
        ...(data.value || [])
    );

    while (data["@odata.nextLink"]) {

        const nextResponse =
            await fetch(
                data["@odata.nextLink"],
                {
                    headers: {
                        Authorization:
                            `Bearer ${token}`
                    }
                }
            );

        data =
            await nextResponse.json();

        users.push(
            ...(data.value || [])
        );
    }

    securityData.mfaUsers =
        users;

    console.log(
        "Registros MFA:",
        users.length
    );
}


/* ============================================================
   DISPOSITIVOS INTUNE
   ============================================================ */

async function loadDevices() {

    const token =
        await getAccessToken();

    const response =
        await fetch(
            "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices",
            {
                headers: {
                    Authorization:
                        `Bearer ${token}`
                }
            }
        );

    if (!response.ok) {

        const error =
            await response.text();

        console.error(
            "Error dispositivos:",
            error
        );

        throw new Error(
            "No se pudieron obtener los dispositivos."
        );
    }

    const data =
        await response.json();

    securityData.devices =
        data.value || [];

    console.log(
        "Dispositivos:",
        securityData.devices.length
    );
}


/* ============================================================
   SIGN-INS
   ============================================================ */

async function loadSignIns() {

    const token =
        await getAccessToken();

    const url =
        "https://graph.microsoft.com/v1.0/auditLogs/signIns?$top=50&$orderby=createdDateTime desc";

    const response =
        await fetch(
            url,
            {
                headers: {
                    Authorization:
                        `Bearer ${token}`
                }
            }
        );

    if (!response.ok) {

        const error =
            await response.text();

        console.error(
            "Error sign-ins:",
            error
        );

        throw new Error(
            "No se pudieron obtener los inicios de sesión."
        );
    }

    const data =
        await response.json();

    securityData.signIns =
        data.value || [];

    console.log(
        "Sign-ins:",
        securityData.signIns.length
    );
}


/* ============================================================
   RENDER DASHBOARD
   ============================================================ */

function renderDashboard() {

    updateCounters();

    renderMfa();

    renderInactiveUsers();

    renderBlockedUsers();

    renderDevices();

    renderSecurityEvents();

}


/* ============================================================
   CONTADORES
   ============================================================ */

function updateCounters() {

    const blocked =
        securityData.blockedUsers.length;

    const inactive =
        securityData.inactiveUsers.length;

    const mfaPending =
        securityData.mfaUsers.filter(
            user =>
                user.isMfaRegistered === false
        ).length;

    const riskyDevices =
        securityData.devices.filter(
            device =>
                device.complianceState &&
                device.complianceState.toLowerCase() !==
                "compliant"
        ).length;


    setElementText(
        "blockedCount",
        blocked
    );

    setElementText(
        "inactiveCount",
        inactive
    );

    setElementText(
        "mfaCount",
        mfaPending
    );

    setElementText(
        "riskDevicesCount",
        riskyDevices
    );
}


/* ============================================================
   MFA
   ============================================================ */

function renderMfa() {

    const total =
        securityData.mfaUsers.length;

    if (!total) {
        return;
    }

    const registered =
        securityData.mfaUsers.filter(
            user =>
                user.isMfaRegistered === true
        ).length;

    const pending =
        total - registered;

    const percentage =
        Math.round(
            (registered / total) * 100
        );

    setElementText(
        "mfaPercentage",
        `${percentage}%`
    );

    setElementText(
        "mfaRegistered",
        registered
    );

    setElementText(
        "mfaPending",
        pending
    );

    const circle =
        document.querySelector(
            ".mfa-circle"
        );

    if (circle) {

        circle.style.setProperty(
            "--percentage",
            `${percentage * 3.6}deg`
        );
    }
}


/* ============================================================
   USUARIOS INACTIVOS
   ============================================================ */

function renderInactiveUsers() {

    const container =
        document.querySelector(
            "#inactiveUsersList"
        );

    if (!container) {
        return;
    }

    container.innerHTML = "";

    const users =
        securityData.inactiveUsers
            .slice(0, 10);

    if (!users.length) {

        container.innerHTML =
            `<p class="empty-state">
                No hay usuarios inactivos.
            </p>`;

        return;
    }

    users.forEach(user => {

        const days =
            getInactiveDays(
                user.signInActivity
            );

        const row =
            document.createElement(
                "div"
            );

        row.className =
            "security-user-row";

        row.innerHTML = `
            <div>
                <strong>
                    ${escapeHtml(
                        user.displayName ||
                        "Sin nombre"
                    )}
                </strong>

                <small>
                    ${escapeHtml(
                        user.userPrincipalName ||
                        ""
                    )}
                </small>
            </div>

            <span class="security-badge warning">
                ${
                    days === null
                        ? "Sin registro"
                        : `${days} días`
                }
            </span>
        `;

        container.appendChild(row);
    });
}


/* ============================================================
   CUENTAS BLOQUEADAS
   ============================================================ */

function renderBlockedUsers() {

    const container =
        document.querySelector(
            "#blockedUsersList"
        );

    if (!container) {
        return;
    }

    container.innerHTML = "";

    const users =
        securityData.blockedUsers
            .slice(0, 10);

    if (!users.length) {

        container.innerHTML =
            `<p class="empty-state">
                No hay cuentas bloqueadas.
            </p>`;

        return;
    }

    users.forEach(user => {

        const row =
            document.createElement(
                "div"
            );

        row.className =
            "security-user-row";

        row.innerHTML = `
            <div>
                <strong>
                    ${escapeHtml(
                        user.displayName ||
                        "Sin nombre"
                    )}
                </strong>

                <small>
                    ${escapeHtml(
                        user.userPrincipalName ||
                        ""
                    )}
                </small>

                ${
                    user.department
                        ? `<small>
                            Área:
                            ${escapeHtml(
                                user.department
                            )}
                           </small>`
                        : ""
                }
            </div>

            <span class="security-badge danger">
                BLOQUEADA
            </span>
        `;

        container.appendChild(row);
    });
}


/* ============================================================
   DISPOSITIVOS
   ============================================================ */

function renderDevices() {

    const container =
        document.querySelector(
            "#devicesList"
        );

    if (!container) {
        return;
    }

    container.innerHTML = "";

    const devices =
        securityData.devices
            .slice(0, 10);

    if (!devices.length) {

        container.innerHTML =
            `<p class="empty-state">
                No hay dispositivos administrados por Intune.
            </p>`;

        return;
    }

    devices.forEach(device => {

        const compliant =
            device.complianceState &&
            device.complianceState
                .toLowerCase() ===
            "compliant";

        const row =
            document.createElement(
                "div"
            );

        row.className =
            "security-device-row";

        row.innerHTML = `
            <div>
                <strong>
                    ${escapeHtml(
                        device.deviceName ||
                        "Sin nombre"
                    )}
                </strong>

                <small>
                    ${escapeHtml(
                        device.operatingSystem ||
                        "Sistema desconocido"
                    )}
                </small>
            </div>

            <span class="security-badge ${
                compliant
                    ? "success"
                    : "danger"
            }">
                ${
                    compliant
                        ? "CUMPLE"
                        : "RIESGO"
                }
            </span>
        `;

        container.appendChild(row);
    });
}


/* ============================================================
   EVENTOS DE SEGURIDAD
   ============================================================ */

function renderSecurityEvents() {

    const container =
        document.querySelector(
            "#securityEventsList"
        );

    if (!container) {
        return;
    }

    container.innerHTML = "";

    const events =
        securityData.signIns
            .slice(0, 15);

    if (!events.length) {

        container.innerHTML =
            `<p class="empty-state">
                No hay eventos recientes.
            </p>`;

        return;
    }

    events.forEach(event => {

        const success =
            event.status &&
            event.status.errorCode === 0;

        const row =
            document.createElement(
                "div"
            );

        row.className =
            "security-event-row";

        row.innerHTML = `

            <div class="event-icon">
                ${
                    success
                        ? "✓"
                        : "⚠️"
                }
            </div>

            <div class="event-info">

                <strong>
                    ${escapeHtml(
                        event.userDisplayName ||
                        event.userPrincipalName ||
                        "Usuario desconocido"
                    )}
                </strong>

                <small>
                    ${
                        event.location?.city ||
                        "Ubicación desconocida"
                    },
                    ${
                        event.location?.countryOrRegion ||
                        ""
                    }
                </small>

                <small>
                    IP:
                    ${
                        event.ipAddress ||
                        "No disponible"
                    }
                </small>

            </div>

            <div class="event-time">
                ${
                    formatDate(
                        event.createdDateTime
                    )
                }
            </div>
        `;

        container.appendChild(row);
    });
}


/* ============================================================
   REFRESCAR
   ============================================================ */

async function refreshSecurity() {

    const button =
        document.querySelector(
            "#refreshSecurity"
        );

    try {

        if (button) {
            button.disabled = true;
            button.textContent =
                "Actualizando...";
        }

        await loadSecurityData();

        renderDashboard();

    } catch (error) {

        console.error(error);

        showError(
            "No se pudieron actualizar los datos."
        );

    } finally {

        if (button) {
            button.disabled = false;
            button.textContent =
                "↻ Actualizar";
        }
    }
}


/* ============================================================
   BOTÓN REFRESCAR
   ============================================================ */

document.addEventListener(
    "click",
    event => {

        if (
            event.target.closest(
                "#refreshSecurity"
            )
        ) {
            refreshSecurity();
        }

    }
);


/* ============================================================
   UTILIDADES
   ============================================================ */

function setElementText(
    id,
    value
) {

    const element =
        document.getElementById(id);

    if (element) {
        element.textContent = value;
    }
}


function formatDate(date) {

    if (!date) {
        return "Sin fecha";
    }

    return new Date(date)
        .toLocaleString(
            "es-PE",
            {
                dateStyle: "short",
                timeStyle: "short"
            }
        );
}


function escapeHtml(value) {

    if (value === null ||
        value === undefined) {

        return "";
    }

    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}


function showLoading() {

    const loading =
        document.getElementById(
            "loadingOverlay"
        );

    if (loading) {
        loading.style.display =
            "flex";
    }
}


function hideLoading() {

    const loading =
        document.getElementById(
            "loadingOverlay"
        );

    if (loading) {
        loading.style.display =
            "none";
    }
}


function showError(message) {

    console.error(message);

    const container =
        document.querySelector(
            ".security-content"
        );

    if (!container) {
        return;
    }

    const error =
        document.createElement(
            "div"
        );

    error.className =
        "security-error";

    error.textContent =
        message;

    container.prepend(error);
}