/* ============================================================
   ALFERZA LIVE OFFICE
   SEGURIDAD - MICROSOFT GRAPH
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


/* ============================================================
   PERMISOS
   ============================================================ */

const scopes = [
    "User.Read",
    "User.Read.All",
    "Presence.Read.All",
    "Sites.Read.All",
    "AuditLog.Read.All",
    "DeviceManagementManagedDevices.Read.All"
];


/* ============================================================
   MSAL
   ============================================================ */

const msalInstance =
    new msal.PublicClientApplication(msalConfig);


/* ============================================================
   DATOS
   ============================================================ */

const securityData = {

    users: [],

    inactiveUsers: [],

    blockedUsers: [],

    mfaUsers: [],

    devices: [],

    signIns: [],

    auditLogs: []

};


/* ============================================================
   CUANDO CARGA LA PÁGINA
   ============================================================ */

document.addEventListener(
    "DOMContentLoaded",
    async () => {

        try {

            await iniciarSesion();

            await cargarSeguridad();

            actualizarDashboard();

        } catch (error) {

            console.error(
                "ERROR SEGURIDAD:",
                error
            );

            mostrarError(
                obtenerMensajeError(error)
            );
        }

    }
);


/* ============================================================
   AUTENTICACIÓN
   ============================================================ */

async function iniciarSesion() {

    const redirectResponse =
        await msalInstance.handleRedirectPromise();


    if (redirectResponse?.account) {

        msalInstance.setActiveAccount(
            redirectResponse.account
        );
    }


    let accounts =
        msalInstance.getAllAccounts();


    if (!accounts.length) {

        /*
         * Si el login de tu aplicación ya dejó
         * una sesión activa, debería aparecer aquí.
         */

        window.location.replace(
            "login.html"
        );

        throw new Error(
            "No existe una sesión de Microsoft 365."
        );
    }


    const account =
        msalInstance.getActiveAccount()
        || accounts[0];


    msalInstance.setActiveAccount(
        account
    );


    console.log(
        "Usuario autenticado:",
        account.username
    );
}


/* ============================================================
   TOKEN
   ============================================================ */

async function obtenerToken() {

    const account =
        msalInstance.getActiveAccount();


    if (!account) {

        throw new Error(
            "No hay una cuenta Microsoft autenticada."
        );
    }


    try {

        const response =
            await msalInstance.acquireTokenSilent({

                scopes: scopes,

                account: account

            });


        return response.accessToken;

    } catch (error) {

        console.warn(
            "Silent token falló. Solicitando token nuevamente..."
        );


        const response =
            await msalInstance.acquireTokenPopup({

                scopes: scopes

            });


        return response.accessToken;
    }
}


/* ============================================================
   PETICIÓN GRAPH
   ============================================================ */

async function graph(
    endpoint,
    beta = false
) {

    const token =
        await obtenerToken();


    const base =
        beta
            ? "https://graph.microsoft.com/beta"
            : "https://graph.microsoft.com/v1.0";


    const url =
        endpoint.startsWith("http")
            ? endpoint
            : `${base}${endpoint}`;


    const response =
        await fetch(
            url,
            {

                method: "GET",

                headers: {

                    Authorization:
                        `Bearer ${token}`,

                    "Content-Type":
                        "application/json"

                }

            }
        );


    if (!response.ok) {

        const text =
            await response.text();


        console.error(
            "GRAPH ERROR",
            response.status,
            text
        );


        throw new Error(
            `Microsoft Graph respondió ${response.status}`
        );
    }


    return response.json();
}


/* ============================================================
   CARGAR TODO
   ============================================================ */

async function cargarSeguridad() {

    mostrarEstadoCarga(
        "Consultando Microsoft 365..."
    );


    await Promise.all([

        cargarUsuarios(),

        cargarMFA(),

        cargarDispositivos(),

        cargarSignIns(),

        cargarAuditoria()

    ]);


    console.log(
        "Datos de seguridad cargados:",
        securityData
    );
}


/* ============================================================
   USUARIOS
   ============================================================ */

async function cargarUsuarios() {

    let url =
        "/users?$select=id,displayName,userPrincipalName,accountEnabled,department,signInActivity";


    let usuarios = [];


    while (url) {

        const data =
            await graph(url, true);


        usuarios.push(
            ...(data.value || [])
        );


        url =
            data["@odata.nextLink"]
            || null;
    }


    securityData.users =
        usuarios;


    securityData.blockedUsers =
        usuarios.filter(
            user =>
                user.accountEnabled === false
        );


    securityData.inactiveUsers =
        usuarios.filter(
            user =>
                obtenerDiasInactividad(
                    user.signInActivity
                ) >= 30
        );


    console.log(
        "Usuarios:",
        usuarios.length
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
   INACTIVIDAD
   ============================================================ */

function obtenerDiasInactividad(
    signInActivity
) {

    if (
        !signInActivity ||
        !signInActivity.lastSignInDateTime
    ) {

        return 9999;
    }


    const ultimoAcceso =
        new Date(
            signInActivity.lastSignInDateTime
        );


    const ahora =
        new Date();


    const diferencia =
        ahora.getTime()
        -
        ultimoAcceso.getTime();


    return Math.floor(
        diferencia /
        (
            1000 *
            60 *
            60 *
            24
        )
    );
}


/* ============================================================
   MFA
   ============================================================ */

async function cargarMFA() {

    let url =
        "/reports/authenticationMethods/userRegistrationDetails";


    let registros = [];


    while (url) {

        const data =
            await graph(url);


        registros.push(
            ...(data.value || [])
        );


        url =
            data["@odata.nextLink"]
            || null;
    }


    securityData.mfaUsers =
        registros;


    console.log(
        "Usuarios MFA:",
        registros.length
    );
}


/* ============================================================
   DISPOSITIVOS INTUNE
   ============================================================ */

async function cargarDispositivos() {

    let url =
        "/deviceManagement/managedDevices";


    let dispositivos = [];


    while (url) {

        const data =
            await graph(url);


        dispositivos.push(
            ...(data.value || [])
        );


        url =
            data["@odata.nextLink"]
            || null;
    }


    securityData.devices =
        dispositivos;


    console.log(
        "Dispositivos Intune:",
        dispositivos.length
    );
}


/* ============================================================
   SIGN INS
   ============================================================ */

async function cargarSignIns() {

    const url =
        "/auditLogs/signIns?$top=100&$orderby=createdDateTime desc";


    const data =
        await graph(url);


    securityData.signIns =
        data.value || [];


    console.log(
        "Sign-ins:",
        securityData.signIns.length
    );
}


/* ============================================================
   AUDITORÍA
   ============================================================

   Buscamos eventos relacionados con:
   - Disable account
   - Update user
   - Account
   - User
   ============================================================ */

async function cargarAuditoria() {

    try {

        const fecha =
            new Date();


        fecha.setDate(
            fecha.getDate() - 180
        );


        const iso =
            fecha.toISOString();


        const filter =
            encodeURIComponent(
                `activityDateTime ge ${iso}`
            );


        const url =
            `/auditLogs/directoryAudits?$filter=${filter}&$top=1000`;


        const data =
            await graph(url);


        securityData.auditLogs =
            data.value || [];


        console.log(
            "Eventos de auditoría:",
            securityData.auditLogs.length
        );

    } catch (error) {

        console.warn(
            "No se pudieron cargar los audit logs:",
            error
        );


        securityData.auditLogs = [];
    }
}


/* ============================================================
   ACTUALIZAR DASHBOARD
   ============================================================ */

function actualizarDashboard() {

    actualizarContadores();

    actualizarMFA();

    actualizarInactivos();

    actualizarBloqueados();

    actualizarDispositivos();

    actualizarEventos();

}


/* ============================================================
   CONTADORES
   ============================================================ */

function actualizarContadores() {


    /* MFA */

    const mfaPendiente =
        securityData.mfaUsers.filter(
            user =>
                user.isMfaRegistered === false
        ).length;


    /* Inactivos */

    const inactivos =
        securityData.inactiveUsers.length;


    /* Bloqueados */

    const bloqueados =
        securityData.blockedUsers.length;


    /* Dispositivos */

    const riesgo =
        securityData.devices.filter(
            device =>
                !esDispositivoSeguro(
                    device
                )
        ).length;


    ponerTexto(
        "mfaPendiente",
        mfaPendiente
    );


    ponerTexto(
        "usuariosInactivos",
        inactivos
    );


    ponerTexto(
        "cuentasBloqueadas",
        bloqueados
    );


    ponerTexto(
        "equiposRiesgo",
        riesgo
    );


    /* Contador de cuentas bloqueadas */

    const counter =
        document.querySelector(
            ".blocked-list"
        )
        ?.parentElement
        ?.querySelector(
            ".counter"
        );


    if (counter) {

        counter.textContent =
            bloqueados;
    }
}


/* ============================================================
   MFA
   ============================================================ */

function actualizarMFA() {

    const total =
        securityData.mfaUsers.length;


    if (!total) {

        return;
    }


    const registrados =
        securityData.mfaUsers.filter(
            user =>
                user.isMfaRegistered === true
        ).length;


    const pendientes =
        total - registrados;


    const porcentaje =
        Math.round(
            (
                registrados /
                total
            ) * 100
        );


    const circle =
        document.querySelector(
            ".mfa-circle"
        );


    if (circle) {

        circle.style.background =
            `conic-gradient(
                #16a34a ${porcentaje}%,
                #ef4444 ${porcentaje}% 100%
            )`;
    }


    const number =
        document.querySelector(
            ".mfa-number"
        );


    if (number) {

        number.textContent =
            `${porcentaje}%`;
    }


    const textos =
        document.querySelectorAll(
            ".mfa-info .legend-item small"
        );


    if (textos.length >= 2) {

        textos[0].textContent =
            `${porcentaje}% de usuarios`;


        textos[1].textContent =
            `${100 - porcentaje}% de usuarios`;
    }
}


/* ============================================================
   USUARIOS INACTIVOS
   ============================================================ */

function actualizarInactivos() {

    const tbody =
        document.getElementById(
            "inactiveUsersTable"
        );


    if (!tbody) {

        return;
    }


    const filtro =
        parseInt(
            document.getElementById(
                "inactiveFilter"
            )?.value || "30"
        );


    const usuarios =
        securityData.inactiveUsers
            .filter(
                user =>
                    obtenerDiasInactividad(
                        user.signInActivity
                    ) >= filtro
            )
            .sort(
                (
                    a,
                    b
                ) =>
                    obtenerDiasInactividad(
                        b.signInActivity
                    )
                    -
                    obtenerDiasInactividad(
                        a.signInActivity
                    )
            );


    tbody.innerHTML = "";


    if (!usuarios.length) {

        tbody.innerHTML = `

            <tr>

                <td
                    colspan="4"
                    style="text-align:center;padding:25px"
                >
                    No hay usuarios con más de
                    ${filtro} días de inactividad.
                </td>

            </tr>

        `;

        return;
    }


    usuarios
        .slice(0, 50)
        .forEach(
            user => {

                const dias =
                    obtenerDiasInactividad(
                        user.signInActivity
                    );


                const nombre =
                    user.displayName
                    ||
                    "Sin nombre";


                const correo =
                    user.userPrincipalName
                    ||
                    "";


                const iniciales =
                    obtenerIniciales(
                        nombre
                    );


                const area =
                    user.department
                    ||
                    "Sin área";


                const ultimoAcceso =
                    user.signInActivity
                    ?.lastSignInDateTime;


                const fecha =
                    ultimoAcceso
                        ? formatearFecha(
                            ultimoAcceso
                        )
                        : "Nunca registrado";


                const clase =
                    dias >= 90
                        ? "danger"
                        : "warning";


                const tr =
                    document.createElement(
                        "tr"
                    );


                tr.innerHTML = `

                    <td>

                        <div class="user-cell">

                            <div class="avatar">
                                ${escapar(iniciales)}
                            </div>

                            <div>

                                <strong>
                                    ${escapar(nombre)}
                                </strong>

                                <small>
                                    ${escapar(correo)}
                                </small>

                            </div>

                        </div>

                    </td>

                    <td>
                        ${escapar(area)}
                    </td>

                    <td>
                        ${fecha}
                    </td>

                    <td>

                        <span class="badge ${clase}">
                            ${dias >= 9999
                                ? "Sin acceso"
                                : `${dias} días`
                            }
                        </span>

                    </td>
                `;


                tbody.appendChild(
                    tr
                );
            }
        );
}


/* ============================================================
   SELECT DE INACTIVOS
   ============================================================ */

document.addEventListener(
    "change",
    event => {

        if (
            event.target.id ===
            "inactiveFilter"
        ) {

            actualizarInactivos();
        }

    }
);


/* ============================================================
   CUENTAS BLOQUEADAS
   ============================================================ */

function actualizarBloqueados() {

    const container =
        document.querySelector(
            ".blocked-list"
        );


    if (!container) {

        return;
    }


    container.innerHTML = "";


    if (
        !securityData.blockedUsers.length
    ) {

        container.innerHTML = `

            <div
                style="
                    padding:20px;
                    text-align:center;
                    color:#64748b;
                "
            >
                No hay cuentas bloqueadas.
            </div>

        `;

        return;
    }


    securityData.blockedUsers
        .slice(0, 20)
        .forEach(
            user => {

                const nombre =
                    user.displayName
                    ||
                    "Sin nombre";


                const area =
                    user.department
                    ||
                    "Sin área";


                const iniciales =
                    obtenerIniciales(
                        nombre
                    );


                const fechaBloqueo =
                    obtenerFechaBloqueo(
                        user
                    );


                const div =
                    document.createElement(
                        "div"
                    );


                div.className =
                    "blocked-user";


                div.innerHTML = `

                    <div class="avatar red-avatar">
                        ${escapar(iniciales)}
                    </div>

                    <div class="blocked-info">

                        <strong>
                            ${escapar(nombre)}
                        </strong>

                        <small>
                            ${escapar(area)}
                        </small>

                        <small>
                            ${fechaBloqueo}
                        </small>

                    </div>

                    <span>
                        Bloqueada
                    </span>
                `;


                container.appendChild(
                    div
                );
            }
        );
}


/* ============================================================
   FECHA DE BLOQUEO
   ============================================================ */

function obtenerFechaBloqueo(
    user
) {

    const nombre =
        (
            user.displayName
            ||
            ""
        )
        .toLowerCase();


    const correo =
        (
            user.userPrincipalName
            ||
            ""
        )
        .toLowerCase();


    const eventos =
        securityData.auditLogs
            .filter(
                event => {

                    const activity =
                        (
                            event.activityDisplayName
                            ||
                            ""
                        )
                        .toLowerCase();


                    const targets =
                        event.targetResources
                        ||
                        [];


                    const coincideUsuario =
                        targets.some(
                            target => {

                                const targetName =
                                    (
                                        target.userPrincipalName
                                        ||
                                        target.displayName
                                        ||
                                        ""
                                    )
                                    .toLowerCase();


                                return (
                                    targetName === correo
                                    ||
                                    targetName === nombre
                                );
                            }
                        );


                    const esRelacionado =
                        activity.includes(
                            "disable"
                        )
                        ||
                        activity.includes(
                            "account"
                        );


                    return (
                        coincideUsuario
                        &&
                        esRelacionado
                    );
                }
            )
            .sort(
                (
                    a,
                    b
                ) =>
                    new Date(
                        b.activityDateTime
                    )
                    -
                    new Date(
                        a.activityDateTime
                    )
            );


    if (!eventos.length) {

        return "Fecha no disponible";
    }


    return `
        Bloqueada:
        ${formatearFecha(
            eventos[0].activityDateTime
        )}
    `;
}


/* ============================================================
   DISPOSITIVOS
   ============================================================ */

function actualizarDispositivos() {

    const container =
        document.querySelector(
            ".device-list"
        );


    if (!container) {

        return;
    }


    container.innerHTML = "";


    if (!securityData.devices.length) {

        container.innerHTML = `

            <div
                style="
                    padding:20px;
                    text-align:center;
                    color:#64748b;
                "
            >
                No hay dispositivos administrados por Intune.
            </div>

        `;

        return;
    }


    securityData.devices
        .slice(0, 30)
        .forEach(
            device => {

                const seguro =
                    esDispositivoSeguro(
                        device
                    );


                const nombre =
                    device.deviceName
                    ||
                    "Sin nombre";


                const sistema =
                    device.operatingSystem
                    ||
                    "Sistema desconocido";


                const estado =
                    device.complianceState
                    ||
                    "unknown";


                const div =
                    document.createElement(
                        "div"
                    );


                div.className =
                    "device-row";


                div.innerHTML = `

                    <div class="device-left">

                        <div class="device-icon">
                            ${
                                sistema
                                    .toLowerCase()
                                    .includes("windows")
                                    ? "PC"
                                    : "DV"
                            }
                        </div>

                        <div>

                            <strong>
                                ${escapar(nombre)}
                            </strong>

                            <small>
                                ${escapar(sistema)}
                            </small>

                        </div>

                    </div>

                    <span
                        class="device-status ${
                            seguro
                                ? "good"
                                : "danger"
                        }"
                    >
                        ${
                            seguro
                                ? "Protegido"
                                : "Riesgo"
                        }
                    </span>

                `;


                container.appendChild(
                    div
                );
            }
        );
}


/* ============================================================
   ESTADO DEL DISPOSITIVO
   ============================================================ */

function esDispositivoSeguro(
    device
) {

    const estado =
        (
            device.complianceState
            ||
            ""
        )
        .toLowerCase();


    if (
        estado ===
        "compliant"
    ) {

        return true;
    }


    return false;
}


/* ============================================================
   EVENTOS DE SEGURIDAD
   ============================================================ */

function actualizarEventos() {

    const container =
        document.querySelector(
            ".security-events"
        );


    if (!container) {

        return;
    }


    container.innerHTML = "";


    const eventos =
        analizarEventos();


    const contador =
        document.querySelector(
            ".alert-counter"
        );


    if (contador) {

        contador.textContent =
            `${eventos.alertas} alertas`;
    }


    if (!eventos.lista.length) {

        container.innerHTML = `

            <div
                style="
                    padding:20px;
                    text-align:center;
                    color:#64748b;
                "
            >
                No se detectaron eventos sospechosos
                recientes.
            </div>

        `;

        return;
    }


    eventos.lista
        .slice(0, 20)
        .forEach(
            evento => {

                const div =
                    document.createElement(
                        "div"
                    );


                div.className =
                    `security-event ${
                        evento.nivel === "danger"
                            ? "danger-event"
                            : "warning-event"
                    }`;


                div.innerHTML = `

                    <div class="event-icon">
                        !
                    </div>

                    <div>

                        <strong>
                            ${escapar(
                                evento.titulo
                            )}
                        </strong>

                        <small>
                            ${escapar(
                                evento.descripcion
                            )}
                        </small>

                    </div>

                    <span>
                        ${evento.tiempo}
                    </span>

                `;


                container.appendChild(
                    div
                );
            }
        );
}


/* ============================================================
   ANALIZAR SIGN-INS
   ============================================================ */

function analizarEventos() {

    const lista = [];


    securityData.signIns
        .forEach(
            signIn => {

                const errorCode =
                    signIn.status
                    ?.errorCode;


                /*
                 * LOGIN FALLIDO
                 */

                if (
                    errorCode &&
                    errorCode !== 0
                ) {

                    lista.push({

                        nivel:
                            "danger",

                        titulo:
                            "Inicio de sesión fallido",

                        descripcion:
                            `${signIn.userDisplayName || signIn.userPrincipalName || "Usuario desconocido"} · IP ${signIn.ipAddress || "desconocida"}`,

                        fecha:
                            signIn.createdDateTime,

                        tiempo:
                            tiempoRelativo(
                                signIn.createdDateTime
                            )

                    });
                }


                /*
                 * RIESGO DETECTADO
                 */

                if (
                    signIn.riskDetail &&
                    signIn.riskDetail !==
                    "none"
                ) {

                    lista.push({

                        nivel:
                            "danger",

                        titulo:
                            "Inicio de sesión con riesgo",

                        descripcion:
                            `${signIn.userDisplayName || signIn.userPrincipalName || "Usuario"} · ${signIn.riskDetail}`,

                        fecha:
                            signIn.createdDateTime,

                        tiempo:
                            tiempoRelativo(
                                signIn.createdDateTime
                            )

                    });
                }


                /*
                 * UBICACIÓN
                 */

                if (
                    signIn.location?.countryOrRegion
                ) {

                    const pais =
                        signIn.location.countryOrRegion;


                    /*
                     * Aquí NO inventamos una
                     * ubicación "habitual".
                     *
                     * Solo mostramos la
                     * información entregada
                     * por Graph.
                     */
                }

            }
        );


    lista.sort(
        (
            a,
            b
        ) =>
            new Date(b.fecha)
            -
            new Date(a.fecha)
    );


    return {

        lista,

        alertas:
            lista.length

    };
}


/* ============================================================
   UTILIDADES
   ============================================================ */

function ponerTexto(
    id,
    valor
) {

    const elemento =
        document.getElementById(
            id
        );


    if (elemento) {

        elemento.textContent =
            valor;
    }
}


/* ============================================================
   INICIALES
   ============================================================ */

function obtenerIniciales(
    nombre
) {

    return nombre
        .split(" ")
        .filter(
            parte =>
                parte.trim()
        )
        .slice(0, 2)
        .map(
            parte =>
                parte[0]
        )
        .join("")
        .toUpperCase();
}


/* ============================================================
   FECHA
   ============================================================ */

function formatearFecha(
    fecha
) {

    if (!fecha) {

        return "Sin fecha";
    }


    return new Date(
        fecha
    ).toLocaleString(
        "es-PE",
        {
            dateStyle: "short",
            timeStyle: "short"
        }
    );
}


/* ============================================================
   TIEMPO RELATIVO
   ============================================================ */

function tiempoRelativo(
    fecha
) {

    if (!fecha) {

        return "";
    }


    const ahora =
        new Date();


    const fechaEvento =
        new Date(
            fecha
        );


    const minutos =
        Math.floor(
            (
                ahora -
                fechaEvento
            )
            /
            60000
        );


    if (minutos < 1) {

        return "Ahora";
    }


    if (minutos < 60) {

        return `Hace ${minutos} min`;
    }


    const horas =
        Math.floor(
            minutos / 60
        );


    if (horas < 24) {

        return `Hace ${horas} h`;
    }


    const dias =
        Math.floor(
            horas / 24
        );


    return `Hace ${dias} días`;
}


/* ============================================================
   ESCAPAR HTML
   ============================================================ */

function escapar(
    valor
) {

    if (
        valor === null ||
        valor === undefined
    ) {

        return "";
    }


    return String(valor)
        .replaceAll(
            "&",
            "&amp;"
        )
        .replaceAll(
            "<",
            "&lt;"
        )
        .replaceAll(
            ">",
            "&gt;"
        )
        .replaceAll(
            '"',
            "&quot;"
        )
        .replaceAll(
            "'",
            "&#039;"
        );
}


/* ============================================================
   ERRORES
   ============================================================ */

function obtenerMensajeError(
    error
) {

    if (
        error?.message?.includes(
            "403"
        )
    ) {

        return `
            Microsoft Graph rechazó la solicitud (403).
            Verifica que la aplicación tenga los permisos
            necesarios y que se haya otorgado Admin Consent.
        `;
    }


    if (
        error?.message?.includes(
            "401"
        )
    ) {

        return `
            La sesión de Microsoft 365 ya no es válida.
            Vuelve a iniciar sesión.
        `;
    }


    return (
        error?.message
        ||
        "Error desconocido."
    );
}


/* ============================================================
   MENSAJE DE ERROR
   ============================================================ */

function mostrarError(
    mensaje
) {

    const content =
        document.querySelector(
            ".content"
        );


    if (!content) {

        return;
    }


    const anterior =
        document.querySelector(
            ".security-error"
        );


    if (anterior) {

        anterior.remove();
    }


    const div =
        document.createElement(
            "div"
        );


    div.className =
        "security-error";


    div.style.cssText = `
        margin-bottom:20px;
        padding:15px 18px;
        border-radius:10px;
        background:#fee2e2;
        border:1px solid #fecaca;
        color:#991b1b;
        font-size:14px;
    `;


    div.innerHTML = `
        <strong>
            Error de seguridad
        </strong>
        <br>
        ${escapar(mensaje)}
    `;


    content.prepend(
        div
    );
}


/* ============================================================
   ESTADO DE CARGA
   ============================================================ */

function mostrarEstadoCarga(
    mensaje
) {

    console.log(
        mensaje
    );
}


/* ============================================================
   BOTÓN ACTUALIZAR
   ============================================================ */

document.addEventListener(
    "click",
    async event => {

        const button =
            event.target.closest(
                "#refreshButton"
            );


        if (!button) {

            return;
        }


        const textoOriginal =
            button.innerHTML;


        try {

            button.disabled =
                true;


            button.innerHTML =
                "↻ Actualizando...";


            await cargarSeguridad();


            actualizarDashboard();

        } catch (error) {

            console.error(
                error
            );


            mostrarError(
                obtenerMensajeError(
                    error
                )
            );

        } finally {

            button.disabled =
                false;


            button.innerHTML =
                textoOriginal;
        }

    }
);