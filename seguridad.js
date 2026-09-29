
/* ============================================================
   ALFERZA LIVE OFFICE
   SEGURIDAD - MICROSOFT GRAPH

   DATOS UTILIZADOS:

   ✓ Usuarios
   ✓ Cuentas bloqueadas
   ✓ Departamentos
   ✓ Auditoría de directorio
   ✓ Actividad administrativa

   NO SE CONSULTA:

   ✗ MFA
   ✗ Sign-ins
   ✗ signInActivity
   ✗ Intune

   Estos servicios no están disponibles actualmente
   para el tenant de ALFERZA.
   ============================================================ */


/* ============================================================
   CONFIGURACIÓN MSAL
   ============================================================ */

const msalConfig = {

    auth: {

        clientId:
            "5d98417c-74a7-4fab-8f2c-41ac127be696",

        authority:
            "https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1",

        redirectUri:
            "https://alferzati.github.io/live-office/blank.html"

    },

    cache: {

        cacheLocation:
            "sessionStorage",

        storeAuthStateInCookie:
            false

    }

};


/* ============================================================
   PERMISOS
   ============================================================ */

const scopes = [

    "User.Read",

    "User.Read.All",

    "AuditLog.Read.All"

];


/* ============================================================
   MSAL
   ============================================================ */

const msalInstance =
    new msal.PublicClientApplication(
        msalConfig
    );


/* ============================================================
   DATOS
   ============================================================ */

const securityData = {

    users: [],

    blockedUsers: [],

    auditLogs: [],

    departments: [],

    auditTypes: [],

    status: {

        users:
            "loading",

        audit:
            "loading"

    }

};


/* ============================================================
   CARGA INICIAL
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
                obtenerMensajeError(
                    error
                )
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


    if (
        redirectResponse?.account
    ) {

        msalInstance.setActiveAccount(
            redirectResponse.account
        );

    }


    const accounts =
        msalInstance.getAllAccounts();


    if (
        !accounts.length
    ) {

        window.location.replace(
            "login.html"
        );

        throw new Error(
            "No existe una sesión de Microsoft 365."
        );

    }


    const account =
        msalInstance.getActiveAccount()
        ||
        accounts[0];


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

                scopes:
                    scopes,

                account:
                    account

            });


        return response.accessToken;

    } catch (error) {

        console.warn(
            "Silent token falló. Solicitando token nuevamente...",
            error
        );


        const response =
            await msalInstance.acquireTokenPopup({

                scopes:
                    scopes,

                account:
                    account

            });


        return response.accessToken;

    }

}


/* ============================================================
   PETICIÓN GRAPH
============================================================ */

async function graph(
    endpoint
) {

    const token =
        await obtenerToken();


    const base =
        "https://graph.microsoft.com/v1.0";


    const url =
        endpoint.startsWith("http")
            ? endpoint
            : `${base}${endpoint}`;


    console.log(
        "GRAPH REQUEST:",
        url
    );


    const response =
        await fetch(
            url,
            {

                method:
                    "GET",

                headers: {

                    Authorization:
                        `Bearer ${token}`,

                    Accept:
                        "application/json"

                }

            }
        );


    const text =
        await response.text();


    let data =
        {};


    try {

        data =
            text
                ? JSON.parse(text)
                : {};

    } catch {

        data = {

            raw:
                text

        };

    }


    if (!response.ok) {

        console.error(
            "GRAPH ERROR:",
            response.status,
            url,
            data
        );


        const error =
            new Error(
                `Graph ${response.status}: ${
                    data?.error?.message
                    ||
                    data?.message
                    ||
                    "Error desconocido"
                }`
            );


        error.status =
            response.status;


        error.graphData =
            data;


        error.url =
            url;


        throw error;

    }


    console.log(
        "GRAPH OK:",
        response.status,
        url
    );


    return data;

}


/* ============================================================
   CARGAR SEGURIDAD
============================================================ */

async function cargarSeguridad() {

    mostrarEstadoCarga(
        "Consultando Microsoft 365..."
    );


    console.log(
        "========================================"
    );

    console.log(
        "INICIANDO CARGA DE SEGURIDAD"
    );

    console.log(
        "========================================"
    );


    const resultados =
        await Promise.allSettled([

            cargarUsuarios(),

            cargarAuditoria()

        ]);


    const nombres = [

        "Usuarios",

        "Auditoría"

    ];


    resultados.forEach(
        (
            resultado,
            index
        ) => {

            console.log(
                `${nombres[index]}:`,
                resultado.status
            );


            if (
                resultado.status ===
                "rejected"
            ) {

                console.error(
                    `Error cargando ${nombres[index]}:`,
                    resultado.reason
                );

            }

        }
    );


    console.log(
        "ESTADO FINAL:",
        securityData.status
    );


    console.log(
        "Datos de seguridad:",
        securityData
    );

}


/* ============================================================
   USUARIOS
============================================================ */

async function cargarUsuarios() {

    try {

        securityData.status.users =
            "loading";


        let url =
            "/users?$select=id,displayName,userPrincipalName,accountEnabled,department&$top=999";


        const usuarios =
            [];


        while (url) {

            const data =
                await graph(
                    url
                );


            usuarios.push(
                ...(data.value || [])
            );


            url =
                data["@odata.nextLink"]
                ||
                null;

        }


        securityData.users =
            usuarios;


        securityData.blockedUsers =
            usuarios.filter(
                user =>
                    user.accountEnabled === false
            );


        securityData.departments =
            calcularDepartamentos(
                usuarios
            );


        securityData.status.users =
            "ok";


        console.log(
            "Usuarios:",
            usuarios.length
        );


        console.log(
            "Cuentas bloqueadas:",
            securityData.blockedUsers.length
        );


        console.log(
            "Departamentos:",
            securityData.departments
        );


    } catch (error) {

        securityData.status.users =
            "error";


        securityData.users =
            [];


        securityData.blockedUsers =
            [];


        securityData.departments =
            [];


        console.error(
            "Error cargando usuarios:",
            error
        );


        throw error;

    }

}


/* ============================================================
   CALCULAR DEPARTAMENTOS
============================================================ */

function calcularDepartamentos(
    usuarios
) {

    const mapa =
        new Map();


    usuarios.forEach(
        user => {

            const departamento =
                (
                    user.department
                    ||
                    "Sin departamento"
                )
                .trim();


            const actual =
                mapa.get(
                    departamento
                )
                ||
                0;


            mapa.set(
                departamento,
                actual + 1
            );

        }
    );


    return Array
        .from(
            mapa.entries()
        )
        .map(
            ([nombre, cantidad]) => ({

                nombre,

                cantidad

            })
        )
        .sort(
            (
                a,
                b
            ) =>
                b.cantidad
                -
                a.cantidad
        );

}


/* ============================================================
   AUDITORÍA
============================================================ */

async function cargarAuditoria() {

    try {

        securityData.status.audit =
            "loading";


        const fecha =
            new Date();


        /*
         * Microsoft Graph actualmente permite
         * consultar este período para este tenant.
         */

        fecha.setDate(
            fecha.getDate() - 30
        );


        const iso =
            fecha.toISOString();


        const filter =
            encodeURIComponent(
                `activityDateTime ge ${iso}`
            );


        let url =
            `/auditLogs/directoryAudits?$filter=${filter}&$top=1000`;


        const eventos =
            [];


        while (url) {

            const data =
                await graph(
                    url
                );


            eventos.push(
                ...(data.value || [])
            );


            url =
                data["@odata.nextLink"]
                ||
                null;

        }


        securityData.auditLogs =
            eventos;


        securityData.auditTypes =
            calcularTiposAuditoria(
                eventos
            );


        securityData.status.audit =
            "ok";


        console.log(
            "Eventos de auditoría:",
            eventos.length
        );


        console.log(
            "Tipos de auditoría:",
            securityData.auditTypes
        );


    } catch (error) {

        securityData.status.audit =
            "error";


        securityData.auditLogs =
            [];


        securityData.auditTypes =
            [];


        console.error(
            "Error cargando auditoría:",
            error
        );


        throw error;

    }

}


/* ============================================================
   TIPOS DE AUDITORÍA
============================================================ */

function calcularTiposAuditoria(
    eventos
) {

    const mapa =
        new Map();


    eventos.forEach(
        evento => {

            const nombre =
                clasificarEvento(
                    evento.activityDisplayName
                );


            const actual =
                mapa.get(
                    nombre
                )
                ||
                0;


            mapa.set(
                nombre,
                actual + 1
            );

        }
    );


    return Array
        .from(
            mapa.entries()
        )
        .map(
            ([nombre, cantidad]) => ({

                nombre,

                cantidad

            })
        )
        .sort(
            (
                a,
                b
            ) =>
                b.cantidad
                -
                a.cantidad
        )
        .slice(
            0,
            6
        );

}


/* ============================================================
   CLASIFICAR EVENTO
============================================================ */

function clasificarEvento(
    actividad
) {

    const texto =
        (
            actividad
            ||
            "Otra actividad"
        )
        .toLowerCase();


    if (
        texto.includes(
            "user"
        )
        ||
        texto.includes(
            "usuario"
        )
    ) {

        return "Usuarios";

    }


    if (
        texto.includes(
            "group"
        )
        ||
        texto.includes(
            "grupo"
        )
    ) {

        return "Grupos";

    }


    if (
        texto.includes(
            "application"
        )
        ||
        texto.includes(
            "app"
        )
    ) {

        return "Aplicaciones";

    }


    if (
        texto.includes(
            "role"
        )
        ||
        texto.includes(
            "rol"
        )
    ) {

        return "Roles";

    }


    if (
        texto.includes(
            "policy"
        )
        ||
        texto.includes(
            "polic"
        )
    ) {

        return "Políticas";

    }


    return "Otras acciones";

}


/* ============================================================
   ACTUALIZAR DASHBOARD
============================================================ */

function actualizarDashboard() {

    actualizarContadores();

    actualizarDepartamentos();

    actualizarResumenAuditoria();

    actualizarBloqueados();

    actualizarResumenDirectorio();

    actualizarTablaAuditoria();

}


/* ============================================================
   CONTADORES
============================================================ */

function actualizarContadores() {

    ponerTexto(
        "totalUsuarios",
        securityData.users.length
    );


    ponerTexto(
        "cuentasBloqueadas",
        securityData.blockedUsers.length
    );


    ponerTexto(
        "eventosAuditoria",
        securityData.auditLogs.length
    );


    ponerTexto(
        "totalDepartamentos",
        securityData.departments.length
    );


    ponerTexto(
        "blockedCounter",
        securityData.blockedUsers.length
    );


    ponerTexto(
        "auditCounter",
        `${securityData.auditLogs.length} eventos`
    );

}


/* ============================================================
   DEPARTAMENTOS
============================================================ */

function actualizarDepartamentos() {

    const container =
        document.getElementById(
            "departmentList"
        );


    if (!container) {

        return;

    }


    container.innerHTML =
        "";


    if (
        !securityData.departments.length
    ) {

        container.innerHTML = `

            <div class="loading">

                No hay información disponible.

            </div>

        `;

        return;

    }


    const max =
        securityData.departments[0]
            .cantidad;


    securityData.departments
        .slice(
            0,
            8
        )
        .forEach(
            departamento => {

                const porcentaje =
                    Math.max(
                        3,
                        (
                            departamento.cantidad /
                            max
                        ) * 100
                    );


                const row =
                    document.createElement(
                        "div"
                    );


                row.className =
                    "department-row";


                row.innerHTML = `

                    <div
                        class="department-name"
                        title="${escapar(
                            departamento.nombre
                        )}"
                    >

                        ${escapar(
                            departamento.nombre
                        )}

                    </div>


                    <div class="department-bar">

                        <div
                            class="department-fill"
                            style="width:${porcentaje}%"
                        ></div>

                    </div>


                    <div class="department-count">

                        ${departamento.cantidad}

                    </div>

                `;


                container.appendChild(
                    row
                );

            }
        );

}


/* ============================================================
   RESUMEN DE AUDITORÍA
============================================================ */

function actualizarResumenAuditoria() {

    const container =
        document.getElementById(
            "auditSummary"
        );


    if (!container) {

        return;

    }


    container.innerHTML =
        "";


    if (
        !securityData.auditTypes.length
    ) {

        container.innerHTML = `

            <div class="loading">

                No hay eventos disponibles.

            </div>

        `;

        return;

    }


    const iconos = {

        "Usuarios":
            "👤",

        "Grupos":
            "👥",

        "Aplicaciones":
            "▣",

        "Roles":
            "🔑",

        "Políticas":
            "⚙",

        "Otras acciones":
            "📋"

    };


    securityData.auditTypes
        .forEach(
            tipo => {

                const row =
                    document.createElement(
                        "div"
                    );


                row.className =
                    "audit-type-row";


                row.innerHTML = `

                    <div class="audit-type-icon">

                        ${
                            iconos[
                                tipo.nombre
                            ]
                            ||
                            "📋"
                        }

                    </div>


                    <div class="audit-type-info">

                        <strong>
                            ${escapar(
                                tipo.nombre
                            )}
                        </strong>

                        <small>
                            Actividades registradas
                        </small>

                    </div>


                    <div class="audit-type-count">

                        ${tipo.cantidad}

                    </div>

                `;


                container.appendChild(
                    row
                );

            }
        );

}


/* ============================================================
   CUENTAS BLOQUEADAS
============================================================ */

function actualizarBloqueados() {

    const container =
        document.getElementById(
            "blockedList"
        );


    if (!container) {

        return;

    }


    container.innerHTML =
        "";


    if (
        !securityData.blockedUsers.length
    ) {

        container.innerHTML = `

            <div class="loading">

                No hay cuentas bloqueadas.

            </div>

        `;

        return;

    }


    securityData.blockedUsers
        .slice(
            0,
            10
        )
        .forEach(
            user => {

                const nombre =
                    user.displayName
                    ||
                    "Sin nombre";


                const correo =
                    user.userPrincipalName
                    ||
                    "Sin correo";


                const area =
                    user.department
                    ||
                    "Sin departamento";


                const iniciales =
                    obtenerIniciales(
                        nombre
                    );


                const div =
                    document.createElement(
                        "div"
                    );


                div.className =
                    "blocked-user";


                div.innerHTML = `

                    <div class="avatar red-avatar">

                        ${escapar(
                            iniciales
                        )}

                    </div>


                    <div class="blocked-info">

                        <strong>

                            ${escapar(
                                nombre
                            )}

                        </strong>


                        <small>

                            ${escapar(
                                area
                            )}

                        </small>


                        <small>

                            ${escapar(
                                correo
                            )}

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
   RESUMEN DEL DIRECTORIO
============================================================ */

function actualizarResumenDirectorio() {

    ponerTexto(
        "summaryUsers",
        securityData.users.length
    );


    ponerTexto(
        "summaryBlocked",
        securityData.blockedUsers.length
    );


    ponerTexto(
        "summaryAudit",
        securityData.auditLogs.length
    );


    ponerTexto(
        "summaryDepartments",
        securityData.departments.length
    );

}


/* ============================================================
   TABLA DE AUDITORÍA
============================================================ */

function actualizarTablaAuditoria() {

    const tbody =
        document.getElementById(
            "auditTable"
        );


    if (!tbody) {

        return;

    }


    tbody.innerHTML =
        "";


    if (
        !securityData.auditLogs.length
    ) {

        tbody.innerHTML = `

            <tr>

                <td
                    colspan="4"
                    class="loading-cell"
                >

                    No hay eventos de auditoría
                    registrados en los últimos 30 días.

                </td>

            </tr>

        `;

        return;

    }


    securityData.auditLogs
        .slice(
            0,
            50
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
        )
        .forEach(
            evento => {

                const accion =
                    evento.activityDisplayName
                    ||
                    "Actividad desconocida";


                const iniciador =
                    obtenerIniciador(
                        evento
                    );


                const recurso =
                    obtenerRecurso(
                        evento
                    );


                const fecha =
                    formatearFecha(
                        evento.activityDateTime
                    );


                const icono =
                    obtenerIconoEvento(
                        accion
                    );


                const tr =
                    document.createElement(
                        "tr"
                    );


                tr.innerHTML = `

                    <td>

                        <div class="audit-action">

                            <div class="audit-icon">

                                ${icono}

                            </div>


                            <div>

                                <strong>

                                    ${escapar(
                                        accion
                                    )}

                                </strong>

                                <small>

                                    ${escapar(
                                        obtenerCategoriaEvento(
                                            accion
                                        )
                                    )}

                                </small>

                            </div>

                        </div>

                    </td>


                    <td>

                        <div class="audit-user">

                            <strong>

                                ${escapar(
                                    iniciador.nombre
                                )}

                            </strong>


                            <small>

                                ${escapar(
                                    iniciador.correo
                                )}

                            </small>

                        </div>

                    </td>


                    <td>

                        <div
                            class="target-name"
                            title="${escapar(
                                recurso
                            )}"
                        >

                            ${escapar(
                                recurso
                            )}

                        </div>

                    </td>


                    <td>

                        <div class="audit-date">

                            ${fecha}

                        </div>

                    </td>

                `;


                tbody.appendChild(
                    tr
                );

            }
        );

}


/* ============================================================
   OBTENER INICIADOR
============================================================ */

function obtenerIniciador(
    evento
) {

    const user =
        evento.initiatedBy
        ?.user;


    const app =
        evento.initiatedBy
        ?.app;


    if (user) {

        return {

            nombre:
                user.displayName
                ||
                "Usuario",

            correo:
                user.userPrincipalName
                ||
                "Cuenta Microsoft"

        };

    }


    if (app) {

        return {

            nombre:
                app.displayName
                ||
                "Aplicación",

            correo:
                app.servicePrincipalName
                ||
                "Aplicación"

        };

    }


    return {

        nombre:
            "No identificado",

        correo:
            "Sin información"

    };

}


/* ============================================================
   OBTENER RECURSO
============================================================ */

function obtenerRecurso(
    evento
) {

    const recursos =
        evento.targetResources
        ||
        [];


    if (!recursos.length) {

        return "Sin recurso especificado";

    }


    const nombres =
        recursos
            .map(
                recurso => {

                    return (
                        recurso.userPrincipalName
                        ||
                        recurso.displayName
                        ||
                        recurso.groupName
                        ||
                        recurso.id
                        ||
                        "Recurso"
                    );

                }
            )
            .filter(
                Boolean
            );


    if (!nombres.length) {

        return "Recurso no identificado";

    }


    return nombres.join(
        ", "
    );

}


/* ============================================================
   CATEGORÍA DEL EVENTO
============================================================ */

function obtenerCategoriaEvento(
    actividad
) {

    const texto =
        (
            actividad
            ||
            ""
        )
        .toLowerCase();


    if (
        texto.includes("user")
        ||
        texto.includes("usuario")
    ) {

        return "Usuarios";

    }


    if (
        texto.includes("group")
        ||
        texto.includes("grupo")
    ) {

        return "Grupos";

    }


    if (
        texto.includes("application")
        ||
        texto.includes("app")
    ) {

        return "Aplicaciones";

    }


    if (
        texto.includes("role")
        ||
        texto.includes("rol")
    ) {

        return "Roles";

    }


    if (
        texto.includes("policy")
        ||
        texto.includes("polic")
    ) {

        return "Políticas";

    }


    return "Otras acciones";

}


/* ============================================================
   ICONO DEL EVENTO
============================================================ */

function obtenerIconoEvento(
    actividad
) {

    const categoria =
        obtenerCategoriaEvento(
            actividad
        );


    const iconos = {

        "Usuarios":
            "👤",

        "Grupos":
            "👥",

        "Aplicaciones":
            "▣",

        "Roles":
            "🔑",

        "Políticas":
            "⚙",

        "Otras acciones":
            "📋"

    };


    return (
        iconos[
            categoria
        ]
        ||
        "📋"
    );

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

        .slice(
            0,
            2
        )

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


    const fechaObjeto =
        new Date(
            fecha
        );


    if (
        Number.isNaN(
            fechaObjeto.getTime()
        )
    ) {

        return "Fecha no válida";

    }


    return fechaObjeto.toLocaleString(
        "es-PE",
        {

            dateStyle:
                "short",

            timeStyle:
                "short"

        }
    );

}


/* ============================================================
   PONER TEXTO
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
   ESCAPAR HTML
============================================================ */

function escapar(
    valor
) {

    if (
        valor === null
        ||
        valor === undefined
    ) {

        return "";

    }


    return String(
        valor
    )

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
   ERROR
============================================================ */

function obtenerMensajeError(
    error
) {

    const mensaje =
        error?.message
        ||
        "Error desconocido.";


    if (
        mensaje.includes(
            "403"
        )
    ) {

        return `
            Microsoft Graph rechazó la solicitud (403).
            Verifica los permisos de Microsoft Graph.
        `;

    }


    if (
        mensaje.includes(
            "401"
        )
    ) {

        return `
            La sesión de Microsoft 365 ya no es válida.
            Vuelve a iniciar sesión.
        `;

    }


    if (
        mensaje.includes(
            "400"
        )
    ) {

        return `
            Microsoft Graph rechazó la solicitud (400).
            Revisa el endpoint y los parámetros enviados.
        `;

    }


    return mensaje;

}


/* ============================================================
   MOSTRAR ERROR
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


    div.innerHTML = `

        <strong>
            Error de seguridad
        </strong>

        <br><br>

        ${escapar(
            mensaje
        )}

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
                "ERROR AL ACTUALIZAR:",
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

/* ============================================================
   FILTROS DEL DASHBOARD
============================================================ */

let filtroActivo = null;


document.addEventListener(
    "click",
    event => {

        const card =
            event.target.closest(
                ".filter-card"
            );


        if (!card) {

            return;

        }


        const filtro =
            card.dataset.filter;


        /*
         * Si hacemos click nuevamente
         * sobre el filtro activo,
         * mostramos todo el dashboard.
         */

        if (
            filtroActivo === filtro
        ) {

            mostrarTodoDashboard();

            return;

        }


        filtroActivo =
            filtro;


        aplicarFiltro(
            filtro
        );

    }
);


/* ============================================================
   SOPORTE PARA TECLADO
============================================================ */

document.addEventListener(
    "keydown",
    event => {

        if (
            event.key !== "Enter"
            &&
            event.key !== " "
        ) {

            return;

        }


        const card =
            event.target.closest(
                ".filter-card"
            );


        if (!card) {

            return;

        }


        event.preventDefault();


        const filtro =
            card.dataset.filter;


        if (
            filtroActivo === filtro
        ) {

            mostrarTodoDashboard();

            return;

        }


        filtroActivo =
            filtro;


        aplicarFiltro(
            filtro
        );

    }
);


/* ============================================================
   APLICAR FILTRO
============================================================ */

function aplicarFiltro(
    filtro
) {

    const sections =
        document.querySelectorAll(
            ".filter-section"
        );


    /*
     * Ocultar todas las secciones
     */

    sections.forEach(
        section => {

            section.classList.add(
                "filter-hidden"
            );

        }
    );


    /*
     * Mostrar solamente
     * las relacionadas con el filtro.
     */

    sections.forEach(
        section => {

            const contenido =
                section.dataset.section
                ||
                "";


            const categorias =
                contenido.split(
                    " "
                );


            if (
                categorias.includes(
                    filtro
                )
            ) {

                section.classList.remove(
                    "filter-hidden"
                );

            }

        }
    );


    /*
     * Marcar tarjeta activa
     */

    document
        .querySelectorAll(
            ".filter-card"
        )
        .forEach(
            card => {

                card.classList.remove(
                    "filter-active"
                );

            }
        );


    const tarjetaActiva =
        document.querySelector(
            `.filter-card[data-filter="${filtro}"]`
        );


    if (
        tarjetaActiva
    ) {

        tarjetaActiva.classList.add(
            "filter-active"
        );

    }

}


/* ============================================================
   MOSTRAR TODO
============================================================ */

function mostrarTodoDashboard() {

    filtroActivo =
        null;


    document
        .querySelectorAll(
            ".filter-section"
        )
        .forEach(
            section => {

                section.classList.remove(
                    "filter-hidden"
                );

            }
        );


    document
        .querySelectorAll(
            ".filter-card"
        )
        .forEach(
            card => {

                card.classList.remove(
                    "filter-active"
                );

            }
        );

}