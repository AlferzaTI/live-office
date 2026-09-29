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
   PAGINACIÓN
============================================================ */

const paginationState = {

    users: 1,

    blocked: 1,

    departments: 1,

    audit: 1

};


/*
   Cantidad de registros cuando estamos
   en el dashboard principal.
*/

const DASHBOARD_ITEMS = 5;


/*
   Cantidad de registros cuando
   estamos dentro de un filtro.
*/

const ITEMS_PER_PAGE = 20;


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


        paginationState.users =
            1;

        paginationState.blocked =
            1;

        paginationState.departments =
            1;


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
                String(
                    user.department
                    ||
                    ""
                )
                .trim()
                ||
                "Sin departamento";


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
            eventos
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


        securityData.auditTypes =
            calcularTiposAuditoria(
                securityData.auditLogs
            );


        securityData.status.audit =
            "ok";


        paginationState.audit =
            1;


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
   ACTUALIZAR DASHBOARD
============================================================ */

function actualizarDashboard() {

    actualizarContadores();

    actualizarUsuarios();

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
        "usersCounter",
        securityData.users.length
    );


    ponerTexto(
        "departmentsCounter",
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
   USUARIOS REGISTRADOS
============================================================ */

function actualizarUsuarios() {

    const tbody =
        document.getElementById(
            "usersTable"
        );


    if (!tbody) {

        return;

    }


    tbody.innerHTML =
        "";


    if (
        !securityData.users.length
    ) {

        tbody.innerHTML = `

            <tr>

                <td
                    colspan="4"
                    class="loading-cell"
                >

                    No hay usuarios disponibles.

                </td>

            </tr>

        `;


        actualizarPaginacion(
            "usersPagination",
            0,
            1,
            () => {}
        );


        return;

    }


    /*
       ========================================================
       DASHBOARD PRINCIPAL
       ========================================================
       Solo muestra los primeros 10 usuarios.
    */

    if (
        filtroActivo === null
    ) {

        const usuarios =
            securityData.users.slice(
                0,
                DASHBOARD_ITEMS
            );


        usuarios.forEach(
            user => {

                renderUsuarioTabla(
                    tbody,
                    user
                );

            }
        );


        actualizarPaginacion(
            "usersPagination",
            0,
            1,
            () => {}
        );


        return;

    }


    /*
       ========================================================
       FILTRO ACTIVO
       ========================================================
       Mantiene 20 registros por página.
    */

    const total =
        securityData.users.length;


    const totalPaginas =
        Math.ceil(
            total /
            ITEMS_PER_PAGE
        );


    if (
        paginationState.users >
        totalPaginas
    ) {

        paginationState.users =
            totalPaginas;

    }


    const pagina =
        paginationState.users;


    const inicio =
        (
            pagina - 1
        )
        *
        ITEMS_PER_PAGE;


    const fin =
        Math.min(
            inicio +
            ITEMS_PER_PAGE,
            total
        );


    const usuarios =
        securityData.users.slice(
            inicio,
            fin
        );


    usuarios.forEach(
        user => {

            renderUsuarioTabla(
                tbody,
                user
            );

        }
    );


    actualizarPaginacion(
        "usersPagination",
        total,
        pagina,
        nuevaPagina => {

            paginationState.users =
                nuevaPagina;

            actualizarUsuarios();

        }
    );

}


/* ============================================================
   RENDER USUARIO
============================================================ */

function renderUsuarioTabla(
    tbody,
    user
) {

    const nombre =
        user.displayName
        ||
        "Sin nombre";


    const correo =
        user.userPrincipalName
        ||
        "Sin correo";


    const departamento =
        String(
            user.department
            ||
            ""
        )
        .trim()
        ||
        "Sin departamento";


    const iniciales =
        obtenerIniciales(
            nombre
        );


    const activo =
        user.accountEnabled !== false;


    const tr =
        document.createElement(
            "tr"
        );


    tr.innerHTML = `

        <td>

            <div class="user-cell">

                <div
                    class="avatar ${
                        activo
                            ? "blue-avatar"
                            : "red-avatar"
                    }"
                >

                    ${escapar(
                        iniciales
                    )}

                </div>


                <div class="user-cell-info">

                    <strong>

                        ${escapar(
                            nombre
                        )}

                    </strong>


                    <small>

                        ID: ${escapar(
                            user.id || "—"
                        )}

                    </small>

                </div>

            </div>

        </td>


        <td>

            <div
                class="email-cell"
                title="${escapar(
                    correo
                )}"
            >

                ${escapar(
                    correo
                )}

            </div>

        </td>


        <td>

            <div
                class="department-cell"
                title="${escapar(
                    departamento
                )}"
            >

                ${escapar(
                    departamento
                )}

            </div>

        </td>


        <td>

            <span
                class="status-badge ${
                    activo
                        ? "status-active"
                        : "status-blocked"
                }"
            >

                ${
                    activo
                        ? "Activo"
                        : "Bloqueado"
                }

            </span>

        </td>

    `;


    tbody.appendChild(
        tr
    );

}


/* ============================================================
   DEPARTAMENTOS
============================================================ */

function actualizarDepartamentos() {

    const tbody =
        document.getElementById(
            "departmentTable"
        );


    if (!tbody) {

        return;

    }


    tbody.innerHTML =
        "";


    if (
        !securityData.departments.length
    ) {

        tbody.innerHTML = `

            <tr>

                <td
                    colspan="3"
                    class="loading-cell"
                >

                    No hay información disponible.

                </td>

            </tr>

        `;


        actualizarPaginacion(
            "departmentsPagination",
            0,
            1,
            () => {}
        );


        return;

    }


    const total =
        securityData.departments.length;


    /*
       ========================================================
       DASHBOARD PRINCIPAL
       Solo muestra los primeros 10 departamentos.
       ========================================================
    */

    if (
        filtroActivo === null
    ) {

        const departamentos =
            securityData.departments.slice(
                0,
                DASHBOARD_ITEMS
            );


        renderDepartamentos(
            tbody,
            departamentos
        );


        actualizarPaginacion(
            "departmentsPagination",
            0,
            1,
            () => {}
        );


        return;

    }


    /*
       ========================================================
       FILTRO ACTIVO
       ========================================================
       Mantiene 20 registros por página.
    */

    const totalPaginas =
        Math.ceil(
            total /
            ITEMS_PER_PAGE
        );


    if (
        paginationState.departments >
        totalPaginas
    ) {

        paginationState.departments =
            totalPaginas;

    }


    const pagina =
        paginationState.departments;


    const inicio =
        (
            pagina - 1
        )
        *
        ITEMS_PER_PAGE;


    const fin =
        Math.min(
            inicio +
            ITEMS_PER_PAGE,
            total
        );


    const departamentos =
        securityData.departments.slice(
            inicio,
            fin
        );


    renderDepartamentos(
        tbody,
        departamentos
    );


    actualizarPaginacion(
        "departmentsPagination",
        total,
        pagina,
        nuevaPagina => {

            paginationState.departments =
                nuevaPagina;

            actualizarDepartamentos();

        }
    );

}


/* ============================================================
   RENDER DEPARTAMENTOS
============================================================ */

function renderDepartamentos(
    tbody,
    departamentos
) {

    const totalUsuarios =
        securityData.users.length;


    const max =
        securityData.departments[0]
            ?.cantidad
            ||
            1;


    departamentos.forEach(
        departamento => {

            const porcentaje =
                totalUsuarios > 0
                    ? (
                        departamento.cantidad /
                        totalUsuarios
                    ) * 100
                    : 0;


            const anchoBarra =
                Math.max(
                    3,
                    (
                        departamento.cantidad /
                        max
                    ) * 100
                );


            const tr =
                document.createElement(
                    "tr"
                );


            tr.innerHTML = `

                <td>

                    <div
                        class="department-name-cell"
                        title="${escapar(
                            departamento.nombre
                        )}"
                    >

                        ${escapar(
                            departamento.nombre
                        )}

                    </div>

                </td>


                <td>

                    <div class="department-count-cell">

                        ${departamento.cantidad}

                    </div>

                </td>


                <td>

                    <div class="department-progress">

                        <div class="department-bar">

                            <div
                                class="department-fill"
                                style="width:${anchoBarra}%"
                            ></div>

                        </div>


                        <span class="department-percent">

                            ${porcentaje.toFixed(1)}%

                        </span>

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

    const tbody =
        document.getElementById(
            "blockedTable"
        );


    if (!tbody) {

        return;

    }


    tbody.innerHTML =
        "";


    if (
        !securityData.blockedUsers.length
    ) {

        tbody.innerHTML = `

            <tr>

                <td
                    colspan="4"
                    class="loading-cell"
                >

                    No hay cuentas bloqueadas.

                </td>

            </tr>

        `;


        actualizarPaginacion(
            "blockedPagination",
            0,
            1,
            () => {}
        );


        return;

    }


    /*
       ========================================================
       DASHBOARD PRINCIPAL
       Solo muestra las primeras 10 cuentas.
       ========================================================
    */

    if (
        filtroActivo === null
    ) {

        const usuarios =
            securityData.blockedUsers.slice(
                0,
                DASHBOARD_ITEMS
            );


        renderBloqueados(
            tbody,
            usuarios
        );


        actualizarPaginacion(
            "blockedPagination",
            0,
            1,
            () => {}
        );


        return;

    }


    /*
       ========================================================
       FILTRO ACTIVO
       ========================================================
       Mantiene 20 registros por página.
    */

    const total =
        securityData.blockedUsers.length;


    const totalPaginas =
        Math.ceil(
            total /
            ITEMS_PER_PAGE
        );


    if (
        paginationState.blocked >
        totalPaginas
    ) {

        paginationState.blocked =
            totalPaginas;

    }


    const pagina =
        paginationState.blocked;


    const inicio =
        (
            pagina - 1
        )
        *
        ITEMS_PER_PAGE;


    const fin =
        Math.min(
            inicio +
            ITEMS_PER_PAGE,
            total
        );


    const usuarios =
        securityData.blockedUsers.slice(
            inicio,
            fin
        );


    renderBloqueados(
        tbody,
        usuarios
    );


    actualizarPaginacion(
        "blockedPagination",
        total,
        pagina,
        nuevaPagina => {

            paginationState.blocked =
                nuevaPagina;

            actualizarBloqueados();

        }
    );

}


/* ============================================================
   RENDER BLOQUEADOS
============================================================ */

function renderBloqueados(
    tbody,
    usuarios
) {

    usuarios.forEach(
        user => {

            const nombre =
                user.displayName
                ||
                "Sin nombre";


            const correo =
                user.userPrincipalName
                ||
                "Sin correo";


            const departamento =
                String(
                    user.department
                    ||
                    ""
                )
                .trim()
                ||
                "Sin departamento";


            const iniciales =
                obtenerIniciales(
                    nombre
                );


            const tr =
                document.createElement(
                    "tr"
                );


            tr.innerHTML = `

                <td>

                    <div class="user-cell">

                        <div class="avatar red-avatar">

                            ${escapar(
                                iniciales
                            )}

                        </div>


                        <div class="user-cell-info">

                            <strong>

                                ${escapar(
                                    nombre
                                )}

                            </strong>


                            <small>

                                Cuenta deshabilitada

                            </small>

                        </div>

                    </div>

                </td>


                <td>

                    <div
                        class="email-cell"
                        title="${escapar(
                            correo
                        )}"
                    >

                        ${escapar(
                            correo
                        )}

                    </div>

                </td>


                <td>

                    <div
                        class="department-cell"
                        title="${escapar(
                            departamento
                        )}"
                    >

                        ${escapar(
                            departamento
                        )}

                    </div>

                </td>


                <td>

                    <span class="status-badge status-blocked">

                        Bloqueada

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


        actualizarPaginacion(
            "auditPagination",
            0,
            1,
            () => {}
        );


        return;

    }


    /*
       ========================================================
       DASHBOARD PRINCIPAL
       Solo muestra los 10 eventos más recientes.
       ========================================================
    */

    if (
        filtroActivo === null
    ) {

        const eventos =
            securityData.auditLogs.slice(
                0,
                DASHBOARD_ITEMS
            );


        renderEventosAuditoria(
            tbody,
            eventos
        );


        actualizarPaginacion(
            "auditPagination",
            0,
            1,
            () => {}
        );


        return;

    }


    /*
       ========================================================
       FILTRO ACTIVO
       ========================================================
       Mantiene 20 registros por página.
    */

    const total =
        securityData.auditLogs.length;


    const totalPaginas =
        Math.ceil(
            total /
            ITEMS_PER_PAGE
        );


    if (
        paginationState.audit >
        totalPaginas
    ) {

        paginationState.audit =
            totalPaginas;

    }


    const pagina =
        paginationState.audit;


    const inicio =
        (
            pagina - 1
        )
        *
        ITEMS_PER_PAGE;


    const fin =
        Math.min(
            inicio +
            ITEMS_PER_PAGE,
            total
        );


    const eventos =
        securityData.auditLogs.slice(
            inicio,
            fin
        );


    renderEventosAuditoria(
        tbody,
        eventos
    );


    actualizarPaginacion(
        "auditPagination",
        total,
        pagina,
        nuevaPagina => {

            paginationState.audit =
                nuevaPagina;

            actualizarTablaAuditoria();

        }
    );

}


/* ============================================================
   RENDER EVENTOS DE AUDITORÍA
============================================================ */

function renderEventosAuditoria(
    tbody,
    eventos
) {

    eventos.forEach(
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
   PAGINACIÓN
============================================================ */

function actualizarPaginacion(
    containerId,
    totalItems,
    paginaActual,
    onPageChange
) {

    const container =
        document.getElementById(
            containerId
        );


    if (!container) {

        return;

    }


    container.innerHTML =
        "";


    if (
        totalItems <= 0
    ) {

        return;

    }


    const totalPaginas =
        Math.ceil(
            totalItems /
            ITEMS_PER_PAGE
        );


    const inicio =
        (
            paginaActual - 1
        )
        *
        ITEMS_PER_PAGE
        +
        1;


    const fin =
        Math.min(
            paginaActual *
            ITEMS_PER_PAGE,
            totalItems
        );


    const info =
        document.createElement(
            "div"
        );


    info.className =
        "pagination-info";


    info.textContent =
        `Mostrando ${inicio}–${fin} de ${totalItems}`;


    const controls =
        document.createElement(
            "div"
        );


    controls.className =
        "pagination-controls";


    const anterior =
        crearBotonPaginacion(
            "‹",
            paginaActual === 1,
            false,
            () => {

                onPageChange(
                    paginaActual - 1
                );

            }
        );


    controls.appendChild(
        anterior
    );


    const paginas =
        obtenerPaginasVisibles(
            paginaActual,
            totalPaginas
        );


    paginas.forEach(
        pagina => {

            if (
                pagina === "..."
            ) {

                const ellipsis =
                    document.createElement(
                        "span"
                    );


                ellipsis.className =
                    "pagination-ellipsis";


                ellipsis.textContent =
                    "…";


                controls.appendChild(
                    ellipsis
                );


                return;

            }


            const boton =
                crearBotonPaginacion(
                    pagina,
                    false,
                    pagina === paginaActual,
                    () => {

                        onPageChange(
                            pagina
                        );

                    }
                );


            controls.appendChild(
                boton
            );

        }
    );


    const siguiente =
        crearBotonPaginacion(
            "›",
            paginaActual === totalPaginas,
            false,
            () => {

                onPageChange(
                    paginaActual + 1
                );

            }
        );


    controls.appendChild(
        siguiente
    );


    container.appendChild(
        info
    );


    container.appendChild(
        controls
    );

}


/* ============================================================
   CREAR BOTÓN PAGINACIÓN
============================================================ */

function crearBotonPaginacion(
    texto,
    disabled,
    active,
    callback
) {

    const button =
        document.createElement(
            "button"
        );


    button.type =
        "button";


    button.className =
        "pagination-button";


    if (active) {

        button.classList.add(
            "active"
        );

    }


    button.disabled =
        disabled;


    button.textContent =
        texto;


    button.addEventListener(
        "click",
        callback
    );


    return button;

}


/* ============================================================
   PÁGINAS VISIBLES
============================================================ */

function obtenerPaginasVisibles(
    actual,
    total
) {

    if (
        total <= 7
    ) {

        return Array.from(
            {
                length:
                    total
            },
            (
                _,
                index
            ) =>
                index + 1
        );

    }


    const paginas = [];


    paginas.push(
        1
    );


    if (
        actual > 4
    ) {

        paginas.push(
            "..."
        );

    }


    const inicio =
        Math.max(
            2,
            actual - 1
        );


    const fin =
        Math.min(
            total - 1,
            actual + 1
        );


    for (
        let i = inicio;
        i <= fin;
        i++
    ) {

        paginas.push(
            i
        );

    }


    if (
        actual < total - 3
    ) {

        paginas.push(
            "..."
        );

    }


    paginas.push(
        total
    );


    return paginas;

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

let filtroActivo =
    null;


/* ============================================================
   CLICK EN TARJETAS
============================================================ */

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

    const content =
        document.querySelector(
            ".content"
        );


    if (content) {

        content.classList.add(
            "filter-mode"
        );

    }


    document
        .querySelectorAll(
            ".filter-section"
        )
        .forEach(
            section => {

                section.classList.add(
                    "filter-hidden"
                );

            }
        );


    document
        .querySelectorAll(
            ".filter-section"
        )
        .forEach(
            section => {

                const contenido =
                    section.dataset.section
                    ||
                    "";


                const categorias =
                    contenido
                        .split(/\s+/)
                        .filter(Boolean);


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


    document
        .querySelectorAll(
            ".dashboard-card:not(.filter-section)"
        )
        .forEach(
            card => {

                card.classList.add(
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


    /*
       Al entrar al filtro volvemos
       siempre a la primera página.
    */

    if (
        filtro === "usuarios"
    ) {

        paginationState.users =
            1;

    }


    if (
        filtro === "bloqueadas"
    ) {

        paginationState.blocked =
            1;

    }


    if (
        filtro === "departamentos"
    ) {

        paginationState.departments =
            1;

    }


    if (
        filtro === "auditoria"
    ) {

        paginationState.audit =
            1;

    }


    /*
       Redibujar inmediatamente.
    */

    actualizarDashboard();


    window.scrollTo({
        top: 0,
        behavior: "smooth"
    });

}


/* ============================================================
   MOSTRAR TODO EL DASHBOARD
============================================================ */

function mostrarTodoDashboard() {

    filtroActivo =
        null;


    const content =
        document.querySelector(
            ".content"
        );


    if (content) {

        content.classList.remove(
            "filter-mode"
        );

    }


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
            ".dashboard-card:not(.filter-section)"
        )
        .forEach(
            card => {

                card.classList.remove(
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


    /*
       Volver a la primera página
       para la siguiente vez que
       se abra un filtro.
    */

    paginationState.users =
        1;

    paginationState.blocked =
        1;

    paginationState.departments =
        1;

    paginationState.audit =
        1;


    /*
       Redibujar para volver
       inmediatamente a los 10.
    */

    actualizarDashboard();

}


/* ============================================================
   BUSCADOR DE TRABAJADORES
============================================================ */

const workerSearch =
    document.getElementById(
        "workerSearch"
    );


const clearWorkerSearch =
    document.getElementById(
        "clearWorkerSearch"
    );


const workerSearchResult =
    document.getElementById(
        "workerSearchResult"
    );


/* ============================================================
   BUSCAR TRABAJADOR
============================================================ */

function buscarTrabajador(
    texto
) {

    const busqueda =
        String(
            texto || ""
        )
        .trim()
        .toLowerCase();


    if (!busqueda) {

        limpiarBusquedaTrabajador();

        return;

    }


    const resultados =
        securityData.users.filter(
            user => {

                const nombre =
                    String(
                        user.displayName || ""
                    )
                    .toLowerCase();


                const correo =
                    String(
                        user.userPrincipalName || ""
                    )
                    .toLowerCase();


                const departamento =
                    String(
                        user.department || ""
                    )
                    .toLowerCase();


                const estado =
                    user.accountEnabled === false
                        ? "bloqueado"
                        : "activo";


                return (

                    nombre.includes(
                        busqueda
                    )

                    ||

                    correo.includes(
                        busqueda
                    )

                    ||

                    departamento.includes(
                        busqueda
                    )

                    ||

                    estado.includes(
                        busqueda
                    )

                );

            }
        );


    mostrarResultadosBusqueda(
        resultados,
        busqueda
    );

}


/* ============================================================
   MOSTRAR RESULTADOS
============================================================ */

function mostrarResultadosBusqueda(
    resultados,
    busqueda
) {

    if (!workerSearchResult) {

        return;

    }


    workerSearchResult.innerHTML =
        "";


    if (!resultados.length) {

        workerSearchResult.innerHTML = `

            <div class="search-no-results">

                <span>
                    🔎
                </span>

                <div>

                    <strong>
                        No se encontraron trabajadores
                    </strong>

                    <small>
                        No hay coincidencias para
                        "${escapar(busqueda)}"
                    </small>

                </div>

            </div>

        `;

        return;

    }


    const titulo =
        resultados.length === 1
            ? "Trabajador encontrado"
            : `${resultados.length} trabajadores encontrados`;


    const encabezado =
        document.createElement(
            "div"
        );


    encabezado.className =
        "search-results-header";


    encabezado.innerHTML = `

        <strong>
            ${titulo}
        </strong>

        <span>
            ${resultados.length}
        </span>

    `;


    workerSearchResult.appendChild(
        encabezado
    );


    const lista =
        document.createElement(
            "div"
        );


    lista.className =
        "search-results-list";


    resultados.forEach(
        user => {

            const nombre =
                user.displayName
                ||
                "Sin nombre";


            const correo =
                user.userPrincipalName
                ||
                "Sin correo";


            const departamento =
                String(
                    user.department || ""
                )
                .trim()
                ||
                "Sin departamento";


            const activo =
                user.accountEnabled !== false;


            const iniciales =
                obtenerIniciales(
                    nombre
                );


            const item =
                document.createElement(
                    "div"
                );


            item.className =
                "search-user";


            item.innerHTML = `

                <div
                    class="avatar ${
                        activo
                            ? "blue-avatar"
                            : "red-avatar"
                    }"
                >

                    ${escapar(
                        iniciales
                    )}

                </div>


                <div class="search-user-main">

                    <strong>
                        ${escapar(
                            nombre
                        )}
                    </strong>

                    <span>
                        ${escapar(
                            correo
                        )}
                    </span>

                </div>


                <div class="search-user-data">

                    <div>

                        <small>
                            Departamento
                        </small>

                        <strong>
                            ${escapar(
                                departamento
                            )}
                        </strong>

                    </div>


                    <div>

                        <small>
                            Estado
                        </small>

                        <strong class="${
                            activo
                                ? "search-active"
                                : "search-blocked"
                        }">

                            ${
                                activo
                                    ? "Activo"
                                    : "Bloqueado"
                            }

                        </strong>

                    </div>


                    <div>

                        <small>
                            ID
                        </small>

                        <strong>
                            ${escapar(
                                user.id || "—"
                            )}
                        </strong>

                    </div>

                </div>

            `;


            /*
               Al hacer clic en un resultado,
               lo mostramos también en la tabla
               principal de usuarios.
            */

            item.addEventListener(
                "click",
                () => {

                    mostrarTrabajadorEnTabla(
                        user
                    );

                }
            );


            lista.appendChild(
                item
            );

        }
    );


    workerSearchResult.appendChild(
        lista
    );

}


/* ============================================================
   MOSTRAR TRABAJADOR EN TABLA PRINCIPAL
============================================================ */

function mostrarTrabajadorEnTabla(
    user
) {

    const tbody =
        document.getElementById(
            "usersTable"
        );


    if (!tbody) {

        return;

    }


    tbody.innerHTML =
        "";


    renderUsuarioTabla(
        tbody,
        user
    );


    /*
       Ocultar paginación porque
       estamos mostrando un resultado
       específico.
    */

    actualizarPaginacion(
        "usersPagination",
        0,
        1,
        () => {}
    );


    /*
       Ir hacia la tabla de usuarios.
    */

    const tabla =
        document.querySelector(
            '[data-section="usuarios"]'
        );


    if (tabla) {

        tabla.scrollIntoView({
            behavior: "smooth",
            block: "start"
        });

    }

}


/* ============================================================
   LIMPIAR BÚSQUEDA
============================================================ */

function limpiarBusquedaTrabajador() {

    if (workerSearchResult) {

        workerSearchResult.innerHTML =
            "";

    }


    /*
       Volver a mostrar la tabla
       normalmente.
    */

    if (
        typeof actualizarUsuarios ===
        "function"
    ) {

        actualizarUsuarios();

    }

}


/* ============================================================
   EVENTO INPUT
============================================================ */

if (workerSearch) {

    workerSearch.addEventListener(
        "input",
        event => {

            buscarTrabajador(
                event.target.value
            );

        }
    );

}


/* ============================================================
   BOTÓN LIMPIAR
============================================================ */

if (clearWorkerSearch) {

    clearWorkerSearch.addEventListener(
        "click",
        () => {

            if (workerSearch) {

                workerSearch.value =
                    "";

                workerSearch.focus();

            }


            limpiarBusquedaTrabajador();

        }
    );

}