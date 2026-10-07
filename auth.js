/* =========================================================
   ALFERZA LIVE OFFICE
   CONTROL DE AUTENTICACIÓN Y PERMISOS
   ========================================================= */

/* =========================================================
   CONFIGURACIÓN
   ========================================================= */

const SESION_KEY = "alferza_login";
const RETURN_KEY = "alferza_return_url";
const PERMISOS_KEY = "alferza_permisos";

const ROOT_PATH = "/live-office/";
const LOGIN_PATH = "/live-office/login.html";
const BLANK_PATH = "/live-office/blank.html";

const CLIENT_ID = "5d98417c-74a7-4fab-8f2c-41ac127be696";
const TENANT_ID = "dbab984f-4bb1-4b60-9dff-da59f54acdf1";

const REDIRECT_URI =
    window.location.origin + BLANK_PATH;

const GRAPH_BASE =
    "https://graph.microsoft.com/v1.0";

const SHAREPOINT_HOST =
    "alferzaholding-my.sharepoint.com";

const SHAREPOINT_PATH =
    "/personal/soporte1_alferza_pe";

const PERMISOS_LIST_NAME =
    "PermisosTI";

const TI_ADMIN_EMAIL =
    "soporte1@alferza.pe";

/* =========================================================
   SCOPES
   ========================================================= */

const SCOPES = [
    "User.Read",
    "Sites.Read.All"
];

/* =========================================================
   MÓDULOS
   ========================================================= */

const MODULOS = [
    "Oficina",
    "Personal",
    "Reservas",
    "Salas",
    "Comunicados",
    "Seguridad",
    "Infraestructura",
    "Tickets",
    "Permisos",
    "Configuracion"
];

/* =========================================================
   MAPA DE PÁGINAS
   ========================================================= */

const PAGINAS_PERMISOS = {

    "index.html": "Oficina",

    "personal.html": "Personal",

    "reserva.html": "Reservas",
    "reservas.html": "Reservas",

    "salas/salas.html": "Salas",
    "salas.html": "Salas",

    "comunicados.html": "Comunicados",

    "seguridad.html": "Seguridad",

    "infraestructura.html": "Infraestructura",

    "tickets.html": "Tickets",

    "permisos.html": "Permisos",

    "configuracion.html": "Configuracion"
};

/* =========================================================
   MSAL
   =========================================================

   IMPORTANTE:
   Se usa authMsalInstance y NO msalInstance
   para evitar conflicto con app.js.
   ========================================================= */

let authMsalInstance = null;
let cuentaActual = null;
let permisosActuales = null;

/* =========================================================
   INICIO
   ========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    () => {

        iniciarControlPermisos();

    }
);

/* =========================================================
   CONTROL PRINCIPAL
   ========================================================= */

async function iniciarControlPermisos() {

    console.log("");
    console.log("========================================");
    console.log("ALFERZA LIVE OFFICE");
    console.log("Verificando permisos...");
    console.log("========================================");

    try {

        /* -----------------------------------------
           VERIFICAR SESIÓN
           ----------------------------------------- */

        if (
            sessionStorage.getItem(
                SESION_KEY
            ) !== "true"
        ) {

            console.warn(
                "No existe sesión ALFERZA."
            );

            guardarPaginaActual();

            window.location.replace(
                LOGIN_PATH
            );

            return;
        }

        /* -----------------------------------------
           CREAR MSAL
           ----------------------------------------- */

        await crearMSAL();

        /* -----------------------------------------
           CUENTA
           ----------------------------------------- */

        cuentaActual =
            obtenerCuentaActual();

        if (!cuentaActual) {

            console.error(
                "No se encontró una cuenta MSAL."
            );

            sessionStorage.removeItem(
                SESION_KEY
            );

            window.location.replace(
                LOGIN_PATH
            );

            return;
        }

        console.log(
            "Cuenta seleccionada:",
            cuentaActual
        );

        /* -----------------------------------------
           PERMISOS
           ----------------------------------------- */

        permisosActuales =
            await obtenerPermisosUsuario();

        console.log("");
        console.log("========================================");
        console.log("PERMISOS OBTENIDOS");
        console.table(
            permisosActuales
        );
        console.log("========================================");

        sessionStorage.setItem(
            PERMISOS_KEY,
            JSON.stringify(
                permisosActuales
            )
        );

        /* -----------------------------------------
           MENÚ
           ----------------------------------------- */

        aplicarPermisosMenu(
            permisosActuales
        );

        /* -----------------------------------------
           PÁGINA ACTUAL
           ----------------------------------------- */

        validarPaginaActual(
            permisosActuales
        );

    } catch (error) {

        console.error(
            "========================================"
        );

        console.error(
            "ERROR COMPROBANDO PERMISOS"
        );

        console.error(
            error
        );

        console.error(
            "========================================"
        );

        permisosActuales =
            crearPermisosDenegados();

        sessionStorage.setItem(
            PERMISOS_KEY,
            JSON.stringify(
                permisosActuales
            )
        );

        aplicarPermisosMenu(
            permisosActuales
        );

        validarPaginaActual(
            permisosActuales
        );
    }
}

/* =========================================================
   CREAR MSAL
   ========================================================= */

async function crearMSAL() {

    if (authMsalInstance) {
        return authMsalInstance;
    }

    authMsalInstance =
        new msal.PublicClientApplication({

            auth: {

                clientId:
                    CLIENT_ID,

                authority:
                    `https://login.microsoftonline.com/${TENANT_ID}`,

                redirectUri:
                    REDIRECT_URI,

                postLogoutRedirectUri:
                    window.location.origin +
                    ROOT_PATH +
                    "login.html",

                navigateToLoginRequestUrl:
                    false
            },

            cache: {

                cacheLocation:
                    "sessionStorage",

                storeAuthStateInCookie:
                    false
            }
        });

    await authMsalInstance.initialize();

    return authMsalInstance;
}

/* =========================================================
   OBTENER CUENTA
   ========================================================= */

function obtenerCuentaActual() {

    const cuentas =
        authMsalInstance.getAllAccounts();

    console.log(
        "Cuentas MSAL encontradas:",
        cuentas
    );

    if (!cuentas.length) {
        return null;
    }

    const cuentaTenant =
        cuentas.find(
            cuenta =>
                cuenta.tenantId === TENANT_ID
        );

    return cuentaTenant || cuentas[0];
}

/* =========================================================
   OBTENER TOKEN
   ========================================================= */

async function obtenerToken() {

    if (!cuentaActual) {

        cuentaActual =
            obtenerCuentaActual();
    }

    if (!cuentaActual) {

        throw new Error(
            "No existe cuenta MSAL."
        );
    }

    try {

        const response =
            await authMsalInstance.acquireTokenSilent({

                account:
                    cuentaActual,

                scopes:
                    SCOPES
            });

        console.log(
            "Token Microsoft obtenido correctamente."
        );

        return response.accessToken;

    } catch (error) {

        console.warn(
            "No se pudo obtener token silenciosamente.",
            error
        );

        const response =
            await authMsalInstance.acquireTokenPopup({

                scopes:
                    SCOPES
            });

        return response.accessToken;
    }
}

/* =========================================================
   GRAPH FETCH
   ========================================================= */

async function graphFetch(
    url,
    options = {}
) {

    const token =
        await obtenerToken();

    console.log(
        "Graph:",
        url
    );

    const response =
        await fetch(
            url,
            {
                ...options,

                headers: {

                    ...(options.headers || {}),

                    Authorization:
                        `Bearer ${token}`,

                    Accept:
                        "application/json"
                }
            }
        );

    if (!response.ok) {

        let detalle = "";

        try {

            detalle =
                await response.text();

        } catch (_) {}

        throw new Error(
            `Graph ${response.status}: ${detalle}`
        );
    }

    return response.json();
}

/* =========================================================
   OBTENER SITIO SHAREPOINT
   ========================================================= */

async function obtenerSitioSharePoint() {

    const url =
        `${GRAPH_BASE}/sites/` +
        `${SHAREPOINT_HOST}:${SHAREPOINT_PATH}`;

    const sitio =
        await graphFetch(url);

    console.log(
        "Sitio SharePoint encontrado:",
        sitio
    );

    return sitio;
}

/* =========================================================
   OBTENER LISTA PermisosTI
   ========================================================= */

async function obtenerListaPermisos(
    sitio
) {

    console.log("");
    console.log(
        "========================================"
    );

    console.log(
        "BUSCANDO LISTA DE PERMISOS"
    );

    console.log(
        "Nombre:",
        PERMISOS_LIST_NAME
    );

    console.log(
        "Sitio:",
        sitio.id
    );

    console.log(
        "========================================"
    );

    /* -----------------------------------------
       PRIMER INTENTO:
       Buscar mediante displayName
       ----------------------------------------- */

    const nombreCodificado =
        encodeURIComponent(
            PERMISOS_LIST_NAME
        );

    const urlFiltro =
        `${GRAPH_BASE}/sites/` +
        `${sitio.id}/lists` +
        `?$filter=displayName eq '${nombreCodificado}'` +
        `&$select=id,name,displayName,webUrl`;

    console.log(
        "Buscando PermisosTI mediante filtro:"
    );

    console.log(
        urlFiltro
    );

    try {

        const respuestaFiltro =
            await graphFetch(
                urlFiltro
            );

        const listasFiltro =
            respuestaFiltro.value || [];

        console.log(
            "Resultado del filtro:",
            listasFiltro
        );

        if (
            listasFiltro.length > 0
        ) {

            const lista =
                listasFiltro[0];

            console.log(
                "PermisosTI encontrada mediante filtro:",
                lista
            );

            return lista;
        }

    } catch (error) {

        console.warn(
            "Falló búsqueda mediante filtro:",
            error
        );
    }

    /* -----------------------------------------
       SEGUNDO INTENTO:
       Enumerar todas las listas
       ----------------------------------------- */

    console.log(
        "Enumerando listas del sitio..."
    );

    const urlListas =
        `${GRAPH_BASE}/sites/` +
        `${sitio.id}/lists` +
        `?$select=id,name,displayName,webUrl` +
        `&$top=200`;

    const respuesta =
        await graphFetch(
            urlListas
        );

    const listas =
        respuesta.value || [];

    console.log(
        "Listas encontradas:",
        listas
    );

    console.table(
        listas.map(
            lista => ({

                id:
                    lista.id,

                name:
                    lista.name,

                displayName:
                    lista.displayName,

                webUrl:
                    lista.webUrl

            })
        )
    );

    const objetivo =
        normalizarTexto(
            PERMISOS_LIST_NAME
        );

    const listaEncontrada =
        listas.find(
            lista => {

                const name =
                    normalizarTexto(
                        lista.name
                    );

                const displayName =
                    normalizarTexto(
                        lista.displayName
                    );

                return (
                    name === objetivo ||
                    displayName === objetivo
                );
            }
        );

    if (listaEncontrada) {

        console.log(
            "PermisosTI encontrada:",
            listaEncontrada
        );

        return listaEncontrada;
    }

    /* -----------------------------------------
       NO ENCONTRADA
       ----------------------------------------- */

    console.error(
        "========================================"
    );

    console.error(
        "NO SE ENCONTRÓ PermisosTI"
    );

    console.error(
        "Graph está viendo estas listas:"
    );

    console.table(
        listas.map(
            lista => ({

                ID:
                    lista.id,

                Nombre:
                    lista.name,

                DisplayName:
                    lista.displayName,

                URL:
                    lista.webUrl

            })
        )
    );

    console.error(
        "========================================"
    );

    throw new Error(
        `No se encontró la lista "${PERMISOS_LIST_NAME}" mediante Microsoft Graph.`
    );
}

/* =========================================================
   OBTENER PERMISOS DEL USUARIO
   ========================================================= */

async function obtenerPermisosUsuario() {

    const cuenta =
        obtenerCuentaActual();

    if (!cuenta) {

        throw new Error(
            "No se encontró la cuenta actual."
        );
    }

    const correo =
        extraerCorreo(
            cuenta
        );

    console.log("");
    console.log(
        "========================================"
    );

    console.log(
        "USUARIO ACTUAL:",
        correo
    );

    console.log(
        "========================================"
    );

    /* -----------------------------------------
       ADMIN TI PRINCIPAL
       ----------------------------------------- */

    if (
        normalizarCorreo(correo) ===
        normalizarCorreo(TI_ADMIN_EMAIL)
    ) {

        console.log(
            "Usuario TI principal detectado."
        );

        return crearPermisosTI();
    }

    /* -----------------------------------------
       SITIO
       ----------------------------------------- */

    const sitio =
        await obtenerSitioSharePoint();

    /* -----------------------------------------
       LISTA
       ----------------------------------------- */

    const lista =
        await obtenerListaPermisos(
            sitio
        );

    console.log(
        "ID REAL DE PermisosTI:",
        lista.id
    );

    /* -----------------------------------------
       ITEMS
       ----------------------------------------- */

    const urlItems =
        `${GRAPH_BASE}/sites/` +
        `${sitio.id}/lists/` +
        `${lista.id}/items` +
        `?$expand=fields&$top=500`;

    const respuesta =
        await graphFetch(
            urlItems
        );

    const items =
        respuesta.value || [];

    console.log(
        "Registros encontrados en PermisosTI:",
        items
    );

    console.log(
        "Cantidad de registros:",
        items.length
    );

    /* -----------------------------------------
       BUSCAR USUARIO
       ----------------------------------------- */

    const usuario =
        items.find(
            item => {

                const fields =
                    item.fields || {};

                const correos = [

                    fields.UsuarioCorreo,

                    fields.UsuarioCorreo0,

                    fields.Correo,

                    fields.Email,

                    fields.Usuario,

                    fields.Title
                ];

                return correos.some(
                    valor => {

                        const correoCampo =
                            extraerCorreo(
                                valor
                            );

                        return (

                            correoCampo &&

                            normalizarCorreo(
                                correoCampo
                            ) ===

                            normalizarCorreo(
                                correo
                            )
                        );
                    }
                );
            }
        );

    /* -----------------------------------------
       USUARIO NO ENCONTRADO
       ----------------------------------------- */

    if (!usuario) {

        console.warn(
            "USUARIO NO ENCONTRADO EN PermisosTI:",
            correo
        );

        console.table(
            items.map(
                item => {

                    const fields =
                        item.fields || {};

                    return {

                        ID:
                            item.id,

                        UsuarioCorreo:
                            fields.UsuarioCorreo,

                        UsuarioCorreo0:
                            fields.UsuarioCorreo0,

                        Correo:
                            fields.Correo,

                        Email:
                            fields.Email,

                        Usuario:
                            fields.Usuario,

                        Title:
                            fields.Title
                    };
                }
            )
        );

        return crearPermisosDenegados();
    }

    /* -----------------------------------------
       REGISTRO ENCONTRADO
       ----------------------------------------- */

    console.log("");
    console.log(
        "REGISTRO DEL USUARIO ENCONTRADO:"
    );

    console.log(
        usuario
    );

    console.log(
        "FIELDS:",
        usuario.fields
    );

    /* -----------------------------------------
       CONSTRUIR PERMISOS
       ----------------------------------------- */

    const fields =
        usuario.fields || {};

    const permisos = {};

    MODULOS.forEach(
        modulo => {

            const valor =
                obtenerCampoFlexible(
                    fields,
                    modulo
                );

            permisos[modulo] =
                valorBooleano(
                    valor
                );

            console.log(
                `Permiso ${modulo}:`,
                permisos[modulo],
                "valor original:",
                valor
            );
        }
    );

    return permisos;
}

/* =========================================================
   CREAR PERMISOS TI
   ========================================================= */

function crearPermisosTI() {

    const permisos = {};

    MODULOS.forEach(
        modulo => {

            permisos[modulo] =
                true;
        }
    );

    return permisos;
}

/* =========================================================
   CREAR PERMISOS DENEGADOS
   ========================================================= */

function crearPermisosDenegados() {

    const permisos = {};

    MODULOS.forEach(
        modulo => {

            permisos[modulo] =
                false;
        }
    );

    return permisos;
}

/* =========================================================
   APLICAR PERMISOS AL MENÚ
   ========================================================= */

function aplicarPermisosMenu(
    permisos
) {

    const elementos =
        document.querySelectorAll(
            "[data-modulo], .menu-item, .nav-item, aside a"
        );

    console.log(
        "Elementos del menú encontrados:",
        elementos.length
    );

    elementos.forEach(
        elemento => {

            let modulo =
                elemento.dataset.modulo;

            if (!modulo) {

                const href =
                    elemento.getAttribute(
                        "href"
                    );

                modulo =
                    encontrarPermisoPagina(
                        href
                    );
            }

            if (!modulo) {
                return;
            }

            const permiso =
                permisos[modulo] === true;

            console.log(
                `Menú ${modulo}:`,
                permiso
            );

            if (!permiso) {

                elemento.style.display =
                    "none";
            }
        }
    );
}

/* =========================================================
   OBTENER NOMBRE DEL MÓDULO
   ========================================================= */

function obtenerNombreModulo(
    modulo
) {

    if (!modulo) {
        return null;
    }

    const texto =
        normalizarTexto(
            modulo
        );

    const equivalencias = {

        "oficina":
            "Oficina",

        "personal":
            "Personal",

        "reservas":
            "Reservas",

        "reserva":
            "Reservas",

        "salas":
            "Salas",

        "comunicados":
            "Comunicados",

        "seguridad":
            "Seguridad",

        "infraestructura":
            "Infraestructura",

        "tickets":
            "Tickets",

        "permisos":
            "Permisos",

        "configuracion":
            "Configuracion"
    };

    return (
        equivalencias[texto] ||
        modulo
    );
}

/* =========================================================
   VALIDAR PÁGINA ACTUAL
   ========================================================= */

function validarPaginaActual(
    permisos
) {

    const pagina =
        obtenerPaginaActual();

    console.log(
        "Página actual:",
        pagina
    );

    const modulo =
        encontrarPermisoPagina(
            pagina
        );

    if (!modulo) {

        console.log(
            "Página sin permiso específico."
        );

        return;
    }

    const permiso =
        permisos[modulo] === true;

    console.log(
        `Validación ${modulo}:`,
        permiso
    );

    if (!permiso) {

        console.warn(
            `Acceso denegado al módulo: ${modulo}`
        );

        bloquearPagina(
            modulo
        );

        return;
    }

    console.log(
        `Acceso permitido al módulo: ${modulo}`
    );
}

/* =========================================================
   OBTENER PÁGINA ACTUAL
   ========================================================= */

function obtenerPaginaActual() {

    let ruta =
        window.location.pathname;

    ruta =
        ruta.replace(
            ROOT_PATH,
            ""
        );

    ruta =
        ruta.replace(
            /^\/+/,
            ""
        );

    if (!ruta) {
        return "index.html";
    }

    return ruta;
}

/* =========================================================
   ENCONTRAR PERMISO DE PÁGINA
   ========================================================= */

function encontrarPermisoPagina(
    pagina
) {

    if (!pagina) {
        return null;
    }

    let ruta =
        pagina;

    ruta =
        ruta.replace(
            /^\/+/,
            ""
        );

    ruta =
        ruta.replace(
            ROOT_PATH.replace(
                /^\/+/,
                ""
            ),
            ""
        );

    ruta =
        ruta.toLowerCase();

    for (
        const [
            rutaMapa,
            modulo
        ]
        of Object.entries(
            PAGINAS_PERMISOS
        )
    ) {

        if (
            rutaMapa.toLowerCase() ===
            ruta
        ) {

            return modulo;
        }
    }

    if (
        ruta === "" ||
        ruta === "/"
    ) {

        return "Oficina";
    }

    return null;
}

/* =========================================================
   BLOQUEAR PÁGINA
   ========================================================= */

function bloquearPagina(
    modulo
) {

    const contenido =
        document.body;

    if (!contenido) {
        return;
    }

    contenido.innerHTML = `

        <div style="
            min-height:100vh;
            display:flex;
            align-items:center;
            justify-content:center;
            background:#f4f6f9;
            font-family:Arial,sans-serif;
            padding:30px;
            box-sizing:border-box;
        ">

            <div style="
                max-width:520px;
                width:100%;
                background:#ffffff;
                border-radius:18px;
                padding:40px;
                text-align:center;
                box-shadow:0 10px 35px rgba(0,0,0,.10);
            ">

                <div style="
                    font-size:48px;
                    margin-bottom:20px;
                ">
                    🔒
                </div>

                <h2 style="
                    margin:0 0 12px;
                    color:#00205c;
                ">
                    Acceso restringido
                </h2>

                <p style="
                    margin:0 0 25px;
                    color:#555;
                    line-height:1.6;
                ">
                    No tienes permisos para acceder a
                    <strong>${modulo}</strong>.
                </p>

                <button
                    onclick="window.location.href='${ROOT_PATH}index.html'"
                    style="
                        border:0;
                        border-radius:10px;
                        padding:12px 22px;
                        background:#00205c;
                        color:#ffffff;
                        font-size:15px;
                        font-weight:600;
                        cursor:pointer;
                    "
                >
                    Volver a Oficina
                </button>

            </div>

        </div>

    `;
}

/* =========================================================
   OBTENER CAMPO FLEXIBLE
   ========================================================= */

function obtenerCampoFlexible(
    fields,
    nombre
) {

    if (!fields) {
        return null;
    }

    const objetivo =
        normalizarTexto(
            nombre
        );

    /* -----------------------------------------
       COINCIDENCIA EXACTA
       ----------------------------------------- */

    if (
        Object.prototype.hasOwnProperty.call(
            fields,
            nombre
        )
    ) {

        return fields[nombre];
    }

    /* -----------------------------------------
       COINCIDENCIA NORMALIZADA
       ----------------------------------------- */

    const clave =
        Object.keys(fields)
            .find(
                key =>
                    normalizarTexto(
                        key
                    ) === objetivo
            );

    if (clave) {

        return fields[clave];
    }

    /* -----------------------------------------
       ALIASES
       ----------------------------------------- */

    const aliases = {

        "oficina": [

            "Oficina",
            "Oficina0",
            "AccesoOficina",
            "PermisoOficina"
        ],

        "personal": [

            "Personal",
            "Personal0",
            "AccesoPersonal",
            "PermisoPersonal"
        ],

        "reservas": [

            "Reservas",
            "Reserva",
            "Reservas0",
            "AccesoReservas",
            "PermisoReservas"
        ],

        "salas": [

            "Salas",
            "Sala",
            "Salas0",
            "AccesoSalas",
            "PermisoSalas"
        ],

        "comunicados": [

            "Comunicados",
            "Comunicado",
            "Comunicados0",
            "AccesoComunicados",
            "PermisoComunicados"
        ],

        "seguridad": [

            "Seguridad",
            "Seguridad0",
            "AccesoSeguridad",
            "PermisoSeguridad"
        ],

        "infraestructura": [

            "Infraestructura",
            "Infraestructura0",
            "AccesoInfraestructura",
            "PermisoInfraestructura"
        ],

        "tickets": [

            "Tickets",
            "Ticket",
            "Tickets0",
            "AccesoTickets",
            "PermisoTickets"
        ],

        "permisos": [

            "Permisos",
            "Permiso",
            "Permisos0",
            "AccesoPermisos",
            "PermisoPermisos"
        ],

        "configuracion": [

            "Configuracion",
            "Configuración",
            "Configuracion0",
            "Configuración0",
            "AccesoConfiguracion",
            "PermisoConfiguracion"
        ]
    };

    const listaAliases =
        aliases[objetivo] || [];

    for (
        const alias
        of listaAliases
    ) {

        if (
            Object.prototype.hasOwnProperty.call(
                fields,
                alias
            )
        ) {

            return fields[alias];
        }

        const aliasNormalizado =
            normalizarTexto(
                alias
            );

        const claveAlias =
            Object.keys(fields)
                .find(
                    key =>
                        normalizarTexto(
                            key
                        ) ===
                        aliasNormalizado
                );

        if (claveAlias) {

            return fields[
                claveAlias
            ];
        }
    }

    return null;
}

/* =========================================================
   CONVERTIR VALOR A BOOLEAN
   ========================================================= */

function valorBooleano(
    valor
) {

    if (
        valor === true ||
        valor === 1
    ) {

        return true;
    }

    if (
        typeof valor === "string"
    ) {

        const texto =
            valor
                .trim()
                .toLowerCase();

        return [

            "true",
            "1",
            "si",
            "sí",
            "yes",
            "activo",
            "permitido",
            "permitida",
            "habilitado",
            "habilitada"

        ].includes(
            texto
        );
    }

    return false;
}

/* =========================================================
   EXTRAER CORREO
   ========================================================= */

function extraerCorreo(
    valor
) {

    if (!valor) {
        return "";
    }

    if (
        typeof valor === "string"
    ) {

        return valor
            .trim()
            .toLowerCase();
    }

    if (
        typeof valor === "object"
    ) {

        return (

            valor.email ||
            valor.Email ||
            valor.mail ||
            valor.Mail ||
            valor.userPrincipalName ||
            valor.UserPrincipalName ||
            ""

        )
            .toString()
            .trim()
            .toLowerCase();
    }

    return "";
}

/* =========================================================
   NORMALIZAR CORREO
   ========================================================= */

function normalizarCorreo(
    correo
) {

    if (!correo) {
        return "";
    }

    return correo
        .toString()
        .trim()
        .toLowerCase();
}

/* =========================================================
   NORMALIZAR TEXTO
   ========================================================= */

function normalizarTexto(
    texto
) {

    if (
        texto === null ||
        texto === undefined
    ) {

        return "";
    }

    return texto
        .toString()
        .trim()
        .toLowerCase()
        .normalize("NFD")
        .replace(
            /[\u0300-\u036f]/g,
            ""
        )
        .replace(
            /[\s_-]+/g,
            ""
        );
}

/* =========================================================
   GUARDAR PÁGINA ACTUAL
   ========================================================= */

function guardarPaginaActual() {

    try {

        sessionStorage.setItem(
            RETURN_KEY,

            window.location.pathname +
            window.location.search +
            window.location.hash
        );

    } catch (error) {

        console.warn(
            "No se pudo guardar la página actual.",
            error
        );
    }
}

/* =========================================================
   FUNCIONES GLOBALES
   ========================================================= */

window.AlferzaAuth = {

    obtenerPermisos:
        () =>
            permisosActuales,

    obtenerCuenta:
        () =>
            cuentaActual,

    tienePermiso:
        modulo =>
            permisosActuales &&
            permisosActuales[
                obtenerNombreModulo(
                    modulo
                )
            ] === true
};