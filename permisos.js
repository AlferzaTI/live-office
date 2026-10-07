/* =========================================================
   ALFERZA LIVE OFFICE
   GESTIÓN DE PERMISOS
========================================================= */


/* =========================================================
   CONFIGURACIÓN
========================================================= */

const PERMISOS_LIST_NAME = "PermisosTI";

const SHAREPOINT_HOST =
    "alferzaholding-my.sharepoint.com";

const SHAREPOINT_SITE_PATH =
    "/personal/soporte1_alferza_pe";

const TENANT_ID =
    "dbab984f-4bb1-4b60-9dff-da59f54acdf1";

const CLIENT_ID =
    "5d98417c-74a7-4fab-8f2c-41ac127be696";

const REDIRECT_URI =
    "https://alferzati.github.io/live-office/blank.html";

const GRUPO_TI = "TI";
const GRUPO_NORMAL = "Normal";

const USUARIOS_TI_INICIALES = [
    "soporte1@alferza.pe"
];

const MODULOS_PERMISOS = [
    "Oficina",
    "Personal",
    "Reservas",
    "Salas",
    "Comunicados",
    "Seguridad",
    "Infraestructura",
    "Tickets",
    "Configuracion"
];


/* =========================================================
   SCOPES GRAPH
========================================================= */

const SCOPES_GRAPH = [
    "User.Read",
    "User.ReadBasic.All",
    "Sites.ReadWrite.All"
];


/* =========================================================
   VARIABLES GLOBALES
========================================================= */

let permisosMsalInstance = null;

let usuarioSeleccionado = null;

let registrosPermisos = [];

let sitioSharePointCache = null;

let listaPermisosCache = null;

let columnasPermisosCache = null;

let mapaCamposCache = null;


/* =========================================================
   INICIO
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    async () => {

        console.log(
            "Permisos cargado"
        );


        const contenido =
            document.querySelector(
                ".content"
            );


        if (contenido) {

            contenido.style.display =
                "block";

            contenido.style.visibility =
                "visible";

            contenido.style.opacity =
                "1";
        }


        configurarMenu();

        configurarEventos();

        await inicializarPermisos();

    }
);


/* =========================================================
   INICIALIZAR
========================================================= */

async function inicializarPermisos() {

    try {

        const cuenta =
            await obtenerCuentaActual();


        if (!cuenta) {

            mostrarMensaje(
                "No se pudo obtener la sesión de Microsoft. Vuelve a iniciar sesión.",
                "error"
            );

            return;
        }


        const correoActual =
            obtenerCorreoCuenta(
                cuenta
            );


        console.log(
            "Cuenta MSAL:",
            correoActual
        );


        const userInfo =
            document.getElementById(
                "userInfo"
            );


        if (userInfo) {

            userInfo.textContent =
                cuenta.name ||
                correoActual ||
                "";
        }


        /*
         * Intentar obtener también el
         * usuario real mediante Graph.
         */

        const usuarioGraph =
            await obtenerUsuarioActualGraph();


        if (usuarioGraph) {

            console.log(
                "Usuario actual desde Graph:",
                usuarioGraph
            );


            if (userInfo) {

                userInfo.textContent =
                    usuarioGraph.displayName ||
                    usuarioGraph.mail ||
                    usuarioGraph.userPrincipalName ||
                    correoActual;
            }
        }


        /*
         * Cargar los registros de permisos.
         */

        await cargarPermisos();


        /*
         * Comprobar si el usuario actual
         * tiene grupo TI.
         */

        const registroActual =
            buscarRegistroPorCorreo(
                correoActual
            );


        const esTI =
            USUARIOS_TI_INICIALES
                .map(normalizarCorreo)
                .includes(
                    normalizarCorreo(
                        correoActual
                    )
                )
            ||
            (
                registroActual &&
                String(
                    obtenerCampo(
                        registroActual,
                        "Grupo"
                    )
                )
                    .trim()
                    .toLowerCase() ===
                    GRUPO_TI.toLowerCase()
            );


        if (!esTI) {

            mostrarAccesoDenegado();

            return;
        }


        console.log(
            "Usuario autorizado para administrar permisos."
        );


    } catch (error) {

        console.error(
            "Error en permisos:",
            error
        );


        mostrarMensaje(
            obtenerMensajeError(
                error
            ),
            "error"
        );

    }
}


/* =========================================================
   CONFIGURAR MSAL
========================================================= */

async function obtenerInstanciaMSAL() {

    /*
     * Por seguridad, si alguna página ya
     * expone una instancia global, la usamos.
     */

    if (
        window.msalInstance
    ) {

        return window.msalInstance;
    }


    /*
     * Tu auth.js NO crea MSAL.
     * Por eso esta página crea su propia
     * instancia usando la misma aplicación.
     */

    if (
        !window.msal ||
        !window.msal.PublicClientApplication
    ) {

        throw new Error(
            "La librería MSAL Browser no está cargada."
        );
    }


    if (
        !permisosMsalInstance
    ) {

        permisosMsalInstance =
            new window.msal.PublicClientApplication({

                auth: {

                    clientId:
                        CLIENT_ID,

                    authority:
                        `https://login.microsoftonline.com/${TENANT_ID}`,

                    redirectUri:
                        REDIRECT_URI

                },

                cache: {

                    cacheLocation:
                        "sessionStorage",

                    storeAuthStateInCookie:
                        false
                }

            });


        /*
         * Recuperar la cuenta que ya tiene
         * la sesión iniciada.
         */

        const cuentas =
            permisosMsalInstance.getAllAccounts();


        if (
            cuentas &&
            cuentas.length > 0
        ) {

            permisosMsalInstance.setActiveAccount(
                cuentas[0]
            );
        }
    }


    return permisosMsalInstance;
}


/* =========================================================
   OBTENER CUENTA ACTUAL
========================================================= */

async function obtenerCuentaActual() {

    const instancia =
        await obtenerInstanciaMSAL();


    let cuenta =
        instancia.getActiveAccount();


    if (!cuenta) {

        const cuentas =
            instancia.getAllAccounts();


        if (
            cuentas &&
            cuentas.length > 0
        ) {

            cuenta =
                cuentas[0];


            instancia.setActiveAccount(
                cuenta
            );
        }
    }


    if (!cuenta) {

        return null;
    }


    return cuenta;
}


/* =========================================================
   CORREO DE CUENTA
========================================================= */

function obtenerCorreoCuenta(
    cuenta
) {

    if (!cuenta) {

        return "";
    }


    return normalizarCorreo(

        cuenta.username ||
        cuenta.mail ||
        cuenta.userPrincipalName ||
        ""

    );
}


/* =========================================================
   TOKEN GRAPH
========================================================= */

async function obtenerTokenGraph() {

    const instancia =
        await obtenerInstanciaMSAL();


    const cuenta =
        await obtenerCuentaActual();


    if (!cuenta) {

        throw new Error(
            "No existe una cuenta Microsoft activa en esta página. Cierra sesión y vuelve a iniciar sesión."
        );
    }


    try {

        /*
         * Primero intenta reutilizar
         * el token existente.
         */

        const resultado =
            await instancia.acquireTokenSilent({

                scopes:
                    SCOPES_GRAPH,

                account:
                    cuenta

            });


        return resultado.accessToken;


    } catch (silentError) {

        console.warn(
            "acquireTokenSilent no pudo obtener el token:",
            silentError
        );


        /*
         * Si hace falta consentimiento,
         * abre el popup de Microsoft.
         */

        try {

            const resultado =
                await instancia.acquireTokenPopup({

                    scopes:
                        SCOPES_GRAPH,

                    account:
                        cuenta

                });


            return resultado.accessToken;


        } catch (popupError) {

            console.error(
                "Error obteniendo token Graph:",
                popupError
            );


            throw popupError;
        }
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
        await obtenerTokenGraph();


    const headers = {

        Authorization:
            `Bearer ${token}`,

        Accept:
            "application/json",

        ...(options.headers || {})

    };


    if (
        options.body &&
        !headers["Content-Type"]
    ) {

        headers["Content-Type"] =
            "application/json";
    }


    const response =
        await fetch(

            url,

            {

                ...options,

                headers

            }

        );


    if (!response.ok) {

        let detalle = "";


        try {

            const errorData =
                await response.json();


            detalle =
                errorData?.error?.message ||
                "";

        } catch (error) {

            // No hay cuerpo JSON.
        }


        throw new Error(

            `Graph ${response.status}: ${
                detalle ||
                response.statusText
            }`

        );
    }


    if (
        response.status === 204
    ) {

        return null;
    }


    return await response.json();
}


/* =========================================================
   USUARIO ACTUAL DESDE GRAPH
========================================================= */

async function obtenerUsuarioActualGraph() {

    try {

        return await graphFetch(

            "https://graph.microsoft.com/v1.0/me?$select=id,displayName,mail,userPrincipalName"

        );

    } catch (error) {

        console.warn(
            "No se pudo obtener /me:",
            error
        );


        return null;
    }
}


/* =========================================================
   OBTENER SITIO SHAREPOINT
========================================================= */

async function obtenerSitioSharePoint() {

    if (
        sitioSharePointCache
    ) {

        return sitioSharePointCache;
    }


    const url =
        `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOST}:${SHAREPOINT_SITE_PATH}`;


    sitioSharePointCache =
        await graphFetch(
            url
        );


    return sitioSharePointCache;
}


/* =========================================================
   OBTENER LISTA
========================================================= */

async function obtenerListaPermisos() {

    if (
        listaPermisosCache
    ) {

        return listaPermisosCache;
    }


    const sitio =
        await obtenerSitioSharePoint();


    const url =
        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${encodeURIComponent(PERMISOS_LIST_NAME)}`;


    listaPermisosCache =
        await graphFetch(
            url
        );


    return listaPermisosCache;
}


/* =========================================================
   OBTENER COLUMNAS
========================================================= */

async function obtenerColumnasPermisos() {

    if (
        columnasPermisosCache
    ) {

        return columnasPermisosCache;
    }


    const sitio =
        await obtenerSitioSharePoint();


    const lista =
        await obtenerListaPermisos();


    const url =
        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/columns?$select=name,displayName,hidden,readOnly&$top=200`;


    const resultado =
        await graphFetch(
            url
        );


    columnasPermisosCache =
        resultado.value || [];


    return columnasPermisosCache;
}


/* =========================================================
   MAPA DE CAMPOS
========================================================= */

async function obtenerMapaCampos() {

    if (
        mapaCamposCache
    ) {

        return mapaCamposCache;
    }


    const columnas =
        await obtenerColumnasPermisos();


    const mapa = {};


    for (
        const modulo of
        MODULOS_PERMISOS
    ) {

        const columna =
            encontrarColumna(

                columnas,

                [
                    modulo
                ]

            );


        if (!columna) {

            throw new Error(

                `No se encontró la columna de Microsoft Lists: ${modulo}`

            );
        }


        mapa[modulo] =
            columna.name;
    }


    const titleColumn =
        encontrarColumna(

            columnas,

            [
                "Title",
                "Título"
            ]

        );


    if (!titleColumn) {

        throw new Error(
            "No se encontró la columna Title/Título."
        );
    }


    mapa.Title =
        titleColumn.name;


    for (
        const nombre of
        [
            "UsuarioCorreo",
            "NombreUsuario",
            "Grupo"
        ]
    ) {

        const columna =
            encontrarColumna(

                columnas,

                [
                    nombre
                ]

            );


        if (!columna) {

            throw new Error(

                `No se encontró la columna: ${nombre}`

            );
        }


        mapa[nombre] =
            columna.name;
    }


    mapaCamposCache =
        mapa;


    console.log(
        "Mapa de columnas PermisosTI:",
        mapa
    );


    return mapa;
}


/* =========================================================
   ENCONTRAR COLUMNA
========================================================= */

function encontrarColumna(
    columnas,
    nombres
) {

    const objetivos =
        nombres.map(
            normalizarTexto
        );


    return (

        columnas.find(
            columna =>
                objetivos.includes(
                    normalizarTexto(
                        columna.displayName
                    )
                )
        )

        ||

        columnas.find(
            columna =>
                objetivos.includes(
                    normalizarTexto(
                        columna.name
                    )
                )
        )

        ||

        null
    );
}


/* =========================================================
   OBTENER TODOS LOS ITEMS
========================================================= */

async function obtenerTodosLosItems() {

    const sitio =
        await obtenerSitioSharePoint();


    const lista =
        await obtenerListaPermisos();


    let url =
        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items?expand=fields&$top=500`;


    const items = [];


    while (url) {

        const resultado =
            await graphFetch(
                url
            );


        if (
            Array.isArray(
                resultado.value
            )
        ) {

            items.push(
                ...resultado.value
            );
        }


        url =
            resultado["@odata.nextLink"] ||
            null;
    }


    return items;
}


/* =========================================================
   CARGAR PERMISOS
========================================================= */

async function cargarPermisos() {

    const tabla =
        document.getElementById(
            "tablaPermisos"
        );


    if (!tabla) {

        return [];
    }


    tabla.innerHTML =
        `
        <div class="tabla-vacia">

            <span>
                Cargando usuarios...
            </span>

        </div>
        `;


    try {

        registrosPermisos =
            await obtenerTodosLosItems();


        renderizarTablaPermisos(
            registrosPermisos
        );


        return registrosPermisos;


    } catch (error) {

        console.error(
            "Error cargando PermisosTI:",
            error
        );


        tabla.innerHTML =
            `
            <div class="tabla-vacia">

                <span>
                    No se pudieron cargar los permisos.
                </span>

            </div>
            `;


        throw error;
    }
}


/* =========================================================
   RENDERIZAR TABLA
========================================================= */

function renderizarTablaPermisos(
    registros
) {

    const tabla =
        document.getElementById(
            "tablaPermisos"
        );


    if (!tabla) {

        return;
    }


    if (
        !registros ||
        registros.length === 0
    ) {

        tabla.innerHTML =
            `
            <div class="tabla-vacia">

                <span>
                    No hay usuarios configurados.
                </span>

            </div>
            `;

        return;
    }


    let filas = "";


    registros.forEach(
        registro => {

            const correo =
                obtenerCampo(
                    registro,
                    "UsuarioCorreo"
                ) ||
                "";


            const nombre =
                obtenerCampo(
                    registro,
                    "NombreUsuario"
                ) ||

                obtenerCampo(
                    registro,
                    "Title"
                ) ||

                correo ||

                "Sin nombre";


            const grupo =
                String(

                    obtenerCampo(
                        registro,
                        "Grupo"
                    ) ||

                    GRUPO_NORMAL

                )
                    .trim();


            const cantidad =
                contarPermisos(
                    registro
                );


            const esCompleto =

                grupo.toLowerCase() ===
                    GRUPO_TI.toLowerCase()

                ||

                cantidad ===
                    MODULOS_PERMISOS.length;


            const accesoTexto =

                grupo.toLowerCase() ===
                    GRUPO_TI.toLowerCase()

                    ? "Acceso total"

                    : `${cantidad}/${MODULOS_PERMISOS.length}`;


            const claseGrupo =

                grupo.toLowerCase() ===
                    GRUPO_TI.toLowerCase()

                    ? "badge-ti"

                    : "badge-normal";


            filas +=
                `
                <tr>

                    <td>

                        <div class="permission-user-cell">

                            <div class="permission-user-avatar">
                                👤
                            </div>

                            <div class="permission-user-info">

                                <strong>
                                    ${escapeHtml(nombre)}
                                </strong>

                                <span>
                                    ${escapeHtml(correo)}
                                </span>

                            </div>

                        </div>

                    </td>


                    <td>

                        <span class="badge ${claseGrupo}">
                            ${escapeHtml(grupo)}
                        </span>

                    </td>


                    <td>

                        <span
                            class="
                                access-summary
                                ${
                                    esCompleto
                                        ? "full"
                                        : "partial"
                                }
                            "
                        >
                            ${escapeHtml(
                                accesoTexto
                            )}
                        </span>

                    </td>


                    <td>

                        <button

                            type="button"

                            class="table-action-btn"

                            data-editar-permiso="${escapeHtml(
                                registro.id
                            )}"

                        >
                            Editar
                        </button>

                    </td>

                </tr>
                `;
        }
    );


    tabla.innerHTML =
        `
        <div class="permissions-table-container">

            <table class="permissions-table">

                <thead>

                    <tr>

                        <th>
                            Usuario
                        </th>

                        <th>
                            Grupo
                        </th>

                        <th>
                            Acceso
                        </th>

                        <th>
                            Acción
                        </th>

                    </tr>

                </thead>

                <tbody>

                    ${filas}

                </tbody>

            </table>

        </div>
        `;
}


/* =========================================================
   OBTENER CAMPO
========================================================= */

function obtenerCampo(
    registro,
    nombreBuscado
) {

    if (!registro) {

        return "";
    }


    const fields =
        registro.fields ||
        {};


    const objetivo =
        normalizarTexto(
            nombreBuscado
        );


    const clave =
        Object.keys(
            fields
        ).find(
            key =>
                normalizarTexto(
                    key
                ) === objetivo
        );


    if (!clave) {

        return "";
    }


    return fields[clave];
}


/* =========================================================
   CONTAR PERMISOS
========================================================= */

function contarPermisos(
    registro
) {

    const grupo =
        String(

            obtenerCampo(
                registro,
                "Grupo"
            ) ||

            ""

        )
            .trim()
            .toLowerCase();


    if (
        grupo ===
        GRUPO_TI.toLowerCase()
    ) {

        return MODULOS_PERMISOS.length;
    }


    let cantidad = 0;


    MODULOS_PERMISOS.forEach(
        modulo => {

            if (
                valorBooleano(
                    obtenerCampo(
                        registro,
                        modulo
                    )
                )
            ) {

                cantidad++;
            }
        }
    );


    return cantidad;
}


/* =========================================================
   BUSCAR REGISTRO POR CORREO
========================================================= */

function buscarRegistroPorCorreo(
    correo
) {

    const objetivo =
        normalizarCorreo(
            correo
        );


    return registrosPermisos.find(
        registro =>

            normalizarCorreo(

                obtenerCampo(
                    registro,
                    "UsuarioCorreo"
                )

            ) === objetivo

    ) || null;
}


/* =========================================================
   BUSCAR USUARIO EN GRAPH
========================================================= */

async function buscarUsuario() {

    const input =
        document.getElementById(
            "correoUsuario"
        );


    const btn =
        document.getElementById(
            "btnBuscarUsuario"
        );


    if (!input) {

        return;
    }


    const correo =
        normalizarCorreo(
            input.value
        );


    if (!correo) {

        alert(
            "Ingresa un correo corporativo."
        );

        return;
    }


    if (
        !correo.endsWith(
            "@alferza.pe"
        )
    ) {

        alert(
            "Ingresa un correo corporativo de @alferza.pe."
        );

        return;
    }


    try {

        if (btn) {

            btn.disabled =
                true;

            btn.textContent =
                "Buscando...";
        }


        ocultarPanelUsuario();


        usuarioSeleccionado =
            null;


        /*
         * Buscar directamente mediante UPN.
         */

        const url =
            `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(correo)}?$select=id,displayName,mail,userPrincipalName`;


        let usuario =
            null;


        try {

            usuario =
                await graphFetch(
                    url
                );


        } catch (errorDirecto) {

            console.warn(

                "La búsqueda directa falló. Probando filtro.",

                errorDirecto

            );


            /*
             * Segundo intento usando filtro.
             */

            const filtroCorreo =
                correo.replace(
                    /'/g,
                    "''"
                );


            const urlFiltro =
                `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${filtroCorreo}' or mail eq '${filtroCorreo}'&$select=id,displayName,mail,userPrincipalName`;


            const resultado =
                await graphFetch(
                    urlFiltro
                );


            if (
                !resultado.value ||
                resultado.value.length === 0
            ) {

                throw errorDirecto;
            }


            usuario =
                resultado.value[0];
        }


        if (!usuario) {

            throw new Error(
                "El usuario no fue encontrado en Microsoft Entra ID."
            );
        }


        usuarioSeleccionado = {

            id:
                usuario.id ||
                "",

            nombre:
                usuario.displayName ||
                "Sin nombre",

            correo:
                normalizarCorreo(

                    usuario.mail ||

                    usuario.userPrincipalName ||

                    correo

                ),

            userPrincipalName:
                usuario.userPrincipalName ||
                ""

        };


        console.log(
            "Usuario encontrado en Graph:",
            usuarioSeleccionado
        );


        /*
         * Mostrar usuario.
         */

        const nombreUsuario =
            document.getElementById(
                "nombreUsuario"
            );


        const correoMostrado =
            document.getElementById(
                "correoMostrado"
            );


        if (nombreUsuario) {

            nombreUsuario.textContent =
                usuarioSeleccionado.nombre;
        }


        if (correoMostrado) {

            correoMostrado.textContent =
                usuarioSeleccionado.correo;
        }


        /*
         * Buscar si ya tiene permisos
         * en Microsoft Lists.
         */

        const registro =
            buscarRegistroPorCorreo(

                usuarioSeleccionado.correo

            );


        if (registro) {

            cargarRegistroEnFormulario(
                registro
            );

        } else {

            prepararNuevoUsuario();
        }


        mostrarPanelUsuario();


    } catch (error) {

        console.error(
            "Error buscando usuario en Graph:",
            error
        );


        mostrarMensaje(

            obtenerMensajeBusqueda(
                error
            ),

            "error"

        );


        ocultarPanelUsuario();


    } finally {

        if (btn) {

            btn.disabled =
                false;

            btn.textContent =
                "Buscar";
        }
    }
}


/* =========================================================
   CARGAR REGISTRO EXISTENTE
========================================================= */

function cargarRegistroEnFormulario(
    registro
) {

    const grupo =
        document.getElementById(
            "grupoUsuario"
        );


    const grupoRegistrado =
        String(

            obtenerCampo(
                registro,
                "Grupo"
            ) ||

            GRUPO_NORMAL

        ).trim();


    if (grupo) {

        grupo.value =

            grupoRegistrado.toLowerCase() ===
                GRUPO_TI.toLowerCase()

                ? GRUPO_TI

                : GRUPO_NORMAL;
    }


    MODULOS_PERMISOS.forEach(
        modulo => {

            const checkbox =
                document.querySelector(

                    `.permiso-modulo[data-modulo="${modulo}"]`

                );


            if (!checkbox) {

                return;
            }


            checkbox.checked =
                valorBooleano(

                    obtenerCampo(
                        registro,
                        modulo
                    )

                );
        }
    );


    cambiarGrupo();
}


/* =========================================================
   PREPARAR NUEVO USUARIO
========================================================= */

function prepararNuevoUsuario() {

    const grupo =
        document.getElementById(
            "grupoUsuario"
        );


    if (grupo) {

        grupo.value =
            GRUPO_NORMAL;
    }


    MODULOS_PERMISOS.forEach(
        modulo => {

            const checkbox =
                document.querySelector(

                    `.permiso-modulo[data-modulo="${modulo}"]`

                );


            if (!checkbox) {

                return;
            }


            checkbox.checked =
                false;

            checkbox.disabled =
                false;
        }
    );
}


/* =========================================================
   CAMBIAR GRUPO
========================================================= */

function cambiarGrupo() {

    const grupo =
        document.getElementById(
            "grupoUsuario"
        );


    if (!grupo) {

        return;
    }


    const esTI =
        grupo.value ===
        GRUPO_TI;


    const controles =
        document.querySelectorAll(
            ".permiso-modulo"
        );


    controles.forEach(
        control => {

            if (esTI) {

                control.checked =
                    true;

                control.disabled =
                    true;


            } else {

                control.disabled =
                    false;

            }
        }
    );
}


/* =========================================================
   EDITAR REGISTRO
========================================================= */

async function editarRegistro(
    itemId
) {

    const registro =
        registrosPermisos.find(

            item =>
                String(
                    item.id
                ) ===
                String(
                    itemId
                )

        );


    if (!registro) {

        alert(
            "No se encontró el registro."
        );

        return;
    }


    const correo =
        obtenerCampo(
            registro,
            "UsuarioCorreo"
        );


    if (!correo) {

        alert(
            "El registro no tiene correo."
        );

        return;
    }


    const input =
        document.getElementById(
            "correoUsuario"
        );


    if (input) {

        input.value =
            correo;
    }


    await buscarUsuario();
}


/* =========================================================
   GUARDAR PERMISOS
========================================================= */

async function guardarPermisos() {

    const btn =
        document.getElementById(
            "btnGuardarPermisos"
        );


    const grupoSelect =
        document.getElementById(
            "grupoUsuario"
        );


    if (!usuarioSeleccionado) {

        alert(
            "Primero busca un usuario válido en Graph."
        );

        return;
    }


    const grupo =
        grupoSelect
            ? grupoSelect.value
            : GRUPO_NORMAL;


    const correo =
        usuarioSeleccionado.correo;


    const nombre =
        usuarioSeleccionado.nombre;


    const permisos = {};


    MODULOS_PERMISOS.forEach(
        modulo => {

            const checkbox =
                document.querySelector(

                    `.permiso-modulo[data-modulo="${modulo}"]`

                );


            permisos[modulo] =
                checkbox
                    ? checkbox.checked
                    : false;
        }
    );


    /*
     * TI siempre tiene todo habilitado.
     */

    if (
        grupo ===
        GRUPO_TI
    ) {

        MODULOS_PERMISOS.forEach(
            modulo => {

                permisos[modulo] =
                    true;
            }
        );
    }


    try {

        if (btn) {

            btn.disabled =
                true;

            btn.textContent =
                "Guardando...";
        }


        const sitio =
            await obtenerSitioSharePoint();


        const lista =
            await obtenerListaPermisos();


        const campos =
            await obtenerMapaCampos();


        const registroExistente =
            buscarRegistroPorCorreo(
                correo
            );


        const fields = {};


        fields[
            campos.Title
        ] =
            nombre ||
            correo;


        fields[
            campos.UsuarioCorreo
        ] =
            correo;


        fields[
            campos.NombreUsuario
        ] =
            nombre ||
            correo;


        fields[
            campos.Grupo
        ] =
            grupo;


        MODULOS_PERMISOS.forEach(
            modulo => {

                fields[
                    campos[modulo]
                ] =
                    Boolean(
                        permisos[
                            modulo
                        ]
                    );
            }
        );


        /*
         * ACTUALIZAR
         */

        if (registroExistente) {

            const url =

                `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items/${registroExistente.id}/fields`;


            await graphFetch(

                url,

                {

                    method:
                        "PATCH",

                    body:
                        JSON.stringify(
                            fields
                        )

                }

            );


            mostrarMensaje(

                "Permisos actualizados correctamente.",

                "success"

            );


        /*
         * CREAR
         */

        } else {

            const url =

                `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items`;


            await graphFetch(

                url,

                {

                    method:
                        "POST",

                    body:
                        JSON.stringify({
                            fields
                        })

                }

            );


            mostrarMensaje(

                "Usuario y permisos guardados correctamente.",

                "success"

            );
        }


        /*
         * Recargar tabla.
         */

        await cargarPermisos();


        /*
         * Volver a cargar el registro
         * actualizado en los switches.
         */

        const actualizado =
            buscarRegistroPorCorreo(
                correo
            );


        if (actualizado) {

            cargarRegistroEnFormulario(
                actualizado
            );
        }


    } catch (error) {

        console.error(
            "Error guardando permisos:",
            error
        );


        mostrarMensaje(

            `No se pudieron guardar los permisos. ${obtenerMensajeError(error)}`,

            "error"

        );


    } finally {

        if (btn) {

            btn.disabled =
                false;

            btn.textContent =
                "Guardar permisos";
        }
    }
}


/* =========================================================
   MOSTRAR PANEL
========================================================= */

function mostrarPanelUsuario() {

    const panel =
        document.getElementById(
            "panelPermisos"
        );


    if (panel) {

        panel.style.display =
            "block";
    }
}


/* =========================================================
   OCULTAR PANEL
========================================================= */

function ocultarPanelUsuario() {

    const panel =
        document.getElementById(
            "panelPermisos"
        );


    if (panel) {

        panel.style.display =
            "none";
    }
}


/* =========================================================
   ACCESO DENEGADO
========================================================= */

function mostrarAccesoDenegado() {

    const contenido =
        document.querySelector(
            ".content"
        );


    if (!contenido) {

        return;
    }


    contenido.innerHTML =
        `
        <section class="page-header">

            <div>

                <h1>
                    Permisos
                </h1>

                <p>
                    Administración de accesos al sistema
                </p>

            </div>

        </section>


        <div class="permisos-card">

            <div class="tabla-vacia">

                <span>
                    No tienes permisos para administrar esta sección.
                </span>

            </div>

        </div>
        `;
}


/* =========================================================
   MENSAJES
========================================================= */

function mostrarMensaje(
    mensaje,
    tipo = "error"
) {

    const contenido =
        document.querySelector(
            ".content"
        );


    if (!contenido) {

        return;
    }


    const anterior =
        contenido.querySelector(
            ".mensaje-permisos"
        );


    if (anterior) {

        anterior.remove();
    }


    const aviso =
        document.createElement(
            "div"
        );


    aviso.className =
        `permission-message ${tipo} mensaje-permisos`;


    aviso.textContent =
        mensaje;


    contenido.prepend(
        aviso
    );


    window.setTimeout(
        () => {

            if (
                aviso &&
                aviso.parentNode
            ) {

                aviso.remove();
            }

        },
        5000
    );
}


/* =========================================================
   MENSAJE DE BÚSQUEDA
========================================================= */

function obtenerMensajeBusqueda(
    error
) {

    const mensaje =
        String(
            error?.message ||
            ""
        );


    if (
        mensaje.includes(
            "Graph 403"
        )
    ) {

        return (

            "Microsoft Graph rechazó la consulta. Verifica que User.ReadBasic.All esté habilitado para la aplicación."

        );
    }


    if (
        mensaje.includes(
            "Graph 401"
        )
    ) {

        return (

            "La sesión de Microsoft expiró. Cierra sesión y vuelve a ingresar."

        );
    }


    if (
        mensaje.includes(
            "Graph 404"
        )
    ) {

        return (

            "El usuario no existe o no fue encontrado en Microsoft Entra ID."

        );
    }


    if (
        mensaje.includes(
            "AADSTS65001"
        ) ||

        mensaje.includes(
            "consent"
        )
    ) {

        return (

            "Microsoft solicita consentimiento para los permisos de Graph. Vuelve a iniciar sesión y acepta los permisos."

        );
    }


    return (

        `No se pudo buscar el usuario. ${mensaje}`

    );
}


/* =========================================================
   NORMALIZAR TEXTO
========================================================= */

function normalizarTexto(
    valor
) {

    return String(
        valor || ""
    )
        .normalize(
            "NFD"
        )
        .replace(
            /[\u0300-\u036f]/g,
            ""
        )
        .toLowerCase()
        .replace(
            /[^a-z0-9]/g,
            ""
        );
}


/* =========================================================
   NORMALIZAR CORREO
========================================================= */

function normalizarCorreo(
    correo
) {

    return String(
        correo || ""
    )
        .trim()
        .toLowerCase();
}


/* =========================================================
   VALOR BOOLEANO
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


    const texto =
        String(
            valor || ""
        )
            .trim()
            .toLowerCase();


    return (

        texto === "true" ||

        texto === "1" ||

        texto === "yes" ||

        texto === "sí" ||

        texto === "si"

    );
}


/* =========================================================
   ESCAPAR HTML
========================================================= */

function escapeHtml(
    valor
) {

    return String(
        valor ?? ""
    )
        .replace(
            /&/g,
            "&amp;"
        )
        .replace(
            /</g,
            "&lt;"
        )
        .replace(
            />/g,
            "&gt;"
        )
        .replace(
            /"/g,
            "&quot;"
        )
        .replace(
            /'/g,
            "&#039;"
        );
}


/* =========================================================
   ERROR GENERAL
========================================================= */

function obtenerMensajeError(
    error
) {

    return String(

        error?.message ||

        "Error desconocido."

    );
}


/* =========================================================
   MENÚ
========================================================= */

function configurarMenu() {

    const menuToggle =
        document.getElementById(
            "menuToggle"
        );


    const sidebar =
        document.getElementById(
            "sidebar"
        );


    if (
        menuToggle &&
        sidebar
    ) {

        menuToggle.addEventListener(
            "click",
            () => {

                sidebar.classList.toggle(
                    "open"
                );

            }
        );
    }
}


/* =========================================================
   EVENTOS
========================================================= */

function configurarEventos() {

    const btnBuscar =
        document.getElementById(
            "btnBuscarUsuario"
        );


    if (btnBuscar) {

        btnBuscar.addEventListener(
            "click",
            buscarUsuario
        );
    }


    const grupo =
        document.getElementById(
            "grupoUsuario"
        );


    if (grupo) {

        grupo.addEventListener(
            "change",
            cambiarGrupo
        );
    }


    const btnGuardar =
        document.getElementById(
            "btnGuardarPermisos"
        );


    if (btnGuardar) {

        btnGuardar.addEventListener(
            "click",
            guardarPermisos
        );
    }


    const correo =
        document.getElementById(
            "correoUsuario"
        );


    if (correo) {

        correo.addEventListener(
            "keydown",
            event => {

                if (
                    event.key ===
                    "Enter"
                ) {

                    event.preventDefault();

                    buscarUsuario();
                }
            }
        );
    }


    const tabla =
        document.getElementById(
            "tablaPermisos"
        );


    if (tabla) {

        tabla.addEventListener(
            "click",
            event => {

                const boton =
                    event.target.closest(
                        "[data-editar-permiso]"
                    );


                if (!boton) {

                    return;
                }


                const itemId =
                    boton.getAttribute(
                        "data-editar-permiso"
                    );


                editarRegistro(
                    itemId
                );
            }
        );
    }
}