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

let usuarioSeleccionado = null;

let registrosPermisos = [];

let sitioSharePointCache = null;

let listaPermisosCache = null;

let columnasPermisosCache = null;


/* =========================================================
   INICIO
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    async () => {

        console.log("Permisos cargado");

        const contenido =
            document.querySelector(".content");

        if (contenido) {
            contenido.style.display = "block";
            contenido.style.visibility = "visible";
            contenido.style.opacity = "1";
        }

        ocultarPanelUsuario();

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
                "No se pudo obtener el usuario actual.",
                "error"
            );

            return;
        }


        const correoActual =
            obtenerCorreoCuenta(cuenta);


        console.log(
            "Usuario actual:",
            correoActual
        );


        const userInfo =
            document.getElementById("userInfo");

        if (userInfo) {

            userInfo.textContent =
                cuenta.nombre ||
                correoActual ||
                "";
        }


        /* =====================================================
           CARGAR LISTA DE PERMISOS
        ====================================================== */

        await cargarPermisos();


        /* =====================================================
           COMPROBAR SI ES TI
        ====================================================== */

        const registroActual =
            buscarRegistroPorCorreo(
                correoActual
            );


        const esTI =
            USUARIOS_TI_INICIALES
                .map(normalizarCorreo)
                .includes(
                    normalizarCorreo(correoActual)
                )
            ||
            (
                registroActual &&
                String(
                    obtenerCampo(
                        registroActual,
                        "Grupo"
                    )
                ).trim().toLowerCase() ===
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
            "Error inicializando permisos:",
            error
        );

        mostrarMensaje(
            obtenerMensajeError(error),
            "error"
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
                    event.key === "Enter"
                ) {

                    event.preventDefault();

                    buscarUsuario();
                }
            }
        );
    }


    /* =====================================================
       EDITAR DESDE TABLA
    ====================================================== */

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

                if (!boton) return;

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
   OCULTAR PANEL
========================================================= */

function ocultarPanelUsuario() {

    const panel =
        document.getElementById(
            "panelPermisos"
        );

    if (panel) {

        panel.style.display = "none";
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

        panel.style.display = "block";
    }
}


/* =========================================================
   CUENTA ACTUAL
========================================================= */

async function obtenerCuentaActual() {

    try {

        let instancia = null;


        if (
            typeof msalInstance !==
            "undefined"
        ) {

            instancia =
                msalInstance;
        }


        if (
            !instancia &&
            window.msalInstance
        ) {

            instancia =
                window.msalInstance;
        }


        if (!instancia) {

            console.error(
                "MSAL no está disponible."
            );

            return null;
        }


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

                try {

                    instancia.setActiveAccount(
                        cuenta
                    );

                } catch (e) {
                }
            }
        }


        if (!cuenta) {

            return null;
        }


        return cuenta;

    } catch (error) {

        console.error(
            "Error obteniendo cuenta:",
            error
        );

        return null;
    }
}


/* =========================================================
   CORREO DE CUENTA
========================================================= */

function obtenerCorreoCuenta(cuenta) {

    if (!cuenta) {
        return "";
    }

    return (
        cuenta.username ||
        cuenta.mail ||
        cuenta.userPrincipalName ||
        ""
    )
        .trim()
        .toLowerCase();
}


/* =========================================================
   TOKEN GRAPH
========================================================= */

async function obtenerTokenGraph() {

    let instancia = null;


    if (
        typeof msalInstance !==
        "undefined"
    ) {

        instancia =
            msalInstance;
    }


    if (
        !instancia &&
        window.msalInstance
    ) {

        instancia =
            window.msalInstance;
    }


    if (!instancia) {

        throw new Error(
            "MSAL no está disponible."
        );
    }


    const cuenta =
        await obtenerCuentaActual();


    if (!cuenta) {

        throw new Error(
            "No existe una sesión activa de Microsoft."
        );
    }


    try {

        const resultado =
            await instancia.acquireTokenSilent(
                {
                    scopes:
                        SCOPES_GRAPH,
                    account:
                        cuenta
                }
            );


        return resultado.accessToken;

    } catch (error) {

        console.warn(
            "No se pudo obtener el token silenciosamente. Solicitando consentimiento.",
            error
        );


        try {

            const resultado =
                await instancia.acquireTokenPopup(
                    {
                        scopes:
                            SCOPES_GRAPH,
                        account:
                            cuenta
                    }
                );


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


    if (
        !response.ok
    ) {

        let detalle = "";


        try {

            const errorData =
                await response.json();

            detalle =
                errorData?.error?.message ||
                "";
        } catch (e) {
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
   NORMALIZAR TEXTO
========================================================= */

function normalizarTexto(valor) {

    return String(
        valor || ""
    )
        .normalize("NFD")
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

function normalizarCorreo(correo) {

    return String(
        correo || ""
    )
        .trim()
        .toLowerCase();
}


/* =========================================================
   OBTENER NOMBRE INTERNO
========================================================= */

async function obtenerNombreCampo(
    nombresPosibles,
    obligatorio = true
) {

    const columnas =
        await obtenerColumnasPermisos();


    const candidatos =
        nombresPosibles.map(
            normalizarTexto
        );


    let columna =
        columnas.find(
            campo =>
                candidatos.includes(
                    normalizarTexto(
                        campo.displayName
                    )
                )
        );


    if (!columna) {

        columna =
            columnas.find(
                campo =>
                    candidatos.includes(
                        normalizarTexto(
                            campo.name
                        )
                    )
            );
    }


    if (!columna) {

        if (obligatorio) {

            throw new Error(
                `No se encontró la columna: ${nombresPosibles.join(", ")}`
            );
        }


        return null;
    }


    return columna.name;
}


/* =========================================================
   MAPA DE COLUMNAS
========================================================= */

async function obtenerMapaCampos() {

    const mapa = {};


    mapa.Title =
        await obtenerNombreCampo(
            [
                "Title",
                "Título"
            ]
        );


    mapa.UsuarioCorreo =
        await obtenerNombreCampo(
            [
                "UsuarioCorreo"
            ]
        );


    mapa.NombreUsuario =
        await obtenerNombreCampo(
            [
                "NombreUsuario"
            ]
        );


    mapa.Grupo =
        await obtenerNombreCampo(
            [
                "Grupo"
            ]
        );


    for (
        const modulo of
        MODULOS_PERMISOS
    ) {

        mapa[modulo] =
            await obtenerNombreCampo(
                [
                    modulo
                ]
            );
    }


    return mapa;
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


    try {

        tabla.innerHTML =
            `
            <div class="tabla-vacia">
                <span>
                    Cargando usuarios...
                </span>
            </div>
            `;


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

async function renderizarTablaPermisos(
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

            const fields =
                registro.fields ||
                {};


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
                obtenerCampo(
                    registro,
                    "Grupo"
                ) ||
                GRUPO_NORMAL;


            const cantidad =
                contarPermisos(
                    registro
                );


            const esCompleto =
                grupo === GRUPO_TI ||
                cantidad ===
                    MODULOS_PERMISOS.length;


            let accesoTexto = "";


            if (
                grupo === GRUPO_TI
            ) {

                accesoTexto =
                    "Acceso total";

            } else {

                accesoTexto =
                    `${cantidad}/${MODULOS_PERMISOS.length}`;
            }


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

                        ${
                            grupo === GRUPO_TI

                                ? `
                                    <span class="badge badge-ti">
                                        TI
                                    </span>
                                  `

                                : `
                                    <span class="badge badge-normal">
                                        Normal
                                    </span>
                                  `
                        }

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
            )
        )
            .trim();


    if (
        grupo ===
        GRUPO_TI
    ) {

        return MODULOS_PERMISOS.length;
    }


    let cantidad = 0;


    MODULOS_PERMISOS.forEach(
        modulo => {

            const valor =
                obtenerCampo(
                    registro,
                    modulo
                );


            if (
                valor === true ||
                valor === 1 ||
                String(
                    valor
                ).toLowerCase() ===
                    "true" ||
                String(
                    valor
                ).toLowerCase() ===
                    "sí" ||
                String(
                    valor
                ).toLowerCase() ===
                    "si"
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

            btn.disabled = true;
            btn.textContent =
                "Buscando...";
        }


        ocultarPanelUsuario();

        usuarioSeleccionado = null;


        /* =====================================================
           GRAPH /USERS
        ====================================================== */

        const url =
            `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(correo)}?$select=id,displayName,mail,userPrincipalName`;


        let usuario;


        try {

            usuario =
                await graphFetch(
                    url
                );

        } catch (error) {

            console.warn(
                "Búsqueda directa falló. Intentando por filtro...",
                error
            );


            const urlFiltro =
                `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${correo}' or mail eq '${correo}'&$select=id,displayName,mail,userPrincipalName`;


            const resultado =
                await graphFetch(
                    urlFiltro
                );


            if (
                !resultado.value ||
                resultado.value.length === 0
            ) {

                throw error;
            }


            usuario =
                resultado.value[0];
        }


        if (!usuario) {

            throw new Error(
                "Usuario no encontrado."
            );
        }


        const correoUsuario =
            normalizarCorreo(
                usuario.mail ||
                usuario.userPrincipalName ||
                correo
            );


        usuarioSeleccionado = {
            id:
                usuario.id,

            nombre:
                usuario.displayName ||
                "Sin nombre",

            correo:
                correoUsuario,

            userPrincipalName:
                usuario.userPrincipalName ||
                ""
        };


        console.log(
            "Usuario encontrado en Graph:",
            usuarioSeleccionado
        );


        /* =====================================================
           MOSTRAR INFORMACIÓN
        ====================================================== */

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


        /* =====================================================
           BUSCAR PERMISOS YA EXISTENTES
        ====================================================== */

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
            "Error buscando usuario:",
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

            btn.disabled = false;

            btn.textContent =
                "Buscar";
        }
    }
}


/* =========================================================
   CARGAR REGISTRO EN FORMULARIO
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
            grupoRegistrado === GRUPO_TI
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


            const valor =
                obtenerCampo(
                    registro,
                    modulo
                );


            checkbox.checked =
                valor === true ||
                valor === 1 ||
                String(
                    valor
                ).toLowerCase() ===
                    "true" ||
                String(
                    valor
                ).toLowerCase() ===
                    "sí" ||
                String(
                    valor
                ).toLowerCase() ===
                    "si";
        }
    );


    cambiarGrupo();


    console.log(
        "Permisos existentes cargados:",
        registro
    );
}


/* =========================================================
   NUEVO USUARIO
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


            checkbox.checked = false;
            checkbox.disabled = false;
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
        grupo.value === GRUPO_TI;


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
                String(itemId)
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


    try {

        await buscarUsuario();

    } catch (error) {

        console.error(
            error
        );
    }
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


    if (
        grupo === GRUPO_TI
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

            btn.disabled = true;

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
                        permisos[modulo]
                    );
            }
        );


        /* =====================================================
           ACTUALIZAR
        ====================================================== */

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


        /* =====================================================
           CREAR
        ====================================================== */

        } else {

            const url =
                `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items`;


            await graphFetch(
                url,
                {
                    method:
                        "POST",

                    body:
                        JSON.stringify(
                            {
                                fields
                            }
                        )
                }
            );


            mostrarMensaje(
                "Usuario y permisos guardados correctamente.",
                "success"
            );
        }


        /* =====================================================
           RECARGAR LISTA
        ====================================================== */

        await cargarPermisos();


        /* =====================================================
           VOLVER A CARGAR REGISTRO
        ====================================================== */

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

            btn.disabled = false;

            btn.textContent =
                "Guardar permisos";
        }
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
   MENSAJE
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
            "Graph 404"
        )
    ) {

        return (
            "El correo no existe en Microsoft Entra ID."
        );
    }


    if (
        mensaje.includes(
            "Graph 403"
        )
    ) {

        return (
            "Microsoft Graph rechazó la consulta. Verifica que la aplicación tenga User.ReadBasic.All."
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


    return (
        `No se pudo buscar el usuario. ${mensaje}`
    );
}


/* =========================================================
   MENSAJE GENERAL DE ERROR
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