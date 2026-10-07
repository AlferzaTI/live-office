/* =========================================================
   ALFERZA LIVE OFFICE
   GESTIÓN DE PERMISOS
========================================================= */


/* =========================================================
   CONFIGURACIÓN MSAL
========================================================= */

const PERMISOS_TENANT_ID =
    "dbab984f-4bb1-4b60-9dff-da59f54acdf1";


const PERMISOS_CLIENT_ID =
    "5d98417c-74a7-4fab-8f2c-41ac127be696";


const permisosMsalConfig = {

    auth: {

        clientId:
            PERMISOS_CLIENT_ID,

        authority:
            `https://login.microsoftonline.com/${PERMISOS_TENANT_ID}`,

        redirectUri:
            "https://AlferzaTI.github.io/live-office/blank.html"

    },

    cache: {

        cacheLocation:
            "sessionStorage",

        storeAuthStateInCookie:
            false

    }

};


/* =========================================================
   SCOPES
   MISMO ESQUEMA DE APP.JS
========================================================= */

const permisosScopes = [

    "User.Read",

    "Presence.Read.All",

    "Sites.Read.All",

    "AuditLog.Read.All",

    "DeviceManagementManagedDevices.Read.All"

];


/* =========================================================
   INSTANCIA MSAL
========================================================= */

const permisosMsalInstance =
    new msal.PublicClientApplication(
        permisosMsalConfig
    );


/* =========================================================
   CONFIGURACIÓN SHAREPOINT
========================================================= */

const PERMISOS_LIST_NAME =
    "PermisosTI";


const SHAREPOINT_HOST =
    "alferzaholding-my.sharepoint.com";


const SHAREPOINT_SITE_PATH =
    "/personal/soporte1_alferza_pe";


/* =========================================================
   GRUPOS
========================================================= */

const GRUPO_TI =
    "TI";


const GRUPO_NORMAL =
    "Normal";


/* =========================================================
   USUARIOS TI INICIALES
========================================================= */

const USUARIOS_TI_INICIALES = [

    "soporte1@alferza.pe"

];


/* =========================================================
   MÓDULOS
========================================================= */

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
   VARIABLES
========================================================= */

let usuarioSeleccionado =
    null;


let registrosPermisos =
    [];


let sitioSharePoint =
    null;


let listaPermisos =
    null;


let columnasPermisos =
    null;


let mapaCampos =
    null;


/* =========================================================
   INICIO
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    async function () {

        console.log(
            "Permisos cargado."
        );


        configurarMenu();


        configurarEventos();


        try {

            await inicializarPermisos();

        }

        catch (error) {

            console.error(
                "Error inicializando Permisos:",
                error
            );


            mostrarMensaje(
                obtenerMensajeError(error),
                "error"
            );

        }

    }
);


/* =========================================================
   INICIALIZAR PERMISOS
========================================================= */

async function inicializarPermisos() {

    /*
     * Recuperamos la cuenta MSAL que ya inició sesión
     * en login.js.
     */

    const cuenta =
        obtenerCuentaMSAL();


    if (!cuenta) {

        mostrarMensaje(

            "No existe una sesión de Microsoft disponible. Cierra sesión y vuelve a ingresar.",

            "error"

        );

        return;

    }


    const correoActual =
        obtenerCorreoCuenta(
            cuenta
        );


    console.log(
        "Cuenta actual:",
        correoActual
    );


    /*
     * Mostrar usuario en navbar.
     */

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
     * Cargar permisos registrados.
     */

    await cargarPermisos();


    /*
     * Verificar si el usuario actual es TI.
     */

    const registroActual =
        buscarRegistroPorCorreo(
            correoActual
        );


    const esTI =
        USUARIOS_TI_INICIALES
            .map(
                normalizarCorreo
            )
            .includes(
                normalizarCorreo(
                    correoActual
                )
            )
        ||

        (

            registroActual &&

            normalizarTexto(

                obtenerCampo(
                    registroActual,
                    "Grupo"
                )

            ) ===
                normalizarTexto(
                    GRUPO_TI
                )

        );


    if (!esTI) {

        mostrarAccesoDenegado();

        return;

    }


    console.log(
        "Usuario TI autorizado."
    );

}


/* =========================================================
   OBTENER CUENTA MSAL
========================================================= */

function obtenerCuentaMSAL() {

    const cuentas =
        permisosMsalInstance
            .getAllAccounts();


    if (
        cuentas &&
        cuentas.length > 0
    ) {

        const cuenta =
            cuentas[0];


        try {

            permisosMsalInstance
                .setActiveAccount(
                    cuenta
                );

        }

        catch (error) {

            console.warn(
                "No se pudo establecer cuenta activa:",
                error
            );

        }


        return cuenta;

    }


    return null;

}


/* =========================================================
   OBTENER CORREO
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
   OBTENER TOKEN GRAPH
========================================================= */

async function obtenerTokenGraph() {

    const cuenta =
        obtenerCuentaMSAL();


    if (!cuenta) {

        throw new Error(

            "No existe una sesión de Microsoft."

        );

    }


    permisosMsalInstance
        .setActiveAccount(
            cuenta
        );


    try {

        /*
         * MISMO mecanismo que app.js.
         */

        const respuesta =
            await permisosMsalInstance
                .acquireTokenSilent({

                    scopes:
                        permisosScopes,

                    account:
                        cuenta

                });


        return respuesta.accessToken;

    }

    catch (error) {

        console.error(

            "No se pudo obtener el token silenciosamente:",

            error

        );


        /*
         * NO hacemos loginPopup.
         *
         * La sesión se genera desde login.js.
         */

        throw new Error(

            "No se pudo obtener el token de Microsoft Graph. Cierra sesión y vuelve a ingresar."

        );

    }

}


/* =========================================================
   GRAPH FETCH
========================================================= */

async function graphFetch(
    url,
    opciones = {}
) {

    const TOKEN =
        await obtenerTokenGraph();


    const headers = {

        Authorization:
            `Bearer ${TOKEN}`,

        Accept:
            "application/json"

    };


    if (
        opciones.body
    ) {

        headers[
            "Content-Type"
        ] =
            "application/json";

    }


    const respuesta =
        await fetch(

            url,

            {

                method:
                    opciones.method ||
                    "GET",

                headers:

                    {

                        ...headers,

                        ...(opciones.headers ||
                            {})

                    },

                body:
                    opciones.body ||
                    undefined

            }

        );


    if (!respuesta.ok) {

        let detalle =
            "";


        try {

            const errorData =
                await respuesta.json();


            detalle =
                errorData
                    ?.error
                    ?.message ||
                "";

        }

        catch (error) {

            // Sin respuesta JSON.

        }


        throw new Error(

            `Graph ${respuesta.status}: ${
                detalle ||
                respuesta.statusText
            }`

        );

    }


    if (
        respuesta.status ===
        204
    ) {

        return null;

    }


    return await respuesta.json();

}


/* =========================================================
   OBTENER SITIO SHAREPOINT
========================================================= */

async function obtenerSitioSharePoint() {

    if (
        sitioSharePoint
    ) {

        return sitioSharePoint;

    }


    const url =

        `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOST}:${SHAREPOINT_SITE_PATH}`;


    sitioSharePoint =
        await graphFetch(
            url
        );


    return sitioSharePoint;

}


/* =========================================================
   OBTENER LISTA PERMISOS
========================================================= */

async function obtenerListaPermisos() {

    if (
        listaPermisos
    ) {

        return listaPermisos;

    }


    const sitio =
        await obtenerSitioSharePoint();


    const url =

        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${encodeURIComponent(PERMISOS_LIST_NAME)}`;


    listaPermisos =
        await graphFetch(
            url
        );


    return listaPermisos;

}


/* =========================================================
   OBTENER COLUMNAS
========================================================= */

async function obtenerColumnasPermisos() {

    if (
        columnasPermisos
    ) {

        return columnasPermisos;

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


    columnasPermisos =
        resultado.value ||
        [];


    return columnasPermisos;

}


/* =========================================================
   OBTENER MAPA DE COLUMNAS
========================================================= */

async function obtenerMapaCampos() {

    if (
        mapaCampos
    ) {

        return mapaCampos;

    }


    const columnas =
        await obtenerColumnasPermisos();


    mapaCampos = {};


    /*
     * TITLE
     */

    const columnaTitle =
        encontrarColumna(

            columnas,

            [
                "Title",
                "Título"
            ]

        );


    if (!columnaTitle) {

        throw new Error(

            "No se encontró la columna Title/Título en PermisosTI."

        );

    }


    mapaCampos.Title =
        columnaTitle.name;


    /*
     * USUARIOCORREO
     */

    const columnaCorreo =
        encontrarColumna(

            columnas,

            [
                "UsuarioCorreo"
            ]

        );


    if (!columnaCorreo) {

        throw new Error(

            "No se encontró la columna UsuarioCorreo."

        );

    }


    mapaCampos.UsuarioCorreo =
        columnaCorreo.name;


    /*
     * NOMBREUSUARIO
     */

    const columnaNombre =
        encontrarColumna(

            columnas,

            [
                "NombreUsuario"
            ]

        );


    if (!columnaNombre) {

        throw new Error(

            "No se encontró la columna NombreUsuario."

        );

    }


    mapaCampos.NombreUsuario =
        columnaNombre.name;


    /*
     * GRUPO
     */

    const columnaGrupo =
        encontrarColumna(

            columnas,

            [
                "Grupo"
            ]

        );


    if (!columnaGrupo) {

        throw new Error(

            "No se encontró la columna Grupo."

        );

    }


    mapaCampos.Grupo =
        columnaGrupo.name;


    /*
     * MÓDULOS
     */

    MODULOS_PERMISOS
        .forEach(
            modulo => {

                const columna =
                    encontrarColumna(

                        columnas,

                        [
                            modulo
                        ]

                    );


                if (!columna) {

                    throw new Error(

                        `No se encontró la columna ${modulo}.`

                    );

                }


                mapaCampos[
                    modulo
                ] =
                    columna.name;

            }
        );


    console.log(
        "Mapa de columnas:",
        mapaCampos
    );


    return mapaCampos;

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


    /*
     * Primero displayName.
     */

    let columna =
        columnas.find(

            item =>

                objetivos.includes(

                    normalizarTexto(
                        item.displayName
                    )

                )

        );


    /*
     * Después nombre interno.
     */

    if (!columna) {

        columna =
            columnas.find(

                item =>

                    objetivos.includes(

                        normalizarTexto(
                            item.name
                        )

                    )

            );

    }


    return columna ||
        null;

}


/* =========================================================
   OBTENER ITEMS DE PERMISOS
========================================================= */

async function obtenerTodosLosItems() {

    const sitio =
        await obtenerSitioSharePoint();


    const lista =
        await obtenerListaPermisos();


    let url =

        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items?expand=fields&$top=500`;


    const items =
        [];


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
            resultado[
                "@odata.nextLink"
            ] ||
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

    }

    catch (error) {

        console.error(

            "Error cargando PermisosTI:",

            error

        );


        tabla.innerHTML =

            `
            <div class="tabla-vacia">

                <span>
                    No se pudieron cargar los usuarios.
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


    let filas =
        "";


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

                ).trim();


            const cantidad =
                contarPermisos(
                    registro
                );


            const esTI =
                normalizarTexto(
                    grupo
                ) ===
                normalizarTexto(
                    GRUPO_TI
                );


            const accesoCompleto =
                esTI ||
                cantidad ===
                    MODULOS_PERMISOS.length;


            const accesoTexto =
                esTI

                    ? "Acceso total"

                    : `${cantidad}/${MODULOS_PERMISOS.length}`;


            const claseAcceso =
                accesoCompleto

                    ? "full"

                    : "partial";


            const claseGrupo =
                esTI

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
                                    ${escapeHtml(
                                        nombre
                                    )}
                                </strong>

                                <span>
                                    ${escapeHtml(
                                        correo
                                    )}
                                </span>

                            </div>

                        </div>

                    </td>


                    <td>

                        <span
                            class="badge ${claseGrupo}"
                        >
                            ${escapeHtml(
                                grupo
                            )}
                        </span>

                    </td>


                    <td>

                        <span
                            class="
                                access-summary
                                ${claseAcceso}
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
   BUSCAR USUARIO CON GRAPH
========================================================= */

async function buscarUsuario() {

    const input =
        document.getElementById(
            "correoUsuario"
        );


    const boton =
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

        if (boton) {

            boton.disabled =
                true;

            boton.textContent =
                "Buscando...";

        }


        ocultarPanelUsuario();


        usuarioSeleccionado =
            null;


        /*
         * MISMO GRAPH QUE APP.JS.
         *
         * En app.js ya utilizas:
         *
         * GET /users?$top=999
         *
         * Aquí hacemos una búsqueda filtrada
         * para no traer todos los usuarios.
         */

        const correoFiltro =
            correo.replace(
                /'/g,
                "''"
            );


        const url =

            `https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq '${correoFiltro}' or mail eq '${correoFiltro}'&$select=id,displayName,mail,userPrincipalName`;


        console.log(
            "Buscando usuario:",
            correo
        );


        const resultado =
            await graphFetch(
                url
            );


        console.log(
            "Respuesta Graph:",
            resultado
        );


        if (
            !resultado.value ||
            resultado.value.length === 0
        ) {

            /*
             * Segundo intento con búsqueda
             * directa por UPN.
             */

            const urlDirecta =

                `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(correo)}?$select=id,displayName,mail,userPrincipalName`;


            const usuarioDirecto =
                await graphFetch(
                    urlDirecta
                );


            if (!usuarioDirecto) {

                throw new Error(

                    "El usuario no fue encontrado en Microsoft Graph."

                );

            }


            usuarioSeleccionado = {

                id:
                    usuarioDirecto.id,

                nombre:
                    usuarioDirecto.displayName ||
                    "Sin nombre",

                correo:
                    normalizarCorreo(

                        usuarioDirecto.mail ||

                        usuarioDirecto.userPrincipalName ||

                        correo

                    ),

                userPrincipalName:
                    usuarioDirecto.userPrincipalName ||
                    ""

            };

        }

        else {

            const usuario =
                resultado.value[0];


            usuarioSeleccionado = {

                id:
                    usuario.id,

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

        }


        console.log(

            "Usuario encontrado:",

            usuarioSeleccionado

        );


        /*
         * Mostrar nombre real.
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
         * Revisar si ya tiene permisos.
         */

        const registro =
            buscarRegistroPorCorreo(

                usuarioSeleccionado.correo

            );


        if (registro) {

            cargarRegistroEnFormulario(
                registro
            );

        }

        else {

            prepararNuevoUsuario();

        }


        mostrarPanelUsuario();


    }

    catch (error) {

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

    }

    finally {

        if (boton) {

            boton.disabled =
                false;

            boton.textContent =
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

            normalizarTexto(
                grupoRegistrado
            ) ===
            normalizarTexto(
                GRUPO_TI
            )

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

            }

            else {

                control.disabled =
                    false;

            }

        }

    );

}


/* =========================================================
   GUARDAR PERMISOS
========================================================= */

async function guardarPermisos() {

    const boton =
        document.getElementById(
            "btnGuardarPermisos"
        );


    const grupoSelect =
        document.getElementById(
            "grupoUsuario"
        );


    if (!usuarioSeleccionado) {

        alert(

            "Primero busca un usuario mediante Graph."

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


    const permisos =
        {};


    MODULOS_PERMISOS.forEach(

        modulo => {

            const checkbox =
                document.querySelector(

                    `.permiso-modulo[data-modulo="${modulo}"]`

                );


            permisos[
                modulo
            ] =

                checkbox
                    ? checkbox.checked
                    : false;

        }

    );


    /*
     * TI siempre tiene todo.
     */

    if (
        grupo ===
        GRUPO_TI
    ) {

        MODULOS_PERMISOS.forEach(

            modulo => {

                permisos[
                    modulo
                ] =
                    true;

            }

        );

    }


    try {

        if (boton) {

            boton.disabled =
                true;

            boton.textContent =
                "Guardando...";

        }


        const sitio =
            await obtenerSitioSharePoint();


        const lista =
            await obtenerListaPermisos();


        const campos =
            await obtenerMapaCampos();


        const existente =
            buscarRegistroPorCorreo(
                correo
            );


        const fields =
            {};


        /*
         * TITLE
         */

        fields[
            campos.Title
        ] =
            nombre ||
            correo;


        /*
         * CORREO
         */

        fields[
            campos.UsuarioCorreo
        ] =
            correo;


        /*
         * NOMBRE
         */

        fields[
            campos.NombreUsuario
        ] =
            nombre ||
            correo;


        /*
         * GRUPO
         */

        fields[
            campos.Grupo
        ] =
            grupo;


        /*
         * MÓDULOS
         */

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

        if (existente) {

            const url =

                `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items/${existente.id}/fields`;


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

        }


        /*
         * CREAR
         */

        else {

            const url =

                `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items`;


            await graphFetch(

                url,

                {

                    method:
                        "POST",

                    body:
                        JSON.stringify({

                            fields:
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
         * Actualizar tabla.
         */

        await cargarPermisos();


        /*
         * Volver a cargar el registro.
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


    }

    catch (error) {

        console.error(

            "Error guardando permisos:",

            error

        );


        mostrarMensaje(

            `No se pudieron guardar los permisos. ${obtenerMensajeError(error)}`,

            "error"

        );

    }

    finally {

        if (boton) {

            boton.disabled =
                false;

            boton.textContent =
                "Guardar permisos";

        }

    }

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


    return fields[
        clave
    ];

}


/* =========================================================
   CONTAR PERMISOS
========================================================= */

function contarPermisos(
    registro
) {

    const grupo =
        normalizarTexto(

            obtenerCampo(
                registro,
                "Grupo"
            )

        );


    if (
        grupo ===
        normalizarTexto(
            GRUPO_TI
        )
    ) {

        return MODULOS_PERMISOS.length;

    }


    let cantidad =
        0;


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

        texto ===
            "true" ||

        texto ===
            "1" ||

        texto ===
            "yes" ||

        texto ===
            "sí" ||

        texto ===
            "si"

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


    setTimeout(

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
            "Graph 401"
        )
    ) {

        return (

            "La sesión de Microsoft no es válida. Cierra sesión y vuelve a ingresar."

        );

    }


    if (
        mensaje.includes(
            "Graph 403"
        )
    ) {

        return (

            "Microsoft Graph rechazó el acceso a usuarios. Verifica los permisos de la aplicación."

        );

    }


    if (
        mensaje.includes(
            "Graph 404"
        )
    ) {

        return (

            "El usuario no existe en Microsoft Entra ID."

        );

    }


    return (

        `No se pudo buscar el usuario. ${obtenerMensajeError(error)}`

    );

}


/* =========================================================
   ERROR
========================================================= */

function obtenerMensajeError(
    error
) {

    return (

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

            function () {

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


    const inputCorreo =
        document.getElementById(
            "correoUsuario"
        );


    if (inputCorreo) {

        inputCorreo.addEventListener(

            "keydown",

            function (event) {

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

            function (event) {

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