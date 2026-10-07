/* =========================================================
   ALFERZA LIVE OFFICE
   GESTIÓN DE PERMISOS
========================================================= */


/* =========================================================
   CONFIGURACIÓN
========================================================= */

const PERMISOS_TENANT_ID =
    "dbab984f-4bb1-4b60-9dff-da59f54acdf1";


const PERMISOS_CLIENT_ID =
    "5d98417c-74a7-4fab-8f2c-41ac127be696";


const PERMISOS_REDIRECT_URI =
    "https://AlferzaTI.github.io/live-office/blank.html";


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
   SCOPES
========================================================= */

const PERMISOS_SCOPES = [

    "User.Read",

    "User.ReadBasic.All",

    "Sites.ReadWrite.All"

];


/* =========================================================
   CONFIGURACIÓN MSAL
========================================================= */

const permisosMsalConfig = {

    auth: {

        clientId:
            PERMISOS_CLIENT_ID,

        authority:
            `https://login.microsoftonline.com/${PERMISOS_TENANT_ID}`,

        redirectUri:
            PERMISOS_REDIRECT_URI

    },

    cache: {

        cacheLocation:
            "sessionStorage",

        storeAuthStateInCookie:
            false

    }

};


/* =========================================================
   INSTANCIA MSAL
========================================================= */

const permisosMsalInstance =
    new msal.PublicClientApplication(
        permisosMsalConfig
    );


/* =========================================================
   VARIABLES
========================================================= */

let usuariosGraph =
    [];


let registrosPermisos =
    [];


let usuarioSeleccionado =
    null;


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


        configurarEventos();


        try {

            const cuenta =
                obtenerCuentaMSAL();


            if (!cuenta) {

                mostrarMensaje(

                    "No existe una sesión de Microsoft disponible. Cierra sesión y vuelve a ingresar.",

                    "error"

                );

                return;

            }


            mostrarUsuarioNavbar(
                cuenta
            );


            /*
             * 1. Cargar usuarios directamente
             *    desde Microsoft Graph.
             */

            await cargarUsuariosGraph();


            /*
             * 2. Intentar cargar configuración
             *    guardada en PermisosTI.
             *
             *    Si falla, los usuarios de Graph
             *    siguen funcionando.
             */

            try {

                await cargarPermisos();

            }

            catch (errorLista) {

                console.error(

                    "No se pudo cargar PermisosTI:",

                    errorLista

                );


                mostrarMensaje(

                    "Los usuarios de Microsoft se cargaron correctamente, pero no se pudo acceder a PermisosTI.",

                    "error"

                );

            }


            /*
             * 3. Validar que el usuario actual
             *    sea TI.
             */

            const correoActual =
                obtenerCorreoCuenta(
                    cuenta
                );


            const registroActual =
                buscarRegistroPorCorreo(
                    correoActual
                );


            const esTIInicial =
                USUARIOS_TI_INICIALES
                    .map(
                        normalizarCorreo
                    )
                    .includes(
                        normalizarCorreo(
                            correoActual
                        )
                    );


            const esTILista =
                registroActual &&

                normalizarTexto(
                    obtenerCampo(
                        registroActual,
                        "Grupo"
                    )
                ) ===
                normalizarTexto(
                    GRUPO_TI
                );


            if (
                !esTIInicial &&
                !esTILista
            ) {

                mostrarAccesoDenegado();

                return;

            }


            console.log(
                "Usuario autorizado para Permisos."
            );

        }

        catch (error) {

            console.error(

                "Error inicializando Permisos:",

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
);


/* =========================================================
   EVENTOS
========================================================= */

function configurarEventos() {

    const botonBuscar =
        document.getElementById(
            "btnBuscarUsuario"
        );


    if (botonBuscar) {

        botonBuscar.addEventListener(

            "click",

            buscarUsuario

        );

    }


    const inputBuscar =
        document.getElementById(
            "correoUsuario"
        );


    if (inputBuscar) {

        inputBuscar.addEventListener(

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


    const selector =
        document.getElementById(
            "selectorUsuario"
        );


    if (selector) {

        selector.addEventListener(

            "change",

            function () {

                seleccionarUsuario(
                    this.value
                );

            }

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


    const guardar =
        document.getElementById(
            "btnGuardarPermisos"
        );


    if (guardar) {

        guardar.addEventListener(

            "click",

            guardarPermisos

        );

    }

}


/* =========================================================
   CUENTA MSAL
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
                "No se pudo establecer la cuenta activa.",
                error
            );

        }


        return cuenta;

    }


    return null;

}


/* =========================================================
   NAVBAR
========================================================= */

function mostrarUsuarioNavbar(
    cuenta
) {

    const elemento =
        document.getElementById(
            "userInfo"
        );


    if (!elemento) {

        return;

    }


    elemento.textContent =

        cuenta.name ||

        cuenta.username ||

        "";

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
   TOKEN GRAPH
========================================================= */

async function obtenerTokenGraph() {

    const cuenta =
        obtenerCuentaMSAL();


    if (!cuenta) {

        throw new Error(

            "No existe una sesión de Microsoft."

        );

    }


    try {

        const respuesta =

            await permisosMsalInstance
                .acquireTokenSilent({

                    scopes:
                        PERMISOS_SCOPES,

                    account:
                        cuenta

                });


        return respuesta.accessToken;

    }

    catch (errorSilent) {

        console.warn(

            "El token silencioso requiere interacción:",

            errorSilent

        );


        /*
         * Solo pedimos interacción si realmente
         * hace falta el scope adicional.
         *
         * Si el consentimiento de administrador
         * ya está configurado, no debería pedir
         * aprobación administrativa.
         */

        try {

            const respuesta =

                await permisosMsalInstance
                    .acquireTokenPopup({

                        scopes:
                            PERMISOS_SCOPES,

                        account:
                            cuenta

                    });


            return respuesta.accessToken;

        }

        catch (errorPopup) {

            console.error(

                "Error obteniendo token Graph:",

                errorPopup

            );


            throw new Error(

                "No se pudo obtener autorización para Microsoft Graph. Verifica User.ReadBasic.All y Sites.ReadWrite.All."

            );

        }

    }

}


/* =========================================================
   GRAPH FETCH
========================================================= */

async function graphFetch(
    url,
    opciones = {}
) {

    const token =
        await obtenerTokenGraph();


    const headers = {

        Authorization:
            `Bearer ${token}`,

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

                headers: {

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

            const datos =
                await respuesta.json();


            detalle =
                datos
                    ?.error
                    ?.message ||
                "";

        }

        catch (error) {

            // Sin contenido JSON.

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
   CARGAR TODOS LOS USUARIOS DE GRAPH
========================================================= */

async function cargarUsuariosGraph() {

    const selector =
        document.getElementById(
            "selectorUsuario"
        );


    if (selector) {

        selector.innerHTML =

            `
            <option value="">
                Cargando usuarios de Microsoft...
            </option>
            `;

    }


    let url =

        "https://graph.microsoft.com/v1.0/users?$select=id,displayName,mail,userPrincipalName,accountEnabled&$top=999";


    const usuarios =
        [];


    while (url) {

        const respuesta =
            await graphFetch(
                url
            );


        if (
            Array.isArray(
                respuesta.value
            )
        ) {

            usuarios.push(
                ...respuesta.value
            );

        }


        url =
            respuesta[
                "@odata.nextLink"
            ] ||

            null;

    }


    /*
     * Solo usuarios ALFERZA.
     */

    usuariosGraph =
        usuarios
            .filter(

                usuario => {

                    const correo =

                        usuario.mail ||

                        usuario.userPrincipalName ||

                        "";


                    return normalizarCorreo(
                        correo
                    ).endsWith(
                        "@alferza.pe"
                    );

                }

            )
            .sort(

                (a, b) => {

                    const nombreA =
                        String(
                            a.displayName ||
                            ""
                        ).toLowerCase();


                    const nombreB =
                        String(
                            b.displayName ||
                            ""
                        ).toLowerCase();


                    return nombreA.localeCompare(
                        nombreB,
                        "es"
                    );

                }

            );


    renderizarUsuariosGraph();


    console.log(

        `Usuarios ALFERZA encontrados: ${usuariosGraph.length}`

    );

}


/* =========================================================
   RENDERIZAR SELECTOR
========================================================= */

function renderizarUsuariosGraph(
    lista = usuariosGraph
) {

    const selector =
        document.getElementById(
            "selectorUsuario"
        );


    if (!selector) {

        return;

    }


    selector.innerHTML =
        "";


    const opcionInicial =
        document.createElement(
            "option"
        );


    opcionInicial.value =
        "";


    opcionInicial.textContent =

        lista.length ===
            usuariosGraph.length

            ? `Selecciona un usuario (${usuariosGraph.length})`

            : `${lista.length} usuarios encontrados`;


    selector.appendChild(
        opcionInicial
    );


    lista.forEach(

        usuario => {

            const correo =
                normalizarCorreo(

                    usuario.mail ||

                    usuario.userPrincipalName ||

                    ""

                );


            const opcion =
                document.createElement(
                    "option"
                );


            opcion.value =
                usuario.id;


            opcion.textContent =

                `${usuario.displayName || "Sin nombre"} — ${correo}`;


            selector.appendChild(
                opcion
            );

        }

    );


    if (
        lista.length === 0
    ) {

        opcionInicial.textContent =
            "No se encontraron usuarios.";

    }

}


/* =========================================================
   BUSCAR USUARIO
========================================================= */

async function buscarUsuario() {

    const input =
        document.getElementById(
            "correoUsuario"
        );


    const selector =
        document.getElementById(
            "selectorUsuario"
        );


    const boton =
        document.getElementById(
            "btnBuscarUsuario"
        );


    if (!input || !selector) {

        return;

    }


    const texto =
        normalizarTexto(
            input.value
        );


    if (!texto) {

        renderizarUsuariosGraph();

        selector.focus();

        return;

    }


    const resultados =
        usuariosGraph.filter(

            usuario => {

                const nombre =
                    normalizarTexto(
                        usuario.displayName
                    );


                const correo =
                    normalizarTexto(

                        usuario.mail ||

                        usuario.userPrincipalName ||

                        ""

                    );


                return (

                    nombre.includes(
                        texto
                    ) ||

                    correo.includes(
                        texto
                    )

                );

            }

        );


    renderizarUsuariosGraph(
        resultados
    );


    if (
        resultados.length === 1
    ) {

        selector.value =
            resultados[0].id;


        await seleccionarUsuario(
            resultados[0].id
        );

    }

    else if (
        resultados.length > 1
    ) {

        selector.focus();


        mostrarMensaje(

            `${resultados.length} usuarios encontrados. Selecciona uno de la lista.`,

            "info"

        );

    }

    else {

        mostrarMensaje(

            "No se encontró ningún usuario.",

            "error"

        );

    }


    if (boton) {

        boton.textContent =
            "Buscar";

    }

}


/* =========================================================
   SELECCIONAR USUARIO
========================================================= */

async function seleccionarUsuario(
    usuarioId
) {

    if (!usuarioId) {

        usuarioSeleccionado =
            null;

        ocultarPanelUsuario();

        return;

    }


    const usuario =
        usuariosGraph.find(

            item =>
                String(
                    item.id
                ) ===
                String(
                    usuarioId
                )

        );


    if (!usuario) {

        return;

    }


    const correo =
        normalizarCorreo(

            usuario.mail ||

            usuario.userPrincipalName ||

            ""

        );


    usuarioSeleccionado = {

        id:
            usuario.id,

        nombre:
            usuario.displayName ||
            "Sin nombre",

        correo:
            correo,

        userPrincipalName:
            usuario.userPrincipalName ||
            ""

    };


    console.log(

        "Usuario seleccionado:",

        usuarioSeleccionado

    );


    const nombreUsuario =
        document.getElementById(
            "nombreUsuario"
        );


    const correoMostrado =
        document.getElementById(
            "correoMostrado"
        );


    const input =
        document.getElementById(
            "correoUsuario"
        );


    if (nombreUsuario) {

        nombreUsuario.textContent =
            usuarioSeleccionado.nombre;

    }


    if (correoMostrado) {

        correoMostrado.textContent =
            usuarioSeleccionado.correo;

    }


    if (input) {

        input.value =
            usuarioSeleccionado.correo;

    }


    /*
     * Buscar configuración existente.
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
   SHAREPOINT
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
   OBTENER LISTA
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


    const respuesta =
        await graphFetch(
            url
        );


    columnasPermisos =
        respuesta.value ||
        [];


    return columnasPermisos;

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


    let columna =
        columnas.find(

            item =>

                objetivos.includes(

                    normalizarTexto(
                        item.displayName
                    )

                )

        );


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
   MAPA DE CAMPOS
========================================================= */

async function obtenerMapaCampos() {

    if (
        mapaCampos
    ) {

        return mapaCampos;

    }


    const columnas =
        await obtenerColumnasPermisos();


    mapaCampos =
        {};


    const nombresBase = [

        "Title",

        "UsuarioCorreo",

        "NombreUsuario",

        "Grupo"

    ];


    for (
        const nombre of
        nombresBase
    ) {

        const columna =
            encontrarColumna(

                columnas,

                nombre ===
                    "Title"

                    ? [
                        "Title",
                        "Título"
                    ]

                    : [
                        nombre
                    ]

            );


        if (!columna) {

            throw new Error(

                `No se encontró la columna ${nombre} en PermisosTI.`

            );

        }


        mapaCampos[nombre] =
            columna.name;

    }


    MODULOS_PERMISOS.forEach(

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

                    `No se encontró la columna ${modulo} en PermisosTI.`

                );

            }


            mapaCampos[
                modulo
            ] =
                columna.name;

        }

    );


    console.log(
        "Mapa PermisosTI:",
        mapaCampos
    );


    return mapaCampos;

}


/* =========================================================
   OBTENER ITEMS
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

        const respuesta =
            await graphFetch(
                url
            );


        if (
            Array.isArray(
                respuesta.value
            )
        ) {

            items.push(
                ...respuesta.value
            );

        }


        url =
            respuesta[
                "@odata.nextLink"
            ] ||
            null;

    }


    return items;

}


/* =========================================================
   CARGAR CONFIGURACIONES
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
                Cargando configuraciones...
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
                );


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


            const completo =
                esTI ||

                cantidad ===
                    MODULOS_PERMISOS.length;


            const textoAcceso =
                esTI

                    ? "Acceso total"

                    : `${cantidad}/${MODULOS_PERMISOS.length}`;


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
                            class="badge ${
                                esTI
                                    ? "badge-ti"
                                    : "badge-normal"
                            }"
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
                                ${
                                    completo
                                        ? "full"
                                        : "partial"
                                }
                            "
                        >
                            ${escapeHtml(
                                textoAcceso
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


    /*
     * Botones editar.
     */

    tabla
        .querySelectorAll(
            "[data-editar-permiso]"
        )
        .forEach(

            boton => {

                boton.addEventListener(

                    "click",

                    function () {

                        editarRegistro(

                            this.getAttribute(
                                "data-editar-permiso"
                            )

                        );

                    }

                );

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
        normalizarCorreo(

            obtenerCampo(
                registro,
                "UsuarioCorreo"
            )

        );


    const usuario =
        usuariosGraph.find(

            item =>

                normalizarCorreo(

                    item.mail ||

                    item.userPrincipalName ||

                    ""

                ) === correo

        );


    if (!usuario) {

        alert(

            "El usuario ya no fue encontrado en Microsoft Graph."

        );

        return;

    }


    const selector =
        document.getElementById(
            "selectorUsuario"
        );


    if (selector) {

        selector.value =
            usuario.id;

    }


    await seleccionarUsuario(
        usuario.id
    );

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

            "Primero selecciona un usuario."

        );

        return;

    }


    const grupo =
        grupoSelect

            ? grupoSelect.value

            : GRUPO_NORMAL;


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
     * TI = acceso total.
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


        const registroExistente =
            buscarRegistroPorCorreo(

                usuarioSeleccionado.correo

            );


        const fields =
            {};


        fields[
            campos.Title
        ] =

            usuarioSeleccionado.nombre ||
            usuarioSeleccionado.correo;


        fields[
            campos.UsuarioCorreo
        ] =

            usuarioSeleccionado.correo;


        fields[
            campos.NombreUsuario
        ] =

            usuarioSeleccionado.nombre ||
            usuarioSeleccionado.correo;


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

        if (
            registroExistente
        ) {

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

                "Permisos guardados correctamente.",

                "success"

            );

        }


        /*
         * Volver a cargar
         * configuraciones.
         */

        await cargarPermisos();


        /*
         * Recuperar el registro
         * actualizado.
         */

        const actualizado =
            buscarRegistroPorCorreo(

                usuarioSeleccionado.correo

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
   MOSTRAR MENSAJE
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
   ERROR DE BÚSQUEDA
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

            "La sesión de Microsoft no es válida."

        );

    }


    if (
        mensaje.includes(
            "Graph 403"
        )
    ) {

        return (

            "Graph rechazó el acceso. Verifica User.ReadBasic.All."

        );

    }


    if (
        mensaje.includes(
            "Graph 404"
        )
    ) {

        return (

            "No se encontró el usuario."

        );

    }


    return (

        `No se pudo buscar el usuario. ${mensaje}`

    );

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

    return (

        error?.message ||

        "Error desconocido."

    );

}