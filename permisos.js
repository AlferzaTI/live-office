(function () {

/* =========================================================
   ALFERZA LIVE OFFICE
   GESTIÓN DE PERMISOS
========================================================= */


/* =========================================================
   CONFIGURACIÓN MICROSOFT
========================================================= */

const MSAL_CONFIG = {

    auth: {

        clientId:
            "5d98417c-74a7-4fab-8f2c-41ac127be696",

        authority:
            "https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1",

        redirectUri:
            new URL(
                "blank.html",
                window.location.href
            ).href

    },

    cache: {

        cacheLocation:
            "sessionStorage",

        storeAuthStateInCookie:
            false

    }

};


/* =========================================================
   SHAREPOINT
========================================================= */

const SHAREPOINT_HOST =
    "alferzaholding-my.sharepoint.com";


const SHAREPOINT_SITE_PATH =
    "/personal/soporte1_alferza_pe";


const PERMISOS_LIST_NAME =
    "PermisosTI";


/* =========================================================
   SCOPES
   LOS MISMOS DE TICKETS.JS
========================================================= */

const SCOPES_GRAPH = [

    "User.Read",

    "Sites.ReadWrite.All"

];


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
   VALIDAR MSAL
========================================================= */

if (
    typeof msal === "undefined"
) {

    mostrarErrorInicial(
        "No cargó la librería de Microsoft (MSAL)."
    );


    throw new Error(
        "MSAL no está disponible."
    );

}


/* =========================================================
   MSAL
========================================================= */

const msalInstance =
    new msal.PublicClientApplication(
        MSAL_CONFIG
    );


/* =========================================================
   VARIABLES
========================================================= */

let cuentaActual =
    null;


let sitioSharePoint =
    null;


let listaPermisos =
    null;


let usuariosGraph =
    [];


let registrosPermisos =
    [];


let usuarioSeleccionado =
    null;


let mapaCampos =
    null;


/* =========================================================
   ELEMENTOS
========================================================= */

const selectorUsuario =
    document.getElementById(
        "selectorUsuario"
    );


const correoUsuario =
    document.getElementById(
        "correoUsuario"
    );


const btnBuscarUsuario =
    document.getElementById(
        "btnBuscarUsuario"
    );


const panelPermisos =
    document.getElementById(
        "panelPermisos"
    );


const nombreUsuario =
    document.getElementById(
        "nombreUsuario"
    );


const correoMostrado =
    document.getElementById(
        "correoMostrado"
    );


const grupoUsuario =
    document.getElementById(
        "grupoUsuario"
    );


const btnGuardarPermisos =
    document.getElementById(
        "btnGuardarPermisos"
    );


const tablaPermisos =
    document.getElementById(
        "tablaPermisos"
    );


const userInfo =
    document.getElementById(
        "userInfo"
    );


/* =========================================================
   INICIO
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    iniciarPermisos
);


/* =========================================================
   MOSTRAR USUARIO EN NAVBAR
========================================================= */

function mostrarUsuarioNavbar() {

    if (!userInfo || !cuentaActual) {

        return;

    }


    userInfo.textContent =

        cuentaActual.name ||

        cuentaActual.username ||

        "";

}
/* =========================================================
   INICIAR
========================================================= */

async function iniciarPermisos() {

    console.log(
        "Iniciando módulo Permisos..."
    );


    configurarEventos();


    ocultarPanel();


    try {

        const cuenta =
            obtenerCuenta();


        if (!cuenta) {

            mostrarMensaje(

                "No existe una sesión de Microsoft.",

                "error"

            );


            return;

        }


        cuentaActual =
            cuenta;


        mostrarUsuarioNavbar();


        await cargarUsuariosGraph();


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


        const correoActual =
    obtenerCorreoCuentaActual();


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


/* =========================================================
   CONFIGURAR EVENTOS
========================================================= */

function configurarEventos() {


    if (
        btnBuscarUsuario
    ) {

        btnBuscarUsuario.addEventListener(

            "click",

            buscarUsuario

        );

    }


    if (
        correoUsuario
    ) {

        correoUsuario.addEventListener(

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


    if (
        selectorUsuario
    ) {

        selectorUsuario.addEventListener(

            "change",

            function () {

                seleccionarUsuario(
                    this.value
                );

            }

        );

    }


    if (
        grupoUsuario
    ) {

        grupoUsuario.addEventListener(

            "change",

            cambiarGrupo

        );

    }


    if (
        btnGuardarPermisos
    ) {

        btnGuardarPermisos.addEventListener(

            "click",

            guardarPermisos

        );

    }


    configurarMenu();

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
   OBTENER CUENTA
========================================================= */

function obtenerCuenta() {

    const cuentas =
        msalInstance.getAllAccounts();


    if (
        cuentas &&
        cuentas.length > 0
    ) {

        cuentaActual =
            cuentas[0];


        msalInstance.setActiveAccount(
            cuentaActual
        );


        return cuentaActual;

    }


    return null;

}


/* =========================================================
   TOKEN
   MISMO MECANISMO QUE TICKETS.JS
========================================================= */

async function obtenerToken() {

    const cuentas =
        msalInstance.getAllAccounts();


    if (!cuentas.length) {

        const error =
            new Error(
                "No existe una sesión de Microsoft."
            );


        error.codigo =
            "SIN_SESION";


        throw error;

    }


    cuentaActual =
        cuentas[0];


    msalInstance.setActiveAccount(
        cuentaActual
    );


    try {

        const resultado =
            await msalInstance
                .acquireTokenSilent({

                    scopes:
                        SCOPES_GRAPH,

                    account:
                        cuentaActual

                });


        return resultado.accessToken;

    }

    catch (error) {

        console.error(

            "No se pudo obtener el token silenciosamente:",

            error

        );


        const nuevoError =
            new Error(

                "No se pudo obtener el token de Microsoft."

            );


        nuevoError.codigo =
            "REQUIERE_INTERACCION";


        throw nuevoError;

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
        await obtenerToken();


    const headers = {

        Authorization:
            `Bearer ${token}`,

        "Content-Type":
            "application/json",

        ...(opciones.headers || {})

    };


    const respuesta =
        await fetch(

            url,

            {

                ...opciones,

                headers

            }

        );


    if (!respuesta.ok) {

        const texto =
            await respuesta.text();


        throw new Error(

            `Graph ${respuesta.status}: ${texto}`

        );

    }


    if (
        respuesta.status ===
        204
    ) {

        return null;

    }


    const texto =
        await respuesta.text();


    if (!texto) {

        return null;

    }


    try {

        return JSON.parse(
            texto
        );

    }

    catch {

        return texto;

    }

}


/* =========================================================
   OBTENER USUARIOS GRAPH
========================================================= */

async function cargarUsuariosGraph() {

    if (
        selectorUsuario
    ) {

        selectorUsuario.innerHTML =

            `
            <option value="">
                Cargando usuarios de Microsoft...
            </option>
            `;

    }


    let url =

        "https://graph.microsoft.com/v1.0/users" +

        "?$select=id,displayName,mail,userPrincipalName,accountEnabled" +

        "&$top=999";


    const usuarios =
        [];


    while (url) {

        const resultado =
            await graphFetch(
                url
            );


        if (
            resultado &&
            Array.isArray(
                resultado.value
            )
        ) {

            usuarios.push(
                ...resultado.value
            );

        }


        url =
            resultado[
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
                        )
                            .toLowerCase();


                    const nombreB =
                        String(
                            b.displayName ||
                            ""
                        )
                            .toLowerCase();


                    return nombreA.localeCompare(

                        nombreB,

                        "es"

                    );

                }

            );


    renderizarSelectorUsuarios(
        usuariosGraph
    );


    console.log(

        "Usuarios ALFERZA obtenidos:",

        usuariosGraph.length

    );

}


/* =========================================================
   RENDERIZAR SELECTOR
========================================================= */

function renderizarSelectorUsuarios(
    usuarios
) {

    if (
        !selectorUsuario
    ) {

        return;

    }


    selectorUsuario.innerHTML =
        "";


    const opcionInicial =
        document.createElement(
            "option"
        );


    opcionInicial.value =
        "";


    opcionInicial.textContent =

        `Selecciona un usuario (${usuarios.length})`;


    selectorUsuario.appendChild(
        opcionInicial
    );


    usuarios.forEach(

        usuario => {

            const correo =

                usuario.mail ||

                usuario.userPrincipalName ||

                "";


            const opcion =
                document.createElement(
                    "option"
                );


            opcion.value =
                usuario.id;


            opcion.textContent =

                `${usuario.displayName || "Sin nombre"} — ${correo}`;


            selectorUsuario.appendChild(
                opcion
            );

        }

    );


    if (
        !usuarios.length
    ) {

        opcionInicial.textContent =
            "No se encontraron usuarios.";

    }

}


/* =========================================================
   BUSCAR USUARIO
========================================================= */

async function buscarUsuario() {

    const texto =
        normalizarTexto(

            correoUsuario
                ? correoUsuario.value
                : ""

        );


    if (!texto) {

        renderizarSelectorUsuarios(
            usuariosGraph
        );


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


    renderizarSelectorUsuarios(
        resultados
    );


    if (
        resultados.length ===
        1
    ) {

        selectorUsuario.value =
            resultados[0].id;


        seleccionarUsuario(
            resultados[0].id
        );

    }

    else if (
        resultados.length ===
        0
    ) {

        mostrarMensaje(

            "No se encontró ningún usuario.",

            "error"

        );

    }

    else {

        mostrarMensaje(

            `${resultados.length} usuarios encontrados. Selecciona uno de la lista.`,

            "info"

        );

    }

}


/* =========================================================
   SELECCIONAR USUARIO
========================================================= */

function seleccionarUsuario(
    usuarioId
) {

    if (!usuarioId) {

        usuarioSeleccionado =
            null;


        ocultarPanel();


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


    if (nombreUsuario) {

        nombreUsuario.textContent =

            usuarioSeleccionado.nombre;

    }


    if (correoMostrado) {

        correoMostrado.textContent =

            usuarioSeleccionado.correo;

    }


    if (correoUsuario) {

        correoUsuario.value =

            usuarioSeleccionado.correo;

    }


    /*
     * Buscar configuración guardada.
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


    mostrarPanel();

}


/* =========================================================
   MOSTRAR PANEL
========================================================= */

function mostrarPanel() {

    if (
        panelPermisos
    ) {

        panelPermisos.style.display =
            "block";

    }

}


/* =========================================================
   OCULTAR PANEL
========================================================= */

function ocultarPanel() {

    if (
        panelPermisos
    ) {

        panelPermisos.style.display =
            "none";

    }

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

        `https://graph.microsoft.com/v1.0/sites/` +

        `${SHAREPOINT_HOST}:${SHAREPOINT_SITE_PATH}`;


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


    /*
     * IGUAL QUE TICKETS.JS:
     *
     * Primero obtenemos todas las listas
     * y luego buscamos por nombre.
     *
     * Esto evita el 404 de:
     * /lists/PermisosTI
     */

    const url =

        `https://graph.microsoft.com/v1.0/sites/` +

        `${sitio.id}/lists` +

        `?$select=id,name,displayName&$top=200`;


    const resultado =
        await graphFetch(
            url
        );


    if (
        !resultado.value ||
        !resultado.value.length
    ) {

        throw new Error(

            "No se encontraron listas en el sitio de SharePoint."

        );

    }


    const objetivo =
        normalizarTexto(
            PERMISOS_LIST_NAME
        );


    listaPermisos =
        resultado.value.find(

            lista =>

                normalizarTexto(
                    lista.displayName
                ) === objetivo

                ||

                normalizarTexto(
                    lista.name
                ) === objetivo

        );


    if (!listaPermisos) {

        throw new Error(

            `No se encontró la lista "${PERMISOS_LIST_NAME}". ` +

            `Listas disponibles: ` +

            resultado.value

                .map(

                    lista =>

                        lista.displayName ||
                        lista.name

                )

                .join(
                    ", "
                )

        );

    }


    console.log(

        "Lista PermisosTI encontrada:",

        listaPermisos

    );


    return listaPermisos;

}


/* =========================================================
   OBTENER COLUMNAS
========================================================= */

async function obtenerColumnasPermisos() {

    const sitio =
        await obtenerSitioSharePoint();


    const lista =
        await obtenerListaPermisos();


    const url =

        `https://graph.microsoft.com/v1.0/sites/` +

        `${sitio.id}/lists/${lista.id}/columns` +

        `?$select=name,displayName,hidden,readOnly&$top=200`;


    const resultado =
        await graphFetch(
            url
        );


    return resultado.value || [];

}


/* =========================================================
   OBTENER MAPA DE CAMPOS
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


    /*
     * COLUMNAS GENERALES
     */

    mapaCampos.Title =
        buscarColumna(
            columnas,
            [
                "Title",
                "Título"
            ]
        );


    mapaCampos.UsuarioCorreo =
        buscarColumna(
            columnas,
            [
                "UsuarioCorreo"
            ]
        );


    mapaCampos.NombreUsuario =
        buscarColumna(
            columnas,
            [
                "NombreUsuario"
            ]
        );


    mapaCampos.Grupo =
        buscarColumna(
            columnas,
            [
                "Grupo"
            ]
        );


    if (!mapaCampos.Title) {

        throw new Error(
            "No se encontró la columna Title/Título."
        );

    }


    if (!mapaCampos.UsuarioCorreo) {

        throw new Error(
            "No se encontró la columna UsuarioCorreo."
        );

    }


    if (!mapaCampos.NombreUsuario) {

        throw new Error(
            "No se encontró la columna NombreUsuario."
        );

    }


    if (!mapaCampos.Grupo) {

        throw new Error(
            "No se encontró la columna Grupo."
        );

    }


    /*
     * COLUMNAS DE PERMISOS
     */

    MODULOS_PERMISOS.forEach(

        modulo => {

            mapaCampos[modulo] =
                buscarColumna(

                    columnas,

                    [
                        modulo
                    ]

                );


            if (
                !mapaCampos[modulo]
            ) {

                throw new Error(

                    `No se encontró la columna "${modulo}" en PermisosTI.`

                );

            }

        }

    );


    console.log(

        "Mapa de campos PermisosTI:",

        mapaCampos

    );


    return mapaCampos;

}


/* =========================================================
   BUSCAR COLUMNA
========================================================= */

function buscarColumna(
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


    return columna
        ? columna.name
        : null;

}


/* =========================================================
   OBTENER TODOS LOS REGISTROS
========================================================= */

async function obtenerTodosLosItems() {

    const sitio =
        await obtenerSitioSharePoint();


    const lista =
        await obtenerListaPermisos();


    let url =

        `https://graph.microsoft.com/v1.0/sites/` +

        `${sitio.id}/lists/${lista.id}/items` +

        `?$expand=fields&$top=500`;


    const items =
        [];


    while (url) {

        const resultado =
            await graphFetch(
                url
            );


        if (
            resultado &&
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

    if (
        tablaPermisos
    ) {

        tablaPermisos.innerHTML =

            `
            <div class="tabla-vacia">

                <span>
                    Cargando configuraciones...
                </span>

            </div>
            `;

    }


    try {

        registrosPermisos =
            await obtenerTodosLosItems();


        renderizarTablaPermisos(
            registrosPermisos
        );


        console.log(

            "Configuraciones cargadas:",

            registrosPermisos.length

        );


    }

    catch (error) {

        console.error(

            "Error cargando PermisosTI:",

            error

        );


        if (
            tablaPermisos
        ) {

            tablaPermisos.innerHTML =

                `
                <div class="tabla-vacia">

                    <span>
                        No se pudieron cargar las configuraciones.
                    </span>

                </div>
                `;

        }


        throw error;

    }

}


/* =========================================================
   RENDERIZAR TABLA
========================================================= */

function renderizarTablaPermisos(
    registros
) {

    if (
        !tablaPermisos
    ) {

        return;

    }


    if (
        !registros ||
        !registros.length
    ) {

        tablaPermisos.innerHTML =

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
                                    ${escaparHTML(
                                        nombre
                                    )}
                                </strong>

                                <span>
                                    ${escaparHTML(
                                        correo
                                    )}
                                </span>

                            </div>

                        </div>

                    </td>


                    <td>

                        <span
                            class="
                                badge
                                ${
                                    esTI
                                        ? "badge-ti"
                                        : "badge-normal"
                                }
                            "
                        >
                            ${escaparHTML(
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

                            ${
                                esTI

                                    ? "Acceso total"

                                    : `${cantidad}/${MODULOS_PERMISOS.length}`

                            }

                        </span>

                    </td>


                    <td>

                        <button

                            type="button"

                            class="table-action-btn"

                            data-editar-permiso="${escaparHTML(
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


    tablaPermisos.innerHTML =

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


    tablaPermisos
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
   CARGAR REGISTRO
========================================================= */

function cargarRegistroEnFormulario(
    registro
) {

    const grupo =
        obtenerCampo(
            registro,
            "Grupo"
        );


    if (grupoUsuario) {

        grupoUsuario.value =

            normalizarTexto(
                grupo
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

    if (grupoUsuario) {

        grupoUsuario.value =
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
        grupoUsuario
            ? grupoUsuario.value
            : GRUPO_NORMAL;


    const esTI =
        grupo ===
        GRUPO_TI;


    document
        .querySelectorAll(
            ".permiso-modulo"
        )
        .forEach(

            checkbox => {

                if (esTI) {

                    checkbox.checked =
                        true;

                    checkbox.disabled =
                        true;

                }

                else {

                    checkbox.disabled =
                        false;

                }

            }

        );

}


/* =========================================================
   GUARDAR PERMISOS
========================================================= */

async function guardarPermisos() {

    if (!usuarioSeleccionado) {

        alert(
            "Primero selecciona un usuario."
        );


        return;

    }


    const grupo =
        grupoUsuario
            ? grupoUsuario.value
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
     * TI siempre tiene todos los módulos.
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

        if (btnGuardarPermisos) {

            btnGuardarPermisos.disabled =
                true;

            btnGuardarPermisos.textContent =
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

                usuarioSeleccionado.correo

            );


        const fields =
            {};


        /*
         * NOMBRE
         */

        fields[
            campos.Title
        ] =

            usuarioSeleccionado.nombre ||
            usuarioSeleccionado.correo;


        /*
         * CORREO
         */

        fields[
            campos.UsuarioCorreo
        ] =

            usuarioSeleccionado.correo;


        /*
         * NOMBRE USUARIO
         */

        fields[
            campos.NombreUsuario
        ] =

            usuarioSeleccionado.nombre ||
            usuarioSeleccionado.correo;


        /*
         * GRUPO
         */

        fields[
            campos.Grupo
        ] =
            grupo;


        /*
         * PERMISOS
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
         * ACTUALIZAR REGISTRO EXISTENTE
         */

        if (existente) {

            const url =

                `https://graph.microsoft.com/v1.0/sites/` +

                `${sitio.id}/lists/${lista.id}/items/` +

                `${existente.id}/fields`;


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
         * CREAR REGISTRO NUEVO
         */

        else {

            const url =

                `https://graph.microsoft.com/v1.0/sites/` +

                `${sitio.id}/lists/${lista.id}/items`;


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
         * Recargar configuraciones.
         */

        await cargarPermisos();


        /*
         * Volver a cargar usuario.
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

            "No se pudieron guardar los permisos. " +

            obtenerMensajeError(
                error
            ),

            "error"

        );

    }

    finally {

        if (btnGuardarPermisos) {

            btnGuardarPermisos.disabled =
                false;

            btnGuardarPermisos.textContent =
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

            "El usuario ya no se encuentra en Microsoft Graph."

        );


        return;

    }


    if (selectorUsuario) {

        selectorUsuario.value =
            usuario.id;

    }


    seleccionarUsuario(
        usuario.id
    );

}


/* =========================================================
   VALIDAR ADMIN TI
========================================================= */

async function validarAdministradorTI() {

    const correoActual =
        normalizarCorreo(

            cuentaActual?.username ||
            cuentaActual?.name ||
            ""

        );


    const usuarioActualCorreo =
        obtenerCorreoCuentaActual();


    const esTIInicial =
        USUARIOS_TI_INICIALES
            .map(
                normalizarCorreo
            )
            .includes(
                usuarioActualCorreo
            );


    if (
        esTIInicial
    ) {

        console.log(
            "Administrador TI autorizado."
        );


        return true;

    }


    const registro =
        buscarRegistroPorCorreo(
            usuarioActualCorreo
        );


    if (
        registro &&
        normalizarTexto(

            obtenerCampo(
                registro,
                "Grupo"
            )

        ) ===
            normalizarTexto(
                GRUPO_TI
            )
    ) {

        console.log(
            "Administrador TI autorizado mediante PermisosTI."
        );


        return true;

    }


    mostrarAccesoDenegado();


    return false;

}


/* =========================================================
   CORREO CUENTA ACTUAL
========================================================= */

function obtenerCorreoCuentaActual() {

    if (!cuentaActual) {

        return "";

    }


    return normalizarCorreo(

        cuentaActual.username ||

        cuentaActual.mail ||

        cuentaActual.userPrincipalName ||

        ""

    );

}


/* =========================================================
   MOSTRAR ACCESO DENEGADO
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
   OBTENER CAMPO
========================================================= */

function obtenerCampo(
    registro,
    nombre
) {

    if (!registro) {

        return "";

    }


    const fields =
        registro.fields ||
        {};


    const objetivo =
        normalizarTexto(
            nombre
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

function escaparHTML(
    valor
) {

    if (
        valor === null ||
        valor === undefined
    ) {

        return "";

    }


    return String(
        valor
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

        function () {

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
   ERROR INICIAL
========================================================= */

function mostrarErrorInicial(
    mensaje
) {

    const usuario =
        document.getElementById(
            "userInfo"
        );


    if (usuario) {

        usuario.textContent =
            mensaje;

    }

}


/* =========================================================
   ERROR
========================================================= */

function obtenerMensajeError(
    error
) {

    if (
        !error
    ) {

        return "Error desconocido.";

    }


    const mensaje =
        String(
            error.message ||
            error
        );


    if (
        mensaje.includes(
            "Graph 401"
        )
    ) {

        return (

            "La sesión de Microsoft no es válida. Cierra sesión y vuelve a iniciar sesión."

        );

    }


    if (
        mensaje.includes(
            "Graph 403"
        )
    ) {

        return (

            "Microsoft Graph rechazó la operación. Revisa los permisos ya concedidos a ALFERZA Live Office."

        );

    }


    return mensaje;

}


/* =========================================================
   EXPONER FUNCIONES
========================================================= */

window.seleccionarUsuario =
    seleccionarUsuario;


window.buscarUsuario =
    buscarUsuario;


window.guardarPermisos =
    guardarPermisos;


window.editarRegistro =
    editarRegistro;

})();