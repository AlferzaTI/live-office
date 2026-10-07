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
   GRUPOS
========================================================= */

const GRUPO_TI =
    "TI";


const GRUPO_NORMAL =
    "Normal";


/*
 * USUARIOS TI INICIALES
 *
 * Esto sirve para que puedas entrar a la
 * página de permisos aunque todavía no
 * exista un registro en PermisosTI.
 *
 * Agrega aquí los correos iniciales de TI.
 */

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

const SCOPES_GRAPH = [

    "User.Read",

    "Sites.ReadWrite.All"

];


/* =========================================================
   MSAL
========================================================= */

if (
    typeof msal === "undefined"
) {

    throw new Error(
        "MSAL no está disponible."
    );

}


const msalInstance =
    new msal.PublicClientApplication(
        MSAL_CONFIG
    );


/* =========================================================
   VARIABLES
========================================================= */

let cuentaActual =
    null;


let usuarioActualGraph =
    null;


let sitioSharePoint =
    null;


let listaPermisos =
    null;


let columnasPermisos =
    [];


let usuariosPermisos =
    [];


let usuarioSeleccionado =
    null;


let grupoSeleccionado =
    GRUPO_NORMAL;


/* =========================================================
   DOM
========================================================= */

const loadingOverlay =
    document.getElementById(
        "loadingOverlay"
    );


const loadingMessage =
    document.getElementById(
        "loadingMessage"
    );


const grupoUsuarioActual =
    document.getElementById(
        "grupoUsuarioActual"
    );


const correoUsuario =
    document.getElementById(
        "correoUsuario"
    );


const btnBuscarUsuario =
    document.getElementById(
        "btnBuscarUsuario"
    );


const btnNuevoUsuario =
    document.getElementById(
        "btnNuevoUsuario"
    );


const buscarMensaje =
    document.getElementById(
        "buscarMensaje"
    );


const usuarioSection =
    document.getElementById(
        "usuarioSection"
    );


const usuarioGrupoBadge =
    document.getElementById(
        "usuarioGrupoBadge"
    );


const usuarioSeleccionadoCorreo =
    document.getElementById(
        "usuarioSeleccionadoCorreo"
    );


const usuarioSeleccionadoNombre =
    document.getElementById(
        "usuarioSeleccionadoNombre"
    );


const btnConfigurarPermisos =
    document.getElementById(
        "btnConfigurarPermisos"
    );


const btnActualizarPermisos =
    document.getElementById(
        "btnActualizarPermisos"
    );


const totalUsuarios =
    document.getElementById(
        "totalUsuarios"
    );


const usuariosTI =
    document.getElementById(
        "usuariosTI"
    );


const usuariosNormales =
    document.getElementById(
        "usuariosNormales"
    );


const usuariosTableBody =
    document.getElementById(
        "usuariosTableBody"
    );


/* =========================================================
   DOM MODAL
========================================================= */

const permisosModal =
    document.getElementById(
        "permisosModal"
    );


const cerrarPermisosModalBtn =
    document.getElementById(
        "cerrarPermisosModal"
    );


const cancelarPermisosBtn =
    document.getElementById(
        "cancelarPermisos"
    );


const guardarPermisosBtn =
    document.getElementById(
        "guardarPermisos"
    );


const modalUsuarioCorreo =
    document.getElementById(
        "modalUsuarioCorreo"
    );


const grupoTI =
    document.getElementById(
        "grupoTI"
    );


const grupoNormal =
    document.getElementById(
        "grupoNormal"
    );


const btnTodosPermisos =
    document.getElementById(
        "btnTodosPermisos"
    );


const modalPermisosMensaje =
    document.getElementById(
        "modalPermisosMensaje"
    );


/* =========================================================
   UTILIDADES
========================================================= */

function normalizarTexto(
    valor
) {

    return String(
        valor || ""
    )

        .toLowerCase()

        .normalize("NFD")

        .replace(
            /[\u0300-\u036f]/g,
            ""
        )

        .trim();

}


function normalizarCampo(
    valor
) {

    return normalizarTexto(
        valor
    )

        .replace(
            /[\s_-]/g,
            ""
        );

}


function escaparHTML(
    valor
) {

    if (
        valor === null ||
        valor === undefined
    ) {

        return "";

    }


    return String(valor)

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
   LOADING
========================================================= */

function cambiarMensajeLoading(
    mensaje
) {

    loadingMessage.textContent =
        mensaje;

}


function ocultarLoading() {

    loadingOverlay.classList.add(
        "hidden"
    );

}


/* =========================================================
   TOKEN
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


    try {

        const resultado =
            await msalInstance.acquireTokenSilent({

                scopes:
                    SCOPES_GRAPH,

                account:
                    cuentaActual

            });


        return resultado.accessToken;

    }

    catch (error) {

        console.warn(
            "No se pudo obtener token silencioso:",
            error
        );


        const nuevoError =
            new Error(
                "Se requiere autorización de Microsoft."
            );


        nuevoError.codigo =
            "REQUIERE_INTERACCION";


        throw nuevoError;

    }

}


/* =========================================================
   GRAPH
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
        respuesta.status === 204
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
   SITIO
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
   LISTA
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
        `https://graph.microsoft.com/v1.0/sites/` +
        `${sitio.id}/lists` +
        `?$select=id,name,displayName&$top=200`;


    const resultado =
        await graphFetch(
            url
        );


    const objetivo =
        normalizarTexto(
            PERMISOS_LIST_NAME
        );


    listaPermisos =
        (resultado.value || []).find(

            lista =>

                normalizarTexto(
                    lista.displayName
                ) === objetivo ||

                normalizarTexto(
                    lista.name
                ) === objetivo

        );


    if (
        !listaPermisos
    ) {

        throw new Error(

            `No se encontró la lista "${PERMISOS_LIST_NAME}".`

        );

    }


    return listaPermisos;

}


/* =========================================================
   COLUMNAS
========================================================= */

async function obtenerColumnasPermisos() {

    if (
        columnasPermisos.length
    ) {

        return columnasPermisos;

    }


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


    columnasPermisos =
        resultado.value || [];


    console.log(
        "Columnas de PermisosTI:",
        columnasPermisos
    );


    return columnasPermisos;

}


/* =========================================================
   RESOLVER NOMBRE INTERNO DE COLUMNA
========================================================= */

function obtenerNombreInternoColumna(
    displayName,
    fallback
) {

    const objetivo =
        normalizarCampo(
            displayName
        );


    const columna =
        columnasPermisos.find(

            item =>

                normalizarCampo(
                    item.displayName
                ) === objetivo

        );


    if (
        columna
    ) {

        return columna.name;

    }


    return fallback;

}


/* =========================================================
   MAPEO DE CAMPOS
========================================================= */

function obtenerCampoCorreo() {

    return obtenerNombreInternoColumna(
        "UsuarioCorreo",
        "UsuarioCorreo"
    );

}


function obtenerCampoNombre() {

    return obtenerNombreInternoColumna(
        "NombreUsuario",
        "NombreUsuario"
    );

}


function obtenerCampoGrupo() {

    return obtenerNombreInternoColumna(
        "Grupo",
        "Grupo"
    );

}


function obtenerCampoModulo(
    modulo
) {

    return obtenerNombreInternoColumna(
        modulo,
        modulo
    );

}


/* =========================================================
   USUARIO ACTUAL
========================================================= */

async function obtenerUsuarioActual() {

    const token =
        await obtenerToken();


    const respuesta =
        await fetch(

            "https://graph.microsoft.com/v1.0/me",

            {

                headers: {

                    Authorization:
                        `Bearer ${token}`

                }

            }

        );


    if (
        !respuesta.ok
    ) {

        throw new Error(

            `No se pudo obtener el usuario (${respuesta.status}).`

        );

    }


    usuarioActualGraph =
        await respuesta.json();


    cuentaActual =
        msalInstance
            .getAllAccounts()[0];


    return usuarioActualGraph;

}


/* =========================================================
   CARGAR USUARIOS
========================================================= */

async function cargarUsuariosPermisos() {

    const sitio =
        await obtenerSitioSharePoint();


    const lista =
        await obtenerListaPermisos();


    const url =
        `https://graph.microsoft.com/v1.0/sites/` +
        `${sitio.id}/lists/${lista.id}/items` +
        `?$expand=fields&$top=999`;


    const resultado =
        await graphFetch(
            url
        );


    usuariosPermisos =
        resultado.value || [];


    console.log(
        "Usuarios de permisos:",
        usuariosPermisos
    );


    actualizarResumenUsuarios();


    renderizarUsuariosPermisos();


    return usuariosPermisos;

}


/* =========================================================
   BUSCAR USUARIO EN LISTA
========================================================= */

function buscarUsuarioEnLista(
    correo
) {

    const correoNormalizado =
        normalizarTexto(
            correo
        );


    const campoCorreo =
        obtenerCampoCorreo();


    return usuariosPermisos.find(

        item => {

            const fields =
                item.fields || {};


            const valor =
                normalizarTexto(
                    fields[campoCorreo] ||
                    fields.UsuarioCorreo ||
                    ""
                );


            return (
                valor ===
                correoNormalizado
            );

        }

    ) || null;

}


/* =========================================================
   USUARIO ES TI
========================================================= */

function usuarioEsTI(
    correo,
    registro = null
) {

    const correoNormalizado =
        normalizarTexto(
            correo
        );


    const inicial =
        USUARIOS_TI_INICIALES.some(

            correoInicial =>

                normalizarTexto(
                    correoInicial
                ) === correoNormalizado

        );


    if (
        inicial
    ) {

        return true;

    }


    if (
        registro
    ) {

        const fields =
            registro.fields || {};


        const campoGrupo =
            obtenerCampoGrupo();


        const grupo =
            normalizarTexto(

                fields[campoGrupo] ||

                fields.Grupo ||

                ""

            );


        if (
            grupo ===
            normalizarTexto(GRUPO_TI)
        ) {

            return true;

        }

    }


    return false;

}


/* =========================================================
   PERMISOS EFECTIVOS
========================================================= */

function obtenerPermisosRegistro(
    registro
) {

    const permisos = {};


    MODULOS_PERMISOS.forEach(

        modulo => {

            permisos[modulo] =
                false;

        }

    );


    if (
        !registro
    ) {

        return permisos;

    }


    const fields =
        registro.fields || {};


    const correo =
        fields[
            obtenerCampoCorreo()
        ] ||
        fields.UsuarioCorreo ||
        "";


    const campoGrupo =
        obtenerCampoGrupo();


    const grupo =
        fields[campoGrupo] ||
        fields.Grupo ||
        GRUPO_NORMAL;


    if (
        usuarioEsTI(
            correo,
            registro
        ) ||
        normalizarTexto(grupo) ===
            normalizarTexto(GRUPO_TI)
    ) {

        MODULOS_PERMISOS.forEach(

            modulo => {

                permisos[modulo] =
                    true;

            }

        );


        return permisos;

    }


    MODULOS_PERMISOS.forEach(

        modulo => {

            const campo =
                obtenerCampoModulo(
                    modulo
                );


            permisos[modulo] =
                fields[campo] === true;

        }

    );


    return permisos;

}


/* =========================================================
   RESUMEN USUARIOS
========================================================= */

function actualizarResumenUsuarios() {

    const total =
        usuariosPermisos.length;


    const ti =
        usuariosPermisos.filter(

            registro => {

                const fields =
                    registro.fields || {};


                const correo =
                    fields[
                        obtenerCampoCorreo()
                    ] ||
                    fields.UsuarioCorreo ||
                    "";


                return usuarioEsTI(
                    correo,
                    registro
                );

            }

        ).length;


    const normales =
        total - ti;


    totalUsuarios.textContent =
        total;


    usuariosTI.textContent =
        ti;


    usuariosNormales.textContent =
        normales;

}


/* =========================================================
   RENDERIZAR USUARIOS
========================================================= */

function renderizarUsuariosPermisos() {

    if (
        !usuariosPermisos.length
    ) {

        usuariosTableBody.innerHTML = `

            <tr>

                <td
                    colspan="4"
                    class="table-loading"
                >

                    No existen usuarios configurados.

                </td>

            </tr>

        `;


        return;

    }


    const campoCorreo =
        obtenerCampoCorreo();


    const campoNombre =
        obtenerCampoNombre();


    const campoGrupo =
        obtenerCampoGrupo();


    usuariosTableBody.innerHTML =

        usuariosPermisos.map(

            registro => {

                const fields =
                    registro.fields || {};


                const correo =
                    fields[campoCorreo] ||
                    fields.UsuarioCorreo ||
                    "Sin correo";


                const nombre =
                    fields[campoNombre] ||
                    fields.NombreUsuario ||
                    "Usuario";


                const esTI =
                    usuarioEsTI(
                        correo,
                        registro
                    );


                const grupo =
                    esTI
                        ? GRUPO_TI
                        : (
                            fields[campoGrupo] ||
                            fields.Grupo ||
                            GRUPO_NORMAL
                        );


                const permisos =
                    obtenerPermisosRegistro(
                        registro
                    );


                const cantidadAcceso =
                    Object.values(
                        permisos
                    ).filter(
                        Boolean
                    ).length;


                const accesoClase =
                    esTI
                        ? "full"
                        : (
                            cantidadAcceso > 0
                                ? "partial"
                                : ""
                        );


                const accesoTexto =
                    esTI
                        ? "Acceso total"
                        : `${cantidadAcceso}/${MODULOS_PERMISOS.length} módulos`;


                return `

                    <tr>

                        <td>

                            <div class="permission-user-cell">

                                <div class="permission-user-avatar">
                                    👤
                                </div>

                                <div class="permission-user-info">

                                    <strong>
                                        ${escaparHTML(nombre)}
                                    </strong>

                                    <span>
                                        ${escaparHTML(correo)}
                                    </span>

                                </div>

                            </div>

                        </td>


                        <td>

                            <span
                                class="group-badge ${
                                    esTI
                                        ? "ti"
                                        : "normal"
                                }"
                            >

                                ${
                                    escaparHTML(
                                        String(grupo)
                                            .toUpperCase()
                                    )
                                }

                            </span>

                        </td>


                        <td>

                            <span
                                class="access-summary ${accesoClase}"
                            >

                                ${accesoTexto}

                            </span>

                        </td>


                        <td>

                            <button
                                type="button"
                                class="table-action-btn"
                                data-correo="${escaparHTML(correo)}"
                            >

                                ✎ Editar

                            </button>

                        </td>

                    </tr>

                `;

            }

        ).join("");


    document
        .querySelectorAll(
            ".table-action-btn"
        )
        .forEach(

            boton => {

                boton.addEventListener(

                    "click",

                    () => {

                        const correo =
                            boton.dataset.correo;

                        seleccionarUsuario(
                            correo
                        );

                        abrirModalConfiguracion();

                    }

                );

            }

        );

}


/* =========================================================
   SELECCIONAR USUARIO
========================================================= */

function seleccionarUsuario(
    correo
) {

    const registro =
        buscarUsuarioEnLista(
            correo
        );


    const correoLimpio =
        normalizarTexto(
            correo
        );


    if (
        !correoLimpio
    ) {

        return;

    }


    usuarioSeleccionado = {

        registro:
            registro,

        correo:
            correoLimpio,

        nombre:
            registro?.fields?.[
                obtenerCampoNombre()
            ] ||

            registro?.fields?.NombreUsuario ||

            correoLimpio,

        grupo:

            registro?.fields?.[
                obtenerCampoGrupo()
            ] ||

            registro?.fields?.Grupo ||

            (
                usuarioEsTI(
                    correoLimpio,
                    registro
                )

                    ? GRUPO_TI

                    : GRUPO_NORMAL
            )

    };


    usuarioSeleccionadoCorreo.textContent =
        usuarioSeleccionado.correo;


    usuarioSeleccionadoNombre.textContent =
        usuarioSeleccionado.nombre;


    const esTI =
        usuarioEsTI(
            usuarioSeleccionado.correo,
            usuarioSeleccionado.registro
        );


    const grupoReal =
        esTI
            ? GRUPO_TI
            : GRUPO_NORMAL;


    usuarioSeleccionado.grupo =
        grupoReal;


    usuarioGrupoBadge.textContent =
        grupoReal.toUpperCase();


    usuarioGrupoBadge.className =
        `group-badge ${
            esTI
                ? "ti"
                : "normal"
        }`;


    usuarioSection.style.display =
        "block";


    correoUsuario.value =
        usuarioSeleccionado.correo;


    return usuarioSeleccionado;

}


/* =========================================================
   MOSTRAR MENSAJE BÚSQUEDA
========================================================= */

function mostrarBuscarMensaje(
    mensaje,
    tipo
) {

    buscarMensaje.textContent =
        mensaje;


    buscarMensaje.className =
        `permission-message ${tipo}`;

}


/* =========================================================
   BUSCAR POR CORREO
========================================================= */

function ejecutarBusquedaUsuario() {

    const correo =
        correoUsuario.value
            .trim();


    if (
        !correo
    ) {

        mostrarBuscarMensaje(
            "Ingresa un correo corporativo.",
            "error"
        );


        correoUsuario.focus();


        return;

    }


    if (
        !correo.includes("@")
    ) {

        mostrarBuscarMensaje(
            "Ingresa un correo válido.",
            "error"
        );


        correoUsuario.focus();


        return;

    }


    const usuario =
        seleccionarUsuario(
            correo
        );


    if (
        usuario.registro
    ) {

        mostrarBuscarMensaje(

            "Usuario encontrado. Puedes editar sus permisos.",

            "success"

        );

    }

    else {

        mostrarBuscarMensaje(

            "Usuario no configurado. Puedes crear sus permisos.",

            "info"

        );

    }


    abrirModalConfiguracion();

}


/* =========================================================
   NUEVO USUARIO
========================================================= */

function nuevoUsuario() {

    const correo =
        correoUsuario.value.trim();


    if (
        correo
    ) {

        if (
            !correo.includes("@")
        ) {

            mostrarBuscarMensaje(

                "Ingresa un correo válido.",

                "error"

            );


            return;

        }


        seleccionarUsuario(
            correo
        );

    }

    else {

        usuarioSeleccionado = {

            registro:
                null,

            correo:
                "",

            nombre:
                "",

            grupo:
                GRUPO_NORMAL

        };

    }


    abrirModalConfiguracion();

}


/* =========================================================
   ABRIR MODAL
========================================================= */

function abrirModalConfiguracion() {

    if (
        !usuarioSeleccionado
    ) {

        return;

    }


    modalUsuarioCorreo.textContent =

        usuarioSeleccionado.correo ||

        "Nuevo usuario";


    const esTI =
        usuarioEsTI(
            usuarioSeleccionado.correo,
            usuarioSeleccionado.registro
        );


    if (
        esTI
    ) {

        grupoSeleccionado =
            GRUPO_TI;

    }

    else {

        grupoSeleccionado =

            usuarioSeleccionado.grupo ===
            GRUPO_TI

                ? GRUPO_TI

                : GRUPO_NORMAL;

    }


    actualizarGrupoVisual();


    const permisosExistentes =

        obtenerPermisosRegistro(

            usuarioSeleccionado.registro

        );


    MODULOS_PERMISOS.forEach(

        modulo => {

            const boton =
                document.querySelector(
                    `.permission-toggle[data-permission="${modulo}"]`
                );


            if (
                boton
            ) {

                establecerEstadoToggle(

                    boton,

                    permisosExistentes[modulo]

                );

            }

        }

    );


    aplicarReglasGrupo();


    modalPermisosMensaje.textContent =
        "";


    modalPermisosMensaje.className =
        "permission-message";


    permisosModal.style.display =
        "flex";


    document.body.style.overflow =
        "hidden";

}


/* =========================================================
   CERRAR MODAL
========================================================= */

function cerrarModalConfiguracion() {

    permisosModal.style.display =
        "none";


    document.body.style.overflow =
        "";

}


/* =========================================================
   GRUPO
========================================================= */

function actualizarGrupoVisual() {

    grupoTI.classList.toggle(

        "selected",

        grupoSeleccionado ===
            GRUPO_TI

    );


    grupoNormal.classList.toggle(

        "selected",

        grupoSeleccionado ===
            GRUPO_NORMAL

    );

}


/* =========================================================
   TOGGLE
========================================================= */

function establecerEstadoToggle(
    boton,
    activo
) {

    const item =
        boton.closest(
            ".permission-item"
        );


    const label =
        boton.querySelector(
            ".toggle-label"
        );


    boton.classList.toggle(
        "active",
        Boolean(activo)
    );


    if (
        label
    ) {

        label.textContent =
            activo
                ? "SÍ"
                : "NO";

    }


    if (
        item
    ) {

        item.classList.toggle(
            "allowed",
            Boolean(activo)
        );

    }

}


/* =========================================================
   LEER TOGGLE
========================================================= */

function leerEstadoToggle(
    modulo
) {

    const boton =
        document.querySelector(

            `.permission-toggle[data-permission="${modulo}"]`

        );


    if (
        !boton
    ) {

        return false;

    }


    return boton.classList.contains(
        "active"
    );

}


/* =========================================================
   APLICAR REGLAS DE GRUPO
========================================================= */

function aplicarReglasGrupo() {

    const esTI =
        grupoSeleccionado ===
        GRUPO_TI;


    MODULOS_PERMISOS.forEach(

        modulo => {

            const boton =
                document.querySelector(

                    `.permission-toggle[data-permission="${modulo}"]`

                );


            if (
                !boton
            ) {

                return;

            }


            if (
                esTI
            ) {

                establecerEstadoToggle(
                    boton,
                    true
                );


                boton.classList.add(
                    "locked"
                );


                boton.disabled =
                    true;

            }

            else {

                boton.classList.remove(
                    "locked"
                );


                boton.disabled =
                    false;

            }

        }

    );

}


/* =========================================================
   GUARDAR
========================================================= */

async function guardarPermisos() {

    if (
        !usuarioSeleccionado
    ) {

        return;

    }


    const correo =
        usuarioSeleccionado.correo
            .trim();


    if (
        !correo
    ) {

        modalPermisosMensaje.className =
            "permission-message error";


        modalPermisosMensaje.textContent =
            "Debes ingresar un correo.";


        return;

    }


    if (
        !correo.includes("@")
    ) {

        modalPermisosMensaje.className =
            "permission-message error";


        modalPermisosMensaje.textContent =
            "El correo ingresado no es válido.";


        return;

    }


    try {

        guardarPermisosBtn.disabled =
            true;


        guardarPermisosBtn.textContent =
            "⏳ Guardando...";


        const sitio =
            await obtenerSitioSharePoint();


        const lista =
            await obtenerListaPermisos();


        const campoCorreo =
            obtenerCampoCorreo();


        const campoNombre =
            obtenerCampoNombre();


        const campoGrupo =
            obtenerCampoGrupo();


        const registroExistente =
            buscarUsuarioEnLista(
                correo
            );


        let grupo =
            grupoSeleccionado;


        const esTI =
            grupo ===
            GRUPO_TI;


        const fields = {};


        /*
         * Siempre guardamos Title.
         */

        fields.Title =
            usuarioSeleccionado.nombre ||
            correo;


        fields[campoCorreo] =
            correo;


        fields[campoNombre] =
            usuarioSeleccionado.nombre ||
            correo;


        fields[campoGrupo] =
            grupo;


        /*
         * Permisos por módulo.
         *
         * TI = todo en true.
         * Normal = según los toggles.
         */

        MODULOS_PERMISOS.forEach(

            modulo => {

                const campoModulo =
                    obtenerCampoModulo(
                        modulo
                    );


                fields[campoModulo] =
                    esTI

                        ? true

                        : leerEstadoToggle(
                            modulo
                        );

            }

        );


        if (
            registroExistente
        ) {

            const itemId =
                registroExistente.id;


            const url =
                `https://graph.microsoft.com/v1.0/sites/` +
                `${sitio.id}/lists/${lista.id}/items/` +
                `${itemId}/fields`;


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


        }

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

        }


        mostrarModalMensaje(

            registroExistente

                ? "Permisos actualizados correctamente."

                : "Usuario agregado correctamente.",

            "success"

        );


        await cargarUsuariosPermisos();


        seleccionarUsuario(
            correo
        );


        setTimeout(

            () => {

                cerrarModalConfiguracion();

            },

            900

        );

    }

    catch (error) {

        console.error(

            "Error guardando permisos:",

            error

        );


        mostrarModalMensaje(

            "No se pudieron guardar los permisos: " +

            String(
                error.message ||
                error
            ).slice(
                0,
                500
            ),

            "error"

        );

    }

    finally {

        guardarPermisosBtn.disabled =
            false;


        guardarPermisosBtn.textContent =
            "💾 Guardar permisos";

    }

}


/* =========================================================
   MENSAJE MODAL
========================================================= */

function mostrarModalMensaje(
    mensaje,
    tipo
) {

    modalPermisosMensaje.textContent =
        mensaje;


    modalPermisosMensaje.className =
        `permission-message ${tipo}`;

}


/* =========================================================
   TODOS LOS PERMISOS
========================================================= */

function activarTodosPermisos() {

    MODULOS_PERMISOS.forEach(

        modulo => {

            const boton =
                document.querySelector(

                    `.permission-toggle[data-permission="${modulo}"]`

                );


            if (
                boton &&
                !boton.disabled
            ) {

                establecerEstadoToggle(
                    boton,
                    true
                );

            }

        }

    );

}


/* =========================================================
   CAMBIAR GRUPO
========================================================= */

function seleccionarGrupo(
    grupo
) {

    grupoSeleccionado =
        grupo;


    actualizarGrupoVisual();


    aplicarReglasGrupo();

}


/* =========================================================
   EVENTOS DE TOGGLE
========================================================= */

document
    .querySelectorAll(
        ".permission-toggle"
    )
    .forEach(

        boton => {

            boton.addEventListener(

                "click",

                () => {

                    if (
                        boton.disabled
                    ) {

                        return;

                    }


                    const activo =
                        boton.classList.contains(
                            "active"
                        );


                    establecerEstadoToggle(

                        boton,

                        !activo

                    );

                }

            );

        }

    );


/* =========================================================
   EVENTOS
========================================================= */

btnBuscarUsuario.addEventListener(

    "click",

    ejecutarBusquedaUsuario

);


btnNuevoUsuario.addEventListener(

    "click",

    nuevoUsuario

);


correoUsuario.addEventListener(

    "keydown",

    event => {

        if (
            event.key === "Enter"
        ) {

            ejecutarBusquedaUsuario();

        }

    }

);


btnConfigurarPermisos.addEventListener(

    "click",

    abrirModalConfiguracion

);


btnActualizarPermisos.addEventListener(

    "click",

    async () => {

        try {

            btnActualizarPermisos.disabled =
                true;


            btnActualizarPermisos.textContent =
                "⏳ Actualizando...";


            await cargarUsuariosPermisos();


        }

        catch (error) {

            console.error(
                error
            );

        }

        finally {

            btnActualizarPermisos.disabled =
                false;


            btnActualizarPermisos.textContent =
                "↻ Actualizar";

        }

    }

);


grupoTI.addEventListener(

    "click",

    () => {

        seleccionarGrupo(
            GRUPO_TI
        );

    }

);


grupoNormal.addEventListener(

    "click",

    () => {

        seleccionarGrupo(
            GRUPO_NORMAL
        );

    }

);


btnTodosPermisos.addEventListener(

    "click",

    activarTodosPermisos

);


cerrarPermisosModalBtn.addEventListener(

    "click",

    cerrarModalConfiguracion

);


cancelarPermisosBtn.addEventListener(

    "click",

    cerrarModalConfiguracion

);


guardarPermisosBtn.addEventListener(

    "click",

    guardarPermisos

);


/* =========================================================
   CLICK FUERA DEL MODAL
========================================================= */

permisosModal.addEventListener(

    "click",

    event => {

        if (
            event.target ===
            permisosModal
        ) {

            cerrarModalConfiguracion();

        }

    }

);


/* =========================================================
   ESCAPE
========================================================= */

document.addEventListener(

    "keydown",

    event => {

        if (
            event.key !==
            "Escape"
        ) {

            return;

        }


        if (
            permisosModal.style.display ===
            "flex"
        ) {

            cerrarModalConfiguracion();

        }

    }

);


/* =========================================================
   ESPERAR CUENTA
========================================================= */

async function esperarCuenta(
    ms = 6000
) {

    const fin =
        Date.now() + ms;


    while (
        Date.now() < fin
    ) {

        if (
            msalInstance
                .getAllAccounts()
                .length
        ) {

            return true;

        }


        await new Promise(

            resolve =>

                setTimeout(
                    resolve,
                    400
                )

        );

    }


    return false;

}


/* =========================================================
   MOSTRAR LOGIN
========================================================= */

async function solicitarLogin() {

    try {

        const cuentas =
            msalInstance.getAllAccounts();


        if (
            cuentas.length
        ) {

            await msalInstance.acquireTokenPopup({

                scopes:
                    SCOPES_GRAPH,

                account:
                    cuentas[0]

            });

        }

        else {

            await msalInstance.loginPopup({

                scopes:
                    SCOPES_GRAPH

            });

        }


        window.location.reload();

    }

    catch (error) {

        console.error(
            "Error iniciando sesión:",
            error
        );


        cambiarMensajeLoading(
            "No se pudo iniciar sesión."
        );

    }

}


/* =========================================================
   VERIFICAR ACCESO A PERMISOS
========================================================= */

async function verificarAcceso() {

    const correo =
        normalizarTexto(

            usuarioActualGraph.mail ||

            usuarioActualGraph.userPrincipalName ||

            cuentaActual?.username ||

            ""

        );


    if (
        !correo
    ) {

        throw new Error(
            "No se pudo obtener el correo del usuario actual."
        );

    }


    const registro =
        buscarUsuarioEnLista(
            correo
        );


    const esTI =
        usuarioEsTI(
            correo,
            registro
        );


    if (
        !esTI
    ) {

        grupoUsuarioActual.textContent =
            "SIN ACCESO";


        throw new Error(
            "El usuario actual no pertenece al grupo TI."
        );

    }


    grupoUsuarioActual.textContent =
        "TI - ACCESO TOTAL";


    return true;

}


/* =========================================================
   INICIALIZACIÓN
========================================================= */

async function iniciarPermisos() {

    try {

        cambiarMensajeLoading(
            "Verificando sesión..."
        );


        if (
            !(await esperarCuenta())
        ) {

            cambiarMensajeLoading(
                "Sesión requerida."
            );


            setTimeout(
                solicitarLogin,
                500
            );


            return;

        }


        cambiarMensajeLoading(
            "Obteniendo usuario..."
        );


        await obtenerUsuarioActual();


        cambiarMensajeLoading(
            "Cargando estructura de permisos..."
        );


        await obtenerListaPermisos();


        await obtenerColumnasPermisos();


        cambiarMensajeLoading(
            "Verificando grupo TI..."
        );


        await cargarUsuariosPermisos();


        await verificarAcceso();


        cambiarMensajeLoading(
            "Carga completada."
        );


        ocultarLoading();

    }

    catch (error) {

        console.error(
            "Error inicializando permisos:",
            error
        );


        /*
         * Si no pertenece a TI,
         * no puede utilizar esta sección.
         */

        if (

            error.message &&
            error.message.includes(
                "no pertenece al grupo TI"
            )

        ) {

            cambiarMensajeLoading(
                "Acceso restringido. Solo usuarios TI."
            );


            setTimeout(

                () => {

                    window.location.replace(
                        "index.html"
                    );

                },

                1300

            );


            return;

        }


        cambiarMensajeLoading(

            String(
                error.message ||
                error
            ).slice(
                0,
                250
            )

        );

    }

}


/* =========================================================
   ARRANQUE
========================================================= */

iniciarPermisos();

})();