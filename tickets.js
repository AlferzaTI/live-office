(function () {

/* =========================================================
   ALFERZA LIVE OFFICE
   TICKETS
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
            new URL("blank.html", window.location.href).href

    },

    cache: {

        cacheLocation:
            "sessionStorage",

        storeAuthStateInCookie:
            false

    }

};


/* =========================================================
   MICROSOFT LIST
========================================================= */

const SHAREPOINT_HOST =
    "alferzaholding-my.sharepoint.com";


const SHAREPOINT_SITE_PATH =
    "/personal/soporte1_alferza_pe";


const TICKETS_LIST_NAME =
    "TicketsTI";


/* =========================================================
   SCOPES GRAPH
========================================================= */

const SCOPES_GRAPH = [

    "User.Read",

    "Sites.ReadWrite.All"

];


/* =========================================================
   MSAL
========================================================= */

if (typeof msal === "undefined") {

    document.getElementById("usuarioNombre").textContent =
        "No cargó la librería de Microsoft (MSAL)";

    document.getElementById("usuarioCorreo").textContent =
        "Revisa tu conexión o bloqueadores de contenido";

    throw new Error("MSAL no está disponible.");

}


const msalInstance =
    new msal.PublicClientApplication(
        MSAL_CONFIG
    );


/* =========================================================
   VARIABLES
========================================================= */

let cuentaActual = null;

let sitioSharePoint = null;

let listaTickets = null;

let ticketsData = [];

let ticketSeleccionadoResolver = null;


/* =========================================================
   DOM
========================================================= */

const ticketForm =
    document.getElementById("ticketForm");


const ticketTitulo =
    document.getElementById("ticketTitulo");


const ticketCategoria =
    document.getElementById("ticketCategoria");


const ticketPrioridad =
    document.getElementById("ticketPrioridad");


const ticketDescripcion =
    document.getElementById("ticketDescripcion");


const crearTicketBtn =
    document.getElementById("crearTicketBtn");


const ticketMensaje =
    document.getElementById("ticketMensaje");


const usuarioNombre =
    document.getElementById("usuarioNombre");


const usuarioCorreo =
    document.getElementById("usuarioCorreo");


const ticketsTableBody =
    document.getElementById("ticketsTableBody");


const ticketsEmpty =
    document.getElementById("ticketsEmpty");


const totalTickets =
    document.getElementById("totalTickets");


const ticketsPendientes =
    document.getElementById("ticketsPendientes");


const ticketsProceso =
    document.getElementById("ticketsProceso");


const ticketsResueltos =
    document.getElementById("ticketsResueltos");


const buscarTicket =
    document.getElementById("buscarTicket");


const filtroEstado =
    document.getElementById("filtroEstado");


const filtroPrioridad =
    document.getElementById("filtroPrioridad");


const actualizarTickets =
    document.getElementById("actualizarTickets");


const ticketModal =
    document.getElementById("ticketModal");


const cerrarModal =
    document.getElementById("cerrarModal");


/* =========================================================
   DOM - MODAL RESOLVER
========================================================= */

const resolverModal =
    document.getElementById("resolverModal");


const cerrarResolverModal =
    document.getElementById("cerrarResolverModal");


const cancelarResolver =
    document.getElementById("cancelarResolver");


const confirmarResolver =
    document.getElementById("confirmarResolver");


const resolverTicketId =
    document.getElementById("resolverTicketId");


const resolverEstado =
    document.getElementById("resolverEstado");


const resolverDescripcion =
    document.getElementById("resolverDescripcion");


const resolverMensaje =
    document.getElementById("resolverMensaje");


/* =========================================================
   UTILIDADES
========================================================= */

function escaparHTML(valor) {

    if (
        valor === null ||
        valor === undefined
    ) {

        return "";

    }


    return String(valor)

        .replace(/&/g, "&amp;")

        .replace(/</g, "&lt;")

        .replace(/>/g, "&gt;")

        .replace(/"/g, "&quot;")

        .replace(/'/g, "&#039;");

}


/* =========================================================
   TOKEN
========================================================= */

async function obtenerToken() {

    const cuentas =
        msalInstance.getAllAccounts();


    if (!cuentas.length) {

        const e = new Error(
            "No existe una sesión de Microsoft."
        );

        e.codigo = "SIN_SESION";

        throw e;

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

        /*
         * No abrimos popup automáticamente:
         * el navegador lo bloquea si no viene
         * de un clic del usuario.
         */

        console.warn(
            "Token silencioso no disponible:",
            error
        );

        const e = new Error(
            "Se requiere autorización de Microsoft."
        );

        e.codigo = "REQUIERE_INTERACCION";

        throw e;

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


    return respuesta.json();

}


/* =========================================================
   OBTENER SITIO SHAREPOINT
========================================================= */

async function obtenerSitioSharePoint() {

    if (sitioSharePoint) {

        return sitioSharePoint;

    }


    const url =
        `https://graph.microsoft.com/v1.0/sites/` +
        `${SHAREPOINT_HOST}:${SHAREPOINT_SITE_PATH}`;


    sitioSharePoint =
        await graphFetch(url);


    return sitioSharePoint;

}


/* =========================================================
   OBTENER LISTA
========================================================= */

async function obtenerListaTickets() {

    if (listaTickets) {

        return listaTickets;

    }


    const sitio =
        await obtenerSitioSharePoint();


    const url =
        `https://graph.microsoft.com/v1.0/sites/` +
        `${sitio.id}/lists` +
        `?$select=id,name,displayName&$top=200`;


    const resultado =
        await graphFetch(url);


    if (
        !resultado.value ||
        !resultado.value.length
    ) {

        throw new Error(

            `No se encontró la lista "${TICKETS_LIST_NAME}".`

        );

    }


    const objetivo =
        TICKETS_LIST_NAME.toLowerCase();


    listaTickets =
        resultado.value.find(
            l =>
                String(l.displayName || "")
                    .toLowerCase() === objetivo ||
                String(l.name || "")
                    .toLowerCase() === objetivo
        );


    if (!listaTickets) {

        throw new Error(
            `No se encontró la lista "${TICKETS_LIST_NAME}". ` +
            `Listas disponibles: ` +
            resultado.value
                .map(l => l.displayName)
                .join(", ")
        );

    }


    return listaTickets;

}


/* =========================================================
   OBTENER USUARIO
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


    if (!respuesta.ok) {

        throw new Error(
            `No se pudo obtener el usuario (${respuesta.status}).`
        );

    }


    const usuario =
        await respuesta.json();


    cuentaActual =
        msalInstance.getAllAccounts()[0];


    usuarioNombre.textContent =
        usuario.displayName ||
        cuentaActual?.name ||
        "Usuario";


    usuarioCorreo.textContent =
        usuario.mail ||
        usuario.userPrincipalName ||
        "Sin correo";


    return usuario;

}


/* =========================================================
   CARGAR TICKETS
========================================================= */

async function cargarTickets() {

    try {

        ticketsTableBody.innerHTML = `

            <tr>

                <td
                    colspan="7"
                    class="table-loading"
                >
                    Cargando tickets...
                </td>

            </tr>

        `;


        const sitio =
            await obtenerSitioSharePoint();


        const lista =
            await obtenerListaTickets();


        /*
         * IMPORTANTE:
         * $expand debe llevar el signo $
         * para que Graph devuelva los campos.
         */

        const url =
            `https://graph.microsoft.com/v1.0/sites/` +
            `${sitio.id}/lists/${lista.id}/items` +
            `?$expand=fields&$top=999`;


        const resultado =
            await graphFetch(url);


        ticketsData =
            resultado.value || [];


        console.log(
            "Tickets cargados:",
            ticketsData
        );


        renderizarTickets();

        actualizarResumen();

    }

    catch (error) {

        console.error(
            "Error cargando tickets:",
            error
        );


        ticketsTableBody.innerHTML = `

            <tr>

                <td
                    colspan="7"
                    class="table-loading"
                >
                    No se pudieron cargar los tickets.
                    <br>
                    <small>${escaparHTML(String(error.message).slice(0, 300))}</small>
                </td>

            </tr>

        `;

    }

}


/* =========================================================
   OBTENER CAMPO
========================================================= */

function obtenerCampo(
    fields,
    nombres
) {

    for (
        const nombre of nombres
    ) {

        if (
            fields &&
            fields[nombre] !== undefined &&
            fields[nombre] !== null
        ) {

            return fields[nombre];

        }

    }


    return "";

}


/* =========================================================
   FORMATEAR FECHA
========================================================= */

function formatearFecha(
    fecha
) {

    if (!fecha) {

        return "—";

    }


    const date =
        new Date(fecha);


    if (
        Number.isNaN(date.getTime())
    ) {

        return fecha;

    }


    return date.toLocaleString(
        "es-PE",
        {

            day:
                "2-digit",

            month:
                "2-digit",

            year:
                "numeric",

            hour:
                "2-digit",

            minute:
                "2-digit"

        }
    );

}


/* =========================================================
   CLASE ESTADO
========================================================= */

function claseEstado(
    estado
) {

    const valor =
        String(estado)
            .toLowerCase()
            .trim();


    if (
        valor === "en proceso"
    ) {

        return "estado-proceso";

    }


    if (
        valor === "resuelto"
    ) {

        return "estado-resuelto";

    }


    if (
        valor === "sin resolver"
    ) {

        return "estado-sin-resolver";

    }


    if (
        valor === "cerrado"
    ) {

        return "estado-cerrado";

    }


    return "estado-pendiente";

}


/* =========================================================
   CLASE PRIORIDAD
========================================================= */

function clasePrioridad(
    prioridad
) {

    const valor =
        String(prioridad)
            .toLowerCase()
            .trim();


    if (
        valor === "baja"
    ) {

        return "prioridad-baja";

    }


    if (
        valor === "alta"
    ) {

        return "prioridad-alta";

    }


    if (
        valor === "crítica" ||
        valor === "critica"
    ) {

        return "prioridad-critica";

    }


    return "prioridad-media";

}


/* =========================================================
   RENDERIZAR TICKETS
========================================================= */

function renderizarTickets() {

    const busqueda =
        buscarTicket.value
            .toLowerCase()
            .trim();


    const estadoFiltro =
        filtroEstado.value;


    const prioridadFiltro =
        filtroPrioridad.value;


    const filtrados =
        ticketsData.filter(
            item => {

                const fields =
                    item.fields || {};


                const titulo =
                    obtenerCampo(
                        fields,
                        [
                            "Title",
                            "Título"
                        ]
                    );


                const ticketId =
                    obtenerCampo(
                        fields,
                        [
                            "TicketID"
                        ]
                    );


                const categoria =
                    obtenerCampo(
                        fields,
                        [
                            "Categoria",
                            "Categoría"
                        ]
                    );


                const prioridad =
                    obtenerCampo(
                        fields,
                        [
                            "Prioridad"
                        ]
                    );


                const estado =
                    obtenerCampo(
                        fields,
                        [
                            "Estado"
                        ]
                    );


                const coincideBusqueda =

                    !busqueda ||

                    String(titulo)
                        .toLowerCase()
                        .includes(busqueda) ||

                    String(ticketId)
                        .toLowerCase()
                        .includes(busqueda) ||

                    String(categoria)
                        .toLowerCase()
                        .includes(busqueda);


                const coincideEstado =

                    !estadoFiltro ||

                    estado === estadoFiltro;


                const coincidePrioridad =

                    !prioridadFiltro ||

                    prioridad === prioridadFiltro;


                return (

                    coincideBusqueda &&

                    coincideEstado &&

                    coincidePrioridad

                );

            }
        );


    if (!filtrados.length) {

        ticketsTableBody.innerHTML = "";

        ticketsEmpty.style.display =
            "block";

        return;

    }


    ticketsEmpty.style.display =
        "none";


    ticketsTableBody.innerHTML =
        filtrados.map(
            item => {

                const fields =
                    item.fields || {};


                const titulo =
                    obtenerCampo(
                        fields,
                        [
                            "Title",
                            "Título"
                        ]
                    );


                const ticketId =
                    obtenerCampo(
                        fields,
                        [
                            "TicketID"
                        ]
                    ) ||
                    `TKT-${String(item.id)
                        .padStart(6, "0")}`;


                const categoria =
                    obtenerCampo(
                        fields,
                        [
                            "Categoria",
                            "Categoría"
                        ]
                    );


                const prioridad =
                    obtenerCampo(
                        fields,
                        [
                            "Prioridad"
                        ]
                    );


                const estado =
                    obtenerCampo(
                        fields,
                        [
                            "Estado"
                        ]
                    ) ||
                    "Pendiente";


                const fecha =
                    obtenerCampo(
                        fields,
                        [
                            "FechaCreacion",
                            "FechaCreación"
                        ]
                    );


                /*
                 * Los tickets Resueltos o Cerrados
                 * ya no necesitan acción.
                 *
                 * Los Pendientes, En proceso y
                 * Sin resolver pueden volver a gestionarse.
                 */

                const accion =

                    estado === "Resuelto" ||
                    estado === "Cerrado"

                        ? `
                            <span class="sin-accion">
                                —
                            </span>
                        `

                        : `
                            <button
                                type="button"
                                class="btn-resolver"
                                onclick="abrirResolverTicket(event, '${escaparHTML(item.id)}')"
                            >
                                ✓ Resolver
                            </button>
                        `;


                return `

                    <tr
                        onclick="mostrarDetalleTicket('${escaparHTML(item.id)}')"
                    >

                        <td>

                            <span class="ticket-id">

                                ${escaparHTML(ticketId)}

                            </span>

                        </td>


                        <td>

                            ${escaparHTML(titulo)}

                        </td>


                        <td>

                            ${escaparHTML(categoria)}

                        </td>


                        <td>

                            <span
                                class="prioridad-badge ${clasePrioridad(prioridad)}"
                            >

                                ${escaparHTML(prioridad)}

                            </span>

                        </td>


                        <td>

                            <span
                                class="estado-badge ${claseEstado(estado)}"
                            >

                                ${escaparHTML(estado)}

                            </span>

                        </td>


                        <td>

                            ${formatearFecha(fecha)}

                        </td>


                        <td class="ticket-action-cell">

                            ${accion}

                        </td>

                    </tr>

                `;

            }
        ).join("");

}


/* =========================================================
   RESUMEN
========================================================= */

function actualizarResumen() {

    const total =
        ticketsData.length;


    const pendientes =
        ticketsData.filter(
            item =>
                obtenerCampo(
                    item.fields,
                    ["Estado"]
                ) === "Pendiente"
        ).length;


    const proceso =
        ticketsData.filter(
            item =>
                obtenerCampo(
                    item.fields,
                    ["Estado"]
                ) === "En proceso"
        ).length;


    const resueltos =
        ticketsData.filter(
            item => {

                const estado =
                    obtenerCampo(
                        item.fields,
                        ["Estado"]
                    );


                return (

                    estado === "Resuelto" ||

                    estado === "Cerrado"

                );

            }
        ).length;


    totalTickets.textContent =
        total;


    ticketsPendientes.textContent =
        pendientes;


    ticketsProceso.textContent =
        proceso;


    ticketsResueltos.textContent =
        resueltos;

}


/* =========================================================
   GENERAR ID
========================================================= */

function generarTicketID() {

    const numero =
        Date.now()
            .toString()
            .slice(-6);


    return `TKT-${numero}`;

}


/* =========================================================
   CREAR TICKET
========================================================= */

async function crearTicket(
    event
) {

    event.preventDefault();


    const titulo =
        ticketTitulo.value.trim();


    const categoria =
        ticketCategoria.value;


    const prioridad =
        ticketPrioridad.value;


    const descripcion =
        ticketDescripcion.value.trim();


    if (
        !titulo ||
        !categoria ||
        !prioridad ||
        !descripcion
    ) {

        mostrarMensaje(
            "Completa todos los campos obligatorios.",
            "error"
        );

        return;

    }


    try {

        crearTicketBtn.disabled =
            true;


        crearTicketBtn.innerHTML =
            "⏳ Creando...";


        const usuario =
            await obtenerUsuarioActual();


        const sitio =
            await obtenerSitioSharePoint();


        const lista =
            await obtenerListaTickets();


        const ticketID =
            generarTicketID();


        const fecha =
            new Date()
                .toISOString();


        const body = {

            fields: {

                Title:
                    titulo,

                TicketID:
                    ticketID,

                Usuario:
                    usuario.displayName ||
                    "",

                Correo:
                    usuario.mail ||
                    usuario.userPrincipalName ||
                    "",

                Categoria:
                    categoria,

                Prioridad:
                    prioridad,

                Descripcion:
                    descripcion,

                Estado:
                    "Pendiente",

                FechaCreacion:
                    fecha,

                AsignadoA:
                    "Soporte TI",

                TipoSolucion:
                    ""

            }

        };


        const url =
            `https://graph.microsoft.com/v1.0/sites/` +
            `${sitio.id}/lists/${lista.id}/items`;


        await graphFetch(

            url,

            {

                method:
                    "POST",

                body:
                    JSON.stringify(body)

            }

        );


        mostrarMensaje(

            `Ticket ${ticketID} creado correctamente.`,

            "success"

        );


        ticketForm.reset();


        ticketPrioridad.value =
            "Media";


        await cargarTickets();

    }

    catch (error) {

        console.error(
            "Error creando ticket:",
            error
        );


        mostrarMensaje(

            "No se pudo crear el ticket. " +
            "Revisa la conexión con Microsoft Lists.",

            "error"

        );

    }

    finally {

        crearTicketBtn.disabled =
            false;


        crearTicketBtn.innerHTML =
            "<span>🎫</span> Crear ticket";

    }

}


/* =========================================================
   MENSAJE
========================================================= */

function mostrarMensaje(
    mensaje,
    tipo
) {

    ticketMensaje.textContent =
        mensaje;


    ticketMensaje.className =
        `ticket-message ${tipo}`;


    setTimeout(
        () => {

            ticketMensaje.className =
                "ticket-message";

            ticketMensaje.textContent =
                "";

        },
        5000
    );

}


/* =========================================================
   MOSTRAR DETALLE DEL TICKET
========================================================= */

function mostrarDetalleTicket(
    itemId
) {

    const ticket =
        ticketsData.find(
            item =>
                String(item.id) ===
                String(itemId)
        );


    if (!ticket) {

        return;

    }


    const fields =
        ticket.fields || {};


    const ticketId =
        obtenerCampo(
            fields,
            ["TicketID"]
        ) ||
        `TKT-${String(ticket.id)
            .padStart(6, "0")}`;


    const titulo =
        obtenerCampo(
            fields,
            [
                "Title",
                "Título"
            ]
        );


    const usuario =
        obtenerCampo(
            fields,
            ["Usuario"]
        );


    const categoria =
        obtenerCampo(
            fields,
            [
                "Categoria",
                "Categoría"
            ]
        );


    const prioridad =
        obtenerCampo(
            fields,
            ["Prioridad"]
        );


    const estado =
        obtenerCampo(
            fields,
            ["Estado"]
        ) ||
        "Pendiente";


    const fecha =
        obtenerCampo(
            fields,
            [
                "FechaCreacion",
                "FechaCreación"
            ]
        );


    const descripcion =
        obtenerCampo(
            fields,
            [
                "Descripcion",
                "Descripción"
            ]
        );


    const solucion =
        obtenerCampo(
            fields,
            [
                "TipoSolucion"
            ]
        );


    document.getElementById(
        "modalTicketId"
    ).textContent =
        ticketId;


    const estadoElement =
        document.getElementById(
            "modalTicketEstado"
        );


    estadoElement.textContent =
        estado;


    estadoElement.className =
        `estado-badge ${claseEstado(estado)}`;


    document.getElementById(
        "modalTicketTitulo"
    ).textContent =
        titulo;


    document.getElementById(
        "modalTicketUsuario"
    ).textContent =
        usuario || "—";


    document.getElementById(
        "modalTicketCategoria"
    ).textContent =
        categoria || "—";


    document.getElementById(
        "modalTicketPrioridad"
    ).textContent =
        prioridad || "—";


    document.getElementById(
        "modalTicketFecha"
    ).textContent =
        formatearFecha(fecha);


    document.getElementById(
        "modalTicketDescripcion"
    ).textContent =
        descripcion || "—";


    /*
     * Mostrar el resultado / solución
     * únicamente cuando exista.
     */

    const solucionBox =
        document.getElementById(
            "modalTicketSolucionBox"
        );


    const solucionElement =
        document.getElementById(
            "modalTicketSolucion"
        );


    if (solucion) {

        solucionElement.textContent =
            solucion;


        solucionBox.style.display =
            "block";

    }

    else {

        solucionElement.textContent =
            "—";


        solucionBox.style.display =
            "none";

    }


    ticketModal.classList.add(
        "show"
    );


    document.body.style.overflow =
        "hidden";

}


/* =========================================================
   CERRAR MODAL DETALLE
========================================================= */

function cerrarTicketModal() {

    ticketModal.classList.remove(
        "show"
    );


    document.body.style.overflow =
        "";

}


/* =========================================================
   ABRIR MODAL RESOLVER
========================================================= */

function abrirResolverTicket(
    event,
    itemId
) {

    /*
     * Evita que el click del botón
     * abra también el detalle del ticket.
     */

    event.stopPropagation();


    const ticket =
        ticketsData.find(
            item =>
                String(item.id) ===
                String(itemId)
        );


    if (!ticket) {

        return;

    }


    const fields =
        ticket.fields || {};


    const ticketId =
        obtenerCampo(
            fields,
            ["TicketID"]
        ) ||
        `TKT-${String(ticket.id)
            .padStart(6, "0")}`;


    const estado =
        obtenerCampo(
            fields,
            ["Estado"]
        ) ||
        "Pendiente";


    const solucion =
        obtenerCampo(
            fields,
            ["TipoSolucion"]
        );


    ticketSeleccionadoResolver =
        ticket;


    resolverTicketId.textContent =
        ticketId;


    /*
     * Si el ticket ya estaba Sin resolver,
     * dejamos seleccionada esa opción.
     *
     * En cualquier otro estado, por defecto
     * se propone Resuelto.
     */

    resolverEstado.value =
        estado === "Sin resolver"
            ? "Sin resolver"
            : "Resuelto";


    resolverDescripcion.value =
        solucion || "";


    resolverMensaje.textContent =
        "";


    resolverMensaje.className =
        "ticket-message";


    resolverModal.classList.add(
        "show"
    );


    document.body.style.overflow =
        "hidden";

}


/* =========================================================
   CERRAR MODAL RESOLVER
========================================================= */

function cerrarResolverTicket() {

    resolverModal.classList.remove(
        "show"
    );


    document.body.style.overflow =
        "";


    ticketSeleccionadoResolver =
        null;


    resolverDescripcion.value =
        "";


    resolverMensaje.textContent =
        "";


    resolverMensaje.className =
        "ticket-message";

}


/* =========================================================
   MENSAJE MODAL RESOLVER
========================================================= */

function mostrarResolverMensaje(
    mensaje,
    tipo
) {

    resolverMensaje.textContent =
        mensaje;


    resolverMensaje.className =
        `ticket-message ${tipo}`;

}


/* =========================================================
   GUARDAR RESULTADO DEL TICKET
========================================================= */

async function guardarResultadoTicket() {

    if (!ticketSeleccionadoResolver) {

        return;

    }


    const estado =
        resolverEstado.value;


    const descripcion =
        resolverDescripcion.value.trim();


    if (!estado) {

        mostrarResolverMensaje(

            "Selecciona el resultado del ticket.",

            "error"

        );

        return;

    }


    if (!descripcion) {

        mostrarResolverMensaje(

            "Ingresa una descripción del resultado.",

            "error"

        );

        resolverDescripcion.focus();

        return;

    }


    try {

        confirmarResolver.disabled =
            true;


        confirmarResolver.innerHTML =
            "⏳ Guardando...";


        const sitio =
            await obtenerSitioSharePoint();


        const lista =
            await obtenerListaTickets();


        const itemId =
            ticketSeleccionadoResolver.id;


        /*
         * Actualizamos directamente los campos
         * del elemento de Microsoft Lists.
         */

        const url =
            `https://graph.microsoft.com/v1.0/sites/` +
            `${sitio.id}/lists/${lista.id}/items/` +
            `${itemId}/fields`;


        const body = {

            Estado:
                estado,

            TipoSolucion:
                descripcion

        };


        await graphFetch(

            url,

            {

                method:
                    "PATCH",

                body:
                    JSON.stringify(body)

            }

        );


        /*
         * Actualizar también los datos locales
         * para que la interfaz responda inmediatamente.
         */

        if (
            ticketSeleccionadoResolver.fields
        ) {

            ticketSeleccionadoResolver.fields.Estado =
                estado;


            ticketSeleccionadoResolver.fields.TipoSolucion =
                descripcion;

        }


        mostrarResolverMensaje(

            "Resultado guardado correctamente.",

            "success"

        );


        /*
         * Esperamos un momento para que el usuario
         * vea el mensaje antes de cerrar.
         */

        setTimeout(
            async () => {

                cerrarResolverTicket();

                renderizarTickets();

                actualizarResumen();

                await cargarTickets();

            },
            700
        );

    }

    catch (error) {

        console.error(
            "Error guardando resultado del ticket:",
            error
        );


        mostrarResolverMensaje(

            "No se pudo guardar el resultado. " +
            "Verifica que la columna TipoSolucion " +
            "exista en Microsoft Lists y que tenga permisos de edición.",

            "error"

        );

    }

    finally {

        confirmarResolver.disabled =
            false;


        confirmarResolver.innerHTML =
            "✓ Guardar resultado";

    }

}


/* =========================================================
   EVENTOS
========================================================= */

ticketForm.addEventListener(
    "submit",
    crearTicket
);


actualizarTickets.addEventListener(
    "click",
    cargarTickets
);


buscarTicket.addEventListener(
    "input",
    renderizarTickets
);


filtroEstado.addEventListener(
    "change",
    renderizarTickets
);


filtroPrioridad.addEventListener(
    "change",
    renderizarTickets
);


cerrarModal.addEventListener(
    "click",
    cerrarTicketModal
);


/* =========================================================
   EVENTOS MODAL DETALLE
========================================================= */

ticketModal.addEventListener(
    "click",
    event => {

        if (
            event.target === ticketModal
        ) {

            cerrarTicketModal();

        }

    }
);


/* =========================================================
   EVENTOS MODAL RESOLVER
========================================================= */

cerrarResolverModal.addEventListener(
    "click",
    cerrarResolverTicket
);


cancelarResolver.addEventListener(
    "click",
    cerrarResolverTicket
);


confirmarResolver.addEventListener(
    "click",
    guardarResultadoTicket
);


resolverModal.addEventListener(
    "click",
    event => {

        if (
            event.target === resolverModal
        ) {

            cerrarResolverTicket();

        }

    }
);


/* =========================================================
   TECLA ESCAPE
========================================================= */

document.addEventListener(
    "keydown",
    event => {

        if (
            event.key !== "Escape"
        ) {

            return;

        }


        if (
            resolverModal.classList.contains("show")
        ) {

            cerrarResolverTicket();

            return;

        }


        if (
            ticketModal.classList.contains("show")
        ) {

            cerrarTicketModal();

        }

    }
);


/* =========================================================
   INICIALIZACIÓN
========================================================= */

async function esperarCuenta(
    ms = 6000
) {

    const fin =
        Date.now() + ms;


    while (Date.now() < fin) {

        if (
            msalInstance.getAllAccounts().length
        ) {

            return true;

        }

        await new Promise(
            r => setTimeout(r, 400)
        );

    }

    return false;

}


function mostrarBotonLogin(
    texto
) {

    usuarioNombre.textContent =
        "Sesión de Microsoft requerida";

    usuarioCorreo.textContent =
        texto;

    ticketsTableBody.innerHTML = `
        <tr>
            <td colspan="7" class="table-loading">
                <button
                    type="button"
                    class="btn-primary"
                    id="btnLoginTickets"
                >
                    Iniciar sesión / autorizar con Microsoft
                </button>
            </td>
        </tr>
    `;

    document
        .getElementById("btnLoginTickets")
        .addEventListener(
            "click",
            async () => {

                try {

                    const cuentas =
                        msalInstance.getAllAccounts();

                    if (cuentas.length) {

                        await msalInstance.acquireTokenPopup({
                            scopes: SCOPES_GRAPH,
                            account: cuentas[0]
                        });

                    }

                    else {

                        await msalInstance.loginPopup({
                            scopes: SCOPES_GRAPH
                        });

                    }

                    iniciarTickets();

                }

                catch (error) {

                    console.error(
                        "Error de inicio de sesión:",
                        error
                    );

                    usuarioCorreo.textContent =
                        "No se pudo iniciar sesión: " +
                        String(error.message || error)
                            .slice(0, 200);

                }

            }
        );

}


async function iniciarTickets() {

    try {

        /*
         * auth.js se carga después de este archivo,
         * así que damos unos segundos para que
         * termine de restaurar la sesión.
         */

        if (!(await esperarCuenta())) {

            mostrarBotonLogin(
                "Inicia sesión para ver y crear tickets"
            );

            return;

        }


        await obtenerUsuarioActual();

        await cargarTickets();

    }

    catch (error) {

        console.error(
            "Error inicializando Tickets:",
            error
        );


        if (
            error.codigo === "REQUIERE_INTERACCION" ||
            error.codigo === "SIN_SESION"
        ) {

            mostrarBotonLogin(
                "Autoriza los permisos de Microsoft"
            );

            return;

        }


        usuarioNombre.textContent =
            "No se pudo obtener el usuario";

        usuarioCorreo.textContent =
            String(error.message || error)
                .slice(0, 200);

    }

}


/* Funciones usadas desde onclick en el HTML */

window.abrirResolverTicket =
    abrirResolverTicket;

window.mostrarDetalleTicket =
    mostrarDetalleTicket;


iniciarTickets();

})();