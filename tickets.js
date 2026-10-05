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
            "https://alferzati.github.io/live-office/blank.html"

    },

    cache: {

        cacheLocation: "sessionStorage",

        storeAuthStateInCookie: false

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
SCOPES
========================================================= */

const SCOPES_GRAPH = [

    "User.Read",

    "Sites.Read.All",

    "Sites.ReadWrite.All"

];

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

let cuentaActual = null;

let sitioSharePoint = null;

let listaTickets = null;

let ticketsData = [];

let ticketSeleccionadoResolver = null;

let campoTipoSolucion = null;

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
MODAL RESOLVER
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
ESCAPAR HTML
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
NORMALIZAR TEXTO
========================================================= */

function normalizarTexto(valor) {

    return String(valor || "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .replace(/\s+/g, "")
        .trim();

}

/* =========================================================
TOKEN
========================================================= */

async function obtenerToken() {

    const cuentas =
        msalInstance.getAllAccounts();

    if (!cuentas.length) {

        throw new Error(
            "No existe una sesión de Microsoft."
        );

    }

    cuentaActual =
        cuentas[0];

    try {

        const resultado =
            await msalInstance.acquireTokenSilent({

                scopes: SCOPES_GRAPH,

                account: cuentaActual

            });

        return resultado.accessToken;

    }

    catch (error) {

        const resultado =
            await msalInstance.acquireTokenPopup({

                scopes: SCOPES_GRAPH

            });

        return resultado.accessToken;

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
        `?$filter=displayName eq '${TICKETS_LIST_NAME}'`;

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

    listaTickets =
        resultado.value[0];

    return listaTickets;

}

/* =========================================================
OBTENER COLUMNA TIPO SOLUCIÓN
========================================================= */

async function obtenerCampoTipoSolucion() {

    if (campoTipoSolucion) {

        return campoTipoSolucion;

    }

    const sitio =
        await obtenerSitioSharePoint();

    const lista =
        await obtenerListaTickets();

    const url =
        `https://graph.microsoft.com/v1.0/sites/` +
        `${sitio.id}/lists/${lista.id}/columns` +
        `?$select=id,name,displayName`;

    const resultado =
        await graphFetch(url);

    const columnas =
        resultado.value || [];

    const columna =
        columnas.find(
            columna => {

                const nombre =
                    normalizarTexto(
                        columna.name
                    );

                const displayName =
                    normalizarTexto(
                        columna.displayName
                    );

                return (

                    nombre === "tiposolucion" ||

                    nombre === "tiposolución" ||

                    displayName === "tiposolucion" ||

                    displayName === "tiposolucion"

                );

            }
        );

    if (!columna) {

        throw new Error(
            "No se encontró la columna 'Tipo solución' en la lista TicketsTI."
        );

    }

    campoTipoSolucion =
        columna.name;

    console.log(
        "Columna Tipo solución detectada:",
        campoTipoSolucion
    );

    return campoTipoSolucion;

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
            "No se pudo obtener el usuario."
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

        ticketsEmpty.style.display =
            "none";

        const sitio =
            await obtenerSitioSharePoint();

        const lista =
            await obtenerListaTickets();

        const url =
            `https://graph.microsoft.com/v1.0/sites/` +
            `${sitio.id}/lists/${lista.id}/items` +
            `?expand=fields&$top=999`;

        const resultado =
            await graphFetch(url);

        ticketsData =
            resultado.value || [];

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

function formatearFecha(fecha) {

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

            day: "2-digit",

            month: "2-digit",

            year: "numeric",

            hour: "2-digit",

            minute: "2-digit"

        }
    );

}

/* =========================================================
CLASE ESTADO
========================================================= */

function claseEstado(estado) {

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

function clasePrioridad(prioridad) {

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

                const puedeResolver =

                    estado !== "Resuelto" &&

                    estado !== "Cerrado";

                return `

                    <tr
                        onclick="mostrarDetalleTicket('${item.id}')"
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

                        <td>

                            ${
                                puedeResolver

                                    ?

                                `
                                <button
                                    type="button"
                                    class="btn-resolver"
                                    onclick="
                                        event.stopPropagation();
                                        abrirResolverTicket('${item.id}');
                                    "
                                >
                                    ✓ Resolver
                                </button>
                                `

                                    :

                                `
                                <span class="sin-accion">
                                    —
                                </span>
                                `
                            }

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

async function crearTicket(event) {

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
                    "Soporte TI"

            }

        };

        const url =
            `https://graph.microsoft.com/v1.0/sites/` +
            `${sitio.id}/lists/${lista.id}/items`;

        await graphFetch(

            url,

            {

                method: "POST",

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
MENSAJE GENERAL
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
ABRIR MODAL RESOLVER
========================================================= */

function abrirResolverTicket(itemId) {

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

    const estado =
        obtenerCampo(
            fields,
            ["Estado"]
        );

    if (
        estado === "Resuelto" ||
        estado === "Cerrado"
    ) {

        return;

    }

    ticketSeleccionadoResolver =
        ticket;

    const ticketId =
        obtenerCampo(
            fields,
            ["TicketID"]
        ) ||
        `TKT-${String(ticket.id)
            .padStart(6, "0")}`;

    resolverTicketId.textContent =
        ticketId;

    resolverEstado.value =
        "Resuelto";

    resolverDescripcion.value =
        "";

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

    ticketSeleccionadoResolver =
        null;

    if (
        !ticketModal.classList.contains("show")
    ) {

        document.body.style.overflow =
            "";

    }

}

/* =========================================================
MENSAJE MODAL RESOLVER
========================================================= */

function mostrarMensajeResolver(
    mensaje,
    tipo
) {

    resolverMensaje.textContent =
        mensaje;

    resolverMensaje.className =
        `ticket-message ${tipo}`;

}

/* =========================================================
GUARDAR RESULTADO
========================================================= */

async function guardarResultadoTicket() {

    if (
        !ticketSeleccionadoResolver
    ) {

        return;

    }

    const estado =
        resolverEstado.value;

    const descripcion =
        resolverDescripcion.value.trim();

    if (!descripcion) {

        mostrarMensajeResolver(

            "Escribe una descripción antes de guardar.",

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

        const nombreCampoSolucion =
            await obtenerCampoTipoSolucion();

        const url =
            `https://graph.microsoft.com/v1.0/sites/` +
            `${sitio.id}/lists/${lista.id}/items/` +
            `${itemId}/fields`;

        const body = {

            Estado:
                estado

        };

        body[nombreCampoSolucion] =
            descripcion;

        console.log(
            "Actualizando ticket:",
            body
        );

        await graphFetch(

            url,

            {

                method: "PATCH",

                body:
                    JSON.stringify(body)

            }

        );

        const ticketId =
            obtenerCampo(
                ticketSeleccionadoResolver.fields,
                ["TicketID"]
            ) ||
            `TKT-${String(itemId)
                .padStart(6, "0")}`;

        cerrarResolverTicket();

        mostrarMensaje(

            `Ticket ${ticketId} actualizado: ${estado}.`,

            "success"

        );

        await cargarTickets();

    }

    catch (error) {

        console.error(
            "Error actualizando ticket:",
            error
        );

        mostrarMensajeResolver(

            error.message ||
            "No se pudo actualizar el ticket.",

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
MOSTRAR DETALLE
========================================================= */

function mostrarDetalleTicket(itemId) {

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

    const nombresSolucion = [
        "TipoSolucion",
        "Tipo solución",
        "Tipo Solucion",
        "Tipo_x0020_soluci_x00f3_n"
    ];

    if (campoTipoSolucion) {

        nombresSolucion.push(
            campoTipoSolucion
        );

    }

    const solucion =
        obtenerCampo(
            fields,
            nombresSolucion
        );

    const descripcion =
        obtenerCampo(
            fields,
            [
                "Descripcion",
                "Descripción"
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
CERRAR DETALLE
========================================================= */

function cerrarTicketModal() {

    ticketModal.classList.remove(
        "show"
    );

    if (
        !resolverModal.classList.contains("show")
    ) {

        document.body.style.overflow =
            "";

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
ESCAPE
========================================================= */

document.addEventListener(
    "keydown",
    event => {

        if (
            event.key === "Escape"
        ) {

            cerrarTicketModal();

            cerrarResolverTicket();

        }

    }
);

/* =========================================================
INICIALIZACIÓN
========================================================= */

async function iniciarTickets() {

    try {

        await obtenerUsuarioActual();

        await cargarTickets();

    }

    catch (error) {

        console.error(
            "Error inicializando Tickets:",
            error
        );

        usuarioNombre.textContent =
            "No se pudo obtener el usuario";

        usuarioCorreo.textContent =
            "Revisa tu sesión de Microsoft 365";

    }

}

iniciarTickets();