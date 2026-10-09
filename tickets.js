
(function () {
    "use strict";

    /* =========================================================
       ALFERZA LIVE OFFICE
       MÓDULO DE TICKETS
    ========================================================= */

    /* =========================================================
       CONFIGURACIÓN
    ========================================================= */

    const MSAL_CONFIG = {
        auth: {
            clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",
            authority: "https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1",
            redirectUri: new URL("blank.html", window.location.href).href
        },
        cache: {
            cacheLocation: "sessionStorage",
            storeAuthStateInCookie: false
        }
    };

    const SHAREPOINT_HOST = "alferzaholding-my.sharepoint.com";
    const SHAREPOINT_SITE_PATH = "/personal/soporte1_alferza_pe";
    const TICKETS_LIST_NAME = "TicketsTI";
    const SCOPES_GRAPH = ["User.Read", "Sites.ReadWrite.All"];

    /* =========================================================
       VALIDACIÓN DE MSAL
    ========================================================= */

    if (typeof msal === "undefined") {
        document.getElementById("usuarioNombre").textContent =
            "No cargó la librería de Microsoft (MSAL)";

        document.getElementById("usuarioCorreo").textContent =
            "Revisa tu conexión o bloqueadores de contenido";

        throw new Error("MSAL no está disponible.");
    }

    const msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);

    /* =========================================================
       ESTADO DEL MÓDULO
    ========================================================= */

    let cuentaActual = null;
    let sitioSharePoint = null;
    let listaTickets = null;
    let ticketsData = [];
    let ticketSeleccionadoResolver = null;
    let nombreCampoSolucion = null;

    /* =========================================================
       REFERENCIAS HTML
    ========================================================= */

    const ticketForm = document.getElementById("ticketForm");
    const ticketTitulo = document.getElementById("ticketTitulo");
    const ticketCategoria = document.getElementById("ticketCategoria");
    const ticketPrioridad = document.getElementById("ticketPrioridad");
    const ticketDescripcion = document.getElementById("ticketDescripcion");
    const crearTicketBtn = document.getElementById("crearTicketBtn");
    const ticketMensaje = document.getElementById("ticketMensaje");

    const usuarioNombre = document.getElementById("usuarioNombre");
    const usuarioCorreo = document.getElementById("usuarioCorreo");

    const registrosSection = document.getElementById("registrosSection");
    const ticketsTableBody = document.getElementById("ticketsTableBody");
    const ticketsEmpty = document.getElementById("ticketsEmpty");

    const totalTickets = document.getElementById("totalTickets");
    const ticketsPendientes = document.getElementById("ticketsPendientes");
    const ticketsProceso = document.getElementById("ticketsProceso");
    const ticketsResueltos = document.getElementById("ticketsResueltos");

    const buscarTicket = document.getElementById("buscarTicket");
    const filtroEstado = document.getElementById("filtroEstado");
    const filtroPrioridad = document.getElementById("filtroPrioridad");
    const actualizarTickets = document.getElementById("actualizarTickets");
    const abrirReporte = document.getElementById("abrirReporte");

    const ticketModal = document.getElementById("ticketModal");
    const cerrarModal = document.getElementById("cerrarModal");

    const resolverModal = document.getElementById("resolverModal");
    const cerrarResolverModal = document.getElementById("cerrarResolverModal");
    const cancelarResolver = document.getElementById("cancelarResolver");
    const confirmarResolver = document.getElementById("confirmarResolver");
    const resolverTicketId = document.getElementById("resolverTicketId");
    const resolverEstado = document.getElementById("resolverEstado");
    const resolverDescripcion = document.getElementById("resolverDescripcion");
    const resolverMensaje = document.getElementById("resolverMensaje");

    const reporteModal = document.getElementById("reporteModal");
    const cerrarReporteModal = document.getElementById("cerrarReporteModal");
    const cancelarReporte = document.getElementById("cancelarReporte");
    const generarReportePDFBtn = document.getElementById("generarReportePDF");
    const reporteMes = document.getElementById("reporteMes");
    const reporteCantidad = document.getElementById("reporteCantidad");
    const reportePeriodo = document.getElementById("reportePeriodo");
    const reporteMensaje = document.getElementById("reporteMensaje");

    /* =========================================================
       PERMISOS DE REGISTROS Y REPORTES
       TI: acceso a registros y reportes.
       Normal: puede crear tickets, pero no ver los registros.
    ========================================================= */

    // Denegar visualización mientras se validan los permisos.
    if (registrosSection) {
        registrosSection.style.display = "none";
    }

    if (abrirReporte) {
        abrirReporte.hidden = true;
        abrirReporte.style.display = "none";
    }

    function esUsuarioTI() {
        const grupo = String(
            window.alferzaPermisos?.grupo || ""
        ).trim().toUpperCase();

        return grupo === "TI";
    }

    function aplicarPermisosReportes() {
        const autorizado = esUsuarioTI();

        // Controlar la sección completa: título, filtros, tabla y botones.
        if (registrosSection) {
            registrosSection.style.display = autorizado ? "" : "none";
        }

        // Controlar también el botón de reportes.
        if (abrirReporte) {
            abrirReporte.hidden = !autorizado;
            abrirReporte.style.display = autorizado ? "" : "none";
        }

        // Cerrar el reporte si el usuario no tiene permisos.
        if (!autorizado && reporteModal) {
            cerrarReporte();
        }

        console.log(
            "[Tickets] Grupo:",
            window.alferzaPermisos?.grupo || "SIN VALIDAR",
            "| Registros permitidos:",
            autorizado,
            "| Reportes permitidos:",
            autorizado
        );

        return autorizado;
    }

    function inicializarPermisosReportes() {
        // Capturar el evento de validación de acceso.
        document.addEventListener(
            "alferza:access-granted",
            function (event) {
                if (event.detail) {
                    window.alferzaPermisos = event.detail;
                }

                aplicarPermisosReportes();
            }
        );

        // Cubrir el caso en que la validación ya esté en curso.
        if (window.alferzaAccessReady) {
            window.alferzaAccessReady
                .then(function (resultado) {
                    if (
                        resultado &&
                        resultado.autorizado === true &&
                        resultado.permisos
                    ) {
                        window.alferzaPermisos = resultado.permisos;
                    } else if (!esUsuarioTI()) {
                        window.alferzaPermisos = {
                            grupo: "Normal",
                            modulos: []
                        };
                    }

                    aplicarPermisosReportes();
                })
                .catch(function (error) {
                    console.error(
                        "[Tickets] Error validando permisos:",
                        error
                    );

                    window.alferzaPermisos = {
                        grupo: "Normal",
                        modulos: []
                    };

                    aplicarPermisosReportes();
                });

            return;
        }

        // Sin permisos confirmados, ocultar la sección.
        aplicarPermisosReportes();
    }

    /* =========================================================
       UTILIDADES
    ========================================================= */

    function escaparHTML(valor) {
        if (valor === null || valor === undefined) {
            return "";
        }

        return String(valor)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    function normalizarNombreColumna(valor) {
        return String(valor || "")
            .toLowerCase()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .replace(/[\s_-]/g, "")
            .trim();
    }

    function obtenerCampo(fields, nombres) {
        for (const nombre of nombres) {
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

    function formatearFecha(fecha) {
        if (!fecha) {
            return "—";
        }

        const date = new Date(fecha);

        if (Number.isNaN(date.getTime())) {
            return fecha;
        }

        return date.toLocaleString("es-PE", {
            day: "2-digit",
            month: "2-digit",
            year: "numeric",
            hour: "2-digit",
            minute: "2-digit"
        });
    }

    function obtenerFechaTicket(item) {
        const fields = item.fields || {};

        return obtenerCampo(fields, [
            "FechaCreacion",
            "FechaCreación"
        ]) || item.createdDateTime || "";
    }

    function claseEstado(estado) {
        const valor = String(estado).toLowerCase().trim();

        if (valor === "en proceso") return "estado-proceso";
        if (valor === "resuelto") return "estado-resuelto";
        if (valor === "sin resolver") return "estado-sin-resolver";
        if (valor === "cerrado") return "estado-cerrado";

        return "estado-pendiente";
    }

    function clasePrioridad(prioridad) {
        const valor = String(prioridad).toLowerCase().trim();

        if (valor === "baja") return "prioridad-baja";
        if (valor === "alta") return "prioridad-alta";

        if (valor === "crítica" || valor === "critica") {
            return "prioridad-critica";
        }

        return "prioridad-media";
    }

    /* =========================================================
       MICROSOFT GRAPH Y SHAREPOINT
    ========================================================= */

    async function obtenerToken() {
        const cuentas = msalInstance.getAllAccounts();

        if (!cuentas.length) {
            const error = new Error("No existe una sesión de Microsoft.");
            error.codigo = "SIN_SESION";
            throw error;
        }

        cuentaActual = cuentas[0];

        try {
            const resultado = await msalInstance.acquireTokenSilent({
                scopes: SCOPES_GRAPH,
                account: cuentaActual
            });

            return resultado.accessToken;
        } catch (error) {
            console.warn("Token silencioso no disponible:", error);

            const nuevoError = new Error(
                "Se requiere autorización de Microsoft."
            );

            nuevoError.codigo = "REQUIERE_INTERACCION";
            throw nuevoError;
        }
    }

    async function graphFetch(url, opciones = {}) {
        const token = await obtenerToken();

        const headers = {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
            ...(opciones.headers || {})
        };

        const respuesta = await fetch(url, {
            ...opciones,
            headers
        });

        if (!respuesta.ok) {
            const texto = await respuesta.text();

            throw new Error(
                `Graph ${respuesta.status}: ${texto}`
            );
        }

        if (respuesta.status === 204) {
            return null;
        }

        const texto = await respuesta.text();

        if (!texto) {
            return null;
        }

        try {
            return JSON.parse(texto);
        } catch {
            return texto;
        }
    }

    async function obtenerSitioSharePoint() {
        if (sitioSharePoint) {
            return sitioSharePoint;
        }

        const url =
            `https://graph.microsoft.com/v1.0/sites/` +
            `${SHAREPOINT_HOST}:${SHAREPOINT_SITE_PATH}`;

        sitioSharePoint = await graphFetch(url);

        return sitioSharePoint;
    }

    async function obtenerListaTickets() {
        if (listaTickets) {
            return listaTickets;
        }

        const sitio = await obtenerSitioSharePoint();

        const url =
            `https://graph.microsoft.com/v1.0/sites/` +
            `${sitio.id}/lists?$select=id,name,displayName&$top=200`;

        const resultado = await graphFetch(url);

        if (!resultado.value || !resultado.value.length) {
            throw new Error(
                `No se encontró la lista "${TICKETS_LIST_NAME}".`
            );
        }

        const objetivo = TICKETS_LIST_NAME.toLowerCase();

        listaTickets = resultado.value.find(function (lista) {
            return (
                String(lista.displayName || "").toLowerCase() === objetivo ||
                String(lista.name || "").toLowerCase() === objetivo
            );
        });

        if (!listaTickets) {
            const disponibles = resultado.value
                .map(lista => lista.displayName || lista.name)
                .join(", ");

            throw new Error(
                `No se encontró la lista "${TICKETS_LIST_NAME}". ` +
                `Listas disponibles: ${disponibles}`
            );
        }

        return listaTickets;
    }

    async function obtenerNombreCampoSolucion() {
        if (nombreCampoSolucion) {
            return nombreCampoSolucion;
        }

        const sitio = await obtenerSitioSharePoint();
        const lista = await obtenerListaTickets();

        const url =
            `https://graph.microsoft.com/v1.0/sites/` +
            `${sitio.id}/lists/${lista.id}/columns` +
            `?$select=name,displayName,hidden,readOnly&$top=200`;

        const resultado = await graphFetch(url);
        const columnas = resultado.value || [];

        console.log("Columnas reales de TicketsTI:", columnas);

        let columna = columnas.find(function (item) {
            return normalizarNombreColumna(item.displayName) ===
                "tiposolucion";
        });

        if (!columna) {
            columna = columnas.find(function (item) {
                const nombre = normalizarNombreColumna(item.name);

                return (
                    nombre === "tiposolucion" ||
                    nombre === "confirmarresolve"
                );
            });
        }

        if (!columna) {
            throw new Error(
                "No se encontró la columna de solución en Microsoft Lists. " +
                "Revisa que exista una columna cuyo nombre visible sea TipoSolucion."
            );
        }

        nombreCampoSolucion = columna.name;

        console.log("Campo de solución detectado:", {
            displayName: columna.displayName,
            name: columna.name
        });

        return nombreCampoSolucion;
    }

    function obtenerSolucion(fields) {
        if (!fields) {
            return "";
        }

        if (
            nombreCampoSolucion &&
            fields[nombreCampoSolucion] !== undefined &&
            fields[nombreCampoSolucion] !== null
        ) {
            return fields[nombreCampoSolucion];
        }

        const posiblesNombres = [
            "TipoSolucion",
            "ConfirmarResolve"
        ];

        for (const nombre of posiblesNombres) {
            if (
                fields[nombre] !== undefined &&
                fields[nombre] !== null
            ) {
                return fields[nombre];
            }
        }

        const claveEncontrada = Object.keys(fields).find(function (clave) {
            const normalizada = normalizarNombreColumna(clave);

            return (
                normalizada === "tiposolucion" ||
                normalizada === "confirmarresolve"
            );
        });

        return claveEncontrada ? fields[claveEncontrada] : "";
    }

    async function obtenerUsuarioActual() {
        const token = await obtenerToken();

        const respuesta = await fetch(
            "https://graph.microsoft.com/v1.0/me",
            {
                headers: {
                    Authorization: `Bearer ${token}`
                }
            }
        );

        if (!respuesta.ok) {
            throw new Error(
                `No se pudo obtener el usuario (${respuesta.status}).`
            );
        }

        const usuario = await respuesta.json();

        cuentaActual = msalInstance.getAllAccounts()[0];

        usuarioNombre.textContent =
            usuario.displayName || cuentaActual?.name || "Usuario";

        usuarioCorreo.textContent =
            usuario.mail ||
            usuario.userPrincipalName ||
            "Sin correo";

        return usuario;
    }

    /* =========================================================
       CARGA Y VISUALIZACIÓN DE TICKETS
    ========================================================= */

    async function cargarTickets() {
        // Evitar consultas desde los controles de registros para usuarios normales.
        if (!esUsuarioTI()) {
            aplicarPermisosReportes();
            return;
        }

        try {
            ticketsTableBody.innerHTML = `
                <tr>
                    <td colspan="7" class="table-loading">
                        Cargando tickets...
                    </td>
                </tr>
            `;

            ticketsEmpty.style.display = "none";

            const sitio = await obtenerSitioSharePoint();
            const lista = await obtenerListaTickets();

            try {
                await obtenerNombreCampoSolucion();
            } catch (error) {
                console.warn(
                    "No se pudo detectar inicialmente la columna de solución:",
                    error
                );
            }

            const url =
                `https://graph.microsoft.com/v1.0/sites/` +
                `${sitio.id}/lists/${lista.id}/items` +
                `?$expand=fields&$top=999`;

            const resultado = await graphFetch(url);

            ticketsData = resultado.value || [];

            console.log("Tickets cargados:", ticketsData);

            renderizarTickets();
            actualizarResumen();
            actualizarCantidadReporte();
        } catch (error) {
            console.error("Error cargando tickets:", error);

            ticketsTableBody.innerHTML = `
                <tr>
                    <td colspan="7" class="table-loading">
                        No se pudieron cargar los tickets.<br>
                        <small>
                            ${escaparHTML(String(error.message).slice(0, 500))}
                        </small>
                    </td>
                </tr>
            `;
        }
    }

    function renderizarTickets() {
        if (!esUsuarioTI()) {
            return;
        }

        const busqueda = buscarTicket.value.toLowerCase().trim();
        const estadoFiltro = filtroEstado.value;
        const prioridadFiltro = filtroPrioridad.value;

        const filtrados = ticketsData.filter(function (item) {
            const fields = item.fields || {};

            const titulo = obtenerCampo(fields, ["Title", "Título"]);
            const ticketId = obtenerCampo(fields, ["TicketID"]);
            const categoria = obtenerCampo(fields, ["Categoria", "Categoría"]);
            const prioridad = obtenerCampo(fields, ["Prioridad"]);
            const estado = obtenerCampo(fields, ["Estado"]);

            const coincideBusqueda =
                !busqueda ||
                String(titulo).toLowerCase().includes(busqueda) ||
                String(ticketId).toLowerCase().includes(busqueda) ||
                String(categoria).toLowerCase().includes(busqueda);

            const coincideEstado =
                !estadoFiltro || estado === estadoFiltro;

            const coincidePrioridad =
                !prioridadFiltro || prioridad === prioridadFiltro;

            return (
                coincideBusqueda &&
                coincideEstado &&
                coincidePrioridad
            );
        });

        if (!filtrados.length) {
            ticketsTableBody.innerHTML = "";
            ticketsEmpty.style.display = "block";
            return;
        }

        ticketsEmpty.style.display = "none";

        ticketsTableBody.innerHTML = filtrados.map(function (item) {
            const fields = item.fields || {};

            const titulo = obtenerCampo(fields, ["Title", "Título"]);

            const ticketId = obtenerCampo(fields, ["TicketID"]) ||
                `TKT-${String(item.id).padStart(6, "0")}`;

            const categoria = obtenerCampo(fields, ["Categoria", "Categoría"]);
            const prioridad = obtenerCampo(fields, ["Prioridad"]);
            const estado = obtenerCampo(fields, ["Estado"]) || "Pendiente";
            const fecha = obtenerCampo(fields, ["FechaCreacion", "FechaCreación"]);

            const accion =
                estado === "Resuelto" || estado === "Cerrado"
                    ? `<span class="sin-accion">—</span>`
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
                <tr onclick="mostrarDetalleTicket('${escaparHTML(item.id)}')">
                    <td>
                        <span class="ticket-id">
                            ${escaparHTML(ticketId)}
                        </span>
                    </td>
                    <td>${escaparHTML(titulo)}</td>
                    <td>${escaparHTML(categoria)}</td>
                    <td>
                        <span class="prioridad-badge ${clasePrioridad(prioridad)}">
                            ${escaparHTML(prioridad)}
                        </span>
                    </td>
                    <td>
                        <span class="estado-badge ${claseEstado(estado)}">
                            ${escaparHTML(estado)}
                        </span>
                    </td>
                    <td>${formatearFecha(fecha)}</td>
                    <td class="ticket-action-cell">${accion}</td>
                </tr>
            `;
        }).join("");
    }

    function actualizarResumen() {
        const total = ticketsData.length;

        const pendientes = ticketsData.filter(function (item) {
            return obtenerCampo(item.fields, ["Estado"]) === "Pendiente";
        }).length;

        const proceso = ticketsData.filter(function (item) {
            return obtenerCampo(item.fields, ["Estado"]) === "En proceso";
        }).length;

        const resueltos = ticketsData.filter(function (item) {
            const estado = obtenerCampo(item.fields, ["Estado"]);

            return estado === "Resuelto" || estado === "Cerrado";
        }).length;

        totalTickets.textContent = total;
        ticketsPendientes.textContent = pendientes;
        ticketsProceso.textContent = proceso;
        ticketsResueltos.textContent = resueltos;
    }

    /* =========================================================
       CREACIÓN DE TICKETS
       DISPONIBLE PARA TODOS LOS USUARIOS AUTENTICADOS
    ========================================================= */

    function generarTicketID() {
        const numero = Date.now().toString().slice(-6);

        return `TKT-${numero}`;
    }

    async function crearTicket(event) {
        event.preventDefault();

        const titulo = ticketTitulo.value.trim();
        const categoria = ticketCategoria.value;
        const prioridad = ticketPrioridad.value;
        const descripcion = ticketDescripcion.value.trim();

        if (!titulo || !categoria || !prioridad || !descripcion) {
            mostrarMensaje(
                "Completa todos los campos obligatorios.",
                "error"
            );
            return;
        }

        try {
            crearTicketBtn.disabled = true;
            crearTicketBtn.innerHTML = "⏳ Creando...";

            const usuario = await obtenerUsuarioActual();
            const sitio = await obtenerSitioSharePoint();
            const lista = await obtenerListaTickets();

            const ticketID = generarTicketID();
            const fecha = new Date().toISOString();

            const fields = {
                Title: titulo,
                TicketID: ticketID,
                Usuario: usuario.displayName || "",
                Correo: usuario.mail || usuario.userPrincipalName || "",
                Categoria: categoria,
                Prioridad: prioridad,
                Descripcion: descripcion,
                Estado: "Pendiente",
                FechaCreacion: fecha,
                AsignadoA: "Soporte TI"
            };

            try {
                const campoSolucion = await obtenerNombreCampoSolucion();
                fields[campoSolucion] = "";
            } catch (error) {
                console.warn(
                    "No se pudo detectar TipoSolucion al crear el ticket. " +
                    "El ticket se creará sin ese campo.",
                    error
                );
            }

            const url =
                `https://graph.microsoft.com/v1.0/sites/` +
                `${sitio.id}/lists/${lista.id}/items`;

            await graphFetch(url, {
                method: "POST",
                body: JSON.stringify({ fields })
            });

            mostrarMensaje(
                `Ticket ${ticketID} creado correctamente.`,
                "success"
            );

            ticketForm.reset();
            ticketPrioridad.value = "Media";

            // Los usuarios normales pueden crear tickets sin ver registros.
            if (esUsuarioTI()) {
                await cargarTickets();
            }
        } catch (error) {
            console.error("Error creando ticket:", error);

            mostrarMensaje(
                "No se pudo crear el ticket. " +
                String(error.message || error).slice(0, 400),
                "error"
            );
        } finally {
            crearTicketBtn.disabled = false;
            crearTicketBtn.innerHTML = "<span>🎫</span> Crear ticket";
        }
    }

    function mostrarMensaje(mensaje, tipo) {
        ticketMensaje.textContent = mensaje;
        ticketMensaje.className = `ticket-message ${tipo}`;

        setTimeout(function () {
            ticketMensaje.className = "ticket-message";
            ticketMensaje.textContent = "";
        }, 5000);
    }

    /* =========================================================
       DETALLE DE TICKETS
    ========================================================= */

    function mostrarDetalleTicket(itemId) {
        if (!esUsuarioTI()) {
            return;
        }

        const ticket = ticketsData.find(function (item) {
            return String(item.id) === String(itemId);
        });

        if (!ticket) {
            return;
        }

        const fields = ticket.fields || {};

        const ticketId = obtenerCampo(fields, ["TicketID"]) ||
            `TKT-${String(ticket.id).padStart(6, "0")}`;

        const titulo = obtenerCampo(fields, ["Title", "Título"]);
        const usuario = obtenerCampo(fields, ["Usuario"]);
        const categoria = obtenerCampo(fields, ["Categoria", "Categoría"]);
        const prioridad = obtenerCampo(fields, ["Prioridad"]);
        const estado = obtenerCampo(fields, ["Estado"]) || "Pendiente";
        const fecha = obtenerCampo(fields, ["FechaCreacion", "FechaCreación"]);
        const descripcion = obtenerCampo(fields, ["Descripcion", "Descripción"]);
        const solucion = obtenerSolucion(fields);

        document.getElementById("modalTicketId").textContent = ticketId;

        const estadoElement = document.getElementById("modalTicketEstado");

        estadoElement.textContent = estado;
        estadoElement.className = `estado-badge ${claseEstado(estado)}`;

        document.getElementById("modalTicketTitulo").textContent = titulo;
        document.getElementById("modalTicketUsuario").textContent = usuario || "—";
        document.getElementById("modalTicketCategoria").textContent = categoria || "—";
        document.getElementById("modalTicketPrioridad").textContent = prioridad || "—";
        document.getElementById("modalTicketFecha").textContent = formatearFecha(fecha);
        document.getElementById("modalTicketDescripcion").textContent = descripcion || "—";

        const solucionBox = document.getElementById("modalTicketSolucionBox");
        const solucionElement = document.getElementById("modalTicketSolucion");

        if (solucion) {
            solucionElement.textContent = solucion;
            solucionBox.style.display = "block";
        } else {
            solucionElement.textContent = "—";
            solucionBox.style.display = "none";
        }

        ticketModal.classList.add("show");
        document.body.style.overflow = "hidden";
    }

    function cerrarTicketModal() {
        ticketModal.classList.remove("show");
        document.body.style.overflow = "";
    }

    /* =========================================================
       RESOLUCIÓN DE TICKETS
    ========================================================= */

    function abrirResolverTicket(event, itemId) {
        if (event) {
            event.stopPropagation();
        }

        if (!esUsuarioTI()) {
            return;
        }

        const ticket = ticketsData.find(function (item) {
            return String(item.id) === String(itemId);
        });

        if (!ticket) {
            return;
        }

        const fields = ticket.fields || {};

        const ticketId = obtenerCampo(fields, ["TicketID"]) ||
            `TKT-${String(ticket.id).padStart(6, "0")}`;

        const estado = obtenerCampo(fields, ["Estado"]) || "Pendiente";
        const solucion = obtenerSolucion(fields);

        ticketSeleccionadoResolver = ticket;

        resolverTicketId.textContent = ticketId;
        resolverEstado.value =
            estado === "Sin resolver" ? "Sin resolver" : "Resuelto";

        resolverDescripcion.value = solucion || "";
        resolverMensaje.textContent = "";
        resolverMensaje.className = "ticket-message";

        resolverModal.classList.add("show");
        document.body.style.overflow = "hidden";
    }

    function cerrarResolverTicket() {
        resolverModal.classList.remove("show");
        document.body.style.overflow = "";

        ticketSeleccionadoResolver = null;
        resolverDescripcion.value = "";
        resolverMensaje.textContent = "";
        resolverMensaje.className = "ticket-message";
    }

    function mostrarResolverMensaje(mensaje, tipo) {
        resolverMensaje.textContent = mensaje;
        resolverMensaje.className = `ticket-message ${tipo}`;
    }

    async function guardarResultadoTicket() {
        if (!esUsuarioTI() || !ticketSeleccionadoResolver) {
            return;
        }

        const estado = resolverEstado.value;
        const descripcion = resolverDescripcion.value.trim();

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
            confirmarResolver.disabled = true;
            confirmarResolver.innerHTML = "⏳ Guardando...";

            const sitio = await obtenerSitioSharePoint();
            const lista = await obtenerListaTickets();
            const campoSolucion = await obtenerNombreCampoSolucion();
            const itemId = ticketSeleccionadoResolver.id;

            const url =
                `https://graph.microsoft.com/v1.0/sites/` +
                `${sitio.id}/lists/${lista.id}/items/` +
                `${itemId}/fields`;

            const body = { Estado: estado };
            body[campoSolucion] = descripcion;

            console.log("Actualizando ticket:", {
                itemId,
                campoSolucion,
                estado,
                descripcion
            });

            await graphFetch(url, {
                method: "PATCH",
                body: JSON.stringify(body)
            });

            if (ticketSeleccionadoResolver.fields) {
                ticketSeleccionadoResolver.fields.Estado = estado;
                ticketSeleccionadoResolver.fields[campoSolucion] = descripcion;
            }

            mostrarResolverMensaje(
                "Resultado guardado correctamente.",
                "success"
            );

            setTimeout(async function () {
                cerrarResolverTicket();
                renderizarTickets();
                actualizarResumen();
                actualizarCantidadReporte();

                await cargarTickets();
            }, 700);
        } catch (error) {
            console.error("Error guardando resultado del ticket:", error);

            mostrarResolverMensaje(
                "No se pudo guardar el resultado. " +
                String(error?.message || error || "Error desconocido").slice(0, 600),
                "error"
            );
        } finally {
            confirmarResolver.disabled = false;
            confirmarResolver.innerHTML = "✓ Guardar resultado";
        }
    }

    /* =========================================================
       REPORTES: FECHAS Y FILTROS
    ========================================================= */

    function establecerMesReporteActual() {
        const fecha = new Date();
        const anio = fecha.getFullYear();
        const mes = String(fecha.getMonth() + 1).padStart(2, "0");

        reporteMes.value = `${anio}-${mes}`;
    }

    function obtenerNombreMesReporte(valorMes) {
        if (!valorMes) {
            return "";
        }

        const partes = valorMes.split("-");

        if (partes.length !== 2) {
            return "";
        }

        const anio = Number(partes[0]);
        const mes = Number(partes[1]);

        if (!anio || !mes) {
            return "";
        }

        const fecha = new Date(anio, mes - 1, 1);

        const nombre = fecha.toLocaleDateString("es-PE", {
            month: "long",
            year: "numeric"
        });

        return nombre.charAt(0).toUpperCase() + nombre.slice(1);
    }

    function obtenerTicketsDelMes(valorMes) {
        if (!valorMes) {
            return [];
        }

        const partes = valorMes.split("-");

        if (partes.length !== 2) {
            return [];
        }

        const anio = Number(partes[0]);
        const mes = Number(partes[1]);

        if (!anio || !mes) {
            return [];
        }

        return ticketsData.filter(function (item) {
            const fecha = obtenerFechaTicket(item);

            if (!fecha) {
                return false;
            }

            const date = new Date(fecha);

            if (Number.isNaN(date.getTime())) {
                return false;
            }

            return (
                date.getFullYear() === anio &&
                date.getMonth() + 1 === mes
            );
        });
    }

    function actualizarCantidadReporte() {
        if (!reporteMes || !reporteCantidad || !reportePeriodo) {
            return;
        }

        const tickets = obtenerTicketsDelMes(reporteMes.value);

        reporteCantidad.textContent = tickets.length;
        reportePeriodo.textContent =
            obtenerNombreMesReporte(reporteMes.value) || "—";
    }

    function mostrarReporteMensaje(mensaje, tipo) {
        reporteMensaje.textContent = mensaje;
        reporteMensaje.className = `ticket-message ${tipo}`;
    }

    function agruparTicketsPorCampo(tickets, nombresCampo) {
        const conteo = {};

        tickets.forEach(function (item) {
            const valor =
                obtenerCampo(item.fields || {}, nombresCampo) ||
                "Sin especificar";

            const clave = String(valor);

            if (!conteo[clave]) {
                conteo[clave] = 0;
            }

            conteo[clave]++;
        });

        return Object.entries(conteo).sort(function (a, b) {
            return b[1] - a[1];
        });
    }

    function textoPDF(valor) {
        if (valor === null || valor === undefined || valor === "") {
            return "—";
        }

        return String(valor);
    }

    function agregarPiePagina(doc, anchoPagina, altoPagina) {
        const paginas = doc.internal.getNumberOfPages();

        for (let pagina = 1; pagina <= paginas; pagina++) {
            doc.setPage(pagina);
            doc.setDrawColor(220, 220, 220);
            doc.line(10, altoPagina - 13, anchoPagina - 10, altoPagina - 13);

            doc.setFont("helvetica", "normal");
            doc.setFontSize(7);
            doc.setTextColor(110);

            doc.text(
                "ALFERZA LIVE OFFICE - Mesa de ayuda TI",
                10,
                altoPagina - 7
            );

            doc.text(
                `Página ${pagina} de ${paginas}`,
                anchoPagina - 10,
                altoPagina - 7,
                { align: "right" }
            );
        }
    }

    /* =========================================================
       GENERACIÓN DEL INFORME PDF
    ========================================================= */

    function generarInformePDF() {
        // Verificación de permisos antes de generar el archivo.
        if (!aplicarPermisosReportes()) {
            return;
        }

        if (typeof window.jspdf === "undefined") {
            mostrarReporteMensaje(
                "No se pudo cargar la librería PDF.",
                "error"
            );
            return;
        }

        if (typeof generarReportePDFBtn === "undefined" || !generarReportePDFBtn) {
            mostrarReporteMensaje(
                "No se encontró el botón para generar el PDF.",
                "error"
            );
            return;
        }

        const jsPDF = window.jspdf.jsPDF;

        if (typeof jsPDF !== "function") {
            mostrarReporteMensaje(
                "La librería jsPDF no está disponible.",
                "error"
            );
            return;
        }

        const valorMes = reporteMes.value;

        if (!valorMes) {
            mostrarReporteMensaje(
                "Selecciona el mes del reporte.",
                "error"
            );

            reporteMes.focus();
            return;
        }

        const tickets = obtenerTicketsDelMes(valorMes);

        if (!tickets.length) {
            mostrarReporteMensaje(
                "No existen atenciones registradas durante el mes seleccionado.",
                "error"
            );
            return;
        }

        try {
            generarReportePDFBtn.disabled = true;
            generarReportePDFBtn.innerHTML = "⏳ Generando...";

            const doc = new jsPDF({
                orientation: "landscape",
                unit: "mm",
                format: "a4"
            });

            const anchoPagina = doc.internal.pageSize.getWidth();
            const altoPagina = doc.internal.pageSize.getHeight();

            const nombreMes = obtenerNombreMesReporte(valorMes);

            const fechaGeneracion = new Date().toLocaleString("es-PE", {
                day: "2-digit",
                month: "2-digit",
                year: "numeric",
                hour: "2-digit",
                minute: "2-digit"
            });

            // Encabezado.
            doc.setFont("helvetica", "bold");
            doc.setFontSize(20);
            doc.setTextColor(0, 32, 92);
            doc.text("ALFERZA LIVE OFFICE", 14, 16);

            doc.setFont("helvetica", "normal");
            doc.setFontSize(10);
            doc.setTextColor(90);
            doc.text("Mesa de ayuda TI", 14, 23);

            doc.setFont("helvetica", "bold");
            doc.setFontSize(16);
            doc.setTextColor(0, 32, 92);
            doc.text(
                "INFORME MENSUAL DE ATENCIONES",
                anchoPagina - 14,
                16,
                { align: "right" }
            );

            doc.setFont("helvetica", "normal");
            doc.setFontSize(10);
            doc.setTextColor(70);
            doc.text(nombreMes, anchoPagina - 14, 23, { align: "right" });
            doc.text(
                `Generado: ${fechaGeneracion}`,
                anchoPagina - 14,
                29,
                { align: "right" }
            );

            doc.setDrawColor(215, 215, 215);
            doc.line(14, 34, anchoPagina - 14, 34);

            // Resumen de estados.
            const total = tickets.length;

            const pendientes = tickets.filter(function (item) {
                return obtenerCampo(item.fields, ["Estado"]) === "Pendiente";
            }).length;

            const proceso = tickets.filter(function (item) {
                return obtenerCampo(item.fields, ["Estado"]) === "En proceso";
            }).length;

            const resueltos = tickets.filter(function (item) {
                const estado = obtenerCampo(item.fields, ["Estado"]);

                return estado === "Resuelto" || estado === "Cerrado";
            }).length;

            const sinResolver = tickets.filter(function (item) {
                return obtenerCampo(item.fields, ["Estado"]) === "Sin resolver";
            }).length;

            doc.setFont("helvetica", "bold");
            doc.setFontSize(10);
            doc.setTextColor(35);
            doc.text("RESUMEN DE ATENCIONES", 14, 43);

            doc.autoTable({
                startY: 46,
                margin: { left: 14, right: 14 },
                head: [[
                    "Total",
                    "Pendientes",
                    "En proceso",
                    "Resueltos / Cerrados",
                    "Sin resolver"
                ]],
                body: [[
                    total,
                    pendientes,
                    proceso,
                    resueltos,
                    sinResolver
                ]],
                theme: "grid",
                styles: {
                    font: "helvetica",
                    fontSize: 9,
                    cellPadding: 3,
                    halign: "center",
                    valign: "middle"
                },
                headStyles: {
                    fillColor: [0, 32, 92],
                    textColor: 255,
                    fontStyle: "bold"
                }
            });

            // Atenciones por categoría y prioridad.
            const categorias = agruparTicketsPorCampo(
                tickets,
                ["Categoria", "Categoría"]
            );

            const prioridades = agruparTicketsPorCampo(
                tickets,
                ["Prioridad"]
            );

            const inicioResumen = doc.lastAutoTable.finalY + 8;

            doc.setFont("helvetica", "bold");
            doc.setFontSize(10);
            doc.setTextColor(35);
            doc.text("ATENCIONES POR CATEGORÍA", 14, inicioResumen);

            doc.autoTable({
                startY: inicioResumen + 3,
                margin: { left: 14 },
                tableWidth: 95,
                head: [["Categoría", "Cantidad"]],
                body: categorias.map(fila => [
                    textoPDF(fila[0]),
                    fila[1]
                ]),
                theme: "grid",
                styles: {
                    font: "helvetica",
                    fontSize: 8,
                    cellPadding: 2.5
                },
                headStyles: {
                    fillColor: [0, 32, 92],
                    textColor: 255,
                    fontStyle: "bold"
                }
            });

            const segundaTablaX = 125;

            doc.setFont("helvetica", "bold");
            doc.setFontSize(10);
            doc.setTextColor(35);
            doc.text(
                "ATENCIONES POR PRIORIDAD",
                segundaTablaX,
                inicioResumen
            );

            doc.autoTable({
                startY: inicioResumen + 3,
                margin: { left: segundaTablaX },
                tableWidth: 90,
                head: [["Prioridad", "Cantidad"]],
                body: prioridades.map(fila => [
                    textoPDF(fila[0]),
                    fila[1]
                ]),
                theme: "grid",
                styles: {
                    font: "helvetica",
                    fontSize: 8,
                    cellPadding: 2.5
                },
                headStyles: {
                    fillColor: [0, 32, 92],
                    textColor: 255,
                    fontStyle: "bold"
                }
            });

            // Tabla general.
            doc.addPage();

            doc.setFont("helvetica", "bold");
            doc.setFontSize(13);
            doc.setTextColor(0, 32, 92);
            doc.text("DETALLE DE ATENCIONES", 14, 16);

            doc.setFont("helvetica", "normal");
            doc.setFontSize(9);
            doc.setTextColor(90);
            doc.text(
                `Periodo: ${nombreMes} | Total de atenciones: ${total}`,
                14,
                22
            );

            const filasGenerales = tickets.map(function (item) {
                const fields = item.fields || {};

                const ticketId = obtenerCampo(fields, ["TicketID"]) ||
                    `TKT-${String(item.id).padStart(6, "0")}`;

                const asunto = obtenerCampo(fields, ["Title", "Título"]);
                const categoria = obtenerCampo(fields, ["Categoria", "Categoría"]);
                const prioridad = obtenerCampo(fields, ["Prioridad"]);
                const estado = obtenerCampo(fields, ["Estado"]) || "Pendiente";
                const fecha = obtenerFechaTicket(item);
                const usuario = obtenerCampo(fields, ["Usuario"]);
                const asignado = obtenerCampo(fields, ["AsignadoA"]);

                return [
                    textoPDF(ticketId),
                    textoPDF(asunto),
                    textoPDF(categoria),
                    textoPDF(prioridad),
                    textoPDF(estado),
                    formatearFecha(fecha),
                    textoPDF(usuario),
                    textoPDF(asignado)
                ];
            });

            doc.autoTable({
                startY: 27,
                margin: {
                    top: 15,
                    right: 8,
                    bottom: 18,
                    left: 8
                },
                head: [[
                    "Ticket",
                    "Asunto",
                    "Categoría",
                    "Prioridad",
                    "Estado",
                    "Fecha",
                    "Solicitante",
                    "Asignado"
                ]],
                body: filasGenerales,
                theme: "striped",
                styles: {
                    font: "helvetica",
                    fontSize: 7,
                    cellPadding: 2,
                    overflow: "linebreak",
                    valign: "middle"
                },
                headStyles: {
                    fillColor: [0, 32, 92],
                    textColor: 255,
                    fontStyle: "bold",
                    halign: "center",
                    valign: "middle"
                },
                columnStyles: {
                    0: { cellWidth: 22 },
                    1: { cellWidth: 43 },
                    2: { cellWidth: 28 },
                    3: { cellWidth: 20, halign: "center" },
                    4: { cellWidth: 25, halign: "center" },
                    5: { cellWidth: 31, halign: "center" },
                    6: { cellWidth: 42 },
                    7: { cellWidth: 28 }
                }
            });

            // Descripciones y resultados.
            doc.addPage();

            doc.setFont("helvetica", "bold");
            doc.setFontSize(13);
            doc.setTextColor(0, 32, 92);
            doc.text("DETALLE DE ATENCIÓN Y RESULTADO", 14, 16);

            doc.setFont("helvetica", "normal");
            doc.setFontSize(9);
            doc.setTextColor(90);
            doc.text(
                "Descripción reportada y resultado registrado por soporte TI.",
                14,
                22
            );

            const filasDetalle = tickets.map(function (item) {
                const fields = item.fields || {};

                const ticketId = obtenerCampo(fields, ["TicketID"]) ||
                    `TKT-${String(item.id).padStart(6, "0")}`;

                const descripcion = obtenerCampo(
                    fields,
                    ["Descripcion", "Descripción"]
                );

                const solucion = obtenerSolucion(fields);

                return [
                    textoPDF(ticketId),
                    textoPDF(descripcion),
                    textoPDF(solucion)
                ];
            });

            doc.autoTable({
                startY: 27,
                margin: {
                    top: 15,
                    right: 10,
                    bottom: 18,
                    left: 10
                },
                head: [[
                    "Ticket",
                    "Descripción de la atención / solicitud",
                    "Resultado / Solución"
                ]],
                body: filasDetalle,
                theme: "striped",
                styles: {
                    font: "helvetica",
                    fontSize: 7.5,
                    cellPadding: 2.5,
                    overflow: "linebreak",
                    valign: "top"
                },
                headStyles: {
                    fillColor: [0, 32, 92],
                    textColor: 255,
                    fontStyle: "bold",
                    halign: "center"
                },
                columnStyles: {
                    0: { cellWidth: 25 },
                    1: { cellWidth: 115 },
                    2: { cellWidth: 115 }
                }
            });

            agregarPiePagina(doc, anchoPagina, altoPagina);

            const nombreArchivo = `Informe_Atenciones_${valorMes}.pdf`;

            doc.save(nombreArchivo);

            mostrarReporteMensaje(
                "Informe PDF generado correctamente.",
                "success"
            );

            setTimeout(function () {
                cerrarReporte();
            }, 1000);
        } catch (error) {
            console.error("Error generando informe PDF:", error);

            mostrarReporteMensaje(
                "No se pudo generar el PDF: " +
                String(error?.message || error || "Error desconocido").slice(0, 500),
                "error"
            );
        } finally {
            generarReportePDFBtn.disabled = false;
            generarReportePDFBtn.innerHTML = "🧾 Generar PDF";
        }
    }

    function abrirModalReporte() {
        if (!aplicarPermisosReportes()) {
            return;
        }

        establecerMesReporteActual();
        actualizarCantidadReporte();

        reporteMensaje.textContent = "";
        reporteMensaje.className = "ticket-message";

        reporteModal.classList.add("show");
        document.body.style.overflow = "hidden";

        setTimeout(function () {
            reporteMes.focus();
        }, 100);
    }

    function cerrarReporte() {
        if (!reporteModal) {
            return;
        }

        reporteModal.classList.remove("show");
        document.body.style.overflow = "";

        if (reporteMensaje) {
            reporteMensaje.textContent = "";
            reporteMensaje.className = "ticket-message";
        }
    }

    /* =========================================================
       INICIO DE SESIÓN
    ========================================================= */

    async function esperarCuenta(ms = 6000) {
        const fin = Date.now() + ms;

        while (Date.now() < fin) {
            if (msalInstance.getAllAccounts().length) {
                return true;
            }

            await new Promise(resolve => setTimeout(resolve, 400));
        }

        return false;
    }

    function mostrarBotonLogin(texto) {
        usuarioNombre.textContent = "Sesión de Microsoft requerida";
        usuarioCorreo.textContent = texto;

        // No mostrar datos de registros mientras no exista sesión válida.
        if (registrosSection) {
            registrosSection.style.display = "none";
        }

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

        const btnLoginTickets = document.getElementById("btnLoginTickets");

        if (!btnLoginTickets) {
            return;
        }

        btnLoginTickets.addEventListener("click", async function () {
            try {
                const cuentas = msalInstance.getAllAccounts();

                if (cuentas.length) {
                    await msalInstance.acquireTokenPopup({
                        scopes: SCOPES_GRAPH,
                        account: cuentas[0]
                    });
                } else {
                    await msalInstance.loginPopup({
                        scopes: SCOPES_GRAPH
                    });
                }

                iniciarTickets();
            } catch (error) {
                console.error("Error de inicio de sesión:", error);

                usuarioCorreo.textContent =
                    "No se pudo iniciar sesión: " +
                    String(error.message || error).slice(0, 300);
            }
        });
    }

    async function iniciarTickets() {
        try {
            if (!(await esperarCuenta())) {
                mostrarBotonLogin(
                    "Inicia sesión para crear tickets"
                );
                return;
            }

            await obtenerUsuarioActual();

            // La consulta de registros solo se ejecuta para el grupo TI.
            if (esUsuarioTI()) {
                await cargarTickets();
            } else {
                ticketsData = [];
                actualizarResumen();
            }
        } catch (error) {
            console.error("Error inicializando Tickets:", error);

            if (
                error.codigo === "REQUIERE_INTERACCION" ||
                error.codigo === "SIN_SESION"
            ) {
                mostrarBotonLogin(
                    "Autoriza los permisos de Microsoft"
                );
                return;
            }

            usuarioNombre.textContent = "No se pudo obtener el usuario";
            usuarioCorreo.textContent =
                String(error.message || error).slice(0, 300);
        }
    }

    /* =========================================================
       EVENTOS
    ========================================================= */

    ticketForm.addEventListener("submit", crearTicket);

    actualizarTickets.addEventListener("click", function () {
        if (aplicarPermisosReportes()) {
            cargarTickets();
        }
    });

    abrirReporte.addEventListener("click", abrirModalReporte);

    buscarTicket.addEventListener("input", renderizarTickets);
    filtroEstado.addEventListener("change", renderizarTickets);
    filtroPrioridad.addEventListener("change", renderizarTickets);

    cerrarModal.addEventListener("click", cerrarTicketModal);

    ticketModal.addEventListener("click", function (event) {
        if (event.target === ticketModal) {
            cerrarTicketModal();
        }
    });

    cerrarResolverModal.addEventListener("click", cerrarResolverTicket);
    cancelarResolver.addEventListener("click", cerrarResolverTicket);
    confirmarResolver.addEventListener("click", guardarResultadoTicket);

    resolverModal.addEventListener("click", function (event) {
        if (event.target === resolverModal) {
            cerrarResolverTicket();
        }
    });

    cerrarReporteModal.addEventListener("click", cerrarReporte);
    cancelarReporte.addEventListener("click", cerrarReporte);
    generarReportePDFBtn.addEventListener("click", generarInformePDF);

    reporteMes.addEventListener("change", actualizarCantidadReporte);

    reporteModal.addEventListener("click", function (event) {
        if (event.target === reporteModal) {
            cerrarReporte();
        }
    });

    document.addEventListener("keydown", function (event) {
        if (event.key !== "Escape") {
            return;
        }

        if (reporteModal.classList.contains("show")) {
            cerrarReporte();
            return;
        }

        if (resolverModal.classList.contains("show")) {
            cerrarResolverTicket();
            return;
        }

        if (ticketModal.classList.contains("show")) {
            cerrarTicketModal();
        }
    });

    /* =========================================================
       FUNCIONES DISPONIBLES PARA LOS BOTONES DE LA TABLA
    ========================================================= */

    window.abrirResolverTicket = abrirResolverTicket;
    window.mostrarDetalleTicket = mostrarDetalleTicket;

    /* =========================================================
       INICIALIZACIÓN
    ========================================================= */

    establecerMesReporteActual();
    actualizarCantidadReporte();

    inicializarPermisosReportes();
    iniciarTickets();

})();