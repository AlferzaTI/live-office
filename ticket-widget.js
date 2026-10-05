(() => {
    "use strict";

    // ============================================================
    // CONFIGURACIÓN
    // ============================================================

    const MSAL_CONFIG = {
        auth: {
            clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",
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

    const SHAREPOINT_HOST = "alferzaholding-my.sharepoint.com";
    const SHAREPOINT_SITE_PATH = "/personal/soporte1_alferza_pe";
    const TICKETS_LIST_NAME = "TicketsTI";

    const SCOPES = [
        "User.Read",
        "Sites.Read.All",
        "Sites.ReadWrite.All"
    ];

    let msalInstance = null;
    let cuentaActual = null;
    let sitioSharePoint = null;
    let listaTickets = null;

    let imagenesSeleccionadas = [];

    // ============================================================
    // CARGAR CSS DEL WIDGET
    // ============================================================

    function cargarCSS() {
        const scriptActual =
            document.currentScript ||
            document.querySelector('script[src*="ticket-widget.js"]');

        if (!scriptActual) return;

        const scriptURL = new URL(scriptActual.src);

        const cssURL = new URL(
            "ticket-widget.css",
            scriptURL
        ).href;

        if (!document.querySelector(`link[href="${cssURL}"]`)) {
            const link = document.createElement("link");
            link.rel = "stylesheet";
            link.href = cssURL;
            document.head.appendChild(link);
        }
    }

    // ============================================================
    // INICIALIZAR MSAL
    // ============================================================

    async function inicializarMSAL() {
        if (msalInstance) return;

        if (typeof msal === "undefined") {
            console.error("MSAL no está cargado.");
            return;
        }

        msalInstance =
            new msal.PublicClientApplication(MSAL_CONFIG);

        const cuentas = msalInstance.getAllAccounts();

        if (cuentas.length > 0) {
            cuentaActual = cuentas[0];
        }
    }

    // ============================================================
    // TOKEN
    // ============================================================

    async function obtenerToken() {
        await inicializarMSAL();

        if (!msalInstance) {
            throw new Error("MSAL no está disponible.");
        }

        let cuenta =
            cuentaActual ||
            msalInstance.getAllAccounts()[0];

        if (!cuenta) {
            const loginResponse =
                await msalInstance.loginPopup({
                    scopes: SCOPES
                });

            cuentaActual = loginResponse.account;
            cuenta = loginResponse.account;
        }

        try {
            const respuesta =
                await msalInstance.acquireTokenSilent({
                    scopes: SCOPES,
                    account: cuenta
                });

            return respuesta.accessToken;

        } catch (error) {

            const respuesta =
                await msalInstance.acquireTokenPopup({
                    scopes: SCOPES
                });

            return respuesta.accessToken;
        }
    }

    // ============================================================
    // GRAPH
    // ============================================================

    async function graphFetch(url, options = {}) {

        const token = await obtenerToken();

        const headers = {
            Authorization: `Bearer ${token}`,
            ...(options.headers || {})
        };

        if (
            options.body &&
            typeof options.body !== "string"
        ) {
            headers["Content-Type"] =
                "application/json";

            options.body =
                JSON.stringify(options.body);
        }

        return fetch(url, {
            ...options,
            headers
        });
    }

    // ============================================================
    // SHAREPOINT
    // ============================================================

    async function obtenerSitioSharePoint() {

        if (sitioSharePoint) {
            return sitioSharePoint;
        }

        const url =
            `https://graph.microsoft.com/v1.0/sites/` +
            `${SHAREPOINT_HOST}:${SHAREPOINT_SITE_PATH}`;

        const respuesta =
            await graphFetch(url);

        if (!respuesta.ok) {
            throw new Error(
                "No se pudo obtener el sitio de SharePoint."
            );
        }

        sitioSharePoint =
            await respuesta.json();

        return sitioSharePoint;
    }

    async function obtenerListaTickets() {

        if (listaTickets) {
            return listaTickets;
        }

        const sitio =
            await obtenerSitioSharePoint();

        const url =
            `https://graph.microsoft.com/v1.0/sites/` +
            `${sitio.id}/lists`;

        const respuesta =
            await graphFetch(url);

        if (!respuesta.ok) {
            throw new Error(
                "No se pudieron obtener las listas."
            );
        }

        const datos =
            await respuesta.json();

        listaTickets =
            datos.value.find(
                lista =>
                    lista.displayName ===
                    TICKETS_LIST_NAME
            );

        if (!listaTickets) {
            throw new Error(
                `No se encontró la lista "${TICKETS_LIST_NAME}".`
            );
        }

        return listaTickets;
    }

    // ============================================================
    // USUARIO
    // ============================================================

    async function obtenerUsuario() {

        const respuesta =
            await graphFetch(
                "https://graph.microsoft.com/v1.0/me"
            );

        if (!respuesta.ok) {
            throw new Error(
                "No se pudo obtener el usuario."
            );
        }

        return await respuesta.json();
    }

    // ============================================================
    // GENERAR ID
    // ============================================================

    function generarTicketID() {

        const ahora = new Date();

        const año =
            ahora.getFullYear();

        const numero =
            Math.floor(
                1000 + Math.random() * 9000
            );

        return `TKT-${año}-${numero}`;
    }

    // ============================================================
    // CREAR WIDGET
    // ============================================================

    function crearWidget() {

        if (
            document.getElementById(
                "ticketWidgetButton"
            )
        ) {
            return;
        }

        const scriptActual =
            document.currentScript ||
            document.querySelector(
                'script[src*="ticket-widget.js"]'
            );

        let imagenGIF = "img/ticket.gif";

        if (scriptActual) {
            imagenGIF =
                new URL(
                    "img/ticket.gif",
                    scriptActual.src
                ).href;
        }

        const widget =
            document.createElement("div");

        widget.id = "ticketWidgetButton";
        widget.className = "ticket-widget";

        widget.innerHTML = `
            <img
                src="${imagenGIF}"
                alt="Registrar ticket"
                title="Registrar ticket"
            >
        `;

        widget.addEventListener(
            "click",
            abrirTicketWidget
        );

        document.body.appendChild(widget);
    }

    // ============================================================
    // CREAR MODAL
    // ============================================================

    function crearModal() {

        if (
            document.getElementById(
                "ticketWidgetModal"
            )
        ) {
            return;
        }

        const modal =
            document.createElement("div");

        modal.id =
            "ticketWidgetModal";

        modal.className =
            "ticket-widget-overlay";

        modal.innerHTML = `

            <div class="ticket-widget-modal">

                <div class="ticket-widget-header">

                    <div>
                        <span class="ticket-widget-icon">
                            🎫
                        </span>

                        <div>
                            <h2>Registrar ticket</h2>
                            <p>Solicita soporte al Área de TI</p>
                        </div>
                    </div>

                    <button
                        type="button"
                        class="ticket-widget-close"
                        id="cerrarTicketWidget"
                    >
                        ×
                    </button>

                </div>

                <form
                    id="ticketWidgetForm"
                    class="ticket-widget-form"
                >

                    <div class="ticket-widget-user">

                        <div>
                            <span>Usuario</span>
                            <strong id="widgetUsuarioNombre">
                                Cargando...
                            </strong>
                        </div>

                        <div>
                            <span>Correo</span>
                            <strong id="widgetUsuarioCorreo">
                                Cargando...
                            </strong>
                        </div>

                    </div>

                    <div class="ticket-widget-field">

                        <label for="widgetTicketTitulo">
                            Asunto
                        </label>

                        <input
                            type="text"
                            id="widgetTicketTitulo"
                            maxlength="150"
                            placeholder="Ej. No puedo acceder a mi correo"
                            required
                        >

                    </div>

                    <div class="ticket-widget-row">

                        <div class="ticket-widget-field">

                            <label for="widgetTicketCategoria">
                                Categoría
                            </label>

                            <select
                                id="widgetTicketCategoria"
                                required
                            >
                                <option value="">
                                    Seleccionar
                                </option>
                                <option value="Correo">
                                    Correo
                                </option>
                                <option value="Internet">
                                    Internet
                                </option>
                                <option value="Equipo">
                                    Equipo
                                </option>
                                <option value="Impresora">
                                    Impresora
                                </option>
                                <option value="Acceso">
                                    Acceso
                                </option>
                                <option value="Software">
                                    Software
                                </option>
                                <option value="Microsoft 365">
                                    Microsoft 365
                                </option>
                                <option value="Otro">
                                    Otro
                                </option>
                            </select>

                        </div>

                        <div class="ticket-widget-field">

                            <label for="widgetTicketPrioridad">
                                Prioridad
                            </label>

                            <select
                                id="widgetTicketPrioridad"
                                required
                            >
                                <option value="Baja">
                                    Baja
                                </option>
                                <option
                                    value="Media"
                                    selected
                                >
                                    Media
                                </option>
                                <option value="Alta">
                                    Alta
                                </option>
                                <option value="Crítica">
                                    Crítica
                                </option>
                            </select>

                        </div>

                    </div>

                    <div class="ticket-widget-field">

                        <label for="widgetTicketDescripcion">
                            Descripción
                        </label>

                        <textarea
                            id="widgetTicketDescripcion"
                            maxlength="2000"
                            rows="4"
                            placeholder="Describe brevemente el problema..."
                            required
                        ></textarea>

                    </div>

                    <div class="ticket-widget-field">

                        <label>
                            Imagen
                        </label>

                        <div
                            id="ticketImageDropzone"
                            class="ticket-image-dropzone"
                        >

                            <input
                                type="file"
                                id="ticketImageInput"
                                accept="image/*"
                                multiple
                                hidden
                            >

                            <div class="ticket-upload-icon">
                                📎
                            </div>

                            <strong>
                                Adjuntar imágenes
                            </strong>

                            <span>
                                Haz clic, arrastra o pega
                                con Ctrl + V
                            </span>

                            <button
                                type="button"
                                id="seleccionarImagenBtn"
                                class="ticket-upload-button"
                            >
                                Seleccionar imagen
                            </button>

                        </div>

                        <div
                            id="ticketImagePreview"
                            class="ticket-image-preview"
                        ></div>

                    </div>

                    <div
                        id="ticketWidgetMensaje"
                        class="ticket-widget-message"
                    ></div>

                    <div class="ticket-widget-actions">

                        <button
                            type="button"
                            id="cancelarTicketWidget"
                            class="ticket-widget-btn secondary"
                        >
                            Cancelar
                        </button>

                        <button
                            type="submit"
                            id="registrarTicketWidget"
                            class="ticket-widget-btn primary"
                        >
                            Registrar ticket
                        </button>

                    </div>

                </form>

            </div>
        `;

        document.body.appendChild(modal);

        configurarEventosModal();
    }

    // ============================================================
    // EVENTOS
    // ============================================================

    function configurarEventosModal() {

        document
            .getElementById(
                "cerrarTicketWidget"
            )
            .addEventListener(
                "click",
                cerrarTicketWidget
            );

        document
            .getElementById(
                "cancelarTicketWidget"
            )
            .addEventListener(
                "click",
                cerrarTicketWidget
            );

        document
            .getElementById(
                "ticketWidgetModal"
            )
            .addEventListener(
                "click",
                event => {

                    if (
                        event.target.id ===
                        "ticketWidgetModal"
                    ) {
                        cerrarTicketWidget();
                    }

                }
            );

        document
            .getElementById(
                "ticketWidgetForm"
            )
            .addEventListener(
                "submit",
                registrarTicket
            );

        const input =
            document.getElementById(
                "ticketImageInput"
            );

        const boton =
            document.getElementById(
                "seleccionarImagenBtn"
            );

        const dropzone =
            document.getElementById(
                "ticketImageDropzone"
            );

        boton.addEventListener(
            "click",
            () => input.click()
        );

        input.addEventListener(
            "change",
            event => {

                agregarImagenes(
                    event.target.files
                );

                input.value = "";

            }
        );

        dropzone.addEventListener(
            "dragover",
            event => {

                event.preventDefault();

                dropzone.classList.add(
                    "dragging"
                );

            }
        );

        dropzone.addEventListener(
            "dragleave",
            () => {

                dropzone.classList.remove(
                    "dragging"
                );

            }
        );

        dropzone.addEventListener(
            "drop",
            event => {

                event.preventDefault();

                dropzone.classList.remove(
                    "dragging"
                );

                agregarImagenes(
                    event.dataTransfer.files
                );

            }
        );

        document.addEventListener(
            "paste",
            manejarPegadoImagen,
            true
        );

        document.addEventListener(
            "keydown",
            event => {

                if (
                    event.key === "Escape" &&
                    document
                        .getElementById(
                            "ticketWidgetModal"
                        )
                        ?.classList.contains("visible")
                ) {
                    cerrarTicketWidget();
                }

            }
        );
    }

    // ============================================================
    // ABRIR
    // ============================================================

    async function abrirTicketWidget() {

        crearModal();

        const modal =
            document.getElementById(
                "ticketWidgetModal"
            );

        modal.classList.add("visible");

        document.body.classList.add(
            "ticket-modal-open"
        );

        await cargarDatosUsuario();
    }

    // ============================================================
    // CERRAR
    // ============================================================

    function cerrarTicketWidget() {

        const modal =
            document.getElementById(
                "ticketWidgetModal"
            );

        if (!modal) return;

        modal.classList.remove(
            "visible"
        );

        document.body.classList.remove(
            "ticket-modal-open"
        );

        limpiarFormulario();
    }

    // ============================================================
    // USUARIO
    // ============================================================

    async function cargarDatosUsuario() {

        const nombre =
            document.getElementById(
                "widgetUsuarioNombre"
            );

        const correo =
            document.getElementById(
                "widgetUsuarioCorreo"
            );

        try {

            const usuario =
                await obtenerUsuario();

            nombre.textContent =
                usuario.displayName ||
                "Usuario";

            correo.textContent =
                usuario.mail ||
                usuario.userPrincipalName ||
                "";

        } catch (error) {

            console.error(error);

            nombre.textContent =
                "No disponible";

            correo.textContent =
                "No disponible";
        }
    }

    // ============================================================
    // IMÁGENES
    // ============================================================

    function manejarPegadoImagen(event) {

        const modal =
            document.getElementById(
                "ticketWidgetModal"
            );

        if (
            !modal ||
            !modal.classList.contains("visible")
        ) {
            return;
        }

        const items =
            event.clipboardData?.items;

        if (!items) return;

        const archivos = [];

        for (const item of items) {

            if (
                item.type &&
                item.type.startsWith("image/")
            ) {

                const archivo =
                    item.getAsFile();

                if (archivo) {
                    archivos.push(archivo);
                }
            }
        }

        if (archivos.length > 0) {

            event.preventDefault();

            agregarImagenes(archivos);
        }
    }

    function agregarImagenes(archivos) {

        const archivosValidos =
            Array.from(archivos).filter(
                archivo =>
                    archivo.type.startsWith(
                        "image/"
                    )
            );

        for (const archivo of archivosValidos) {

            if (
                imagenesSeleccionadas.length >= 5
            ) {

                mostrarMensaje(
                    "Puedes adjuntar máximo 5 imágenes.",
                    "error"
                );

                break;
            }

            if (
                archivo.size >
                5 * 1024 * 1024
            ) {

                mostrarMensaje(
                    `"${archivo.name}" supera los 5 MB.`,
                    "error"
                );

                continue;
            }

            imagenesSeleccionadas.push(
                archivo
            );
        }

        renderizarPreviews();
    }

    function renderizarPreviews() {

        const contenedor =
            document.getElementById(
                "ticketImagePreview"
            );

        if (!contenedor) return;

        contenedor.innerHTML = "";

        imagenesSeleccionadas.forEach(
            (archivo, index) => {

                const reader =
                    new FileReader();

                reader.onload = event => {

                    const item =
                        document.createElement(
                            "div"
                        );

                    item.className =
                        "ticket-image-item";

                    item.innerHTML = `

                        <img
                            src="${event.target.result}"
                            alt="Imagen adjunta"
                        >

                        <div class="ticket-image-info">
                            <span>
                                ${escapeHTML(
                                    archivo.name
                                )}
                            </span>
                            <small>
                                ${formatearTamaño(
                                    archivo.size
                                )}
                            </small>
                        </div>

                        <button
                            type="button"
                            class="ticket-image-remove"
                            title="Eliminar imagen"
                        >
                            ×
                        </button>
                    `;

                    item
                        .querySelector(
                            ".ticket-image-remove"
                        )
                        .addEventListener(
                            "click",
                            () => {

                                imagenesSeleccionadas.splice(
                                    index,
                                    1
                                );

                                renderizarPreviews();
                            }
                        );

                    contenedor.appendChild(
                        item
                    );
                };

                reader.readAsDataURL(
                    archivo
                );
            }
        );
    }

    function formatearTamaño(bytes) {

        if (bytes < 1024) {
            return `${bytes} B`;
        }

        if (bytes < 1024 * 1024) {
            return `${(
                bytes / 1024
            ).toFixed(1)} KB`;
        }

        return `${(
            bytes /
            (1024 * 1024)
        ).toFixed(1)} MB`;
    }

    // ============================================================
    // REGISTRAR TICKET
    // ============================================================

    async function registrarTicket(event) {

        event.preventDefault();

        const boton =
            document.getElementById(
                "registrarTicketWidget"
            );

        const titulo =
            document.getElementById(
                "widgetTicketTitulo"
            ).value.trim();

        const categoria =
            document.getElementById(
                "widgetTicketCategoria"
            ).value;

        const prioridad =
            document.getElementById(
                "widgetTicketPrioridad"
            ).value;

        const descripcion =
            document.getElementById(
                "widgetTicketDescripcion"
            ).value.trim();

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

            boton.disabled = true;

            boton.textContent =
                "Registrando...";

            mostrarMensaje(
                "",
                ""
            );

            const usuario =
                await obtenerUsuario();

            const lista =
                await obtenerListaTickets();

            const sitio =
                await obtenerSitioSharePoint();

            const ticketID =
                generarTicketID();

            const fecha =
                new Date().toISOString();

            const body = {
                fields: {
                    Title: titulo,
                    TicketID: ticketID,

                    Usuario:
                        usuario.displayName || "",

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
                `${sitio.id}/lists/` +
                `${lista.id}/items`;

            const respuesta =
                await graphFetch(
                    url,
                    {
                        method: "POST",
                        body
                    }
                );

            if (!respuesta.ok) {

                const errorTexto =
                    await respuesta.text();

                console.error(
                    errorTexto
                );

                throw new Error(
                    "No se pudo registrar el ticket."
                );
            }

            const ticketCreado =
                await respuesta.json();

            /*
             * IMPORTANTE:
             * Las imágenes se muestran y se preparan
             * correctamente en el formulario.
             *
             * SharePoint List Attachments no se pueden
             * subir directamente mediante Graph.
             *
             * Por eso aquí dejamos el ticket creado
             * correctamente y luego podemos conectar
             * las imágenes mediante SharePoint REST.
             */

            console.log(
                "Ticket creado:",
                ticketCreado
            );

            mostrarMensaje(
                `Ticket ${ticketID} registrado correctamente.`,
                "success"
            );

            boton.textContent =
                "✓ Registrado";

            setTimeout(() => {

                cerrarTicketWidget();

                /*
                 * Si estamos en tickets.html,
                 * intentamos actualizar la tabla
                 * automáticamente.
                 */

                if (
                    typeof window.cargarTickets ===
                    "function"
                ) {
                    window.cargarTickets();
                }

            }, 1300);

        } catch (error) {

            console.error(
                "Error registrando ticket:",
                error
            );

            mostrarMensaje(
                error.message ||
                "Ocurrió un error al registrar el ticket.",
                "error"
            );

            boton.disabled = false;

            boton.textContent =
                "Registrar ticket";
        }
    }

    // ============================================================
    // MENSAJES
    // ============================================================

    function mostrarMensaje(
        mensaje,
        tipo
    ) {

        const elemento =
            document.getElementById(
                "ticketWidgetMensaje"
            );

        if (!elemento) return;

        elemento.textContent =
            mensaje;

        elemento.className =
            "ticket-widget-message";

        if (tipo) {
            elemento.classList.add(
                tipo
            );
        }
    }

    // ============================================================
    // LIMPIAR
    // ============================================================

    function limpiarFormulario() {

        const form =
            document.getElementById(
                "ticketWidgetForm"
            );

        if (form) {
            form.reset();
        }

        imagenesSeleccionadas = [];

        renderizarPreviews();

        const boton =
            document.getElementById(
                "registrarTicketWidget"
            );

        if (boton) {

            boton.disabled = false;

            boton.textContent =
                "Registrar ticket";
        }

        mostrarMensaje("", "");
    }

    // ============================================================
    // SEGURIDAD HTML
    // ============================================================

    function escapeHTML(texto) {

        const div =
            document.createElement(
                "div"
            );

        div.textContent =
            texto;

        return div.innerHTML;
    }

    // ============================================================
    // INICIO
    // ============================================================

    function iniciar() {

        cargarCSS();

        crearWidget();
    }

    if (
        document.readyState ===
        "loading"
    ) {

        document.addEventListener(
            "DOMContentLoaded",
            iniciar
        );

    } else {

        iniciar();
    }

})();