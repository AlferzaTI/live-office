
/* =========================================================
   ALFERZA LIVE OFFICE - INFRAESTRUCTURA
   Inventario, monitoreo y gestión de trabajadores
========================================================= */

"use strict";

/* ======================== CONFIGURACIÓN ======================== */

const CONFIG = {
    clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",
    tenantId: "dbab984f-4bb1-4b60-9dff-da59f54acdf1",
    redirectUri: "https://alferzati.github.io/live-office/blank.html",
    hostname: "alferzaholding-my.sharepoint.com",
    sitePath: "/personal/soporte1_alferza_pe",
    driveId: "b!1l0EwOKXpU24cUGYaNjPRiTQWIyIIChCkDDL197LcoOFOD7ThItqTIJ_Z1KoYrrS",
    listaServicios: "MonitoreoTI",
    listaTrabajadores: "TrabajadoresEquipos",
    archivoEquipos: "Procedimientos T.I/equipos.json",
    scopes: ["User.Read", "Sites.Read.All", "Sites.ReadWrite.All"],
    cachePrefix: "alferza_ti_",
    equiposPorPagina: 40,
    intervaloActualizacion: 120000
};

const GRAPH = "https://graph.microsoft.com/v1.0";
const $ = id => document.getElementById(id);

/* ======================== MICROSOFT AUTH ======================== */

const msalInfraestructura = typeof msal !== "undefined"
    ? new msal.PublicClientApplication({
        auth: {
            clientId: CONFIG.clientId,
            authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
            redirectUri: CONFIG.redirectUri
        },
        cache: {
            cacheLocation: "sessionStorage",
            storeAuthStateInCookie: false
        }
    })
    : null;

/* ======================== ELEMENTOS HTML ======================== */

const ui = {
    loading: $("loadingOverlay"),
    loadingMessage: $("loadingMessage"),
    servicios: $("serviciosGrid"),
    estadoGeneral: $("estadoGeneral"),

    tabla: $("equiposTableBody"),
    vacio: $("equiposEmpty"),
    resultados: $("equiposResultados"),
    total: $("totalEquipos"),
    conectados: $("equiposConectados"),
    desconectados: $("equiposDesconectados"),
    estado: $("equiposEstado"),
    ultimaRevision: $("lastUpdate"),

    filtroEstado: $("filtroEstadoEquipo"),
    filtroArea: $("filtroAreaEquipo"),
    busqueda: $("buscarEquipo"),

    modal: $("gestionTrabajadorModal"),
    busquedaEquipo: $("gestionIP"),
    buscarEquipo: $("buscarEquipoIP"),
    equipoEncontrado: $("equipoEncontrado"),
    nombreEquipo: $("gestionEquipoNombre"),
    detalleEquipo: $("gestionEquipoDetalle"),
    trabajador: $("gestionTrabajador"),
    mensaje: $("gestionTrabajadorMensaje"),

    abrirGestion: $("btnGestionarTrabajadores"),
    registrar: $("btnRegistrarTrabajador"),
    editar: $("btnEditarTrabajador"),
    eliminar: $("btnEliminarTrabajador"),
    cerrar: $("cerrarGestionTrabajadorModal"),
    cancelar: $("cancelarGestionTrabajador")
};

/* ======================== ESTADO ======================== */

let cuentaActual = null;
let tokenActual = null;

let sitioActual = null;
let listaTrabajadoresActual = null;
let columnasTrabajadores = null;
let camposTrabajadores = null;

let promesaSitio = null;
let promesaTrabajadores = null;
let promesaCargaEquipos = null;

let equiposData = [];
let equiposFiltrados = [];
let trabajadoresData = [];
let trabajadoresPorEquipo = {};

let ubicacionEquipos = null;
let etagEquipos = null;
let paginaActual = 1;
let equipoGestionSeleccionado = null;

let trabajadoresCargados = false;
let trabajadoresListos = false;
let ultimaActualizacion = 0;

/* ======================== UTILIDADES ======================== */

function normalizar(valor) {
    return String(valor ?? "").trim().toLowerCase();
}

function sinTildes(valor) {
    return String(valor ?? "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim()
        .toLowerCase();
}

function escaparHTML(valor) {
    return String(valor ?? "").replace(/[&<>"']/g, caracter => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#039;"
    })[caracter]);
}

function fechaActual() {
    return new Date().toLocaleString("es-PE", {
        dateStyle: "short",
        timeStyle: "medium"
    });
}

function leerCache(clave) {
    try {
        return JSON.parse(
            localStorage.getItem(CONFIG.cachePrefix + clave)
        );
    } catch {
        return null;
    }
}

function guardarCache(clave, valor) {
    try {
        localStorage.setItem(
            CONFIG.cachePrefix + clave,
            JSON.stringify(valor)
        );
    } catch {
        // La caché es opcional.
    }
}

function mostrarCarga(mensaje = "") {
    if (ui.loading) ui.loading.style.display = "flex";
    if (ui.loadingMessage) ui.loadingMessage.textContent = mensaje;
}

function ocultarCarga() {
    if (ui.loading) ui.loading.style.display = "none";
}

/* ======================== MENSAJES DEL FORMULARIO ======================== */

function configurarEstilosMensajes() {
    if (document.getElementById("alferza-mensajes-estilos")) return;

    const estilo = document.createElement("style");
    estilo.id = "alferza-mensajes-estilos";

    estilo.textContent = `
        #gestionTrabajadorMensaje {
            display: none;
            align-items: flex-start;
            gap: 10px;
            padding: 12px 14px;
            margin: 14px 0;
            border: 1px solid transparent;
            border-radius: 10px;
            font-size: 13px;
            line-height: 1.5;
            font-weight: 500;
            overflow-wrap: anywhere;
            box-sizing: border-box;
            animation: alferzaMensajeEntrada .2s ease;
        }

        #gestionTrabajadorMensaje.error {
            color: #b42318;
            background: #fef3f2;
            border-color: #fecdca;
        }

        #gestionTrabajadorMensaje.success {
            color: #067647;
            background: #ecfdf3;
            border-color: #abefc6;
        }

        #gestionTrabajadorMensaje.info {
            color: #175cd3;
            background: #eff8ff;
            border-color: #b2ddff;
        }

        #gestionTrabajadorMensaje.warning {
            color: #b54708;
            background: #fffaeb;
            border-color: #fedf89;
        }

        #gestionTrabajadorMensaje .alferza-mensaje-icono {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            flex-shrink: 0;
            width: 20px;
            height: 20px;
            border-radius: 50%;
            background: currentColor;
            color: white;
            font-size: 12px;
            font-weight: 800;
        }

        @keyframes alferzaMensajeEntrada {
            from { opacity: 0; transform: translateY(-4px); }
            to { opacity: 1; transform: translateY(0); }
        }
    `;

    document.head.appendChild(estilo);
}

function mostrarMensaje(texto = "", tipo = "error") {
    if (!ui.mensaje) return;

    configurarEstilosMensajes();

    ui.mensaje.replaceChildren();
    ui.mensaje.className = `gestion-mensaje ${tipo}`;
    ui.mensaje.style.display = texto ? "flex" : "none";
    ui.mensaje.setAttribute("role", tipo === "error" ? "alert" : "status");
    ui.mensaje.setAttribute("aria-live", "polite");

    if (!texto) return;

    const iconos = {
        success: "✓",
        error: "!",
        info: "i",
        warning: "!"
    };

    const icono = document.createElement("span");
    icono.className = "alferza-mensaje-icono";
    icono.textContent = iconos[tipo] || "i";

    const contenido = document.createElement("span");
    contenido.textContent = texto;

    ui.mensaje.append(icono, contenido);
}

function mensajeErrorAmigable(error, operacion = "realizar la operación") {
    // Los detalles técnicos se conservan en la consola, no en el formulario.
    console.error(`[ALFERZA] Error al ${operacion}:`, error);

    if (error?.status === 401) {
        return "Tu sesión ha caducado. Vuelve a iniciar sesión.";
    }

    if (error?.status === 403) {
        const detalle = String(error?.message || "");

        if (/LinkTitle|read.only|readonly/i.test(detalle)) {
            return "La lista de SharePoint tiene una columna no editable configurada incorrectamente. Revisa las columnas de TrabajadoresEquipos.";
        }

        return "No tienes permisos para modificar este registro. Contacta con TI.";
    }

    if (error?.status === 404) {
        return "No se encontró el registro solicitado. Busca nuevamente el equipo.";
    }

    if (error?.status === 409) {
        return "El registro cambió en SharePoint. Busca nuevamente el equipo e inténtalo otra vez.";
    }

    if (error?.name === "AbortError") {
        return "La solicitud tardó demasiado. Inténtalo nuevamente.";
    }

    if (!navigator.onLine) {
        return "No hay conexión a Internet. Comprueba tu conexión.";
    }

    return "No se pudo completar la operación. Comprueba los datos e inténtalo nuevamente.";
}

/* ======================== AUTENTICACIÓN ======================== */

async function obtenerTokenSilencioso() {
    if (!msalInfraestructura) return null;

    const cuentas = msalInfraestructura.getAllAccounts();
    if (!cuentas.length) return null;

    cuentaActual = cuentaActual || cuentas[0];

    try {
        const resultado = await msalInfraestructura.acquireTokenSilent({
            scopes: CONFIG.scopes,
            account: cuentaActual
        });

        tokenActual = resultado.accessToken;
        return tokenActual;
    } catch (error) {
        console.warn("[ALFERZA] No se pudo obtener el token silenciosamente.", error);
        return null;
    }
}

async function iniciarSesion(interactivo = false) {
    const token = await obtenerTokenSilencioso();

    if (token || !interactivo || !msalInfraestructura) {
        return token;
    }

    const resultado = cuentaActual
        ? await msalInfraestructura.acquireTokenPopup({
            scopes: CONFIG.scopes,
            account: cuentaActual
        })
        : await msalInfraestructura.loginPopup({
            scopes: CONFIG.scopes
        });

    cuentaActual = resultado.account;
    tokenActual = resultado.accessToken || await obtenerTokenSilencioso();

    return tokenActual;
}

/* ======================== MICROSOFT GRAPH ======================== */

async function graphFetch(url, opciones = {}, intento = 0) {
    const { timeoutMs = 30000, ...fetchOptions } = opciones;
    const controlador = new AbortController();
    const temporizador = setTimeout(
        () => controlador.abort(),
        timeoutMs
    );

    try {
        const respuesta = await fetch(url, {
            ...fetchOptions,
            signal: controlador.signal,
            headers: {
                Authorization: `Bearer ${tokenActual}`,
                ...(fetchOptions.headers || {})
            }
        });

        if ([429, 503].includes(respuesta.status) && intento < 3) {
            const retryAfter = Number(respuesta.headers.get("Retry-After"));
            const espera = Math.min(
                (retryAfter || 2 ** intento) * 1000,
                10000
            );

            await new Promise(resolve => setTimeout(resolve, espera));
            return graphFetch(url, opciones, intento + 1);
        }

        if (respuesta.status === 401 && intento < 1) {
            if (await obtenerTokenSilencioso()) {
                return graphFetch(url, opciones, intento + 1);
            }
        }

        return respuesta;
    } finally {
        clearTimeout(temporizador);
    }
}

async function obtenerJSON(url, opciones = {}) {
    const respuesta = await graphFetch(url, opciones);

    if (!respuesta.ok) {
        const detalle = await respuesta.text().catch(() => "");

        const error = new Error(
            `Microsoft Graph (${respuesta.status}): ${detalle || respuesta.statusText}`
        );

        error.status = respuesta.status;
        throw error;
    }

    if (respuesta.status === 204) return null;

    const tipo = respuesta.headers.get("content-type") || "";

    if (!tipo.includes("json")) return null;

    return respuesta.json();
}

/* ======================== SHAREPOINT: SITIO Y LISTAS ======================== */

async function obtenerSitio() {
    if (sitioActual) return sitioActual;

    if (!promesaSitio) {
        promesaSitio = obtenerJSON(
            `${GRAPH}/sites/${CONFIG.hostname}:${CONFIG.sitePath}?$select=id`
        )
            .then(sitio => {
                sitioActual = sitio;
                return sitio;
            })
            .catch(error => {
                promesaSitio = null;
                throw error;
            });
    }

    return promesaSitio;
}

async function obtenerListaTrabajadores() {
    if (listaTrabajadoresActual) return listaTrabajadoresActual;

    const sitio = await obtenerSitio();
    let url =
        `${GRAPH}/sites/${sitio.id}/lists` +
        `?$select=id,displayName&$top=200`;

    while (url) {
        const datos = await obtenerJSON(url);

        const lista = (datos.value || []).find(item =>
            normalizar(item.displayName) === normalizar(CONFIG.listaTrabajadores)
        );

        if (lista) {
            listaTrabajadoresActual = lista;
            return lista;
        }

        url = datos["@odata.nextLink"] || null;
    }

    throw new Error(`No se encontró la lista ${CONFIG.listaTrabajadores}.`);
}

async function obtenerColumnasTrabajadores() {
    if (columnasTrabajadores) return columnasTrabajadores;

    const sitio = await obtenerSitio();
    const lista = await obtenerListaTrabajadores();

    let url =
        `${GRAPH}/sites/${sitio.id}/lists/${lista.id}/columns` +
        `?$top=200`;

    const columnas = [];

    while (url) {
        const datos = await obtenerJSON(url);
        columnas.push(...(datos.value || []));
        url = datos["@odata.nextLink"] || null;
    }

    columnasTrabajadores = columnas;
    return columnasTrabajadores;
}

function buscarColumna(columnas, nombres) {
    const candidatos = nombres.map(sinTildes);

    return columnas.find(columna => {
        const nombreInterno = sinTildes(columna.name);
        const nombreVisible = sinTildes(columna.displayName);

        if (
            nombreInterno === "linktitle" ||
            nombreInterno.startsWith("linktitle")
        ) {
            return false;
        }

        if (columna.readOnly === true) return false;
        if (columna.computed === true) return false;

        return (
            candidatos.includes(nombreInterno) ||
            candidatos.includes(nombreVisible)
        );
    }) || null;
}

async function obtenerCamposTrabajadores() {
    if (camposTrabajadores) return camposTrabajadores;

    const columnas = await obtenerColumnasTrabajadores();

    const columnaEquipo = buscarColumna(columnas, [
        "Equipo1",
        "Equipo",
        "NombreEquipo",
        "Nombre del equipo"
    ]);

    const columnaTrabajador = buscarColumna(columnas, [
        "NombreTrabajador",
        "Trabajador",
        "Nombre del trabajador"
    ]);

    if (!columnaEquipo || !columnaTrabajador) {
        console.error(
            "[ALFERZA] Columnas de TrabajadoresEquipos:",
            columnas.map(c => ({
                name: c.name,
                displayName: c.displayName,
                readOnly: c.readOnly,
                computed: c.computed
            }))
        );

        throw new Error(
            "No se identificaron las columnas editables del equipo y del trabajador en TrabajadoresEquipos."
        );
    }

    camposTrabajadores = {
        equipo: columnaEquipo.name,
        trabajador: columnaTrabajador.name
    };

    return camposTrabajadores;
}

/* ======================== INVENTARIO: EQUIPOS.JSON ======================== */

async function buscarEquiposJson(raiz) {
    const nombre = CONFIG.archivoEquipos.split("/").pop();

    const datos = await obtenerJSON(
        `${raiz}/root/search(q='${encodeURIComponent(nombre)}')` +
        `?$select=id,name,eTag&$top=50`
    );

    const archivo = (datos.value || []).find(item =>
        normalizar(item.name) === normalizar(nombre)
    );

    if (!archivo) {
        throw new Error("No se encontró equipos.json en SharePoint.");
    }

    return {
        meta: `${raiz}/items/${archivo.id}`,
        contenido: `${raiz}/items/${archivo.id}/content`,
        etag: archivo.eTag || null
    };
}

async function leerJsonEquipos(respuesta) {
    const buffer = await respuesta.arrayBuffer();
    const bytes = new Uint8Array(buffer);

    let codificacion = "utf-8";

    if (bytes[0] === 0xFF && bytes[1] === 0xFE) {
        codificacion = "utf-16le";
    }

    if (bytes[0] === 0xFE && bytes[1] === 0xFF) {
        codificacion = "utf-16be";
    }

    const texto = new TextDecoder(codificacion)
        .decode(buffer)
        .replace(/^\uFEFF/, "");

    let datos;

    try {
        datos = JSON.parse(texto);
    } catch {
        throw new Error("El archivo equipos.json no contiene JSON válido.");
    }

    if (Array.isArray(datos)) return datos;

    if (datos && typeof datos === "object") {
        for (const clave of [
            "value",
            "equipos",
            "Equipos",
            "data",
            "items"
        ]) {
            if (Array.isArray(datos[clave])) return datos[clave];
        }

        if (datos.Nombre) return [datos];
    }

    throw new Error("El formato de equipos.json no es válido.");
}

async function obtenerEquiposDesdeSharePoint() {
    const raiz = `${GRAPH}/drives/${CONFIG.driveId}`;
    const ruta = CONFIG.archivoEquipos
        .split("/")
        .map(encodeURIComponent)
        .join("/");

    let ubicacion = ubicacionEquipos;
    let etag = null;

    if (ubicacion) {
        try {
            const meta = await obtenerJSON(`${ubicacion.meta}?$select=eTag`);
            etag = meta.eTag || null;
        } catch (error) {
            if (error.status !== 404) throw error;
            ubicacion = null;
        }
    }

    if (!ubicacion) {
        try {
            const meta = await obtenerJSON(
                `${raiz}/root:/${ruta}?$select=id,eTag`
            );

            ubicacion = {
                meta: `${raiz}/items/${meta.id}`,
                contenido: `${raiz}/items/${meta.id}/content`
            };

            etag = meta.eTag || null;
        } catch (error) {
            if (error.status !== 404) throw error;

            ubicacion = await buscarEquiposJson(raiz);
            etag = ubicacion.etag;
        }
    }

    ubicacionEquipos = ubicacion;

    if (etag && etag === etagEquipos && equiposData.length) {
        return null;
    }

    const respuesta = await graphFetch(
        ubicacion.contenido,
        { timeoutMs: 90000 }
    );

    if (!respuesta.ok) {
        const error = new Error("No se pudo descargar el inventario.");
        error.status = respuesta.status;
        throw error;
    }

    const equipos = await leerJsonEquipos(respuesta);
    etagEquipos = etag;

    return equipos;
}

/* ======================== IDENTIFICADORES DE EQUIPOS ======================== */

function identificadoresEquipo(equipo) {
    if (!equipo) return [];

    if (typeof equipo !== "object") {
        return [normalizar(equipo)].filter(Boolean);
    }

    return [
        equipo.Nombre,
        equipo.nombre,
        equipo.NombreEquipo,
        equipo.nombreEquipo,
        equipo.DNSHostName,
        equipo.Hostname,
        equipo.hostname,
        equipo.ComputerName,
        equipo.computerName,
        equipo.IP
    ].map(normalizar).filter(Boolean);
}

function encontrarEquipo(entrada) {
    const consulta = sinTildes(entrada);

    if (!consulta) {
        return {
            equipo: null,
            multiples: false,
            coincidencias: []
        };
    }

    const identificadores = equipo => identificadoresEquipo(equipo)
        .map(sinTildes);

    let coincidencias = equiposData.filter(equipo =>
        identificadores(equipo).some(valor => valor === consulta)
    );

    if (coincidencias.length === 1) {
        return {
            equipo: coincidencias[0],
            multiples: false,
            coincidencias
        };
    }

    if (coincidencias.length > 1) {
        return {
            equipo: null,
            multiples: true,
            coincidencias
        };
    }

    coincidencias = equiposData.filter(equipo =>
        identificadores(equipo).some(valor => valor.includes(consulta))
    );

    return {
        equipo: coincidencias.length === 1 ? coincidencias[0] : null,
        multiples: coincidencias.length > 1,
        coincidencias
    };
}

/* ======================== TRABAJADORES ======================== */

async function obtenerTrabajadores() {
    if (trabajadoresCargados) return trabajadoresPorEquipo;

    if (!promesaTrabajadores) {
        promesaTrabajadores = cargarTrabajadores().finally(() => {
            promesaTrabajadores = null;
        });
    }

    return promesaTrabajadores;
}

async function cargarTrabajadores() {
    try {
        const sitio = await obtenerSitio();
        const lista = await obtenerListaTrabajadores();
        const campos = await obtenerCamposTrabajadores();

        let url =
            `${GRAPH}/sites/${sitio.id}/lists/${lista.id}/items` +
            `?$expand=fields&$top=200`;

        const elementos = [];

        while (url) {
            const datos = await obtenerJSON(url);
            elementos.push(...(datos.value || []));
            url = datos["@odata.nextLink"] || null;
        }

        trabajadoresData = elementos;
        trabajadoresPorEquipo = {};

        elementos.forEach(item => {
            const fields = item.fields || {};
            const equipo = fields[campos.equipo];
            const nombre = fields[campos.trabajador];

            if (!equipo || !nombre) return;

            trabajadoresPorEquipo[normalizar(equipo)] = {
                nombre,
                itemId: item.id,
                equipo
            };
        });

        trabajadoresCargados = true;
        return trabajadoresPorEquipo;

    } finally {
        trabajadoresListos = true;
        renderizarTablaEquipos();
    }
}

function obtenerTrabajadorEquipo(equipo) {
    const identificadores = identificadoresEquipo(equipo);

    for (const identificador of identificadores) {
        const asignacion = trabajadoresPorEquipo[normalizar(identificador)];

        if (asignacion) return asignacion;
    }

    const claves = Object.keys(trabajadoresPorEquipo);

    for (const clave of claves) {
        if (identificadores.some(id => normalizar(id) === clave)) {
            return trabajadoresPorEquipo[clave];
        }
    }

    return null;
}

function actualizarAsignacionLocal(equipo, asignacion) {
    const identificadores = identificadoresEquipo(equipo);
    const itemId = asignacion?.itemId ||
        obtenerTrabajadorEquipo(equipo)?.itemId;

    for (const clave of Object.keys(trabajadoresPorEquipo)) {
        const actual = trabajadoresPorEquipo[clave];

        if (
            identificadores.includes(normalizar(clave)) ||
            (
                itemId &&
                String(actual.itemId) === String(itemId)
            )
        ) {
            delete trabajadoresPorEquipo[clave];
        }
    }

    if (asignacion) {
        trabajadoresPorEquipo[normalizar(asignacion.equipo)] = asignacion;
    }

    renderizarTablaEquipos();
}

/* ======================== MONITOREO DE SERVICIOS ======================== */

async function obtenerMonitoreoTI() {
    const sitio = await obtenerSitio();

    let url =
        `${GRAPH}/sites/${sitio.id}/lists` +
        `?$select=id,displayName&$top=200`;

    let listaServicios = null;

    while (url) {
        const datos = await obtenerJSON(url);

        listaServicios = (datos.value || []).find(item =>
            normalizar(item.displayName) === normalizar(CONFIG.listaServicios)
        );

        if (listaServicios) break;

        url = datos["@odata.nextLink"] || null;
    }

    if (!listaServicios) {
        throw new Error(`No se encontró la lista ${CONFIG.listaServicios}.`);
    }

    url =
        `${GRAPH}/sites/${sitio.id}/lists/${listaServicios.id}/items` +
        `?$expand=fields&$top=200`;

    const elementos = [];

    while (url) {
        const datos = await obtenerJSON(url);
        elementos.push(...(datos.value || []));
        url = datos["@odata.nextLink"] || null;
    }

    return elementos;
}

function obtenerCampo(fields, nombres, defecto = "") {
    const claves = Object.keys(fields).reduce((resultado, clave) => {
        resultado[sinTildes(clave)] = fields[clave];
        return resultado;
    }, {});

    for (const nombre of nombres) {
        const valor = claves[sinTildes(nombre)];

        if (
            valor !== undefined &&
            valor !== null &&
            String(valor).trim()
        ) {
            return valor;
        }
    }

    return defecto;
}

function formatearRevision(valor) {
    if (!valor) return "--";

    const fecha = new Date(valor);

    if (Number.isNaN(fecha.getTime())) return String(valor);

    return fecha.toLocaleString("es-PE", {
        timeZone: "America/Lima",
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false
    });
}

function renderizarServicios(items) {
    if (!ui.servicios) return;

    if (!items.length) {
        ui.servicios.innerHTML = `
            <div class="service-card service-loading">
                <span>No hay servicios registrados.</span>
            </div>`;

        if (ui.estadoGeneral) {
            ui.estadoGeneral.textContent = "● SIN DATOS";
            ui.estadoGeneral.className = "section-status warning";
        }

        return;
    }

    let online = 0;

    ui.servicios.innerHTML = items.map(item => {
        const fields = item.fields || {};

        const nombre = obtenerCampo(
            fields,
            ["Servicio", "Nombre", "Title"],
            "Servicio"
        );

        const estado = obtenerCampo(
            fields,
            ["Estado", "Status"],
            "Sin estado"
        );

        const descripcion = obtenerCampo(
            fields,
            [
                "Descripcion",
                "Description",
                "Detalle",
                "Mensaje",
                "Observacion"
            ],
            "Sin descripción"
        );

        const latencia = obtenerCampo(
            fields,
            ["Latencia", "Ping", "Tiempo"],
            "--"
        );

        const revision = formatearRevision(
            obtenerCampo(
                fields,
                [
                    "UltimaRevision",
                    "Ultima revision",
                    "FechaRevision",
                    "Fecha"
                ],
                ""
            )
        );

        const estadoNormalizado = sinTildes(estado);
        let clase = "neutral";

        if (/operativo|online|conectado/.test(estadoNormalizado)) {
            clase = "online";
            online++;
        } else if (/mantenimiento|advertencia|warning|lento/.test(estadoNormalizado)) {
            clase = "warning";
        } else if (/caido|offline|desconectado|error/.test(estadoNormalizado)) {
            clase = "offline";
        }

        return `
            <article class="service-card ${clase}">
                <div class="service-main">
                    <div class="service-icon">🖥️</div>
                    <div class="service-info">
                        <div class="service-heading">
                            <h3>${escaparHTML(nombre)}</h3>
                            <span class="service-status ${clase}">
                                <span class="status-dot"></span>
                                ${escaparHTML(estado)}
                            </span>
                        </div>
                        <p>${escaparHTML(descripcion)}</p>
                    </div>
                </div>
                <div class="service-meta">
                    <div class="service-latency">
                        <span>Latencia</span>
                        <strong>${escaparHTML(latencia)}</strong>
                    </div>
                    <div class="service-revision">
                        <span class="revision-icon">🕐</span>
                        <div>
                            <span>Última revisión</span>
                            <strong>${escaparHTML(revision)}</strong>
                        </div>
                    </div>
                </div>
            </article>`;
    }).join("");

    if (ui.estadoGeneral) {
        const estado = online === items.length
            ? ["● OPERATIVO", "online"]
            : online
                ? ["● REVISAR", "warning"]
                : ["● NO DISPONIBLE", "offline"];

        ui.estadoGeneral.textContent = estado[0];
        ui.estadoGeneral.className = `section-status ${estado[1]}`;
    }
}

async function cargarServicios() {
    try {
        renderizarServicios(await obtenerMonitoreoTI());
    } catch (error) {
        console.error("[ALFERZA] Error al cargar servicios:", error);

        if (ui.servicios) {
            ui.servicios.innerHTML = `
                <div class="service-card service-loading">
                    <span>No se pudieron cargar los servicios.</span>
                </div>`;
        }

        if (ui.estadoGeneral) {
            ui.estadoGeneral.textContent = "● ERROR";
            ui.estadoGeneral.className = "section-status offline";
        }
    }
}

/* ======================== TABLA DE INVENTARIO ======================== */


function actualizarResumenEquipos() {
    const total = equiposData.length;

    const conectados = equiposData.filter(
        equipo => normalizar(equipo.Estado) === "conectado"
    ).length;

    const desconectados = equiposData.filter(
        equipo => normalizar(equipo.Estado) === "desconectado"
    ).length;

    // Todo equipo que no esté conectado ni desconectado
    // se considera no verificable.
    const noVerificables = equiposData.filter(equipo => {
        const estado = normalizar(equipo.Estado);

        return estado !== "conectado" &&
               estado !== "desconectado";
    }).length;

    if (ui.total) {
        ui.total.textContent = total;
    }

    if (ui.conectados) {
        ui.conectados.textContent = conectados;
    }

    if (ui.desconectados) {
        ui.desconectados.textContent = desconectados;
    }

    if (ui.noVerificables) {
        ui.noVerificables.textContent = noVerificables;
    }

    console.log("[ALFERZA] Resumen de equipos:", {
        total,
        conectados,
        desconectados,
        noVerificables,
        suma: conectados + desconectados + noVerificables,
        diferencia: total - conectados - desconectados - noVerificables
    });
}

function actualizarFiltroAreas() {
    if (!ui.filtroArea) return;

    const seleccion = ui.filtroArea.value;

    const areas = [...new Set(
        equiposData.map(equipo => equipo.Area).filter(Boolean)
    )].sort((a, b) => String(a).localeCompare(b, "es"));

    ui.filtroArea.replaceChildren(
        new Option("Todas las áreas", "todos")
    );

    areas.forEach(area => {
        ui.filtroArea.add(new Option(area, area));
    });

    if ([...ui.filtroArea.options].some(opcion => opcion.value === seleccion)) {
        ui.filtroArea.value = seleccion;
    }
}

function aplicarDatosEquipos(equipos) {
    equiposData = Array.isArray(equipos) ? equipos : [];

    actualizarResumenEquipos();
    actualizarFiltroAreas();
    aplicarFiltrosEquipos(true);
}

function aplicarFiltrosEquipos(conservarPagina = false) {
    const texto = normalizar(ui.busqueda?.value);
    const estado = ui.filtroEstado?.value || "todos";
    const area = ui.filtroArea?.value || "todos";

    equiposFiltrados = equiposData.filter(equipo =>
        (
            !texto ||
            [
                equipo.Nombre,
                equipo.IP,
                equipo.DNSHostName
            ].some(valor => normalizar(valor).includes(texto))
        ) &&
        (
            estado === "todos" ||
            String(equipo.Estado || "") === estado
        ) &&
        (
            area === "todos" ||
            String(equipo.Area || "") === area
        )
    );

    if (!conservarPagina) paginaActual = 1;

    renderizarTablaEquipos();
}

function claseEstadoEquipo(estado) {
    const valor = normalizar(estado);

    if (valor === "conectado") return "online";
    if (valor === "desconectado") return "offline";

    return "neutral";
}

function renderizarTablaEquipos() {
    if (!ui.tabla) return;

    const total = equiposFiltrados.length;

    const paginas = Math.max(
        1,
        Math.ceil(total / CONFIG.equiposPorPagina)
    );

    paginaActual = Math.min(paginaActual, paginas);

    const inicio = (paginaActual - 1) * CONFIG.equiposPorPagina;

    const pagina = equiposFiltrados.slice(
        inicio,
        inicio + CONFIG.equiposPorPagina
    );

    if (ui.resultados) {
        ui.resultados.textContent = `Equipos encontrados: ${total}`;
    }

    if (ui.vacio) {
        ui.vacio.style.display = total ? "none" : "flex";
    }

    if (!total) {
        ui.tabla.innerHTML = "";
        renderizarPaginacion(paginas);
        return;
    }

    ui.tabla.innerHTML = pagina.map(equipo => {
        const estado = equipo.Estado || "No verificable";
        const asignacion = obtenerTrabajadorEquipo(equipo);

        const trabajador = asignacion
            ? `<span class="trabajador-name" title="${escaparHTML(asignacion.nombre)}">${escaparHTML(asignacion.nombre)}</span>`
            : `<span style="color:#98a2b3;font-size:12px">${trabajadoresListos ? "Sin asignar" : "Cargando…"}</span>`;

        return `
            <tr>
                <td>
                    <div class="equipo-name">
                        <div class="equipo-icon">🖥️</div>
                        <div>
                            <strong title="${escaparHTML(equipo.Nombre)}">${escaparHTML(equipo.Nombre)}</strong>
                            <span title="${escaparHTML(equipo.DNSHostName)}">${escaparHTML(equipo.DNSHostName || "Sin DNS")}</span>
                        </div>
                    </div>
                </td>
                <td>
                    <span class="equipo-status ${claseEstadoEquipo(estado)}">
                        <span class="equipo-status-dot"></span>
                        ${escaparHTML(estado)}
                    </span>
                </td>
                <td class="trabajador-cell">${trabajador}</td>
                <td><span class="equipo-ip">${escaparHTML(equipo.IP || "--")}</span></td>
                <td>
                    <div class="equipo-so">
                        ${escaparHTML(equipo.SistemaOperativo || "--")}
                        ${equipo.VersionSO ? `<br><small>${escaparHTML(equipo.VersionSO)}</small>` : ""}
                    </div>
                </td>
                <td><span class="equipo-area">${escaparHTML(equipo.Area || "--")}</span></td>
                <td><span class="equipo-date">${escaparHTML(equipo.UltimoRegistroAD || "--")}</span></td>
            </tr>`;
    }).join("");

    renderizarPaginacion(paginas);
}

function generarPaginas(actual, total) {
    if (total <= 7) {
        return Array.from({ length: total }, (_, i) => i + 1);
    }

    const paginas = [1];

    if (actual > 4) paginas.push("...");

    for (
        let i = Math.max(2, actual - 1);
        i <= Math.min(total - 1, actual + 1);
        i++
    ) {
        paginas.push(i);
    }

    if (actual < total - 3) paginas.push("...");

    paginas.push(total);

    return paginas;
}

function renderizarPaginacion(totalPaginas) {
    document.querySelector(".equipos-pagination")?.remove();

    if (totalPaginas <= 1) return;

    const contenedor = document.createElement("div");
    contenedor.className = "equipos-pagination";

    function boton(texto, pagina, deshabilitado = false) {
        const elemento = document.createElement("button");

        elemento.type = "button";
        elemento.className = "pagination-btn";
        elemento.textContent = texto;
        elemento.disabled = deshabilitado;

        if (Number(texto) === paginaActual) {
            elemento.classList.add("active");
        }

        elemento.addEventListener("click", () => {
            paginaActual = pagina;
            renderizarTablaEquipos();
        });

        contenedor.appendChild(elemento);
    }

    boton("‹", paginaActual - 1, paginaActual === 1);

    generarPaginas(paginaActual, totalPaginas).forEach(pagina => {
        if (pagina === "...") {
            const puntos = document.createElement("span");
            puntos.className = "pagination-ellipsis";
            puntos.textContent = "...";
            contenedor.appendChild(puntos);
        } else {
            boton(String(pagina), pagina);
        }
    });

    boton("›", paginaActual + 1, paginaActual === totalPaginas);

    const tabla = document.querySelector(".equipos-table-container");

    if (tabla?.parentNode) {
        tabla.parentNode.insertBefore(contenedor, tabla.nextSibling);
    }
}

async function cargarEquipos() {
    // Evita iniciar dos descargas del inventario al mismo tiempo.
    if (promesaCargaEquipos) return promesaCargaEquipos;

    promesaCargaEquipos = (async () => {
        try {
            if (ui.estado) {
                ui.estado.textContent = "● ACTUALIZANDO";
                ui.estado.className = "section-status warning";
            }

            const equipos = await obtenerEquiposDesdeSharePoint();

            if (equipos) {
                aplicarDatosEquipos(equipos);
                guardarCache("equipos", {
                    etag: etagEquipos,
                    datos: equipos
                });
            }

            if (ui.ultimaRevision) {
                ui.ultimaRevision.textContent = fechaActual();
            }

            if (ui.estado) {
                ui.estado.textContent = "● ACTUALIZADO";
                ui.estado.className = "section-status online";
            }
        } catch (error) {
            console.error("[ALFERZA] Error al cargar inventario:", error);

            if (ui.estado) {
                ui.estado.textContent = equiposData.length
                    ? "● SIN ACTUALIZAR"
                    : "● ERROR";

                ui.estado.className = "section-status warning";
            }
        }
    })();

    try {
        await promesaCargaEquipos;
    } finally {
        promesaCargaEquipos = null;
    }
}

/* ======================== MODAL DE TRABAJADORES ======================== */

function reiniciarFormularioGestion() {
    equipoGestionSeleccionado = null;

    if (ui.busquedaEquipo) ui.busquedaEquipo.value = "";

    if (ui.trabajador) {
        ui.trabajador.value = "";
        ui.trabajador.disabled = true;
    }

    if (ui.equipoEncontrado) {
        ui.equipoEncontrado.style.display = "none";
    }

    mostrarMensaje("");

    if (ui.registrar) {
        ui.registrar.disabled = true;
        ui.registrar.textContent = "Registrar";
    }

    if (ui.editar) {
        ui.editar.disabled = true;
        ui.editar.textContent = "Editar";
    }

    if (ui.eliminar) {
        ui.eliminar.disabled = true;
        ui.eliminar.textContent = "Eliminar";
    }

    if (ui.buscarEquipo) {
        ui.buscarEquipo.disabled = false;
        ui.buscarEquipo.textContent = "Buscar";
    }
}

function abrirGestionTrabajadores() {
    if (!ui.modal) return;

    reiniciarFormularioGestion();
    ui.modal.style.display = "flex";

    setTimeout(() => ui.busquedaEquipo?.focus(), 100);
}

function cerrarGestionTrabajadores() {
    if (ui.modal) ui.modal.style.display = "none";

    reiniciarFormularioGestion();
}

function limpiarResultadoBusqueda() {
    equipoGestionSeleccionado = null;

    if (ui.equipoEncontrado) {
        ui.equipoEncontrado.style.display = "none";
    }

    if (ui.trabajador) {
        ui.trabajador.value = "";
        ui.trabajador.disabled = true;
    }

    if (ui.registrar) ui.registrar.disabled = true;
    if (ui.editar) ui.editar.disabled = true;
    if (ui.eliminar) ui.eliminar.disabled = true;
}

/* ======================== BUSCAR EQUIPO ======================== */

async function buscarEquipoPorIP() {
    const consulta = ui.busquedaEquipo?.value?.trim();

    if (!consulta) {
        mostrarMensaje("Ingresa la IP o el nombre del equipo.", "warning");
        ui.busquedaEquipo?.focus();
        return;
    }

    const boton = ui.buscarEquipo;

    try {
        if (boton) {
            boton.disabled = true;
            boton.textContent = "Buscando…";
        }

        mostrarMensaje("Buscando el equipo en el inventario…", "info");

        if (!equiposData.length) {
            await cargarEquipos();
        }

        if (!equiposData.length) {
            mostrarMensaje("No se pudo cargar el inventario. Inténtalo de nuevo.", "error");
            limpiarResultadoBusqueda();
            return;
        }

        const resultado = encontrarEquipo(consulta);

        if (!resultado.equipo) {
            limpiarResultadoBusqueda();

            if (resultado.multiples) {
                mostrarMensaje(
                    "Hay varios equipos coincidentes. Escribe un nombre, IP o DNS más específico.",
                    "warning"
                );
            } else {
                mostrarMensaje(
                    "No se encontró ningún equipo con esa IP o nombre.",
                    "warning"
                );
            }

            return;
        }

        const equipo = resultado.equipo;
        equipoGestionSeleccionado = equipo;

        const nombre =
            equipo.Nombre ||
            equipo.nombre ||
            equipo.NombreEquipo ||
            equipo.nombreEquipo ||
            equipo.DNSHostName ||
            equipo.Hostname ||
            equipo.ComputerName ||
            equipo.IP ||
            "Equipo sin nombre";

        if (ui.equipoEncontrado) {
            ui.equipoEncontrado.style.display = "block";
        }

        if (ui.nombreEquipo) {
            ui.nombreEquipo.textContent = nombre;
        }

        if (ui.detalleEquipo) {
            ui.detalleEquipo.textContent =
                `IP: ${equipo.IP || "No registrada"} | DNS: ${equipo.DNSHostName || "No registrado"}`;
        }

        await obtenerTrabajadores();

        const asignacion = obtenerTrabajadorEquipo(equipo);

        if (ui.trabajador) {
            ui.trabajador.disabled = false;
            ui.trabajador.value = asignacion?.nombre || "";
        }

        if (ui.registrar) ui.registrar.disabled = Boolean(asignacion);
        if (ui.editar) ui.editar.disabled = !asignacion;
        if (ui.eliminar) ui.eliminar.disabled = !asignacion;

        mostrarMensaje(
            asignacion
                ? "Equipo encontrado. Puedes editar o eliminar al trabajador asignado."
                : "Equipo encontrado. Escribe el nombre del trabajador y pulsa Registrar.",
            "success"
        );

    } catch (error) {
        limpiarResultadoBusqueda();
        mostrarMensaje(
            mensajeErrorAmigable(error, "buscar el equipo"),
            "error"
        );
    } finally {
        if (boton) {
            boton.disabled = false;
            boton.textContent = "Buscar";
        }
    }
}

/* ======================== VALIDACIÓN DE CAMPOS EDITABLES ======================== */

function validarCamposDeEscritura(campos) {
    const noEditables = [
        "linktitle",
        "linktitlenomenu",
        "linktitle2"
    ];

    for (const campo of [campos.equipo, campos.trabajador]) {
        const nombre = normalizar(campo);

        if (
            noEditables.includes(nombre) ||
            nombre.startsWith("linktitle")
        ) {
            throw new Error(
                "La columna detectada no es editable. Revisa los nombres internos de TrabajadoresEquipos."
            );
        }
    }
}

/* ======================== REGISTRAR TRABAJADOR ======================== */

async function registrarTrabajadorPorIP() {
    const equipo = equipoGestionSeleccionado;
    const nombre = ui.trabajador?.value?.trim();

    if (!equipo) {
        mostrarMensaje("Primero busca un equipo.", "warning");
        return;
    }

    if (!nombre) {
        mostrarMensaje("Escribe el nombre del trabajador.", "warning");
        ui.trabajador?.focus();
        return;
    }

    const nombreEquipo =
        equipo.Nombre ||
        equipo.nombre ||
        equipo.NombreEquipo ||
        equipo.nombreEquipo ||
        equipo.DNSHostName ||
        equipo.Hostname ||
        equipo.ComputerName ||
        equipo.IP;

    if (!nombreEquipo) {
        mostrarMensaje("El equipo no tiene un identificador válido.", "error");
        return;
    }

    const boton = ui.registrar;

    try {
        if (boton) {
            boton.disabled = true;
            boton.textContent = "Registrando…";
        }

        await obtenerTrabajadores();

        const existente = obtenerTrabajadorEquipo(equipo);

        if (existente) {
            if (ui.trabajador) ui.trabajador.value = existente.nombre;
            if (ui.editar) ui.editar.disabled = false;
            if (ui.eliminar) ui.eliminar.disabled = false;

            mostrarMensaje(
                "Este equipo ya tiene un trabajador asignado. Utiliza Editar para cambiarlo.",
                "warning"
            );

            return;
        }

        mostrarMensaje("Registrando la asignación…", "info");

        const sitio = await obtenerSitio();
        const lista = await obtenerListaTrabajadores();
        const campos = await obtenerCamposTrabajadores();

        validarCamposDeEscritura(campos);

        const resultado = await obtenerJSON(
            `${GRAPH}/sites/${sitio.id}/lists/${lista.id}/items`,
            {
                method: "POST",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    fields: {
                        [campos.equipo]: nombreEquipo,
                        [campos.trabajador]: nombre
                    }
                })
            }
        );

        const asignacion = {
            nombre,
            itemId: resultado.id,
            equipo: nombreEquipo
        };

        trabajadoresData.push(resultado);
        trabajadoresCargados = true;
        trabajadoresListos = true;

        actualizarAsignacionLocal(equipo, asignacion);

        if (ui.trabajador) {
            ui.trabajador.value = nombre;
            ui.trabajador.disabled = false;
        }

        if (ui.editar) ui.editar.disabled = false;
        if (ui.eliminar) ui.eliminar.disabled = false;

        mostrarMensaje("Trabajador registrado correctamente.", "success");

    } catch (error) {
        mostrarMensaje(
            mensajeErrorAmigable(error, "registrar al trabajador"),
            "error"
        );
    } finally {
        if (boton) {
            boton.textContent = "Registrar";
            boton.disabled = Boolean(obtenerTrabajadorEquipo(equipo));
        }
    }
}

/* ======================== EDITAR TRABAJADOR ======================== */

async function editarTrabajadorPorIP() {
    const equipo = equipoGestionSeleccionado;
    const nombre = ui.trabajador?.value?.trim();
    const asignacion = equipo && obtenerTrabajadorEquipo(equipo);

    if (!equipo || !asignacion?.itemId) {
        mostrarMensaje(
            "No se encontró la asignación. Busca nuevamente el equipo.",
            "warning"
        );
        return;
    }

    if (!nombre) {
        mostrarMensaje("Escribe el nombre del trabajador.", "warning");
        ui.trabajador?.focus();
        return;
    }

    const boton = ui.editar;

    try {
        if (boton) {
            boton.disabled = true;
            boton.textContent = "Guardando…";
        }

        mostrarMensaje("Guardando los cambios…", "info");

        const sitio = await obtenerSitio();
        const lista = await obtenerListaTrabajadores();
        const campos = await obtenerCamposTrabajadores();

        validarCamposDeEscritura(campos);

        await obtenerJSON(
            `${GRAPH}/sites/${sitio.id}/lists/${lista.id}/items/${asignacion.itemId}/fields`,
            {
                method: "PATCH",
                headers: {
                    "Content-Type": "application/json"
                },
                body: JSON.stringify({
                    [campos.trabajador]: nombre
                })
            }
        );

        asignacion.nombre = nombre;

        actualizarAsignacionLocal(equipo, asignacion);

        if (ui.trabajador) ui.trabajador.value = nombre;

        mostrarMensaje("Trabajador actualizado correctamente.", "success");

    } catch (error) {
        mostrarMensaje(
            mensajeErrorAmigable(error, "actualizar al trabajador"),
            "error"
        );
    } finally {
        if (boton) {
            boton.textContent = "Editar";
            boton.disabled = false;
        }
    }
}

/* ======================== ELIMINAR TRABAJADOR ======================== */

async function eliminarTrabajadorPorIP() {
    const equipo = equipoGestionSeleccionado;
    const asignacion = equipo && obtenerTrabajadorEquipo(equipo);

    if (!equipo || !asignacion?.itemId) {
        mostrarMensaje(
            "No se encontró la asignación. Busca nuevamente el equipo.",
            "warning"
        );
        return;
    }

    const nombreEquipo =
        equipo.Nombre ||
        equipo.DNSHostName ||
        equipo.IP ||
        "este equipo";

    if (!confirm(
        `¿Deseas eliminar la asignación de ${asignacion.nombre} para ${nombreEquipo}?`
    )) {
        return;
    }

    const boton = ui.eliminar;

    try {
        if (boton) {
            boton.disabled = true;
            boton.textContent = "Eliminando…";
        }

        mostrarMensaje("Eliminando la asignación…", "info");

        const sitio = await obtenerSitio();
        const lista = await obtenerListaTrabajadores();

        const respuesta = await graphFetch(
            `${GRAPH}/sites/${sitio.id}/lists/${lista.id}/items/${asignacion.itemId}`,
            { method: "DELETE" }
        );

        if (!respuesta.ok) {
            const detalle = await respuesta.text().catch(() => "");
            const error = new Error(
                `No se pudo eliminar la asignación: ${detalle}`
            );
            error.status = respuesta.status;
            throw error;
        }

        trabajadoresData = trabajadoresData.filter(
            item => String(item.id) !== String(asignacion.itemId)
        );

        actualizarAsignacionLocal(equipo, null);

        if (ui.trabajador) ui.trabajador.value = "";
        if (ui.registrar) ui.registrar.disabled = false;
        if (ui.editar) ui.editar.disabled = true;

        mostrarMensaje("Asignación eliminada correctamente.", "success");

    } catch (error) {
        mostrarMensaje(
            mensajeErrorAmigable(error, "eliminar la asignación"),
            "error"
        );
    } finally {
        if (boton) {
            boton.textContent = "Eliminar";
            boton.disabled = !obtenerTrabajadorEquipo(equipo);
        }
    }
}

/* ======================== PERMISOS DEL INVENTARIO ======================== */

function usuarioPuedeVerInventario() {
    return String((window.alferzaPermisos || {}).grupo || "")
        .trim()
        .toUpperCase() === "TI";
}

async function verificarPermisosInventario() {
    if (window.alferzaAccessReady?.then) {
        try {
            await window.alferzaAccessReady;
        } catch {
            return false;
        }
    }

    return usuarioPuedeVerInventario();
}

function configurarAccesoInventario() {
    const inventario = $("inventarioTISection");
    if (!inventario) return;

    const actualizar = () => {
        inventario.style.display = usuarioPuedeVerInventario()
            ? ""
            : "none";
    };

    inventario.style.display = "none";

    window.addEventListener("alferza:access-granted", actualizar);

    if (window.alferzaAccessReady?.then) {
        window.alferzaAccessReady
            .then(actualizar)
            .catch(() => {
                inventario.style.display = "none";
            });
    } else {
        actualizar();
    }
}

/* ======================== CARGA DEL DASHBOARD ======================== */

function mostrarBotonLogin() {
    mostrarCarga();

    if (!ui.loading || $("btnLoginMs")) return;

    const boton = document.createElement("button");

    boton.id = "btnLoginMs";
    boton.type = "button";
    boton.className = "loading-btn";
    boton.textContent = "Iniciar sesión con Microsoft";

    boton.addEventListener("click", async () => {
        boton.disabled = true;

        try {
            if (await iniciarSesion(true)) {
                boton.remove();
                ocultarCarga();
                await iniciarCargas();
            }
        } catch (error) {
            console.error("[ALFERZA] Error al iniciar sesión:", error);

            if (ui.loadingMessage) {
                ui.loadingMessage.textContent =
                    "No se pudo iniciar sesión. Inténtalo nuevamente.";
            }
        } finally {
            boton.disabled = false;
        }
    });

    ui.loading
        .querySelector(".loading-content, .loading-card")
        ?.appendChild(boton);
}

async function iniciarCargas() {
    ultimaActualizacion = Date.now();

    await cargarServicios();

    if (await verificarPermisosInventario()) {
        await cargarEquipos();

        if (equiposData.length) {
            try {
                await obtenerTrabajadores();
            } catch (error) {
                console.error("[ALFERZA] Error al cargar trabajadores:", error);
            }
        }
    }
}

async function cargarDashboard() {
    if (!msalInfraestructura) {
        mostrarCarga("No se pudo inicializar Microsoft Authentication Library.");
        return;
    }

    const cache = leerCache("equipos");

    if (
        msalInfraestructura.getAllAccounts().length &&
        Array.isArray(cache?.datos) &&
        cache.datos.length
    ) {
        etagEquipos = cache.etag || null;
        aplicarDatosEquipos(cache.datos);
    } else {
        mostrarCarga();
    }

    try {
        if (!await iniciarSesion(false)) {
            mostrarBotonLogin();
            return;
        }

        ocultarCarga();
        await iniciarCargas();

    } catch (error) {
        console.error("[ALFERZA] Error al iniciar el dashboard:", error);
        mostrarBotonLogin();
    }
}

async function actualizarDashboard() {
    if (document.hidden || !tokenActual) return;

    try {
        if (!await obtenerTokenSilencioso()) return;

        ultimaActualizacion = Date.now();

        await cargarServicios();

        if (await verificarPermisosInventario()) {
            await cargarEquipos();
        }

    } catch (error) {
        console.error("[ALFERZA] Error al actualizar el dashboard:", error);
    }
}

/* ======================== EVENTOS ======================== */

function configurarEventos() {
    ui.abrirGestion?.addEventListener(
        "click",
        abrirGestionTrabajadores
    );

    ui.buscarEquipo?.addEventListener(
        "click",
        buscarEquipoPorIP
    );

    ui.registrar?.addEventListener(
        "click",
        registrarTrabajadorPorIP
    );

    ui.editar?.addEventListener(
        "click",
        editarTrabajadorPorIP
    );

    ui.eliminar?.addEventListener(
        "click",
        eliminarTrabajadorPorIP
    );

    ui.cerrar?.addEventListener(
        "click",
        cerrarGestionTrabajadores
    );

    ui.cancelar?.addEventListener(
        "click",
        cerrarGestionTrabajadores
    );

    ui.modal?.addEventListener("click", event => {
        if (event.target === ui.modal) {
            cerrarGestionTrabajadores();
        }
    });

    ui.busquedaEquipo?.addEventListener("keydown", event => {
        if (event.key === "Enter") {
            event.preventDefault();
            buscarEquipoPorIP();
        }

        if (event.key === "Escape") {
            cerrarGestionTrabajadores();
        }
    });

    ui.trabajador?.addEventListener("keydown", event => {
        if (event.key === "Enter") {
            event.preventDefault();

            if (!ui.registrar?.disabled) {
                registrarTrabajadorPorIP();
            } else if (!ui.editar?.disabled) {
                editarTrabajadorPorIP();
            }
        }

        if (event.key === "Escape") {
            cerrarGestionTrabajadores();
        }
    });

    let temporizador;

    ui.busqueda?.addEventListener("input", () => {
        clearTimeout(temporizador);

        temporizador = setTimeout(
            () => aplicarFiltrosEquipos(),
            200
        );
    });

    ui.filtroEstado?.addEventListener(
        "change",
        () => aplicarFiltrosEquipos()
    );

    ui.filtroArea?.addEventListener(
        "change",
        () => aplicarFiltrosEquipos()
    );
}

/* ======================== INICIALIZACIÓN ======================== */

function iniciarInfraestructura() {
    configurarEstilosMensajes();
    configurarAccesoInventario();
    configurarEventos();

    cargarDashboard();

    setInterval(
        actualizarDashboard,
        CONFIG.intervaloActualizacion
    );

    document.addEventListener("visibilitychange", () => {
        if (
            !document.hidden &&
            Date.now() - ultimaActualizacion > 30000
        ) {
            actualizarDashboard();
        }
    });
}

if (document.readyState === "loading") {
    document.addEventListener(
        "DOMContentLoaded",
        iniciarInfraestructura
    );
} else {
    iniciarInfraestructura();
}
