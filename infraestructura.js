/* =========================================================
   CONTROL DE ACCESO AL INVENTARIO TI
========================================================= */

(function controlarAccesoInventario() {
    const inventario = document.getElementById("inventarioTISection");
    if (!inventario) return;

    inventario.style.display = "none";

    function actualizarVisibilidadInventario() {
        const permisos = window.alferzaPermisos || {};
        const grupo = String(permisos.grupo || "").trim().toUpperCase();
        inventario.style.display = grupo === "TI" ? "" : "none";
    }

    window.addEventListener("alferza:access-granted", actualizarVisibilidadInventario);

    if (window.alferzaAccessReady &&
        typeof window.alferzaAccessReady.then === "function") {
        window.alferzaAccessReady
            .then(actualizarVisibilidadInventario)
            .catch(() => {
                inventario.style.display = "none";
            });
    } else {
        actualizarVisibilidadInventario();
    }
})();

/* =========================================================
   CONFIGURACIÓN
========================================================= */

const msalConfigInfraestructura = {
    auth: {
        clientId: "5d98417c-74a7-4fab-8f2c-41ac127be696",
        authority: "https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1",
        redirectUri: "https://alferzati.github.io/live-office/blank.html"
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
    }
};

const msalInstanceInfraestructura =
    typeof msal !== "undefined"
        ? new msal.PublicClientApplication(msalConfigInfraestructura)
        : null;

const SHAREPOINT_HOSTNAME = "alferzaholding-my.sharepoint.com";
const SHAREPOINT_SITE = "/personal/soporte1_alferza_pe";
const SHAREPOINT_LIST_SERVICIOS = "MonitoreoTI";
const SHAREPOINT_LIST_TRABAJADORES = "TrabajadoresEquipos";
const RUTA_EQUIPOS = "Procedimientos T.I/equipos.json";

const EQUIPOS_POR_PAGINA = 40;
const INTERVALO_ACTUALIZACION = 120000;

const SCOPES_GRAPH = [
    "User.Read",
    "Sites.Read.All",
    "Sites.ReadWrite.All"
];

const CACHE_PREFIJO = "alferza_ti_";

/* =========================================================
   VARIABLES
========================================================= */

let cuentaActual = null;
let tokenActual = null;
let sitioSharePoint = null;
let driveSharePoint = null;
let listaTrabajadores = null;
let columnasTrabajadores = null;

let trabajadoresData = [];
let trabajadoresPorEquipo = {};
let trabajadoresCargados = false;
let trabajadoresListos = false;
let cargandoTrabajadores = false;

let equiposData = [];
let equiposFiltrados = [];
let paginaActual = 1;
let equipoSeleccionado = null;
let equipoGestionSeleccionado = null;

let etagEquipos = null;
let promesaSitio = null;
let promesaDrive = null;
let promesaTrabajadores = null;
let cargandoEquipos = false;
let ultimaActualizacion = 0;
let ubicacionEquipos = null;

/* =========================================================
   ELEMENTOS HTML
========================================================= */

const loadingOverlay = document.getElementById("loadingOverlay");
const loadingMessage = document.getElementById("loadingMessage");

const serviciosGrid = document.getElementById("serviciosGrid");
const estadoGeneral = document.getElementById("estadoGeneral");

const equiposTableBody = document.getElementById("equiposTableBody");
const equiposEmpty = document.getElementById("equiposEmpty");
const equiposResultados = document.getElementById("equiposResultados");

const totalEquipos = document.getElementById("totalEquipos");
const equiposConectados = document.getElementById("equiposConectados");
const equiposDesconectados = document.getElementById("equiposDesconectados");

const filtroEstadoEquipo = document.getElementById("filtroEstadoEquipo");
const filtroAreaEquipo = document.getElementById("filtroAreaEquipo");
const buscarEquipo = document.getElementById("buscarEquipo");
const equiposEstado = document.getElementById("equiposEstado");
const lastUpdate = document.getElementById("lastUpdate");

const gestionTrabajadorModal = document.getElementById("gestionTrabajadorModal");
const gestionIP = document.getElementById("gestionIP");
const buscarEquipoIP = document.getElementById("buscarEquipoIP");
const equipoEncontrado = document.getElementById("equipoEncontrado");

const gestionEquipoNombre = document.getElementById("gestionEquipoNombre");
const gestionEquipoDetalle = document.getElementById("gestionEquipoDetalle");
const gestionTrabajador = document.getElementById("gestionTrabajador");
const gestionTrabajadorMensaje = document.getElementById("gestionTrabajadorMensaje");

const btnGestionarTrabajadores = document.getElementById("btnGestionarTrabajadores");
const btnRegistrarTrabajador = document.getElementById("btnRegistrarTrabajador");
const btnEditarTrabajador = document.getElementById("btnEditarTrabajador");
const btnEliminarTrabajador = document.getElementById("btnEliminarTrabajador");

const cerrarGestionTrabajadorModal =
    document.getElementById("cerrarGestionTrabajadorModal");

const cancelarGestionTrabajador =
    document.getElementById("cancelarGestionTrabajador");

/* =========================================================
   UTILIDADES
========================================================= */

function escaparHTML(valor) {
    if (valor === null || valor === undefined) return "";

    return String(valor)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function mostrarLoading() {}

function ocultarLoading() {
    if (loadingOverlay) {
        loadingOverlay.style.display = "none";
    }
}

function obtenerFechaActual() {
    return new Date().toLocaleString("es-PE", {
        dateStyle: "short",
        timeStyle: "medium"
    });
}

function normalizarTexto(valor) {
    return String(valor || "").trim().toLowerCase();
}

function sinTildes(valor) {
    return String(valor === null || valor === undefined ? "" : valor)
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .trim();
}

/*
 * Diagnósticos visuales desactivados.
 * Los errores no generan paneles emergentes.
 */

function mostrarDiagnostico() {}

/* =========================================================
   CACHÉ
========================================================= */

function leerCache(clave) {
    try {
        const valor = localStorage.getItem(CACHE_PREFIJO + clave);
        return valor ? JSON.parse(valor) : null;
    } catch (error) {
        return null;
    }
}

function guardarCache(clave, valor) {
    try {
        localStorage.setItem(
            CACHE_PREFIJO + clave,
            JSON.stringify(valor)
        );
    } catch (error) {
        // No se muestra ningún mensaje.
    }
}

/* =========================================================
   AUTENTICACIÓN MICROSOFT
========================================================= */

async function obtenerTokenSilencioso() {
    if (!msalInstanceInfraestructura) return null;

    const cuentas = msalInstanceInfraestructura.getAllAccounts();

    if (!cuentas.length) return null;

    cuentaActual = cuentas[0];

    try {
        const respuesta =
            await msalInstanceInfraestructura.acquireTokenSilent({
                scopes: SCOPES_GRAPH,
                account: cuentaActual
            });

        tokenActual = respuesta.accessToken;
        return tokenActual;
    } catch (error) {
        mostrarDiagnostico();
        return null;
    }
}

async function iniciarSesion(interactivo) {
    const tokenSilencioso = await obtenerTokenSilencioso();

    if (tokenSilencioso || !interactivo) {
        return tokenSilencioso;
    }

    const respuesta = cuentaActual
        ? await msalInstanceInfraestructura.acquireTokenPopup({
            scopes: SCOPES_GRAPH,
            account: cuentaActual
        })
        : await msalInstanceInfraestructura.loginPopup({
            scopes: SCOPES_GRAPH
        });

    cuentaActual = respuesta.account;
    tokenActual = respuesta.accessToken ||
        await obtenerTokenSilencioso();

    return tokenActual;
}

/* =========================================================
   MICROSOFT GRAPH
========================================================= */

function registrarFallo(status, url) {
    // Registro visual desactivado.
}

async function graphFetch(url, opciones, intento) {
    intento = intento || 0;

    const { timeoutMs = 30000, ...resto } = opciones || {};

    const controlador = new AbortController();
    const temporizador = setTimeout(
        () => controlador.abort(),
        timeoutMs
    );

    try {
        const respuesta = await fetch(url, {
            ...resto,
            signal: controlador.signal,
            headers: {
                Authorization: `Bearer ${tokenActual}`,
                ...(resto.headers || {})
            }
        });

        if (!respuesta.ok) {
            registrarFallo(respuesta.status, url);
        }

        if (
            (respuesta.status === 429 || respuesta.status === 503) &&
            intento < 3
        ) {
            const espera = (
                parseInt(respuesta.headers.get("Retry-After"), 10) ||
                2 ** intento
            ) * 1000;

            await new Promise(resolve =>
                setTimeout(resolve, Math.min(espera, 10000))
            );

            return graphFetch(url, opciones, intento + 1);
        }

        if (
            respuesta.status === 401 &&
            intento < 1 &&
            await obtenerTokenSilencioso()
        ) {
            return graphFetch(url, opciones, intento + 1);
        }

        return respuesta;
    } finally {
        clearTimeout(temporizador);
    }
}

/* =========================================================
   SHAREPOINT: SITIO Y DRIVE
========================================================= */

function obtenerSitioSharePoint() {
    if (!promesaSitio) {
        promesaSitio = (async () => {
            const clave = "sitio:" + SHAREPOINT_SITE;
            const guardado = leerCache(clave);

            if (guardado && guardado.id) {
                return guardado;
            }

            const respuesta = await graphFetch(
                `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOSTNAME}:${SHAREPOINT_SITE}?$select=id`
            );

            if (!respuesta.ok) {
                throw new Error("No se pudo obtener el sitio SharePoint.");
            }

            const sitio = await respuesta.json();

            guardarCache(clave, { id: sitio.id });

            return { id: sitio.id };
        })().catch(error => {
            promesaSitio = null;
            throw error;
        });
    }

    return promesaSitio;
}

function obtenerDriveSharePoint() {
    if (!promesaDrive) {
        promesaDrive = Promise.resolve({
            id: "b!1l0EwOKXpU24cUGYaNjPRiTQWIyIIChCkDDL197LcoOFOD7ThItqTIJ_Z1KoYrrS"
        });
    }

    return promesaDrive;
}

/* =========================================================
   INVENTARIO: EQUIPOS.JSON
========================================================= */

async function buscarEquiposJson(raiz) {
    const nombre = RUTA_EQUIPOS.split("/").pop();

    const respuesta = await graphFetch(
        `${raiz}/root/search(q='${encodeURIComponent(nombre)}')?$select=id,name,eTag,parentReference&$top=25`
    );

    if (respuesta.ok) {
        const elementos = (await respuesta.json()).value || [];

        const encontrados = elementos.filter(
            item => normalizarTexto(item.name) === normalizarTexto(nombre)
        );

        if (encontrados.length > 0) {
            const elemento = encontrados[0];

            return {
                meta: `${raiz}/items/${elemento.id}`,
                contenido: `${raiz}/items/${elemento.id}/content`,
                etag: elemento.eTag || null
            };
        }
    }

    throw new Error("No se encontró el archivo de inventario.");
}

async function leerJsonEquipos(respuesta) {
    const buffer = await respuesta.arrayBuffer();
    const bytes = new Uint8Array(buffer);

    let codificacion = "utf-8";

    if (bytes[0] === 0xFF && bytes[1] === 0xFE) {
        codificacion = "utf-16le";
    } else if (bytes[0] === 0xFE && bytes[1] === 0xFF) {
        codificacion = "utf-16be";
    }

    const texto = new TextDecoder(codificacion)
        .decode(buffer)
        .replace(/^\uFEFF/, "");

    let datos;

    try {
        datos = JSON.parse(texto);
    } catch (error) {
        throw new Error("El archivo de inventario no contiene JSON válido.");
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
            if (Array.isArray(datos[clave])) {
                return datos[clave];
            }
        }

        if (datos.Nombre) return [datos];
    }

    throw new Error("Formato de inventario no válido.");
}

async function obtenerEquiposDesdeSharePoint() {
    const drive = await obtenerDriveSharePoint();
    const raiz = `https://graph.microsoft.com/v1.0/drives/${drive.id}`;

    const ruta = RUTA_EQUIPOS
        .split("/")
        .map(parte => encodeURIComponent(parte))
        .join("/");

    const candidatos = [];

    if (ubicacionEquipos) {
        candidatos.push(ubicacionEquipos);
    }

    candidatos.push({
        meta: `${raiz}/root:/${ruta}`,
        contenido: `${raiz}/root:/${ruta}:/content`
    });

    let ubicacion = null;
    let etag = null;

    for (const candidato of candidatos) {
        const respuesta = await graphFetch(
            `${candidato.meta}?$select=eTag`
        );

        if (respuesta.ok) {
            etag = (await respuesta.json()).eTag || null;
            ubicacion = candidato;
            break;
        }

        if (respuesta.status !== 404) {
            throw new Error("No se pudo consultar el inventario.");
        }
    }

    if (!ubicacion) {
        ubicacion = await buscarEquiposJson(raiz);
        etag = ubicacion.etag;
    }

    ubicacionEquipos = ubicacion;

    if (etag && etag === etagEquipos && equiposData.length > 0) {
        return null;
    }

    const respuesta = await graphFetch(
        ubicacion.contenido,
        { timeoutMs: 90000 }
    );

    if (!respuesta.ok) {
        throw new Error("No se pudo descargar el inventario.");
    }

    const datos = await leerJsonEquipos(respuesta);

    etagEquipos = etag;

    return datos;
}

/* =========================================================
   SHAREPOINT: TRABAJADORES Y EQUIPOS
========================================================= */

async function obtenerListaTrabajadores() {
    if (listaTrabajadores) return listaTrabajadores;

    const sitio = await obtenerSitioSharePoint();

    const respuesta = await graphFetch(
        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists`
    );

    if (!respuesta.ok) {
        throw new Error("No se pudieron consultar las listas.");
    }

    const listas = (await respuesta.json()).value || [];

    listaTrabajadores = listas.find(
        lista =>
            normalizarTexto(lista.displayName) ===
            normalizarTexto(SHAREPOINT_LIST_TRABAJADORES)
    );

    if (!listaTrabajadores) {
        throw new Error("No se encontró la lista de trabajadores.");
    }

    return listaTrabajadores;
}

async function obtenerColumnasTrabajadores() {
    if (columnasTrabajadores) return columnasTrabajadores;

    const lista = await obtenerListaTrabajadores();
    const sitio = await obtenerSitioSharePoint();

    const respuesta = await graphFetch(
        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/columns`
    );

    if (!respuesta.ok) {
        throw new Error("No se pudieron consultar las columnas.");
    }

    columnasTrabajadores = (await respuesta.json()).value || [];

    return columnasTrabajadores;
}

function obtenerNombreInternoColumna(columnas, nombreVisible) {
    const columna = columnas.find(
        item =>
            normalizarTexto(item.displayName) ===
            normalizarTexto(nombreVisible)
    );

    return columna ? columna.name : null;
}

function obtenerTrabajadores() {
    if (trabajadoresCargados) {
        return Promise.resolve(trabajadoresPorEquipo);
    }

    if (!promesaTrabajadores) {
        promesaTrabajadores = cargarTrabajadoresDesdeSharePoint()
            .finally(() => {
                promesaTrabajadores = null;
            });
    }

    return promesaTrabajadores;
}

async function cargarTrabajadoresDesdeSharePoint() {
    try {
        const sitio = await obtenerSitioSharePoint();
        const lista = await obtenerListaTrabajadores();

        let url =
            `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items` +
            `?expand=fields($select=Equipo1,NombreTrabajador)&$top=999`;

        const elementos = [];

        while (url) {
            const respuesta = await graphFetch(url);

            if (!respuesta.ok) {
                throw new Error("Error cargando asignaciones.");
            }

            const datos = await respuesta.json();

            elementos.push(...(datos.value || []));
            url = datos["@odata.nextLink"] || null;
        }

        trabajadoresData = elementos;
        trabajadoresPorEquipo = {};

        elementos.forEach(elemento => {
            const fields = elemento.fields || {};

            if (fields.Equipo1 && fields.NombreTrabajador) {
                trabajadoresPorEquipo[
                    normalizarTexto(fields.Equipo1)
                ] = {
                    nombre: fields.NombreTrabajador,
                    itemId: elemento.id,
                    equipo: fields.Equipo1
                };
            }
        });

        if (
            elementos.length > 0 &&
            Object.keys(trabajadoresPorEquipo).length === 0
        ) {
            await cargarTrabajadoresConColumnas();
        }

        trabajadoresCargados = true;
        trabajadoresListos = true;

        if (equiposData.length > 0) {
            renderizarTablaEquipos();
        }

        return trabajadoresPorEquipo;
    } catch (error) {
        trabajadoresListos = true;

        if (equiposData.length > 0) {
            renderizarTablaEquipos();
        }

        return trabajadoresPorEquipo;
    }
}

async function cargarTrabajadoresConColumnas() {
    const columnas = await obtenerColumnasTrabajadores();

    const campoEquipo = obtenerNombreInternoColumna(
        columnas,
        "Equipo1"
    );

    const campoTrabajador = obtenerNombreInternoColumna(
        columnas,
        "Trabajador"
    );

    if (!campoEquipo || !campoTrabajador) return;

    trabajadoresPorEquipo = {};

    trabajadoresData.forEach(elemento => {
        const fields = elemento.fields || {};
        const equipo = fields[campoEquipo];
        const trabajador = fields[campoTrabajador];

        if (equipo && trabajador) {
            trabajadoresPorEquipo[normalizarTexto(equipo)] = {
                nombre: trabajador,
                itemId: elemento.id,
                equipo
            };
        }
    });
}

function obtenerTrabajadorEquipo(nombreEquipo) {
    return trabajadoresPorEquipo[normalizarTexto(nombreEquipo)] || null;
}

/* =========================================================
   MONITOREO DE SERVICIOS
========================================================= */

async function obtenerMonitoreoTI() {
    const sitio = await obtenerSitioSharePoint();

    const url =
        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/` +
        `${encodeURIComponent(SHAREPOINT_LIST_SERVICIOS)}/items?expand=fields&$top=999`;

    const respuesta = await graphFetch(url);

    if (!respuesta.ok) {
        throw new Error("No se pudo consultar el monitoreo.");
    }

    return (await respuesta.json()).value || [];
}

function obtenerCampo(fields, nombres, defecto) {
    const mapa = {};

    Object.keys(fields).forEach(clave => {
        mapa[sinTildes(clave)] = fields[clave];
    });

    for (const nombre of nombres) {
        const valor = mapa[sinTildes(nombre)];

        if (
            valor !== undefined &&
            valor !== null &&
            String(valor).trim() !== ""
        ) {
            return valor;
        }
    }

    return defecto;
}

function adivinarDescripcion(fields, excluir) {
    const claves = excluir.map(sinTildes);
    let mejor = "";

    Object.entries(fields).forEach(([clave, valor]) => {
        if (typeof valor !== "string") return;
        if (clave.startsWith("@") || clave.startsWith("_")) return;
        if (claves.includes(sinTildes(valor))) return;
        if (/^\d{4}-\d{2}-\d{2}T/.test(valor)) return;

        if (valor.length > mejor.length) {
            mejor = valor;
        }
    });

    return mejor.length >= 15 ? mejor : "";
}

function formatearRevision(valor) {
    if (!valor) return "--";

    const texto = String(valor).trim();

    if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(texto)) {
        return texto;
    }

    const fecha = new Date(texto);

    if (isNaN(fecha.getTime())) return texto;

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
    if (!serviciosGrid) return;

    if (!items || items.length === 0) {
        serviciosGrid.innerHTML = `
            <div class="service-card service-loading">
                <span>No hay servicios registrados.</span>
            </div>
        `;

        estadoGeneral.textContent = "● SIN DATOS";
        estadoGeneral.className = "section-status warning";

        return;
    }

    let serviciosOnline = 0;

    serviciosGrid.innerHTML = items.map(item => {
        const fields = item.fields || {};

        const nombre = obtenerCampo(
            fields,
            ["Servicio", "Nombre", "Title", "LinkTitle"],
            "Servicio"
        );

        const estado = obtenerCampo(
            fields,
            ["Estado", "Status"],
            "Sin estado"
        );

        const descripcion =
            obtenerCampo(
                fields,
                [
                    "Descripcion",
                    "Description",
                    "Detalle",
                    "Mensaje",
                    "Observacion",
                    "Comentario"
                ],
                ""
            ) ||
            adivinarDescripcion(fields, [nombre, estado]) ||
            "Sin descripción";

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
                    "Ultima_revision",
                    "Fecha",
                    "FechaRevision"
                ],
                null
            )
        );

        const estadoNormalizado = sinTildes(estado);
        let claseEstado = "neutral";

        if (
            estadoNormalizado.includes("operativo") ||
            estadoNormalizado.includes("online") ||
            estadoNormalizado.includes("conectado")
        ) {
            claseEstado = "online";
            serviciosOnline++;
        } else if (
            estadoNormalizado.includes("mantenimiento") ||
            estadoNormalizado.includes("advertencia") ||
            estadoNormalizado.includes("warning") ||
            estadoNormalizado.includes("lento")
        ) {
            claseEstado = "warning";
        } else if (
            estadoNormalizado.includes("caido") ||
            estadoNormalizado.includes("offline") ||
            estadoNormalizado.includes("desconectado") ||
            estadoNormalizado.includes("error")
        ) {
            claseEstado = "offline";
        }

        return `
            <article class="service-card ${claseEstado}">
                <div class="service-main">
                    <div class="service-icon">🖥️</div>
                    <div class="service-info">
                        <div class="service-heading">
                            <h3>${escaparHTML(nombre)}</h3>
                            <span class="service-status ${claseEstado}">
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
            </article>
        `;
    }).join("");

    if (serviciosOnline === items.length) {
        estadoGeneral.textContent = "● OPERATIVO";
        estadoGeneral.className = "section-status online";
    } else if (serviciosOnline > 0) {
        estadoGeneral.textContent = "● REVISAR";
        estadoGeneral.className = "section-status warning";
    } else {
        estadoGeneral.textContent = "● NO DISPONIBLE";
        estadoGeneral.className = "section-status offline";
    }
}

/* =========================================================
   TABLA DE INVENTARIO
========================================================= */

function actualizarResumenEquipos() {
    const total = equiposData.length;

    const conectados = equiposData.filter(
        equipo => normalizarTexto(equipo.Estado) === "conectado"
    ).length;

    const desconectados = equiposData.filter(
        equipo => normalizarTexto(equipo.Estado) === "desconectado"
    ).length;

    totalEquipos.textContent = total;
    equiposConectados.textContent = conectados;
    equiposDesconectados.textContent = desconectados;
}

function actualizarFiltroAreas() {
    const areas = [
        ...new Set(equiposData.map(equipo => equipo.Area).filter(Boolean))
    ].sort((a, b) => String(a).localeCompare(String(b), "es"));

    filtroAreaEquipo.innerHTML =
        `<option value="todos">Todas las áreas</option>`;

    areas.forEach(area => {
        const opcion = document.createElement("option");
        opcion.value = area;
        opcion.textContent = area;
        filtroAreaEquipo.appendChild(opcion);
    });
}

function aplicarFiltrosEquipos(conservarPagina) {
    const texto = normalizarTexto(buscarEquipo.value);
    const estado = filtroEstadoEquipo.value;
    const area = filtroAreaEquipo.value;

    equiposFiltrados = equiposData.filter(equipo => {
        const coincideTexto =
            !texto ||
            normalizarTexto(equipo.Nombre).includes(texto) ||
            normalizarTexto(equipo.IP).includes(texto) ||
            normalizarTexto(equipo.DNSHostName).includes(texto);

        const coincideEstado =
            estado === "todos" ||
            String(equipo.Estado || "") === estado;

        const coincideArea =
            area === "todos" ||
            String(equipo.Area || "") === area;

        return coincideTexto && coincideEstado && coincideArea;
    });

    if (conservarPagina !== true) {
        paginaActual = 1;
    }

    renderizarTablaEquipos();
}

function obtenerClaseEstado(estado) {
    const normalizado = normalizarTexto(estado);

    if (normalizado === "conectado") return "online";
    if (normalizado === "desconectado") return "offline";

    return "neutral";
}

function renderizarTablaEquipos() {
    const totalResultados = equiposFiltrados.length;

    const totalPaginas = Math.max(
        1,
        Math.ceil(totalResultados / EQUIPOS_POR_PAGINA)
    );

    if (paginaActual > totalPaginas) {
        paginaActual = totalPaginas;
    }

    const inicio = (paginaActual - 1) * EQUIPOS_POR_PAGINA;
    const equiposPagina = equiposFiltrados.slice(
        inicio,
        inicio + EQUIPOS_POR_PAGINA
    );

    if (equiposResultados) {
        equiposResultados.textContent =
            totalResultados === 0
                ? "0 equipos encontrados"
                : `Equipos encontrados: ${totalResultados}`;
    }

    if (totalResultados === 0) {
        equiposTableBody.innerHTML = "";
        equiposEmpty.style.display = "flex";
        renderizarPaginacion(0);
        return;
    }

    equiposEmpty.style.display = "none";

    equiposTableBody.innerHTML = equiposPagina.map(equipo => {
        const estado = equipo.Estado || "No verificable";
        const claseEstado = obtenerClaseEstado(estado);
        const trabajador = obtenerTrabajadorEquipo(equipo.Nombre);

        const trabajadorHTML = trabajador
            ? `<span class="trabajador-name" title="${escaparHTML(trabajador.nombre)}">${escaparHTML(trabajador.nombre)}</span>`
            : !trabajadoresListos
                ? `<span style="color:#98a2b3;font-size:12px;">Cargando…</span>`
                : `<span style="color:#98a2b3;font-size:12px;">Sin asignar</span>`;

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
                    <span class="equipo-status ${claseEstado}">
                        <span class="equipo-status-dot"></span>
                        ${escaparHTML(estado)}
                    </span>
                </td>

                <td class="trabajador-cell">${trabajadorHTML}</td>

                <td>
                    <span class="equipo-ip">${escaparHTML(equipo.IP || "--")}</span>
                </td>

                <td>
                    <div class="equipo-so">
                        ${escaparHTML(equipo.SistemaOperativo || "--")}
                        ${equipo.VersionSO
                            ? `<br><small>${escaparHTML(equipo.VersionSO)}</small>`
                            : ""}
                    </div>
                </td>

                <td>
                    <span class="equipo-area" title="${escaparHTML(equipo.Area)}">
                        ${escaparHTML(equipo.Area || "--")}
                    </span>
                </td>

                <td>
                    <span class="equipo-date">${escaparHTML(equipo.UltimoRegistroAD || "--")}</span>
                </td>
            </tr>
        `;
    }).join("");

    renderizarPaginacion(totalPaginas);
}

function renderizarPaginacion(totalPaginas) {
    const existente = document.querySelector(".equipos-pagination");

    if (existente) existente.remove();
    if (totalPaginas <= 1) return;

    const contenedor = document.createElement("div");
    contenedor.className = "equipos-pagination";

    const anterior = document.createElement("button");
    anterior.type = "button";
    anterior.className = "pagination-btn pagination-prev";
    anterior.innerHTML = "‹";
    anterior.disabled = paginaActual === 1;

    anterior.addEventListener("click", () => {
        if (paginaActual > 1) {
            paginaActual--;
            renderizarTablaEquipos();
        }
    });

    contenedor.appendChild(anterior);

    generarPaginas(paginaActual, totalPaginas).forEach(pagina => {
        if (pagina === "...") {
            const puntos = document.createElement("span");
            puntos.className = "pagination-ellipsis";
            puntos.textContent = "...";
            contenedor.appendChild(puntos);
            return;
        }

        const boton = document.createElement("button");
        boton.type = "button";
        boton.className = "pagination-btn";
        boton.textContent = pagina;

        if (pagina === paginaActual) {
            boton.classList.add("active");
        }

        boton.addEventListener("click", () => {
            paginaActual = pagina;
            renderizarTablaEquipos();
        });

        contenedor.appendChild(boton);
    });

    const siguiente = document.createElement("button");
    siguiente.type = "button";
    siguiente.className = "pagination-btn pagination-next";
    siguiente.innerHTML = "›";
    siguiente.disabled = paginaActual === totalPaginas;

    siguiente.addEventListener("click", () => {
        if (paginaActual < totalPaginas) {
            paginaActual++;
            renderizarTablaEquipos();
        }
    });

    contenedor.appendChild(siguiente);

    const tableContainer = document.querySelector(
        ".equipos-table-container"
    );

    if (tableContainer) {
        tableContainer.parentNode.insertBefore(
            contenedor,
            tableContainer.nextSibling
        );
    }
}

function generarPaginas(actual, total) {
    if (total <= 7) {
        return Array.from({ length: total }, (_, indice) => indice + 1);
    }

    const paginas = [1];

    if (actual > 4) paginas.push("...");

    const inicio = Math.max(2, actual - 1);
    const fin = Math.min(total - 1, actual + 1);

    for (let pagina = inicio; pagina <= fin; pagina++) {
        paginas.push(pagina);
    }

    if (actual < total - 3) paginas.push("...");

    paginas.push(total);

    return paginas;
}

function aplicarDatosEquipos(equipos) {
    equiposData = equipos;

    const areaPrevia = filtroAreaEquipo.value;

    actualizarResumenEquipos();
    actualizarFiltroAreas();

    if ([...filtroAreaEquipo.options].some(
        opcion => opcion.value === areaPrevia
    )) {
        filtroAreaEquipo.value = areaPrevia;
    }

    aplicarFiltrosEquipos(true);
}

/* =========================================================
   GESTIÓN DE TRABAJADORES
   Mensajes visuales desactivados
========================================================= */

function mostrarMensajeGestion() {
    if (gestionTrabajadorMensaje) {
        gestionTrabajadorMensaje.style.display = "none";
        gestionTrabajadorMensaje.textContent = "";
    }
}

function abrirGestionTrabajadores() {
    if (!gestionTrabajadorModal) return;

    equipoGestionSeleccionado = null;

    gestionIP.value = "";
    gestionTrabajador.value = "";
    gestionTrabajador.disabled = true;

    equipoEncontrado.style.display = "none";
    mostrarMensajeGestion();

    btnRegistrarTrabajador.disabled = true;
    btnEditarTrabajador.disabled = true;
    btnEliminarTrabajador.disabled = true;

    gestionTrabajadorModal.style.display = "flex";

    setTimeout(() => gestionIP.focus(), 100);
}

function cerrarGestionTrabajadores() {
    gestionTrabajadorModal.style.display = "none";
    equipoGestionSeleccionado = null;

    gestionIP.value = "";
    gestionTrabajador.value = "";
    gestionTrabajador.disabled = true;

    equipoEncontrado.style.display = "none";
    mostrarMensajeGestion();

    btnRegistrarTrabajador.disabled = true;
    btnEditarTrabajador.disabled = true;
    btnEliminarTrabajador.disabled = true;
}

async function buscarEquipoPorIP() {
    const ip = gestionIP.value.trim();

    if (!ip) return;

    mostrarMensajeGestion();

    equipoEncontrado.style.display = "none";
    gestionTrabajador.disabled = true;

    btnRegistrarTrabajador.disabled = true;
    btnEditarTrabajador.disabled = true;
    btnEliminarTrabajador.disabled = true;

    equipoGestionSeleccionado = null;

    buscarEquipoIP.disabled = true;
    buscarEquipoIP.textContent = "Buscando...";

    try {
        if (!equiposData || equiposData.length === 0) {
            await cargarEquipos();
        }

        const equipo = equiposData.find(
            item => normalizarTexto(item.IP) === normalizarTexto(ip)
        );

        if (!equipo) return;

        equipoGestionSeleccionado = equipo;

        gestionEquipoNombre.textContent = equipo.Nombre || "Sin nombre";

        gestionEquipoDetalle.textContent =
            `${equipo.IP || ip} · ` +
            `${equipo.DNSHostName || "Sin DNS"} · ` +
            `${equipo.Area || "Sin área"}`;

        equipoEncontrado.style.display = "flex";
        gestionTrabajador.disabled = false;

        let trabajador = obtenerTrabajadorEquipo(equipo.Nombre);

        if (!trabajadoresCargados) {
            try {
                await obtenerTrabajadores();
                trabajador = obtenerTrabajadorEquipo(equipo.Nombre);
            } catch (error) {
                // Sin mensaje visual.
            }
        }

        if (trabajador) {
            gestionTrabajador.value = trabajador.nombre;
            btnRegistrarTrabajador.disabled = true;
            btnEditarTrabajador.disabled = false;
            btnEliminarTrabajador.disabled = false;
        } else {
            gestionTrabajador.value = "";
            btnRegistrarTrabajador.disabled = false;
            btnEditarTrabajador.disabled = true;
            btnEliminarTrabajador.disabled = true;
        }

        gestionTrabajador.focus();
    } catch (error) {
        mostrarMensajeGestion();
    } finally {
        buscarEquipoIP.disabled = false;
        buscarEquipoIP.textContent = "Buscar";
    }
}

async function registrarTrabajadorPorIP() {
    if (!equipoGestionSeleccionado) return;

    const nombre = gestionTrabajador.value.trim();

    if (!nombre) {
        gestionTrabajador.focus();
        return;
    }

    const equipo = equipoGestionSeleccionado;

    let asignacionActual = obtenerTrabajadorEquipo(equipo.Nombre);

    if (!asignacionActual && !trabajadoresCargados) {
        await obtenerTrabajadores();
        asignacionActual = obtenerTrabajadorEquipo(equipo.Nombre);
    }

    if (asignacionActual) return;

    try {
        btnRegistrarTrabajador.disabled = true;
        btnRegistrarTrabajador.textContent = "Registrando...";

        const sitio = await obtenerSitioSharePoint();
        const lista = await obtenerListaTrabajadores();
        const columnas = await obtenerColumnasTrabajadores();

        const campoEquipo = obtenerNombreInternoColumna(
            columnas,
            "Equipo1"
        );

        const campoTrabajador = obtenerNombreInternoColumna(
            columnas,
            "NombreTrabajador"
        );

        if (!campoEquipo || !campoTrabajador) {
            throw new Error("No se encontraron las columnas necesarias.");
        }

        if (
            campoTrabajador === "Title" ||
            campoTrabajador === "LinkTitle"
        ) {
            throw new Error("La columna del trabajador no es válida.");
        }

        const fields = {};
        fields[campoEquipo] = equipo.Nombre;
        fields[campoTrabajador] = nombre;

        const respuesta = await graphFetch(
            `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items`,
            {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ fields })
            }
        );

        if (!respuesta.ok) {
            throw new Error("No se pudo registrar la asignación.");
        }

        const resultado = await respuesta.json();

        trabajadoresPorEquipo[normalizarTexto(equipo.Nombre)] = {
            nombre,
            itemId: resultado.id,
            equipo: equipo.Nombre
        };

        trabajadoresCargados = true;
        trabajadoresListos = true;

        renderizarTablaEquipos();

        gestionTrabajador.value = nombre;
        btnRegistrarTrabajador.disabled = true;
        btnEditarTrabajador.disabled = false;
        btnEliminarTrabajador.disabled = false;

        mostrarMensajeGestion();
    } catch (error) {
        mostrarMensajeGestion();
    } finally {
        btnRegistrarTrabajador.textContent = "Registrar";

        if (
            equipoGestionSeleccionado &&
            !obtenerTrabajadorEquipo(equipoGestionSeleccionado.Nombre)
        ) {
            btnRegistrarTrabajador.disabled = false;
        }
    }
}

async function editarTrabajadorPorIP() {
    if (!equipoGestionSeleccionado) return;

    const nombre = gestionTrabajador.value.trim();

    if (!nombre) {
        gestionTrabajador.focus();
        return;
    }

    const equipo = equipoGestionSeleccionado;
    const asignacion = obtenerTrabajadorEquipo(equipo.Nombre);

    if (!asignacion || !asignacion.itemId) return;

    try {
        btnEditarTrabajador.disabled = true;
        btnEditarTrabajador.textContent = "Guardando...";

        const sitio = await obtenerSitioSharePoint();
        const lista = await obtenerListaTrabajadores();
        const columnas = await obtenerColumnasTrabajadores();

        const campoTrabajador = obtenerNombreInternoColumna(
            columnas,
            "NombreTrabajador"
        );

        if (!campoTrabajador) {
            throw new Error("No se encontró la columna del trabajador.");
        }

        const fields = {};
        fields[campoTrabajador] = nombre;

        const respuesta = await graphFetch(
            `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items/${asignacion.itemId}/fields`,
            {
                method: "PATCH",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(fields)
            }
        );

        if (!respuesta.ok) {
            throw new Error("No se pudo actualizar la asignación.");
        }

        trabajadoresPorEquipo[normalizarTexto(equipo.Nombre)].nombre = nombre;

        renderizarTablaEquipos();
        gestionTrabajador.value = nombre;

        mostrarMensajeGestion();
    } catch (error) {
        mostrarMensajeGestion();
    } finally {
        btnEditarTrabajador.disabled = false;
        btnEditarTrabajador.textContent = "Editar";
    }
}

async function eliminarTrabajadorPorIP() {
    if (!equipoGestionSeleccionado) return;

    const equipo = equipoGestionSeleccionado;
    const asignacion = obtenerTrabajadorEquipo(equipo.Nombre);

    if (!asignacion || !asignacion.itemId) return;

    /*
     * Confirmación eliminada.
     * La asignación se elimina directamente al pulsar el botón.
     */

    try {
        btnEliminarTrabajador.disabled = true;
        btnEliminarTrabajador.textContent = "Eliminando...";

        const sitio = await obtenerSitioSharePoint();
        const lista = await obtenerListaTrabajadores();

        const respuesta = await graphFetch(
            `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items/${asignacion.itemId}`,
            { method: "DELETE" }
        );

        if (!respuesta.ok) {
            throw new Error("No se pudo eliminar la asignación.");
        }

        delete trabajadoresPorEquipo[normalizarTexto(equipo.Nombre)];

        renderizarTablaEquipos();

        gestionTrabajador.value = "";

        btnRegistrarTrabajador.disabled = false;
        btnEditarTrabajador.disabled = true;
        btnEliminarTrabajador.disabled = true;

        mostrarMensajeGestion();
    } catch (error) {
        mostrarMensajeGestion();
    } finally {
        btnEliminarTrabajador.textContent = "Eliminar";
    }
}

/* =========================================================
   EVENTOS DE GESTIÓN
========================================================= */

if (btnGestionarTrabajadores) {
    btnGestionarTrabajadores.addEventListener(
        "click",
        abrirGestionTrabajadores
    );
}

if (buscarEquipoIP) {
    buscarEquipoIP.addEventListener("click", buscarEquipoPorIP);
}

if (gestionIP) {
    gestionIP.addEventListener("keydown", event => {
        if (event.key === "Enter") {
            event.preventDefault();
            buscarEquipoPorIP();
        }

        if (event.key === "Escape") {
            cerrarGestionTrabajadores();
        }
    });
}

if (gestionTrabajador) {
    gestionTrabajador.addEventListener("keydown", event => {
        if (event.key === "Escape") {
            cerrarGestionTrabajadores();
        }
    });
}

if (btnRegistrarTrabajador) {
    btnRegistrarTrabajador.addEventListener(
        "click",
        registrarTrabajadorPorIP
    );
}

if (btnEditarTrabajador) {
    btnEditarTrabajador.addEventListener(
        "click",
        editarTrabajadorPorIP
    );
}

if (btnEliminarTrabajador) {
    btnEliminarTrabajador.addEventListener(
        "click",
        eliminarTrabajadorPorIP
    );
}

if (cerrarGestionTrabajadorModal) {
    cerrarGestionTrabajadorModal.addEventListener(
        "click",
        cerrarGestionTrabajadores
    );
}

if (cancelarGestionTrabajador) {
    cancelarGestionTrabajador.addEventListener(
        "click",
        cerrarGestionTrabajadores
    );
}

if (gestionTrabajadorModal) {
    gestionTrabajadorModal.addEventListener("click", event => {
        if (event.target === gestionTrabajadorModal) {
            cerrarGestionTrabajadores();
        }
    });
}

if (buscarEquipo) {
    let temporizadorBusqueda = null;

    buscarEquipo.addEventListener("input", () => {
        clearTimeout(temporizadorBusqueda);

        temporizadorBusqueda = setTimeout(
            () => aplicarFiltrosEquipos(),
            200
        );
    });
}

if (filtroEstadoEquipo) {
    filtroEstadoEquipo.addEventListener(
        "change",
        () => aplicarFiltrosEquipos()
    );
}

if (filtroAreaEquipo) {
    filtroAreaEquipo.addEventListener(
        "change",
        () => aplicarFiltrosEquipos()
    );
}

/* =========================================================
   PERMISOS DE INVENTARIO
========================================================= */

function usuarioPuedeVerInventario() {
    const permisos = window.alferzaPermisos || {};

    return String(permisos.grupo || "")
        .trim()
        .toUpperCase() === "TI";
}

async function verificarPermisosInventario() {
    if (
        window.alferzaAccessReady &&
        typeof window.alferzaAccessReady.then === "function"
    ) {
        try {
            await window.alferzaAccessReady;
        } catch (error) {
            return false;
        }
    }

    return usuarioPuedeVerInventario();
}

/* =========================================================
   CARGA DE INVENTARIO
========================================================= */

async function cargarEquipos() {
    if (cargandoEquipos) return;

    cargandoEquipos = true;

    try {
        equiposEstado.textContent = "● ACTUALIZANDO";
        equiposEstado.className = "section-status warning";

        const equipos = await obtenerEquiposDesdeSharePoint();

        if (equipos) {
            aplicarDatosEquipos(equipos);

            guardarCache("equipos", {
                etag: etagEquipos,
                datos: equipos
            });
        }

        lastUpdate.textContent = obtenerFechaActual();

        equiposEstado.textContent = "● ACTUALIZADO";
        equiposEstado.className = "section-status online";
    } catch (error) {
        if (equiposData.length > 0) {
            equiposEstado.textContent = "● SIN ACTUALIZAR";
            equiposEstado.className = "section-status warning";
        } else {
            equiposEstado.textContent = "";
            equiposEstado.className = "section-status";
            equiposTableBody.innerHTML = "";

            if (equiposResultados) {
                equiposResultados.textContent = "";
            }
        }
    } finally {
        cargandoEquipos = false;
    }
}

/* =========================================================
   CARGA DE SERVICIOS
========================================================= */

async function cargarServicios() {
    try {
        const servicios = await obtenerMonitoreoTI();
        renderizarServicios(servicios);
    } catch (error) {
        if (serviciosGrid) {
            serviciosGrid.innerHTML = "";
        }

        if (estadoGeneral) {
            estadoGeneral.textContent = "";
            estadoGeneral.className = "section-status";
        }
    }
}

/* =========================================================
   INICIO DE SESIÓN
========================================================= */

function mostrarBotonLogin() {
    if (loadingOverlay) {
        loadingOverlay.style.display = "flex";
    }

    if (loadingMessage) {
        loadingMessage.textContent = "";
    }

    if (document.getElementById("btnLoginMs")) return;

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
                ocultarLoading();
                iniciarCargas();
            }
        } catch (error) {
            if (loadingMessage) {
                loadingMessage.textContent = "";
            }
        } finally {
            boton.disabled = false;
        }
    });

    const caja =
        loadingOverlay.querySelector(".loading-content") ||
        loadingOverlay.querySelector(".loading-card");

    if (caja) {
        caja.appendChild(boton);
    }
}

/* =========================================================
   CARGA DEL DASHBOARD
========================================================= */

async function iniciarCargas() {
    ultimaActualizacion = Date.now();

    // Los servicios están disponibles para usuarios autorizados.
    cargarServicios();

    // El inventario solo se consulta para el grupo TI.
    if (await verificarPermisosInventario()) {
        cargarEquipos()
            .then(() => {
                if (equiposData.length > 0) {
                    return obtenerTrabajadores();
                }
            })
            .catch(() => {});
    }
}

async function cargarDashboard() {
    try {
        if (!msalInstanceInfraestructura) {
            mostrarLoading();
            return;
        }

        let conCache = false;

        if (msalInstanceInfraestructura.getAllAccounts().length > 0) {
            const cache = leerCache("equipos");

            if (
                cache &&
                Array.isArray(cache.datos) &&
                cache.datos.length > 0
            ) {
                etagEquipos = cache.etag || null;
                aplicarDatosEquipos(cache.datos);
                conCache = true;
            }
        }

        if (!conCache) {
            mostrarLoading();
        }

        const token = await iniciarSesion(false);

        if (!token) {
            mostrarBotonLogin();
            return;
        }

        ocultarLoading();
        iniciarCargas();
    } catch (error) {
        mostrarDiagnostico();
        mostrarBotonLogin();
    }
}

async function actualizarDashboard() {
    if (document.hidden || !tokenActual) return;

    try {
        if (!(await obtenerTokenSilencioso())) return;

        ultimaActualizacion = Date.now();

        await Promise.all([
            cargarServicios(),
            verificarPermisosInventario().then(puedeVer => {
                if (puedeVer) return cargarEquipos();
            })
        ]);
    } catch (error) {
        // Sin mensajes visuales.
    }
}

/* =========================================================
   INICIALIZACIÓN
========================================================= */

function iniciarInfraestructura() {
    cargarDashboard();

    setInterval(
        actualizarDashboard,
        INTERVALO_ACTUALIZACION
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