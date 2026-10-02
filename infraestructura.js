/* =========================================
   INFRAESTRUCTURA - MONITOREO TI
========================================= */


/* =========================================
   CONFIGURACIÓN MSAL
========================================= */

const msalConfigInfraestructura = {

    auth: {
        clientId:
            "5d98417c-74a7-4fab-8f2c-41ac127be696",

        authority:
            "https://login.microsoftonline.com/dbab984f-4bb1-4b60-9dff-da59f54acdf1",

        redirectUri:
            "https://AlferzaTI.github.io/live-office/blank.html"
    },

    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false
    }
};


/* Si el script de MSAL no cargó (CDN bloqueado/lento) no rompemos todo el archivo. */
const msalInstanceInfraestructura =
    (typeof msal !== "undefined")
        ? new msal.PublicClientApplication(msalConfigInfraestructura)
        : null;


/* =========================================
   SHAREPOINT
========================================= */

const SHAREPOINT_HOSTNAME =
    "alferzaholding-my.sharepoint.com";

const SHAREPOINT_SITE =
    "/personal/soporte1_alferza_pe";

const SHAREPOINT_LIST_SERVICIOS =
    "MonitoreoTI";

const SHAREPOINT_LIST_TRABAJADORES =
    "TrabajadoresEquipos";

const RUTA_EQUIPOS =
    "Procedimientos T.I/equipos.json";


/* =========================================
   CONFIGURACIÓN
========================================= */

const EQUIPOS_POR_PAGINA = 40;

const INTERVALO_ACTUALIZACION =
    120000;


/* =========================================
   VARIABLES
========================================= */

let cuentaActual = null;

let tokenActual = null;

let sitioSharePoint = null;

let driveSharePoint = null;


/*
 * Datos de trabajadores.
 *
 * IMPORTANTE:
 * Estos datos se cargan en segundo plano.
 * Nunca bloquean la carga inicial del inventario.
 */

let listaTrabajadores = null;

let columnasTrabajadores = null;

let trabajadoresData = [];

let trabajadoresPorEquipo = {};

let trabajadoresCargados = false;

let cargandoTrabajadores = false;


/* =========================================
   EQUIPOS
========================================= */

let equiposData = [];

let equiposFiltrados = [];

let paginaActual = 1;


/* =========================================
   MODAL
========================================= */

let equipoSeleccionado = null;


/* =========================================
   ELEMENTOS DOM
========================================= */

const loadingOverlay =
    document.getElementById(
        "loadingOverlay"
    );

const loadingMessage =
    document.getElementById(
        "loadingMessage"
    );

const serviciosGrid =
    document.getElementById(
        "serviciosGrid"
    );

const estadoGeneral =
    document.getElementById(
        "estadoGeneral"
    );

const equiposTableBody =
    document.getElementById(
        "equiposTableBody"
    );

const equiposEmpty =
    document.getElementById(
        "equiposEmpty"
    );

const equiposResultados =
    document.getElementById(
        "equiposResultados"
    );

const totalEquipos =
    document.getElementById(
        "totalEquipos"
    );

const equiposConectados =
    document.getElementById(
        "equiposConectados"
    );

const equiposDesconectados =
    document.getElementById(
        "equiposDesconectados"
    );

const filtroEstadoEquipo =
    document.getElementById(
        "filtroEstadoEquipo"
    );

const filtroAreaEquipo =
    document.getElementById(
        "filtroAreaEquipo"
    );

const buscarEquipo =
    document.getElementById(
        "buscarEquipo"
    );

const equiposEstado =
    document.getElementById(
        "equiposEstado"
    );

const lastUpdate =
    document.getElementById(
        "lastUpdate"
    );


/* =========================================
   ELEMENTOS MODAL
========================================= */

const trabajadorModal =
    document.getElementById(
        "trabajadorModal"
    );

const modalEquipoNombre =
    document.getElementById(
        "modalEquipoNombre"
    );

const trabajadorNombre =
    document.getElementById(
        "trabajadorNombre"
    );

const trabajadorError =
    document.getElementById(
        "trabajadorError"
    );

const guardarTrabajador =
    document.getElementById(
        "guardarTrabajador"
    );

const cerrarTrabajadorModal =
    document.getElementById(
        "cerrarTrabajadorModal"
    );

const cancelarTrabajador =
    document.getElementById(
        "cancelarTrabajador"
    );


/* =========================================
   UTILIDADES
========================================= */

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


function mostrarLoading(mensaje) {

    if (loadingOverlay) {
        loadingOverlay.style.display =
            "flex";
    }

    if (loadingMessage) {
        loadingMessage.textContent =
            mensaje || "Cargando...";
    }
}


function ocultarLoading() {

    if (loadingOverlay) {
        loadingOverlay.style.display =
            "none";
    }
}


function obtenerFechaActual() {

    const ahora =
        new Date();

    return ahora.toLocaleString(
        "es-PE",
        {
            dateStyle: "short",
            timeStyle: "medium"
        }
    );
}


function normalizarTexto(valor) {

    return String(
        valor || ""
    )
        .trim()
        .toLowerCase();
}


/* =========================================
   AUTENTICACIÓN
========================================= */

/* =========================================
   DIAGNÓSTICO EN PANTALLA
   Muestra el error real y qué revisar.
   (Estilos en línea: no depende del CSS.)
========================================= */

function ayudaError(texto) {
    const t = String(texto);
    const ayudas = [
        [/AADSTS65001|consent_required/i,
            "Falta consentimiento de administrador. En Azure AD → Registros de aplicaciones → ALFERZA Live Office → Permisos de API, agrega Sites.Read.All y Sites.ReadWrite.All (delegados) y pulsa «Conceder consentimiento de administrador»."],
        [/AADSTS70011|invalid_scope/i,
            "Un permiso (scope) no existe o no está configurado en la app de Azure AD."],
        [/AADSTS50011|redirect_uri/i,
            "La URL de redirección no coincide. Registra " + msalConfigInfraestructura.auth.redirectUri + " como tipo SPA en Azure AD y abre la página desde ese mismo dominio."],
        [/AADSTS700054|AADSTS9002326|cross-origin/i,
            "La URL de redirección debe estar registrada como plataforma «Aplicación de página única (SPA)», no «Web»."],
        [/popup_window_error|empty_window_error|user_cancelled|monitor_window_timeout/i,
            "El navegador bloqueó o cerró la ventana de inicio de sesión. Permite las ventanas emergentes para este sitio e inténtalo de nuevo."],
        [/crypto_nonexistent|insecure/i,
            "MSAL solo funciona en https:// o http://localhost. No abras el archivo con doble clic (file://) ni por IP."],
        [/ 401 |InvalidAuthenticationToken/i,
            "Graph rechazó el token (401). Cierra sesión, recarga y vuelve a iniciar sesión."],
        [/ 403 |accessDenied|Access denied/i,
            "Tu usuario no tiene acceso a ese sitio/archivo de SharePoint (403). Debe tener permiso de lectura sobre el OneDrive soporte1_alferza_pe."],
        [/No se encontró ".*equipos/i,
            "El archivo no está en la ruta configurada. Abajo se lista lo que hay en la raíz del OneDrive: ajusta RUTA_EQUIPOS al inicio de infraestructura.js."],
        [/no es un JSON válido|no contiene un arreglo/i,
            "El archivo se descargó pero su contenido no tiene el formato esperado (un arreglo JSON de equipos)."],
        [/ 404 |itemNotFound|Invalid hostname/i,
            "No se encontró el sitio, la lista o el archivo (404). Revisa SHAREPOINT_SITE, el nombre de la lista y la ruta «Procedimientos T.I/equipos.json»."],
        [/Failed to fetch|NetworkError|AbortError/i,
            "Sin respuesta de Microsoft Graph (red, VPN, bloqueador de anuncios o tiempo agotado)."]
    ];
    const hallada = ayudas.find(a => a[0].test(t));
    return hallada ? hallada[1] : "";
}

const peticionesFallidas = [];

/* Guarda qué petición a Graph falló (sin el token) para mostrarla en el diagnóstico. */
function registrarFallo(status, url) {
    let corta = String(url).replace("https://graph.microsoft.com/v1.0", "");
    try { corta = decodeURIComponent(corta); } catch (e) {}
    const linea = status + "  " + corta.slice(0, 170);
    if (!peticionesFallidas.includes(linea)) peticionesFallidas.push(linea);
    console.warn("[Graph] petición fallida:", status, corta);
}

const diagnosticoLineas = [];

function mostrarDiagnostico(paso, error) {
    try {
        const mensaje = String(
            (error && (error.errorCode ? error.errorCode + ": " : "") +
            (error.errorMessage || error.message)) || error || ""
        ).slice(0, 700);
        diagnosticoLineas.push({ paso, mensaje, ayuda: ayudaError(mensaje + " " + (error && error.errorCode || "")) });

        let panel = document.getElementById("panelDiagnostico");
        if (!panel) {
            panel = document.createElement("div");
            panel.id = "panelDiagnostico";
            panel.style.cssText =
                "position:fixed;right:16px;bottom:16px;z-index:99999;max-width:520px;max-height:70vh;" +
                "overflow:auto;background:#fff;color:#172033;border:2px solid #c0392b;border-radius:12px;" +
                "box-shadow:0 10px 40px rgba(0,0,0,.35);padding:14px 16px;font:12px/1.5 Arial,sans-serif;";
            document.body.appendChild(panel);
        }
        panel.textContent = "";

        const titulo = document.createElement("strong");
        titulo.textContent = "Diagnóstico de carga";
        titulo.style.cssText = "display:block;margin-bottom:6px;color:#c0392b;font-size:13px;";
        panel.appendChild(titulo);

        const contexto = document.createElement("div");
        contexto.style.cssText = "color:#667085;margin-bottom:8px;";
        let origenRedirect = "";
        try { origenRedirect = new URL(msalConfigInfraestructura.auth.redirectUri).origin; } catch (e) {}
        contexto.textContent =
            "Página: " + location.origin +
            " | MSAL: " + (msalInstanceInfraestructura ? "cargado" : "NO cargó") +
            " | Sesiones: " + (msalInstanceInfraestructura ? msalInstanceInfraestructura.getAllAccounts().length : 0) +
            " | Token: " + (tokenActual ? "sí" : "no");
        panel.appendChild(contexto);

        if (origenRedirect && location.origin.toLowerCase() !== origenRedirect.toLowerCase()) {
            const aviso = document.createElement("div");
            aviso.style.cssText = "background:#fff4e5;border:1px solid #f0b36b;border-radius:8px;padding:8px;margin-bottom:8px;";
            aviso.textContent =
                "Ojo: la página se abre desde " + location.origin + " pero el redirectUri del código es " + origenRedirect +
                ". Deben ser el mismo dominio o el inicio de sesión por popup no puede completarse.";
            panel.appendChild(aviso);
        }

        if (peticionesFallidas.length) {
            const pf = document.createElement("div");
            pf.style.cssText = "background:#fdecea;border:1px solid #f1a9a0;border-radius:8px;padding:8px;margin-bottom:8px;";
            const pt = document.createElement("strong");
            pt.textContent = "Peticiones a Microsoft Graph que fallaron:";
            pf.appendChild(pt);
            peticionesFallidas.slice(-8).forEach(l => {
                const d = document.createElement("div");
                d.textContent = l;
                d.style.cssText = "font-family:Consolas,monospace;font-size:11px;word-break:break-all;";
                pf.appendChild(d);
            });
            panel.appendChild(pf);
        }

        diagnosticoLineas.forEach(l => {
            const bloque = document.createElement("div");
            bloque.style.cssText = "border-top:1px solid #e4e7ec;padding:7px 0;";
            const p = document.createElement("strong");
            p.textContent = l.paso;
            const m = document.createElement("div");
            m.textContent = l.mensaje;
            m.style.cssText = "font-family:Consolas,monospace;font-size:11px;word-break:break-word;color:#344054;";
            bloque.appendChild(p);
            bloque.appendChild(m);
            if (l.ayuda) {
                const a = document.createElement("div");
                a.textContent = "→ " + l.ayuda;
                a.style.cssText = "margin-top:3px;color:#0a4da2;font-weight:700;";
                bloque.appendChild(a);
            }
            panel.appendChild(bloque);
        });

        const cerrar = document.createElement("button");
        cerrar.type = "button";
        cerrar.textContent = "Cerrar";
        cerrar.style.cssText = "margin-top:8px;padding:5px 12px;border:1px solid #cbd5e1;border-radius:7px;background:#f8fafc;cursor:pointer;";
        cerrar.addEventListener("click", () => panel.remove());
        panel.appendChild(cerrar);
    } catch (e) {
        console.error("No se pudo mostrar el diagnóstico:", e);
    }
}

window.addEventListener("unhandledrejection", e => mostrarDiagnostico("Error no controlado", e.reason));
window.addEventListener("error", e => mostrarDiagnostico("Error de script", e.error || e.message));

const SCOPES_GRAPH = ["User.Read", "Sites.Read.All", "Sites.ReadWrite.All"];
const CACHE_PREFIJO = "alferza_ti_";

let etagEquipos = null;
let promesaSitio = null;
let promesaDrive = null;
let promesaTrabajadores = null;
let trabajadoresListos = false;
let cargandoEquipos = false;
let ultimaActualizacion = 0;

function leerCache(clave) {
    try {
        const valor = localStorage.getItem(CACHE_PREFIJO + clave);
        return valor ? JSON.parse(valor) : null;
    } catch (e) {
        return null;
    }
}

function guardarCache(clave, valor) {
    try {
        localStorage.setItem(CACHE_PREFIJO + clave, JSON.stringify(valor));
    } catch (e) {
        console.warn("No se pudo guardar en caché:", clave);
    }
}

/* Token desde caché de MSAL, sin popups. Devuelve null si no hay sesión. */
async function obtenerTokenSilencioso() {
    if (!msalInstanceInfraestructura) return null;
    const cuentas = msalInstanceInfraestructura.getAllAccounts();
    if (!cuentas.length) return null;
    cuentaActual = cuentas[0];
    try {
        const r = await msalInstanceInfraestructura.acquireTokenSilent({
            scopes: SCOPES_GRAPH,
            account: cuentaActual
        });
        tokenActual = r.accessToken;
        return tokenActual;
    } catch (error) {
        console.warn("Token silencioso no disponible:", error);
        mostrarDiagnostico("Token silencioso (se pedirá iniciar sesión)", error);
        return null;
    }
}

/* interactivo=true abre popup: debe llamarse desde un clic del usuario. */
async function iniciarSesion(interactivo) {
    const silencioso = await obtenerTokenSilencioso();
    if (silencioso || !interactivo) return silencioso;
    const r = cuentaActual
        ? await msalInstanceInfraestructura.acquireTokenPopup({ scopes: SCOPES_GRAPH, account: cuentaActual })
        : await msalInstanceInfraestructura.loginPopup({ scopes: SCOPES_GRAPH });
    cuentaActual = r.account;
    tokenActual = r.accessToken || (await obtenerTokenSilencioso());
    return tokenActual;
}

/* fetch a Graph con timeout, reintento en 429/503 y renovación de token en 401. */
async function graphFetch(url, opciones, intento) {
    intento = intento || 0;
    const { timeoutMs = 30000, ...resto } = opciones || {};
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), timeoutMs);
    try {
        const response = await fetch(url, {
            ...resto,
            signal: ctrl.signal,
            headers: { Authorization: `Bearer ${tokenActual}`, ...(resto.headers || {}) }
        });
        if (!response.ok) registrarFallo(response.status, url);
        if ((response.status === 429 || response.status === 503) && intento < 3) {
            const espera = (parseInt(response.headers.get("Retry-After"), 10) || 2 ** intento) * 1000;
            await new Promise(r => setTimeout(r, Math.min(espera, 10000)));
            return graphFetch(url, opciones, intento + 1);
        }
        if (response.status === 401 && intento < 1 && (await obtenerTokenSilencioso())) {
            return graphFetch(url, opciones, intento + 1);
        }
        return response;
    } finally {
        clearTimeout(timer);
    }
}


/* =========================================
   OBTENER SITIO SHAREPOINT
========================================= */

/* Una sola petición aunque varias funciones la pidan a la vez; el id se guarda en caché. */
function obtenerSitioSharePoint() {
    if (!promesaSitio) {
        promesaSitio = (async () => {
            const clave = "sitio:" + SHAREPOINT_SITE;
            const guardado = leerCache(clave);
            if (guardado && guardado.id) return guardado;
            const response = await graphFetch(
                `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOSTNAME}:${SHAREPOINT_SITE}?$select=id`
            );
            if (!response.ok) {
                throw new Error(`No se pudo obtener el sitio SharePoint. ${response.status} ${await response.text()}`);
            }
            const sitio = await response.json();
            guardarCache(clave, { id: sitio.id });
            return { id: sitio.id };
        })().catch(error => { promesaSitio = null; throw error; });
    }
    return promesaSitio;
}


/* =========================================
   OBTENER DRIVE
========================================= */

function obtenerDriveSharePoint() {
    if (!promesaDrive) {
        promesaDrive = (async () => {
            const sitio = await obtenerSitioSharePoint();
            const clave = "drive:" + sitio.id;
            const guardado = leerCache(clave);
            if (guardado && guardado.id) return guardado;
            const response = await graphFetch(
                `https://graph.microsoft.com/v1.0/sites/${sitio.id}/drive?$select=id`
            );
            if (!response.ok) {
                throw new Error(`No se pudo obtener el drive de SharePoint. ${response.status} ${await response.text()}`);
            }
            const drive = await response.json();
            guardarCache(clave, { id: drive.id });
            return { id: drive.id };
        })().catch(error => { promesaDrive = null; throw error; });
    }
    return promesaDrive;
}


/* =========================================
   OBTENER ARCHIVO EQUIPOS
========================================= */

/* Devuelve null si equipos.json no cambió (no se vuelve a descargar). */
let ubicacionEquipos = null;

/* Si el archivo no está en RUTA_EQUIPOS, lo busca por nombre en el drive. */
async function buscarEquiposJson(raiz) {
    const nombre = RUTA_EQUIPOS.split("/").pop();
    const r = await graphFetch(
        `${raiz}/root/search(q='${encodeURIComponent(nombre)}')?$select=id,name,eTag,parentReference&$top=25`
    );
    if (r.ok) {
        const items = ((await r.json()).value || [])
            .filter(i => normalizarTexto(i.name) === normalizarTexto(nombre));
        if (items.length > 0) {
            const it = items[0];
            console.warn(
                `"${nombre}" no estaba en "${RUTA_EQUIPOS}". Se encontró en: ` +
                `${(it.parentReference && it.parentReference.path) || "(ruta desconocida)"}. ` +
                `Puedes actualizar RUTA_EQUIPOS en infraestructura.js.`
            );
            return {
                meta: `${raiz}/items/${it.id}`,
                contenido: `${raiz}/items/${it.id}/content`,
                etag: it.eTag || null
            };
        }
    }

    let contenidoRaiz = "";
    try {
        const l = await graphFetch(`${raiz}/root/children?$select=name,folder&$top=50`);
        if (l.ok) {
            contenidoRaiz = ((await l.json()).value || [])
                .map(x => (x.folder ? "[carpeta] " : "") + x.name)
                .join(", ");
        }
    } catch (e) { /* solo informativo */ }

    throw new Error(
        `No se encontró "${RUTA_EQUIPOS}" en el OneDrive/sitio configurado. ` +
        `Contenido de la raíz: ${contenidoRaiz || "(no se pudo listar)"}`
    );
}

/* Acepta UTF-8 y UTF-16 (PowerShell suele exportar en UTF-16), y arreglo u objeto contenedor. */
async function leerJsonEquipos(response) {
    const buffer = await response.arrayBuffer();
    const b = new Uint8Array(buffer);
    let codificacion = "utf-8";
    if (b[0] === 0xFF && b[1] === 0xFE) codificacion = "utf-16le";
    else if (b[0] === 0xFE && b[1] === 0xFF) codificacion = "utf-16be";

    const texto = new TextDecoder(codificacion).decode(buffer).replace(/^\uFEFF/, "");

    let datos;
    try {
        datos = JSON.parse(texto);
    } catch (e) {
        throw new Error(`equipos.json no es un JSON válido: ${e.message}. Inicio del archivo: ${texto.slice(0, 60)}`);
    }

    if (Array.isArray(datos)) return datos;

    if (datos && typeof datos === "object") {
        for (const clave of ["value", "equipos", "Equipos", "data", "items"]) {
            if (Array.isArray(datos[clave])) return datos[clave];
        }
        if (datos.Nombre) return [datos];   /* un solo equipo exportado como objeto */
    }

    throw new Error(
        "equipos.json no contiene un arreglo de equipos. Claves encontradas: " +
        Object.keys(datos || {}).slice(0, 10).join(", ")
    );
}

/* Devuelve null si equipos.json no cambió (no se vuelve a descargar). */
async function obtenerEquiposDesdeSharePoint() {
    const drive = await obtenerDriveSharePoint();
    const raiz = `https://graph.microsoft.com/v1.0/drives/${drive.id}`;
    const ruta = RUTA_EQUIPOS.split("/").map(p => encodeURIComponent(p)).join("/");

    const candidatos = [];
    if (ubicacionEquipos) candidatos.push(ubicacionEquipos);
    candidatos.push({
        meta: `${raiz}/root:/${ruta}`,
        contenido: `${raiz}/root:/${ruta}:/content`
    });

    let ubicacion = null;
    let etag = null;

    for (const c of candidatos) {
        const r = await graphFetch(`${c.meta}?$select=eTag`);
        if (r.ok) {
            etag = (await r.json()).eTag || null;
            ubicacion = c;
            break;
        }
        if (r.status !== 404) {
            throw new Error(`No se pudo consultar equipos.json. ${r.status} ${await r.text()}`);
        }
    }

    if (!ubicacion) {
        ubicacion = await buscarEquiposJson(raiz);
        etag = ubicacion.etag;
    }
    ubicacionEquipos = ubicacion;

    if (etag && etag === etagEquipos && equiposData.length > 0) return null;

    const response = await graphFetch(ubicacion.contenido, { timeoutMs: 90000 });
    if (!response.ok) {
        throw new Error(`No se pudo descargar equipos.json. ${response.status} ${await response.text()}`);
    }

    const datos = await leerJsonEquipos(response);
    etagEquipos = etag;
    return datos;
}


/* =========================================
   OBTENER LISTA TRABAJADORES
========================================= */

async function obtenerListaTrabajadores(
    token
) {

    if (
        listaTrabajadores
    ) {
        return listaTrabajadores;
    }


    const sitio =
        await obtenerSitioSharePoint(
            token
        );


    const url =
        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists`;


    const response =
        await fetch(
            url,
            {
                headers: {
                    Authorization:
                        `Bearer ${token}`
                }
            }
        );


    if (
        !response.ok
    ) {

        const texto =
            await response.text();

        throw new Error(
            `No se pudieron obtener las listas SharePoint. ${response.status} ${texto}`
        );
    }


    const data =
        await response.json();


    const listas =
        data.value || [];


    listaTrabajadores =
        listas.find(
            lista =>
                normalizarTexto(
                    lista.displayName
                ) ===
                normalizarTexto(
                    SHAREPOINT_LIST_TRABAJADORES
                )
        );


    if (
        !listaTrabajadores
    ) {

        throw new Error(
            `No se encontró la lista "${SHAREPOINT_LIST_TRABAJADORES}" en el sitio SharePoint.`
        );
    }


    return listaTrabajadores;
}


/* =========================================
   OBTENER COLUMNAS
========================================= */

async function obtenerColumnasTrabajadores(
    token
) {

    if (
        columnasTrabajadores
    ) {
        return columnasTrabajadores;
    }


    const lista =
        await obtenerListaTrabajadores(
            token
        );


    const sitio =
        await obtenerSitioSharePoint(
            token
        );


    const url =
        `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/columns`;


    const response =
        await fetch(
            url,
            {
                headers: {
                    Authorization:
                        `Bearer ${token}`
                }
            }
        );


    if (
        !response.ok
    ) {

        const texto =
            await response.text();

        throw new Error(
            `No se pudieron obtener las columnas de TrabajadoresEquipos. ${response.status} ${texto}`
        );
    }


    const data =
        await response.json();


    columnasTrabajadores =
        data.value || [];


    return columnasTrabajadores;
}


/* =========================================
   BUSCAR NOMBRE INTERNO DE COLUMNA
========================================= */

function obtenerNombreInternoColumna(
    columnas,
    nombreVisible
) {

    const columna =
        columnas.find(
            item =>
                normalizarTexto(
                    item.displayName
                ) ===
                normalizarTexto(
                    nombreVisible
                )
        );


    return columna
        ? columna.name
        : null;
}


/* =========================================
   CARGAR TRABAJADORES
   -----------------------------------------
   ESTA FUNCIÓN YA NO BLOQUEA EL INVENTARIO.
========================================= */

/* Comparte una sola carga entre todos los llamadores (evita duplicados y listas vacías). */
function obtenerTrabajadores() {
    if (trabajadoresCargados) return Promise.resolve(trabajadoresPorEquipo);
    if (!promesaTrabajadores) {
        promesaTrabajadores = cargarTrabajadoresDesdeSharePoint()
            .finally(() => { promesaTrabajadores = null; });
    }
    return promesaTrabajadores;
}

async function cargarTrabajadoresDesdeSharePoint() {
    try {
        const sitio = await obtenerSitioSharePoint(tokenActual);
        const lista = await obtenerListaTrabajadores(tokenActual);

        let url =
            `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${lista.id}/items` +
            `?expand=fields($select=Equipo1,NombreTrabajador)&$top=999`;

        const items = [];

        while (url) {
            const response = await graphFetch(url);

            if (!response.ok) {
                const texto = await response.text();

                throw new Error(
                    `Error cargando trabajadores: ${response.status} ${texto}`
                );
            }

            const data = await response.json();

            items.push(...(data.value || []));

            url = data["@odata.nextLink"] || null;
        }

        trabajadoresData = items;
        trabajadoresPorEquipo = {};

        items.forEach(item => {
            const fields = item.fields || {};

            if (
                fields.Equipo1 &&
                fields.NombreTrabajador
            ) {
                trabajadoresPorEquipo[
                    normalizarTexto(fields.Equipo1)
                ] = {
                    nombre: fields.NombreTrabajador,
                    itemId: item.id,
                    equipo: fields.Equipo1
                };
            }
        });

        console.log(
            "Asignaciones de trabajadores cargadas:",
            Object.keys(trabajadoresPorEquipo).length
        );

        if (
            items.length > 0 &&
            Object.keys(trabajadoresPorEquipo).length === 0
        ) {
            await cargarTrabajadoresConColumnas(
                tokenActual,
                sitio,
                lista
            );
        }

        trabajadoresCargados = true;
        trabajadoresListos = true;

        if (equiposData.length > 0) {
            renderizarTablaEquipos();
        }

        return trabajadoresPorEquipo;

    } catch (error) {

        console.warn(
            "No se pudieron cargar las asignaciones de trabajadores:",
            error
        );

        trabajadoresListos = true;

        if (equiposData.length > 0) {
            renderizarTablaEquipos();
        }

        return trabajadoresPorEquipo;
    }
}

/* =========================================
   RESPALDO PARA COLUMNAS
========================================= */

async function cargarTrabajadoresConColumnas(
    token,
    sitio,
    lista
) {

    const columnas =
        await obtenerColumnasTrabajadores(
            token
        );


    const campoEquipo =
        obtenerNombreInternoColumna(
            columnas,
            "Equipo1"
        );


    const campoTrabajador =
        obtenerNombreInternoColumna(
            columnas,
            "Trabajador"
        );


    if (
        !campoEquipo ||
        !campoTrabajador
    ) {

        return;
    }


    trabajadoresPorEquipo =
        {};


    trabajadoresData.forEach(
        item => {

            const fields =
                item.fields || {};


            const equipo =
                fields[campoEquipo];


            const trabajador =
                fields[campoTrabajador];


            if (
                equipo &&
                trabajador
            ) {

                trabajadoresPorEquipo[
                    normalizarTexto(
                        equipo
                    )
                ] = {

                    nombre:
                        trabajador,

                    itemId:
                        item.id,

                    equipo:
                        equipo
                };
            }
        }
    );
}


/* =========================================
   OBTENER ASIGNACIÓN
========================================= */

function obtenerTrabajadorEquipo(
    nombreEquipo
) {

    return trabajadoresPorEquipo[
        normalizarTexto(
            nombreEquipo
        )
    ] || null;
}


/* =========================================
   CARGAR SERVICIOS DESDE SHAREPOINT
========================================= */

async function obtenerMonitoreoTI() {
    const sitio = await obtenerSitioSharePoint();
    const url = `https://graph.microsoft.com/v1.0/sites/${sitio.id}/lists/${encodeURIComponent(SHAREPOINT_LIST_SERVICIOS)}/items?expand=fields&$top=999`;
    const response = await graphFetch(url);
    if (!response.ok) {
        throw new Error(`No se pudo obtener MonitoreoTI. ${response.status} ${await response.text()}`);
    }
    return (await response.json()).value || [];
}

function obtenerCampo(fields, nombres, defecto) {
    const limpiar = texto =>
        String(texto)
            .replace(/_x([0-9a-f]{4})_/gi, (m, h) => String.fromCharCode(parseInt(h, 16)))
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .toLowerCase()
            .trim();

    const mapa = {};
    Object.keys(fields).forEach(clave => { mapa[limpiar(clave)] = fields[clave]; });

    for (const nombre of nombres) {
        const valor = mapa[limpiar(nombre)];
        if (valor !== undefined && valor !== null && valor !== "") return valor;
    }
    return defecto;
}

/* Respaldo: el texto más largo que no sea nombre ni estado */
function adivinarDescripcion(fields, excluir) {
    let mejor = "";
    Object.entries(fields).forEach(([clave, valor]) => {
        if (clave.startsWith("@") || clave.startsWith("_")) return;
        if (typeof valor !== "string") return;
        if (excluir.includes(valor)) return;
        if (/^\d{4}-\d{2}-\d{2}T/.test(valor)) return;   /* fechas ISO */
        if (valor.length > mejor.length) mejor = valor;
    });
    return mejor.length > 15 ? mejor : "";
}
/* =========================================
   RENDER SERVICIOS
========================================= */

function renderizarServicios(
    items
) {

    if (
        !serviciosGrid
    ) {
        return;
    }


    if (
        !items ||
        items.length === 0
    ) {

        serviciosGrid.innerHTML = `
            <div class="service-card service-loading">
                <span>
                    No hay servicios registrados.
                </span>
            </div>
        `;


        estadoGeneral.textContent =
            "● SIN DATOS";

        estadoGeneral.className =
            "section-status warning";


        return;
    }


    let serviciosOnline =
        0;


    serviciosGrid.innerHTML =
        items
            .map(
                item => {

                    const fields =
                        item.fields || {};


                    const nombre =
                        fields.Nombre ||
                        fields.Title ||
                        "Servicio";


                    const descripcion =
                        fields.Descripcion ||
                        fields.Descripción ||
                        "Sin descripción";


                    const estado =
                        fields.Estado ||
                        "Sin estado";


                    const latencia =
                        fields.Latencia ||
                        fields.Ping ||
                        "--";


                    const revisionRaw =
                        fields.Ultimarevision ||
                        fields.UltimaRevision ||
                        fields["ÚltimaRevision"] ||
                        fields.Fecha ||
                        null;


                    const revision = revisionRaw
                    ? new Date(revisionRaw).toLocaleString("es-PE", {
                        timeZone: "America/Lima",
                        day: "2-digit",
                        month: "2-digit",
                        year: "numeric",
                        hour: "2-digit",
                        minute: "2-digit",
                        hour12: false
                    })
                    : "--";

                    const estadoNormalizado =
                        normalizarTexto(
                            estado
                        );


                    let claseEstado =
                        "neutral";


                    if (
                        estadoNormalizado.includes(
                            "operativo"
                        ) ||
                        estadoNormalizado.includes(
                            "online"
                        ) ||
                        estadoNormalizado.includes(
                            "conectado"
                        )
                    ) {

                        claseEstado =
                            "online";

                        serviciosOnline++;


                    } else if (
                        estadoNormalizado.includes(
                            "mantenimiento"
                        ) ||
                        estadoNormalizado.includes(
                            "advertencia"
                        ) ||
                        estadoNormalizado.includes(
                            "warning"
                        )
                    ) {

                        claseEstado =
                            "warning";


                    } else if (
                        estadoNormalizado.includes(
                            "caido"
                        ) ||
                        estadoNormalizado.includes(
                            "offline"
                        ) ||
                        estadoNormalizado.includes(
                            "desconectado"
                        )
                    ) {

                        claseEstado =
                            "offline";
                    }


                    return `
                        <article
                            class="service-card ${claseEstado}"
                        >

                            <div class="service-main">

                                <div class="service-icon">
                                    🖥️
                                </div>

                                <div class="service-info">

                                    <div class="service-heading">

                                        <h3>
                                            ${escaparHTML(
                                                nombre
                                            )}
                                        </h3>

                                        <span
                                            class="service-status ${claseEstado}"
                                        >

                                            <span class="status-dot"></span>

                                            ${escaparHTML(
                                                estado
                                            )}

                                        </span>

                                    </div>

                                    <p>
                                        ${escaparHTML(
                                            descripcion
                                        )}
                                    </p>

                                </div>

                            </div>


                            <div class="service-meta">

                                <div class="service-latency">

                                    <span>
                                        Latencia
                                    </span>

                                    <strong>
                                        ${escaparHTML(
                                            latencia
                                        )}
                                    </strong>

                                </div>


                                <div class="service-revision">

                                    <span class="revision-icon">
                                        🕐
                                    </span>

                                    <div>

                                        <span>
                                            Última revisión
                                        </span>

                                        <strong>
                                            ${escaparHTML(
                                                revision
                                            )}
                                        </strong>

                                    </div>

                                </div>

                            </div>

                        </article>
                    `;
                }
            )
            .join("");


    if (
        serviciosOnline ===
        items.length
    ) {

        estadoGeneral.textContent =
            "● OPERATIVO";

        estadoGeneral.className =
            "section-status online";


    } else if (
        serviciosOnline > 0
    ) {

        estadoGeneral.textContent =
            "● REVISAR";

        estadoGeneral.className =
            "section-status warning";


    } else {

        estadoGeneral.textContent =
            "● NO DISPONIBLE";

        estadoGeneral.className =
            "section-status offline";
    }
}


/* =========================================
   RESUMEN EQUIPOS
========================================= */

function actualizarResumenEquipos() {

    const total =
        equiposData.length;


    const conectados =
        equiposData.filter(
            equipo =>
                normalizarTexto(
                    equipo.Estado
                ) ===
                "conectado"
        ).length;


    const desconectados =
        equiposData.filter(
            equipo =>
                normalizarTexto(
                    equipo.Estado
                ) ===
                "desconectado"
        ).length;


    totalEquipos.textContent =
        total;


    equiposConectados.textContent =
        conectados;


    equiposDesconectados.textContent =
        desconectados;
}


/* =========================================
   ÁREAS
========================================= */

function actualizarFiltroAreas() {

    const areas =
        [
            ...new Set(
                equiposData
                    .map(
                        equipo =>
                            equipo.Area
                    )
                    .filter(Boolean)
            )
        ]
        .sort(
            (a, b) =>
                String(a)
                    .localeCompare(
                        String(b),
                        "es"
                    )
        );


    filtroAreaEquipo.innerHTML = `
        <option value="todos">
            Todas las áreas
        </option>
    `;


    areas.forEach(
        area => {

            const option =
                document.createElement(
                    "option"
                );


            option.value =
                area;


            option.textContent =
                area;


            filtroAreaEquipo.appendChild(
                option
            );
        }
    );
}


/* =========================================
   FILTRAR EQUIPOS
========================================= */

function aplicarFiltrosEquipos(conservarPagina) {

    const texto =
        normalizarTexto(
            buscarEquipo.value
        );


    const estado =
        filtroEstadoEquipo.value;


    const area =
        filtroAreaEquipo.value;


    equiposFiltrados =
        equiposData.filter(
            equipo => {

                const coincideTexto =
                    !texto ||
                    normalizarTexto(
                        equipo.Nombre
                    ).includes(texto) ||
                    normalizarTexto(
                        equipo.IP
                    ).includes(texto) ||
                    normalizarTexto(
                        equipo.DNSHostName
                    ).includes(texto);


                const coincideEstado =
                    estado === "todos" ||
                    String(
                        equipo.Estado || ""
                    ) === estado;


                const coincideArea =
                    area === "todos" ||
                    String(
                        equipo.Area || ""
                    ) === area;


                return (
                    coincideTexto &&
                    coincideEstado &&
                    coincideArea
                );
            }
        );


    if (!conservarPagina) {
        paginaActual = 1;
    }


    renderizarTablaEquipos();
}


/* =========================================
   CLASE ESTADO
========================================= */

function obtenerClaseEstado(
    estado
) {

    const estadoNormalizado =
        normalizarTexto(
            estado
        );


    if (
        estadoNormalizado ===
        "conectado"
    ) {

        return "online";
    }


    if (
        estadoNormalizado ===
        "desconectado"
    ) {

        return "offline";
    }


    return "neutral";
}


/* =========================================
   RENDER TABLA
========================================= */

function renderizarTablaEquipos() {

    const totalResultados =
        equiposFiltrados.length;


    const totalPaginas =
        Math.max(
            1,
            Math.ceil(
                totalResultados /
                EQUIPOS_POR_PAGINA
            )
        );


    if (
        paginaActual >
        totalPaginas
    ) {

        paginaActual =
            totalPaginas;
    }


    const inicio =
        (
            paginaActual - 1
        ) *
        EQUIPOS_POR_PAGINA;


    const fin =
        inicio +
        EQUIPOS_POR_PAGINA;


    const equiposPagina =
        equiposFiltrados.slice(
            inicio,
            fin
        );


    if (
        equiposResultados
    ) {

        if (
            totalResultados === 0
        ) {

            equiposResultados.textContent =
                "0 equipos encontrados";

        } else {

            equiposResultados.textContent =
                `Equipos encontrados: ${totalResultados}`;
        }
    }


    if (
        totalResultados === 0
    ) {

        equiposTableBody.innerHTML =
            "";

        equiposEmpty.style.display =
            "flex";


        renderizarPaginacion(
            0
        );


        return;
    }


    equiposEmpty.style.display =
        "none";


    equiposTableBody.innerHTML =
        equiposPagina
            .map(
                equipo => {

                    const estado =
                        equipo.Estado ||
                        "No verificable";


                    const claseEstado =
                        obtenerClaseEstado(
                            estado
                        );


                    /*
                     * SOLO SE CONSULTA LA MEMORIA.
                     *
                     * No se hace ninguna petición
                     * a SharePoint aquí.
                     */

                    const trabajador =
                        obtenerTrabajadorEquipo(
                            equipo.Nombre
                        );


                    const trabajadorHTML =
                        trabajador
                            ? `
                                <span
                                    class="trabajador-name"
                                    title="${escaparHTML(
                                        trabajador.nombre
                                    )}"
                                >
                                    ${escaparHTML(
                                        trabajador.nombre
                                    )}
                                </span>
                            `
                            : !trabajadoresListos
                            ? `<span style="color:#98a2b3;font-size:12px;">Cargando…</span>`
                            : `
                                <button
                                    type="button"
                                    class="registrar-trabajador"
                                    data-equipo="${escaparHTML(
                                        equipo.Nombre
                                    )}"
                                >
                                    Registrar
                                </button>
                            `;


                    return `
                        <tr>

                            <td>

                                <div class="equipo-name">

                                    <div class="equipo-icon">
                                        🖥️
                                    </div>

                                    <div>

                                        <strong
                                            title="${escaparHTML(
                                                equipo.Nombre
                                            )}"
                                        >
                                            ${escaparHTML(
                                                equipo.Nombre
                                            )}
                                        </strong>

                                        <span
                                            title="${escaparHTML(
                                                equipo.DNSHostName
                                            )}"
                                        >
                                            ${escaparHTML(
                                                equipo.DNSHostName ||
                                                "Sin DNS"
                                            )}
                                        </span>

                                    </div>

                                </div>

                            </td>


                            <td>

                                <span
                                    class="equipo-status ${claseEstado}"
                                >

                                    <span
                                        class="equipo-status-dot"
                                    ></span>

                                    ${escaparHTML(
                                        estado
                                    )}

                                </span>

                            </td>


                            <td class="trabajador-cell">

                                ${trabajadorHTML}

                            </td>


                            <td>

                                <span class="equipo-ip">

                                    ${escaparHTML(
                                        equipo.IP ||
                                        "--"
                                    )}

                                </span>

                            </td>


                            <td>

                                <div class="equipo-so">

                                    ${escaparHTML(
                                        equipo.SistemaOperativo ||
                                        "--"
                                    )}

                                    ${
                                        equipo.VersionSO
                                            ? `
                                                <br>

                                                <small>
                                                    ${escaparHTML(
                                                        equipo.VersionSO
                                                    )}
                                                </small>
                                            `
                                            : ""
                                    }

                                </div>

                            </td>


                            <td>

                                <span
                                    class="equipo-area"
                                    title="${escaparHTML(
                                        equipo.Area
                                    )}"
                                >
                                    ${escaparHTML(
                                        equipo.Area ||
                                        "--"
                                    )}
                                </span>

                            </td>


                            <td>

                                <span class="equipo-date">

                                    ${escaparHTML(
                                        equipo.UltimoRegistroAD ||
                                        "--"
                                    )}

                                </span>

                            </td>

                        </tr>
                    `;
                }
            )
            .join("");


    renderizarPaginacion(
        totalPaginas
    );
}


/* =========================================
   PAGINACIÓN
========================================= */

function renderizarPaginacion(
    totalPaginas
) {

    const existente =
        document.querySelector(
            ".equipos-pagination"
        );


    if (
        existente
    ) {

        existente.remove();
    }


    if (
        totalPaginas <= 1
    ) {

        return;
    }


    const contenedor =
        document.createElement(
            "div"
        );


    contenedor.className =
        "equipos-pagination";


    const botonAnterior =
        document.createElement(
            "button"
        );


    botonAnterior.type =
        "button";


    botonAnterior.className =
        "pagination-btn pagination-prev";


    botonAnterior.innerHTML =
        "‹";


    botonAnterior.disabled =
        paginaActual === 1;


    botonAnterior.addEventListener(
        "click",
        () => {

            if (
                paginaActual > 1
            ) {

                paginaActual--;

                renderizarTablaEquipos();
            }
        }
    );


    contenedor.appendChild(
        botonAnterior
    );


    const paginas =
        generarPaginas(
            paginaActual,
            totalPaginas
        );


    paginas.forEach(
        pagina => {

            if (
                pagina === "..."
            ) {

                const puntos =
                    document.createElement(
                        "span"
                    );


                puntos.className =
                    "pagination-ellipsis";


                puntos.textContent =
                    "...";


                contenedor.appendChild(
                    puntos
                );


                return;
            }


            const boton =
                document.createElement(
                    "button"
                );


            boton.type =
                "button";


            boton.className =
                "pagination-btn";


            boton.textContent =
                pagina;


            if (
                pagina === paginaActual
            ) {

                boton.classList.add(
                    "active"
                );
            }


            boton.addEventListener(
                "click",
                () => {

                    paginaActual =
                        pagina;


                    renderizarTablaEquipos();
                }
            );


            contenedor.appendChild(
                boton
            );
        }
    );


    const botonSiguiente =
        document.createElement(
            "button"
        );


    botonSiguiente.type =
        "button";


    botonSiguiente.className =
        "pagination-btn pagination-next";


    botonSiguiente.innerHTML =
        "›";


    botonSiguiente.disabled =
        paginaActual ===
        totalPaginas;


    botonSiguiente.addEventListener(
        "click",
        () => {

            if (
                paginaActual <
                totalPaginas
            ) {

                paginaActual++;

                renderizarTablaEquipos();
            }
        }
    );


    contenedor.appendChild(
        botonSiguiente
    );


    const tableContainer =
        document.querySelector(
            ".equipos-table-container"
        );


    if (
        tableContainer
    ) {

        tableContainer.parentNode.insertBefore(
            contenedor,
            tableContainer.nextSibling
        );
    }
}


/* =========================================
   GENERAR PÁGINAS
========================================= */

function generarPaginas(
    actual,
    total
) {

    if (
        total <= 7
    ) {

        return Array.from(
            {
                length: total
            },
            (_, index) =>
                index + 1
        );
    }


    const paginas =
        [];


    paginas.push(
        1
    );


    if (
        actual > 4
    ) {

        paginas.push(
            "..."
        );
    }


    const inicio =
        Math.max(
            2,
            actual - 1
        );


    const fin =
        Math.min(
            total - 1,
            actual + 1
        );


    for (
        let i = inicio;
        i <= fin;
        i++
    ) {

        paginas.push(
            i
        );
    }


    if (
        actual <
        total - 3
    ) {

        paginas.push(
            "..."
        );
    }


    paginas.push(
        total
    );


    return paginas;
}


/* =========================================
   ABRIR MODAL
========================================= */

async function abrirModalTrabajador(
    nombreEquipo
) {

    equipoSeleccionado =
        nombreEquipo;


    const equipoInfo = equiposData.find(e => e.Nombre === nombreEquipo);

    modalEquipoNombre.textContent =
        (equipoInfo && equipoInfo.IP)
            ? `${nombreEquipo} — ${equipoInfo.IP}`
            : nombreEquipo;


    trabajadorNombre.value =
        "";


    trabajadorError.style.display =
        "none";


    trabajadorError.textContent =
        "";


    guardarTrabajador.disabled =
        false;


    guardarTrabajador.textContent =
        "Guardar";


    trabajadorModal.style.display =
        "flex";


    /*
     * Si ya tenemos trabajadores cargados,
     * mostramos inmediatamente la asignación.
     */

    const trabajador =
        obtenerTrabajadorEquipo(
            nombreEquipo
        );


    if (
        trabajador
    ) {

        trabajadorNombre.value =
            trabajador.nombre;

    } else {

        /*
         * Si todavía no terminaron de cargar,
         * intentamos cargar la información.
         *
         * Esto ocurre únicamente al abrir
         * el modal, nunca durante el inventario.
         */

        if (
            !trabajadoresCargados
        ) {

            trabajadorError.textContent =
                "Consultando asignación...";

            trabajadorError.style.display =
                "block";


            try {

                await obtenerTrabajadores(
                    tokenActual
                );


                const trabajadorActual =
                    obtenerTrabajadorEquipo(
                        nombreEquipo
                    );


                if (
                    trabajadorActual
                ) {

                    trabajadorNombre.value =
                        trabajadorActual.nombre;

                    trabajadorError.style.display =
                        "none";

                } else {

                    trabajadorError.style.display =
                        "none";
                }


            } catch (error) {

                console.warn(
                    "Error consultando trabajador:",
                    error
                );


                trabajadorError.textContent =
                    "No se pudo consultar la asignación.";

                trabajadorError.style.display =
                    "block";
            }
        }
    }


    setTimeout(
        () => {

            trabajadorNombre.focus();

        },
        50
    );
}


/* =========================================
   CERRAR MODAL
========================================= */

function cerrarModalTrabajador() {

    trabajadorModal.style.display =
        "none";


    equipoSeleccionado =
        null;


    trabajadorNombre.value =
        "";


    trabajadorError.textContent =
        "";


    trabajadorError.style.display =
        "none";


    guardarTrabajador.disabled =
        false;


    guardarTrabajador.textContent =
        "Guardar";
}


/* =========================================
   ERROR MODAL
========================================= */

function mostrarErrorTrabajador(
    mensaje
) {

    trabajadorError.textContent =
        mensaje;


    trabajadorError.style.display =
        "block";
}


/* =========================================
   GUARDAR TRABAJADOR
========================================= */

async function guardarAsignacionTrabajador() {
    const nombre = trabajadorNombre.value.trim();

    if (!equipoSeleccionado) {
        mostrarErrorTrabajador(
            "No se ha seleccionado ningún equipo."
        );
        return;
    }

    if (!nombre) {
        mostrarErrorTrabajador(
            "Ingresa el nombre del trabajador."
        );
        trabajadorNombre.focus();
        return;
    }

    try {
        guardarTrabajador.disabled = true;
        trabajadorError.style.display = "none";
        guardarTrabajador.textContent = "Guardando...";

        const sitio =
            await obtenerSitioSharePoint(tokenActual);

        const lista =
            await obtenerListaTrabajadores(tokenActual);

        const columnas =
            await obtenerColumnasTrabajadores(tokenActual);

        console.log(
            "COLUMNAS DE TRABAJADORES:",
            columnas.map(columna => ({
                displayName: columna.displayName,
                name: columna.name,
                required: columna.required,
                readOnly: columna.readOnly
            }))
        );

        // ==========================================
        // BUSCAR COLUMNAS
        // ==========================================

        const campoEquipo =
            obtenerNombreInternoColumna(
                columnas,
                "Equipo1"
            );

        const campoTrabajador =
            obtenerNombreInternoColumna(
                columnas,
                "NombreTrabajador"
            );

        if (!campoEquipo) {
            throw new Error(
                'No se encontró la columna "Equipo1".'
            );
        }

        if (!campoTrabajador) {
            throw new Error(
                'No se encontró la columna "NombreTrabajador".'
            );
        }

        // Nunca permitir que la columna del trabajador
        // sea Title o LinkTitle.
        if (
            campoTrabajador === "Title" ||
            campoTrabajador === "LinkTitle"
        ) {
            throw new Error(
                `La columna "NombreTrabajador" está resolviendo incorrectamente a "${campoTrabajador}".`
            );
        }

        // ==========================================
        // COMPROBAR SI YA EXISTE ASIGNACIÓN
        // ==========================================

        let asignacionActual =
            obtenerTrabajadorEquipo(
                equipoSeleccionado
            );

        if (
            !asignacionActual &&
            !trabajadoresCargados
        ) {
            await cargarTrabajadoresDesdeSharePoint();

            asignacionActual =
                obtenerTrabajadorEquipo(
                    equipoSeleccionado
                );
        }

        // ==========================================
        // CAMPOS QUE SE ENVIARÁN A SHAREPOINT
        // ==========================================

        const fields = {};

        fields[campoEquipo] =
            equipoSeleccionado;

        fields[campoTrabajador] =
            nombre;

        console.log(
            "CAMPOS ENVIADOS A SHAREPOINT:",
            fields
        );

        let response;

        // ==========================================
        // ACTUALIZAR REGISTRO EXISTENTE
        // ==========================================

        if (
            asignacionActual &&
            asignacionActual.itemId
        ) {

            const url =
                `https://graph.microsoft.com/v1.0/sites/${sitio.id}` +
                `/lists/${lista.id}` +
                `/items/${asignacionActual.itemId}/fields`;

            console.log(
                "Actualizando asignación:",
                url
            );

            response = await graphFetch(
                url,
                {
                    method: "PATCH",

                    headers: {
                        "Content-Type":
                            "application/json"
                    },

                    body:
                        JSON.stringify(fields)
                }
            );

        }

        // ==========================================
        // CREAR NUEVO REGISTRO
        // ==========================================

        else {

            const url =
                `https://graph.microsoft.com/v1.0/sites/${sitio.id}` +
                `/lists/${lista.id}` +
                `/items`;

            console.log(
                "Creando nueva asignación:",
                url
            );

            response = await graphFetch(
                url,
                {
                    method: "POST",

                    headers: {
                        "Content-Type":
                            "application/json"
                    },

                    body:
                        JSON.stringify({
                            fields: fields
                        })
                }
            );
        }

        // ==========================================
        // VALIDAR RESPUESTA
        // ==========================================

        if (!response.ok) {

            const texto =
                await response.text();

            console.error(
                "ERROR SHAREPOINT:",
                response.status,
                texto
            );

            throw new Error(
                `SharePoint respondió ${response.status}: ${texto}`
            );
        }

        const resultado =
            await response.json();

        console.log(
            "ASIGNACIÓN GUARDADA CORRECTAMENTE:",
            resultado
        );

        // ==========================================
        // ACTUALIZAR DATOS LOCALES
        // ==========================================

        trabajadoresPorEquipo[
            normalizarTexto(
                equipoSeleccionado
            )
        ] = {
            nombre: nombre,

            itemId:
                resultado.id ||
                asignacionActual?.itemId,

            equipo:
                equipoSeleccionado
        };

        trabajadoresCargados = true;
        trabajadoresListos = true;

        // ==========================================
        // ACTUALIZAR INTERFAZ
        // ==========================================

        cerrarModalTrabajador();

        renderizarTablaEquipos();

    } catch (error) {

        console.error(
            "Error guardando trabajador:",
            error
        );

        mostrarErrorTrabajador(
            obtenerMensajeError(error)
        );

    } finally {

        guardarTrabajador.disabled = false;

        guardarTrabajador.textContent =
            "Guardar";
    }
}
/* =========================================
   MENSAJES DE ERROR
========================================= */

function obtenerMensajeError(
    error
) {

    const mensaje =
        String(
            error?.message ||
            error ||
            ""
        );


    if (
        mensaje.includes(
            "403"
        )
    ) {

        return (
            "No tienes permisos para modificar esta lista."
        );
    }


    if (
        mensaje.includes(
            "401"
        )
    ) {

        return (
            "La sesión expiró. Recarga la página e inicia sesión nuevamente."
        );
    }


    return (
        mensaje ||
        "No se pudo guardar la asignación."
    );
}


/* =========================================
   EVENTOS TABLA
========================================= */

if (
    equiposTableBody
) {

    equiposTableBody.addEventListener(
        "click",
        event => {

            const boton =
                event.target.closest(
                    ".registrar-trabajador"
                );


            if (
                !boton
            ) {
                return;
            }


            const equipo =
                boton.dataset.equipo;


            if (
                !equipo
            ) {
                return;
            }


            abrirModalTrabajador(
                equipo
            );
        }
    );
}


/* =========================================
   EVENTOS FILTROS
========================================= */

if (
    buscarEquipo
) {

    let temporizadorBusqueda = null;

    buscarEquipo.addEventListener(
        "input",
        () => {
            clearTimeout(temporizadorBusqueda);
            temporizadorBusqueda = setTimeout(aplicarFiltrosEquipos, 200);
        }
    );
}


if (
    filtroEstadoEquipo
) {

    filtroEstadoEquipo.addEventListener(
        "change",
        () => {

            aplicarFiltrosEquipos();
        }
    );
}


if (
    filtroAreaEquipo
) {

    filtroAreaEquipo.addEventListener(
        "change",
        () => {

            aplicarFiltrosEquipos();
        }
    );
}


/* =========================================
   EVENTOS MODAL
========================================= */

if (
    cerrarTrabajadorModal
) {

    cerrarTrabajadorModal.addEventListener(
        "click",
        cerrarModalTrabajador
    );
}


if (
    cancelarTrabajador
) {

    cancelarTrabajador.addEventListener(
        "click",
        cerrarModalTrabajador
    );
}


if (
    guardarTrabajador
) {

    guardarTrabajador.addEventListener(
        "click",
        guardarAsignacionTrabajador
    );
}


if (
    trabajadorNombre
) {

    trabajadorNombre.addEventListener(
        "keydown",
        event => {

            if (
                event.key ===
                "Enter"
            ) {

                event.preventDefault();

                guardarAsignacionTrabajador();
            }


            if (
                event.key ===
                "Escape"
            ) {

                cerrarModalTrabajador();
            }
        }
    );
}


if (
    trabajadorModal
) {

    trabajadorModal.addEventListener(
        "click",
        event => {

            if (
                event.target ===
                trabajadorModal
            ) {

                cerrarModalTrabajador();
            }
        }
    );
}


/* =========================================
   CARGAR EQUIPOS
   -----------------------------------------
   ESTA ES LA PARTE CRÍTICA.
   
   NO CARGA TRABAJADORES.
   NO ESPERA TRABAJADORES.
   NO CONSULTA LA LISTA.
========================================= */

/* Aplica datos nuevos SIN perder búsqueda, filtros ni página actual. */
function aplicarDatosEquipos(equipos) {
    equiposData = equipos;
    const areaPrevia = filtroAreaEquipo.value;
    actualizarResumenEquipos();
    actualizarFiltroAreas();
    if ([...filtroAreaEquipo.options].some(o => o.value === areaPrevia)) {
        filtroAreaEquipo.value = areaPrevia;
    }
    aplicarFiltrosEquipos(true);
}

async function cargarEquipos() {
    if (cargandoEquipos) return;
    cargandoEquipos = true;
    try {
        equiposEstado.textContent = "● ACTUALIZANDO";
        equiposEstado.className = "section-status warning";

        const equipos = await obtenerEquiposDesdeSharePoint();
        if (equipos) {
            aplicarDatosEquipos(equipos);
            guardarCache("equipos", { etag: etagEquipos, datos: equipos });
        }

        lastUpdate.textContent = obtenerFechaActual();
        equiposEstado.textContent = "● ACTUALIZADO";
        equiposEstado.className = "section-status online";
    } catch (error) {
        console.error("Error cargando equipos:", error);
        mostrarDiagnostico("Inventario (equipos.json)", error);
        if (equiposData.length > 0) {
            equiposEstado.textContent = "● SIN ACTUALIZAR";
            equiposEstado.className = "section-status warning";
        } else {
            equiposEstado.textContent = "● ERROR";
            equiposEstado.className = "section-status offline";
            equiposTableBody.innerHTML = `
                <tr><td colspan="7" class="equipos-loading">
                    <div class="table-loading"><span>No se pudo cargar el inventario.</span></div>
                </td></tr>`;
            if (equiposResultados) equiposResultados.textContent = "Error al obtener los equipos.";
        }
    } finally {
        cargandoEquipos = false;
    }
}


/* =========================================
   CARGAR SERVICIOS
========================================= */

async function cargarServicios() {

    try {

        const servicios =
            await obtenerMonitoreoTI(
                tokenActual
            );


        renderizarServicios(
            servicios
        );


    } catch (error) {

        console.error(
            "Error cargando servicios:",
            error
        );

        mostrarDiagnostico("Servicios (lista " + SHAREPOINT_LIST_SERVICIOS + ")", error);


        serviciosGrid.innerHTML = `
            <div
                class="service-card offline"
            >

                <div class="service-main">

                    <div class="service-icon">
                        ⚠️
                    </div>

                    <div class="service-info">

                        <div class="service-heading">

                            <h3>
                                Error de conexión
                            </h3>

                            <span
                                class="service-status offline"
                            >

                                <span class="status-dot"></span>

                                ERROR

                            </span>

                        </div>

                        <p>
                            No se pudieron obtener
                            los servicios desde SharePoint.
                        </p>

                    </div>

                </div>

            </div>
        `;


        estadoGeneral.textContent =
            "● ERROR";


        estadoGeneral.className =
            "section-status offline";
    }
}


/* =========================================
   CARGA INICIAL
========================================= */

function mostrarBotonLogin(mensaje) {
    mostrarLoading(mensaje);
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
            console.error("Error de inicio de sesión:", error);
            mostrarDiagnostico("Inicio de sesión", error);
            loadingMessage.textContent = "No se pudo iniciar sesión. Permite las ventanas emergentes e inténtalo de nuevo.";
        } finally {
            boton.disabled = false;
        }
    });
    loadingOverlay.querySelector(".loading-card").appendChild(boton);
}

/* Cada sección se pinta apenas llegan sus datos; ninguna espera a la otra. */
function iniciarCargas() {
    ultimaActualizacion = Date.now();
    cargarServicios();

    /* 1.º se llena la tabla con equipos.json; 2.º se consultan las asignaciones. */
    cargarEquipos().then(() => {
        if (equiposData.length > 0) {
            return obtenerTrabajadores();
        }
    }).catch(error => console.warn("Carga de trabajadores falló:", error));
}

async function cargarDashboard() {
    try {
        if (!msalInstanceInfraestructura) {
            mostrarLoading("No se pudo cargar el componente de Microsoft (MSAL). Revisa tu conexión o bloqueadores y recarga la página.");
            return;
        }

        /* Si ya hay sesión, mostramos el último inventario guardado al instante. */
        let conCache = false;
        if (msalInstanceInfraestructura.getAllAccounts().length > 0) {
            const cache = leerCache("equipos");
            if (cache && Array.isArray(cache.datos) && cache.datos.length > 0) {
                etagEquipos = cache.etag || null;
                aplicarDatosEquipos(cache.datos);
                conCache = true;
            }
        }
        if (!conCache) mostrarLoading("Iniciando sesión...");

        const token = await iniciarSesion(false);
        if (!token) {
            mostrarBotonLogin("Inicia sesión para ver el monitoreo.");
            return;
        }
        ocultarLoading();
        iniciarCargas();
    } catch (error) {
        console.error("Error inicializando infraestructura:", error);
        mostrarDiagnostico("Inicialización", error);
        mostrarBotonLogin("No se pudo cargar el monitoreo.");
    }
}


/* =========================================
   ACTUALIZACIÓN AUTOMÁTICA
========================================= */

async function actualizarDashboard() {
    if (document.hidden || !tokenActual) return;   /* no gastar red en pestañas ocultas */
    try {
        if (!(await obtenerTokenSilencioso())) return;
        ultimaActualizacion = Date.now();
        await Promise.all([cargarServicios(), cargarEquipos()]);
    } catch (error) {
        console.error("Error en actualización automática:", error);
    }
}


/* =========================================
   INICIO
========================================= */

function iniciarInfraestructura() {
    cargarDashboard();

    /* Si a los 20 s no hay datos ni errores visibles, lo decimos en pantalla. */
    setTimeout(() => {
        if (equiposData.length === 0 && !document.getElementById("panelDiagnostico")) {
            mostrarDiagnostico(
                "Sin respuesta tras 20 s",
                new Error("Aún no llegan datos. Si ves el botón «Iniciar sesión con Microsoft», púlsalo. Si no, abre F12 → Consola y envíame el primer error en rojo.")
            );
        }
    }, 20000);

    setInterval(actualizarDashboard, INTERVALO_ACTUALIZACION);
    document.addEventListener("visibilitychange", () => {
        if (!document.hidden && Date.now() - ultimaActualizacion > 30000) {
            actualizarDashboard();
        }
    });
}

if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", iniciarInfraestructura);
} else {
    iniciarInfraestructura();
}